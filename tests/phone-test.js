'use strict';
// Phone link: the tool window starts the link, a "phone" (HTTP requests from this test) opens the page, is refused
// without the code, sends files (with an unsafe name), downloads what the PC offers; Open all opens the files and
// closing the window stops the link.   DOCSTUDIO_SRC=<a .md file> DOCSTUDIO_OUT=<dir> electron tests/phone-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const http = require('http');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT || path.join(os.tmpdir(), 'docstudio-phone-out');
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-phonetest-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail });
// One HTTP request, as the phone's browser would make it.
const request = (url, method, body) => new Promise((resolve) => {
  const req = http.request(url, { method: method || 'GET', headers: { 'User-Agent': 'Mozilla/5.0 (Linux; Android 14; SM-S918B)' } }, (res) => {
    const parts = [];
    res.on('data', (c) => parts.push(c));
    res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: Buffer.concat(parts) }));
  });
  req.on('error', (err) => resolve({ status: 0, error: err.code }));
  req.setTimeout(4000, () => req.destroy());
  if (body) req.write(body);
  req.end();
});

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1400, 900);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const run = async (js) => {
    const r = await win.webContents.executeJavaScript(`(async () => { try { return await (${js}); } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${js.slice(0, 120)}\n${r.__err}`);
    return r;
  };
  try {
    const md = path.join(outDir, 'notes.md');
    fs.copyFileSync(src, md);
    win.webContents.send('app:open-path', md);
    await wait(1500);
    await run(`(Tools.phone(), true)`);
    let url = null;
    for (let i = 0; i < 30 && !url; i++) { await wait(200); url = await run(`(document.querySelector('.phone-qr') || {}).dataset ? document.querySelector('.phone-qr').dataset.url : null`); }
    check('the window shows a QR code for a home-network address', /^http:\/\/(10|172|192)\.[\d.]+:\d+\/\?k=[0-9a-f]{32}$/.test(url || ''), url);
    const base = url.split('/?k=')[0], k = url.split('k=')[1];
    const page = await request(url);
    check('the phone page opens', page.status === 200 && /Send to your PC/.test(page.body.toString()));
    await wait(300);
    check('the window says the phone connected', /Phone connected \(Android\)/.test(await run(`document.querySelector('.phone-status').textContent`)));
    check('without the code the page is refused', (await request(`${base}/?k=${'0'.repeat(32)}`)).status === 403 && (await request(`${base}/files`)).status === 403);
    // Phone -> PC.
    const photo = await run(`(async () => { const c = document.createElement('canvas'); c.width = 64; c.height = 48; const g = c.getContext('2d'); g.fillStyle = '#c22'; g.fillRect(0, 0, 64, 48); return [...new Uint8Array(await new Promise((r) => c.toBlob(r, 'image/png')).then((b) => b.arrayBuffer()))]; })()`);
    const up1 = await request(`${base}/upload?k=${k}&name=IMG_2041.png`, 'POST', Buffer.from(photo));
    const up2 = await request(`${base}/upload?k=${k}&name=${encodeURIComponent('../../evil?name.txt')}`, 'POST', Buffer.from('Hello from the phone\n'));
    await wait(400);
    check('files from the phone arrive', up1.status === 200 && up2.status === 200);
    const rows = await run(`[...document.querySelectorAll('.phone-list')[0].querySelectorAll('.merge-name')].map((n) => n.textContent)`);
    check('they are listed, with an unsafe name made safe', rows.join() === 'evil_name.txt,IMG_2041.png', rows);
    // PC -> phone.
    await run(`([...document.querySelectorAll('.phone-tool button')].find((b) => b.textContent === 'Add the open document').click(), true)`);
    await wait(500);
    const list = JSON.parse((await request(`${base}/files?k=${k}`)).body.toString());
    check('the open document is offered to the phone', list.length === 1 && list[0].name === 'notes.md', list);
    const dl = await request(`${base}/download?k=${k}&id=${list[0].id}`);
    check('the phone downloads it exactly', dl.status === 200 && dl.body.equals(fs.readFileSync(md)) && /attachment/.test(dl.headers['content-disposition']));
    await wait(300);
    check('the window shows it was downloaded', /Downloaded/.test(await run(`document.querySelectorAll('.phone-list')[1].innerText`)));
    // Closing with unopened files asks first.
    await run(`([...document.querySelectorAll('.phone-tool button')].find((b) => b.textContent === 'Cancel').click(), true)`);
    await wait(500);
    check('closing with unopened files asks first', /not been opened or saved/.test(await run(`(document.querySelector('.msgbox') || {}).innerText || ''`)));
    await run(`([...document.querySelectorAll('.msgbox button')].find((b) => b.textContent === 'Keep the window open').click(), true)`);
    await wait(300);
    // Open all, which closes the window and the link.
    const before = await run('tabs.length');
    await run(`([...document.querySelectorAll('.phone-tool .tool-actions button')].pop().click(), true)`);
    await wait(2500);
    const opened = await run(`tabs.slice(${before}).map((t) => t.name + ':' + t.kind)`);
    check('Open all opens both as new documents', opened.join() === 'evil_name.txt:text,IMG_2041.png:image', opened);
    check('the window closed', !(await run(`!!document.querySelector('.phone-tool')`)));
    check('the link stopped', (await request(`${base}/files?k=${k}`)).status === 0);
    fs.writeFileSync(path.join(outDir, 'phone.png'), (await win.webContents.capturePage()).toPNG());
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 600)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
