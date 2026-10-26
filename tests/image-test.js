'use strict';
// Image editing: View/Edit, crop and rotate, adjust, filters, draw, text, undo, remove background (AI), resize,
// saving as PNG/JPEG/WebP/BMP, converting, and the Scan tool's window.
//   DOCSTUDIO_OUT=<dir> electron tests/image-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const outDir = process.env.DOCSTUDIO_OUT || path.join(os.tmpdir(), 'docstudio-image-out');
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-image-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail });

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
  const errors = [];
  win.webContents.on('console-message', (_e, level, message) => { if (level >= 3) errors.push(message); });
  try {
    // A made-up photo: sky, grass and a red ball, saved as a PNG file and opened.
    const photo = path.join(outDir, 'photo.png');
    const png = await run(`(async () => {
      const c = document.createElement('canvas'); c.width = 1200; c.height = 800; const g = c.getContext('2d');
      const sky = g.createLinearGradient(0, 0, 0, 800); sky.addColorStop(0, '#6aa8e8'); sky.addColorStop(0.6, '#cfe6fb'); sky.addColorStop(0.6, '#4f9a3c'); sky.addColorStop(1, '#2f6b22');
      g.fillStyle = sky; g.fillRect(0, 0, 1200, 800);
      const ball = g.createRadialGradient(540, 380, 30, 600, 450, 200); ball.addColorStop(0, '#ff8a80'); ball.addColorStop(1, '#b71c1c');
      g.fillStyle = ball; g.beginPath(); g.arc(600, 450, 200, 0, 7); g.fill();
      return [...new Uint8Array(await new Promise((r) => c.toBlob(r, 'image/png')).then((b) => b.arrayBuffer()))];
    })()`);
    fs.writeFileSync(photo, Buffer.from(png));
    win.webContents.send('app:open-path', photo);
    await wait(1500);
    check('a picture opens in View with an Edit mode', await run(`active.kind === 'image' && active.mode === 'view' && active.modes.join() === 'view,edit'`));
    await run(`setMode(active, 'edit')`);
    await wait(500);
    check('Edit shows the seven tools and Undo/Redo/Revert', await run(`[...document.querySelectorAll('.ie-tab')].map((b) => b.textContent).join() === 'Crop,Adjust,Filters,Draw,Text,Background,Resize' && active.tools.length === 3`));
    const ed = 'active.imageEditor';
    // Adjust and filters are settings on top of the picture.
    await run(`(${ed}.push(), ${ed}.params.adjust.brightness = 30, ${ed}.params.filter = 'mono', ${ed}.render(), ${ed}.changed(), true)`);
    const grey = await run(`(() => { const c = ${ed}.renderFull(); const d = c.getContext('2d').getImageData(600, 450, 1, 1).data; return [d[0], d[1], d[2]]; })()`);
    check('filters and light are applied (Mono makes the red ball grey)', Math.abs(grey[0] - grey[1]) < 3 && Math.abs(grey[1] - grey[2]) < 3, grey);
    check('an edit marks the picture unsaved', await run('active.dirty'));
    await run(`(${ed}.undo(), true)`);
    check('Undo puts the settings back', await run(`${ed}.params.filter === 'none' && ${ed}.params.adjust.brightness === 0`));
    // Crop to a square, rotate, flip.
    await run(`(${ed}.setTool('crop'), ${ed}.crop.ratio = 1, ${ed}.fitRatio(), ${ed}.applyCrop(), true)`);
    check('crop to a square', JSON.stringify(await run(`[${ed}.width, ${ed}.height]`)) === '[800,800]');
    await run(`(${ed}.setTool('crop'), ${ed}.crop.angle = 10, ${ed}.applyCrop(), true)`);
    check('straighten keeps the size (zoomed so no corners show)', JSON.stringify(await run(`[${ed}.width, ${ed}.height]`)) === '[800,800]');
    await run(`(${ed}.setTool('crop'), ${ed}.crop.ratio = 16 / 9, ${ed}.fitRatio(), ${ed}.applyCrop(), ${ed}.rotate(1), true)`);
    check('crop 16:9 then rotate right gives a tall picture', JSON.stringify(await run(`[${ed}.width, ${ed}.height]`)) === '[450,800]');
    // Draw and text add to the picture.
    const before = await run(`${ed}.past.length`);
    await run(`(${ed}.setTool('text'), ${ed}.txt.text = 'Hello', ${ed}.bakeText(), true)`);
    check('text is added as one Undo step', (await run(`${ed}.past.length`)) === before + 1);
    // Remove background with the AI model; the sky corner goes, the ball stays.
    await run(`(${ed}.undo(), ${ed}.undo(), ${ed}.undo(), ${ed}.undo(), ${ed}.undo(), ${ed}.setTool('background'), true)`);
    const t0 = Date.now();
    await run(`${ed}.removeBackground()`);
    const mask = await run(`(() => { const m = ${ed}.session && ${ed}.session.mask; if (!m) return null; const g = m.getContext('2d'); const a = (x, y) => g.getImageData(Math.round(x * m.width), Math.round(y * m.height), 1, 1).data[3]; return { corner: a(0.03, 0.03), centre: a(0.5, 0.56) }; })()`);
    check('remove background finds the subject', mask && mask.corner < 30 && mask.centre > 220, { mask, ms: Date.now() - t0 });
    await run(`(${ed}.session.fill = '#ffffff', ${ed}.applySession(), true)`);
    const corner = await run(`(() => { const d = ${ed}.renderFull().getContext('2d').getImageData(5, 5, 1, 1).data; return [...d]; })()`);
    check('a white background replaces it', corner.join() === '255,255,255,255', corner);
    // Resize.
    await run(`(() => { const W = ${ed}.width, H = ${ed}.height; ${ed}.commit((() => { const c = document.createElement('canvas'); c.width = W / 2; c.height = H / 2; c.getContext('2d').drawImage(${ed}.base, 0, 0, c.width, c.height); return c; })()); return true; })()`);
    // Save as each type.
    const types = await run(`(async () => { const out = {}; for (const x of ['png', 'jpg', 'webp', 'bmp']) { const b = await ${ed}.export(x); out[x] = [...new Uint8Array(b.slice(0, 4))].map((v) => v.toString(16).padStart(2, '0')).join(''); } return out; })()`);
    check('saves as PNG, JPEG, WebP and BMP', types.png === '89504e47' && types.jpg.startsWith('ffd8ff') && types.webp === '52494646' && types.bmp.startsWith('424d'), types);
    await run(`save(active)`);
    await wait(500);
    const saved = fs.readFileSync(photo);
    check('Save writes the edited PNG over the file', saved.slice(0, 4).toString('hex') === '89504e47' && !(await run('active.dirty')) && saved.length !== png.length, saved.length);
    check('the original was kept as a backup', fs.readdirSync(path.join(app.getPath('userData'), 'backups')).length === 1);
    // Convert the picture to BMP and PDF.
    const conv = await run(`(async () => { const b = await Convert.run(active, 'bmp'); const p = await Convert.run(active, 'pdf'); return [new Uint8Array(b)[0], new Uint8Array(b)[1], new TextDecoder().decode(new Uint8Array(p).slice(0, 5))]; })()`);
    check('converts to BMP and PDF', conv[0] === 0x42 && conv[1] === 0x4d && conv[2] === '%PDF-', conv);
    // View shows the picture.
    await run(`setMode(active, 'view')`);
    await wait(400);
    check('View shows the saved picture', await run(`!!active.body.querySelector('.image-view img')`));
    fs.writeFileSync(path.join(outDir, 'image.png'), (await win.webContents.capturePage()).toPNG());
    // The Scan tool window (the scanner list depends on this PC).
    await run(`(Tools.scan(), true)`);
    await wait(1500);
    check('Scan opens its tool window', await run(`!!document.querySelector('.scan-tool') && document.querySelectorAll('.scan-tool .tool-field select').length === 4`));
    await run(`(document.querySelectorAll('.backdrop').forEach((b) => b.remove()), true)`);
    check('no errors in the console', !errors.filter((e) => !/Autofill|ResizeObserver/.test(e)).length, errors.slice(0, 5));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 600)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
