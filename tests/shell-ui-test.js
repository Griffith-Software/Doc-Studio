'use strict';
// The Doc Studio window: app bar, action bar, sidebar pages, page bar, thumbnails, layouts, starring, search,
// theme, and the tools (merge, split, compress, OCR):
//   DOCSTUDIO_DOCX=<multi-page.docx> DOCSTUDIO_PDF=<report.pdf> DOCSTUDIO_OUT=<dir> electron tests/shell-ui-test.js
const { app, BrowserWindow, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-shellui-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => { checks.push({ name, ok: !!ok, detail }); console.log('...', ok ? 'ok' : 'FAIL', name); };

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1500, 920);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const errors = [];
  win.webContents.on('console-message', (_e, level, msg) => { if (level >= 3) errors.push(msg); });
  const run = async (js) => {
    const r = await win.webContents.executeJavaScript(`(async () => { try { return await (${js}); } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${js.slice(0, 160)}\n${r.__err}`);
    return r;
  };
  const until = async (js, ms = 20000) => { const t = Date.now(); while (Date.now() - t < ms) { if (await run(js)) return true; await wait(200); } return false; };
  const visible = (sel) => run(`(() => { const e = document.querySelector(${JSON.stringify(sel)}); return !!e && !!e.offsetParent; })()`);
  try {
    await wait(600);
    check('app name and title', (await run('document.title')) === 'Doc Studio' && app.getName() === 'Doc Studio');
    check('title bar drawn by the page (Windows caption buttons over it)', await visible('#appbar') && await visible('#brand img') && (await run(`document.querySelector('.brand-name').textContent`)) === 'Doc Studio');
    check('action bar', await run(`['#btn-open', '#btn-open-more', '#ab-save', '#ab-print', '#ab-edit', '#ab-annotate', '#ab-convert', '#ab-more'].every((s) => document.querySelector(s))`) && await run(`document.querySelector('#ab-save').disabled`));
    const side = await run(`[...document.querySelectorAll('#sidebar .side-label')].map((x) => x.textContent)`);
    check('sidebar items', ['Home', 'Documents', 'Recent', 'Starred', 'Cloud Storage', 'Convert', 'OCR (Scan to Text)', 'Compare', 'Merge', 'Split', 'Compress'].every((x) => side.includes(x)), side);
    check('home page with new-document cards and tools', await visible('#welcome .new-card') && await run(`document.querySelectorAll('#welcome .tool-tile').length === 8`));

    // Open: one click shows one file dialog (and a quick double-click still only one); it closes when a file is picked.
    const { ipcMain } = require('electron');
    let dialogs = 0;
    ipcMain.removeHandler('file:openDialog');
    ipcMain.handle('file:openDialog', async () => { dialogs++; await wait(700); return []; });
    await run(`(document.querySelector('#btn-open').click(), true)`);
    await wait(1200);
    check('Open shows one file dialog', dialogs === 1, dialogs);
    dialogs = 0;
    await run(`(document.querySelector('#btn-open').click(), document.querySelector('#btn-open').click(), true)`);
    await wait(1200);
    check('a double-click on Open still shows one dialog', dialogs === 1, dialogs);
    const mdFile = path.join(outDir, 'notes.md');
    fs.writeFileSync(mdFile, ['# Notes', '', 'Some *instructions* to read.', ''].join(String.fromCharCode(10)));
    ipcMain.removeHandler('file:openDialog');
    ipcMain.handle('file:openDialog', async () => { dialogs++; return [mdFile]; });
    await run(`(document.querySelector('#btn-open').click(), true)`);
    check('a Markdown file opens in Preview', await until(`!!(active && active.kind === 'md' && active.mode === 'preview')`, 8000), await run(`active && active.mode`));
    await run(`(async () => { for (const t of [...tabs]) { t.dirty = false; await closeTab(t); } return true; })()`);

    // Sidebar pages.
    await run(`(Shell.go('documents'), true)`);
    check('Documents page lists a folder', await until(`!!document.querySelector('#browser .file-list, #browser .empty')`, 8000) && await run(`document.querySelector('.browse-chips .chip.on') && document.querySelectorAll('.browse-chips .chip.on').length === 1`));
    await run(`(Shell.go('cloud'), true)`);
    check('Cloud Storage page', await until(`!!document.querySelector('#browser .cloud-grid, #browser .empty-cloud')`, 5000));
    await run(`(Shell.go('recent'), true)`);
    check('Recent page', await until(`!!document.querySelector('#browser h1') && document.querySelector('#browser h1').textContent === 'Recent'`));

    // Open a Word document: page bar, thumbnails, layouts.
    win.webContents.send('app:open-path', process.env.DOCSTUDIO_DOCX);
    await until(`!!(active && active.kind === 'docx' && active.viewRendered)`);
    await wait(800);
    check('opening a file shows it (sidebar back on Home)', !(await run(`document.body.classList.contains('on-page')`)) && await run(`document.querySelector('.side-item.on').dataset.page === 'home'`));
    const pg = await run(`({ n: document.querySelector('#pg-num').value, c: document.querySelector('#pg-count').textContent })`);
    check('page number and count', pg.n === '1' && /\/ [2-9]/.test(pg.c), pg);
    check('page thumbnails', await until(`document.querySelectorAll('#thumbs .thumb').length > 1 && !!document.querySelector('#thumbs .thumb.on')`, 5000) && await until(`!!document.querySelector('#thumbs .thumb-paper')`, 5000));
    await run(`(document.querySelector('#pg-next').click(), true)`);
    check('next page', await until(`document.querySelector('#pg-num').value === '2'`, 3000), await run(`document.querySelector('#pg-num').value`));
    await run(`(document.querySelector('#lay-two').click(), true)`);
    check('two pages side by side', await run(`active.el.classList.contains('layout-two') && document.querySelector('#lay-two').classList.contains('on')`));
    await run(`(document.querySelector('#lay-single').click(), true)`);
    await run(`(document.querySelector('#lay-thumbs').click(), true)`);
    check('thumbnails can be hidden', await run(`document.querySelector('#thumbs').hidden`));
    await run(`(document.querySelector('#lay-thumbs').click(), true)`);
    check('zoom from the page bar', await run(`(document.querySelector('#zm-in').click(), active.zoom > 1)`) && await run(`(document.querySelector('#zm-sel').value = '1', document.querySelector('#zm-sel').dispatchEvent(new Event('change')), active.zoom === 1)`));

    // Star it, find it under Starred and in search.
    await run(`(document.querySelector('#db-star').click(), true)`);
    check('star the document', await until(`Shell.isStarred(active.path)`, 3000));
    await run(`(Shell.go('starred'), true)`);
    check('Starred page lists it', await until(`[...document.querySelectorAll('#browser .file-name')].some((x) => x.textContent === active.name)`, 5000));
    await run(`(activate(active), true)`);
    const docs = await run(`Shell.searchDocs(${JSON.stringify(path.basename(process.env.DOCSTUDIO_DOCX).slice(0, 6))}).then((r) => r.map((x) => x.name))`);
    check('search finds documents', docs.length >= 1, docs);
    await run(`(document.querySelector('#menu-search').value = 'word count', Menus.search('word count'), true)`);
    check('search finds commands', await until(`!document.querySelector('#menu-results').hidden && [...document.querySelectorAll('#menu-results .menu-label')].some((x) => x.textContent === 'Word count')`, 4000));
    await run(`(document.querySelector('#menu-results').hidden = true, document.querySelector('#menu-search').value = '', true)`);

    // Menus: under More while viewing, as a menu bar while editing.
    check('menu bar hidden while viewing', await run(`document.querySelector('#menubar').hidden`));
    await run(`(document.querySelector('#ab-more').click(), true)`);
    const more = await until(`!!document.querySelector('.menu-panel:not(#menu-results)')`, 3000) && await run(`[...document.querySelector('.menu-panel:not(#menu-results)').querySelectorAll('.menu-label')].map((x) => x.textContent)`);
    check('More lists every menu', more && ['File', 'Edit', 'View', 'Tools', 'Help', 'Theme', 'Settings…'].every((x) => more.includes(x)), more);
    await run(`(Menus.close(), true)`);
    win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'F', modifiers: ['alt'] });
    check('Alt+F opens File under More', await until(`!!document.querySelector('.menu-panel:not(#menu-results)') && document.querySelector('#ab-more').classList.contains('open')`, 3000));
    await run(`(Menus.close(), true)`);
    await run(`(document.querySelector('#ab-edit').click(), true)`);
    await until(`!!active.editor && active.mode === 'edit'`);
    await wait(500);
    check('Edit switches to editing with menus and tools', await run(`active.mode === 'edit' && !document.querySelector('#menubar').hidden && !document.querySelector('#toolbar').hidden && document.querySelector('#ab-edit').classList.contains('on')`));
    await run(`(document.querySelector('#ab-annotate').click(), true)`);
    const ann = await until(`!!document.querySelector('.menu-panel:not(#menu-results)')`, 3000) && await run(`[...document.querySelector('.menu-panel:not(#menu-results)').querySelectorAll('.menu-label')].map((x) => x.textContent)`);
    check('Annotate offers comment tools', ann && ann.some((x) => /comment/i.test(x)), ann);
    await run(`(Menus.close(), true)`);
    await run(`(document.querySelector('#ab-edit').click(), true)`);
    await until(`active.mode === 'view'`);

    // PDF: read-only view with the same page bar.
    win.webContents.send('app:open-path', process.env.DOCSTUDIO_PDF);
    await until(`!!(active && active.kind === 'pdf' && active.pdfEditor)`);
    await wait(800);
    check('PDF view uses the page renderer, read-only', await run(`active.pdfEditor.readOnly && !!active.pdfEditor.root.classList.contains('readonly') && document.querySelector('#pg-count').textContent === '/ ' + active.pdfEditor.pages.length`));
    check('PDF thumbnails on the right', await run(`(() => { const s = active.pdfEditor.side.getBoundingClientRect(), m = active.pdfEditor.main.getBoundingClientRect(); return s.left > m.left; })()`));
    await run(`(document.querySelector('#ab-annotate').click(), true)`);
    check('Annotate on a PDF switches to highlighting', await until(`active.mode === 'edit' && active.pdfEditor.tool === 'highlight'`, 5000));
    await run(`setMode(active, 'view')`);
    check('PDF print pages are prepared', await run(`active.beforePrint().then(() => active.printBox.querySelectorAll('img').length === active.pdfEditor.pages.length)`));

    // Tools on real data.
    const pdfBytes = `(await desk.read(${JSON.stringify(process.env.DOCSTUDIO_PDF)})).data`;
    const merged = await run(`(async () => { const d = ${pdfBytes}; const c = document.createElement('canvas'); c.width = 400; c.height = 300; const g = c.getContext('2d'); g.fillStyle = '#c00'; g.fillRect(0, 0, 400, 300); const png = await (await new Promise((r) => c.toBlob(r, 'image/png'))).arrayBuffer(); const r = await Tools.mergeFiles([{ name: 'a.pdf', data: d }, { name: 'b.png', data: png }, { name: 'c.pdf', data: d }]); return r.pages; })()`);
    const n = await run(`active.pdfEditor.pages.length`);
    check('merge PDFs and a picture into one PDF', merged === n * 2 + 1, { merged, n });
    const parts = await run(`(async () => { const d = ${pdfBytes}; const out = await Tools.splitPdf(d, Tools.parseRanges('1, 2-end', ${n})); const L = DDPdf.PDFLib.PDFDocument; return Promise.all(out.map(async (b) => (await L.load(b)).getPageCount())); })()`);
    check('split a PDF by ranges', JSON.stringify(parts) === JSON.stringify([1, n - 1]), parts);
    const sizes = await run(`(async () => { const d = ${pdfBytes}; const b = { set() {} }; const a = await Tools.compressPdf(d, 'lossless', b); const s = await Tools.compressPdf(d, 'small', b); const L = DDPdf.PDFLib.PDFDocument; return { orig: d.byteLength, lossless: a.byteLength, small: s.byteLength, pages: (await L.load(s)).getPageCount() }; })()`);
    check('compress a PDF (both kinds stay valid)', sizes.lossless > 0 && sizes.small > 0 && sizes.pages === n, sizes);
    const zipped = await run(`(async () => { const d = (await desk.read(${JSON.stringify(process.env.DOCSTUDIO_DOCX)})).data; const o = await Tools.compressZip(d, 0.7, 1200, { set() {} }); const z = await JSZip.loadAsync(o); return { ok: !!z.file('word/document.xml'), before: d.byteLength, after: o.byteLength }; })()`);
    check('compress an Office file (package still valid)', zipped.ok && zipped.after > 0, zipped);
    // OCR with Windows' text recognition.
    const text = await run(`(async () => { const c = document.createElement('canvas'); c.width = 1400; c.height = 300; const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, 1400, 300); g.fillStyle = '#000'; g.font = '64px Arial'; g.fillText('Quarterly results improved', 40, 120); g.fillText('Invoice number 4821', 40, 230); const png = await (await new Promise((r) => c.toBlob(r, 'image/png'))).arrayBuffer(); try { return (await desk.ocr([{ data: png, ext: 'png' }]))[0]; } catch (e) { return 'ERROR ' + e.message; } })()`);
    check('OCR reads text from a picture', /Quarterly results improved/i.test(text) && /4821/.test(text), text);

    // Theme: dark and light.
    await run(`desk.theme('dark')`);
    await wait(400);
    check('dark mode', await run(`matchMedia('(prefers-color-scheme: dark)').matches && getComputedStyle(document.querySelector('#sidebar')).backgroundColor !== 'rgb(247, 248, 250)'`));
    fs.writeFileSync(path.join(outDir, 'dark.png'), (await win.webContents.capturePage()).toPNG());
    await run(`desk.theme('light')`);
    await wait(400);
    check('light mode', await run(`!matchMedia('(prefers-color-scheme: dark)').matches && getComputedStyle(document.querySelector('#sidebar')).backgroundColor === 'rgb(247, 248, 250)'`));
    fs.writeFileSync(path.join(outDir, 'light.png'), (await win.webContents.capturePage()).toPNG());
    await run(`desk.theme('system')`);

    // Default apps: Settings has the button; the list shows every registered type with the app that opens it.
    await run(`(document.querySelector('#btn-settings').click(), true)`);
    await wait(500);
    check('Settings has the Default apps button', await run(`!!document.querySelector('.settings-defaults button')`));
    await run(`(document.querySelector('.settings-defaults button').click(), true)`);
    check('Default apps lists 35 types with their apps', await until(`document.querySelectorAll('.defaults-row').length === 35 && [...document.querySelectorAll('.defaults-state')].every((s) => s.textContent)`, 15000),
      await run(`[...document.querySelectorAll('.defaults-row')].slice(0, 4).map((r) => r.textContent)`));
    await run(`(document.querySelectorAll('.backdrop').forEach((b) => b.remove()), true)`);

    // New document from the tab "+" and from Open ▾.
    const before = await run('tabs.length');
    await run(`(document.querySelector('.tab-new').click(), true)`);
    await wait(300);
    await run(`(document.querySelector('.sg-menu button') && document.querySelector('.sg-menu button').click(), true)`);
    check('new document from the tab +', await until(`tabs.length === ${before + 1}`, 8000));
    check('no errors in the console', !errors.filter((e) => !/Autofill|ResizeObserver|favicon/.test(e)).length, errors.slice(0, 5));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 700)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
