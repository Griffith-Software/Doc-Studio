'use strict';
// Split view inside one window: DOCSTUDIO_MD=<file.md> DOCSTUDIO_DOCX=<file.docx> DOCSTUDIO_XLSX=<file.xlsx> DOCSTUDIO_OUT=<dir> electron tests/split-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-split-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => { checks.push({ name, ok: !!ok, detail }); console.log('...', ok ? 'ok' : 'FAIL', name); };

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1600, 950);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const errors = [];
  win.webContents.on('console-message', (_e, level, msg) => { if (level >= 3) errors.push(msg); });
  const run = async (js) => {
    const r = await win.webContents.executeJavaScript(`(async () => { try { return await (${js}); } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${js.slice(0, 160)}\n${r.__err}`);
    return r;
  };
  const until = async (js, ms = 20000) => { const t = Date.now(); while (Date.now() - t < ms) { if (await run(js)) return true; await wait(200); } return false; };
  const click = async (sel) => { const r = await run(`(() => { const e = ${sel}; const b = e.getBoundingClientRect(); return { x: b.left + b.width / 2, y: b.top + Math.min(b.height / 2, 60) }; })()`); win.webContents.sendInputEvent({ type: 'mouseDown', x: Math.round(r.x), y: Math.round(r.y), button: 'left', clickCount: 1 }); win.webContents.sendInputEvent({ type: 'mouseUp', x: Math.round(r.x), y: Math.round(r.y), button: 'left', clickCount: 1 }); await wait(400); };
  const shot = async (n) => fs.writeFileSync(path.join(outDir, n), (await win.webContents.capturePage()).toPNG());
  const visible = (tabExpr) => run(`(() => { const t = ${tabExpr}; const r = t.el.getBoundingClientRect(); return !!t.el.offsetParent && r.width > 100 ? Math.round(r.width) : 0; })()`);
  try {
    await wait(500);
    win.webContents.send('app:open-path', process.env.DOCSTUDIO_MD);
    await until(`!!(active && active.kind === 'md')`);
    win.webContents.send('app:open-path', process.env.DOCSTUDIO_DOCX);
    await until(`!!(active && active.kind === 'docx' && active.viewRendered)`);
    await wait(500);
    check('page bar has a split view button', await run(`!!document.querySelector('#db-split') && !!document.querySelector('#db-split').offsetParent`));

    // Split: Word on the left, Markdown on the right.
    await run(`(Split.open(tabs.find((t) => t.kind === 'md')), true)`);
    await wait(600);
    const wl = await visible(`tabs.find((t) => t.kind === 'docx')`), wr = await visible(`tabs.find((t) => t.kind === 'md')`);
    check('both documents are shown side by side', wl > 300 && wr > 300 && Math.abs(wl - wr) < 40, { wl, wr });
    check('each half has a header with its name', await run(`[...document.querySelectorAll('#stage .split-head')].filter((h) => h.offsetParent).map((h) => h.querySelector('.split-name').textContent).join('|')`).then((s) => /docx/.test(s) && /\.md/.test(s)));
    check('divider shown', await run(`!document.querySelector('#split-divider').hidden && !!document.querySelector('#split-divider').offsetParent`));
    await shot('split-1.png');

    // Clicking in a half makes it the active document.
    await click(`tabs.find((t) => t.kind === 'md').el.querySelector('.pane-body')`);
    check('clicking the right half makes it active', await until(`active.kind === 'md'`, 3000));
    check('page bar follows the active half', await run(`document.querySelector('#mode-switch').textContent.includes('Preview')`));
    await click(`tabs.find((t) => t.kind === 'docx').el.querySelector('.pane-body')`);
    check('clicking the left half makes it active', await until(`active.kind === 'docx'`, 3000));

    // Editing works in each half.
    await run(`setMode(tabs.find((t) => t.kind === 'docx'), 'edit')`);
    await until(`!!active.editor`);
    check('Word editing in the left half', await run(`active.kind === 'docx' && active.mode === 'edit' && !document.querySelector('#toolbar').hidden`) && (await visible(`tabs.find((t) => t.kind === 'md')`)) > 300);
    await run(`(activate(tabs.find((t) => t.kind === 'md')), setMode(active, 'edit'))`);
    await run(`(active.textarea.focus(), active.textarea.setRangeText('\\nTyped in split view', active.textarea.value.length, active.textarea.value.length, 'end'), active.textarea.dispatchEvent(new Event('input')), true)`);
    check('Markdown editing in the right half', await until(`active.dirty && active.textarea.value.includes('Typed in split view')`, 3000));

    // A third document chosen from the tab bar replaces the half you are working in.
    win.webContents.send('app:open-path', process.env.DOCSTUDIO_XLSX);
    await until(`!!(active && active.kind === 'sheet')`);
    await wait(600);
    check('a newly opened document goes into the active half', await run(`Split.state.right === active.id && Split.sideOf(tabs.find((t) => t.kind === 'docx')) === 'left' && !Split.sideOf(tabs.find((t) => t.kind === 'md'))`));
    check('the spreadsheet is drawn in its half', (await visible(`tabs.find((t) => t.kind === 'sheet')`)) > 300 && await run(`!!active.sheetView.grid.offsetParent`));
    await run(`(activate(tabs.find((t) => t.kind === 'md')), true)`);
    check('choosing a tab from the tab bar fills the active half', await run(`Split.state.right === tabs.find((t) => t.kind === 'md').id`));

    // Swap and resize.
    await run(`(Split.swap(), true)`);
    check('swap sides', await run(`Split.sideOf(tabs.find((t) => t.kind === 'md')) === 'left'`));
    await run(`(Split.swap(), true)`);
    const dr = await run(`(() => { const r = document.querySelector('#split-divider').getBoundingClientRect(); return { x: r.left + 3, y: r.top + 200 }; })()`);
    win.webContents.sendInputEvent({ type: 'mouseDown', x: Math.round(dr.x), y: Math.round(dr.y), button: 'left', clickCount: 1 });
    for (let i = 1; i <= 5; i++) { win.webContents.sendInputEvent({ type: 'mouseMove', x: Math.round(dr.x - i * 40), y: Math.round(dr.y), button: 'left', modifiers: ['leftButtonDown'] }); await wait(40); }
    win.webContents.sendInputEvent({ type: 'mouseUp', x: Math.round(dr.x - 200), y: Math.round(dr.y), button: 'left', clickCount: 1 });
    await wait(400);
    check('drag the divider to resize', await run(`Split.state.ratio < 0.4`), await run(`Split.state.ratio`));
    await shot('split-2.png');

    // Home page hides the halves; back to documents shows them again.
    await run(`(Shell.go('home'), true)`);
    check('Home page covers the split', await run(`document.body.classList.contains('on-page') && !document.querySelector('#split-divider').offsetParent`));
    await run(`(Shell.go('home'), true)`);
    check('back to the documents keeps the split', await run(`!document.body.classList.contains('on-page') && !!document.querySelector('#split-divider').offsetParent`));

    // Closing one half leaves the other on its own (unsaved prompt avoided).
    await run(`(async () => { const t = tabs.find((x) => x.kind === 'md'); t.dirty = false; await closeTab(t); return true; })()`);
    check('closing one half ends the split', await run(`!Split.isOn() && !document.querySelector('#stage').classList.contains('split')`));

    // Split with nothing chosen: the right half offers documents.
    await run(`(activate(tabs.find((t) => t.kind === 'docx')), Split.open(null), true)`);
    check('the empty half offers a choice', await run(`!document.querySelector('#split-picker').hidden && document.querySelectorAll('#split-picker .pick-row').length >= 4`));
    await click(`[...document.querySelectorAll('#split-picker .pick-row')].find((b) => b.textContent.includes('New Markdown'))`);
    check('a new document from the picker fills the right half', await until(`Split.isOn() && Split.state.right === active.id && active.kind === 'md'`, 8000));
    await shot('split-3.png');
    await run(`(Split.close(), true)`);
    check('close split view', await run(`!Split.isOn()`));
    check('split view in the View menu and tab menu', await run(`Menus.allItems().then((l) => l.some((i) => i.path.includes('Split view')))`));
    check('no errors', !errors.filter((e) => !/Autofill|ResizeObserver/.test(e)).length, errors.slice(0, 5));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 500)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
