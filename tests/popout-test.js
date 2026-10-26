'use strict';
// Several windows: pop a document out (with unsaved edits), side by side, move between windows, drag requests,
// arrange all windows. DOCSTUDIO_MD=<file.md> DOCSTUDIO_DOCX=<file.docx> DOCSTUDIO_OUT=<dir> electron tests/popout-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-popout-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => { checks.push({ name, ok: !!ok, detail }); console.log('...', ok ? 'ok' : 'FAIL', name); };

app.whenReady().then(async () => {
  let main;
  while (!(main = BrowserWindow.getAllWindows()[0])) await wait(100);
  await new Promise((r) => (main.webContents.isLoading() ? main.webContents.once('did-finish-load', r) : r()));
  const errors = [];
  const watch = (w) => w.webContents.on('console-message', (_e, level, msg) => { if (level >= 3) errors.push(msg); });
  watch(main);
  const run = async (w, js) => {
    const r = await w.webContents.executeJavaScript(`(async () => { try { return await (${js}); } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${js.slice(0, 160)}\n${r.__err}`);
    return r;
  };
  const until = async (fn, ms = 20000) => { const t = Date.now(); while (Date.now() - t < ms) { if (await fn()) return true; await wait(200); } return false; };
  const wins = () => BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed());
  try {
    const md = path.join(outDir, 'instructions.md');
    const docx = path.join(outDir, 'report.docx');
    fs.copyFileSync(process.env.DOCSTUDIO_MD, md);
    fs.copyFileSync(process.env.DOCSTUDIO_DOCX, docx);
    await wait(500);
    main.webContents.send('app:open-path', md);
    await until(() => run(main, `!!(active && active.kind === 'md')`));
    main.webContents.send('app:open-path', docx);
    await until(() => run(main, `!!(active && active.kind === 'docx')`));
    await wait(800);
    check('page bar has a pop-out button', await run(main, `!!document.querySelector('#db-popout') && !!document.querySelector('#db-popout').offsetParent`));
    check('tabs can be dragged', await run(main, `[...document.querySelectorAll('#tabs .tab')].every((t) => t.draggable)`));

    // Edit the Word document (unsaved), then pop it out side by side.
    await run(main, `setMode(active, 'edit')`);
    await until(() => run(main, `!!active.editor`));
    await wait(500);
    await run(main, `(() => { const ed = active.editor; const art = ed.articles[0]; const w = document.createTreeWalker(art, NodeFilter.SHOW_TEXT); for (let n = w.nextNode(); n; n = w.nextNode()) { const i = n.data.indexOf('Last paragraph'); if (i < 0) continue; const r = document.createRange(); r.setStart(n, i + 'Last paragraph'.length); r.collapse(true); art.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, ' POPPED-OUT EDIT'); return true; } return false; })()`);
    await until(() => run(main, `active.dirty`), 3000);
    check('document has unsaved edits before popping out', await run(main, `active.dirty`));
    await run(main, `Popout.send(active, 'side')`);
    check('a second window opens', await until(() => wins().length === 2, 15000));
    const second = wins().find((w) => w !== main);
    watch(second);
    await until(() => run(second, `!!(typeof active !== 'undefined' && active && active.kind === 'docx' && active.editor)`), 20000);
    await wait(800);
    check('the new window has the document, in Edit, still unsaved', await run(second, `active.name === 'report.docx' && active.mode === 'edit' && active.dirty`));
    const text = await run(second, `active.editor.build().then(async (b) => { const z = await JSZip.loadAsync(b.bytes); return (await z.file('word/document.xml').async('string')).includes('POPPED-OUT EDIT'); })`);
    check('the unsaved edit came along', text);
    check('it left the first window', await until(() => run(main, `tabs.length === 1 && active.kind === 'md'`), 8000));
    const b1 = main.getBounds(), b2 = second.getBounds();
    check('side by side: each window takes half the screen', Math.abs(b1.width - b2.width) < 40 && b2.x >= b1.x + b1.width - 20, { b1, b2 });
    check('pop-out window is marked', await run(second, `document.body.classList.contains('popout')`));

    // Saving from the new window writes the original file.
    await run(second, `save(active)`);
    await until(() => run(second, `!active.dirty`), 15000);
    const JSZip = require('jszip');
    const saved = await (await JSZip.loadAsync(fs.readFileSync(docx))).file('word/document.xml').async('string');
    check('saving in the new window writes the original file', saved.includes('POPPED-OUT EDIT'));

    // Move the Markdown file into the second window, then back again (the same as dragging the tab across).
    const id2 = second.id;
    await run(main, `Popout.send(active, ${id2})`);
    check('move to another window', await until(() => run(second, `tabs.length === 2 && tabs.some((t) => t.kind === 'md')`), 15000) && await until(() => run(main, `tabs.length === 0`), 8000));
    const id1 = main.id;
    await run(main, `desk.requestTab(${id2}, (await desk.windowId(), 0))`).catch(() => {});
    const mdTabId = await run(second, `tabs.find((t) => t.kind === 'md').id`);
    await run(main, `desk.requestTab(${id2}, ${mdTabId})`);
    check('dragging a tab onto the other window brings it over', await until(() => run(main, `tabs.length === 1 && active.kind === 'md'`), 15000) && await until(() => run(second, `tabs.length === 1`), 8000));

    // A clean file pops out by reopening it from disk; a drop position places the window.
    await run(main, `Popout.send(active, { x: 300, y: 200 })`);
    check('drag out of the window opens a third window there', await until(() => wins().length === 3, 15000));
    const third = wins().find((w) => w !== main && w !== second);
    check('it has the Markdown file', await until(() => run(third, `!!(typeof active !== 'undefined' && active && active.kind === 'md' && active.name === 'instructions.md')`), 15000));
    if (!checks[checks.length - 1].ok) console.log('THIRD', await run(third, `JSON.stringify({ tabs: tabs.map((t) => [t.name, t.kind, t.mode]), loading: !!document.querySelector('.loading'), title: document.title, welcome: !document.querySelector('#welcome').hidden })`), await run(main, `JSON.stringify(tabs.map((t) => [t.name, t.dirty, t.path]))`));

    // Arrange all windows side by side.
    const n = await run(second, `desk.tileWindows()`);
    await wait(500);
    const xs = wins().map((w) => w.getBounds().x).sort((a, b) => a - b);
    check('arrange all windows side by side', n === 3 && xs[1] > xs[0] && xs[2] > xs[1], xs);
    check('other windows are listed for "Move to"', (await run(second, `desk.windows()`)).length === 2);
    check('menus offer the window commands', await run(second, `Menus.allItems().then((l) => ['Open in a new window', 'Open side by side with this window', 'Arrange all windows side by side'].every((x) => l.some((i) => i.label === x)))`));
    check('dialogs in a pop-out window belong to it (no errors)', !errors.filter((e) => !/Autofill|ResizeObserver/.test(e)).length, errors.slice(0, 5));
    // Closing a window with nothing unsaved closes just that window.
    third.close();
    check('closing one window leaves the others', await until(() => wins().length === 2, 8000));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 500)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
