'use strict';
// Two toolbar lines, overflow menu, toolbar settings (place, reorder, menus only, reset), window buttons on the
// action bar, sidebar toggle and compact-mode labels: DOCSTUDIO_OUT=<dir> electron tests/toolbar-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-toolbar-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => { checks.push({ name, ok: !!ok, detail }); console.log('...', ok ? 'ok' : 'FAIL', name); };

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1900, 1000);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const errors = [];
  win.webContents.on('console-message', (_e, level, msg) => { if (level >= 3) errors.push(msg); });
  const run = async (js) => {
    const r = await win.webContents.executeJavaScript(`(async () => { try { return await (${js}); } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${js.slice(0, 160)}\n${r.__err}`);
    return r;
  };
  const until = async (js, ms = 20000) => { const t = Date.now(); while (Date.now() - t < ms) { if (await run(js)) return true; await wait(200); } return false; };
  const where = (label) => run(`(() => { const el = [...document.querySelectorAll('#tb-primary > *, #tool-left > *, #tb-holder > *')].find((t) => Toolbar.labelOf(t) === ${JSON.stringify(label)}); if (!el) return 'missing'; return el.closest('#tb-primary') ? 'line1' : el.closest('#tool-left') ? 'line2' : el.closest('#tb-holder') ? 'hidden' : 'elsewhere'; })()`);
  const lines = () => run(`({ docbar: document.querySelector('#docbar').getBoundingClientRect().height, toolbar: document.querySelector('#toolbar').getBoundingClientRect().height })`);
  try {
    await wait(500);
    run(`localStorage.removeItem('docstudio.toolbar.docx')`);
    check('window buttons on the action bar, right side', await run(`['#db-split', '#db-popout', '#db-full'].every((s) => document.querySelector(s).closest('#ab-right')) && document.querySelector('#ab-right').getBoundingClientRect().right > innerWidth - 40`));
    await run(`(newDocument('docx'), true)`);
    await until(`!!(active && active.editor && active.mode === 'edit')`);
    await wait(800);
    check('main tools on line 1 beside the zoom', (await where('Bold')) === 'line1' && (await where('Font')) === 'line1' && (await where('Align left')) === 'line1' && (await where('Undo')) === 'line1');
    check('secondary tools on line 2', (await where('Strikethrough')) === 'line2' && (await where('Insert table')) === 'line2' && (await where('Page break')) === 'line2');
    check('menus share line 2', await run(`!!document.querySelector('#toolbar #menubar') && !document.querySelector('#menubar').hidden`));
    const h1 = await lines();
    check('no third line (each bar is a single row)', h1.docbar < 70 && h1.toolbar < 60, h1);
    check('toolbar settings gear at the end of line 2', await run(`(() => { const g = document.querySelector('#tb-settings'); const t = document.querySelector('#toolbar').getBoundingClientRect(); return !!g.offsetParent && g.getBoundingClientRect().right > t.right - 40; })()`));
    fs.writeFileSync(path.join(outDir, 'wide.png'), (await win.webContents.capturePage()).toPNG());

    // Narrow window: what doesn't fit goes into the » menu.
    win.setSize(1150, 900);
    await wait(900);
    const h2 = await lines();
    check('narrow window: still two lines', h2.docbar < 70 && h2.toolbar < 60, h2);
    check('narrow window: » menu holds the rest', await run(`!!document.querySelector('#tb-overflow') && document.querySelector('#tb-overflow').isConnected`));
    await run(`(document.querySelector('#tb-overflow').click(), true)`);
    await wait(500);
    const extra = await run(`[...document.querySelectorAll('.menu-panel:not(#menu-results) .menu-label')].map((x) => x.textContent)`);
    check('» menu lists the hidden tools', extra.length > 0, extra);
    await run(`(Menus.close(), true)`);
    fs.writeFileSync(path.join(outDir, 'narrow.png'), (await win.webContents.capturePage()).toPNG());
    win.setSize(1900, 1000);
    await wait(900);

    // Toolbar settings: Strikethrough to line 1 and to the front; Superscript to menus only.
    await run(`(Toolbar.settings(), true)`);
    await wait(400);
    check('settings list every tool', await run(`document.querySelectorAll('.tb-row').length >= 25`));
    await run(`(() => { const row = [...document.querySelectorAll('.tb-row')].find((r) => r.querySelector('.tb-name').textContent === 'Strikethrough'); [...row.querySelectorAll('.tb-seg button')].find((b) => b.textContent === 'Line 1').click(); return true; })()`);
    for (let i = 0; i < 20; i++) await run(`(() => { const row = [...document.querySelectorAll('.tb-row')].find((r) => r.querySelector('.tb-name').textContent === 'Strikethrough'); const up = [...row.querySelectorAll('.tb-move')].find((b) => b.textContent === '↑'); if (!up.disabled) up.click(); return true; })()`);
    await run(`(() => { const row = [...document.querySelectorAll('.tb-row')].find((r) => r.querySelector('.tb-name').textContent === 'Superscript'); [...row.querySelectorAll('.tb-seg button')].find((b) => b.textContent === 'Menus only').click(); return true; })()`);
    await run(`([...document.querySelectorAll('.dialog button')].find((b) => b.textContent === 'Save').click(), true)`);
    await wait(500);
    check('moved tool shows on line 1', (await where('Strikethrough')) === 'line1');
    check('reordered: it comes first', await run(`Toolbar.labelOf(document.querySelector('#tb-primary').firstElementChild) === 'Strikethrough'`));
    check('"Menus only" leaves the toolbar but stays in the menus', (await where('Superscript')) === 'hidden' && await run(`Menus.allItems().then((l) => l.some((i) => i.label === 'Superscript'))`));
    check('choices are remembered per kind of document', await run(`JSON.parse(localStorage.getItem('docstudio.toolbar.docx')).place.Superscript === 'menu'`));
    await run(`(newDocument('docx'), true)`);
    await until(`!!(active && active.editor && active.mode === 'edit' && tabs.length === 2)`);
    await wait(600);
    check('a second Word document uses the same toolbar', (await where('Strikethrough')) === 'line1' && (await where('Superscript')) === 'hidden');
    await run(`(Toolbar.settings(), true)`);
    await wait(300);
    await run(`([...document.querySelectorAll('.dialog button')].find((b) => b.textContent === 'Reset to default').click(), true)`);
    await wait(400);
    check('reset to default', (await where('Strikethrough')) === 'line2' && (await where('Superscript')) === 'line2' && await run(`localStorage.getItem('docstudio.toolbar.docx') === null`));

    // Markdown: a short toolbar, everything on line 1.
    await run(`(newDocument('md'), true)`);
    await until(`!!(active && active.kind === 'md' && active.textarea)`);
    await wait(500);
    check('short toolbars stay on line 1', (await where('Bold')) === 'line1' && (await where('Table')) === 'line1');

    // Sidebar button and compact-mode labels.
    check('sidebar button at the top left', await run(`(() => { const b = document.querySelector('#btn-sidebar').getBoundingClientRect(); return b.left < 60 && b.top < 60; })()`));
    await run(`(document.querySelector('#btn-sidebar').click(), true)`);
    check('sidebar button hides the sidebar', await run(`getComputedStyle(document.querySelector('#sidebar')).display === 'none' && document.querySelector('#btn-sidebar').title === 'Show the sidebar'`));
    check('menu says "Show the sidebar" when it is hidden', await run(`Menus.allItems().then((l) => l.some((i) => i.label === 'Show the sidebar'))`));
    await run(`(document.querySelector('#btn-sidebar').click(), true)`);
    check('and shows it again', await run(`getComputedStyle(document.querySelector('#sidebar')).display !== 'none' && localStorage.getItem('docstudio.sidebar') === 'true'`));
    await run(`(Menus.toggleCompact(), true)`);
    check('compact mode: menu offers to show the sidebar and menus', await run(`Menus.allItems().then((l) => l.some((i) => /^Show the sidebar and menus/.test(i.label)))`));
    check('compact mode: the sidebar button brings them back', await run(`(document.querySelector('#btn-sidebar').click(), !document.body.classList.contains('compact') && getComputedStyle(document.querySelector('#sidebar')).display !== 'none')`));
    check('no errors', !errors.filter((e) => !/Autofill|ResizeObserver/.test(e)).length, errors.slice(0, 5));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 500)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
