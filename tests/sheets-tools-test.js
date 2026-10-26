'use strict';
// Google Sheets-style features: DOCSTUDIO_SRC=<budget.xlsx> DOCSTUDIO_OUT=<dir> electron tests/sheets-tools-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-sheets-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => { checks.push({ name, ok: !!ok, detail }); console.log('...', ok ? 'ok' : 'FAIL', name); };

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1400, 900);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const run = async (js) => {
    const r = await win.webContents.executeJavaScript(`(async () => { try { return await (${js}); } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${js.slice(0, 160)}\n${r.__err}`);
    return r;
  };
  const until = async (js, ms = 20000) => { const t = Date.now(); while (Date.now() - t < ms) { if (await run(js)) return true; await wait(200); } return false; };
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  const V = 'active.sheetView';
  const sel = (ref) => run(`(() => { const g = SheetModel.parseRange(${JSON.stringify(ref)}); const v = ${V}; v.select(g.r1, g.c1); v.select(g.r2, g.c2, true); return true; })()`);
  const answer = (obj) => run(`(DocTools.form = async () => (${JSON.stringify(obj)}), true)`);
  const val = (a) => run(`(() => { const p = SheetModel.parseAddr(${JSON.stringify(a)}); const c = ${V}.sheet.get(p.r, p.c); return c ? (c.value && c.value.code ? c.value.code : c.value) : null; })()`);
  try {
    const work = path.join(outDir, 'work.xlsx');
    fs.copyFileSync(src, work);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.sheetView)`);
    await wait(500);
    const menus = await run(`Menus.build().then(ms => Object.fromEntries(ms.map(m => [m.name, m.items.filter(i => !i.sep).map(i => i.label)])))`);
    check('Data menu', ['Sort range…', 'Data cleanup', 'Split text to columns…', 'Randomize range', 'Column stats', 'Named ranges…', 'Protect sheet', 'Data validation'].every((x) => menus.Data.includes(x)), menus.Data);
    check('Insert menu', ['Chart', 'Pivot table…', 'Image over cells…', 'Function…', 'Checkbox', 'Dropdown…', 'Note…'].every((x) => menus.Insert.includes(x)), menus.Insert);

    // Chart from Item/Qty.
    await sel('A4:B8');
    await run(`${V}.insertChart({ type: 'column', title: 'Quantities' })`);
    await wait(800);
    check('chart drawn in the grid', await run(`!!document.querySelector('.pane.active canvas')`));
    // Picture.
    await sel('J20');
    await run(`(async () => { const c = document.createElement('canvas'); c.width = 80; c.height = 50; const g = c.getContext('2d'); g.fillStyle = '#43a047'; g.fillRect(0, 0, 80, 50); const b = await (await new Promise(r => c.toBlob(r, 'image/png'))).arrayBuffer(); await ${V}.insertPicture(b, 'png', 80, 50); return true; })()`);
    // Note.
    await sel('A4');
    await answer({ text: 'Laptops for the new hires' });
    await run(`${V}.editNote()`);
    check('note shown with a marker', await run(`!!document.querySelector('.pane.active .sg-note-mark')`));
    // Checkboxes and dropdown.
    await sel('J4:J6');
    await run(`${V}.setValidation({ checkbox: true })`);
    check('checkbox cells are FALSE and drawn as boxes', (await val('J4')) === false && await run(`[...document.querySelectorAll('.pane.active .sc-check')].length >= 3`));
    await run(`(() => { ${V}.apply([{ r: 4, c: 10, patch: { value: true } }]); return true; })()`);
    check('checkbox tick', (await val('J4')) === true);
    await sel('K4');
    await run(`${V}.setValidation({ type: 'list', values: ['Open', 'Done'] })`);
    check('dropdown arrow on the active cell', await run(`(${V}.select(4, 11), ${V}.render(true), !!document.querySelector('.pane.active .sg-dd-btn'))`));

    // Data cleanup in a scratch area L2:M6.
    await run(`(() => { const v = ${V}; v.apply([{ r: 2, c: 12, patch: { value: 'Name' } }, { r: 2, c: 13, patch: { value: 'City' } }, { r: 3, c: 12, patch: { value: '  Ann   Lee ' } }, { r: 3, c: 13, patch: { value: 'Oslo' } }, { r: 4, c: 12, patch: { value: 'Bob' } }, { r: 4, c: 13, patch: { value: 'Rome' } }, { r: 5, c: 12, patch: { value: 'Bob' } }, { r: 5, c: 13, patch: { value: 'Rome' } }, { r: 6, c: 12, patch: { value: 'Cy,Dee' } }]); return true; })()`);
    await sel('L3:M6');
    await run(`(${V}.trimWhitespace(), true)`);
    check('trim whitespace', (await val('L3')) === 'Ann Lee');
    await sel('L2:M6');
    await answer({ header: true });
    await run(`${V}.removeDuplicates()`);
    check('remove duplicates', (await val('L5')) === 'Cy,Dee' && (await val('L6')) == null, [await val('L5'), await val('L6')]);
    await sel('L5');
    await answer({ sep: ',', custom: '' });
    await run(`${V}.splitToColumns()`);
    check('split text to columns', (await val('L5')) === 'Cy' && (await val('M5')) === 'Dee');

    // Named range.
    await sel('D4:D8');
    await answer({ name: 'Totals', ref: 'Budget!$D$4:$D$8' });
    await run(`${V}.addNamedRange()`);
    await run(`(${V}.setFromText(12, 12, '=SUM(Totals)'), true)`);
    check('named range in a formula', Math.abs((await val('L12')) - 10372.5) < 0.01, await val('L12'));

    // Pivot table.
    await sel('A2:D8');
    await answer({ rows: '1', cols: '', values: '4', fn: 'SUM' });
    await run(`${V}.pivotTable()`);
    check('pivot table sheet with live totals', await run(`${V}.sheet.name.startsWith('Pivot Table') && ${V}.sheet.get(2, 2) && !!${V}.sheet.get(2, 2).formula`), await run(`${V}.sheet.name`));
    const gt = await run(`(() => { const s = ${V}.sheet; let r = 2; while (s.get(r, 1) && s.get(r, 1).value !== 'Grand total') r++; return s.get(r, 2) ? s.get(r, 2).value : null; })()`);
    check('pivot grand total', Math.abs(gt - 10372.5) < 0.01, gt);
    await run(`${V}.showSheet(0)`);

    // Convert to table (scratch block).
    await run(`(() => { const v = ${V}; const items = []; ['Region', 'Sales'].forEach((t, i) => items.push({ r: 20, c: 1 + i, patch: { value: t } })); [['North', 5], ['South', 7]].forEach((row, k) => row.forEach((x, i) => items.push({ r: 21 + k, c: 1 + i, patch: { value: x } }))); v.apply(items); return true; })()`);
    await sel('A20:B22');
    await run(`${V}.convertToTable()`);
    check('table created', await run(`active.workbook.tables.some(t => t.g.r1 === 20 && t.cols.join() === 'Region,Sales')`));

    // View and protection.
    await run(`${V}.toggleGridlines()`);
    check('gridlines off', await run(`${V}.sheet.showGrid === false`));
    await run(`${V}.toggleGridlines()`);
    await run(`(${V}.toggleShowFormulas(), true)`);
    check('show formulas', await run(`[...document.querySelectorAll('.pane.active .sc span')].some(s => s.textContent.startsWith('=SUM'))`));
    await run(`(${V}.toggleShowFormulas(), true)`);
    await run(`(${V}.setZoom(1.25), true)`);
    check('zoom', await run(`${V}.grid.style.zoom === '1.25'`));
    await run(`(${V}.setZoom(1), true)`);
    await run(`${V}.toggleProtect()`);
    await run(`(window.__err = null, ${V}.opts.error = (m) => { window.__err = m; }, ${V}.setFromText(30, 1, 'blocked'), true)`);
    check('protected sheet blocks edits', (await val('A30')) == null && await run(`/protected/.test(window.__err || '')`));
    await run(`${V}.toggleProtect()`);
    await shot('sheets-1.png');

    // Sheet tabs.
    const n0 = await run(`active.workbook.sheets.length`);
    await run(`${V}.duplicateSheet(0)`);
    check('duplicate sheet', (await run(`active.workbook.sheets.length`)) === n0 + 1 && await run(`active.workbook.sheets[1].name === 'Copy of Budget'`));
    await run(`${V}.setTabColor(1, '#e53935')`);
    check('tab colour', await run(`${V}.tabColorOf(active.workbook.sheets[1]) === '#e53935'`));
    await run(`${V}.moveSheet(1, 1)`);
    check('move sheet', await run(`active.workbook.sheets[2].name === 'Copy of Budget'`));
    await run(`${V}.setSheetState(2, 'hidden')`);
    check('hide sheet', await run(`active.workbook.sheets[2].state === 'hidden' && ![...document.querySelectorAll('.pane.active .sg-tab')].some(b => b.textContent === 'Copy of Budget')`));
    await run(`${V}.setSheetState(2, 'visible')`);
    await run(`${V}.deleteSheet(2, true)`);
    check('delete sheet', (await run(`active.workbook.sheets.length`)) === n0 && !(await run(`active.workbook.sheets.some(s => s.name === 'Copy of Budget')`)), await run(`active.workbook.sheets.map(s => s.name)`));

    // Undo the delete.
    await run(`${V}.undo()`);
    await wait(500);
    check('undo restores the sheet', await run(`active.workbook.sheets.some(s => s.name === 'Copy of Budget')`));

    await run(`save(active)`);
    await until(`!active.dirty`);
    fs.copyFileSync(work, path.join(outDir, 'edited.xlsx'));
    const z = await JSZip.loadAsync(fs.readFileSync(work));
    const names = Object.keys(z.files);
    check('package has chart, picture, comments, vml, table', ['xl/charts/', 'xl/media/', 'xl/comments', 'vmlDrawing', 'xl/tables/'].every((p) => names.some((n) => n.includes(p))), names.filter((n) => /chart|media|comment|vml|table/.test(n)));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 700)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
