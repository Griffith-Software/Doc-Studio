'use strict';
// Paste special, insert/delete cells, groups, filter views, slicers, protected ranges, theme, copy sheet:
//   DOCSTUDIO_SRC=<budget.xlsx> DOCSTUDIO_OUT=<dir> electron tests/sheets-extra-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-sheetsx-')));
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
  const V = 'W.sheetView';
  const sel = (ref) => run(`(() => { const g = SheetModel.parseRange(${JSON.stringify(ref)}); const v = ${V}; v.select(g.r1, g.c1); v.select(g.r2, g.c2, true); return true; })()`);
  const answer = (obj) => run(`(DocTools.form = async () => (${JSON.stringify(obj)}), true)`);
  const cell = (a) => run(`(() => { const p = SheetModel.parseAddr(${JSON.stringify(a)}); const c = ${V}.sheet.get(p.r, p.c); return c ? { v: c.value && c.value.code ? c.value.code : c.value, f: c.formula || null } : { v: null, f: null }; })()`);
  const saved = async (part) => { const z = await JSZip.loadAsync(Buffer.from(await run(`${V}.wb.save().then(b => Array.from(new Uint8Array(b)))`))); const f = z.file(part); return f ? f.async('string') : null; };
  try {
    const work = path.join(outDir, 'work.xlsx');
    const other = path.join(outDir, 'other.xlsx');
    fs.copyFileSync(src, work);
    fs.copyFileSync(src, other);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.sheetView)`);
    await run(`(window.W = active, true)`);
    await wait(500);
    await run(`(${V}.opts.error = (m) => { window.lastError = m; }, true)`);
    const menus = await run(`Menus.build().then(ms => Object.fromEntries(ms.map(m => [m.name, m.items.filter(i => !i.sep).map(i => i.label)])))`);
    check('menus have the new items', menus.Edit.includes('Paste special') && menus.Edit.includes('Delete cells') && menus.Insert.includes('Cells') && menus.Insert.includes('Drawing…') && menus.Format.includes('Theme')
      && ['Filter views', 'Add a slicer', 'Protected ranges', 'Group'].every((x) => menus.Data.includes(x)) && menus.Tools.includes('Copy this sheet to'), menus);

    // Scratch data P2:R7 (Name, Team, Score) and formulas pointing at it.
    await run(`(() => { const v = ${V}; const rows = [['Name', 'Team', 'Score'], ['Ann', 'Red', 5], ['Bob', 'Blue', 7], ['Cy', 'Red', 3], ['Dee', 'Green', 9], ['Eve', 'Blue', 4]];
      const items = []; rows.forEach((row, i) => row.forEach((x, j) => items.push({ r: 2 + i, c: 16 + j, patch: { value: x } })));
      items.push({ r: 3, c: 20, patch: { formula: 'R5' } }, { r: 4, c: 20, patch: { formula: 'SUM(R3:R7)' } }, { r: 9, c: 17, patch: { value: 42 } }, { r: 5, c: 20, patch: { formula: 'Q9' } });
      v.apply(items); return true; })()`);

    // Paste special.
    await sel('P2:R3');
    await run(`(${V}.onCopy({ preventDefault() {}, clipboardData: { setData() {} } }), true)`);
    await sel('X2');
    await run(`(${V}.pasteSpecial('transpose'), true)`);
    check('paste transposed', (await cell('X3')).v === 'Team' && (await cell('Y2')).v === 'Ann', [await cell('X3'), await cell('Y2')]);
    await sel('T4');
    await run(`(${V}.onCopy({ preventDefault() {}, clipboardData: { setData() {} } }), true)`);
    await sel('Z2');
    await run(`(${V}.pasteSpecial('values'), true)`);
    const z2 = await cell('Z2');
    check('paste values only', z2.v === 28 && !z2.f, z2);

    // Insert / delete cells with shifting.
    await sel('R5');
    await run(`${V}.shiftCells('down')`);
    check('insert cells shifts down and updates references', (await cell('R6')).v === 3 && (await cell('T3')).f === 'R6' && (await cell('T4')).f === 'SUM(R3:R8)', [await cell('R6'), await cell('T3'), await cell('T4')]);
    await sel('R5');
    await run(`${V}.shiftCells('up')`);
    check('delete cells shifts up', (await cell('R5')).v === 3 && (await cell('T3')).f === 'R5' && (await cell('T4')).f === 'SUM(R3:R7)', [await cell('R5'), await cell('T3'), await cell('T4')]);
    await sel('Q9');
    await run(`${V}.shiftCells('left')`);
    check('deleted reference becomes #REF!', (await cell('T5')).f === '#REF!', await cell('T5'));

    // Group rows 10-12 and columns X:Y; collapse/expand.
    await sel('A10:A12');
    await run(`${V}.groupSel('row', true)`);
    await sel('A11');
    await run(`${V}.collapseGroup('row', true)`);
    let ws = await saved('xl/worksheets/sheet1.xml');
    check('row group saved and collapsed', /<row [^>]*r="11"[^>]*>/.test(ws) && /<row [^>]*r="11"[^>]*(outlineLevel="1"[^>]*hidden="1"|hidden="1"[^>]*outlineLevel="1")/.test(ws) && /outlineLevelRow="1"/.test(ws), (ws.match(/<row [^>]*r="1[0-3]"[^>]*>/g) || []));
    await run(`${V}.collapseGroup('row', false)`);
    await sel('X1:Y1');
    await run(`${V}.groupSel('col', true)`);
    ws = await saved('xl/worksheets/sheet1.xml');
    check('column group saved', /<col [^>]*min="24"[^>]*outlineLevel="1"/.test(ws) && /outlineLevelCol="1"/.test(ws), (ws.match(/<col [^>]*>/g) || []));

    // Filter views.
    await sel('P2');
    await run(`${V}.toggleFilter()`);
    await run(`${V}.applyFilter(SheetModel.parseRange(${V}.sheet.autoFilter.ref), 17, { values: new Set(['Red']), blanks: false })`);
    await run(`${V}.saveFilterView('Red team')`);
    await run(`${V}.applyFilter(SheetModel.parseRange(${V}.sheet.autoFilter.ref), 17, null)`);
    check('filter cleared', !(await run(`!!(${V}.sheet.rows.get(4) || {}).hidden`)));
    await run(`${V}.applyFilterView('Red team')`);
    check('filter view applies', await run(`!!(${V}.sheet.rows.get(4) || {}).hidden && !(${V}.sheet.rows.get(3) || {}).hidden`));
    const custom = await saved('docProps/custom.xml');
    check('filter view kept in the workbook', custom && custom.includes('DocStudio.Sheets') && custom.includes('Red team'), custom && custom.slice(0, 400));

    // Slicer.
    await sel('Q3');
    await run(`${V}.addSlicer()`);
    await until(`!!document.querySelector('.sg-slicer')`, 5000);
    check('slicer panel shown', await run(`document.querySelectorAll('.sg-slicer .sg-slicer-item').length === 3`), await run(`document.querySelectorAll('.sg-slicer-item').length`));
    await run(`([...document.querySelectorAll('.sg-slicer-item')].find(b => b.textContent === 'Blue').click(), true)`);
    await until(`!!(${V}.sheet.rows.get(3) || {}).hidden`, 5000);
    check('slicer filters', await run(`!!(${V}.sheet.rows.get(3) || {}).hidden && !(${V}.sheet.rows.get(4) || {}).hidden`));

    // Protected range.
    await sel('P2:R2');
    await run(`${V}.protectRange()`);
    await run(`(${V}.apply([{ r: 2, c: 16, patch: { value: 'Changed' } }]), true)`);
    check('protected range blocks edits', (await cell('P2')).v === 'Name' && /protected/.test(await run('window.lastError || ""')));
    await sel('P2');
    await run(`${V}.unprotectRange()`);
    await run(`(${V}.apply([{ r: 2, c: 16, patch: { value: 'Person' } }]), true)`);
    check('unprotected again', (await cell('P2')).v === 'Person');

    // Theme.
    await run(`${V}.applyTheme('Ocean')`);
    const theme = await saved('xl/theme/theme1.xml');
    check('theme colours and fonts', /<a:accent1><a:srgbClr val="0F6FC6"\/>/.test(theme) && /<a:minorFont>\s*<a:latin typeface="Segoe UI"/.test(theme));

    // Drawing (the canvas result goes in as a picture).
    await run(`(async () => { const c = document.createElement('canvas'); c.width = 60; c.height = 40; c.getContext('2d').fillRect(0, 0, 60, 40); const b = await (await new Promise(r => c.toBlob(r, 'image/png'))).arrayBuffer(); DocsExtra.drawing = async () => ({ bytes: b, ext: 'png', width: 60, height: 40 }); return true; })()`);
    const before = await run(`${V}.wb.save().then(b => b.byteLength)`);
    await run(`(async () => { const ms = await Menus.build(); const ins = ms.find(m => m.name === 'Insert'); await ins.items.find(i => i.label === 'Drawing…').run(); return true; })()`);
    const zz = await JSZip.loadAsync(Buffer.from(await run(`${V}.wb.save().then(b => Array.from(new Uint8Array(b)))`)));
    check('drawing inserted as a picture', Object.keys(zz.files).some((f) => /^xl\/media\//.test(f)), before);

    // Copy the sheet into another open workbook.
    win.webContents.send('app:open-path', other);
    await until(`!!(active && active !== W && active.sheetView)`);
    await run(`(window.O = active, true)`);
    await run(`W.sheetView.copySheetTo(O.sheetView)`);
    const names = await run(`O.sheetView.wb.sheets.map(s => s.name)`);
    check('sheet copied into the other workbook', names.length === 3 && /\(2\)$/.test(names[2]), names);
    check('copied values and formulas', await run(`(() => { const s = O.sheetView.wb.sheets[2]; const a = s.get(3, 16); const f = s.get(4, 20); return a && a.value === 'Ann' && f && f.formula === 'SUM(R3:R7)'; })()`));

    // Save both and open in Excel later (office-validate).
    await run(`save(W)`);
    await until(`!W.dirty`);
    await run(`save(O)`);
    await until(`!O.dirty`);
    fs.copyFileSync(work, path.join(outDir, 'edited.xlsx'));
    fs.copyFileSync(other, path.join(outDir, 'other-edited.xlsx'));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 700)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
