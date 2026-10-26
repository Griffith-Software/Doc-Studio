'use strict';
// Applies one spreadsheet operation per copy of DOCSTUDIO_SRC and saves each to DOCSTUDIO_OUT/<op>.xlsx, so each
// result can be checked in Excel on its own. electron tests/xlsx-bisect.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-bis-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');
const wait = (ms) => new Promise((r) => setTimeout(r, ms));

const OPS = {
  edit: `active.sheetView.setFromText(2, 3, '11')`,
  structref: `active.sheetView.setFromText(4, 8, '=SUM(SalesTable[Revenue])')`,
  spill: `active.sheetView.setFromText(5, 8, '=SEQUENCE(2,2,10)')`,
  insrows: `(async () => { active.sheetView.select(3, 1); active.sheetView.select(4, 1, true); await active.sheetView.insertRows(false); })()`,
  delrows: `(async () => { active.sheetView.select(3, 1); await active.sheetView.deleteRows(); })()`,
  inscol: `(async () => { active.sheetView.select(1, 4); await active.sheetView.insertCols(false); })()`,
  delcol: `(async () => { active.sheetView.select(1, 4); await active.sheetView.deleteCols(); })()`,
  sort: `(async () => { active.sheetView.select(3, 3); await active.sheetView.sortSelection(false); })()`,
  filter: `(async () => { const v = active.sheetView; const f = v.filterAt(1, 2); await v.applyFilter(SheetModel.parseRange(f.ref), 2, { values: new Set(['Widget']), blanks: false }, f); })()`,
  sheetfilter: `(async () => { const v = active.sheetView; v.setFromText(20, 1, 'Name'); v.setFromText(20, 2, 'Score'); v.setFromText(21, 1, 'Ann'); v.setFromText(21, 2, '5'); v.setFromText(22, 1, 'Bo'); v.setFromText(22, 2, '9'); v.select(21, 1); await v.toggleFilter(); await v.applyFilter(v.filterRange(), 1, { values: new Set(['Bo']), blanks: false }); })()`,
  tablefilteroff: `(async () => { const v = active.sheetView; v.select(3, 2); await v.toggleFilter(); })()`,
  hidecol: `(async () => { active.sheetView.select(1, 7); await active.sheetView.hideCols(true); })()`,
  cf: `(async () => { active.sheetView.select(2, 2); active.sheetView.select(7, 2, true); await active.sheetView.addCfRule({ type: 'containsText', attrs: { operator: 'containsText', text: 'Giz' }, formulas: ['NOT(ISERROR(SEARCH("Giz",B2)))'], dxf: { fill: '#C6EFCE', color: '#006100' } }); })()`,
  freeze: `(async () => { active.sheetView.select(3, 2); await active.sheetView.toggleFreeze(); await active.sheetView.toggleFreeze(); })()`,
  rowheight: `(async () => { active.sheetView.select(3, 2); active.sheetView.sheet.setRowHeight(3, 30); })()`,
  newsheet: `(async () => { active.sheetView.addSheet(); })()`,
  merge: `(async () => { active.sheetView.select(10, 1); active.sheetView.select(10, 3, true); active.sheetView.toggleMerge(); })()`,
};

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const only = (process.env.OPS || '').split(',').filter(Boolean);
  for (const [name, js] of Object.entries(OPS)) {
    if (only.length && !only.includes(name)) continue;
    const file = path.join(outDir, `${name}.xlsx`);
    fs.copyFileSync(src, file);
    win.webContents.send('app:open-path', file);
    await wait(2500);
    const r = await win.webContents.executeJavaScript(`(async () => { try { await ${js}; await new Promise(r => setTimeout(r, 300)); active.backedUp = true; await save(active); const n = active.name; await closeTab(active); return 'ok ' + n; } catch (e) { return 'ERR ' + e.message; } })()`);
    console.log(name, r);
  }
  app.exit(0);
});
