'use strict';
// Excel feature test on a workbook written by real Excel (table, chart, CF, data validation, external link,
// dynamic array, frozen panes). DOCSTUDIO_SRC=<excel-made.xlsx> DOCSTUDIO_OUT=<dir> electron tests/xlsx-features-test.js

const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-xlf-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail });

async function parts(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const out = {};
  for (const n of Object.keys(zip.files)) if (!zip.files[n].dir) out[n] = await zip.file(n).async('string');
  return out;
}

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1400, 900);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const run = async (js) => {
    const cut = /^\(.*\)\(\)$/s.test(js.trim()) ? -1 : js.trimEnd().replace(/;$/, '').lastIndexOf(';');
    const body = cut < 0 ? `return await (${js});` : `${js.slice(0, cut + 1)} return await (${js.slice(cut + 1).trim().replace(/;$/, '') || 'undefined'});`;
    const r = await win.webContents.executeJavaScript(`(async () => { try { ${body} } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${js.slice(0, 120)}\n${r.__err}`);
    return r;
  };
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  const val = (a, sheet) => run(`(() => { const s = ${sheet ? `active.workbook.sheet(${JSON.stringify(sheet)})` : 'active.sheetView.sheet'}; const p = SheetModel.parseAddr(${JSON.stringify(a)}); const c = s.get(p.r, p.c); const v = c ? c.value : null; return v && v.code ? v.code : v; })()`);
  const fml = (a) => run(`(() => { const p = SheetModel.parseAddr(${JSON.stringify(a)}); const c = active.sheetView.sheet.get(p.r, p.c); return c ? c.formula : null; })()`);

  const work = path.join(outDir, 'work.xlsx');
  fs.copyFileSync(src, work);
  win.webContents.send('app:open-path', work);
  await wait(3500);
  await shot('1-open.png');
  check('opened', await run(`active.kind === 'sheet'`));
  check('chart drawn', await run(`!!document.querySelector('.pane.active .sg-chart canvas')`));
  check('external link value', (await val('H2')) === 2.5, await val('H2'));
  check('dynamic array values', (await val('L1')) === 1 && (await val('M3')) === 6, [await val('L1'), await val('M3')]);
  check('CF: E2 (25) not highlighted, E3 (40) highlighted', await run(`(() => { const v = active.sheetView; const a = v.cfFor(2, 5, v.sheet.get(2,5)); const b = v.cfFor(3, 5, v.sheet.get(3,5)); return (!a || !a.fill) && b && !!b.fill; })()`));
  check('CF: data bar on Units', await run(`(() => { const v = active.sheetView; const x = v.cfFor(5, 3, v.sheet.get(5,3)); return !!(x && x.bar); })()`));
  check('frozen header row', await run(`active.sheetView.F.rows === 1`));
  // Structured reference typed by the user.
  await run(`active.sheetView.setFromText(4, 8, '=SUM(SalesTable[Revenue])'); true`);
  check('structured reference evaluates', (await val('H4')) === 196.5, await val('H4'));
  await run(`active.sheetView.setFromText(5, 8, '=SEQUENCE(2,2,10)'); true`);
  check('typed dynamic array spills', (await val('H5')) === 10 && (await val('I6')) === 13, [await val('H5'), await val('I6')]);

  // Insert two rows above row 3 (inside the table, CF, chart and SUM ranges).
  await run(`active.sheetView.select(3, 1); active.sheetView.select(4, 1, true); active.sheetView.insertRows(false)`);
  await wait(500);
  check('SUM range grew', (await fml('H1')) === 'SUM(E2:E9)', await fml('H1'));
  check('shared formula moved', (await fml('E5')) === 'C5*D5', await fml('E5'));
  check('total unchanged', (await val('H1')) === 196.5, await val('H1'));
  check('cross-sheet formula updated', (await run(`active.workbook.sheet('Summary').get(2,2).formula`)).includes('Sales!B2:B9'), await run(`active.workbook.sheet('Summary').get(2,2).formula`));
  check('table ref grew', await run(`active.workbook.tables[0].g.r2 === 9`));
  // Delete the two rows again.
  await run(`active.sheetView.select(3, 1); active.sheetView.select(4, 1, true); active.sheetView.deleteRows()`);
  await wait(500);
  check('delete rows restores ranges', (await fml('H1')) === 'SUM(E2:E7)' && (await val('E3')) === 40, [await fml('H1'), await val('E3')]);
  // Insert a column inside the table (before Price).
  await run(`active.sheetView.select(1, 4); active.sheetView.insertCols(false)`);
  await wait(500);
  check('column insert shifts formulas', (await fml('F2')) === 'C2*E2', await fml('F2'));
  check('table got a new column', await run(`active.workbook.tables[0].cols.length === 6 && active.workbook.tables[0].cols[3] !== 'Price'`));
  // Undo the column insert (whole-workbook snapshot).
  await run(`active.sheetView.undo(); true`);
  await wait(800);
  check('undo column insert', (await fml('E2')) === 'C2*D2' && await run(`active.workbook.tables[0].cols.length === 5`), await fml('E2'));
  await run(`active.sheetView.redo(); true`);
  await wait(800);
  check('redo column insert', (await fml('F2')) === 'C2*E2');
  await run(`active.sheetView.undo(); true`);
  await wait(800);

  // Sort the table by Units descending (header detected).
  await run(`active.sheetView.select(3, 3); active.sheetView.sortSelection(false)`);
  check('sorted by Units desc', (await val('C2')) === 12 && (await val('C7')) === 3 && (await val('A2')) === 'West', [await val('C2'), await val('C7'), await val('A2')]);
  check('formulas follow their rows', (await fml('E2')) === 'C2*D2' && (await val('E2')) === 48);
  // Filter: show only Widget rows.
  check('table filter buttons drawn', await run(`document.querySelectorAll('.pane.active .sg-filter-btn').length >= 5`));
  await run(`(() => { const v = active.sheetView; const f = v.filterAt(1, 2); return v.applyFilter(SheetModel.parseRange(f.ref), 2, { values: new Set(['Widget']), blanks: false }, f); })()`);
  await wait(300);
  const hidden = await run(`[2,3,4,5,6,7].map(r => !!(active.sheetView.sheet.rows.get(r) || {}).hidden)`);
  const products = await run(`[2,3,4,5,6,7].map(r => active.sheetView.sheet.get(r,2).value)`);
  check('filter hides non-matching rows', products.every((p, i) => (p === 'Widget') === !hidden[i]), { hidden, products });
  // Hide column G, set row height, freeze.
  await run(`active.sheetView.select(1, 7); active.sheetView.hideCols(true)`);
  await wait(300);
  check('column hidden', await run(`active.sheetView.sheet.isColHidden(7)`));
  // Conditional formatting: add text rule.
  await run(`active.sheetView.select(2, 2); active.sheetView.select(7, 2, true); active.sheetView.addCfRule({ type: 'containsText', attrs: { operator: 'containsText', text: 'Giz' }, formulas: ['NOT(ISERROR(SEARCH("Giz",B2)))'], dxf: { fill: '#C6EFCE', color: '#006100' } })`);
  await wait(300);
  check('new CF rule applies', await run(`(() => { const v = active.sheetView; for (let r = 2; r <= 7; r++) { const c = v.sheet.get(r,2); if (c.value === 'Gizmo') { const x = v.cfFor(r, 2, c); return !!(x && x.fill === '#C6EFCE'); } } return false; })()`));
  // Edit a value: chart follows.
  // Row 3 is a visible (Widget) row after the sort and filter.
  await run(`active.sheetView.setFromText(3, 3, '100'); true`);
  await wait(400);
  const plotted = await run(`(() => { const d = active.sheetView.drawings.find(x => x.chart); return d ? d.chart.data.datasets[0].data : null; })()`);
  const rev3 = await val('E3');
  check('chart follows edits and skips filtered rows', Array.isArray(plotted) && plotted.includes(rev3) && plotted.length === 2, { plotted, rev3 });
  await shot('2-features.png');

  await run(`active.path = active.path.replace('work.xlsx', 'features.xlsx'); save(active)`);
  await wait(800);
  const s = await parts(path.join(outDir, 'features.xlsx'));
  check('table filter saved in the table', /<autoFilter [^>]*ref="A1:E7"[^>]*><filterColumn colId="1"><filters><filter val="Widget"\/>/.test(s['xl/tables/table1.xml']) && !s['xl/worksheets/sheet1.xml'].includes('<autoFilter'));
  check('hidden rows saved', /<row r="2"[^>]*hidden="1"/.test(s['xl/worksheets/sheet1.xml']));
  check('hidden column saved', /<col [^>]*min="7" max="7"[^>]*hidden="1"|<col [^>]*hidden="1"[^>]*min="7"/.test(s['xl/worksheets/sheet1.xml']));
  check('no sheet filter name for a table filter', !s['xl/workbook.xml'].includes('_xlnm._FilterDatabase'));
  check('new CF + dxf saved', s['xl/worksheets/sheet1.xml'].includes('containsText') && s['xl/styles.xml'].includes('C6EFCE'));
  check('calcChain dropped', !s['xl/calcChain.xml'] && !s['[Content_Types].xml'].includes('calcChain'));
  check('external link part kept', !!s['xl/externalLinks/externalLink1.xml']);
  fs.writeFileSync(path.join(outDir, 'features-sheet1.xml'), s['xl/worksheets/sheet1.xml']);

  const summary = { passed: checks.filter((x) => x.ok).length, failed: checks.filter((x) => !x.ok) };
  console.log(JSON.stringify(summary, null, 2));
  app.exit(0);
}).catch((e) => { console.log('ERROR', e.message); console.log(JSON.stringify(checks, null, 1)); app.exit(1); });
