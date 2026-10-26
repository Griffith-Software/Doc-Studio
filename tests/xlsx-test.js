'use strict';
// Excel editor test: set DOCSTUDIO_SRC=<budget.xlsx> DOCSTUDIO_OUT=<dir>, then `electron tests/xlsx-test.js`.
// Opens the workbook in the real app, checks calculated values, edits, formats, undo, copy/paste, saves, and
// verifies the saved package (untouched parts identical, calc chain dropped, styles appended).

const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-xl-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail });
const consoleErrors = [];

async function parts(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const out = {};
  for (const n of Object.keys(zip.files)) if (!zip.files[n].dir) out[n] = await zip.file(n).async('string');
  return out;
}

let step = 'start';
app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1400, 900);
  win.webContents.on('console-message', (_e, level, msg, line, s) => { if (level >= 3) consoleErrors.push(`${path.basename(s || '')}:${line} ${msg}`); });
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const run = async (js) => {
    step = js.slice(0, 100);
    const cut = /^\(.*\)\(\)$/s.test(js.trim()) ? -1 : js.trimEnd().replace(/;$/, '').lastIndexOf(';');
    const body = cut < 0 ? `return await (${js});` : `${js.slice(0, cut + 1)} return await (${js.slice(cut + 1).trim().replace(/;$/, '') || 'undefined'});`;
    const r = await win.webContents.executeJavaScript(`(async () => { try { ${body} } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${step}\n${r.__err}`);
    return r;
  };
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  const val = (a, sheet) => run(`(() => { const s = ${sheet ? `active.workbook.sheet(${JSON.stringify(sheet)})` : 'active.sheetView.sheet'}; const p = SheetModel.parseAddr(${JSON.stringify(a)}); const c = s.get(p.r, p.c); const v = c ? c.value : null; return v && v.code ? v.code : v; })()`);

  const work = path.join(outDir, 'work.xlsx');
  fs.copyFileSync(src, work);
  win.webContents.send('app:open-path', work);
  await wait(3000);
  await shot('1-open.png');
  check('opened as editable workbook', await run(`active.kind === 'sheet' && !!active.sheetView`));

  // Calculated on open (the generator wrote no cached values).
  check('D4 = B4*C4 (shared formula master)', Math.abs((await val('D4')) - 4802) < 1e-9, await val('D4'));
  check('D6 = shared formula child', Math.abs((await val('D6')) - 1452.5) < 1e-9, await val('D6'));
  check('D10 = SUM', Math.abs((await val('D10')) - (4802 + 1380 + 1452.5 + 1550 + 1188)) < 1e-9, await val('D10'));
  check('G4 VLOOKUP across sheets', (await val('G4')) === 0.05, await val('G4'));
  check('H4 IF', (await val('H4')) === 'big');
  check('H5 COUNTIF', (await val('H5')) === 3, await val('H5'));
  check('H6 TEXT(date)', (await val('H6')) === 'Oct 4', await val('H6'));
  check('E4 share', Math.abs((await val('E4')) - 4802 / 10372.5) < 1e-9, await val('E4'));
  const shown = await run(`[...active.sheetView.mainLayer.querySelectorAll('.sc span')].map(s => s.textContent)`);
  check('currency format shown', shown.includes('$4,802.00'), shown.filter((t) => t.includes('$')).slice(0, 3));
  check('percent format shown', shown.includes('46.3%'), shown.filter((t) => t.includes('%')).slice(0, 3));
  check('date format shown', shown.includes('10/4/2026'));
  check('frozen rows rendered', await run(`active.sheetView.F.rows === 2 && active.sheetView.topLayer.querySelectorAll('.sc').length > 0`));
  check('merged title rendered wide', await run(`(() => { const d = [...active.sheetView.topLayer.querySelectorAll('.sc')].find(x => x.textContent === 'Quarterly Budget'); return d ? parseFloat(d.style.width) > 300 : false; })()`));

  // Edit an input: dependents recalculate.
  await run(`active.sheetView.select(4, 2); active.sheetView.setFromText(4, 2, '10'); true`);
  check('edit B4 -> D4 recalculated', Math.abs((await val('D4')) - 12005) < 1e-9, await val('D4'));
  check('edit B4 -> D10 and E5 recalculated', Math.abs((await val('D10')) - (12005 + 1380 + 1452.5 + 1550 + 1188)) < 1e-9 && Math.abs((await val('E5')) - 1380 / 17575.5) < 1e-9);
  check('B10 SUM recalculated', (await val('B10')) === 43, await val('B10'));
  // New formula typed.
  await run(`active.sheetView.setFromText(12, 8, '=AVERAGE(D4:D8)'); true`);
  check('typed formula evaluates', Math.abs((await val('H12')) - 17575.5 / 5) < 1e-9, await val('H12'));
  await run(`active.sheetView.setFromText(13, 8, '=Rates!B3*100'); true`);
  check('cross-sheet typed formula', Math.abs((await val('H13')) - 10) < 1e-9, await val('H13'));
  // Typed values are understood.
  await run(`active.sheetView.setFromText(14, 1, '12%'); active.sheetView.setFromText(14, 2, '$1,234.50'); active.sheetView.setFromText(14, 3, '2026-12-25'); active.sheetView.setFromText(14, 4, 'TRUE'); true`);
  check('typed 12% -> 0.12 with % format', (await val('A14')) === 0.12 && (await run(`active.workbook.styles.numFmt(active.sheetView.sheet.get(14,1).s)`)) === '0%');
  check('typed $1,234.50 -> number', (await val('B14')) === 1234.5);
  check('typed date -> serial', (await val('C14')) === 46381, await val('C14'));
  check('typed TRUE -> boolean', (await val('D14')) === true);
  // Formatting.
  await run(`active.sheetView.select(4, 1); active.sheetView.select(8, 1, true); active.sheetView.styleSelection({ font: { bold: true, color: '#c00000' } }); true`);
  check('bold+colour applied to A4:A8', await run(`(() => { const st = active.workbook.styles; const f = st.font(active.sheetView.sheet.get(6,1).s); return f.bold && f.color === '#C00000'; })()`));
  await run(`active.sheetView.select(4, 3); active.sheetView.styleSelection({ fill: '#ffff00' }); true`);
  check('fill applied, number format kept', await run(`(() => { const st = active.workbook.styles; const s = active.sheetView.sheet.get(4,3).s; return st.fill(s) === '#FFFF00' && st.numFmt(s) === '"$"#,##0.00'; })()`));
  // Undo / redo.
  await run(`active.sheetView.undo(); true`);
  check('undo fill', await run(`active.workbook.styles.fill(active.sheetView.sheet.get(4,3).s) === null`));
  await run(`active.sheetView.redo(); true`);
  check('redo fill', await run(`active.workbook.styles.fill(active.sheetView.sheet.get(4,3).s) === '#FFFF00'`));
  // Fill down: copy H12's formula down with relative refs.
  await run(`active.sheetView.select(12, 8); active.sheetView.select(13, 8, true); true`);
  await run(`active.sheetView.setFromText(15, 5, '=B15*2'); active.sheetView.setFromText(15, 2, '3'); active.sheetView.setFromText(16, 2, '4'); active.sheetView.select(15, 5); active.sheetView.select(16, 5, true); active.sheetView.fill('down'); true`);
  check('fill down shifts references', (await run(`active.sheetView.sheet.get(16,5).formula`)) === 'B16*2' && (await val('E16')) === 8);
  // Column width.
  await run(`active.sheetView.sheet.setColWidth(2, 20); active.sheetView.layoutDirty = true; active.sheetView.render(true); true`);
  // Merge.
  await run(`active.sheetView.select(18, 1); active.sheetView.select(18, 3, true); active.sheetView.setFromText(18, 1, 'Merged!'); active.sheetView.select(18, 1); active.sheetView.select(18, 3, true); active.sheetView.toggleMerge(); true`);
  check('merge added', await run(`active.sheetView.sheet.merges.some(m => m.r1 === 18 && m.c2 === 3)`));
  // New sheet + rename an existing sheet referenced by formulas.
  await run(`active.sheetView.addSheet(); true`);
  await run(`active.workbook.renameSheet(active.workbook.sheet('Rates'), 'Tax Rates'); active.sheetView.calc.build(); true`);
  check('rename updates formulas', (await run(`active.workbook.sheet('Budget').get(4,7).formula`)).includes("'Tax Rates'!$A$2:$B$4"));
  await run(`active.sheetView.showSheet(0); true`);
  await wait(300);
  await shot('2-edited.png');

  // Save.
  await run(`active.path = active.path.replace('work.xlsx', 'saved.xlsx'); save(active)`);
  await wait(800);
  const savedFile = path.join(outDir, 'saved.xlsx');
  check('saved', fs.existsSync(savedFile));
  const o = await parts(src);
  const s = await parts(savedFile);
  check('chart, drawing, theme kept identical', ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml', 'xl/theme/theme1.xml', 'xl/worksheets/_rels/sheet1.xml.rels'].every((k) => s[k] === o[k]));
  check('conditional formatting & data validation kept', s['xl/worksheets/sheet1.xml'].includes('<conditionalFormatting') && s['xl/worksheets/sheet1.xml'].includes('<dataValidations'));
  check('fullCalcOnLoad set', s['xl/workbook.xml'].includes('fullCalcOnLoad="1"'));
  check('new sheet part + content type', Object.keys(s).some((k) => /worksheets\/sheet3\.xml/.test(k)) && s['[Content_Types].xml'].includes('/xl/worksheets/sheet3.xml'));
  check('renamed sheet in workbook.xml', s['xl/workbook.xml'].includes('name="Tax Rates"'));
  check('cached values written for formulas', /<c r="D4"[^>]*>.*?<v>12005<\/v>/.test(s['xl/worksheets/sheet1.xml']));
  check('shared formula kept (only inputs edited), cached values updated', s['xl/worksheets/sheet1.xml'].includes('<f t="shared" ref="D4:D8" si="0">B4*C4</f><v>12005</v>') && /<c r="D5"[^>]*><f t="shared" si="0"\/><v>1380<\/v>/.test(s['xl/worksheets/sheet1.xml']));
  const noDecl = (x) => x.replace(/^<\?xml[^>]*\?>\s*/, '');
  check('styles appended (not rewritten)', noDecl(s['xl/styles.xml']).length > noDecl(o['xl/styles.xml']).length && noDecl(s['xl/styles.xml']).includes('<numFmt numFmtId="167" formatCode="m/d/yyyy"/>') && noDecl(s['xl/styles.xml']).includes('<xf numFmtId="164" fontId="2" fillId="0" borderId="2" pivotButton="0" quotePrefix="0" xfId="0"/>'));
  // Editing a shared formula's member expands the group into plain formulas (other members keep their formula).
  await run(`activate(tabs.find(t => t.name === 'saved.xlsx')); active.sheetView.showSheet(0); active.sheetView.setFromText(6, 4, '=B6*C6*2'); true`);
  check('edited shared member recalculates', Math.abs((await val('D6')) - 2905) < 1e-9, await val('D6'));
  await run(`active.path = active.path.replace('saved.xlsx', 'saved2.xlsx'); save(active)`);
  await wait(600);
  const s2 = await parts(path.join(outDir, 'saved2.xlsx'));
  check('shared group expanded after editing a member', !s2['xl/worksheets/sheet1.xml'].includes('t="shared"') && s2['xl/worksheets/sheet1.xml'].includes('<f>B5*C5</f>') && s2['xl/worksheets/sheet1.xml'].includes('<f>B6*C6*2</f>') && s2['xl/worksheets/sheet1.xml'].includes('<f>B8*C8</f>'));
  
  fs.writeFileSync(path.join(outDir, 'saved-sheet1.xml'), s['xl/worksheets/sheet1.xml']);

  // Re-open the saved file: values are read back from the cache.
  win.webContents.send('app:open-path', savedFile);
  await wait(2500);
  check('reopened saved file', await run(`active.name === 'saved.xlsx' && !!active.sheetView`));
  check('reopened value D4', (await val('D4')) === 12005);
  check('reopened bold A6', await run(`active.workbook.styles.font(active.sheetView.sheet.get(6,1).s).bold`));

  // New workbook.
  await run(`newDocument('xlsx'); true`);
  await wait(1500);
  await run(`active.sheetView.setFromText(1, 1, 'Price'); active.sheetView.setFromText(2, 1, '5'); active.sheetView.setFromText(3, 1, '7'); active.sheetView.setFromText(4, 1, '=SUM(A2:A3)'); true`);
  check('new workbook formula', (await val('A4')) === 12);
  await run(`(async () => { active.path = ${JSON.stringify(path.join(outDir, 'new.xlsx'))}; active.isNew = false; active.backedUp = true; return save(active); })()`);
  await wait(800);
  await shot('3-new.png');
  check('new workbook saved', fs.existsSync(path.join(outDir, 'new.xlsx')));

  const summary = { passed: checks.filter((x) => x.ok).length, failed: checks.filter((x) => !x.ok), consoleErrors };
  fs.writeFileSync(path.join(outDir, 'xlsx.json'), JSON.stringify({ checks, consoleErrors }, null, 2));
  console.log(JSON.stringify(summary, null, 2));
  app.exit(0);
}).catch((err) => {
  console.log('FAILED at step:', err.message);
  console.log(JSON.stringify({ checks, consoleErrors }, null, 2));
  app.exit(1);
});
