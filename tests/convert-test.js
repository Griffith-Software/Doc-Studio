'use strict';
// Older/other-suite formats edited through conversion:
//   DOCSTUDIO_SRC=<folder with rich.doc, budget.xls, budget.ods, budget.xlsb, deck.ppt, deck.odp, rich.rtf>
//   DOCSTUDIO_OUT=<dir> electron tests/convert-test.js
// Needs Microsoft Office (or LibreOffice) for the Word/PowerPoint parts; spreadsheets fall back to SheetJS.
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const srcDir = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-conv-')));
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
  const until = async (js, ms = 30000) => { const t = Date.now(); while (Date.now() - t < ms) { if (await run(js)) return true; await wait(250); } return false; };
  const open = async (name) => {
    const work = path.join(outDir, name);
    fs.copyFileSync(path.join(srcDir, name), work);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.name === ${JSON.stringify(name)} && !document.querySelector('.pane.active .loading'))`);
    return work;
  };
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  const typeInWord = (text) => run(`(() => { const a = active.editor.articles[0]; const p = a.querySelector('p'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); a.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, ${JSON.stringify(text)}); return true; })()`);
  try {
    const tools = await run('desk.convertTools()');
    console.log('converters:', JSON.stringify(tools));

    // .doc: shown through Word (full layout), edited, saved as .doc again.
    let work = await open('rich.doc');
    check('.doc opens as a Word layout', await run(`active.kind === 'docx' && active.original && active.original.ext === 'doc' && !!document.querySelector('.pane.active .docx-view section.docx')`));
    await shot('doc-view.png');
    await run(`setMode(active, 'edit')`);
    await until(`!!active.editor`);
    await typeInWord('DOC EDIT ');
    await run(`save(active)`);
    await until(`!active.dirty`);
    const docBytes = fs.readFileSync(work);
    check('.doc saved as a real Word 97-2003 file', docBytes[0] === 0xd0 && docBytes[1] === 0xcf && docBytes.length > 10000 && fs.existsSync(work), docBytes.length);
    await run(`closeTab(active)`);

    // .xls: Edit converts through Excel (or SheetJS), cell change, save back as .xls.
    work = await open('budget.xls');
    await run(`setMode(active, 'edit')`);
    await until(`active.kind === 'sheet' && !!active.sheetView`);
    check('.xls opens in the workbook editor', await run(`active.original.ext === 'xls' && active.workbook.sheets.length >= 1`));
    await run(`(active.sheetView.setFromText(1, 1, 'XLS EDIT'), active.sheetView.changed(), true)`);
    fs.writeFileSync(path.join(outDir, 'xls-native.xlsx'), Buffer.from(await run(`active.workbook.save().then(b => Array.from(new Uint8Array(b)))`)));
    console.log('xls edited');
    await run(`save(active)`);
    await until(`!active.dirty`);
    const xls = fs.readFileSync(work);
    check('.xls saved as BIFF (OLE) file', xls[0] === 0xd0 && xls[1] === 0xcf, xls.length);
    await run(`closeTab(active)`);

    // .ods through Excel or SheetJS.
    work = await open('budget.ods');
    await run(`setMode(active, 'edit')`);
    await until(`active.kind === 'sheet' && !!active.sheetView`);
    await run(`(active.sheetView.setFromText(1, 1, 'ODS EDIT'), active.sheetView.changed(), true)`);
    await run(`save(active)`);
    await until(`!active.dirty`);
    const ods = fs.readFileSync(work);
    check('.ods saved as an OpenDocument zip', ods[0] === 0x50 && ods.includes(Buffer.from('application/vnd.oasis.opendocument.spreadsheet')), ods.length);
    await run(`closeTab(active)`);

    // .ppt: converted for viewing (own renderer), edited, saved as .ppt.
    work = await open('deck.ppt');
    check('.ppt shows its slides', await run(`active.kind === 'pptx' && document.querySelectorAll('.pane.active .pp-view .pp-slide').length === 4`));
    await run(`setMode(active, 'edit')`);
    await until(`!!active.pptEditor`);
    await run(`(async () => { const ed = active.pptEditor; const shape = ed.slideEl.querySelector('.pp-own[data-idx="0"]'); ed.editText(shape); shape.querySelector('.pp-r').firstChild.data = 'PPT EDIT'; ed.commitText(); return true; })()`);
    await run(`save(active)`);
    await until(`!active.dirty`);
    const ppt = fs.readFileSync(work);
    check('.ppt saved as PowerPoint 97-2003', ppt[0] === 0xd0 && ppt[1] === 0xcf, ppt.length);
    await run(`closeTab(active)`);

    // .odp through PowerPoint.
    work = await open('deck.odp');
    check('.odp shows real slides', await run(`active.kind === 'pptx' && document.querySelectorAll('.pane.active .pp-view .pp-slide').length >= 4`));
    await run(`closeTab(active)`);

    // Export a .docx to other formats (what Save As uses).
    const fmts = await run(`(async () => { const t = { kind: 'docx' }; return exportChoices(t, 'docx'); })()`);
    check('docx can be saved as odt/rtf/doc/pdf here', ['odt', 'rtf', 'doc', 'pdf'].every((x) => fmts.includes(x)), fmts);
    const sheetFmts = await run(`exportChoices({ kind: 'sheet' }, 'xlsx')`);
    check('workbook can be saved as ods/xls/xlsb', ['ods', 'xls', 'xlsb'].every((x) => sheetFmts.includes(x)), sheetFmts);
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 600)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
