'use strict';
// Menus, menu search, find & replace, versions, zoom and conversions.
//   DOCSTUDIO_SRC=<folder with notes.md, rich.docx, budget.xlsx, deck.pptx, report.pdf, picture.png> DOCSTUDIO_OUT=<dir>
//   electron tests/shell-test.js
// Conversions are checked through Doc Studio's own routes (Office is skipped even when installed).
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const srcDir = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-shell-')));
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
    if (r && r.__err) throw new Error(`${js.slice(0, 160)}\n${r.__err}`);
    return r;
  };
  const until = async (js, ms = 20000) => { const t = Date.now(); while (Date.now() - t < ms) { if (await run(js)) return true; await wait(200); } return false; };
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  const open = async (name) => {
    const work = path.join(outDir, name);
    fs.copyFileSync(path.join(srcDir, name), work);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.name === ${JSON.stringify(name)} && !document.querySelector('.pane.active .loading'))`);
    await wait(500);
    return work;
  };
  // Convert the active tab with Doc Studio's own route; returns the bytes as a Buffer (or a list for pictures).
  const convert = async (ext) => {
    const r = await run(`Convert.run(active, ${JSON.stringify(ext)}, { office: false }).then((b) => b && b.files ? { files: b.files.map((f) => Array.from(new Uint8Array(f.data))) } : Array.from(new Uint8Array(b)))`);
    return r.files ? r.files.map((f) => Buffer.from(f)) : Buffer.from(r);
  };
  const save = (name, buf) => { fs.writeFileSync(path.join(outDir, name), buf); return buf; };
  const isZip = (b) => b[0] === 0x50 && b[1] === 0x4b;
  const isPdf = (b) => b.subarray(0, 5).toString() === '%PDF-';
  const pdfPages = (b) => (b.toString('latin1').match(/\/Type\s*\/Page[^s]/g) || []).length;
  try {
    // ---------------- menus
    const md = await open('notes.md');
    // Markdown opens in Preview; the editing menus are checked in Edit.
    check('Markdown opens in Preview', await run(`active.mode === 'preview'`));
    await run(`setMode(active, 'edit')`);
    await wait(400);
    const menus = await run(`[...document.querySelectorAll('#menu-names .menu-name')].map(b => b.textContent)`);
    check('menu bar shows File..Help for Markdown', ['File', 'Edit', 'View', 'Insert', 'Format', 'Tools', 'Help'].every((m) => menus.includes(m)), menus);
    check('document title in the menu bar', await run(`document.querySelector('#doc-title').textContent === 'notes.md'`));
    await run(`Menus.show(0)`);
    await wait(300);
    const fileItems = await run(`[...document.querySelectorAll('.menu-panel .menu-label')].map(x => x.textContent)`);
    check('File menu lists convert, versions, details', ['Download / convert to', 'Convert to…', 'Version history…', 'Details', 'Make a copy…'].every((x) => fileItems.includes(x)), fileItems);
    await shot('menu-file.png');
    await run(`Menus.close()`);
    const insertItems = await run(`Menus.build().then(ms => ms.find(m => m.name === 'Insert').items.filter(i => !i.sep).map(i => i.label))`);
    check('toolbar actions appear in Insert/Format', insertItems.length >= 3, insertItems);
    // search the menus
    await run(`(async () => { const s = document.querySelector('#menu-search'); s.focus(); s.value = 'pdf'; await Menus.search('pdf'); return true; })()`);
    await until(`document.querySelectorAll('#menu-results .menu-label').length > 0`, 5000);
    const results = await run(`[...document.querySelectorAll('#menu-results .menu-label')].map(x => x.textContent)`);
    check('menu search finds PDF commands', results.some((x) => /PDF/.test(x)), results);
    await shot('menu-search.png');
    await run(`(() => { const s = document.querySelector('#menu-search'); s.value = ''; s.dispatchEvent(new Event('input')); s.blur(); return true; })()`);

    // ---------------- find & replace (Markdown editor: one undoable step)
    await run(`setMode(active, 'edit')`);
    const before = await run(`active.textarea.value`);
    const word = (before.match(/\b[A-Za-z]{5,}\b/) || ['notes'])[0];
    const count = (before.match(new RegExp(`\\b${word}\\b`, 'gi')) || []).length;
    await run(`(openFind(true), document.querySelector('#find-input').value = ${JSON.stringify(word)}, runFind(), document.querySelector('#replace-input').value = 'REPLACED', replaceAll(), true)`);
    const after = await run(`active.textarea.value`);
    check('replace all in a text editor', (after.match(/REPLACED/g) || []).length >= count && count > 0, { word, count });
    await run(`(closeFind(), command('undo'), true)`);
    check('replace all is one undo step', (await run(`active.textarea.value`)) === before);
    // regex + case
    await run(`(openFind(true), document.querySelector('#find-regex').click(), document.querySelector('#find-input').value = '^#+ (.*)$', runFind(), true)`);
    const regexHits = await run(`find.hits.length`);
    check('regex search', regexHits > 0, regexHits);
    await run(`(document.querySelector('#find-regex').click(), closeFind(), true)`);

    // ---------------- versions
    await run(`(active.textarea.value += '\\nVersion one', active.textarea.dispatchEvent(new Event('input')), save(active))`);
    await wait(400);
    await run(`(active.textarea.value += '\\nVersion two', active.textarea.dispatchEvent(new Event('input')), save(active))`);
    await wait(400);
    const versions = await run(`desk.versions(active.path)`);
    check('each save keeps a version', versions.length >= 2, versions.length);
    const v0 = await run(`desk.readVersion(active.path, ${JSON.stringify(versions[versions.length - 1].id)}).then(b => new TextDecoder().decode(b))`);
    check('older version readable', v0.includes('Version one') && !v0.includes('Version two'));

    // ---------------- zoom
    await run(`zoomTab(active, 1)`);
    check('zoom in enlarges the editor text', await run(`active.zoom > 1 && parseFloat(active.textarea.style.fontSize) > 14`));
    await run(`zoomTab(active, 0)`);

    // ---------------- Markdown conversions
    const mdDocx = save('md.docx', await convert('docx'));
    check('md → docx', isZip(mdDocx) && (await JSZip.loadAsync(mdDocx)).file('word/document.xml'));
    const mdPdf = save('md.pdf', await convert('pdf'));
    check('md → pdf', isPdf(mdPdf) && pdfPages(mdPdf) >= 1, pdfPages(mdPdf));
    const mdOdt = save('md.odt', await convert('odt'));
    const mdOdtXml = await (await JSZip.loadAsync(mdOdt)).file('content.xml').async('string');
    check('md → odt with headings and lists', /<text:h /.test(mdOdtXml) && (/<text:list/.test(mdOdtXml) || !before.match(/^\s*[-*] /m)), mdOdtXml.slice(0, 300));
    const mdRtf = save('md.rtf', await convert('rtf'));
    check('md → rtf', mdRtf.subarray(0, 6).toString() === '{\\rtf1');
    const mdHtml = save('md.html', await convert('html'));
    check('md → html', /<!DOCTYPE html>/.test(mdHtml.toString()) && /<h1|<h2/.test(mdHtml.toString()));
    const mdEpub = save('md.epub', await convert('epub'));
    const ez = await JSZip.loadAsync(mdEpub);
    check('md → epub', Object.keys(ez.files)[0] === 'mimetype' && !!ez.file('OEBPS/content.opf') && !!ez.file('OEBPS/nav.xhtml'));
    await run(`(active.dirty = false, closeTab(active))`);

    // ---------------- Word conversions
    await open('rich.docx');
    const dPdf = save('docx.pdf', await convert('pdf'));
    check('docx → pdf (own pages)', isPdf(dPdf) && pdfPages(dPdf) >= 1, pdfPages(dPdf));
    const dMd = save('docx.md', await convert('md')).toString();
    check('docx → md keeps headings and text', /^#+ /m.test(dMd) && dMd.length > 100, dMd.slice(0, 300));
    const dOdt = save('docx.odt', await convert('odt'));
    check('docx → odt', isZip(dOdt));
    const dTxt = save('docx.txt', await convert('txt')).toString();
    check('docx → txt', dTxt.trim().length > 50);
    await run(`zoomTab(active, 1)`);
    check('zoom on the Word page', await run(`active.viewBox.querySelector('.docx-view').style.zoom !== ''`));
    await run(`zoomTab(active, 0)`);
    await run(`closeTab(active)`);

    // ---------------- spreadsheet conversions
    await open('budget.xlsx');
    const sCsv = save('xlsx.csv', await convert('csv')).toString();
    check('xlsx → csv (current sheet)', sCsv.split('\n').length > 3 && sCsv.includes(','), sCsv.slice(0, 120));
    const sOds = save('xlsx.ods', await convert('ods'));
    check('xlsx → ods', isZip(sOds) && sOds.includes(Buffer.from('opendocument.spreadsheet')));
    const sPdf = save('xlsx.pdf', await convert('pdf'));
    check('xlsx → pdf', isPdf(sPdf));
    const sMd = save('xlsx.md', await convert('md')).toString();
    check('xlsx → markdown table', /^\| .* \|$/m.test(sMd) && /\| --- \|/.test(sMd));
    await run(`closeTab(active)`);

    // ---------------- slides
    await open('deck.pptx');
    const pPdf = save('pptx.pdf', await convert('pdf'));
    check('pptx → pdf, one page per slide', isPdf(pPdf) && pdfPages(pPdf) === 4, pdfPages(pPdf));
    const pPng = await convert('png');
    pPng.forEach((b, i) => save(`slide-${i + 1}.png`, b));
    check('pptx → png per slide', pPng.length === 4 && pPng.every((b) => b.subarray(1, 4).toString() === 'PNG' && b.length > 5000), pPng.map((b) => b.length));
    const pMd = save('pptx.md', await convert('md')).toString();
    check('pptx → md outline', /^## /m.test(pMd) && pMd.includes('Highlights'), pMd.slice(0, 200));
    await run(`closeTab(active)`);

    // ---------------- PDF sources
    await open('report.pdf');
    const fTxt = save('pdf.txt', await convert('txt')).toString();
    check('pdf → txt', fTxt.trim().length > 50, fTxt.slice(0, 100));
    const fDocx = save('pdf.docx', await convert('docx'));
    check('pdf → docx', isZip(fDocx));
    const fPng = await convert('png');
    check('pdf → png per page', fPng.length === 3, fPng.length);
    await run(`closeTab(active)`);

    // ---------------- pictures
    await open('picture.png');
    const iPdf = save('picture.pdf', await convert('pdf'));
    check('png → pdf', isPdf(iPdf));
    const iJpg = await convert('jpg');
    check('png → jpg', iJpg[0] === 0xff && iJpg[1] === 0xd8);
    await run(`closeTab(active)`);
    await shot('done.png');
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 600)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
