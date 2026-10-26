'use strict';
// Google Docs-style Word features: DOCSTUDIO_SRC=<rich.docx> DOCSTUDIO_OUT=<dir> electron tests/docs-tools-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-docs-')));
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
  // Put the caret in (or select) text in the body.
  const caret = (text, len, atEnd) => run(`(() => {
    const ed = active.editor; const art = ed.articles[0];
    const w = document.createTreeWalker(art, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) { const i = n.data.indexOf(${JSON.stringify(text)}); if (i < 0) continue;
      const r = document.createRange(); const at = i + (${atEnd ? 'true' : 'false'} ? ${JSON.stringify(text)}.length : 0);
      r.setStart(n, at); r.setEnd(n, at + (${len || 0})); art.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); return true; }
    return false; })()`);
  const xml = async (part) => { const z = await JSZip.loadAsync(Buffer.from(await run(`active.editor.build().then(b => Array.from(new Uint8Array(b.bytes)))`))); const f = z.file(part); return f ? f.async('string') : null; };
  try {
    const work = path.join(outDir, 'work.docx');
    fs.copyFileSync(src, work);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.kind === 'docx')`);
    await run(`setMode(active, 'edit')`);
    await until(`!!active.editor`);
    await wait(500);
    const menus = await run(`Menus.build().then(ms => Object.fromEntries(ms.map(m => [m.name, m.items.filter(i => !i.sep).map(i => i.label)])))`);
    check('Insert menu has Docs items', ['Footnote…', 'Comment…', 'Page numbers…', 'Table of contents', 'Special characters…', 'Headers & footers'].every((x) => menus.Insert.includes(x)), menus.Insert);
    check('Format menu has spacing/table/page setup', ['Line & paragraph spacing', 'Table', 'Page setup…', 'Text'].every((x) => menus.Format.includes(x)), menus.Format);

    // Footnote at the end of "intro text."
    await caret('intro text.', 0, true);
    await run(`active.editor.insertFootnote('A footnote from Doc Studio')`);
    let doc = await xml('word/document.xml');
    let notes = await xml('word/footnotes.xml');
    check('footnote reference in the text', /<w:footnoteReference w:id="\d+"\/>/.test(doc));
    check('footnote text in footnotes.xml', notes && notes.includes('A footnote from Doc Studio'));
    check('marker removed from text', !/[-]/.test(doc));
    check('footnote shown after redraw', await run(`active.editor.stage.textContent.includes('A footnote from Doc Studio')`));

    // Comment around a selected word.
    await caret('Untouched', 9);
    await run(`active.editor.insertComment('Please check', 'Tester')`);
    doc = await xml('word/document.xml');
    const comments = await xml('word/comments.xml');
    check('comment range around the selection', /<w:commentRangeStart w:id="(\d+)"\/>.*Untouched.*<w:commentRangeEnd w:id="\1"\/>/s.test(doc), doc.match(/commentRange[^>]*>/g));
    check('comment text and author', comments && comments.includes('Please check') && comments.includes('w:author="Tester"'));
    check('comments list', (await run(`active.editor.comments()`)).some((c) => c.text === 'Please check'));

    // Bookmark, horizontal line, spacing.
    await caret('Last paragraph', 0);
    await run(`active.editor.insertBookmark('Ending')`);
    check('bookmark added', (await run(`active.editor.bookmarks()`)).includes('Ending'));
    await caret('Last paragraph', 0);
    await run(`active.editor.insertHorizontalLine()`);
    doc = await xml('word/document.xml');
    check('horizontal line paragraph', /<w:pBdr><w:bottom w:val="single"/.test(doc));
    await caret('Last paragraph', 0);
    await run(`(active.editor.setSpacing(active.editor.selectedParagraphs(), { line: 2, after: 480 }), true)`);
    doc = await xml('word/document.xml');
    check('double spacing written', /<w:spacing[^>]*w:line="480"[^>]*\/>/.test(doc) && /w:after="480"/.test(doc));

    // Table rows and columns (the document has a 3x3 table).
    await caret('r1c1', 0);
    await run(`active.editor.tableOp('rowBelow')`);
    await caret('r1c1', 0);
    await run(`active.editor.tableOp('colRight')`);
    doc = await xml('word/document.xml');
    const tbl = doc.match(/<w:tbl>.*?<\/w:tbl>/s)[0];
    check('row and column inserted', (tbl.match(/<w:tr[ >]/g) || []).length === 4 && (tbl.match(/<w:gridCol /g) || []).length === 4, { rows: (tbl.match(/<w:tr[ >]/g) || []).length, cols: (tbl.match(/<w:gridCol /g) || []).length });
    await caret('r2c1', 0);
    await run(`active.editor.tableOp('deleteRow')`);
    doc = await xml('word/document.xml');
    check('row deleted', !doc.includes('r2c1'));

    // Page numbers in the footer, page setup.
    await run(`active.editor.insertPageNumbers({ where: 'footer', align: 'center', withTotal: true })`);
    const rels = await xml('word/_rels/document.xml.rels');
    const footerTarget = (rels.match(/Target="(footer\d*\.xml)"/) || [])[1];
    const footer = footerTarget && await xml(`word/${footerTarget}`);
    check('page X of Y in the footer', footer && /PAGE/.test(footer) && /NUMPAGES/.test(footer), footerTarget);
    await run(`active.editor.pageSetup({ width: 11906, height: 16838, orientation: 'landscape', margins: { top: 720, right: 720, bottom: 720, left: 720 }, color: '#fff8e1', columns: 2 })`);
    doc = await xml('word/document.xml');
    const settings = await xml('word/settings.xml');
    check('landscape A4 with margins and 2 columns', /<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"\/>/.test(doc) && /w:top="720"/.test(doc) && /<w:cols [^>]*w:num="2"/.test(doc), doc.match(/<w:sectPr.*?<\/w:sectPr>/s)?.[0]);
    check('page colour', /<w:background w:color="FFF8E1"\/>/.test(doc) && /displayBackgroundShape/.test(settings));
    await shot('docs-1.png');

    // Table of contents (from the Title/Heading paragraphs) + updateFields.
    await caret('Georgia', 0);
    await run(`active.editor.insertTableOfContents()`);
    doc = await xml('word/document.xml');
    check('TOC field inserted', /<w:instrText[^>]*> TOC \\o "1-3" \\h \\z \\u <\/w:instrText>/.test(doc));
    check('Word asked to update fields', /<w:updateFields w:val="true"\/>/.test(await xml('word/settings.xml')));

    // Undo the TOC as one step.
    await run(`active.editor.undo()`);
    doc = await xml('word/document.xml');
    check('undo removes the table of contents', !/ TOC /.test(doc));

    // Generic tools: change case, special character, checklist.
    await caret('Last paragraph', 4);
    await run(`(DocTools.changeCase('upper'), true)`);
    check('change case keeps the rest', await run(`active.editor.articles[0].textContent.includes('LAST paragraph')`));
    await caret('LAST paragraph', 0);
    await run(`(DocTools.toggleChecklist(), true)`);
    check('checklist box added', await run(`[...active.editor.articles[0].querySelectorAll('p')].some(p => p.textContent.startsWith('☐ LAST'))`));
    await caret('LAST paragraph', 0, true);
    await run(`(DocTools.insertText('→'), true)`);
    check('special character inserted', await run(`active.editor.articles[0].textContent.includes('→')`));

    // Save and keep for Word.
    await run(`save(active)`);
    await until(`!active.dirty`);
    fs.copyFileSync(work, path.join(outDir, 'edited.docx'));
    const z = await JSZip.loadAsync(fs.readFileSync(work));
    const ct = await z.file('[Content_Types].xml').async('string');
    check('content types for new parts', ct.includes('/word/footnotes.xml') || !z.file('word/footnotes.xml'));
    await shot('docs-2.png');
    await run(`closeTab(active)`);

    // New document: header/footer creation and page numbers on a file that has none.
    await run(`newDocument('docx')`);
    await until(`!!(active && active.editor)`);
    await run(`(() => { const a = active.editor.articles[0]; const p = a.querySelector('p'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); a.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, 'New document body'); return true; })()`);
    await run(`active.editor.addHeaderFooter('header')`);
    await run(`active.editor.insertPageNumbers({ where: 'footer', align: 'right' })`);
    const nb = Buffer.from(await run(`active.editor.build().then(b => Array.from(new Uint8Array(b.bytes)))`));
    fs.writeFileSync(path.join(outDir, 'new-hf.docx'), nb);
    const nz = await JSZip.loadAsync(nb);
    const nd = await nz.file('word/document.xml').async('string');
    check('header and footer created and referenced', /<w:headerReference w:type="default"/.test(nd) && /<w:footerReference w:type="default"/.test(nd) && Object.keys(nz.files).some((f) => /word\/header\d\.xml/.test(f)) && Object.keys(nz.files).some((f) => /word\/footer\d\.xml/.test(f)));
    await run(`(active.dirty = false, closeTab(active))`);
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 700)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
