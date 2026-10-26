'use strict';
// Suggesting mode, review, equations, pictures, dropdowns, lists, borders, ruler indents, section breaks, styles:
//   DOCSTUDIO_SRC=<rich.docx> DOCSTUDIO_OUT=<dir> electron tests/docs-extra-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-docsx-')));
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
  const caret = (text, len, atEnd) => run(`(() => {
    const ed = active.editor; const art = ed.articles[0];
    const w = document.createTreeWalker(art, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) { const i = n.data.indexOf(${JSON.stringify(text)}); if (i < 0) continue;
      const r = document.createRange(); const at = i + (${atEnd ? 'true' : 'false'} ? ${JSON.stringify(text)}.length : 0);
      r.setStart(n, at); r.setEnd(n, at + (${len || 0})); art.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); return true; }
    return false; })()`);
  const xml = async (part) => { const z = await JSZip.loadAsync(Buffer.from(await run(`active.editor.build().then(b => Array.from(new Uint8Array(b.bytes)))`))); const f = z.file(part); return f ? f.async('string') : null; };
  const answer = (obj) => run(`(DocTools.form = async () => (${JSON.stringify(obj)}), true)`);
  const input = (type, data) => run(`(() => { const a = active.editor.articles[0]; a.dispatchEvent(new InputEvent('beforeinput', { inputType: ${JSON.stringify(type)}, data: ${JSON.stringify(data || null)}, cancelable: true, bubbles: true })); return true; })()`);
  try {
    const work = path.join(outDir, 'work.docx');
    fs.copyFileSync(src, work);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.kind === 'docx')`);
    await run(`setMode(active, 'edit')`);
    await until(`!!active.editor`);
    await wait(400);

    // Suggesting: typed text becomes an insertion, deleted text a deletion.
    await run(`(DocsExtra.setSuggesting(active, true), true)`);
    check('suggesting switch in the toolbar', await run(`active.tabSuggestCheck = active.suggestSel && active.suggestSel.value === 'suggest'`));
    await caret('Last paragraph', 0, true);
    await input('insertText', ' (suggested)');
    await caret('Untouched', 9);
    await input('deleteContentBackward');
    let doc = await xml('word/document.xml');
    check('insertion saved as w:ins', /<w:ins w:id="\d+" w:author="[^"]+" w:date="[^"]+"><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t xml:space="preserve"> \(suggested\)<\/w:t><\/w:r><\/w:ins>/s.test(doc), (doc.match(/<w:ins[\s\S]{0,300}/) || [''])[0]);
    check('deletion saved as w:del with delText', /<w:del w:id="\d+"[^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:delText xml:space="preserve">Untouched<\/w:delText><\/w:r><\/w:del>/s.test(doc));
    const changes = await run(`active.editor.trackedChanges()`);
    check('review lists both suggestions', changes.length >= 2 && changes.some((c) => c.type === 'ins') && changes.some((c) => c.type === 'del'), changes);
    await run(`(DocsExtra.setSuggesting(active, false), true)`);
    const insId = changes.find((c) => c.type === 'ins').id;
    await run(`active.editor.reviewChanges(${JSON.stringify(insId)}, true)`);
    doc = await xml('word/document.xml');
    check('accept keeps the inserted text', !/<w:ins /.test(doc) && doc.includes(' (suggested)'));
    await run(`active.editor.reviewChanges(null, false)`);
    doc = await xml('word/document.xml');
    check('reject all restores the deleted text', !/<w:del /.test(doc) && /<w:t[^>]*>Untouched/.test(doc));

    // Equation.
    await caret('Last paragraph', 0, true);
    await run(`active.editor.insertEquation('x = \\\\frac{-b \\\\pm \\\\sqrt{b^2 - 4ac}}{2a}')`);
    doc = await xml('word/document.xml');
    check('equation written as Word math', /<m:oMath[ >]/.test(doc) && /<m:f>/.test(doc) && /<m:rad>/.test(doc) && /<m:sSup>/.test(doc));

    // Lists: Roman numerals, restart.
    await caret('Alpha item', 0);
    await run(`DocsExtra.listStyle(active.editor, 'roman')`);
    const numbering = await xml('word/numbering.xml');
    check('roman numeral list', /lowerRoman/.test(numbering));

    // Borders & shading, ruler indents (through the paragraph data the dialogs set).
    await caret('Last paragraph', 0);
    await answer({ border: 'box', color: '#c00000', width: '12', shade: '#fff2cc' });
    await run(`(async () => { await DocsExtra.listStyle; return true; })()`);
    await run(`(async () => { DocTools.restore = () => true; const ed = active.editor; const p = ed.caretParagraph(); p.dataset.ddBorder = JSON.stringify({ border: { color: '#c00000', size: 12 }, shade: '#FFF2CC' }); p.dataset.ddInd = JSON.stringify({ left: 720, right: 360, firstLine: 360 }); ed.opts.onChange(); return true; })()`);
    doc = await xml('word/document.xml');
    check('paragraph border and shading', /<w:pBdr><w:top w:val="single" w:sz="12" w:space="4" w:color="c00000"\/>/.test(doc) && /<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"\/>/.test(doc));
    check('ruler indents', /<w:ind w:left="720" w:right="360" w:firstLine="360"\/>/.test(doc), (doc.match(/<w:ind [^>]*>/g) || []).slice(-3));
    await run(`(delete DocTools.restore, true)`);

    // Section break.
    await caret('Last paragraph', 0);
    await run(`active.editor.insertSectionBreak('nextPage')`);
    doc = await xml('word/document.xml');
    check('section break', /<w:pPr><w:sectPr[ >][\s\S]*?<w:type w:val="nextPage"\/>/.test(doc));

    // Picture: size, alt text, crop.
    const img = await run(`(() => { const im = active.editor.stage.querySelector('article img'); if (!im) return null; im.dispatchEvent(new MouseEvent('mousedown', { bubbles: true })); return !!active.editor.selectedImage; })()`);
    check('click selects a picture (with a resize handle)', img && await run(`!!document.querySelector('.dd-img-handle')`));
    await run(`(async () => { const ed = active.editor; const info = ed.imageInfo(ed.selectedImage); await ed.setImage({ id: info.id, cx: info.cx * 1.5, cy: info.cy * 1.5, alt: 'Blue frame', crop: { l: 10, t: 0, r: 5, b: 0 } }); return true; })()`);
    doc = await xml('word/document.xml');
    check('picture resized, alt text and crop', /descr="Blue frame"/.test(doc) && /<a:srcRect l="10000" r="5000"\/>/.test(doc));

    // Dropdown chip.
    await caret('Last paragraph', 0, true);
    await run(`active.editor.insertDropdown(['Draft', 'Final'])`);
    doc = await xml('word/document.xml');
    check('dropdown content control', /<w:dropDownList><w:listItem w:displayText="Draft" w:value="Draft"\/><w:listItem w:displayText="Final"/.test(doc));
    check('chip drawn and clickable', await run(`!!active.editor.stage.querySelector('.dd-chip')`));

    // Update style to match.
    await caret('Last paragraph', 4);
    await run(`(document.execCommand('styleWithCSS', false, true), document.execCommand('bold'), true)`);
    await caret('Last paragraph', 2);
    await run(`active.editor.updateStyleToMatch(active.editor.caretParagraph())`);
    const styles = await xml('word/styles.xml');
    check('style updated to match', /<w:style [^>]*w:styleId="(Normal|[^"]+)"[\s\S]*?<w:rFonts w:ascii=/.test(styles));

    // Chart and drawing as pictures.
    await answer({ type: 'bar', title: 'Sales', data: 'Q, Sales\nQ1, 10\nQ2, 20' });
    const chart = await run(`DocsExtra.chartImage().then(c => c && c.bytes.byteLength)`);
    check('chart maker gives a PNG', chart > 2000, chart);
    await run(`(DocsExtra.togglePageless(active), true)`);
    check('pageless view', await run(`active.body.classList.contains('pageless')`));
    await run(`(DocsExtra.togglePageless(active), true)`);
    await run(`(DocsExtra.toggleRuler(active), true)`);
    check('ruler shown', await run(`!!document.querySelector('.dd-ruler .dd-ind.left')`));
    await run(`(DocsExtra.toggleRuler(active), true)`);

    await run(`save(active)`);
    await until(`!active.dirty`);
    fs.copyFileSync(work, path.join(outDir, 'edited.docx'));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 700)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
