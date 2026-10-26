'use strict';
// Word test on a document written by real Word: headers/footers/footnotes editing, formatted paste, undo across
// structural edits, and pagination. DOCSTUDIO_SRC=<word-made.docx> DOCSTUDIO_OUT=<dir> electron tests/docx-parts-test.js

const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-wp-')));
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
  win.setSize(1300, 900);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const run = async (js) => {
    const cut = /^\(.*\)\(\)$/s.test(js.trim()) ? -1 : js.trimEnd().replace(/;$/, '').lastIndexOf(';');
    const body = cut < 0 ? `return await (${js});` : `${js.slice(0, cut + 1)} return await (${js.slice(cut + 1).trim().replace(/;$/, '') || 'undefined'});`;
    const r = await win.webContents.executeJavaScript(`(async () => { try { ${body} } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${js.slice(0, 120)}\n${r.__err}`);
    return r;
  };
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  const original = await parts(src);

  const work = path.join(outDir, 'work.docx');
  fs.copyFileSync(src, work);
  win.webContents.send('app:open-path', work);
  await wait(3000);
  const viewPages = await run(`document.querySelectorAll('.pane.active .docx-view section.docx').length`);
  check('View paginates long text into several pages', viewPages >= 3, viewPages);
  await shot('1-view.png');
  await run(`setMode(active, 'edit')`);
  await wait(2500);
  await run(`window.__caretIn = (el, atEnd) => { const r = document.createRange(); r.selectNodeContents(el); r.collapse(!atEnd); el.closest('[contenteditable="true"]').focus(); getSelection().removeAllRanges(); getSelection().addRange(r); }; true`);
  check('headers/footers are editable', await run(`[...active.editor.stage.querySelectorAll('section.docx > header')].some(h => h.isContentEditable && h.querySelector('p[data-dd-p]'))`));
  check('page markers drawn in edit view', await run(`active.editor.stage.querySelectorAll('.dd-page-marker').length >= 1`));

  // No edits: every part identical.
  const b0 = await run(`active.editor.build().then(b => Array.from(new Uint8Array(b.bytes)))`);
  fs.writeFileSync(path.join(outDir, 'a-noedit.docx'), Buffer.from(b0));
  const a = await parts(path.join(outDir, 'a-noedit.docx'));
  const sameParts = Object.keys(original).filter((k) => k !== 'word/document.xml').every((k) => a[k] === original[k]);
  check('no edits: headers/footers/notes/styles byte-identical', sameParts);

  // Edit section 1 header, the footer (keeping the PAGE field), and the footnote.
  await run(`(() => { const h = [...active.editor.stage.querySelectorAll('header p')].find(p => p.textContent.includes('Section one header')); __caretIn(h, true); document.execCommand('insertText', false, ' EDITED-H'); return true; })()`);
  await run(`(() => { const f = [...active.editor.stage.querySelectorAll('footer p')].find(p => p.textContent.includes('Page')); const r = document.createRange(); r.setStart(f.firstChild.firstChild || f.firstChild, 0); r.collapse(true); f.closest('[contenteditable]').focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, 'FOOT '); return true; })()`);
  await run(`(() => { const n = [...active.editor.stage.querySelectorAll('p')].find(p => p.textContent.includes('Footnote text from Word')); if (!n) return false; __caretIn(n, true); document.execCommand('insertText', false, ' EDITED-N'); return true; })()`);
  // Formatted paste into the intro paragraph.
  await run(`(() => { const p = [...active.editor.stage.querySelectorAll('article p')].find(p => p.textContent.startsWith('Intro paragraph')); __caretIn(p, true); const dt = new DataTransfer(); dt.setData('text/html', '<p>Pasted <b>bold</b> and <i style="color:#c00000">red italic</i> with <a href="https://example.com/x">a link</a>.</p>'); dt.setData('text/plain', 'Pasted bold and red italic with a link.'); p.closest('[contenteditable]').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); return true; })()`);
  await run(`active.editor.opts.onChange(); true`);
  await shot('2-edited.png');

  // Structural edit + undo + redo.
  await run(`(() => { const p = [...active.editor.stage.querySelectorAll('article p')].find(p => p.textContent.startsWith('Long paragraph 3:')); __caretIn(p, true); return true; })()`);
  await run(`active.editor.insertTable(2, 3)`);
  await wait(800);
  const tablesAfterInsert = await run(`active.editor.stage.querySelectorAll('article table').length`);
  await run(`active.editor.undo()`);
  await wait(800);
  const tablesAfterUndo = await run(`active.editor.stage.querySelectorAll('article table').length`);
  check('undo removes an inserted table', tablesAfterUndo === tablesAfterInsert - 1, { tablesAfterInsert, tablesAfterUndo });
  check('undo keeps earlier text edits', await run(`active.editor.stage.textContent.includes('EDITED-H') && active.editor.stage.textContent.includes('EDITED-N') && active.editor.stage.textContent.includes('red italic')`));
  await run(`active.editor.redo()`);
  await wait(800);
  check('redo puts the table back', (await run(`active.editor.stage.querySelectorAll('article table').length`)) === tablesAfterInsert);
  await run(`active.editor.undo()`);
  await wait(800);

  await run(`active.path = active.path.replace('work.docx', 'parts.docx'); active.backedUp = true; save(active)`);
  await wait(1000);
  const s = await parts(path.join(outDir, 'parts.docx'));
  const headerFile = Object.keys(s).find((k) => /header\d\.xml$/.test(k) && s[k].includes('EDITED-H'));
  check('header edit saved in its header part', !!headerFile);
  check('other header unchanged', Object.keys(original).filter((k) => /header\d\.xml$/.test(k) && k !== headerFile).every((k) => s[k] === original[k]));
  const footerFile = Object.keys(s).find((k) => /footer\d\.xml$/.test(k) && s[k].includes('FOOT '));
  check('footer edit saved, PAGE field kept', !!footerFile && /fldChar|fldSimple|instrText/.test(s[footerFile]), footerFile);
  check('footnote edit saved', s['word/footnotes.xml'].includes('EDITED-N') && s['word/footnotes.xml'].includes('w:type="separator"'));
  const doc = s['word/document.xml'];
  check('pasted bold/italic/colour/link kept', /<w:b\/>[\s\S]{0,200}?<w:t[^>]*>bold<\/w:t>/.test(doc) && /<w:i\/>[\s\S]{0,300}?<w:t[^>]*>red italic<\/w:t>/.test(doc) && /C00000/i.test(doc) && /<w:hyperlink [^>]*r:id=/.test(doc));
  check('footnote reference kept in body', doc.includes('w:footnoteReference'));
  check('section break kept', (doc.match(/<w:sectPr/g) || []).length >= 2);
  check('no table left after undo', !doc.includes('<w:tbl>') || (doc.match(/<w:tbl>/g) || []).length === (original['word/document.xml'].match(/<w:tbl>/g) || []).length);

  const summary = { passed: checks.filter((x) => x.ok).length, failed: checks.filter((x) => !x.ok) };
  console.log(JSON.stringify(summary, null, 2));
  app.exit(0);
}).catch((e) => { console.log('ERROR', e.message); console.log(JSON.stringify(checks, null, 1)); app.exit(1); });
