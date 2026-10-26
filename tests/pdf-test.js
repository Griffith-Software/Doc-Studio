'use strict';
// PDF editor test: DOCSTUDIO_SRC=<report.pdf> DOCSTUDIO_FORM=<form.pdf> DOCSTUDIO_OUT=<dir> electron tests/pdf-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { PDFDocument, PDFName } = require('pdf-lib');

const src = process.env.DOCSTUDIO_SRC;
const formSrc = process.env.DOCSTUDIO_FORM;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-pdf-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail });

async function annots(file) {
  const doc = await PDFDocument.load(fs.readFileSync(file));
  return doc.getPages().map((p) => {
    const a = p.node.Annots();
    return a ? a.asArray().map((ref) => { const d = doc.context.lookup(ref); const st = d.get(PDFName.of('Subtype')); return st ? st.toString() : '?'; }) : [];
  });
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

  const work = path.join(outDir, 'work.pdf');
  fs.copyFileSync(src, work);
  win.webContents.send('app:open-path', work);
  await wait(2500);
  await run(`setMode(active, 'edit')`);
  await wait(2500);
  check('edit mode with pages and thumbnails', await run(`active.pdfEditor.pages.length === 3 && document.querySelectorAll('.pane.active .pdfe-thumb').length === 3`));
  check('text layer rendered', await run(`document.querySelectorAll('.pane.active .pdfe-page[data-index="0"] .textLayer span').length > 5`));
  await shot('1-edit.png');

  // Highlight a text selection on page 1.
  await run(`(() => { const ed = active.pdfEditor; ed.setTool('highlight'); const span = [...ed.pages[0].text.querySelectorAll('span')].find(s => s.textContent.startsWith('Revenue grew')); const r = document.createRange(); r.selectNodeContents(span); getSelection().removeAllRanges(); getSelection().addRange(r); ed.onTextMouseUp(); return true; })()`);
  check('highlight item from text selection', await run(`active.pdfEditor.items.some(i => i.type === 'highlight' && i.text.startsWith('Revenue'))`));
  // Edit existing text (Times line).
  await run(`(() => { const ed = active.pdfEditor; ed.setTool('edittext'); const span = [...ed.pages[0].text.querySelectorAll('span')].find(s => s.textContent.startsWith('The board approved')); ed.startEditText(0, { target: span }); const el = document.activeElement; el.innerText = 'The board approved the REVISED budget on Tuesday.'; el.blur(); return true; })()`);
  const rep = await run(`(() => { const it = active.pdfEditor.items.find(i => i.type === 'replace'); return it ? { font: it.fontName, size: it.size, text: it.text } : null; })()`);
  check('edit text keeps the original font name and size', rep && /Times/i.test(rep.font) && Math.abs(rep.size - 12) < 0.2, rep);
  // Add text, ink, rectangle, note, white-out, signature-like image.
  await run(`(() => { const ed = active.pdfEditor; ed.addItem({ type: 'text', page: 0, x: 72, y: 250, size: 14, color: '#c62828', font: 'Arial', text: 'Added by Doc Studio\\nSecond line' }); ed.addItem({ type: 'ink', page: 0, paths: [[[100, 200], [150, 220], [200, 190]]], color: '#1565c0', width: 2 }); ed.addItem({ type: 'rect', page: 0, rect: [300, 150, 400, 220], color: '#2e7d32', width: 2 }); ed.addItem({ type: 'note', page: 0, x: 500, y: 700, text: 'Check these numbers' }); ed.addItem({ type: 'whiteout', page: 1, rect: [70, 600, 300, 615] }); return true; })()`);
  await run(`(async () => { const c = document.createElement('canvas'); c.width = 120; c.height = 40; const x = c.getContext('2d'); x.fillStyle = '#1a237e'; x.font = '28px cursive'; x.fillText('Sig', 10, 30); const b = await new Promise(r => c.toBlob(r, 'image/png')); const png = new Uint8Array(await b.arrayBuffer()); active.pdfEditor.addItem({ type: 'image', page: 0, rect: [400, 80, 520, 120], png, kind: 'signature' }); return true; })()`);
  // Text on the rotated page.
  await run(`(() => { const ed = active.pdfEditor; ed.addItem({ type: 'text', page: 2, x: 300, y: 300, size: 12, color: '#000000', font: 'Helvetica', text: 'Upright on a rotated page' }); return true; })()`);
  await wait(300);
  await shot('2-items.png');
  // Undo removes the last item, redo restores it.
  const n1 = await run(`active.pdfEditor.items.length`);
  await run(`active.pdfEditor.undo()`);
  const n2 = await run(`active.pdfEditor.items.length`);
  await run(`active.pdfEditor.redo()`);
  check('undo/redo of items', n2 === n1 - 1 && (await run(`active.pdfEditor.items.length`)) === n1, { n1, n2 });

  // Save (items written) then page operations.
  const buildErr = await run(`active.pdfEditor.build().then(() => 'ok', (e) => e.stack)`);
  check('build succeeds', buildErr === 'ok', buildErr);
  await run(`active.path = active.path.replace('work.pdf', 'edited.pdf'); active.backedUp = true; save(active)`);
  await wait(1200);
  const edited = path.join(outDir, 'edited.pdf');
  const a1 = await annots(edited);
  check('annotations saved (Highlight, Ink, Square, Text)', ['/Highlight', '/Ink', '/Square', '/Text'].every((t) => a1[0].includes(t)), a1[0]);
  const text = await run(`(async () => { const d = await DDPdf.pdfjs.getDocument({ data: new Uint8Array(${JSON.stringify(Array.from(fs.readFileSync(edited)))}), isEvalSupported: false }).promise; const out = []; for (let i = 1; i <= d.numPages; i++) { const tc = await (await d.getPage(i)).getTextContent(); out.push(tc.items.map(x => x.str).join(' ')); } return out; })()`);
  check('added text is real PDF text', text[0].includes('Added by Doc Studio') && text[0].includes('Second line'));
  check('replaced text present', text[0].includes('REVISED budget on Tuesday'));
  check('text on rotated page', text[2].includes('Upright on a rotated page'));
  check('after save the editor is clean', await run(`active.pdfEditor.items.length === 0 && !active.dirty`));

  await run(`active.pdfEditor.rotatePage(0, 90)`);
  await run(`active.pdfEditor.insertBlank(1)`);
  await run(`active.pdfEditor.movePage(3, 0)`);
  check('page operations applied', await run(`active.pdfEditor.pages.length === 4 && active.pdfEditor.pages[1].rotation === 90`), await run(`active.pdfEditor.pages.map(p => p.rotation)`));
  await run(`active.pdfEditor.deletePage(2)`);
  await run(`active.pdfEditor.undo()`);
  check('undo a page delete', (await run(`active.pdfEditor.pages.length`)) === 4);
  await run(`active.path = active.path.replace('edited.pdf', 'pages.pdf'); save(active)`);
  await wait(1200);
  const pagesDoc = await PDFDocument.load(fs.readFileSync(path.join(outDir, 'pages.pdf')));
  check('saved page order/rotation', pagesDoc.getPageCount() === 4 && pagesDoc.getPage(1).getRotation().angle === 90);
  const a2 = await annots(path.join(outDir, 'pages.pdf'));
  check('annotations survive page moves', a2[1].includes('/Highlight'), a2);
  await shot('3-pages.png');

  // Forms.
  win.webContents.send('app:open-path', formSrc);
  await wait(2500);
  await run(`setMode(active, 'edit')`);
  await wait(2000);
  check('form fields shown', (await run(`document.querySelectorAll('.pane.active .pdfe-field').length`)) >= 5);
  await run(`(() => { const ed = active.pdfEditor; const f = (n) => ed.fields.find(x => x.name === n); ed.setField(f('name'), 'Ada Lovelace'); ed.setField(f('plan'), 'Pro'); ed.setField(f('newsletter'), true); ed.setField(ed.fields.find(x => x.name === 'contact' && x.option === 'phone'), 'phone'); return true; })()`);
  await shot('4-form.png');
  await run(`active.path = ${JSON.stringify(path.join(outDir, 'form-filled.pdf'))}; active.backedUp = true; save(active)`);
  await wait(1200);
  const filled = await PDFDocument.load(fs.readFileSync(path.join(outDir, 'form-filled.pdf')));
  const form = filled.getForm();
  check('form values saved', form.getTextField('name').getText() === 'Ada Lovelace' && form.getDropdown('plan').getSelected()[0] === 'Pro' && form.getCheckBox('newsletter').isChecked() && form.getRadioGroup('contact').getSelected() === 'phone');

  // New PDF.
  await run(`newDocument('pdf'); true`);
  await wait(2500);
  await run(`(() => { active.pdfEditor.addItem({ type: 'text', page: 0, x: 72, y: 700, size: 18, color: '#000000', font: 'Calibri', text: 'Brand new PDF' }); return true; })()`);
  await run(`active.path = ${JSON.stringify(path.join(outDir, 'new.pdf'))}; active.isNew = false; active.backedUp = true; save(active)`);
  await wait(1200);
  check('new PDF saved', fs.existsSync(path.join(outDir, 'new.pdf')) && (await PDFDocument.load(fs.readFileSync(path.join(outDir, 'new.pdf')))).getPageCount() === 1);

  const summary = { passed: checks.filter((x) => x.ok).length, failed: checks.filter((x) => !x.ok) };
  console.log(JSON.stringify(summary, null, 2));
  app.exit(0);
}).catch((e) => { console.log('ERROR', e.message); console.log(JSON.stringify(checks, null, 1)); app.exit(1); });
