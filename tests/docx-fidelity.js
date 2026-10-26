'use strict';
// Word fidelity test: set DOCSTUDIO_SRC=<rich.docx> and DOCSTUDIO_OUT=<outDir>, then `electron tests/docx-fidelity.js`
// (paths go in the environment because the app opens any file named on the command line).
// Edits a document through the real editor, saves it, and checks that everything not edited is identical
// and that edited paragraphs kept their original formatting.

const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-fid-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => { checks.push({ name, ok: !!ok, detail }); };
const consoleErrors = [];

async function partStrings(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const out = {};
  for (const name of Object.keys(zip.files)) if (!zip.files[name].dir) out[name] = await zip.file(name).async('string');
  return out;
}
// Body paragraphs as XML strings, using the same parser the app uses (via the page).
async function paragraphsXml(win, xml) {
  return win.webContents.executeJavaScript(`(() => {
    const d = new DOMParser().parseFromString(${JSON.stringify(xml)}, 'application/xml');
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const body = d.getElementsByTagNameNS(W, 'body')[0];
    return [...body.getElementsByTagNameNS(W, 'p')].map((p) => new XMLSerializer().serializeToString(p).replace(/ xmlns(:\\w+)?="[^"]*"/g, ''));
  })()`);
}

let step = 'start';
app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1300, 900);
  win.webContents.on('console-message', (_e, level, msg, line, s) => { if (level >= 3) consoleErrors.push(`${path.basename(s || '')}:${line} ${msg}`); });
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const run = async (js) => {
    step = js.slice(0, 90);
    // The last statement of a snippet is its result.
    const cut = /^\(.*\)\(\)$/s.test(js.trim()) ? -1 : js.trimEnd().replace(/;$/, '').lastIndexOf(';');
    const body = cut < 0 ? `return await (${js});` : `${js.slice(0, cut + 1)} return await (${js.slice(cut + 1).trim().replace(/;$/, '') || 'undefined'});`;
    const r = await win.webContents.executeJavaScript(`(async () => { try { ${body} } catch (e) { return { __err: e.stack || String(e) }; } })()`);
    if (r && r.__err) throw new Error(`${step}
${r.__err}`);
    return r;
  };
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());

  const work = path.join(outDir, 'work.docx');
  fs.copyFileSync(src, work);
  win.webContents.send('app:open-path', work);
  await wait(2500);
  await shot('1-view.png');
  await run(`setMode(active, 'edit')`);
  await wait(2500);
  await shot('2-edit.png');

  // --- A: save without edits = identical paragraphs
  const original = await partStrings(src);
  const docPath = 'word/document.xml';
  const origParas = await paragraphsXml(win, original[docPath]);
  const built0 = await run(`active.editor.build().then(b => Array.from(new Uint8Array(b.bytes)))`);
  fs.writeFileSync(path.join(outDir, 'a-noedit.docx'), Buffer.from(built0));
  const a = await partStrings(path.join(outDir, 'a-noedit.docx'));
  const aParas = await paragraphsXml(win, a[docPath]);
  check('no edits: same paragraph count', aParas.length === origParas.length, `${aParas.length} vs ${origParas.length}`);
  check('no edits: every paragraph identical', aParas.every((p, i) => p === origParas[i]));
  check('no edits: other parts identical', Object.keys(original).filter((k) => k !== docPath).every((k) => a[k] === original[k]));

  // --- B: edits
  const helpers = `
    window.__find = (needle) => {
      const w = document.createTreeWalker(active.editor.stage, NodeFilter.SHOW_TEXT);
      for (let n = w.nextNode(); n; n = w.nextNode()) { const i = n.data.indexOf(needle); if (i >= 0) return [n, i]; }
      return null;
    };
    window.__select = (needle, start, len) => {
      const [n, i] = __find(needle);
      const r = document.createRange();
      r.setStart(n, i + (start || 0)); r.setEnd(n, i + (start || 0) + (len == null ? needle.length : len));
      n.parentElement.closest('article').focus();
      getSelection().removeAllRanges(); getSelection().addRange(r);
    };
    true`;
  await run(helpers);
  // B1: type at the end of the centred paragraph
  await run(`__select('be edited.', 10, 0); document.execCommand('insertText', false, ' EDITED'); true`);
  // B2: bold one word in the Georgia run
  await run(`__select('intro'); document.execCommand('bold'); true`);
  // B3: split "Last paragraph." with Enter
  await run(`__select('Last paragraph.', 5, 0); document.execCommand('insertParagraph'); true`);
  // B4: type inside the hyperlink text
  await run(`__select('our website', 4, 0); document.execCommand('insertText', false, 'great '); true`);
  // B5: colour a word in the purple run
  await run(`__select('purple'); document.execCommand('foreColor', false, '#0000ff'); true`);
  await run(`active.editor.opts.onChange(); true`);
  await wait(300);
  await shot('3-edited.png');
  await run(`active.path = active.path.replace('work.docx', 'b-edited.docx'); save(active)`);
  await wait(800);
  const bFile = path.join(outDir, 'b-edited.docx');
  const b = await partStrings(bFile);
  const bParas = await paragraphsXml(win, b[docPath]);
  const textOf = (x) => (x.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join('');
  const find = (arr, s) => arr.find((p) => textOf(p).includes(s)) || '';
  const centred = find(bParas, 'EDITED');
  check('typed text saved', textOf(centred) === 'Centered paragraph that will be EDITED edited.' || textOf(centred).includes('EDITED'), textOf(centred));
  check('centred paragraph keeps alignment', centred.includes('<w:jc w:val="center"/>'));
  check('centred paragraph keeps italic/size/colour', centred.includes('<w:i/>') && centred.includes('<w:sz w:val="32"/>') && centred.includes('C0392B'), centred.slice(0, 600));
  const intro = find(bParas, 'Georgia 14pt');
  check('intro keeps Georgia on every Georgia run', (intro.match(/w:ascii="Georgia"/g) || []).length >= 3, intro.slice(0, 900));
  check('"intro" made bold', /<w:b\/>[\s\S]*?<w:t[^>]*>intro<\/w:t>/.test(intro));
  check('purple run still Arial bold, "purple" now blue', intro.includes('w:ascii="Arial"') && intro.includes('0000FF') && intro.includes('6A1B9A'));
  const link = find(bParas, 'great');
  check('hyperlink kept with its relationship', /<w:hyperlink [^>]*r:id="rId\d+"/.test(link) && textOf(link).includes('our great website'), link.slice(0, 500));
  check('bookmark kept', link.includes('w:bookmarkStart') && link.includes('MyMark') && link.includes('w:bookmarkEnd'));
  check('split paragraph -> two paragraphs', bParas.some((p) => textOf(p) === 'Last ') && bParas.some((p) => textOf(p) === 'paragraph.'));
  const untouched = origParas.filter((p) => !/EDITED|Georgia|our website|Last paragraph|Centered/.test(textOf(p)));
  check('untouched paragraphs byte-identical', untouched.every((p) => bParas.includes(p)), `${untouched.filter((p) => !bParas.includes(p)).length} changed`);
  check('picture kept', b[docPath].includes('<w:drawing>') && Object.keys(b).some((k) => k.startsWith('word/media/')));
  check('header/footer/styles identical', ['word/styles.xml', 'word/numbering.xml'].concat(Object.keys(original).filter((k) => /header|footer/.test(k))).every((k) => b[k] === original[k]));

  // --- C: structural edits: bullets, table, picture, page break
  await run(`__select('paragraph.', 0, 0); true`);
  await run(`active.editor.toggleList('bullet')`);
  await wait(1200);
  await run(helpers);
  await run(`__select('paragraph.', 0, 0); true`);
  await run(`active.editor.insertTable(2, 2)`);
  await wait(1200);
  await run(helpers);
  await run(`__select('Untouched paragraph', 0, 0); true`);
  const png = fs.readFileSync(path.join(path.dirname(src), 'picture.png'));
  await run(`active.editor.insertImage(new Uint8Array(${JSON.stringify(Array.from(png))}).buffer, 'png', 320, 180)`);
  await wait(1200);
  await run(helpers);
  await run(`__select('Last ', 0, 0); true`);
  await run(`active.editor.insertPageBreak()`);
  await wait(1200);
  await shot('4-structural.png');
  await run(`active.path = active.path.replace('b-edited.docx', 'c-structural.docx'); save(active)`);
  await wait(800);
  const c = await partStrings(path.join(outDir, 'c-structural.docx'));
  check('bullet list applied', /<w:numPr><w:ilvl w:val="0"\/><w:numId w:val="\d+"\/><\/w:numPr>[\s\S]*?paragraph\./.test(c[docPath]));
  check('table inserted', (c[docPath].match(/<w:tbl>/g) || []).length === 2);
  check('picture inserted', (c[docPath].match(/<w:drawing>/g) || []).length === 2 && Object.keys(c).filter((k) => k.startsWith('word/media/')).length === 2);
  check('page break inserted', c[docPath].includes('<w:br w:type="page"/>'));
  check('edits from before structural changes survived', c[docPath].includes('EDITED') && c[docPath].includes('great'));

  // --- D: new document
  await run(`newDocument('docx'); true`);
  await wait(2500);
  await run(`(() => { const a = active.editor.articles[0]; const p = a.querySelector('p'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); a.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, 'Hello new document'); return true; })()`);
  await run(`active.editor.setParagraphStyle(active.editor.selectedParagraphs(), 'Heading1'); true`);
  await shot('5-new.png');
  await run(`(async () => { active.path = ${JSON.stringify(path.join(outDir, 'd-new.docx'))}; active.isNew = false; active.backedUp = true; return save(active); })()`);
  await wait(800);
  const dd = await partStrings(path.join(outDir, 'd-new.docx'));
  check('new document saved with text and Heading1', dd[docPath].includes('Hello new document') && dd[docPath].includes('<w:pStyle w:val="Heading1"/>'));

  // Re-open the saved files in the viewer to be sure they parse.
  for (const f of ['b-edited.docx', 'c-structural.docx', 'd-new.docx']) {
    win.webContents.send('app:open-path', path.join(outDir, f));
    await wait(2000);
    const ok = await run(`!document.querySelector('.pane.active .failure') && !!document.querySelector('.pane.active .docx-view section')`);
    check(`re-opens: ${f}`, ok);
  }
  await shot('6-reopened.png');

  const summary = { passed: checks.filter((x) => x.ok).length, failed: checks.filter((x) => !x.ok), consoleErrors };
  fs.writeFileSync(path.join(outDir, 'fidelity.json'), JSON.stringify({ checks, consoleErrors }, null, 2));
  console.log(JSON.stringify(summary, null, 2));
  app.exit(0);
}).catch((err) => {
  console.log('FAILED at step:', err.message);
  console.log(JSON.stringify({ checks, consoleErrors }, null, 2));
  app.exit(1);
});
