'use strict';
// OpenDocument text + native Rich Text editors:
//   DOCSTUDIO_ODT=<file.odt> DOCSTUDIO_RTF=<file.rtf> DOCSTUDIO_OUT=<dir> electron tests/odt-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const odtSrc = process.env.DOCSTUDIO_ODT;
const rtfSrc = process.env.DOCSTUDIO_RTF;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-odt-')));
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
  const until = async (js, ms = 20000) => { const t = Date.now(); while (Date.now() - t < ms) { if (await run(js)) return true; await wait(200); } return false; };
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  // Put the caret in the paragraph containing `text` (offset from its start, or at its end), optionally selecting `len` chars.
  const caret = (text, offset, len) => run(`(() => {
    const ed = active.richEditor; const walker = document.createTreeWalker(ed.article, NodeFilter.SHOW_TEXT);
    let n; while ((n = walker.nextNode())) { const i = n.data.indexOf(${JSON.stringify(text)}); if (i >= 0) {
      const r = document.createRange(); const at = i + ${offset == null ? text.length : offset};
      r.setStart(n, at); r.setEnd(n, Math.min(n.data.length, at + ${len || 0}));
      ed.article.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); return true; } }
    return false; })()`);
  const contentOf = async (file) => { const z = await JSZip.loadAsync(fs.readFileSync(file)); return z.file('content.xml').async('string'); };
  try {
    // ---------------- ODT
    const work = path.join(outDir, 'work.odt');
    fs.copyFileSync(odtSrc, work);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.odtEditor)`);
    const paraCount = await run(`active.odtEditor.paras.length`);
    check('ODT draws its paragraphs with styles', paraCount > 5 && await run(`!!document.querySelector('.pane.active .odt-page .odt-p')`), paraCount);
    await shot('odt-1-view.png');
    await run(`setMode(active, 'edit')`);
    check('edit mode makes the page editable', await run(`active.odtEditor.article.isContentEditable`));
    const firstText = await run(`active.odtEditor.article.querySelector('.odt-p').textContent`);
    // 1. Type at the start of the first paragraph.
    await run(`(() => { const p = active.odtEditor.article.querySelector('.odt-p'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); active.odtEditor.article.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, 'ODT EDIT '); return true; })()`);
    // 2. Bold a word.
    const word = await run(`(() => { const ps = [...active.odtEditor.article.querySelectorAll('.odt-p')]; const p = ps.find((x, i) => i > 0 && /\\w{5,}/.test(x.textContent) && !x.closest('table')); return p ? p.textContent.match(/\\w{5,}/)[0] : null; })()`);
    await caret(word, 0, word.length);
    await run(`(document.execCommand('styleWithCSS', false, true), document.execCommand('bold'), true)`);
    // 3. New paragraph after the word's paragraph (Enter at the end).
    await caret(word, null, 0);
    await run(`(() => { const s = getSelection(); const p = s.anchorNode.parentElement.closest('.odt-p'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertParagraph'); document.execCommand('insertText', false, 'A brand new paragraph'); return true; })()`);
    // 4. Center the first paragraph.
    await run(`(() => { const p = active.odtEditor.article.querySelector('.odt-p'); active.odtEditor.align([p], 'center'); return true; })()`);
    // 5. Change the font colour of a word in the last paragraph outside tables.
    await shot('odt-2-edited.png');
    check('edits mark the tab dirty', await run(`active.dirty`));
    await run(`save(active)`);
    await until(`!active.dirty`);
    const xml = await contentOf(work);
    check('typed text saved', xml.includes('ODT EDIT ') && xml.includes(firstText.slice(0, 10).replace(/&/g, '&amp;')), firstText);
    check('new paragraph saved', xml.includes('A brand new paragraph'));
    check('bold saved as an automatic text style', /<style:style style:name="DDT\d+" style:family="text"[^>]*>(<[^>]+>)*?<style:text-properties[^>]*fo:font-weight="bold"/.test(xml) && new RegExp(`<text:span text:style-name="DDT\\d+">${word}</text:span>`).test(xml), word);
    check('centered paragraph saved', /style:family="paragraph"[^>]*>.*?fo:text-align="center"/.test(xml));
    const z = await JSZip.loadAsync(fs.readFileSync(work));
    const names = Object.keys(z.files);
    check('mimetype stays first and stored', names[0] === 'mimetype' && fs.readFileSync(work).subarray(30, 38).toString() === 'mimetype' && fs.readFileSync(work).readUInt16LE(8) === 0);
    const orig = await JSZip.loadAsync(fs.readFileSync(odtSrc));
    check('styles.xml untouched', (await z.file('styles.xml').async('string')) === (await orig.file('styles.xml').async('string')));
    fs.copyFileSync(work, path.join(outDir, 'edited.odt'));
    // Undo-free sanity: reopen and check the page still draws and the edit is there.
    await run(`closeTab(active)`);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.odtEditor)`);
    check('reopened ODT shows the edits', await run(`active.odtEditor.article.textContent.includes('ODT EDIT') && active.odtEditor.article.textContent.includes('A brand new paragraph')`));
    await run(`closeTab(active)`);

    // Delete a paragraph and insert a picture.
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.odtEditor)`);
    await run(`setMode(active, 'edit')`);
    await run(`(() => { const p = [...active.odtEditor.article.querySelectorAll('.odt-p')].find(x => x.textContent === 'A brand new paragraph'); p.remove(); return true; })()`);
    await run(`(async () => { const c = document.createElement('canvas'); c.width = 60; c.height = 40; const g = c.getContext('2d'); g.fillStyle = '#1e88e5'; g.fillRect(0, 0, 60, 40); const b = await (await new Promise(r => c.toBlob(r, 'image/png'))).arrayBuffer(); const p = active.odtEditor.article.querySelector('.odt-p'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); active.odtEditor.article.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); await active.odtEditor.insertImage(b, 'png', 60, 40); return true; })()`);
    await run(`(active.dirty = true, save(active))`);
    await until(`!active.dirty`);
    const xml2 = await contentOf(work);
    const z2 = await JSZip.loadAsync(fs.readFileSync(work));
    const man = await z2.file('META-INF/manifest.xml').async('string');
    check('deleted paragraph removed', !xml2.includes('A brand new paragraph'));
    check('picture saved with manifest entry', /<draw:image xlink:href="Pictures\/docstudio1\.png"/.test(xml2) && !!z2.file('Pictures/docstudio1.png') && man.includes('Pictures/docstudio1.png'));
    fs.copyFileSync(work, path.join(outDir, 'edited2.odt'));
    await run(`closeTab(active)`);

    // New ODT.
    await run(`newDocument('odt')`);
    await until(`!!(active && active.odtEditor)`);
    await run(`(() => { const p = active.odtEditor.article.querySelector('.odt-p'); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); active.odtEditor.article.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, 'Hello OpenDocument'); return true; })()`);
    await run(`(() => { const ed = active.odtEditor; const p = ed.article.querySelector('.odt-p'); ed.setParagraphStyle([p], 'Heading_20_1'); return true; })()`);
    const newBytes = await run(`active.odtEditor.save().then(b => Array.from(new Uint8Array(b)))`);
    fs.writeFileSync(path.join(outDir, 'new.odt'), Buffer.from(newBytes));
    const nx = await contentOf(path.join(outDir, 'new.odt'));
    check('new ODT has the text in Heading 1', /<text:p text:style-name="Heading_20_1">Hello OpenDocument<\/text:p>/.test(nx), nx.slice(-300));
    await run(`(active.dirty = false, closeTab(active))`);

    // ---------------- RTF (native editor, as used without Word)
    const rtfBytes = Array.from(fs.readFileSync(rtfSrc));
    await run(`(async () => { const tab = makeTab({ path: null, name: 'native.rtf', size: 0, data: new Uint8Array(${JSON.stringify(rtfBytes)}).buffer }); await setupRich(tab, 'rtf'); return true; })()`);
    await until(`!!(active && active.rtfEditor)`);
    await run(`setMode(active, 'edit')`);
    await run(`(() => { const a = active.rtfEditor.article; const r = document.createRange(); r.selectNodeContents(a.firstElementChild); r.collapse(true); a.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, 'RTF EDIT '); return true; })()`);
    await caret('RTF EDIT', 0, 3);
    await run(`(document.execCommand('styleWithCSS', false, true), document.execCommand('bold'), document.execCommand('foreColor', false, '#ff0000'), true)`);
    await shot('rtf-edited.png');
    const rtf = await run(`active.rtfEditor.save().then(b => new TextDecoder().decode(b))`);
    fs.writeFileSync(path.join(outDir, 'native.rtf'), rtf);
    check('RTF written with fonts, colours and the edit', rtf.startsWith('{\\rtf1') && /\{\\fonttbl/.test(rtf) && /\\red255\\green0\\blue0/.test(rtf) && /\\b[^}]*\\cf\d+ RTF\}/.test(rtf) && rtf.includes('EDIT '), rtf.slice(0, 600));
    check('RTF keeps page size', /\\paperw\d+/.test(rtf));
    await run(`(active.dirty = false, closeTab(active))`);
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 600)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
