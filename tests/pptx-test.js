'use strict';
// PowerPoint editor test: DOCSTUDIO_SRC=<deck.pptx> DOCSTUDIO_OUT=<dir> electron tests/pptx-test.js
// The test deck (made by PowerPoint) has: 1 title + subtitle, 2 title + bullets + rounded shape,
// 3 title + table + picture, 4 text box (with speaker notes).
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-pptx-')));
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
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  try {
    const work = path.join(outDir, 'work.pptx');
    fs.copyFileSync(src, work);
    win.webContents.send('app:open-path', work);
    await wait(2500);
    check('view mode draws every slide', await run(`document.querySelectorAll('.pane.active .pp-view .pp-slide').length === 4`));
    await shot('1-view.png');
    await run(`setMode(active, 'edit')`);
    await wait(1500);
    check('edit mode with 4 thumbnails', await run(`document.querySelectorAll('.pane.active .ppe-thumb').length === 4`));

    // 1. Edit the title on slide 1 and make one word bold.
    await run(`(async () => {
      const ed = active.pptEditor;
      const shape = ed.slideEl.querySelector('.pp-own[data-idx="0"]');
      ed.editText(shape);
      const run = shape.querySelector('.pp-r');
      run.firstChild.data = 'Quarterly Review 2026';
      const r = document.createRange(); r.setStart(run.firstChild, 17); r.setEnd(run.firstChild, 21);
      getSelection().removeAllRanges(); getSelection().addRange(r);
      ed.format('bold');
      ed.commitText();
      return true;
    })()`);
    const t1 = await run(`(() => { const d = active.pptEditor.deck.slides[0].doc; const rs = [...d.getElementsByTagNameNS(PptxEditor.NS.a, 'r')].slice(0, 3); return rs.map(r => ({ t: r.textContent, b: (r.getElementsByTagNameNS(PptxEditor.NS.a, 'rPr')[0] || { getAttribute: () => null }).getAttribute('b') })); })()`);
    check('title text edited with a bold run', t1.length >= 2 && t1[0].t === 'Quarterly Review ' && t1[1].t === '2026' && t1[1].b === '1', t1);

    // 2. New paragraph in the slide-2 bullets keeps the bullet level properties.
    await run(`(async () => {
      const ed = active.pptEditor; await ed.go(1);
      const shape = ed.slideEl.querySelector('.pp-own[data-idx="1"]');
      ed.editText(shape);
      const ps = shape.querySelectorAll('.pp-p');
      const last = ps[ps.length - 1];
      const copy = last.cloneNode(true);
      copy.querySelectorAll('.pp-r').forEach((r, i) => { r.textContent = i ? '' : 'Hiring two engineers'; });
      last.after(copy);
      ed.commitText();
      return true;
    })()`);
    const t2 = await run(`(() => { const d = active.pptEditor.deck.slides[1].doc; const body = d.getElementsByTagNameNS(PptxEditor.NS.p, 'txBody')[1]; return [...body.getElementsByTagNameNS(PptxEditor.NS.a, 'p')].map(p => p.textContent); })()`);
    check('bullet paragraph added', t2.length === 4 && t2[3] === 'Hiring two engineers', t2);

    // 3. Move and resize the rounded shape.
    await run(`(async () => {
      const ed = active.pptEditor;
      const shape = ed.slideEl.querySelector('.pp-own[data-idx="2"]');
      ed.select(shape);
      ed.snapshot();
      shape.style.left = (parseFloat(shape.style.left) - 100) + 'px';
      shape.style.width = (parseFloat(shape.style.width) + 50) + 'px';
      await ed.applyBox(shape, true);
      return true;
    })()`);
    const off = await run(`(() => { const d = active.pptEditor.deck.slides[1].doc; const sp = d.getElementsByTagNameNS(PptxEditor.NS.p, 'sp')[2]; const x = sp.getElementsByTagNameNS(PptxEditor.NS.a, 'xfrm')[0]; const o = x.getElementsByTagNameNS(PptxEditor.NS.a, 'off')[0]; const e = x.getElementsByTagNameNS(PptxEditor.NS.a, 'ext')[0]; return { x: +o.getAttribute('x'), cx: +e.getAttribute('cx') }; })()`);
    check('shape moved and resized in EMU', off.x < 620 * 12700 - 100000 && off.cx > 260 * 12700 + 100000, off);

    // 4. Speaker notes: add to slide 2 (it has none), change slide 4's.
    await run(`active.pptEditor.setNotes(1, 'Mention the hiring plan\\nSecond line')`);
    await run(`active.pptEditor.setNotes(3, 'Updated notes for slide four')`);
    check('notes read back', await run(`(async () => (await active.pptEditor.getNotes(1)) === 'Mention the hiring plan\\nSecond line' && (await active.pptEditor.getNotes(3)) === 'Updated notes for slide four')()`));

    // 5. Insert a text box, a shape and a picture on slide 4.
    await run(`(async () => {
      const ed = active.pptEditor; await ed.go(3);
      await ed.insertTextBox();
      await ed.insertShape('ellipse');
      const c = document.createElement('canvas'); c.width = 40; c.height = 30; const g = c.getContext('2d'); g.fillStyle = '#e53935'; g.fillRect(0, 0, 40, 30);
      const bytes = await (await new Promise(r => c.toBlob(r, 'image/png'))).arrayBuffer();
      await ed.insertPicture(bytes, 'png', 400, 300);
      return true;
    })()`);
    check('three items inserted', await run(`active.pptEditor.slideEl.querySelectorAll('.pp-own').length === 4 && active.pptEditor.slideEl.querySelector('.pp-pic img') != null`));
    await shot('2-inserted.png');

    // 6. Undo removes the picture, redo brings it back.
    await run(`active.pptEditor.undo()`);
    const afterUndo = await run(`active.pptEditor.slideEl.querySelectorAll('.pp-own').length`);
    await run(`active.pptEditor.redo()`);
    const afterRedo = await run(`active.pptEditor.slideEl.querySelectorAll('.pp-own').length`);
    check('undo/redo of an insert', afterUndo === 3 && afterRedo === 4, { afterUndo, afterRedo });

    // 7. Slides: duplicate slide 2, delete slide 3 (the table slide), add a slide from a layout, move slide 1 to second.
    await run(`active.pptEditor.duplicateSlide(1)`);
    check('duplicate slide', await run(`active.pptEditor.deck.slides.length === 5`));
    await run(`active.pptEditor.deleteSlide(3)`);
    const layouts = await run(`active.pptEditor.layouts().then(l => l.map(x => x.name))`);
    await run(`active.pptEditor.layouts().then(l => active.pptEditor.addSlide(l.find(x => /Two Content/.test(x.name)).path))`);
    await run(`active.pptEditor.moveSlide(0, 1)`);
    const titles = await run(`active.pptEditor.deck.slides.map(s => s.doc.getElementsByTagNameNS(PptxEditor.NS.a, 't')[0]?.textContent || '')`);
    check('slide operations', titles.length === 5 && titles[0] === 'Highlights' && titles[1].startsWith('Quarterly'), { titles, layouts });
    // Type into the new slide's title placeholder.
    await run(`(async () => {
      const ed = active.pptEditor; const i = ed.deck.slides.findIndex(s => !s.doc.getElementsByTagNameNS(PptxEditor.NS.a, 't').length); await ed.go(i);
      const shape = ed.slideEl.querySelector('.pp-own[data-idx="0"]');
      ed.editText(shape);
      const sp = shape.querySelector('.pp-p'); sp.textContent = 'Two columns';
      ed.commitText();
      return true;
    })()`);
    await shot('3-slides.png');

    // 8. Save and check the package.
    await run(`save(active)`);
    await wait(500);
    check('saved without errors', await run(`!active.dirty`));
    const zip = await JSZip.loadAsync(fs.readFileSync(work));
    const pres = await zip.file('ppt/presentation.xml').async('string');
    const ct = await zip.file('[Content_Types].xml').async('string');
    const presRels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    const ids = [...pres.matchAll(/<p:sldId [^>]*r:id="(rId\d+)"/g)].map((m) => m[1]);
    const targets = ids.map((id) => (new RegExp(`Id="${id}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${id}"`).exec(presRels) || []).slice(1).find(Boolean));
    check('five slides listed, all parts present', ids.length === 5 && targets.every((t) => t && zip.file('ppt/' + t)), targets);
    check('content types cover every slide', targets.every((t) => ct.includes(`/ppt/${t}"`)));
    check('deleted slide removed from the package', !targets.includes('slides/slide3.xml') && !zip.file('ppt/slides/slide3.xml') && !ct.includes('/ppt/slides/slide3.xml"'));
    check('picture saved as media', Object.keys(zip.files).some((f) => /^ppt\/media\/image\d+\.png$/.test(f)) && /Extension="png"/i.test(ct));
    fs.copyFileSync(work, path.join(outDir, 'edited.pptx'));

    // 9. Reopen the saved file: everything still renders.
    await run(`closeTab(active)`);
    win.webContents.send('app:open-path', work);
    await wait(2500);
    check('reopened file draws 5 slides', await run(`document.querySelectorAll('.pane.active .pp-view .pp-slide').length === 5`));
    await shot('4-reopened.png');
    await run(`closeTab(active)`);

    // 10. New presentation from the template.
    await run(`newDocument('pptx')`);
    await wait(2000);
    await run(`(async () => {
      const ed = active.pptEditor;
      const shape = ed.slideEl.querySelector('.pp-own[data-idx="0"]');
      ed.editText(shape);
      shape.querySelector('.pp-p').textContent = 'A brand new deck';
      ed.commitText();
      await ed.addSlide(null);
      return true;
    })()`);
    const newBytes = await run(`active.pptEditor.save().then(b => Array.from(new Uint8Array(b)))`);
    fs.writeFileSync(path.join(outDir, 'new.pptx'), Buffer.from(newBytes));
    const nz = await JSZip.loadAsync(Buffer.from(newBytes));
    const core = await nz.file('docProps/core.xml').async('string');
    check('new deck: 2 slides, title saved, no personal name', (await nz.file('ppt/presentation.xml').async('string')).match(/<p:sldId /g).length === 2 && (await nz.file('ppt/slides/slide1.xml').async('string')).includes('A brand new deck') && /<dc:creator>Doc Studio<\/dc:creator>/.test(core) && /dcterms:created/.test(core), core.slice(0, 400));
    await run(`(active.dirty = false, closeTab(active))`);
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 600)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
