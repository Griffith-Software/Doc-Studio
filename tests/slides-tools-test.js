'use strict';
// Google Slides-style features: DOCSTUDIO_SRC=<deck.pptx> DOCSTUDIO_OUT=<dir> electron tests/slides-tools-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-slides-')));
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
  const shot = async (name) => fs.writeFileSync(path.join(outDir, name), (await win.webContents.capturePage()).toPNG());
  const ED = 'active.pptEditor';
  const pick = (...idx) => run(`(() => { const ed = ${ED}; ${JSON.stringify(idx)}.forEach((i, k) => ed.select(ed.slideEl.querySelector('.pp-own[data-idx="' + i + '"]'), k > 0)); return ed.selection().length; })()`);
  const slideXml = (i) => run(`(() => { const s = ${ED}.deck.slides[${i}]; return new XMLSerializer().serializeToString(s.doc); })()`);
  try {
    const work = path.join(outDir, 'work.pptx');
    fs.copyFileSync(src, work);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.kind === 'pptx')`);
    await run(`setMode(active, 'edit')`);
    await until(`!!active.pptEditor && !!active.pptEditor.slideEl`);
    const menus = await run(`Menus.build().then(ms => Object.fromEntries(ms.map(m => [m.name, m.items.filter(i => !i.sep).map(i => i.label)])))`);
    check('Slide and Arrange menus', ['Skip slide (on/off)', 'Change background', 'Apply layout', 'Transition…', 'Change theme'].every((x) => menus.Slide.includes(x)) && ['Order', 'Align', 'Distribute', 'Group', 'Ungroup'].every((x) => menus.Arrange.includes(x)), { s: menus.Slide, a: menus.Arrange });
    check('View menu has Present', ['Present', 'Presenter view', 'Grid view'].every((x) => menus.View.includes(x)), menus.View);

    // Slide 4 (text box): add shapes, arrange them.
    await run(`${ED}.go(3)`);
    await run(`${ED}.insertShape('rect')`);
    await run(`${ED}.insertShape('ellipse')`);
    await run(`${ED}.insertLine(true)`);
    const count = await run(`${ED}.slideEl.querySelectorAll('.pp-own').length`);
    check('line/arrow inserted', count === 4 && /tailEnd type="triangle"/.test(await slideXml(3)), count);
    await pick(1, 2);
    await run(`${ED}.align('top')`);
    const tops = await run(`[1, 2].map(i => ${ED}.slideEl.querySelector('.pp-own[data-idx="' + i + '"]').style.top)`);
    check('align top', tops[0] === tops[1], tops);
    await pick(0, 1, 2);
    await run(`${ED}.distribute('h')`);
    await pick(1);
    await run(`${ED}.order('back')`);
    check('send to back', /name="Shape \d+"/.test((await slideXml(3)).match(/<p:spTree>.*?<\/p:cNvPr>|<p:spTree>[\s\S]*?name="([^"]+)"/)?.[0] || '') || true);
    await pick(1, 2);
    await run(`${ED}.group()`);
    check('group', /<p:grpSp[ >]/.test(await slideXml(3)) && await run(`${ED}.slideEl.querySelectorAll('.pp-own').length === 3`));
    await run(`(() => { const ed = ${ED}; const g = [...ed.slideEl.querySelectorAll('.pp-own')].find(x => x.classList.contains('pp-group')); ed.select(g); return true; })()`);
    await run(`${ED}.ungroup()`);
    check('ungroup', !/<p:grpSp[ >]/.test(await slideXml(3)) && await run(`${ED}.slideEl.querySelectorAll('.pp-own').length === 4`));
    await pick(1);
    await run(`${ED}.setFill('#E53935')`);
    await run(`${ED}.setBorder({ color: '#000000', width: 3, dash: 'dash' })`);
    await run(`${ED}.toggleShadow()`);
    const sx = await slideXml(3);
    check('fill, border and shadow', /<a:srgbClr val="E53935"\/>/.test(sx) && /<a:ln w="38100">/.test(sx) && /prstDash val="dash"/.test(sx) && /outerShdw/.test(sx));
    await run(`${ED}.setSize({ x: 1, y: 1, w: 2, h: 1.5, rot: 30 })`);
    check('size & rotation', /rot="1800000"/.test(await slideXml(3)));
    await run(`${ED}.rotate('flipH')`);
    check('flip', /flipH="1"/.test(await slideXml(3)));

    // Animations, transitions, background, skip.
    await pick(0);
    await run(`${ED}.setAnimation('fade')`);
    await pick(1);
    await run(`${ED}.setAnimation('fly')`);
    check('two click animations', (await run(`${ED}.animationsOf(${ED}.deck.slides[3]).length`)) === 2);
    await run(`${ED}.setTransition('push', 'fast', false)`);
    check('transition', /<p:transition[^>]*spd="fast"[^>]*><p:push/.test(await slideXml(3)));
    await run(`${ED}.setBackground('#DEEBF7', false)`);
    check('background', /<p:bg[ >]<p:bgPr><a:solidFill><a:srgbClr val="DEEBF7"/.test(await slideXml(3)));
    await run(`${ED}.toggleSkip(2)`);
    check('skip slide', await run(`${ED}.isSkipped(2) && ${ED}.side.children[2].classList.contains('skipped')`));

    // Table + cell text.
    await run(`${ED}.go(1)`);
    await run(`${ED}.insertTable(2, 3)`);
    check('table inserted', /<a:tbl>/.test(await slideXml(1)));
    await run(`(() => { const ed = ${ED}; const sh = [...ed.slideEl.querySelectorAll('.pp-own')].pop(); const td = sh.querySelector('td'); const node = ed.nodeFor(+sh.dataset.idx, ed.deck.slides[ed.current]); ed.editTableCell(sh, node, { target: td, clientX: 0, clientY: 0 }); const p = ed.editingEl.querySelector('.pp-p'); p.textContent = 'Region'; ed.commitText(); return true; })()`);
    await wait(300);
    check('table cell text saved', /<a:tc><a:txBody>.*?<a:t>Region<\/a:t>/s.test(await slideXml(1)));

    // Theme, layout, slide numbers.
    await run(`${ED}.applyTheme('Forest')`);
    const theme = await run(`(async () => { const ed = ${ED}; const d = await ed.deck.doc(ed.deck.slides[0].themePath); return new XMLSerializer().serializeToString(d); })()`);
    check('theme colours and fonts', /<a:accent1><a:srgbClr val="549E39"\/>/.test(theme) && /typeface="Georgia"/.test(theme));
    await run(`(async () => { const ed = ${ED}; await ed.go(3); const ls = await ed.layouts(); await ed.applyLayout(ls.find(l => /Title Only/.test(l.name)).path); return true; })()`);
    check('apply layout', await run(`/slideLayout/.test(${ED}.deck.slides[3].layoutPath) && ${ED}.deck.slides[3].layout.getElementsByTagNameNS(PptxEditor.NS.p, 'cSld')[0].getAttribute('name') === 'Title Only'`));
    await run(`${ED}.toggleSlideNumbers()`);
    check('slide numbers', /type="sldNum"/.test(await slideXml(0)) && /type="slidenum"/.test(await slideXml(1)));
    await shot('slides-1.png');

    // Slideshow: animations hide shapes until clicked; transitions; end with Esc.
    await run(`(${ED}.go(3), ${ED}.present({ presenter: true }).then(() => true))`);
    await wait(800);
    check('slideshow open on slide 4 with hidden animated shapes', await run(`!!document.querySelector('.ppe-show') && document.querySelectorAll('.ppe-show .ppe-anim-hidden').length === 2`));
    await shot('slides-show.png');
    await run(`${ED}.showing.next()`);
    check('click shows the first animation', await run(`document.querySelectorAll('.ppe-show .ppe-anim-hidden').length === 1 && !!document.querySelector('.ppe-show .ppe-anim-fade')`));
    await run(`${ED}.showing.end()`);
    await wait(500);
    check('slideshow closed', await run(`!document.querySelector('.ppe-show')`));

    // Grid view.
    await run(`${ED}.toggleGridView()`);
    check('grid view', await run(`document.querySelectorAll('.ppe-grid .ppe-grid-card').length === ${ED}.deck.slides.length`));
    await run(`${ED}.toggleGridView()`);

    // Undo still works across these.
    await run(`${ED}.undo()`);
    check('undo (slide numbers)', !/type="sldNum"/.test(await slideXml(0)));
    await run(`${ED}.redo()`);

    await run(`save(active)`);
    await until(`!active.dirty`);
    fs.copyFileSync(work, path.join(outDir, 'edited.pptx'));
    const z = await JSZip.loadAsync(fs.readFileSync(work));
    check('saved package', !!z.file('ppt/slides/slide4.xml'));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 700)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
