'use strict';
// Image options, video/audio, real charts, diagrams, comments, theme builder:
//   DOCSTUDIO_SRC=<deck.pptx> DOCSTUDIO_VIDEO=<clip.mp4> DOCSTUDIO_OUT=<dir> electron tests/slides-extra-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const JSZip = require('jszip');

const src = process.env.DOCSTUDIO_SRC;
const video = process.env.DOCSTUDIO_VIDEO;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-slidesx-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => { checks.push({ name, ok: !!ok, detail }); console.log('...', ok ? 'ok' : 'FAIL', name); };

// A short WAV tone for the audio test.
function wav() {
  const rate = 8000, n = rate / 2, buf = Buffer.alloc(44 + n * 2);
  buf.write('RIFF', 0); buf.writeUInt32LE(36 + n * 2, 4); buf.write('WAVE', 8); buf.write('fmt ', 12); buf.writeUInt32LE(16, 16);
  buf.writeUInt16LE(1, 20); buf.writeUInt16LE(1, 22); buf.writeUInt32LE(rate, 24); buf.writeUInt32LE(rate * 2, 28); buf.writeUInt16LE(2, 32); buf.writeUInt16LE(16, 34);
  buf.write('data', 36); buf.writeUInt32LE(n * 2, 40);
  for (let i = 0; i < n; i++) buf.writeInt16LE(Math.round(Math.sin((i / rate) * 2 * Math.PI * 440) * 8000), 44 + i * 2);
  return buf;
}

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
  const lastIdx = () => run(`${ED}.slideEl.querySelectorAll('.pp-own').length - 1`);
  const pick = (i) => run(`(() => { const ed = ${ED}; ed.select(ed.slideEl.querySelector('.pp-own[data-idx="' + ${i} + '"]')); return ed.selection().length; })()`);
  const slideXml = (i) => run(`(() => { const s = ${ED}.deck.slides[${i}]; return new XMLSerializer().serializeToString(s.doc); })()`);
  const bytes = (buf) => `new Uint8Array(${JSON.stringify([...buf])}).buffer`;
  try {
    const work = path.join(outDir, 'work.pptx');
    fs.copyFileSync(src, work);
    win.webContents.send('app:open-path', work);
    await until(`!!(active && active.kind === 'pptx')`);
    await run(`setMode(active, 'edit')`);
    await until(`!!active.pptEditor && !!active.pptEditor.slideEl`);
    await wait(400);
    const menus = await run(`Menus.build().then(ms => Object.fromEntries(ms.map(m => [m.name, m.items.filter(i => !i.sep).map(i => i.label)])))`);
    check('menus have the new items', ['Video…', 'Audio…', 'Chart…', 'Diagram', 'Comment…'].every((x) => menus.Insert.includes(x)) && ['Crop image…', 'Image options…', 'Edit chart data…'].every((x) => menus.Format.includes(x)) && menus.Slide.includes('Theme builder…'), menus);

    // New blank slide to work on.
    await run(`${ED}.addSlide()`).catch(() => run(`${ED}.addSlideDialog()`));
    await wait(300);
    const si = await run(`${ED}.current`);

    // Picture + image options.
    await run(`(async () => { const c = document.createElement('canvas'); c.width = 200; c.height = 120; const g = c.getContext('2d'); g.fillStyle = '#e53935'; g.fillRect(0, 0, 200, 120); g.fillStyle = '#1e88e5'; g.fillRect(100, 0, 100, 120); const b = await (await new Promise(r => c.toBlob(r, 'image/png'))).arrayBuffer(); await ${ED}.insertPicture(b, 'png', 200, 120); return true; })()`);
    await pick(await lastIdx());
    await run(`${ED}.setImageOptions({ crop: { l: 10, t: 0, r: 20, b: 5 }, bright: 20, contrast: -10, transparency: 30, grey: true, reflection: true })`);
    let x = await slideXml(si);
    check('crop saved', /<a:srcRect l="10000" r="20000" b="5000"\/>/.test(x), (x.match(/<a:srcRect[^>]*>/) || [])[0]);
    check('adjustments saved', /<a:alphaModFix amt="70000"\/>/.test(x) && /<a:grayscl\/>/.test(x) && /<a:lum bright="20000" contrast="-10000"\/>/.test(x) && /<a:reflection /.test(x));
    check('adjustments drawn', await run(`(() => { const img = ${ED}.slideEl.querySelector('.pp-own:last-of-type img') || [...${ED}.slideEl.querySelectorAll('.pp-pic img')].pop(); return !!img && /grayscale/.test(img.style.filter) && /brightness/.test(img.style.filter) && +img.style.opacity < 0.8; })()`));
    check('reflection drawn', await run(`!!${ED}.slideEl.querySelector('.pp-reflect')`));

    // Video and audio.
    await run(`${ED}.insertMedia(${bytes(fs.readFileSync(video))}, 'clip.mp4')`);
    await until(`!!${ED}.slideEl.querySelector('video.pp-video')`, 8000);
    check('video shown with a play button', await run(`!!${ED}.slideEl.querySelector('video.pp-video') && !!${ED}.slideEl.querySelector('.pp-play')`));
    await run(`${ED}.insertMedia(${bytes(wav())}, 'tone.wav')`);
    check('audio shown', await until(`!!${ED}.slideEl.querySelector('audio.pp-audio')`, 8000));
    x = await slideXml(si);
    check('media saved as PowerPoint media', /<a:videoFile r:link="rId\d+"\/>/.test(x) && /<a:audioFile r:link="rId\d+"\/>/.test(x) && (x.match(/p14:media/g) || []).length >= 2);

    // Real chart and editing its data.
    await run(`${ED}.insertRealChart({ type: 'column', title: 'Sales', rows: [['Quarter', 'Sales', 'Costs'], ['Q1', '120', '80'], ['Q2', '140', '90'], ['Q3', '160', '95']] })`);
    await wait(500);
    check('chart drawn', await run(`!!${ED}.slideEl.querySelector('.pp-chart canvas')`));
    await pick(await lastIdx());
    const info = await run(`${ED}.chartAt().then(i => i && { rows: i.rows, type: i.type, title: i.title })`);
    check('chart data read back', info && info.title === 'Sales' && info.rows[1][0] === 'Q1' && info.rows[2][1] === '140', info);
    await run(`${ED}.editChartData({ type: 'line', title: 'Sales 2', rows: [['Quarter', 'Sales'], ['Q1', '10'], ['Q2', '20'], ['Q3', '30'], ['Q4', '40']] })`);
    const info2 = await run(`${ED}.chartAt().then(i => i && { rows: i.rows, type: i.type, title: i.title })`);
    check('chart data edited', info2 && info2.type === 'line' && info2.rows.length === 5 && info2.title === 'Sales 2', info2);

    // Diagrams on another new slide.
    await run(`${ED}.addSlide()`);
    await wait(300);
    const di = await run(`${ED}.current`);
    for (const [kind, items] of [['process', ['Plan', 'Build', 'Test']], ['timeline', ['Jan', 'Mar', 'Jun']], ['cycle', ['Plan', 'Do', 'Check', 'Act']], ['hierarchy', ['Boss', 'Design', '- UX', 'Code']]]) await run(`${ED}.insertDiagram(${JSON.stringify(kind)}, ${JSON.stringify(items)})`);
    x = await slideXml(di);
    check('diagrams inserted as groups', (x.match(/<p:grpSp>/g) || []).length === 4 && /prst="chevron"/.test(x) && /prst="ellipse"/.test(x) && /<p:cxnSp>/.test(x) && x.includes('>UX<'), (x.match(/<p:grpSp>/g) || []).length);
    await shot('diagrams.png');

    // Comments.
    await run(`${ED}.addComment('Check this slide')`);
    check('comment pin shown', await run(`!!${ED}.slideEl.querySelector('.pp-comment-pin')`));
    const comments = await run(`${ED}.slideComments()`);
    check('comment stored', comments.length === 1 && comments[0].text === 'Check this slide' && comments[0].author === 'Doc Studio', comments);

    // Theme builder.
    await run(`${ED}.buildTheme({ name: 'Test theme', colors: { accent1: '#123456', accent2: '#abcdef' }, heading: 'Georgia', body: 'Verdana', background: '#fafaf0', titleSize: 40, bodySize: 20 })`);
    const themeXml = await run(`(() => { const s = ${ED}.deck.slides[0]; return new XMLSerializer().serializeToString(${ED}.deck.docs.get(s.themePath)); })()`);
    const masterXml = await run(`(() => { const s = ${ED}.deck.slides[0]; return new XMLSerializer().serializeToString(${ED}.deck.docs.get(s.masterPath)); })()`);
    check('theme colours and fonts', /<a:accent1><a:srgbClr val="123456"\/>/.test(themeXml) && /<a:majorFont><a:latin typeface="Georgia"/.test(themeXml.replace(/>\s+</g, '><')) && /name="Test theme"/.test(themeXml));
    check('master background and sizes', /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FAFAF0"\/>/.test(masterXml) && /<p:titleStyle><a:lvl1pPr[^>]*>[\s\S]*?<a:defRPr[^>]*sz="4000"/.test(masterXml));

    // Slideshow shows the media with controls.
    await run(`${ED}.go(${si})`);
    await run(`(${ED}.present({}), true)`);
    check('slideshow plays media', await until(`!!document.querySelector('.ppe-show video[controls]')`, 8000));
    await run(`(document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })), true)`);
    await wait(600);

    await run(`save(active)`);
    await until(`!active.dirty`);
    fs.copyFileSync(work, path.join(outDir, 'edited.pptx'));
    const z = await JSZip.loadAsync(fs.readFileSync(work));
    check('package has the media, chart workbook and comments', Object.keys(z.files).some((f) => /^ppt\/media\/media\d+\.mp4$/.test(f)) && Object.keys(z.files).some((f) => /^ppt\/embeddings\/.*\.xlsx$/.test(f)) && !!z.file('ppt/commentAuthors.xml') && Object.keys(z.files).some((f) => /^ppt\/comments\/comment\d+\.xml$/.test(f)));
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 700)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
