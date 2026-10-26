'use strict';
// Markdown Split view: both panes scroll together. DOCSTUDIO_SRC=<long.md> DOCSTUDIO_OUT=<dir> electron tests/md-scroll-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const src = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-md-')));
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
  try {
    const work = path.join(outDir, 'long.md');
    fs.copyFileSync(src, work);
    win.webContents.send('app:open-path', work);
    await wait(2000);
    check('opens in Preview', await run(`active.mode === 'preview'`));
    await run(`setMode(active, 'split')`);
    await wait(600);
    check('scroll map built', await run(`Array.isArray(active.scrollMap) && active.scrollMap.length > 5`), await run(`active.scrollMap && active.scrollMap.length`));
    // The heading at the top of the editor should be the heading at the top of the preview.
    const topHeadings = `(() => {
      const ta = active.textarea, pane = active.previewPane;
      const lh = parseFloat(getComputedStyle(ta).lineHeight);
      const pvTop = pane.getBoundingClientRect().top;
      const hs = [...pane.querySelectorAll('h1,h2,h3')];
      const firstVisible = hs.find((x) => x.getBoundingClientRect().top >= pvTop - 4);
      return { pv: firstVisible && firstVisible.textContent.trim(), ta: ta.scrollTop, pane: pane.scrollTop };
    })()`;
    // Editor -> preview: put a heading line at the top of the editor.
    const r1 = await run(`(async () => {
      const ta = active.textarea; const lines = ta.value.split('\\n');
      const idx = lines.findIndex((l, i) => i > lines.length / 2 && /^##\\s/.test(l));
      const target = lines[idx].replace(/^#+\\s*/, '').trim();
      ta.focus(); ta.setSelectionRange(0, 0);
      // scroll so that heading line is at the top: measure with the scroll map's mirror approach
      const before = ta.value.split('\\n').slice(0, idx).join('\\n');
      const m = active.measure; m.textContent = before + '\\n'; const mark = document.createElement('span'); m.append(mark);
      const y = mark.offsetTop - parseFloat(getComputedStyle(ta).paddingTop); m.replaceChildren();
      ta.scrollTop = y;
      await new Promise(r => setTimeout(r, 300));
      const pane = active.previewPane; const h = [...pane.querySelectorAll('h2')].find(x => x.textContent.trim() === target);
      return { target, offset: Math.round(h.getBoundingClientRect().top - pane.getBoundingClientRect().top), pane: pane.scrollTop };
    })()`);
    check('editor scroll brings the same heading to the top of the preview', Math.abs(r1.offset - 16) < 40, r1);
    // Preview -> editor: scroll the preview to another heading.
    const r2 = await run(`(async () => {
      const pane = active.previewPane; const hs = [...pane.querySelectorAll('h2')]; const h = hs[1];
      pane.scrollTop = h.getBoundingClientRect().top - pane.getBoundingClientRect().top + pane.scrollTop - 16;
      await new Promise(r => setTimeout(r, 300));
      const ta = active.textarea; const lines = ta.value.split('\\n');
      const idx = lines.findIndex(l => /^##\\s/.test(l) && l.replace(/^#+\\s*/, '').trim() === h.textContent.trim());
      const m = active.measure; m.textContent = lines.slice(0, idx).join('\\n') + '\\n'; const mark = document.createElement('span'); m.append(mark);
      const y = mark.offsetTop - parseFloat(getComputedStyle(ta).paddingTop); m.replaceChildren();
      return { heading: h.textContent.trim(), expected: Math.round(y), actual: Math.round(ta.scrollTop) };
    })()`);
    check('preview scroll moves the editor to the same heading', Math.abs(r2.expected - r2.actual) < 30, r2);
    // Bottom stays in sync.
    const r3 = await run(`(async () => { const ta = active.textarea, pane = active.previewPane; ta.scrollTop = ta.scrollHeight; await new Promise(r => setTimeout(r, 300)); return { pane: Math.round(pane.scrollTop), max: pane.scrollHeight - pane.clientHeight }; })()`);
    check('editor at the bottom puts the preview at the bottom', Math.abs(r3.pane - r3.max) < 3, r3);
    const r4 = await run(`(async () => { const ta = active.textarea, pane = active.previewPane; pane.scrollTop = 0; await new Promise(r => setTimeout(r, 300)); return { ta: ta.scrollTop }; })()`);
    check('preview at the top puts the editor at the top', r4.ta < 2, r4);
    fs.writeFileSync(path.join(outDir, 'split.png'), (await win.webContents.capturePage()).toPNG());
  } catch (err) {
    check('no exceptions', false, err.message);
  }
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 600)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
