'use strict';
// Smoke test: run with `electron tests/smoke.js <samplesDir> <outDir>`.
// Opens every sample in the real app, checks for errors, exercises edit/save, and saves a screenshot per file.

const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const [samples, outDir] = process.argv.slice(2).filter((a) => !a.startsWith('-') && !a.endsWith('smoke.js'));
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-test-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const results = [];
const consoleErrors = [];

async function waitFor(win, expr, ms = 15000) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (await win.webContents.executeJavaScript(expr).catch(() => false)) return true;
    await wait(150);
  }
  return false;
}

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1200, 820);
  win.webContents.on('console-message', (_e, level, msg, line, src) => { if (level >= 3) consoleErrors.push(`${path.basename(src || '')}:${line} ${msg}`); });
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  await win.webContents.capturePage().then((img) => fs.writeFileSync(path.join(outDir, '00-welcome.png'), img.toPNG()));

  const files = fs.readdirSync(samples).filter((f) => !f.startsWith('saved-')).sort();
  for (const f of files) {
    console.log('OPEN', f);
    const n = await win.webContents.executeJavaScript('tabs.length');
    win.webContents.send('app:open-path', path.join(samples, f));
    await waitFor(win, `tabs.length > ${n} && !document.querySelector('.pane.active .loading')`);
    await wait(f.endsWith('.pdf') || f.endsWith('.pptx') ? 2500 : 900);
    const info = await win.webContents.executeJavaScript(`({ kind: active.kind, mode: active.mode, status: document.querySelector('#status-left').textContent, failure: !!document.querySelector('.pane.active .failure'), failText: document.querySelector('.pane.active .failure')?.innerText || '' })`);
    results.push({ file: f, ...info });
    const img = await win.webContents.capturePage();
    fs.writeFileSync(path.join(outDir, `${f.replace(/\./g, '_')}.png`), img.toPNG());
  }

  // Edit + save round trips.
  const run = (js) => win.webContents.executeJavaScript(js).catch((err) => { console.log('STEP FAILED:', js.slice(0, 160)); throw err; });
  // Markdown: type, save, check file.
  await run(`activate(tabs.find(t => t.name === 'notes.md')); active.textarea.focus(); active.textarea.setSelectionRange(0,0); document.execCommand('insertText', false, 'Added line\\n'); true`);
  await run(`active.path = active.path.replace('notes.md', 'saved-notes.md'); save(active)`);
  const md = fs.readFileSync(path.join(samples, 'saved-notes.md'), 'utf8');
  results.push({ file: 'saved-notes.md', ok: md.startsWith('Added line\n# Notes') && !md.includes('\r\n') });

  // Text in windows-1252 with CRLF keeps its encoding and line endings.
  await run(`activate(tabs.find(t => t.name === 'readme.txt')); active.textarea.focus(); active.textarea.setSelectionRange(active.textarea.value.length, active.textarea.value.length); document.execCommand('insertText', false, 'more café'); active.path = active.path.replace('readme.txt', 'saved-readme.txt'); save(active)`);
  const txt = fs.readFileSync(path.join(samples, 'saved-readme.txt'));
  results.push({ file: 'saved-readme.txt', ok: txt.includes(Buffer.from('caf\xe9', 'latin1')) && txt.includes(Buffer.from('\r\n')) && !txt.includes(Buffer.from([0xc3, 0xa9])) });

  // Word: switch to edit, type, save as a copy, check the saved XML.
  await run(`activate(tabs.find(t => t.name === 'report.docx')); setMode(active, 'edit')`);
  await waitFor(win, `!!active.editor`);
  await wait(500);
  await win.webContents.capturePage().then((img) => fs.writeFileSync(path.join(outDir, 'report_docx-edit.png'), img.toPNG()));
  await run(`(() => { const p = [...active.editor.stage.querySelectorAll('article p')].find(p => p.textContent.startsWith('This is a')); const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); p.closest('article').focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, 'EDITED: '); active.editor.opts.onChange(); return true; })()`);
  await run(`active.path = active.path.replace('report.docx', 'saved-report.docx'); save(active)`);
  const savedDocx = path.join(samples, 'saved-report.docx');
  const JSZip = require('jszip');
  const xml = fs.existsSync(savedDocx) ? await (await JSZip.loadAsync(fs.readFileSync(savedDocx))).file('word/document.xml').async('string') : '';
  const origXml = await (await JSZip.loadAsync(fs.readFileSync(path.join(samples, 'report.docx')))).file('word/document.xml').async('string');
  results.push({
    file: 'saved-report.docx',
    ok: xml.includes('EDITED: ') && xml.includes('Quarterly Report') && xml.includes('North') && xml.includes('Second point'),
    italicKept: xml.includes('<w:i/>'), colourKept: /C0392B/i.test(xml), table: xml.includes('<w:tbl>'), listStyle: xml.includes('w:val="ListBullet"'),
    sameTable: (origXml.match(/<w:tbl>[\s\S]*<\/w:tbl>/) || [''])[0] === (xml.match(/<w:tbl>[\s\S]*<\/w:tbl>/) || ['x'])[0],
  });
  await run(`setMode(active, 'view')`);
  await wait(1500);
  await win.webContents.capturePage().then((img) => fs.writeFileSync(path.join(outDir, 'report_docx-after-save.png'), img.toPNG()));

  // Markdown split view + find.
  await run(`activate(tabs.find(t => t.name === 'saved-notes.md')); openFind(); document.querySelector('#find-input').value = 'needle'; runFind(); true`);
  await wait(300);
  results.push({ file: 'find-md', count: await run(`document.querySelector('#find-count').textContent`) });
  await win.webContents.capturePage().then((img) => fs.writeFileSync(path.join(outDir, 'find-md.png'), img.toPNG()));
  await run(`activate(tabs.find(t => t.name === 'config.json')); openFind(); document.querySelector('#find-input').value = '1'; runFind(); stepFind(1); true`);
  await wait(300);
  await win.webContents.capturePage().then((img) => fs.writeFileSync(path.join(outDir, 'find-text.png'), img.toPNG()));
  await run(`activate(tabs.find(t => t.name === 'report.docx' || t.name === 'saved-report.docx')); openFind(); document.querySelector('#find-input').value = 'needle'; runFind(); true`);
  await wait(300);
  results.push({ file: 'find-docx', count: await run(`document.querySelector('#find-count').textContent`) });

  // New docx and export to PDF via print path.
  await run(`newDocument('docx'); true`);
  await waitFor(win, `!!active.editor`);
  await run(`(() => { const a = active.editor.articles[0]; const r = document.createRange(); r.selectNodeContents(a.querySelector('p')); r.collapse(true); a.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); document.execCommand('insertText', false, 'Brand new document'); return true; })()`);
  await wait(300);
  await win.webContents.capturePage().then((img) => fs.writeFileSync(path.join(outDir, 'new-docx.png'), img.toPNG()));
  await run(`active.beforePrint ? active.beforePrint() : true`);
  const prepared = await run(`preparePrint()`);
  const pdf = await win.webContents.printToPDF({ printBackground: true });
  await run(`endPrint()`);
  fs.writeFileSync(path.join(outDir, 'export.pdf'), pdf);
  results.push({ file: 'export.pdf', ok: prepared && pdf.length > 1000 });

  // Dark theme screenshot of markdown.
  require('electron').nativeTheme.themeSource = 'dark';
  await run(`activate(tabs.find(t => t.name === 'saved-notes.md')); closeFind(); true`);
  await wait(500);
  await win.webContents.capturePage().then((img) => fs.writeFileSync(path.join(outDir, 'dark-md.png'), img.toPNG()));

  fs.writeFileSync(path.join(outDir, 'results.json'), JSON.stringify({ results, consoleErrors }, null, 2));
  console.log(JSON.stringify({ results, consoleErrors }, null, 2));
  app.exit(0);
});
