'use strict';
// AI assistants: the MCP server (integrations/mcp/mcp-server.js) talking to this Doc Studio through the AI bridge.
// Reads, creates, converts, merges, splits, OCRs, edits a picture, and checks the file rules (new files only,
// copy first, AI files may change, an earlier session's file asks the owner, the off switch).
//   DOCSTUDIO_SRC=<folder with notes.md, rich.docx, report.pdf> DOCSTUDIO_OUT=<dir> electron tests/ai-test.js
const { app, BrowserWindow } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawn } = require('child_process');

const srcDir = process.env.DOCSTUDIO_SRC;
const outDir = process.env.DOCSTUDIO_OUT || path.join(os.tmpdir(), 'docstudio-ai-out');
const userData = fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-aitest-'));
app.setPath('userData', userData);
// AI assistants are off until the owner turns them on; this test's profile has them on.
fs.writeFileSync(path.join(userData, 'settings.json'), JSON.stringify({ aiEnabled: true }));
fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail });

// An MCP client: the server runs on this Electron in Node mode, as Claude Code and Codex run it.
function client() {
  const child = spawn(process.execPath, [path.join(__dirname, '..', 'integrations', 'mcp', 'mcp-server.js')], { env: { ...process.env, ELECTRON_RUN_AS_NODE: '1', DOC_STUDIO_DATA: userData, DOC_STUDIO_EXE: 'C:\\no\\such\\Doc Studio.exe' }, stdio: ['pipe', 'pipe', 'ignore'] });
  let buf = '', next = 1;
  const waiting = new Map();
  child.stdout.on('data', (c) => { buf += c; let i; while ((i = buf.indexOf('\n')) >= 0) { const m = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1); if (waiting.has(m.id)) { waiting.get(m.id)(m); waiting.delete(m.id); } } });
  const rpc = (method, params) => new Promise((r) => { const id = next++; waiting.set(id, r); child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n'); });
  const call = async (name, args) => { const r = (await rpc('tools/call', { name, arguments: args })).result; return { error: !!r.isError, text: r.content.map((c) => c.text).join('\n'), data: r.structuredContent }; };
  return { rpc, call, stop: () => child.stdin.end() };
}

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  for (let i = 0; i < 50 && !fs.existsSync(path.join(userData, 'ai-bridge.json')); i++) await wait(100);
  const run = (js) => win.webContents.executeJavaScript(js);
  const one = client();
  try {
    const out = (n) => path.join(outDir, n);
    const init = (await one.rpc('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1' } })).result;
    check('the server introduces itself', init.serverInfo.title === 'Doc Studio' && /new files/i.test(init.instructions));
    const tools = (await one.rpc('tools/list', {})).result.tools.map((t) => t.name);
    check('15 tools', tools.length === 15 && tools.includes('docstudio_edit_image'), tools);
    check('status answers', !(await one.call('docstudio_status', {})).error);
    const read = await one.call('docstudio_read_document', { path: path.join(srcDir, 'rich.docx') });
    check('reads a Word document as Markdown', !read.error && /^#/m.test(read.text), read.text.slice(0, 120));
    const made = await one.call('docstudio_create_document', { content: '# Plan\n\n- One\n- Two\n', output: out('plan.docx') });
    check('creates a Word document', !made.error && fs.readFileSync(out('plan.docx')).slice(0, 2).toString() === 'PK', made.text);
    const pdf = await one.call('docstudio_create_document', { content: '# Hello\n\nA PDF from Markdown.', output: out('hello.pdf') });
    check('creates a PDF', !pdf.error && fs.readFileSync(out('hello.pdf')).slice(0, 5).toString() === '%PDF-', pdf.text);
    const conv = await one.call('docstudio_convert', { source: path.join(srcDir, 'notes.md'), to: 'html', output: out('notes.html') });
    check('converts', !conv.error && /<html/i.test(fs.readFileSync(out('notes.html'), 'utf8')), conv.text);
    const merged = await one.call('docstudio_merge_pdfs', { sources: [path.join(srcDir, 'report.pdf'), out('hello.pdf')], output: out('merged.pdf') });
    check('merges PDFs', !merged.error && merged.data.pages >= 2, merged.text);
    const split = await one.call('docstudio_split_pdf', { source: out('merged.pdf'), ranges: '1, 2-end' });
    check('splits a PDF', !split.error && split.data.created.length === 2, split.text);
    const ocr = await one.call('docstudio_ocr', { source: out('hello.pdf') });
    check('reads text from a PDF page (OCR)', !ocr.error && /PDF from Markdown/i.test(ocr.text), ocr.text.slice(0, 120));
    const pic = await run(`(async () => { const c = document.createElement('canvas'); c.width = 300; c.height = 200; const g = c.getContext('2d'); g.fillStyle = '#7ab'; g.fillRect(0, 0, 300, 200); g.fillStyle = '#c22'; g.beginPath(); g.arc(150, 100, 60, 0, 7); g.fill(); return [...new Uint8Array(await new Promise((r) => c.toBlob(r, 'image/png')).then((b) => b.arrayBuffer()))]; })()`);
    fs.writeFileSync(out('ball.png'), Buffer.from(pic));
    const img = await one.call('docstudio_edit_image', { source: out('ball.png'), steps: [{ op: 'rotate', turns: 1 }, { op: 'remove_background', fill: 'transparent' }, { op: 'resize', width: 100 }], output: out('ball-cut.png') });
    check('edits a picture (rotate, remove background, resize)', !img.error && img.data.width === 100 && img.data.height === 150, img.text);
    // The rules.
    const over = await one.call('docstudio_create_document', { content: 'x', output: path.join(srcDir, 'notes.md') });
    check('will not overwrite a file it did not make', over.error && /copy it first/i.test(over.text));
    const copy = await one.call('docstudio_copy_document', { source: path.join(srcDir, 'notes.md'), output: out('notes-copy.md') });
    const fr = await one.call('docstudio_find_replace', { path: out('notes-copy.md'), find: 'Google', replace: 'GOOGLE', match_case: true });
    check('copies a file and edits the copy', !copy.error && !fr.error && fr.data.replacements > 0);
    check('the original is untouched', !/GOOGLE/.test(fs.readFileSync(path.join(srcDir, 'notes.md'), 'utf8')));
    const same = await one.call('docstudio_write_document', { path: out('plan.docx'), content: '# Plan v2' });
    check('rewrites its own file in the same session', !same.error, same.text);
    // Another session: asked; "Don't allow" refuses, "Allow for this file" lets it (and later calls) through.
    const two = client();
    await two.rpc('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test2', version: '1' } });
    const denied = two.call('docstudio_write_document', { path: out('plan.docx'), content: '# Plan v3' });
    for (let i = 0; i < 30 && !(await run(`!!document.querySelector('.msgbox')`)); i++) await wait(100);
    check('an earlier session\'s file asks the owner', /AI assistant wants to rewrite plan\.docx/.test(await run(`(document.querySelector('.msgbox') || {}).innerText || ''`)));
    await run(`[...document.querySelectorAll('.msgbox button')].find((b) => b.textContent === "Don't allow").click()`);
    check('"Don\'t allow" refuses', (await denied).error && /did not allow/.test((await denied).text));
    const allowed = two.call('docstudio_write_document', { path: out('plan.docx'), content: '# Plan v3' });
    for (let i = 0; i < 30 && !(await run(`!!document.querySelector('.msgbox')`)); i++) await wait(100);
    await run(`[...document.querySelectorAll('.msgbox button')].find((b) => b.textContent === 'Allow for this file').click()`);
    check('"Allow for this file" lets it change', !(await allowed).error);
    const again = await two.call('docstudio_write_document', { path: out('plan.docx'), content: '# Plan v4' });
    check('...and does not ask again', !again.error && !(await run(`!!document.querySelector('.msgbox')`)));
    // The off switch.
    await run(`desk.aiSetEnabled(false)`);
    const listening = fs.existsSync(path.join(userData, 'ai-bridge.json'));
    const off = await two.call('docstudio_status', {});
    await run(`desk.aiSetEnabled(true)`);
    check('the off switch blocks AI assistants', off.error && /turned off/.test(off.text), off.text);
    check('...and nothing listens while it is off', !listening);
    check('switching it on again listens again', fs.existsSync(path.join(userData, 'ai-bridge.json')));
    two.stop();
    // Settings > AI assistants.
    await run(`(document.querySelector('#btn-settings').click(), true)`);
    await wait(500);
    await run(`[...document.querySelectorAll('.settings-defaults button')].find((b) => b.textContent === 'Set up…').click()`);
    await wait(800);
    check('Settings > AI assistants lists the four apps (Claude Desktop: show the extension file)', (await run(`[...document.querySelectorAll('.ai-card-head b')].map((b) => b.textContent).join()`)) === 'Claude Code,Claude Desktop,Gemini CLI,OpenAI Codex' && (await run(`[...document.querySelectorAll('.ai-card button')].some((b) => b.textContent === 'Show the extension file')`)));
    check('Recent is not filled with the AI\'s files', (await run(`desk.recent()`)).every((f) => !f.includes(outDir)));
    fs.writeFileSync(path.join(outDir, 'ai.png'), (await win.webContents.capturePage()).toPNG());
  } catch (err) {
    check('no exceptions', false, err.stack);
  }
  one.stop();
  for (const c of checks) console.log(`${c.ok ? 'PASS' : 'FAIL'} ${c.name}${c.ok || c.detail === undefined ? '' : ` -- ${JSON.stringify(c.detail).slice(0, 600)}`}`);
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  app.exit(checks.every((c) => c.ok) ? 0 : 1);
});
