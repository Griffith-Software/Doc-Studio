'use strict';
// AI bridge: lets AI assistants (Claude, Gemini, Codex) use Doc Studio through the Doc Studio MCP server
// (integrations/mcp/mcp-server.js). The server asks this running Doc Studio to do the work.
//
// Local only: an HTTP endpoint on 127.0.0.1 with a random port and a secret key, both written to
// <userData>/ai-bridge.json while Doc Studio runs (removed when it quits). Work that needs the document engines
// (converting, reading, OCR, pictures, PDFs) runs in a hidden worker window (renderer/ai-worker.js), so the
// owner's open documents are never touched.
//
// The file rules the owner chose:
//   - The AI writes only NEW files; when it names no output, the result gets a free name next to the source.
//   - It may copy an existing file to a new file and work on the copy.
//   - It may change a file only if an AI assistant created it (ai-files.json). In the same AI session that is
//     allowed; a file from an earlier session needs the owner's permission (Allow / Allow for this file / Don't).
//   - It never deletes, and never sends anything off this PC.

const http = require('http');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { app, BrowserWindow } = require('electron');

let server = null;
let token = null;
let worker = null;
let workerReady = null;
let queue = Promise.resolve();
let ctx = null; // { W, mainWindow, freeName, version }

const bridgeFile = () => path.join(app.getPath('userData'), 'ai-bridge.json');
const registryFile = () => path.join(app.getPath('userData'), 'ai-files.json');
const keyOf = (p) => path.resolve(p).toLowerCase();

// ---------------------------------------------------------------- AI-created files

function registry() { try { return JSON.parse(fs.readFileSync(registryFile(), 'utf8')); } catch { return {}; } }
function remember(file, session) {
  const r = registry();
  r[keyOf(file)] = { path: path.resolve(file), session, created: r[keyOf(file)] ? r[keyOf(file)].created : new Date().toISOString(), approved: r[keyOf(file)] ? !!r[keyOf(file)].approved : false };
  fs.writeFileSync(registryFile(), JSON.stringify(r, null, 2));
}
function approve(file) { const r = registry(); if (r[keyOf(file)]) { r[keyOf(file)].approved = true; fs.writeFileSync(registryFile(), JSON.stringify(r, null, 2)); } }

class AiError extends Error { constructor(code, message) { super(message); this.code = code; } }

// Asks the owner in Doc Studio's window. Resolves to 0 (allow once), 1 (allow for this file) or 2 (don't).
async function askOwner(file, what) {
  const w = ctx.mainWindow();
  if (!w) return 2;
  if (w.isMinimized()) w.restore();
  w.show(); w.focus();
  const opts = { type: 'warning', message: `An AI assistant wants to ${what} ${path.basename(file)}`, detail: `${file}\n\nAn AI assistant created this file in an earlier session. Allow it to change the file now?`, buttons: ['Allow once', 'Allow for this file', "Don't allow"], defaultId: 2, cancelId: 2 };
  const answer = w.webContents.executeJavaScript(`desk.confirm(${JSON.stringify(opts)})`);
  // No answer within 3 minutes counts as "don't".
  return Promise.race([answer, new Promise((r) => setTimeout(() => r(2), 180000))]);
}

// Where a result may be written. Returns the absolute path to write.
async function writable(target, session, what) {
  if (typeof target !== 'string' || !path.isAbsolute(target)) throw new AiError('bad_path', 'Give the output as a full path (for example C:\\Users\\me\\Documents\\report.pdf).');
  target = path.resolve(target);
  if (!path.extname(target)) throw new AiError('bad_path', 'The output needs a file extension (such as .docx, .pdf or .png).');
  if (!fs.existsSync(path.dirname(target))) throw new AiError('bad_path', `The folder ${path.dirname(target)} does not exist.`);
  if (!fs.existsSync(target)) return target;
  const entry = registry()[keyOf(target)];
  if (!entry) throw new AiError('exists', `${path.basename(target)} already exists and was not created by an AI assistant, so it can't be changed. Write to a new file, or copy it first with docstudio_copy_document and work on the copy.`);
  if (entry.session === session || entry.approved) return target;
  const r = await askOwner(target, what || 'change');
  if (r === 2) throw new AiError('declined', `The owner did not allow changing ${path.basename(target)}. Write to a new file instead.`);
  if (r === 1) approve(target);
  return target;
}
// A new file next to the source (or in the Documents folder): "name.ext", "name (2).ext"...
async function freshNear(source, ext, suffix) {
  const dir = source ? path.dirname(source) : path.join(app.getPath('documents'));
  const base = source ? path.basename(source, path.extname(source)) : 'Doc Studio';
  return ctx.freeName(path.join(dir, `${base}${suffix || ''}.${ext.replace(/^\./, '')}`));
}
async function writeOut(target, data, session) {
  const tmp = `${target}.docstudio-tmp`;
  await fs.promises.writeFile(tmp, Buffer.from(data, 'base64'));
  await fs.promises.rename(tmp, target);
  remember(target, session);
  return target;
}

// ---------------------------------------------------------------- the hidden worker window

function startWorker() {
  if (worker && !worker.isDestroyed()) return workerReady;
  worker = new BrowserWindow({
    show: false, width: 1200, height: 900,
    webPreferences: { preload: path.join(__dirname, '..', 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true, backgroundThrottling: false },
  });
  worker.docStudioWorker = true;
  worker.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  workerReady = new Promise((resolve, reject) => {
    worker.webContents.once('did-finish-load', () => setTimeout(resolve, 300));
    worker.webContents.once('did-fail-load', (_e, _c, d) => reject(new Error(d)));
  });
  worker.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'), { query: { worker: '1' } });
  worker.on('closed', () => { worker = null; });
  return workerReady;
}
// Closing the owner's last window closes the worker too, so Doc Studio quits as before.
function onWindowClosed() {
  const visible = BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed() && !w.docStudioWorker);
  if (!visible.length && worker && !worker.isDestroyed()) worker.destroy();
}
async function inWorker(op, args) {
  await startWorker();
  const r = await worker.webContents.executeJavaScript(`AIWorker.run(${JSON.stringify(op)}, ${JSON.stringify(args)})`);
  if (r && r.error) throw new AiError(r.code || 'failed', r.error);
  return r;
}

// ---------------------------------------------------------------- tools

const exists = (p) => typeof p === 'string' && path.isAbsolute(p) && fs.existsSync(p) && fs.statSync(p).isFile();
function needFile(p, what) { if (!exists(p)) throw new AiError('not_found', `${what || 'The file'} was not found: ${p}`); return path.resolve(p); }
const extOf = (p) => path.extname(p).slice(1).toLowerCase();

const TOOLS = {
  async status() {
    const w = ctx.mainWindow();
    const open = w ? await w.webContents.executeJavaScript('AIWorker.openDocuments()').catch(() => []) : [];
    return { app: 'Doc Studio', version: ctx.version, open_documents: open, documents_folder: app.getPath('documents') };
  },
  async list_open_documents() {
    const w = ctx.mainWindow();
    return { documents: w ? await w.webContents.executeJavaScript('AIWorker.openDocuments()') : [] };
  },
  async read_document(a) {
    if (a.open_document !== undefined && a.open_document !== null && a.open_document !== '') {
      const w = ctx.mainWindow();
      if (!w) throw new AiError('not_running', 'Doc Studio has no window open.');
      const r = await w.webContents.executeJavaScript(`AIWorker.readOpen(${JSON.stringify(a.open_document)}, ${JSON.stringify(a.max_chars || 0)})`);
      if (r && r.error) throw new AiError(r.code || 'failed', r.error);
      return r;
    }
    return inWorker('read', { path: needFile(a.path), max_chars: a.max_chars });
  },
  async create_document(a, session) {
    const ext = a.output ? extOf(a.output) : String(a.format || 'md').toLowerCase();
    if (!['md', 'txt', 'html', 'docx', 'pdf'].includes(ext)) throw new AiError('bad_format', 'Doc Studio can create .md, .txt, .html, .docx and .pdf documents from Markdown. To get another format, create one of these and use docstudio_convert.');
    // No output given: a new file named after the title in the Documents folder.
    const target = a.output ? await writable(a.output, session, 'replace') : await ctx.freeName(path.join(app.getPath('documents'), `${(a.title || 'Untitled').replace(/[<>:"/\\|?*]/g, '_').slice(0, 80)}.${ext}`));
    const r = await inWorker('create', { ext, content: String(a.content || ''), title: a.title || path.basename(target, path.extname(target)) });
    return { created: await writeOut(target, r.data, session), bytes: Buffer.byteLength(r.data, 'base64') };
  },
  async write_document(a, session) {
    const file = path.resolve(String(a.path || ''));
    const target = await writable(file, session, 'rewrite');
    const ext = extOf(target);
    if (!['md', 'txt', 'html', 'docx', 'pdf'].includes(ext)) throw new AiError('bad_format', 'docstudio_write_document writes .md, .txt, .html, .docx and .pdf from Markdown.');
    const r = await inWorker('create', { ext, content: String(a.content || ''), title: a.title || path.basename(target, path.extname(target)) });
    return { written: await writeOut(target, r.data, session) };
  },
  async find_replace(a, session) {
    const file = needFile(a.path);
    if (!['md', 'markdown', 'txt', 'html', 'htm', 'csv', 'tsv', 'json', 'xml', 'log'].includes(extOf(file))) throw new AiError('bad_format', 'Find and replace works on text files (.md, .txt, .html, .csv, .json, .xml). For other documents, write a new version with docstudio_write_document.');
    const target = await writable(file, session, 'edit');
    const text = await fs.promises.readFile(target, 'utf8');
    if (!a.find) throw new AiError('bad_args', 'Say what text to find.');
    const flags = a.match_case ? 'g' : 'gi';
    const re = new RegExp(String(a.find).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), a.all === false ? flags.replace('g', '') : flags);
    let count = 0;
    const out = text.replace(re, () => { count++; return String(a.replace ?? ''); });
    if (!count) return { changed: false, replacements: 0 };
    await writeOut(target, Buffer.from(out, 'utf8').toString('base64'), session);
    return { changed: true, replacements: count, path: target };
  },
  async copy_document(a, session) {
    const src = needFile(a.source, 'The source file');
    const target = a.output ? await writable(a.output, session, 'replace') : await freshNear(src, extOf(src), ' (copy)');
    if (keyOf(target) === keyOf(src)) throw new AiError('bad_path', 'Copy to a different file.');
    await fs.promises.copyFile(src, target);
    remember(target, session);
    return { copied_to: target };
  },
  async convert(a, session) {
    const src = needFile(a.source, 'The source file');
    const to = String(a.to || '').replace(/^\./, '').toLowerCase();
    if (!to) throw new AiError('bad_args', 'Say which format to convert to, for example "pdf" or "docx".');
    const r = await inWorker('convert', { path: src, to });
    if (r.files) {
      const first = a.output ? await writable(a.output, session, 'replace') : await freshNear(src, to);
      const stem = first.slice(0, -path.extname(first).length);
      const written = [];
      for (const f of r.files) written.push(await writeOut(await ctx.freeName(`${stem}${f.suffix}.${to}`), f.data, session));
      return { created: written };
    }
    const target = a.output ? await writable(a.output, session, 'replace') : await freshNear(src, to);
    return { created: await writeOut(target, r.data, session) };
  },
  async merge_pdfs(a, session) {
    const sources = (a.sources || []).map((p) => needFile(p));
    if (sources.length < 2) throw new AiError('bad_args', 'Give at least two files to merge.');
    const r = await inWorker('merge', { paths: sources });
    const target = a.output ? await writable(a.output, session, 'replace') : await freshNear(sources[0], 'pdf', ' (merged)');
    return { created: await writeOut(target, r.data, session), pages: r.pages };
  },
  async split_pdf(a, session) {
    const src = needFile(a.source, 'The PDF');
    const r = await inWorker('split', { path: src, ranges: a.ranges, every: a.every });
    const dir = a.output_folder ? path.resolve(a.output_folder) : path.dirname(src);
    if (!fs.existsSync(dir)) throw new AiError('bad_path', `The folder ${dir} does not exist.`);
    const written = [];
    for (const f of r.files) written.push(await writeOut(await ctx.freeName(path.join(dir, `${path.basename(src, '.pdf')} (${f.label}).pdf`)), f.data, session));
    return { created: written };
  },
  async compress(a, session) {
    const src = needFile(a.source, 'The file');
    const r = await inWorker('compress', { path: src, level: a.level, quality: a.quality, max_side: a.max_side, format: a.format });
    if (r.notSmaller) return { created: null, message: 'The file is already as small as these settings can make it.', before: r.before, after: r.after };
    const target = a.output ? await writable(a.output, session, 'replace') : await freshNear(src, r.ext, ' (compressed)');
    return { created: await writeOut(target, r.data, session), before: r.before, after: r.after };
  },
  async ocr(a) { return inWorker('ocr', { path: needFile(a.source, 'The scan, photo or PDF') }); },
  async edit_image(a, session) {
    const src = needFile(a.source, 'The picture');
    const format = String(a.format || extOf(a.output || '') || extOf(src)).replace(/^\./, '').toLowerCase();
    const r = await inWorker('image', { path: src, steps: a.steps || [], format });
    const target = a.output ? await writable(a.output, session, 'replace') : await freshNear(src, r.ext, ' (edited)');
    return { created: await writeOut(target, r.data, session), width: r.width, height: r.height };
  },
  async scan(a, session) {
    const make = String(a.make || 'pdf').toLowerCase();
    const r = await inWorker('scan', { make, color: a.color, dpi: a.dpi, paper: a.paper, scanner: a.scanner });
    if (make === 'text') return { text: r.text };
    const target = a.output ? await writable(a.output, session, 'replace') : await ctx.freeName(path.join(app.getPath('documents'), `Scan ${new Date().toISOString().slice(0, 16).replace('T', ' ').replace(':', '.')}.${r.ext}`));
    return { created: await writeOut(target, r.data, session) };
  },
  async show(a) {
    const file = needFile(a.path);
    const w = ctx.mainWindow();
    if (!w) throw new AiError('not_running', 'Doc Studio has no window open.');
    if (w.isMinimized()) w.restore();
    w.show(); w.focus();
    w.webContents.send('app:open-path', file);
    return { shown: file };
  },
};

// ---------------------------------------------------------------- the endpoint

function handle(req, res) {
  const send = (code, body) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(body)); };
  const remote = String(req.socket.remoteAddress || '');
  if (!/^(127\.0\.0\.1|::1|::ffff:127\.0\.0\.1)$/.test(remote)) { req.destroy(); return; }
  const auth = String(req.headers.authorization || '').replace(/^Bearer /, '');
  if (auth.length !== token.length || !crypto.timingSafeEqual(Buffer.from(auth), Buffer.from(token))) { send(401, { ok: false, code: 'unauthorized', error: 'Wrong key.' }); return; }
  if (req.method !== 'POST' || req.url !== '/call') { send(404, { ok: false, error: 'Not found' }); return; }
  const parts = [];
  let size = 0;
  req.on('data', (c) => { size += c.length; if (size > 50 * 1024 * 1024) req.destroy(); else parts.push(c); });
  req.on('end', () => {
    let body;
    try { body = JSON.parse(Buffer.concat(parts).toString('utf8')); } catch { send(400, { ok: false, error: 'Bad request' }); return; }
    if (ctx.enabled && !ctx.enabled()) { send(200, { ok: false, code: 'disabled', error: 'AI assistants are turned off in Doc Studio (Settings > AI assistants). Ask the owner to turn them on.' }); return; }
    const fn = TOOLS[body.tool];
    if (!fn) { send(404, { ok: false, code: 'unknown_tool', error: `Unknown tool ${body.tool}` }); return; }
    // One job at a time: the worker window does one thing at once.
    const job = queue.then(() => fn(body.args || {}, String(body.session || 'unknown')));
    queue = job.catch(() => {});
    job.then((result) => send(200, { ok: true, result }), (err) => send(200, { ok: false, code: err.code || 'failed', error: err.message }));
  });
}

// The endpoint listens only while AI assistants are switched on (Settings > AI assistants; off until the owner
// turns them on), so nothing is listening otherwise.
async function start(options) {
  ctx = options;
  rememberInstall();
  app.on('will-quit', stop);
  if (ctx.enabled()) await listen();
}
async function listen() {
  if (server) return;
  token = crypto.randomBytes(24).toString('hex');
  server = http.createServer(handle);
  server.requestTimeout = 0;
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); });
  fs.writeFileSync(bridgeFile(), JSON.stringify({ port: server.address().port, token, pid: process.pid, version: ctx.version, exe: process.execPath }, null, 2));
}
const setEnabled = (on) => (on ? listen() : stop());
// Where this Doc Studio is installed, so the AI packages work in any install folder and for any Windows user:
// <userData>/install.json tells the MCP server which program to start when Doc Studio is closed, and the Claude
// Code plugin's .mcp.json (in resources\ai, which this program owns) is pointed at this Doc Studio.exe.
// The plugin's .mcp.json itself keeps the standard per-user path (${LOCALAPPDATA}/Programs/Doc Studio), so the
// program folder is never changed; another install folder is found through install.json.
function rememberInstall() {
  try { fs.writeFileSync(path.join(app.getPath('userData'), 'install.json'), JSON.stringify({ exe: process.execPath, version: ctx.version }, null, 2)); } catch { /* not kept */ }
}

function stop() {
  if (server) { server.close(); server = null; }
  try { const f = JSON.parse(fs.readFileSync(bridgeFile(), 'utf8')); if (f.pid === process.pid) fs.rmSync(bridgeFile(), { force: true }); } catch { /* gone */ }
}

module.exports = { start, stop, setEnabled, onWindowClosed, TOOLS, writable, registry };
