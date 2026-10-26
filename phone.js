'use strict';
// Phone link: a small web page on this PC, reachable only from the home network while the Phone tool window is
// open. The phone opens it by scanning a QR code (no app to install), then sends photos and files to the PC and
// downloads the files the PC offers.
//
// Safety: a random one-time code is part of every address (requests without it are refused); only private network
// addresses are answered; the phone can only upload into a temporary folder and download what the PC chose to
// offer; files are limited to 200 MB; everything stops when the window closes or after 15 minutes unused.

const http = require('http');
const crypto = require('crypto');
const fs = require('fs');
const os = require('os');
const path = require('path');

const MAX_FILE = 200 * 1024 * 1024;
const IDLE = 15 * 60 * 1000;

let server = null;
let session = null; // { token, port, win, dir, offered: Map, received: Map, idle, seen }

// Addresses this PC has on the home network (Wi-Fi or Ethernet), most likely first.
function addresses() {
  const out = [];
  for (const [name, list] of Object.entries(os.networkInterfaces())) {
    if (/vEthernet|VirtualBox|VMware|WSL|Hyper-V|Loopback|Bluetooth|Teredo|isatap/i.test(name)) continue;
    for (const a of list || []) {
      if (a.family !== 'IPv4' || a.internal || !isPrivate(a.address) || a.address.startsWith('169.254.')) continue;
      out.push({ name, address: a.address, rank: /wi-?fi|wlan|wireless/i.test(name) ? 0 : /ethernet/i.test(name) ? 1 : 2 });
    }
  }
  return out.sort((x, y) => x.rank - y.rank).map(({ name, address }) => ({ name, address }));
}
function isPrivate(ip) {
  ip = String(ip || '').replace(/^::ffff:/, '');
  if (ip === '127.0.0.1' || ip === '::1') return true;
  const p = ip.split('.').map(Number);
  if (p.length === 4) return p[0] === 10 || (p[0] === 172 && p[1] >= 16 && p[1] <= 31) || (p[0] === 192 && p[1] === 168) || (p[0] === 169 && p[1] === 254);
  return /^f[cd]/i.test(ip) || /^fe80:/i.test(ip);
}
// A file name that is safe on Windows.
function safeName(n) {
  n = path.basename(String(n || '')).replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_').replace(/^\.+/, '').trim();
  if (!n) n = 'file';
  return n.length > 150 ? n.slice(0, 120) + path.extname(n).slice(0, 20) : n;
}
const send = (event, data) => { if (session && session.win && !session.win.isDestroyed()) session.win.webContents.send(event, data); };
function touch() {
  if (!session) return;
  clearTimeout(session.idle);
  session.idle = setTimeout(() => { send('phone:stopped', { reason: 'idle' }); stop(); }, IDLE);
}

async function start(win, page) {
  stop();
  const dir = path.join(os.tmpdir(), 'docstudio-phone', crypto.randomBytes(6).toString('hex'));
  await fs.promises.mkdir(dir, { recursive: true });
  session = { token: crypto.randomBytes(16).toString('hex'), win, dir, page, offered: new Map(), received: new Map(), seen: false };
  server = http.createServer(handle);
  server.requestTimeout = 0; // big uploads on a slow phone connection
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '0.0.0.0', resolve); });
  session.port = server.address().port;
  // Closing the Doc Studio window that opened it closes the link too.
  const owner = session;
  win.once('closed', () => { if (session === owner) stop(); });
  touch();
  return { port: session.port, token: session.token, addresses: addresses() };
}
function stop() {
  if (server) { server.close(); server.closeAllConnections?.(); server = null; }
  if (session) {
    clearTimeout(session.idle);
    fs.promises.rm(session.dir, { recursive: true, force: true }).catch(() => {});
    session = null;
  }
}

function reply(res, code, type, body, extra) {
  res.writeHead(code, {
    'Content-Type': type, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer',
    'Content-Security-Policy': "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; connect-src 'self'; img-src 'self' data:",
    ...(extra || {}),
  });
  res.end(body);
}

function handle(req, res) {
  if (!session || !isPrivate(req.socket.remoteAddress)) { req.destroy(); return; }
  const url = new URL(req.url, 'http://phone');
  const okToken = (url.searchParams.get('k') || '').length === session.token.length && crypto.timingSafeEqual(Buffer.from(url.searchParams.get('k')), Buffer.from(session.token));
  if (!okToken) { reply(res, 403, 'text/plain; charset=utf-8', 'This link has expired. Scan the QR code in Doc Studio again.'); return; }
  touch();
  const route = `${req.method} ${url.pathname}`;
  if (route === 'GET /') {
    if (!session.seen) { session.seen = true; }
    send('phone:connected', { from: String(req.socket.remoteAddress).replace(/^::ffff:/, ''), agent: String(req.headers['user-agent'] || '') });
    reply(res, 200, 'text/html; charset=utf-8', session.page);
    return;
  }
  if (route === 'GET /files') {
    const list = [...session.offered].map(([id, f]) => ({ id, name: f.name, size: f.size }));
    reply(res, 200, 'application/json', JSON.stringify(list));
    return;
  }
  if (route === 'GET /download') {
    const f = session.offered.get(url.searchParams.get('id'));
    if (!f) { reply(res, 404, 'text/plain; charset=utf-8', 'This file is no longer offered.'); return; }
    res.writeHead(200, {
      'Content-Type': 'application/octet-stream', 'Content-Length': f.size, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff',
      'Content-Disposition': `attachment; filename="${f.name.replace(/[^\x20-\x7e]|"/g, '_')}"; filename*=UTF-8''${encodeURIComponent(f.name)}`,
    });
    send('phone:downloaded', { id: url.searchParams.get('id'), name: f.name });
    if (f.data) res.end(Buffer.from(f.data)); else fs.createReadStream(f.path).pipe(res);
    return;
  }
  if (route === 'POST /upload') { upload(req, res, url); return; }
  reply(res, 404, 'text/plain; charset=utf-8', 'Not found');
}

function upload(req, res, url) {
  const name = safeName(url.searchParams.get('name'));
  const total = +req.headers['content-length'] || 0;
  if (total > MAX_FILE) { reply(res, 413, 'text/plain; charset=utf-8', 'This file is larger than 200 MB.'); req.resume(); return; }
  const id = crypto.randomBytes(6).toString('hex');
  const target = path.join(session.dir, `${id}-${name}`);
  const out = fs.createWriteStream(target);
  let got = 0, last = 0, failed = false;
  const dir = session.dir;
  send('phone:incoming', { id, name, size: total });
  const fail = (msg) => { if (failed) return; failed = true; out.destroy(); fs.promises.rm(target, { force: true }).catch(() => {}); send('phone:failed', { id, name, message: msg }); };
  req.on('data', (chunk) => {
    got += chunk.length;
    if (got > MAX_FILE) { fail('The file is larger than 200 MB.'); req.destroy(); return; }
    if (Date.now() - last > 250) { last = Date.now(); send('phone:progress', { id, got, size: total }); touch(); }
  });
  req.on('aborted', () => fail('The phone stopped sending.'));
  req.pipe(out);
  out.on('finish', () => {
    if (failed || !session || session.dir !== dir) return;
    session.received.set(id, { name, path: target, size: got });
    send('phone:received', { id, name, size: got });
    reply(res, 200, 'application/json', JSON.stringify({ ok: true }));
  });
  out.on('error', (err) => { fail(err.message); reply(res, 500, 'text/plain; charset=utf-8', 'The PC could not save the file.'); });
}

function register(ipcMain, { W, freeName, phonePage }) {
  ipcMain.handle('phone:start', (e) => start(W(e), phonePage()));
  ipcMain.handle('phone:stop', () => { stop(); });
  // PC -> phone: files the phone may download (bytes from an open document, or a file on disk).
  ipcMain.handle('phone:offer', async (_e, items) => {
    if (!session) throw new Error('The phone link is not open.');
    const out = [];
    for (const it of items) {
      const id = crypto.randomBytes(6).toString('hex');
      const name = safeName(it.name || (it.path && path.basename(it.path)));
      if (it.data) session.offered.set(id, { name, data: it.data, size: it.data.byteLength });
      else { const st = await fs.promises.stat(it.path); session.offered.set(id, { name, path: it.path, size: st.size }); }
      out.push({ id, name, size: session.offered.get(id).size });
    }
    return out;
  });
  ipcMain.handle('phone:unoffer', (_e, id) => { if (session) session.offered.delete(id); });
  // A received file's bytes (to open it as a new document).
  ipcMain.handle('phone:read', async (_e, id) => {
    const f = session && session.received.get(id);
    if (!f) throw new Error('This file is no longer here. Send it again from the phone.');
    const buf = await fs.promises.readFile(f.path);
    return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  });
  // Copy received files into a folder (never over an existing file).
  ipcMain.handle('phone:saveTo', async (_e, ids, folder) => {
    const saved = [];
    for (const id of ids) {
      const f = session && session.received.get(id);
      if (!f) continue;
      const target = await freeName(path.join(folder, f.name));
      await fs.promises.copyFile(f.path, target);
      saved.push(target);
    }
    return saved;
  });
  ipcMain.handle('phone:forget', (_e, id) => { const f = session && session.received.get(id); if (f) { session.received.delete(id); fs.promises.rm(f.path, { force: true }).catch(() => {}); } });
}

module.exports = { register, stop, addresses, isPrivate, safeName };
