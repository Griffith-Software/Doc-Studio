'use strict';
// Doc Studio main process: window, menu, file access and the Node-only converters
// (DOCX <-> HTML for editing, legacy .doc text extraction).

const { app, BrowserWindow, Menu, dialog, ipcMain, shell, nativeTheme, clipboard } = require('electron');
const path = require('path');
const fs = require('fs');

app.setName('Doc Studio');
// Windows' occlusion detection can wrongly decide the window is covered (when it is hosted by another program,
// screen-shared or remote) and Chromium then stops drawing it, leaving a blank window. Draw it regardless.
app.commandLine.appendSwitch('disable-features', 'CalculateNativeWinOcclusion');

const MAX_RECENT = 12;
let win = null;
const aiBridge = require('./ai/bridge');
// The AI assistants' hidden worker window (ai/bridge.js).
const isWorker = (e) => { const w = BrowserWindow.fromWebContents(e.sender); return !!(w && w.docStudioWorker); };
let pendingOpen = filesFromArgv(process.argv).filter((f) => !sampleExt(f));

// ---------- single instance: a second launch opens its files in the running window

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv) => {
    const all = filesFromArgv(argv);
    // A sample file from Default apps: the owner just chose an app for its type in File Explorer.
    for (const f of all.filter(sampleExt)) BrowserWindow.getAllWindows().forEach((bw) => bw.webContents.send('app:default-picked', sampleExt(f)));
    const files = all.filter((f) => !sampleExt(f));
    const w = BrowserWindow.getFocusedWindow() || win;
    if (w) {
      if (w.isMinimized()) w.restore();
      w.focus();
      files.forEach((f) => w.webContents.send('app:open-path', f));
    }
  });
  app.whenReady().then(() => {
    migrateSettings();
    createWindow();
    // AI assistants (ai/bridge.js): a local, key-protected endpoint for the Doc Studio MCP server.
    // Per-type icons in File Explorer (shell-icons.js).
    require('./shell-icons').register();
    aiBridge.start({ mainWindow: () => (win && !win.isDestroyed() ? win : BrowserWindow.getAllWindows().find((x) => !x.isDestroyed() && !x.docStudioWorker) || null), freeName, version: app.getVersion(), enabled: aiOn }).catch((err) => console.error('AI bridge', err));
  });
}

// Default apps sample files (see 'defaults:reveal'); returns the file type, or '' for any other file.
function sampleExt(f) {
  const m = /[\\/]docstudio-defaults[\\/]Choose the app for \.(\w+) files\.\w+$/i.exec(f);
  return m ? m[1].toLowerCase() : '';
}

function filesFromArgv(argv) {
  const args = argv.slice(app.isPackaged ? 1 : 2);
  return args.filter((a) => !a.startsWith('-') && fs.existsSync(a) && fs.statSync(a).isFile())
    .map((a) => path.resolve(a));
}

// ---------- settings (recent files, theme)

const settingsFile = () => path.join(app.getPath('userData'), 'settings.json');
function readSettings() {
  try { return JSON.parse(fs.readFileSync(settingsFile(), 'utf8')); } catch { return {}; }
}
function writeSettings(s) {
  try { fs.writeFileSync(settingsFile(), JSON.stringify(s, null, 2)); } catch { /* not fatal */ }
}
function addRecent(file) {
  const s = readSettings();
  s.recent = [file, ...(s.recent || []).filter((f) => f.toLowerCase() !== file.toLowerCase())].slice(0, MAX_RECENT);
  writeSettings(s);
  app.addRecentDocument(file);
  buildMenu();
}

// ---------- windows (the main window plus any documents popped out into their own)

// The window that sent an IPC request (falls back to the main window).
const W = (e) => (e && e.sender && BrowserWindow.fromWebContents(e.sender)) || BrowserWindow.getFocusedWindow() || win;
nativeTheme.on('updated', () => { for (const w of BrowserWindow.getAllWindows()) if (!w.isDestroyed() && w.setTitleBarOverlay) { try { w.setTitleBarOverlay(overlayColors()); } catch { /* window without an overlay */ } } });

// opts: { bounds, adopt } - a pop-out window gets its own bounds and the document it adopts.
function createWindow(opts) {
  opts = opts || {};
  const s = readSettings();
  nativeTheme.themeSource = s.theme || 'system';
  const w = new BrowserWindow({
    width: 1200,
    height: 820,
    ...(opts.bounds || {}),
    minWidth: 640,
    minHeight: 420,
    title: 'Doc Studio',
    icon: path.join(__dirname, 'build', 'icon.png'),
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#14161b' : '#f4f5f7',
    // The app bar is drawn by the page; Windows still draws minimise/maximise/close over it.
    titleBarStyle: 'hidden',
    titleBarOverlay: overlayColors(),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      plugins: true, // Chromium's built-in PDF viewer
      spellcheck: true,
    },
  });
  // Doc Studio draws its own menus (per document type, with search); the native menu stays for its shortcuts.
  w.setMenuBarVisibility(false);
  w.loadFile(path.join(__dirname, 'renderer', 'index.html'), opts.adopt ? { query: { popout: '1' } } : undefined);
  if (!win) win = w;

  // Links clicked inside documents open in the default browser, never in the app.
  w.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:|^mailto:/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  w.webContents.on('will-navigate', (e, url) => {
    if (url !== w.webContents.getURL()) {
      e.preventDefault();
      if (/^https?:|^mailto:/i.test(url)) shell.openExternal(url);
    }
  });

  w.webContents.on('did-finish-load', () => {
    if (opts.adopt) { w.webContents.send('app:adopt', opts.adopt); opts.adopt = null; return; }
    if (w !== win) return;
    pendingOpen.forEach((f) => w.webContents.send('app:open-path', f));
    pendingOpen = [];
  });

  // The renderer decides whether there is unsaved work.
  w.on('close', (e) => {
    if (w.allowClose) return;
    e.preventDefault();
    w.webContents.send('app:close-request');
  });
  w.on('closed', () => {
    if (win === w) win = BrowserWindow.getAllWindows().find((x) => x !== w && !x.isDestroyed() && !x.docStudioWorker) || null;
    aiBridge.onWindowClosed();
  });

  // Right-click in text: spelling suggestions, add to dictionary, then cut/copy/paste.
  w.webContents.on('context-menu', (_e, params) => {
    if (!params.isEditable && !params.selectionText) return;
    const items = [];
    if (params.misspelledWord) {
      for (const s of params.dictionarySuggestions.slice(0, 6)) items.push({ label: s, click: () => w.webContents.replaceMisspelling(s) });
      if (!params.dictionarySuggestions.length) items.push({ label: 'No suggestions', enabled: false });
      items.push({ label: 'Add to dictionary', click: () => w.webContents.session.addWordToSpellCheckerDictionary(params.misspelledWord) }, { type: 'separator' });
    }
    if (params.isEditable) items.push({ role: 'cut', enabled: params.editFlags.canCut }, { role: 'copy', enabled: params.editFlags.canCopy }, { role: 'paste', enabled: params.editFlags.canPaste }, { label: 'Paste without formatting', click: () => w.webContents.pasteAndMatchStyle(), enabled: params.editFlags.canPaste }, { type: 'separator' }, { role: 'selectAll' });
    else items.push({ role: 'copy' });
    Menu.buildFromTemplate(items).popup({ window: w });
  });
  buildMenu();
  return w;
}

app.on('window-all-closed', () => app.quit());

// ---------- menu

function send(cmd, arg) { const w = BrowserWindow.getFocusedWindow() || win; if (w) w.webContents.send('app:command', cmd, arg); }

function buildMenu() {
  const s = readSettings();
  const recent = (s.recent || []).filter((f) => fs.existsSync(f));
  const theme = s.theme || 'system';
  const setTheme = (t) => { const cur = readSettings(); cur.theme = t; writeSettings(cur); nativeTheme.themeSource = t; buildMenu(); };
  const template = [
    {
      label: '&File',
      submenu: [
        {
          label: 'New',
          submenu: [
            { label: 'Markdown (.md)', accelerator: 'CmdOrCtrl+Shift+N', click: () => send('new', 'md') },
            { type: 'separator' },
            { label: 'Document (.docx)', accelerator: 'CmdOrCtrl+N', click: () => send('new', 'docx') },
            { label: 'Spreadsheet (.xlsx)', click: () => send('new', 'xlsx') },
            { label: 'Presentation (.pptx)', click: () => send('new', 'pptx') },
            { label: 'PDF Document (.pdf)', click: () => send('new', 'pdf') },
            { type: 'separator' },
            { label: 'OpenDocument Text (.odt)', click: () => send('new', 'odt') },
            { label: 'OpenDocument Spreadsheet (.ods)', click: () => send('new', 'ods') },
            { label: 'OpenDocument Presentation (.odp)', click: () => send('new', 'odp') },
            { label: 'Rich Text (.rtf)', click: () => send('new', 'rtf') },
            { type: 'separator' },
            { label: 'Plain Text (.txt)', click: () => send('new', 'txt') },
            { label: 'Web Page (.html)', click: () => send('new', 'html') },
            { label: 'CSV Table (.csv)', click: () => send('new', 'csv') },
            { label: 'JSON (.json)', click: () => send('new', 'json') },
          ],
        },
        { type: 'separator' },
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: () => send('open') },
        {
          label: 'Open Recent',
          submenu: recent.length
            ? [...recent.map((f) => ({ label: f, click: () => { const w = BrowserWindow.getFocusedWindow() || win; if (w) w.webContents.send('app:open-path', f); } })),
               { type: 'separator' },
               { label: 'Clear Recent', click: () => { const cur = readSettings(); cur.recent = []; writeSettings(cur); app.clearRecentDocuments(); buildMenu(); } }]
            : [{ label: 'No recent files', enabled: false }],
        },
        { type: 'separator' },
        { label: 'Save', accelerator: 'CmdOrCtrl+S', click: () => send('save') },
        { label: 'Save As…', accelerator: 'CmdOrCtrl+Shift+S', click: () => send('saveAs') },
        { label: 'Export as PDF…', accelerator: 'CmdOrCtrl+Shift+E', click: () => send('exportPdf') },
        { label: 'Print…', accelerator: 'CmdOrCtrl+P', click: () => send('print') },
        { type: 'separator' },
        { label: 'Show in Folder', click: () => send('reveal') },
        { label: 'Open Backups Folder', click: () => { const d = path.join(app.getPath('userData'), 'backups'); fs.mkdirSync(d, { recursive: true }); shell.openPath(d); } },
        { label: 'Close Tab', accelerator: 'CmdOrCtrl+W', click: () => send('closeTab') },
        { type: 'separator' },
        { role: 'quit', label: 'Exit' },
      ],
    },
    {
      label: '&Edit',
      submenu: [
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', click: () => send('undo') },
        { label: 'Redo', accelerator: 'CmdOrCtrl+Y', click: () => send('redo') },
        { type: 'separator' },
        { role: 'cut' }, { role: 'copy' }, { role: 'paste' }, { role: 'pasteAndMatchStyle' }, { role: 'selectAll' },
        { type: 'separator' },
        { label: 'Find…', accelerator: 'CmdOrCtrl+F', click: () => send('find') },
        { label: 'Find and Replace…', accelerator: 'CmdOrCtrl+H', click: () => send('replace') },
        { label: 'Search the Menus', accelerator: 'Alt+/', click: () => send('menuSearch') },
        { label: 'Keyboard Shortcuts', accelerator: 'CmdOrCtrl+/', click: () => send('shortcuts') },
        { label: 'Word Count', accelerator: 'CmdOrCtrl+Shift+C', click: () => send('wordCount') },
        { label: 'Compact Mode', accelerator: 'CmdOrCtrl+Shift+F', click: () => send('compact') },
      ],
    },
    {
      label: '&View',
      submenu: [
        { label: 'Toggle Edit / View', accelerator: 'CmdOrCtrl+E', click: () => send('toggleMode') },
        { label: 'Next Tab', accelerator: 'Ctrl+Tab', click: () => send('nextTab', 1) },
        { label: 'Previous Tab', accelerator: 'Ctrl+Shift+Tab', click: () => send('nextTab', -1) },
        { type: 'separator' },
        { label: 'Zoom In', accelerator: 'CmdOrCtrl+=', click: () => send('zoom', 1) },
        { label: 'Zoom Out', accelerator: 'CmdOrCtrl+-', click: () => send('zoom', -1) },
        { label: 'Actual Size', accelerator: 'CmdOrCtrl+0', click: () => send('zoom', 0) },
        { type: 'separator' },
        {
          label: 'Theme',
          submenu: ['system', 'light', 'dark'].map((t) => ({
            label: t[0].toUpperCase() + t.slice(1), type: 'radio', checked: theme === t, click: () => setTheme(t),
          })),
        },
        { role: 'togglefullscreen' },
        ...(app.isPackaged ? [] : [{ role: 'toggleDevTools' }, { role: 'reload' }]),
      ],
    },
    {
      label: '&Help',
      submenu: [
        { label: 'Supported File Types', click: () => send('help') },
        { label: `About Doc Studio ${app.getVersion()}`, click: () => dialog.showMessageBox(BrowserWindow.getFocusedWindow() || win, {
          type: 'info', title: 'About Doc Studio', message: `Doc Studio ${app.getVersion()}`,
          detail: 'View, edit, create and convert documents, spreadsheets, presentations and PDFs, keeping each file\'s own formatting.\n\nDoc Studio by Griffith Software',
        }) },
      ],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// ---------- IPC

ipcMain.handle('file:openDialog', async (_e) => {
  const r = await dialog.showOpenDialog(W(_e), {
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'All supported', extensions: ['docx', 'docm', 'dotx', 'doc', 'dot', 'xltx', 'xltm', 'md', 'markdown', 'pdf', 'txt', 'rtf', 'odt', 'ott', 'odp', 'ods', 'fods', 'xlsx', 'xls', 'xlsb', 'xlsm', 'numbers', 'csv', 'tsv', 'pptx', 'pptm', 'ppsx', 'ppt', 'pps', 'epub', 'html', 'htm', 'json', 'xml', 'log', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'ico'] },
      { name: 'All files', extensions: ['*'] },
    ],
  });
  return r.canceled ? [] : r.filePaths;
});

ipcMain.handle('file:read', async (_e, file) => {
  const buf = await fs.promises.readFile(file);
  const stat = await fs.promises.stat(file);
  if (!isWorker(_e)) addRecent(file); // the AI worker's files stay out of Recent
  // Transfer as a plain ArrayBuffer slice so the renderer gets exactly the file bytes.
  return { path: file, name: path.basename(file), size: stat.size, data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
});

ipcMain.handle('file:write', async (_e, file, data, opts) => {
  // Write to a temp file first, then rename, so a failed save never truncates the original.
  const tmp = file + '.docstudio-tmp';
  await fs.promises.writeFile(tmp, typeof data === 'string' ? data : Buffer.from(data));
  await fs.promises.rename(tmp, file);
  if (!isWorker(_e)) addRecent(file);
  if (!opts || opts.version !== false) await addVersion(file).catch(() => {});
  return true;
});

// ---------- version history: every save keeps a copy (newest 40 per file) in <userData>/versions/<file id>/

const crypto = require('crypto');
const versionDir = (file) => path.join(app.getPath('userData'), 'versions', crypto.createHash('sha1').update(path.resolve(file).toLowerCase()).digest('hex').slice(0, 16));
async function addVersion(file) {
  const dir = versionDir(file);
  await fs.promises.mkdir(dir, { recursive: true });
  await fs.promises.writeFile(path.join(dir, 'source.txt'), file);
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  await fs.promises.copyFile(file, path.join(dir, `${stamp}${path.extname(file)}`));
  const all = (await fs.promises.readdir(dir)).filter((f) => f !== 'source.txt').sort();
  for (const old of all.slice(0, Math.max(0, all.length - 40))) await fs.promises.unlink(path.join(dir, old)).catch(() => {});
}
ipcMain.handle('versions:list', async (_e, file) => {
  const dir = versionDir(file);
  let names = [];
  try { names = (await fs.promises.readdir(dir)).filter((f) => f !== 'source.txt').sort().reverse(); } catch { return []; }
  const out = [];
  for (const n of names) {
    const st = await fs.promises.stat(path.join(dir, n));
    const iso = n.replace(/\.[^.]+$/, '').replace(/^(\d{4}-\d{2}-\d{2}T\d{2})-(\d{2})-(\d{2})-(\d{3})Z$/, '$1:$2:$3.$4Z');
    out.push({ id: n, time: iso, size: st.size });
  }
  return out;
});
ipcMain.handle('versions:read', async (_e, file, id) => {
  if (!/^[\w.-]+$/.test(id)) return null;
  const buf = await fs.promises.readFile(path.join(versionDir(file), id));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
});

ipcMain.handle('file:stat', async (_e, file) => {
  const st = await fs.promises.stat(file);
  return { size: st.size, created: st.birthtime.toISOString(), modified: st.mtime.toISOString() };
});
ipcMain.handle('file:rename', async (_e, file, newName) => {
  if (!newName || /[\\/:*?"<>|]/.test(newName)) throw new Error('A file name can\'t contain \\ / : * ? " < > |');
  const target = path.join(path.dirname(file), newName);
  if (target.toLowerCase() !== file.toLowerCase() && fs.existsSync(target)) throw new Error(`${newName} already exists in that folder.`);
  await fs.promises.rename(file, target);
  addRecent(target);
  return target;
});
ipcMain.handle('recent:list', () => (readSettings().recent || []).filter((f) => fs.existsSync(f)));
ipcMain.handle('ui:fullscreen', (_e) => { if (W(_e)) W(_e).setFullScreen(!W(_e).isFullScreen()); return !!(W(_e) && W(_e).isFullScreen()); });
ipcMain.handle('ui:theme', (_e, t) => {
  if (t) { const cur = readSettings(); cur.theme = t; writeSettings(cur); nativeTheme.themeSource = t; buildMenu(); }
  return readSettings().theme || 'system';
});

ipcMain.handle('file:saveDialog', async (_e, defaultPath, ext, others) => {
  const names = {
    docx: 'Document', doc: 'Document (97-2003)', rtf: 'Rich Text', odt: 'OpenDocument text', xlsx: 'Spreadsheet',
    xlsm: 'Spreadsheet with macros', xls: 'Spreadsheet (97-2003)', xlsb: 'Binary spreadsheet', ods: 'OpenDocument spreadsheet',
    fods: 'Flat OpenDocument spreadsheet', csv: 'CSV (first sheet)', pptx: 'Presentation', ppt: 'Presentation (97-2003)',
    odp: 'OpenDocument presentation', md: 'Markdown', txt: 'Text', html: 'HTML', pdf: 'PDF',
  };
  const list = [ext, ...(Array.isArray(others) ? others : [])].filter((x, i, a) => x && a.indexOf(x) === i);
  const r = await dialog.showSaveDialog(W(_e), {
    defaultPath,
    filters: [...list.map((x) => ({ name: `${names[x] || x.toUpperCase()} (.${x})`, extensions: [x] })), { name: 'All files', extensions: ['*'] }],
  });
  return r.canceled ? null : r.filePath;
});

ipcMain.handle('fs:exists', (_e, p) => typeof p === 'string' && fs.existsSync(p));
ipcMain.handle('ui:reveal', (_e, file) => { shell.showItemInFolder(file); });
ipcMain.handle('clipboard:read', () => clipboard.readText());

// Windows font files for embedding in PDFs (only plain file names inside the Fonts folder are allowed).
ipcMain.handle('font:read', async (_e, file) => {
  if (!/^[\w .-]+\.(ttf|otf)$/i.test(file)) return null;
  const dirs = [path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts'), path.join(process.env.LOCALAPPDATA || '', 'Microsoft', 'Windows', 'Fonts')];
  for (const d of dirs) {
    const p = path.join(d, file);
    try { const buf = await fs.promises.readFile(p); return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); } catch { /* try next */ }
  }
  return null;
});
// Windows' Default apps page for Doc Studio (the installer registers it there; only the owner can set the default).
ipcMain.handle('ui:defaultApps', () => shell.openExternal('ms-settings:defaultapps?registeredAppUser=Doc%20Studio'));
ipcMain.handle('ui:openExternal', (_e, url) => { if (/^https?:|^mailto:/i.test(url)) shell.openExternal(url); });

ipcMain.handle('ui:find', (_e, text, opts) => {
  if (!text) { W(_e).webContents.stopFindInPage('clearSelection'); return; }
  W(_e).webContents.findInPage(text, opts || {});
});
ipcMain.handle('ui:stopFind', (_e) => W(_e).webContents.stopFindInPage('keepSelection'));

ipcMain.handle('ui:exportPdf', async (_e, defaultPath) => {
  const target = await dialog.showSaveDialog(W(_e), { defaultPath, filters: [{ name: 'PDF', extensions: ['pdf'] }] });
  if (target.canceled) return null;
  const data = await W(_e).webContents.printToPDF({ printBackground: true, margins: { marginType: 'default' } });
  await fs.promises.writeFile(target.filePath, data);
  return target.filePath;
});

// Printing asks Windows for its printers first (off the main thread, with a time limit): a printer driver that
// hangs would otherwise freeze the whole window inside the print dialog.
ipcMain.handle('ui:print', async (_e) => {
  const printers = await Promise.race([W(_e).webContents.getPrintersAsync().catch(() => null), new Promise((r) => setTimeout(() => r('timeout'), 6000))]);
  if (printers === 'timeout' || printers === null) return { ok: false, reason: 'no-response' };
  if (!printers.length) return { ok: false, reason: 'no-printers' };
  return new Promise((resolve) => W(_e).webContents.print({ printBackground: true }, (ok, reason) => resolve({ ok, reason })));
});

// ---------- recovery: unsaved work survives a crash or a power cut
// While a document has unsaved edits, the window keeps a copy of it in <userData>/recovery (renderer/recovery.js).
// Saving, closing or discarding it removes the copy, so copies found when Doc Studio starts are work that was lost.
const recoveryDir = () => path.join(app.getPath('userData'), 'recovery');
const recoveryKey = (k) => String(k).replace(/[^\w-]/g, '');
let recoveryOffered = false;
ipcMain.handle('recovery:save', async (_e, key, meta, bytes) => {
  const dir = recoveryDir();
  await fs.promises.mkdir(dir, { recursive: true });
  const base = path.join(dir, recoveryKey(key));
  // Write to a temporary name, then rename, so a power cut mid-write never leaves a half-written copy.
  await fs.promises.writeFile(`${base}.bin.tmp`, Buffer.from(bytes));
  await fs.promises.rename(`${base}.bin.tmp`, `${base}.bin`);
  await fs.promises.writeFile(`${base}.json.tmp`, JSON.stringify({ ...meta, time: Date.now() }));
  await fs.promises.rename(`${base}.json.tmp`, `${base}.json`);
});
ipcMain.handle('recovery:clear', async (_e, keys) => {
  for (const k of [].concat(keys)) for (const ext of ['bin', 'json']) await fs.promises.rm(path.join(recoveryDir(), `${recoveryKey(k)}.${ext}`), { force: true });
});
// Copies left from an earlier run, offered once, to the first window that asks.
ipcMain.handle('recovery:take', async () => {
  if (recoveryOffered) return [];
  recoveryOffered = true;
  let names = [];
  try { names = await fs.promises.readdir(recoveryDir()); } catch { return []; }
  const found = [];
  for (const n of names.filter((x) => x.endsWith('.json'))) {
    const key = n.slice(0, -5);
    try {
      const meta = JSON.parse(await fs.promises.readFile(path.join(recoveryDir(), n), 'utf8'));
      const buf = await fs.promises.readFile(path.join(recoveryDir(), `${key}.bin`));
      found.push({ key, meta, data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) });
    } catch { /* incomplete copy */ }
  }
  return found.sort((a, b) => b.meta.time - a.meta.time);
});

ipcMain.handle('app:closeConfirmed',(_e) => { const w = W(_e); if (w) { w.allowClose = true; w.close(); } });

// A blank Word document for File > New (Word's usual styles: Normal, Title, Heading 1-6, List Paragraph...).
ipcMain.handle('docx:blank', async () => {
  const { Document, Packer, Paragraph } = require('docx');
  const heading = (size, color, before, extra) => ({
    run: { font: 'Calibri Light', size, color, ...(extra || {}) },
    paragraph: { spacing: { before, after: 0 }, keepNext: true, keepLines: true },
  });
  const doc = new Document({
    creator: 'Doc Studio',
    styles: {
      default: {
        document: { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 160, line: 259 } } },
        title: { run: { font: 'Calibri Light', size: 56, kern: 28 }, paragraph: { spacing: { after: 0, line: 240 } } },
        heading1: heading(32, '2F5496', 240),
        heading2: heading(26, '2F5496', 40),
        heading3: heading(24, '1F3763', 40),
        heading4: heading(22, '2F5496', 40, { italics: true }),
        listParagraph: { paragraph: { indent: { left: 720 }, contextualSpacing: true } },
        hyperlink: { run: { color: '0563C1', underline: { type: 'single' } } },
      },
      paragraphStyles: [
        { id: 'Subtitle', name: 'Subtitle', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { color: '5A5A5A', size: 22 }, paragraph: { spacing: { after: 160 } } },
        { id: 'Quote', name: 'Quote', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { italics: true, color: '404040' }, paragraph: { spacing: { before: 200, after: 160 }, indent: { left: 864, right: 864 }, alignment: 'center' } },
        { id: 'NoSpacing', name: 'No Spacing', quickFormat: true, paragraph: { spacing: { after: 0, line: 240 } } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [new Paragraph({})],
    }],
  });
  const buf = await Packer.toBuffer(doc);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
});

// Before Doc Studio first overwrites a file, a copy of the original goes to <userData>/backups (newest 50 kept).
ipcMain.handle('file:backup', async (_e, file) => {
  if (!fs.existsSync(file)) return false;
  const dir = path.join(app.getPath('userData'), 'backups');
  await fs.promises.mkdir(dir, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  await fs.promises.copyFile(file, path.join(dir, `${stamp} ${path.basename(file)}`));
  const all = (await fs.promises.readdir(dir)).sort();
  for (const old of all.slice(0, Math.max(0, all.length - 50))) await fs.promises.unlink(path.join(dir, old)).catch(() => {});
  return true;
});

// Editor HTML -> DOCX bytes.
ipcMain.handle('docx:fromHtml', async (_e, html, title) => {
  const { Document, Packer } = require('docx');
  const { htmlToContent } = require('docx/html');
  const doc = new Document({
    title: title || '',
    creator: 'Doc Studio',
    styles: { default: { document: { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 160, line: 264 } } } } },
    sections: [{
      // US Letter with 1" margins.
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: htmlToContent(html, { maxImageWidth: 624 }),
    }],
  });
  const buf = await Packer.toBuffer(doc);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
});

// Legacy binary .doc: text only.
ipcMain.handle('doc:extract', async (_e, data) => {
  const WordExtractor = require('word-extractor');
  const doc = await new WordExtractor().extract(Buffer.from(data));
  return { body: doc.getBody(), headers: doc.getHeaders({ includeFooters: false }), footers: doc.getFooters() };
});

// Conversions through installed office programs (Word/Excel/PowerPoint over COM, or LibreOffice).
const converter = require('./convert');
ipcMain.handle('convert:tools', () => converter.tools());
ipcMain.handle('convert:route', (_e, from, to) => converter.route(from, to));
ipcMain.handle('convert:run', (_e, data, from, to) => converter.convert(data, from, to));

// ---------- rendering HTML off screen: PDF files and PNG pictures for conversions
// The page is written to a private temp folder and loaded in a hidden, script-free window with Doc Studio's styles.

const os = require('os');
const { pathToFileURL } = require('url');
// A document's links to pictures on the internet stay unloaded: the page may use local files only, so
// converting never contacts other sites.
const OFFSCREEN_CSP = "default-src 'none'; img-src data: blob: file:; style-src 'unsafe-inline' file:; font-src data: file:; media-src data: blob: file:";
async function offscreenPage(html, opts) {
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'docstudio-render-'));
  const file = path.join(dir, 'page.html');
  const css = pathToFileURL(path.join(__dirname, 'renderer', 'styles.css')).href;
  await fs.promises.writeFile(file, `<!DOCTYPE html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${OFFSCREEN_CSP}"><link rel="stylesheet" href="${css}"><style>html,body{background:#fff;color:#000;color-scheme:light;height:auto;overflow:visible;display:block}</style></head><body class="convert-page">${html}</body></html>`);
  const w = new BrowserWindow({
    show: false, width: Math.round((opts && opts.width) || 1000), height: Math.round((opts && opts.height) || 800), useContentSize: true,
    webPreferences: { javascript: false, sandbox: true, contextIsolation: true, offscreen: !!(opts && opts.offscreen) },
  });
  w.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  w.webContents.on('will-navigate', (e) => e.preventDefault());
  await w.loadFile(file);
  // Give pictures and fonts a moment to finish.
  await new Promise((r) => setTimeout(r, 300));
  return { w, done: () => { w.destroy(); fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {}); } };
}
ipcMain.handle('pdf:fromHtml', async (_e, html, opts) => {
  const { w, done } = await offscreenPage(html, {});
  try {
    const data = await w.webContents.printToPDF({
      printBackground: true, preferCSSPageSize: true,
      pageSize: (opts && opts.pageSize) || 'Letter', landscape: !!(opts && opts.landscape),
      margins: opts && opts.margins ? opts.margins : { marginType: 'default' },
    });
    return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
  } finally { done(); }
});
ipcMain.handle('png:fromHtml', async (_e, pages, width, height) => {
  // Each entry of `pages` is drawn alone at width x height and captured.
  const out = [];
  for (const html of pages) {
    const { w, done } = await offscreenPage(`<div style="width:${width}px;height:${height}px;overflow:hidden">${html}</div>`, { width, height, offscreen: true });
    try {
      await new Promise((r) => setTimeout(r, 200));
      const img = await w.webContents.capturePage({ x: 0, y: 0, width: Math.round(width), height: Math.round(height) });
      const buf = img.toPNG();
      out.push(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
    } finally { done(); }
  }
  return out;
});

ipcMain.handle('edit:op', (_e, op) => { if (W(_e) && ['cut', 'copy', 'paste', 'pasteAndMatchStyle', 'selectAll', 'undo', 'redo'].includes(op)) W(_e).webContents[op](); });
ipcMain.handle('spell:set', (_e, on) => { if (W(_e)) W(_e).webContents.session.setSpellCheckerEnabled(!!on); });
ipcMain.handle('app:version', () => app.getVersion());
ipcMain.handle('ui:openBackups', () => { const d = path.join(app.getPath('userData'), 'backups'); fs.mkdirSync(d, { recursive: true }); shell.openPath(d); });

// Dictionary: definitions of a word the user asked to look up (Tools > Dictionary), from dictionaryapi.dev.
ipcMain.handle('dict:lookup', async (_e, word) => {
  if (!/^[\p{L}'-]{1,60}$/u.test(word || '')) return null;
  const { net } = require('electron');
  try {
    const r = await net.fetch(`https://api.dictionaryapi.dev/api/v2/entries/en/${encodeURIComponent(word)}`);
    if (!r.ok) return null;
    return await r.json();
  } catch { return null; }
});

// ---------- app shell: title bar, settings carried over, places, file lists, starred files, OCR

function overlayColors() {
  const dark = nativeTheme.shouldUseDarkColors;
  return { color: dark ? '#111318' : '#1b1e25', symbolColor: '#e8eaed', height: 40 };
}

// Settings, version history and backups from the program's earlier name are copied over once.
function migrateSettings() {
  try {
    const now = app.getPath('userData');
    const before = path.join(path.dirname(now), 'Any Doc'); // the program's earlier name
    if (fs.existsSync(path.join(now, 'settings.json')) || !fs.existsSync(before)) return;
    fs.mkdirSync(now, { recursive: true });
    for (const name of ['settings.json', 'versions', 'backups']) {
      const from = path.join(before, name);
      if (fs.existsSync(from)) fs.cpSync(from, path.join(now, name), { recursive: true, errorOnExist: false });
    }
  } catch { /* nothing to carry over */ }
}

const DOC_EXTS = new Set(['docx', 'docm', 'dotx', 'doc', 'dot', 'md', 'markdown', 'pdf', 'txt', 'rtf', 'odt', 'ott', 'odp', 'ods', 'fods', 'xlsx', 'xls', 'xlsb', 'xlsm', 'csv', 'tsv', 'pptx', 'pptm', 'ppsx', 'ppt', 'pps', 'epub', 'html', 'htm', 'json', 'xml', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg']);
// Search lists documents (not pictures, code or data files).
const SEARCH_EXTS = new Set(['docx', 'docm', 'dotx', 'doc', 'md', 'markdown', 'pdf', 'txt', 'rtf', 'odt', 'ods', 'odp', 'xlsx', 'xls', 'xlsb', 'xlsm', 'csv', 'pptx', 'pptm', 'ppsx', 'ppt', 'epub']);
const extOf = (f) => (f.includes('.') ? f.slice(f.lastIndexOf('.') + 1).toLowerCase() : '');

// Cloud folders synced to this PC (OneDrive, Google Drive, Dropbox, iCloud, Box).
function cloudFolders() {
  const home = os.homedir();
  const out = [];
  const add = (name, kind, p) => { try { if (p && fs.existsSync(p) && fs.statSync(p).isDirectory() && !out.some((x) => x.path.toLowerCase() === p.toLowerCase())) out.push({ name, kind, path: p }); } catch { /* skip */ } };
  for (const v of ['OneDrive', 'OneDriveConsumer', 'OneDriveCommercial']) add(v === 'OneDriveCommercial' ? 'OneDrive (work or school)' : 'OneDrive', 'onedrive', process.env[v]);
  try { for (const d of fs.readdirSync(home)) if (/^OneDrive/i.test(d)) add(d.replace(/^OneDrive\s*-\s*/i, 'OneDrive - '), 'onedrive', path.join(home, d)); } catch { /* skip */ }
  add('Google Drive', 'gdrive', path.join(home, 'Google Drive'));
  add('Google Drive', 'gdrive', path.join(home, 'My Drive'));
  for (const l of 'DEFGHIJKLMNOPQRSTUVWXYZ') add('Google Drive', 'gdrive', `${l}:\\My Drive`);
  add('Dropbox', 'dropbox', path.join(home, 'Dropbox'));
  try { const info = JSON.parse(fs.readFileSync(path.join(process.env.LOCALAPPDATA || '', 'Dropbox', 'info.json'), 'utf8')); for (const v of Object.values(info)) add('Dropbox', 'dropbox', v.path); } catch { /* no Dropbox */ }
  add('iCloud Drive', 'icloud', path.join(home, 'iCloudDrive'));
  add('Box', 'box', path.join(home, 'Box'));
  return out;
}
ipcMain.handle('fs:places', () => ({
  home: os.homedir(),
  documents: app.getPath('documents'),
  desktop: app.getPath('desktop'),
  downloads: app.getPath('downloads'),
  cloud: cloudFolders(),
}));
ipcMain.handle('fs:list', async (_e, dir) => {
  const out = [];
  let entries = [];
  try { entries = await fs.promises.readdir(dir, { withFileTypes: true }); } catch (err) { return { error: err.message, items: [] }; }
  for (const d of entries) {
    if (d.name.startsWith('.') || d.name.startsWith('~$') || /^(desktop\.ini|thumbs\.db)$/i.test(d.name)) continue;
    const p = path.join(dir, d.name);
    try {
      const st = await fs.promises.stat(p);
      if (d.isDirectory()) out.push({ name: d.name, path: p, dir: true, mtime: st.mtimeMs });
      else if (DOC_EXTS.has(extOf(d.name))) out.push({ name: d.name, path: p, dir: false, size: st.size, mtime: st.mtimeMs, ext: extOf(d.name) });
    } catch { /* unreadable */ }
  }
  return { items: out };
});
// A quick index of documents in the usual places, for "Search documents".
let docIndex = null;
ipcMain.handle('docs:index', async () => {
  if (docIndex && Date.now() - docIndex.at < 60000) return docIndex.files;
  const roots = [app.getPath('documents'), app.getPath('desktop'), app.getPath('downloads'), ...cloudFolders().map((c) => c.path)];
  const files = [];
  const seen = new Set();
  const walk = async (dir, depth) => {
    if (files.length >= 6000 || depth > 5 || seen.has(dir.toLowerCase())) return;
    seen.add(dir.toLowerCase());
    let entries = [];
    try { entries = await fs.promises.readdir(dir, { withFileTypes: true }); } catch { return; }
    for (const d of entries) {
      // Program and build folders are skipped: search is for documents, not source code.
      if (d.name.startsWith('.') || d.name.startsWith('~$') || /^(node_modules|AppData|\$RECYCLE\.BIN|dist|build|out|bin|obj|vendor|venv|\.venv|__pycache__|site-packages|target|packages|build-tools|sdk|[\w.-]*-sdk[\w.-]*|jdk[\w.-]*|node-v\d[\w.-]*)$/i.test(d.name)) continue;
      const p = path.join(dir, d.name);
      if (d.isDirectory()) { if (!fs.existsSync(path.join(p, 'package.json')) && !fs.existsSync(path.join(p, '.git'))) await walk(p, depth + 1); }
      else if (SEARCH_EXTS.has(extOf(d.name))) files.push(p);
      if (files.length >= 6000) return;
    }
  };
  for (const r of roots) await walk(r, 0);
  docIndex = { at: Date.now(), files };
  return files;
});
ipcMain.handle('starred:list', () => (readSettings().starred || []).filter((f) => fs.existsSync(f)));
ipcMain.handle('starred:toggle', (_e, file) => {
  const s = readSettings();
  const list = s.starred || [];
  const on = !list.some((f) => f.toLowerCase() === file.toLowerCase());
  s.starred = on ? [file, ...list] : list.filter((f) => f.toLowerCase() !== file.toLowerCase());
  writeSettings(s);
  return on;
});
ipcMain.handle('file:pick', async (_e, opts) => {
  opts = opts || {};
  const r = await dialog.showOpenDialog(W(_e), {
    title: opts.title,
    properties: [opts.folder ? 'openDirectory' : 'openFile', ...(opts.multi ? ['multiSelections'] : []), ...(opts.folder ? ['createDirectory'] : [])],
    filters: opts.folder ? undefined : [...(opts.filters || []), { name: 'All files', extensions: ['*'] }],
  });
  return r.canceled ? [] : r.filePaths;
});

// OCR with the text recognition built into Windows (Windows.Media.Ocr; works offline with the installed languages).
const OCR_PS = String.raw`param([string]$list)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics, ContentType=WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation` + '`' + String.raw`1' })[0]
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(-1) | Out-Null; $task.Result }
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if (-not $engine) { Write-Error 'No OCR language is installed in Windows.'; exit 2 }
foreach ($img in (Get-Content -LiteralPath $list -Encoding UTF8)) {
  if (-not $img) { continue }
  $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($img)) ([Windows.Storage.StorageFile])
  $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
  $dec = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
  $bmp = Await ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
  $res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
  foreach ($line in $res.Lines) { [Console]::Out.WriteLine($line.Text) }
  [Console]::Out.WriteLine([char]12)
  $stream.Dispose()
}
`;
ipcMain.handle('ocr:images', async (_e, images) => {
  const { execFile } = require('child_process');
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'docstudio-ocr-'));
  try {
    const files = [];
    for (const [i, img] of images.entries()) { const f = path.join(dir, `page${i + 1}.${img.ext || 'png'}`); await fs.promises.writeFile(f, Buffer.from(img.data)); files.push(f); }
    const list = path.join(dir, 'list.txt');
    await fs.promises.writeFile(list, files.join('\r\n'), 'utf8');
    const script = path.join(dir, 'ocr.ps1');
    await fs.promises.writeFile(script, '\ufeff' + OCR_PS, 'utf8');
    const out = await new Promise((resolve, reject) => execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, list], { maxBuffer: 64 * 1024 * 1024, windowsHide: true, encoding: 'utf8', timeout: 10 * 60 * 1000 }, (err, stdout, stderr) => (err ? reject(new Error((stderr || err.message).split('\n')[0])) : resolve(stdout))));
    return out.split('\f').map((t) => t.replace(/\r/g, '').trim()).slice(0, images.length);
  } finally { fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {}); }
});

// ---------- default apps
// Windows lets only the owner choose a default app (it protects that choice from programs). So this asks Windows
// which app opens each type, and shows each type's sample file in File Explorer, where the owner picks Open with > Doc Studio >
// Always (Windows only offers "Always" there). Opening the sample tells Doc Studio that type is done.
// The types match build/installer.nsh, which registers Doc Studio for them.

const DEFAULT_TYPES = [
  ['Documents', [['docx', 'Document'], ['doc', 'Document (97-2003)'], ['docm', 'Document with macros'], ['dotx', 'Document template'], ['dot', 'Document template (97-2003)'], ['rtf', 'Rich Text document'], ['odt', 'OpenDocument text'], ['ott', 'OpenDocument text template']]],
  ['Spreadsheets', [['xlsx', 'Spreadsheet'], ['xls', 'Spreadsheet (97-2003)'], ['xlsm', 'Spreadsheet with macros'], ['xlsb', 'Binary spreadsheet'], ['xltx', 'Spreadsheet template'], ['xltm', 'Spreadsheet template with macros'], ['ods', 'OpenDocument spreadsheet'], ['fods', 'OpenDocument flat spreadsheet'], ['numbers', 'Numbers spreadsheet'], ['csv', 'CSV table'], ['tsv', 'Tab-separated table']]],
  ['Presentations', [['pptx', 'Presentation'], ['ppt', 'Presentation (97-2003)'], ['pptm', 'Presentation with macros'], ['ppsx', 'Slide show'], ['pps', 'Slide show (97-2003)'], ['odp', 'OpenDocument presentation']]],
  ['PDF and books', [['pdf', 'PDF document'], ['epub', 'EPUB book']]],
  ['Markdown and text', [['md', 'Markdown document'], ['markdown', 'Markdown document'], ['txt', 'Text document'], ['log', 'Log file'], ['html', 'Web page'], ['htm', 'Web page'], ['json', 'JSON file'], ['xml', 'XML file']]],
];
const DEFAULT_EXTS = new Set(DEFAULT_TYPES.flatMap(([, list]) => list.map(([x]) => x)));
const DEFAULTS_PS = String.raw`param([string]$Mode, [string]$Arg)
$ErrorActionPreference = 'Stop'
Add-Type -TypeDefinition @"
using System; using System.Text; using System.Runtime.InteropServices;
public static class DSAssoc {
  [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)] static extern int AssocQueryString(int flags, int str, string assoc, string extra, StringBuilder outp, ref uint len);
  public static string Query(string assoc, int what) { var sb = new StringBuilder(1024); uint n = 1024; return AssocQueryString(0x400, what, assoc, null, sb, ref n) == 0 ? sb.ToString() : ""; }
  [ComImport, Guid("4e530b0a-e611-4c77-a3ac-9031d022281b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  interface IAAR { [PreserveSig] int QueryCurrentDefault([MarshalAs(UnmanagedType.LPWStr)] string q, int type, int level, [MarshalAs(UnmanagedType.LPWStr)] out string progId); }
  [ComImport, Guid("591209c7-767b-42b2-9fba-44ee4615f2c7")] class AAR { }
  public static string Current(string ext) { string p; return ((IAAR)new AAR()).QueryCurrentDefault(ext, 0, 1, out p) == 0 ? p : ""; }
}
"@
function ProgOf($x) {
  $p = [DSAssoc]::Current(".$x"); if ($p) { return $p }
  $b = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.$x"
  $p = (Get-ItemProperty "$b\UserChoiceLatest\ProgId" -ErrorAction SilentlyContinue).ProgId; if ($p) { return $p }
  $p = (Get-ItemProperty "$b\UserChoice" -ErrorAction SilentlyContinue).ProgId; if ($p) { return $p }
  return (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\.$x" -ErrorAction SilentlyContinue).'(default)'
}
function NameOf($p, $x) {
  if (-not $p) { return [DSAssoc]::Query(".$x", 4) }
  $n = [DSAssoc]::Query($p, 4); if ($n -and $n -ne '%1') { return $n }
  if ($p -match '^Applications\\(.+)$') { $f = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$p" -ErrorAction SilentlyContinue).FriendlyAppName; if ($f) { return $f }; return $Matches[1] }
  $c = (Get-ItemProperty "Registry::HKEY_CLASSES_ROOT\$p\shell\open\command" -ErrorAction SilentlyContinue).'(default)'
  if ($c -match '([^\\"]+)\.exe') { $e = $Matches[1]; $o = @{ winword = 'Word'; excel = 'Excel'; powerpnt = 'PowerPoint' }[$e.ToLower()]; if ($o) { return $o }; return $e }
  return [DSAssoc]::Query(".$x", 4)
}
foreach ($x in ($Arg -split ',')) { $p = ProgOf $x; [Console]::Out.WriteLine("$x" + [char]9 + $p + [char]9 + (NameOf $p $x) + [char]9 + $(if ($p) { [DSAssoc]::Query($p, 2) } else { '' })) }
`;
async function runDefaults(args) {
  const { execFile } = require('child_process');
  const dir = path.join(os.tmpdir(), 'docstudio-defaults');
  await fs.promises.mkdir(dir, { recursive: true });
  const script = path.join(dir, 'defaults.ps1');
  await fs.promises.writeFile(script, '﻿' + DEFAULTS_PS, 'utf8');
  const out = await new Promise((resolve, reject) => execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, ...args], { windowsHide: true, encoding: 'utf8', timeout: 30 * 60 * 1000 }, (err, stdout, stderr) => (err ? reject(new Error((stderr || err.message).split('\n')[0])) : resolve(stdout))));
  const me = path.resolve(process.execPath).toLowerCase();
  const status = {};
  for (const line of out.split(/\r?\n/)) {
    const parts = line.split('\t');
    if (parts.length < 3) continue;
    const [ext, progId, name, exe] = parts;
    // Doc Studio's own registrations (installer ProgIDs, or the app picked from the Open with list).
    const mine = /^(DocStudio\.|Applications\\Doc Studio\.exe$)/i.test(progId || '') || (!!exe && path.resolve(exe).toLowerCase() === me);
    // legacy: chosen as "the Doc Studio app" (before per-type icons), so Explorer shows the plain app icon until
    // the type is chosen again.
    const legacy = mine && !/^DocStudio\./i.test(progId || '');
    status[ext] = { mine, legacy, app: mine ? 'Doc Studio' : (name || (exe ? path.basename(exe) : '')) };
  }
  return { status };
}
ipcMain.handle('defaults:status', async () => ({ types: DEFAULT_TYPES, packaged: app.isPackaged, ...(await runDefaults(['status', [...DEFAULT_EXTS].join(',')])) }));
ipcMain.handle('defaults:reveal', async (_e, ext) => {
  if (!DEFAULT_EXTS.has(ext)) throw new Error('Unknown file type.');
  const dir = path.join(os.tmpdir(), 'docstudio-defaults');
  await fs.promises.mkdir(dir, { recursive: true });
  const sample = path.join(dir, `Choose the app for .${ext} files.${ext}`);
  if (!fs.existsSync(sample)) await fs.promises.writeFile(sample, '');
  shell.showItemInFolder(sample);
  return sample;
});
ipcMain.handle('defaults:check', async (_e, ext) => (await runDefaults(['status', DEFAULT_EXTS.has(ext) ? ext : 'txt'])).status[ext] || {});

// ---------- scanning (Windows Image Acquisition, the scanner support built into Windows)

const SCAN_PS = String.raw`param([string]$Mode, [string]$Id, [string]$Out, [int]$Dpi, [int]$Intent, [string]$Paper)
$ErrorActionPreference = 'Stop'
$dm = New-Object -ComObject WIA.DeviceManager
if ($Mode -eq 'list') {
  foreach ($d in $dm.DeviceInfos) { if ($d.Type -eq 1) { [Console]::Out.WriteLine($d.DeviceID + [char]9 + $d.Properties.Item('Name').Value) } }
  exit 0
}
$info = $null
foreach ($d in $dm.DeviceInfos) { if ($d.DeviceID -eq $Id) { $info = $d } }
if (-not $info) { throw 'SCANNER_GONE' }
$dev = $info.Connect()
$item = $dev.Items.Item(1)
# Sets a scanner setting, kept within the range the scanner allows; settings it doesn't have are skipped.
function SetProp($obj, $propId, $val) {
  foreach ($p in $obj.Properties) {
    if ($p.PropertyID -eq $propId) {
      try { if ($p.SubType -eq 1) { $val = [Math]::Min([Math]::Max($val, $p.SubTypeMin), $p.SubTypeMax) }; $p.Value = $val } catch { }
    }
  }
}
SetProp $item 6146 $Intent
SetProp $item 6147 $Dpi
SetProp $item 6148 $Dpi
if ($Paper -ne 'glass') {
  $w = 8.5; $h = 11
  if ($Paper -eq 'a4') { $w = 8.27; $h = 11.69 }
  SetProp $item 6149 0
  SetProp $item 6150 0
  SetProp $item 6151 ([int]($w * $Dpi))
  SetProp $item 6152 ([int]($h * $Dpi))
}
$png = '{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}'
try { $img = $item.Transfer($png) } catch { $img = $item.Transfer() }
# Some drivers return BMP whatever is asked for; save it under the right extension (Doc Studio reads both).
if ($img.FormatID -ne $png) { $Out = [IO.Path]::ChangeExtension($Out, '.bmp') }
if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out }
$img.SaveFile($Out)
[Console]::Out.WriteLine($Out)
`;
// Scanner errors in plain words (WIA error codes).
const SCAN_ERRORS = [
  [/SCANNER_GONE|0x80210015|0x8021000a/i, 'The scanner is not connected or not turned on.'],
  [/0x80210006/i, 'The scanner is busy. Wait for it to finish, then try again.'],
  [/0x80210005|0x80210004/i, "The scanner isn't responding. Check that it's on and connected, then try again."],
  [/0x80210003/i, 'There is no paper in the document feeder.'],
  [/0x80210002/i, 'The paper is jammed in the scanner.'],
  [/0x80210016/i, "The scanner's lid is open."],
  [/0x80210064/i, 'The scan was cancelled.'],
  [/WIA\.DeviceManager|80040154/i, 'Windows scanner support (WIA) is not available on this PC.'],
];
async function runScan(args, timeout) {
  const { execFile } = require('child_process');
  const dir = path.join(os.tmpdir(), 'docstudio-scan');
  await fs.promises.mkdir(dir, { recursive: true });
  const script = path.join(dir, 'scan.ps1');
  await fs.promises.writeFile(script, '﻿' + SCAN_PS, 'utf8');
  return new Promise((resolve, reject) => execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, ...args], { windowsHide: true, encoding: 'utf8', timeout }, (err, stdout, stderr) => {
    if (!err) { resolve(stdout); return; }
    const text = `${stderr || ''} ${err.message}`;
    const known = SCAN_ERRORS.find(([re]) => re.test(text));
    reject(new Error(known ? known[1] : `The scanner reported a problem: ${(stderr || err.message).split('\n')[0].trim()}`));
  }));
}
ipcMain.handle('scan:devices', async () => (await runScan(['list'], 30000)).split(/\r?\n/).filter((l) => l.includes('\t')).map((l) => { const [id, name] = l.split('\t'); return { id, name: name.trim() }; }));
ipcMain.handle('scan:page', async (_e, o) => {
  const out = path.join(os.tmpdir(), 'docstudio-scan', `page-${Date.now()}.png`);
  const intent = { color: 1, gray: 2, bw: 4 }[o.color] || 1;
  const dpi = [75, 100, 150, 200, 300, 600].includes(+o.dpi) ? +o.dpi : 300;
  const saved = (await runScan(['page', String(o.id), out, String(dpi), String(intent), ['letter', 'a4', 'glass'].includes(o.paper) ? o.paper : 'letter'], 5 * 60 * 1000)).trim().split(/\r?\n/).pop();
  const buf = await fs.promises.readFile(saved);
  fs.promises.rm(saved, { force: true }).catch(() => {});
  return { ext: path.extname(saved).slice(1).toLowerCase(), data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) };
});

// Files that ship inside the program (the background remover's model and runtime): read by name, inside
// renderer/vendor only.
ipcMain.handle('app:asset', async (_e, rel) => {
  if (typeof rel !== 'string' || !/^vendor\/[\w./-]+$/.test(rel) || rel.includes('..')) throw new Error('Not a program file.');
  const buf = await fs.promises.readFile(path.join(__dirname, 'renderer', rel));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
});

// ---------- phone link (phone.js): send files between a phone and this PC over the home network

// A free name for a new file: 'name (2).ext', 'name (3).ext'... when the name is taken.
async function freeName(target) {
  if (!fs.existsSync(target)) return target;
  const ext = path.extname(target), base = target.slice(0, target.length - ext.length);
  for (let i = 2; ; i++) { const t = `${base} (${i})${ext}`; if (!fs.existsSync(t)) return t; }
}
let phonePageHtml = null;
const phone = require('./phone');
phone.register(ipcMain, { W, freeName, phonePage: () => (phonePageHtml = phonePageHtml || fs.readFileSync(path.join(__dirname, 'phone-page.html'), 'utf8')) });
app.on('will-quit', () => {
  phone.stop();
  // Doc Studio's own temporary folders (Default apps samples, the scanner script, the phone link's folder).
  for (const d of ['docstudio-defaults', 'docstudio-scan', 'docstudio-phone']) { try { fs.rmSync(path.join(os.tmpdir(), d), { recursive: true, force: true }); } catch { /* in use */ } }
});

// ---------- AI assistants: Settings > AI assistants (the packages are built into resourcesai by
// tools/make-ai-packages.js; ai/bridge.js answers the MCP server)

// AI assistants are off until the owner turns them on.
const aiOn = () => readSettings().aiEnabled === true;
const aiDir = () => (app.isPackaged ? path.join(process.resourcesPath, 'ai') : path.join(__dirname, 'build', 'ai'));
// Whether an AI app looks installed (its settings folder or program is there).
function aiApps() {
  const home = os.homedir(), local = process.env.LOCALAPPDATA || '', roaming = process.env.APPDATA || '';
  const onPath = (name) => (process.env.PATH || '').split(path.delimiter).some((d) => d && ['.cmd', '.exe', '.ps1', ''].some((x) => fs.existsSync(path.join(d, name + x))));
  return {
    claudeCode: fs.existsSync(path.join(home, '.claude')) || onPath('claude'),
    claudeDesktop: fs.existsSync(path.join(roaming, 'Claude')) || fs.existsSync(path.join(local, 'AnthropicClaude')),
    gemini: fs.existsSync(path.join(home, '.gemini')) || onPath('gemini'),
    codex: fs.existsSync(path.join(home, '.codex')) || onPath('codex'),
  };
}
ipcMain.handle('ai:info', () => {
  const dir = aiDir();
  const registry = aiBridge.registry();
  return { enabled: aiOn(), dir, exe: process.execPath, server: path.join(dir, 'mcp-server.js'), mcpb: path.join(dir, 'doc-studio.mcpb'), gemini: path.join(dir, 'gemini'), ready: fs.existsSync(path.join(dir, 'mcp-server.js')), apps: aiApps(), madeFiles: Object.values(registry).filter((f) => fs.existsSync(f.path)).length };
});
ipcMain.handle('ai:setEnabled', async (_e, on) => { const s = readSettings(); s.aiEnabled = !!on; writeSettings(s); await aiBridge.setEnabled(!!on); return !!on; });
// Claude Desktop doesn't register itself for .mcpb files on every PC (opening one can bring up Windows' app
// chooser), so the file is shown in File Explorer and the window explains Claude's own Install extension steps.
ipcMain.handle('ai:installClaudeDesktop', async () => { const f = path.join(aiDir(), 'doc-studio.mcpb'); if (!fs.existsSync(f)) throw new Error('The Claude Desktop package is missing. Reinstall Doc Studio.'); shell.showItemInFolder(f); return f; });

// ---------- pop out: a document moves into its own window

ipcMain.handle('win:popOut', (e, doc, how) => {
  const from = W(e);
  const { screen } = require('electron');
  const area = screen.getDisplayMatching(from.getBounds()).workArea;
  let bounds;
  if (how === 'side') {
    const half = Math.floor(area.width / 2);
    if (from.isMaximized()) from.unmaximize();
    if (from.isFullScreen()) from.setFullScreen(false);
    from.setBounds({ x: area.x, y: area.y, width: half, height: area.height });
    bounds = { x: area.x + half, y: area.y, width: area.width - half, height: area.height };
  } else if (how && typeof how === 'object' && how.x != null) {
    // Dropped outside the window: the new window appears where the tab was let go.
    const b = from.getBounds();
    const width = Math.min(Math.max(900, Math.round(b.width * 0.75)), area.width), height = Math.min(b.height, area.height);
    bounds = { x: Math.max(area.x, Math.min(Math.round(how.x - 160), area.x + area.width - width)), y: Math.max(area.y, Math.min(Math.round(how.y - 30), area.y + area.height - height)), width, height };
  } else {
    const b = from.getBounds();
    const width = Math.min(b.width, area.width), height = Math.min(b.height, area.height);
    bounds = { x: Math.min(b.x + 48, area.x + area.width - width), y: Math.min(b.y + 48, area.y + area.height - height), width, height };
  }
  const w = createWindow({ bounds, adopt: doc });
  w.focus();
  return true;
});
// The other Doc Studio windows, so a tab can be moved into one of them.
ipcMain.handle('win:list', (e) => BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed() && w.webContents !== e.sender).map((w) => ({ id: w.id, title: w.getTitle() })));
ipcMain.handle('win:moveTo', (e, id, doc) => {
  const w = BrowserWindow.fromId(id);
  if (!w || w.isDestroyed()) return false;
  w.webContents.send('app:adopt', doc);
  if (w.isMinimized()) w.restore();
  w.focus();
  return true;
});
ipcMain.handle('win:tile', (e) => {
  // Arrange every Doc Studio window side by side on this screen.
  const { screen } = require('electron');
  const wins = BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed() && w.isVisible());
  const area = screen.getDisplayMatching(W(e).getBounds()).workArea;
  const each = Math.floor(area.width / wins.length);
  wins.forEach((w, i) => { if (w.isMaximized()) w.unmaximize(); w.setBounds({ x: area.x + i * each, y: area.y, width: i === wins.length - 1 ? area.width - each * i : each, height: area.height }); });
  return wins.length;
});
ipcMain.handle('win:count', () => BrowserWindow.getAllWindows().filter((w) => !w.isDestroyed()).length);

// A moved document is only closed in the window it came from once the receiving window has opened it.
const pendingMoves = new Map(); // token -> webContents id of the sending window
ipcMain.handle('win:id', (e) => W(e).id);
ipcMain.handle('win:expect', (e, token) => { pendingMoves.set(token, e.sender.id); return true; });
ipcMain.handle('win:adopted', (e, token, ok, message) => {
  const fromId = pendingMoves.get(token);
  pendingMoves.delete(token);
  const from = BrowserWindow.getAllWindows().find((w) => !w.isDestroyed() && w.webContents.id === fromId);
  if (from) from.webContents.send('app:adopted', token, ok, message || '');
  return true;
});
// Drag a tab onto another Doc Studio window: that window asks the one the tab came from to send it over.
ipcMain.handle('win:requestTab', (e, fromWinId, tabId) => {
  const from = BrowserWindow.fromId(fromWinId);
  if (!from || from.isDestroyed() || from.webContents === e.sender) return false;
  from.webContents.send('app:give-tab', tabId, W(e).id);
  return true;
});
