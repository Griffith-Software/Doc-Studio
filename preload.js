'use strict';
const { contextBridge, ipcRenderer, webUtils } = require('electron');

// Questions and errors ("Save changes?", "Couldn't open…") are drawn in the page, in Doc Studio's own style,
// instead of Windows message boxes. Same options as Electron's showMessageBox: message, detail, buttons,
// defaultId, cancelId, type ('warning', 'error', 'question', 'info'). Resolves to the index of the button pressed.
const ICONS = {
  warning: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/><path d="M12 9v4M12 17h.01"/></svg>',
  error: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><path d="m15 9-6 6M9 9l6 6"/></svg>',
  info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></svg>',
};
function messageBox(opts) {
  return new Promise((resolve) => {
    const o = opts || {};
    const buttons = o.buttons && o.buttons.length ? o.buttons : ['OK'];
    const def = o.defaultId ?? 0;
    let cancel = o.cancelId;
    if (cancel == null) { cancel = buttons.findIndex((b) => /^(cancel|no|close|ok)$/i.test(b)); if (cancel < 0) cancel = buttons.length - 1; }
    const type = o.type === 'error' ? 'error' : o.type === 'warning' ? 'warning' : 'info';
    const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
    const back = el('div', 'backdrop msgbox-back');
    const box = el('div', `dialog msgbox msgbox-${type}`);
    box.setAttribute('role', 'alertdialog');
    box.setAttribute('aria-modal', 'true');
    const icon = el('span', 'msgbox-icon');
    icon.innerHTML = ICONS[type];
    const text = el('div', 'msgbox-text');
    const title = el('h3', null, o.message || '');
    title.id = `msgbox-${Date.now()}`;
    box.setAttribute('aria-labelledby', title.id);
    text.append(title);
    if (o.detail) text.append(el('p', 'msgbox-detail', o.detail));
    const head = el('div', 'msgbox-head');
    head.append(icon, text);
    const actions = el('div', 'dialog-actions');
    const before = document.activeElement;
    const done = (i) => { back.remove(); try { if (before && before.isConnected) before.focus(); } catch { /* gone */ } resolve(i); };
    // The default button is the filled one, on the right; the others are plain, in their given order.
    const order = buttons.map((_, i) => i).filter((i) => i !== def).concat(def < buttons.length ? [def] : []);
    let first = null;
    for (const i of order) {
      const b = el('button', i === def ? 'msgbox-default' : 'ghost', buttons[i]);
      b.type = 'button';
      b.addEventListener('click', () => done(i));
      actions.append(b);
      if (i === def) first = b;
    }
    box.append(head, actions);
    back.append(box);
    // Keys stay in the box: Esc answers with the cancel button, Enter with the focused one, Tab moves between them.
    back.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); done(cancel); }
      else if (e.key === 'Enter') { e.preventDefault(); const f = document.activeElement; (actions.contains(f) ? f : first || actions.querySelector('button')).click(); }
      else if (e.key === 'Tab') {
        const bs = [...actions.querySelectorAll('button')];
        const k = bs.indexOf(document.activeElement);
        e.preventDefault();
        bs[(k + (e.shiftKey ? bs.length - 1 : 1)) % bs.length].focus();
      }
      e.stopPropagation();
    });
    back.addEventListener('keyup', (e) => e.stopPropagation());
    document.body.append(back);
    (first || actions.querySelector('button')).focus();
  });
}


contextBridge.exposeInMainWorld('desk', {
  openDialog: () => ipcRenderer.invoke('file:openDialog'),
  read: (file) => ipcRenderer.invoke('file:read', file),
  write: (file, data, opts) => ipcRenderer.invoke('file:write', file, data, opts),
  saveDialog: (defaultPath, ext, others) => ipcRenderer.invoke('file:saveDialog', defaultPath, ext, others),
  confirm: (opts) => messageBox(opts),
  error: (message, detail) => messageBox({ type: 'error', message, detail }).then(() => undefined),
  reveal: (file) => ipcRenderer.invoke('ui:reveal', file),
  exists: (p) => ipcRenderer.invoke('fs:exists', p),
  asset: (rel) => ipcRenderer.invoke('app:asset', rel),
  scanners: () => ipcRenderer.invoke('scan:devices'),
  scanPage: (opts) => ipcRenderer.invoke('scan:page', opts),
  aiInfo: () => ipcRenderer.invoke('ai:info'),
  aiSetEnabled: (on) => ipcRenderer.invoke('ai:setEnabled', on),
  aiInstallClaudeDesktop: () => ipcRenderer.invoke('ai:installClaudeDesktop'),
  phoneStart: () => ipcRenderer.invoke('phone:start'),
  phoneStop: () => ipcRenderer.invoke('phone:stop'),
  phoneOffer: (items) => ipcRenderer.invoke('phone:offer', items),
  phoneUnoffer: (id) => ipcRenderer.invoke('phone:unoffer', id),
  phoneRead: (id) => ipcRenderer.invoke('phone:read', id),
  phoneSaveTo: (ids, folder) => ipcRenderer.invoke('phone:saveTo', ids, folder),
  phoneForget: (id) => ipcRenderer.invoke('phone:forget', id),
  // Phone link events: connected, incoming, progress, received, failed, downloaded, stopped. Returns a function that stops listening.
  onPhone: (fn) => { const names = ['connected', 'incoming', 'progress', 'received', 'failed', 'downloaded', 'stopped']; const subs = names.map((n) => { const f = (_e, d) => fn(n, d); ipcRenderer.on(`phone:${n}`, f); return [`phone:${n}`, f]; }); return () => subs.forEach(([n, f]) => ipcRenderer.removeListener(n, f)); },
  openExternal: (url) => ipcRenderer.invoke('ui:openExternal', url),
  defaultApps: () => ipcRenderer.invoke('ui:defaultApps'),
  defaultsStatus: () => ipcRenderer.invoke('defaults:status'),
  revealDefaultSample: (ext) => ipcRenderer.invoke('defaults:reveal', ext),
  checkDefault: (ext) => ipcRenderer.invoke('defaults:check', ext),
  onDefaultPicked: (fn) => ipcRenderer.on('app:default-picked', (_e, ext) => fn(ext)),
  find: (text, opts) => ipcRenderer.invoke('ui:find', text, opts),
  stopFind: () => ipcRenderer.invoke('ui:stopFind'),
  exportPdf: (defaultPath) => ipcRenderer.invoke('ui:exportPdf', defaultPath),
  print: () => ipcRenderer.invoke('ui:print'),
  closeConfirmed: () => ipcRenderer.invoke('app:closeConfirmed'),
  docxBlank: () => ipcRenderer.invoke('docx:blank'),
  backup: (file) => ipcRenderer.invoke('file:backup', file),
  recoverySave: (key, meta, bytes) => ipcRenderer.invoke('recovery:save', key, meta, bytes),
  recoveryClear: (keys) => ipcRenderer.invoke('recovery:clear', keys),
  recoveryTake: () => ipcRenderer.invoke('recovery:take'),
  docxFromHtml: (html, title) => ipcRenderer.invoke('docx:fromHtml', html, title),
  docExtract: (data) => ipcRenderer.invoke('doc:extract', data),
  pathForFile: (file) => webUtils.getPathForFile(file),
  readClipboard: () => ipcRenderer.invoke('clipboard:read'),
  readFont: (file) => ipcRenderer.invoke('font:read', file),
  stat: (file) => ipcRenderer.invoke('file:stat', file),
  rename: (file, newName) => ipcRenderer.invoke('file:rename', file, newName),
  versions: (file) => ipcRenderer.invoke('versions:list', file),
  readVersion: (file, id) => ipcRenderer.invoke('versions:read', file, id),
  recent: () => ipcRenderer.invoke('recent:list'),
  editOp: (op) => ipcRenderer.invoke('edit:op', op),
  setSpellcheck: (on) => ipcRenderer.invoke('spell:set', on),
  appVersion: () => ipcRenderer.invoke('app:version'),
  dictionary: (word) => ipcRenderer.invoke('dict:lookup', word),
  openBackups: () => ipcRenderer.invoke('ui:openBackups'),
  fullscreen: () => ipcRenderer.invoke('ui:fullscreen'),
  theme: (t) => ipcRenderer.invoke('ui:theme', t),
  places: () => ipcRenderer.invoke('fs:places'),
  list: (dir) => ipcRenderer.invoke('fs:list', dir),
  docIndex: () => ipcRenderer.invoke('docs:index'),
  starred: () => ipcRenderer.invoke('starred:list'),
  toggleStar: (file) => ipcRenderer.invoke('starred:toggle', file),
  pick: (opts) => ipcRenderer.invoke('file:pick', opts),
  ocr: (images) => ipcRenderer.invoke('ocr:images', images),
  pdfFromHtml: (html, opts) => ipcRenderer.invoke('pdf:fromHtml', html, opts),
  pngFromHtml: (pages, width, height) => ipcRenderer.invoke('png:fromHtml', pages, width, height),
  convertTools: () => ipcRenderer.invoke('convert:tools'),
  convertRoute: (from, to) => ipcRenderer.invoke('convert:route', from, to),
  convert: (data, from, to) => ipcRenderer.invoke('convert:run', data, from, to),
  popOut: (doc, how) => ipcRenderer.invoke('win:popOut', doc, how),
  windows: () => ipcRenderer.invoke('win:list'),
  moveTo: (id, doc) => ipcRenderer.invoke('win:moveTo', id, doc),
  tileWindows: () => ipcRenderer.invoke('win:tile'),
  windowCount: () => ipcRenderer.invoke('win:count'),
  windowId: () => ipcRenderer.invoke('win:id'),
  expectMove: (token) => ipcRenderer.invoke('win:expect', token),
  adopted: (token, ok, message) => ipcRenderer.invoke('win:adopted', token, ok, message),
  requestTab: (fromWinId, tabId) => ipcRenderer.invoke('win:requestTab', fromWinId, tabId),
  onAdopt: (fn) => ipcRenderer.on('app:adopt', (_e, doc) => fn(doc)),
  onAdopted: (fn) => ipcRenderer.on('app:adopted', (_e, token, ok, message) => fn(token, ok, message)),
  onGiveTab: (fn) => ipcRenderer.on('app:give-tab', (_e, tabId, toWin) => fn(tabId, toWin)),
  onOpenPath: (fn) => ipcRenderer.on('app:open-path', (_e, p) => fn(p)),
  onCommand: (fn) => ipcRenderer.on('app:command', (_e, cmd, arg) => fn(cmd, arg)),
  onCloseRequest: (fn) => ipcRenderer.on('app:close-request', () => fn()),
});
