'use strict';
// Several documents at once: pop a document out into its own window (or beside this one), move it to another
// Doc Studio window, drag a tab out of the window or onto another window, or arrange every window side by side.
// A moved document stays where it was until the receiving window confirms it has opened it.

const Popout = (() => {
  const TAB_TYPE = 'application/x-docstudio-tab';
  const ICON = '<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="7" width="12" height="13" rx="1.5"/><path d="M10 4h10v10"/><path d="M20 4l-8 8"/></svg>';
  let myWin = null;
  desk.windowId().then((id) => { myWin = id; });
  const moving = new Map(); // token -> { tab, timer }

  // What the receiving window needs: a saved, unchanged file is reopened from disk; anything else travels as
  // its current contents (unsaved edits included).
  async function docOf(tab) {
    const doc = { name: tab.name, path: tab.path || null, mode: tab.mode, dirty: !!tab.dirty, isNew: !!tab.isNew, zoom: tab.zoom || 1, layout: tab.layout || null };
    if (tab.path && !tab.dirty && !tab.original) return doc;
    doc.data = await tabBytes(tab);
    if (doc.data && !(doc.data instanceof ArrayBuffer)) doc.data = doc.data.buffer ? doc.data.buffer.slice(doc.data.byteOffset, doc.data.byteOffset + doc.data.byteLength) : doc.data;
    if (tab.original) { doc.original = tab.original; doc.kind = tab.kind; }
    return doc;
  }

  // target: 'window' (new window), 'side' (new window, both side by side), { x, y } (new window at a screen
  // point) or the id of another Doc Studio window.
  async function send(tab, target) {
    if (!tab) return;
    if (tab.sheetView && tab.sheetView.editing) tab.sheetView.endEdit(true);
    if (tab.commitText) tab.commitText();
    let doc;
    try { doc = await docOf(tab); } catch (err) { desk.error(`Couldn't move ${tab.name}`, err.message); return; }
    doc.token = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    await desk.expectMove(doc.token);
    const timer = setTimeout(() => { if (moving.delete(doc.token)) desk.error(`${tab.name} couldn't be opened in the other window`, 'It is still open here.'); }, 60000);
    moving.set(doc.token, { tab, timer });
    flash(`Opening ${tab.name} in ${typeof target === 'number' ? 'the other window' : 'a new window'}…`);
    const ok = typeof target === 'number' ? await desk.moveTo(target, doc) : await desk.popOut(doc, target === 'window' ? null : target);
    if (!ok) { clearTimeout(timer); moving.delete(doc.token); desk.error(`Couldn't move ${tab.name}`, 'The other window is no longer open.'); }
  }
  desk.onAdopted(async (token, ok, message) => {
    const m = moving.get(token);
    if (!m) return;
    clearTimeout(m.timer);
    moving.delete(token);
    if (!ok) { desk.error(`${m.tab.name} couldn't be opened in the other window`, `${message}\nIt is still open here.`); return; }
    m.tab.dirty = false;
    await closeTab(m.tab);
  });
  desk.onGiveTab((tabId, toWin) => { const t = tabs.find((x) => x.id === tabId); if (t) send(t, toWin); });

  // Receiving a document.
  desk.onAdopt(async (doc) => {
    try {
      let tab;
      if (!doc.data && doc.path) {
        await openPath(doc.path);
        tab = tabs.find((t) => t.path && t.path.toLowerCase() === doc.path.toLowerCase());
        if (!tab) throw new Error('The file could not be opened.');
      } else {
        tab = makeTab({ path: doc.path, name: doc.name, size: doc.data ? doc.data.byteLength : 0, data: doc.data });
        if (doc.original) { tab.original = doc.original; tab.kind = doc.kind; }
        tab.isNew = doc.isNew;
        if (doc.mode) tab.startMode = doc.mode;
        await setupTab(tab);
      }
      if (doc.mode && tab.modes.includes(doc.mode) && tab.mode !== doc.mode) await setMode(tab, doc.mode);
      if (doc.layout) tab.layout = doc.layout;
      if (doc.dirty) markDirty(tab);
      if (doc.zoom && doc.zoom !== 1 && typeof zoomTab === 'function') zoomTab(tab, 0, doc.zoom);
      activate(tab);
      refreshChrome();
      desk.adopted(doc.token, true);
    } catch (err) {
      console.error(err);
      desk.adopted(doc.token, false, err.message);
    }
  });

  // ---------------------------------------------------------------- menus

  async function items(tab) {
    const others = await desk.windows();
    return [
      { label: 'Open in a new window', run: () => send(tab, 'window') },
      { label: 'Open side by side with this window', run: () => send(tab, 'side') },
      ...(others.length ? [{ label: 'Move to window', sub: others.map((w) => ({ label: w.title || `Window ${w.id}`, run: () => send(tab, w.id) })) }] : []),
      { sep: true },
      { label: 'Arrange all windows side by side', run: async () => { const n = await desk.tileWindows(); if (n < 2) flash('Open a document in another window first (Pop out).'); } },
    ];
  }
  async function tabMenu(tab, x, y) {
    const others = await desk.windows();
    const starred = typeof Shell !== 'undefined' && tab.path && Shell.isStarred(tab.path);
    SheetView.showMenu(x, y, [
      ...(typeof Split !== 'undefined' && tabs.length > 1 && !Split.sideOf(tab) ? [['Show in split view beside the current document', '', () => { const cur = active !== tab ? active : tabs.find((t) => t !== tab); activate(cur); Split.open(tab); }]] : []),
      ...(typeof Split !== 'undefined' && Split.isOn() ? [['Close split view', '', () => Split.close()]] : []),
      ['Open in a new window', '', () => send(tab, 'window')],
      ['Open side by side with this window', '', () => send(tab, 'side')],
      ...others.map((w) => [`Move to: ${w.title || `window ${w.id}`}`, '', () => send(tab, w.id)]),
      ['Arrange all windows side by side', '', () => desk.tileWindows()],
      null,
      ['Rename…', '', () => renameTab(tab)],
      ...(tab.path ? [['Show in folder', '', () => desk.reveal(tab.path)], [starred ? 'Remove from Starred' : 'Star', '', () => Shell.toggleStar(tab.path).then(() => refreshChrome())]] : []),
      null,
      ['Close', 'Ctrl+W', () => closeTab(tab)],
      ...(tabs.length > 1 ? [['Close other tabs', '', async () => { for (const t of tabs.filter((x) => x !== tab)) if (!(await closeTab(t))) break; }]] : []),
    ]);
  }

  // ---------------------------------------------------------------- tabs: drag out, drag between windows, reorder, right-click

  function decorateTabs() {
    const els = [...document.querySelectorAll('#tabs .tab')];
    els.forEach((el, i) => {
      const tab = tabs[i];
      if (!tab || el.dataset.pop) return;
      el.dataset.pop = '1';
      el.draggable = true;
      el.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData(TAB_TYPE, JSON.stringify({ win: myWin, tab: tab.id }));
        e.dataTransfer.effectAllowed = 'move';
        document.body.classList.add('dragging-tab');
      });
      el.addEventListener('dragend', (e) => {
        document.body.classList.remove('dragging-tab');
        // Let go outside every Doc Studio window: the document gets its own window there.
        const outside = e.screenX < window.screenX || e.screenX > window.screenX + window.outerWidth || e.screenY < window.screenY || e.screenY > window.screenY + window.outerHeight;
        if (e.dataTransfer.dropEffect === 'none' && outside) send(tab, { x: e.screenX, y: e.screenY });
      });
      el.addEventListener('contextmenu', (e) => { e.preventDefault(); tabMenu(tab, e.clientX, e.clientY); });
    });
  }
  const isTabDrag = (e) => e.dataTransfer && [...e.dataTransfer.types].includes(TAB_TYPE);
  document.addEventListener('dragover', (e) => { if (isTabDrag(e)) { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; } });
  document.addEventListener('drop', (e) => {
    if (!isTabDrag(e)) return;
    e.preventDefault();
    let info;
    try { info = JSON.parse(e.dataTransfer.getData(TAB_TYPE)); } catch { return; }
    if (info.win !== myWin) { desk.requestTab(info.win, info.tab); return; }
    // Same window: dropped on another tab moves it there.
    const target = e.target.closest && e.target.closest('#tabs .tab');
    const from = tabs.findIndex((t) => t.id === info.tab);
    if (!target || from < 0) return;
    const to = [...document.querySelectorAll('#tabs .tab')].indexOf(target);
    if (to < 0 || to === from) return;
    const [t] = tabs.splice(from, 1);
    tabs.splice(to, 0, t);
    refreshChrome();
  });
  const refreshChrome0 = refreshChrome;
  refreshChrome = function () { refreshChrome0(); decorateTabs(); }; // eslint-disable-line no-global-assign

  // ---------------------------------------------------------------- page bar button

  const btn = h('button', { id: 'db-popout', class: 'db-icon', type: 'button', title: 'Pop out: open this document in its own window', 'aria-label': 'Pop out', 'data-menu-anchor': '' });
  btn.innerHTML = ICON;
  btn.addEventListener('click', async () => { if (active) Menus.popup(btn, await items(active)); });
  const full = document.querySelector('#db-full');
  full.before(btn);

  if (new URLSearchParams(location.search).get('popout')) document.body.classList.add('popout');
  return { send, items, docOf, ICON };
})();
