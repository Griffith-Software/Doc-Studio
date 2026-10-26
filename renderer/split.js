'use strict';
// Split view: two documents side by side in one window. Each half has a small header; clicking in a half makes
// that document the active one (the page bar, toolbars and menus follow it). Choosing a tab puts it in the half
// you are working in. The divider between the halves can be dragged.

const Split = (() => {
  let state = null; // { left: tabId, right: tabId|null, ratio }
  let fillRight = false; // the next document opened goes into the empty right half
  const byId = (id) => tabs.find((t) => t.id === id) || null;
  const ICON = '<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="1.5"/><path d="M12 4v16"/></svg>';
  const stage = document.querySelector('#stage');
  const divider = h('div', { id: 'split-divider', hidden: true, role: 'separator', 'aria-orientation': 'vertical', title: 'Drag to resize' });
  const picker = h('div', { id: 'split-picker', hidden: true });
  stage.append(divider, picker);

  const isOn = () => !!state;
  const sideOf = (tab) => (!state || !tab ? null : state.left === tab.id ? 'left' : state.right === tab.id ? 'right' : null);

  function open(rightTab) {
    if (!active) return;
    if (rightTab === active) rightTab = null;
    state = { left: active.id, right: rightTab ? rightTab.id : null, ratio: state ? state.ratio : 0.5 };
    fillRight = !rightTab;
    layout();
    refreshChrome();
  }
  function close() {
    if (!state) return;
    state = null;
    fillRight = false;
    layout();
    refreshChrome();
  }
  function swap() {
    if (!state || !state.right) return;
    [state.left, state.right] = [state.right, state.left];
    layout();
    refreshChrome();
  }

  // Put the halves in place (called after every change and every chrome refresh).
  function layout() {
    const left = state && byId(state.left), right = state && state.right != null ? byId(state.right) : null;
    if (state && !left) { if (right) { state.left = right.id; state.right = null; return layout(); } state = null; }
    stage.classList.toggle('split', !!state);
    for (const t of tabs) {
      const side = sideOf(t);
      t.el.classList.toggle('split-l', side === 'left');
      t.el.classList.toggle('split-r', side === 'right');
      t.el.style.left = side === 'right' ? `calc(${state.ratio * 100}% + 3px)` : '';
      t.el.style.right = side === 'left' ? `calc(${(1 - state.ratio) * 100}% + 3px)` : '';
      if (side) head(t);
    }
    divider.hidden = !state;
    if (state) divider.style.left = `calc(${state.ratio * 100}% - 3px)`;
    picker.hidden = !(state && state.right == null);
    if (!picker.hidden) { picker.style.left = `calc(${state.ratio * 100}% + 3px)`; paintPicker(); }
    refit();
  }
  // Documents drawn at a fixed size are fitted again to their new width.
  let refitTimer = 0;
  function refit() {
    clearTimeout(refitTimer);
    refitTimer = setTimeout(() => {
      // (Sheets, slides and Markdown follow their own size; PDF pages are drawn at a size.)
      for (const t of tabs) if (t.pdfEditor && t.el.offsetParent && t.pdfFitWidth !== t.el.clientWidth) { t.pdfFitWidth = t.el.clientWidth; t.pdfEditor.fitWidth(t.layout === 'two' ? 2 : 1); t.zoom = 1; }
    }, 120);
  }

  // The small header on each half: name, swap, close.
  function head(t) {
    if (!t.splitHead) {
      const name = h('span', { class: 'split-name' });
      const swapBtn = h('button', { class: 'pane-split-btn', type: 'button', title: 'Swap sides', text: '⇄' });
      const closeBtn = h('button', { class: 'pane-split-btn', type: 'button', title: 'Close split view (this document stays open)', text: '✕' });
      swapBtn.addEventListener('click', (e) => { e.stopPropagation(); swap(); });
      closeBtn.addEventListener('click', (e) => { e.stopPropagation(); activate(t); close(); });
      t.splitHead = h('div', { class: 'split-head' }, name, swapBtn, closeBtn);
      t.splitHead.addEventListener('mousedown', () => { if (active !== t) activate(t); });
      t.el.prepend(t.splitHead);
    }
    t.splitHead.querySelector('.split-name').textContent = `${t.name}${t.dirty ? ' ●' : ''}`;
  }

  // The empty right half: choose what goes there.
  function paintPicker() {
    const others = tabs.filter((t) => !sideOf(t));
    const row = (label, sub, fn) => { const b = h('button', { class: 'pick-row', type: 'button' }, h('b', { text: label }), h('small', { text: sub || '' })); b.addEventListener('click', fn); return b; };
    picker.replaceChildren(h('div', { class: 'pick-box' },
      h('h3', { text: 'Choose a document for this side' }),
      ...(others.length ? [h('div', { class: 'pick-group', text: 'Open documents' }), ...others.map((t) => row(t.name, t.path || 'Not saved yet', () => { state.right = t.id; fillRight = false; layout(); activate(t); }))] : []),
      h('div', { class: 'pick-group', text: 'Or' }),
      row('Open a file…', 'Ctrl+O', () => { fillRight = true; openDialog(); }),
      row('New document (.docx)', '', () => { fillRight = true; newDocument('docx'); }),
      row('New Markdown note', '', () => { fillRight = true; newDocument('md'); }),
      h('button', { class: 'ghost pick-cancel', type: 'button', text: 'Cancel split view', onclick: () => close() })));
  }

  // Which half a newly chosen document goes into: the one you are working in.
  const activate0 = activate;
  activate = function (tab) { // eslint-disable-line no-global-assign
    if (state && tab && !sideOf(tab)) {
      if (state.right == null && fillRight) { state.right = tab.id; fillRight = false; }
      else if (sideOf(active) === 'right') state.right = tab.id;
      else state.left = tab.id;
    }
    const r = activate0(tab);
    if (state) layout();
    return r;
  };
  // Closing one half leaves the other on its own.
  const closeTab0 = closeTab;
  closeTab = async function (tab) { // eslint-disable-line no-global-assign
    const side = sideOf(tab);
    const r = await closeTab0(tab);
    if (r && side && state) { state = null; fillRight = false; layout(); refreshChrome(); }
    return r;
  };
  const refreshChrome0 = refreshChrome;
  refreshChrome = function () { // eslint-disable-line no-global-assign
    refreshChrome0();
    if (state) layout();
    document.querySelectorAll('#tabs .tab').forEach((el, i) => el.classList.toggle('in-split', !!sideOf(tabs[i]) && tabs[i] !== active));
    const b = document.querySelector('#db-split');
    if (b) { b.classList.toggle('on', !!state); b.title = state ? 'Close split view' : 'Split view: two documents side by side'; }
  };

  // Clicking or typing in a half makes it the active document.
  stage.addEventListener('mousedown', (e) => {
    if (!state) return;
    const pane = e.target.closest('.pane');
    const t = pane && tabs.find((x) => x.el === pane);
    if (t && t !== active) activate(t);
  }, true);
  stage.addEventListener('focusin', (e) => {
    if (!state) return;
    const pane = e.target.closest('.pane');
    const t = pane && tabs.find((x) => x.el === pane);
    if (t && t !== active) activate(t);
  });

  // Drag the divider.
  divider.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || !state) return;
    e.preventDefault();
    document.body.classList.add('resizing-split');
    const r = stage.getBoundingClientRect();
    const move = (ev) => { state.ratio = Math.min(0.85, Math.max(0.15, (ev.clientX - r.left) / r.width)); layout(); };
    const up = () => { window.removeEventListener('mousemove', move, true); window.removeEventListener('mouseup', up, true); document.body.classList.remove('resizing-split'); layout(); };
    window.addEventListener('mousemove', move, true);
    window.addEventListener('mouseup', up, true);
  });
  divider.addEventListener('dblclick', () => { if (state) { state.ratio = 0.5; layout(); } });

  // Menu of documents for the right half.
  function menuItems() {
    if (state) return [{ label: 'Close split view', run: () => close() }, { label: 'Swap sides', run: () => swap(), enabled: () => !!(state && state.right != null) }];
    const others = tabs.filter((t) => t !== active);
    return [
      ...(others.length ? [{ label: 'Show beside this document', sub: others.map((t) => ({ label: t.name, run: () => open(t) })) }] : []),
      { label: 'Open a file beside it…', run: () => { open(null); fillRight = true; openDialog(); } },
      { label: 'New document beside it', sub: [['Markdown note', 'md'], ['Document (.docx)', 'docx'], ['Spreadsheet (.xlsx)', 'xlsx'], ['Plain text', 'txt']].map(([l, x]) => ({ label: l, run: () => { open(null); fillRight = true; newDocument(x); } })) },
      { label: 'Choose later', run: () => open(null) },
    ];
  }
  const btn = h('button', { id: 'db-split', class: 'db-icon', type: 'button', title: 'Split view: two documents side by side', 'aria-label': 'Split view', 'data-menu-anchor': '' });
  btn.innerHTML = ICON;
  btn.addEventListener('click', () => { if (!active) return; if (state) close(); else Menus.popup(btn, menuItems()); });
  document.querySelector('#db-popout').before(btn);
  addEventListener('resize', () => { if (state) refit(); });

  return { open, close, swap, isOn, sideOf, menuItems, get state() { return state ? { ...state } : null; } };
})();
