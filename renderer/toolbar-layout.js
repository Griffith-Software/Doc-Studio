'use strict';
// The document's two toolbar lines.
//   Line 1: page controls, zoom, then the main editing tools (bold, font, alignment...), then View/Edit and layout.
//   Line 2: the menus (File, Edit, View...) while editing, then the secondary tools, then the toolbar settings.
// Tools that don't fit go into a » menu instead of wrapping onto a third line. Toolbar settings (the gear at the
// end of line 2) let people choose, per kind of document, which tools sit on line 1, line 2 or only in the menus,
// and in what order. Window buttons (split view, pop out, full screen) live on the right of the action bar.

const Toolbar = (() => {
  const PREF = (kind) => `docstudio.toolbar.${kind}`;
  const load = (kind) => { try { return JSON.parse(localStorage.getItem(PREF(kind)) || 'null'); } catch { return null; } };
  const store = (kind, cfg) => { try { if (cfg) localStorage.setItem(PREF(kind), JSON.stringify(cfg)); else localStorage.removeItem(PREF(kind)); } catch { /* not kept */ } };
  const GEAR = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 7h10M18 7h2M4 17h4M12 17h8"/><circle cx="16" cy="7" r="2"/><circle cx="10" cy="17" r="2"/></svg>';
  const MORE = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 7l5 5-5 5M13 7l5 5-5 5"/></svg>';

  // Which tools start on line 1, by kind of document.
  const PRIMARY = {
    docx: /^(undo|redo|paragraph style|font|font size|decrease font size|increase font size|bold|italic|underline|text colou?r|highlight|align left|center|align right|justify|bulleted list|numbered list)$/i,
    odt: /^(undo|redo|style|paragraph style|font|font size|size|bold|italic|underline|text colou?r|highlight|align left|center|align right|justify|bulleted list|numbered list)$/i,
    sheet: /^(undo|redo|font|font size|bold|italic|underline|strikethrough|font colou?r|text colou?r|fill colou?r|borders|align left|center|align right|merge.*|number format|currency.*|percent.*)$/i,
    pptx: /^(undo|redo|present|new slide.*|font|font size|size|bold|italic|underline|strikethrough|text colou?r|align left|center|align right)$/i,
    pdf: /^(undo|redo|select.*|add text|edit text.*|highlight.*|freehand.*|draw.*|sticky note.*|note.*|signature.*|ink and text colou?r|highlight colou?r)$/i,
  };
  PRIMARY.rtf = PRIMARY.odt;
  const ALL_PRIMARY = new Set(['md', 'html', 'text', 'csv']); // short toolbars: everything on line 1

  const labelOf = (el) => {
    if (!el || el.nodeType !== 1) return '';
    const t = el.title || el.getAttribute('aria-label') || (el.tagName === 'SELECT' ? '' : el.textContent);
    return String(t || '').replace(/\s*\(.*\)\s*$/, '').replace(/[:.…]+$/, '').trim();
  };
  const isSep = (el) => el && el.classList && el.classList.contains('tool-sep');
  const usable = (el) => el && el.nodeType === 1 && !isSep(el) && !(el.tagName === 'INPUT' && el.type === 'file') && !el.hidden && labelOf(el);

  let last = null; // { tab, items }
  const row1 = () => document.querySelector('#tb-primary');
  const row2 = () => document.querySelector('#tool-left');
  const holder = h('div', { id: 'tb-holder', 'aria-hidden': 'true' });
  document.body.append(holder);

  function config(tab) {
    const kind = tab.kind;
    const saved = load(kind) || {};
    return { kind, order: saved.order || null, place: saved.place || {} };
  }
  function placeOf(cfg, key) {
    if (cfg.place[key]) return cfg.place[key];
    if (ALL_PRIMARY.has(cfg.kind)) return 'top';
    const re = PRIMARY[cfg.kind];
    return re && re.test(key) ? 'top' : 'second';
  }

  // items: everything the editor offers (in its own order, with separators)
  function layout(tab, items) {
    last = { tab, items };
    const top = row1(), second = row2();
    holder.replaceChildren();
    if (!tab || !items.length) { plan = null; top.replaceChildren(); second.replaceChildren(); return; }
    const cfg = config(tab);
    const tools = items.filter(usable);
    const hidden = items.filter((el) => el && el.nodeType === 1 && !isSep(el) && !usable(el)); // file inputs etc. stay in the page
    // Order: the person's own order when set, otherwise the editor's (with its separators).
    let ordered;
    if (cfg.order) {
      const idx = new Map(cfg.order.map((k, i) => [k, i]));
      ordered = tools.map((el, i) => ({ el, i, k: labelOf(el) })).sort((a, b) => (idx.has(a.k) ? idx.get(a.k) : 1000 + a.i) - (idx.has(b.k) ? idx.get(b.k) : 1000 + b.i)).map((x) => x.el);
    } else ordered = items.filter((el) => el && el.nodeType === 1 && (isSep(el) || usable(el)));
    const lines = { top: [], second: [] };
    let lastPlace = null;
    for (const el of ordered) {
      if (isSep(el)) { if (lastPlace) lines[lastPlace].push(el); continue; }
      const p = placeOf(cfg, labelOf(el));
      if (p === 'menu') { holder.append(el); continue; }
      lines[p].push(el);
      lastPlace = p;
    }
    const tidy = (arr) => arr.filter((el, i) => !isSep(el) || (i > 0 && i < arr.length - 1 && !isSep(arr[i - 1])));
    plan = { top: tidy(lines.top), second: tidy(lines.second), hidden, menu: [...holder.children] };
    fit();
    requestAnimationFrame(() => requestAnimationFrame(fit));
  }
  let plan = null;

  // Line 1 overflow moves to the front of line 2; line 2 overflow goes into the » menu.
  const overflowBtn = h('button', { id: 'tb-overflow', class: 'tool', type: 'button', title: 'More tools', 'aria-label': 'More tools', 'data-menu-anchor': '' });
  overflowBtn.innerHTML = MORE;
  let overflowed = [];
  // Always starts again from the plan, so it can run any number of times (after a resize, fonts loading...).
  function fit() {
    const top = row1(), second = row2();
    if (!plan || !top) return;
    overflowed = [];
    overflowBtn.remove();
    top.replaceChildren(...plan.top);
    second.replaceChildren(...plan.second, ...plan.hidden);
    holder.replaceChildren(...plan.menu);
    if (!top.offsetParent) return;
    while (top.scrollWidth > top.clientWidth + 1 && top.lastElementChild) {
      const el = top.lastElementChild;
      if (isSep(el)) { el.remove(); continue; }
      second.prepend(el);
    }
    while (top.lastElementChild && isSep(top.lastElementChild)) top.lastElementChild.remove();
    if (second.offsetParent && second.scrollWidth > second.clientWidth + 1) {
      second.append(overflowBtn);
      while (second.scrollWidth > second.clientWidth + 1) {
        const kids = [...second.children].filter((x) => x !== overflowBtn && !(x.tagName === 'INPUT' && x.type === 'file'));
        const el = kids[kids.length - 1];
        if (!el) break;
        if (isSep(el)) { el.remove(); continue; }
        overflowed.unshift(el);
        holder.append(el);
      }
      if (!overflowed.length) overflowBtn.remove();
    }
  }
  overflowBtn.addEventListener('click', () => {
    Menus.popup(overflowBtn, overflowed.map((el) => {
      const label = labelOf(el);
      if (el.tagName === 'SELECT') return { label, sub: [...el.options].filter((o) => o.value !== '').map((o) => ({ label: o.textContent, checked: () => el.value === o.value, run: () => { el.value = o.value; el.dispatchEvent(new Event('change', { bubbles: true })); } })) };
      return { label: label + (el.tagName === 'INPUT' && el.type === 'color' ? '…' : ''), run: () => el.click() };
    }));
  });
  // The bars change size with the window, the sidebar and split view: fit again.
  let fitTimer = 0;
  const refit = () => { clearTimeout(fitTimer); fitTimer = setTimeout(fit, 60); };
  new ResizeObserver(refit).observe(document.querySelector('#docbar'));
  new ResizeObserver(refit).observe(document.querySelector('#toolbar'));
  let resizeTimer = 0;
  addEventListener('resize', () => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => { if (last && last.tab === active) layout(last.tab, last.items); }, 120); });

  // ---------------------------------------------------------------- toolbar settings

  async function settings() {
    if (!last || !last.tab) return;
    const tab = last.tab;
    const cfg = config(tab);
    const tools = last.items.filter(usable);
    const rows = (cfg.order ? tools.slice().sort((a, b) => { const ia = cfg.order.indexOf(labelOf(a)), ib = cfg.order.indexOf(labelOf(b)); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib); }) : tools)
      .map((el) => ({ el, key: labelOf(el), place: placeOf(cfg, labelOf(el)) }));
    const box = h('div', { class: 'dialog tb-settings' },
      h('h3', { text: `Toolbar - ${FORMATS.label(tab.kind)}` }),
      h('p', { class: 'muted', text: 'Choose where each tool goes and drag (or use ↑ ↓) to reorder. Tools set to "Menus only" are still in the menus and the search. Line 1 tools that do not fit the window move to the start of line 2 (put the ones you want first at the top of the list).' }));
    const back = h('div', { class: 'backdrop' }, box);
    const list = h('div', { class: 'tb-list' });
    const paint = () => {
      list.replaceChildren(...rows.map((r, i) => {
        const icon = h('span', { class: 'tb-icon' });
        if (r.el.tagName === 'BUTTON') icon.innerHTML = r.el.innerHTML; else if (r.el.tagName === 'SELECT') icon.textContent = '▾'; else if (r.el.type === 'color') { icon.textContent = '■'; icon.style.color = r.el.value; }
        const seg = h('div', { class: 'seg tb-seg' }, ...[['top', 'Line 1'], ['second', 'Line 2'], ['menu', 'Menus only']].map(([v, t]) => h('button', { type: 'button', class: r.place === v ? 'on' : '', text: t, onclick: () => { r.place = v; paint(); } })));
        const row = h('div', { class: 'tb-row', draggable: 'true' }, h('span', { class: 'tb-grip', text: '⋮⋮', title: 'Drag to reorder' }), icon, h('span', { class: 'tb-name', text: r.key }), seg,
          h('button', { class: 'ghost tb-move', type: 'button', text: '↑', title: 'Move up', disabled: i === 0, onclick: () => { [rows[i - 1], rows[i]] = [rows[i], rows[i - 1]]; paint(); } }),
          h('button', { class: 'ghost tb-move', type: 'button', text: '↓', title: 'Move down', disabled: i === rows.length - 1, onclick: () => { [rows[i + 1], rows[i]] = [rows[i], rows[i + 1]]; paint(); } }));
        row.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/x-tb-row', String(i)); row.classList.add('dragging'); });
        row.addEventListener('dragend', () => row.classList.remove('dragging'));
        row.addEventListener('dragover', (e) => { if ([...e.dataTransfer.types].includes('text/x-tb-row')) { e.preventDefault(); row.classList.add('drop'); } });
        row.addEventListener('dragleave', () => row.classList.remove('drop'));
        row.addEventListener('drop', (e) => { e.preventDefault(); row.classList.remove('drop'); const from = +e.dataTransfer.getData('text/x-tb-row'); if (isNaN(from) || from === i) return; const [m] = rows.splice(from, 1); rows.splice(i, 0, m); paint(); });
        return row;
      }));
    };
    paint();
    const close = () => back.remove();
    const apply = (save) => {
      if (save) store(tab.kind, { order: rows.map((r) => r.key), place: Object.fromEntries(rows.map((r) => [r.key, r.place])) });
      close();
      refreshChrome();
    };
    box.append(list, h('div', { class: 'dialog-actions', style: 'justify-content:space-between' },
      h('button', { class: 'ghost', type: 'button', text: 'Reset to default', onclick: () => { store(tab.kind, null); close(); refreshChrome(); flash('Toolbar reset'); } }),
      h('span', {}, h('button', { class: 'ghost', type: 'button', text: 'Cancel', onclick: close }), ' ', h('button', { type: 'button', text: 'Save', onclick: () => apply(true) }))));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
  }

  // ---------------------------------------------------------------- moving things into place

  // Line 1 gets a slot for the main tools; line 2 (the old toolbar row) gets the menus and the settings gear.
  const prim = h('div', { id: 'tb-primary', role: 'toolbar', 'aria-label': 'Main tools' });
  document.querySelector('#zoombox').after(prim);
  const bar = document.querySelector('#toolbar');
  bar.prepend(document.querySelector('#menubar'));
  const gear = h('button', { id: 'tb-settings', class: 'tool', type: 'button', title: 'Toolbar settings: choose and order your tools', 'aria-label': 'Toolbar settings' });
  gear.innerHTML = GEAR;
  gear.addEventListener('click', () => settings());
  document.querySelector('#tool-right').append(gear);
  // Window buttons: right end of the action bar.
  const right = h('div', { id: 'ab-right' });
  for (const id of ['#db-split', '#db-popout', '#db-full']) { const b = document.querySelector(id); if (b) { b.classList.remove('db-icon'); b.classList.add('ab', 'ab-icon'); right.append(b); } }
  document.querySelector('#actionbar').append(right);

  return { layout, settings, fit, labelOf, placeOf: (tab, key) => placeOf(config(tab), key) };
})();
