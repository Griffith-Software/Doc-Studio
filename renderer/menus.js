'use strict';
// The in-window menu bar (File, Edit, View, Insert, Format, Data/Slide/Page, Tools, Help) and "Search the menus".
//
// Menus are built fresh each time one opens, from the active document: shell commands are listed here, and every
// toolbar button and drop-down of the current editor is also offered as a menu item (so anything on a toolbar can
// be found by searching the menus). Editors can add their own items through tab.menuExtras().

const Menus = (() => {
  const SEP = { sep: true };
  let open = null; // { index, panel }
  let names = [];

  const item = (label, run, o) => ({ label, run, ...(o || {}) });
  const sub = (label, items, o) => ({ label, sub: items, ...(o || {}) });
  const isEditable = () => !!active && (['edit', 'split'].includes(active.mode) || active.kind === 'sheet');

  // ---------------------------------------------------------------- toolbar → menu items

  // Which menu a toolbar control belongs to, by its title.
  function menuFor(title, kind) {
    const t = title.toLowerCase();
    if (kind === 'sheet' && /sort|filter|freeze|conditional|recalc/.test(t)) return 'Data';
    if (kind === 'pptx' && /slide/.test(t)) return 'Slide';
    if (kind === 'pdf' && /page|rotate/.test(t)) return 'Page';
    if (kind === 'pdf' && /zoom|fit/.test(t)) return 'View';
    if (kind === 'pdf') return 'Insert';
    if (/^(insert|add|new )|page break|^link|picture|image|table|text box|horizontal|^rule|code block|signature|sticky/.test(t)) return 'Insert';
    if (/^delete|^remove/.test(t) && kind === 'sheet') return 'Edit';
    return 'Format';
  }
  function toolbarItems() {
    const out = {};
    if (!active) return out;
    const add = (m, it) => { (out[m] = out[m] || []).push(it); };
    for (const t of active.tools || []) {
      if (!t || t.nodeType !== 1 || t.hidden) continue;
      if (t.tagName === 'SELECT') {
        const label = (t.title || t.getAttribute('aria-label') || '').replace(/\s*\(.*\)$/, '');
        if (!label) continue;
        const opts = [...t.options].filter((o) => o.value !== '');
        if (!opts.length) continue;
        add(menuFor(label, active.kind), sub(label, opts.map((o) => item(o.textContent, () => { t.value = o.value; t.dispatchEvent(new Event('change')); }, { checked: () => t.value === o.value }))));
      } else if (t.tagName === 'BUTTON' && t.title) {
        const m = /^(.*?)\s*\((.+)\)$/.exec(t.title);
        add(menuFor(t.title, active.kind), item(m ? m[1] : t.title, () => t.click(), { key: m && /ctrl|shift|alt|f\d/i.test(m[2]) ? m[2] : '', enabled: () => !t.disabled, checked: t.classList.contains('on') ? () => true : undefined }));
      } else if (t.tagName === 'INPUT' && t.type === 'color' && t.title) {
        add('Format', item(`${t.title}…`, () => t.click()));
      }
    }
    return out;
  }

  // ---------------------------------------------------------------- the menus

  async function build() {
    const k = active ? active.kind : null;
    const editable = FORMATS.isEditable(k);
    const fromTools = toolbarItems();
    const extras = active ? (active.menuExtras ? await active.menuExtras() : (typeof DocTools !== 'undefined' ? DocTools.docMenus(active) : {})) : {};
    // Toolbar items already offered by an explicit menu item (same label) aren't repeated.
    const norm = (l) => String(l || '').toLowerCase().replace(/[…().:]/g, '').replace(/\s+/g, ' ').trim();
    const explicit = new Set(Object.values(extras).flat().filter((x) => x && x.label).map((x) => norm(x.label)));
    for (const k of Object.keys(fromTools)) fromTools[k] = fromTools[k].filter((x) => !explicit.has(norm(x.label)));
    const merge = (name, base) => [...base, ...(extras[name] && extras[name].length ? [...(base.length ? [SEP] : []), ...extras[name]] : []), ...(fromTools[name] && fromTools[name].length ? [...(base.length || (extras[name] || []).length ? [SEP] : []), ...fromTools[name]] : [])];
    const recent = await desk.recent();
    const targets = active ? await Convert.targets(active) : [];
    const menus = [];

    menus.push({ name: 'File', items: [
      sub('New', [
        item('Markdown', () => newDocument('md'), { key: 'Ctrl+Shift+N' }), SEP,
        item('Document (.docx)', () => newDocument('docx'), { key: 'Ctrl+N' }), item('Spreadsheet (.xlsx)', () => newDocument('xlsx')), item('Presentation (.pptx)', () => newDocument('pptx')), item('PDF', () => newDocument('pdf')), SEP,
        item('OpenDocument text', () => newDocument('odt')), item('OpenDocument spreadsheet', () => newDocument('ods')), item('OpenDocument presentation', () => newDocument('odp')), item('Rich Text', () => newDocument('rtf')), SEP,
        item('Plain text', () => newDocument('txt')), item('Web page (HTML)', () => newDocument('html')), item('CSV table', () => newDocument('csv')), item('JSON', () => newDocument('json')),
      ]),
      item('Open…', () => openDialog(), { key: 'Ctrl+O' }),
      sub('Open recent', recent.length ? recent.map((f) => item(f, () => openPath(f))) : [item('No recent files', null, { enabled: () => false })]),
      item('Make a copy…', () => makeCopy(active), { enabled: () => !!active && !!active.data }),
      SEP,
      item('Save', () => save(active, false), { key: 'Ctrl+S', enabled: () => !!active }),
      item('Save as…', () => save(active, true), { key: 'Ctrl+Shift+S', enabled: () => !!active }),
      sub('Download / convert to', targets.length ? targets.map((t) => item(`${t.label} (.${t.ext})`, () => Convert.convertTab(active, t.ext))) : [item('Nothing to convert', null, { enabled: () => false })], { enabled: () => !!active }),
      item('Convert to…', () => Convert.dialog(active), { enabled: () => !!active }),
      item('Convert files…', () => Convert.batch()),
      SEP,
      item('Rename…', () => renameTab(active), { enabled: () => !!active }),
      item('Version history…', () => showVersions(active), { enabled: () => !!(active && active.path) }),
      item('Show in folder', () => desk.reveal(active.path), { enabled: () => !!(active && active.path) }),
      item('Details', () => showDetails(active), { enabled: () => !!active }),
      SEP,
      item('Print…', () => doPrint(), { key: 'Ctrl+P', enabled: () => !!active }),
      item('Export as PDF…', () => doExportPdf(), { key: 'Ctrl+Shift+E', enabled: () => !!active }),
      SEP,
      item('Close tab', () => closeTab(active), { key: 'Ctrl+W', enabled: () => !!active }),
    ] });

    menus.push({ name: 'Edit', items: merge('Edit', [
      item('Undo', () => command('undo'), { key: 'Ctrl+Z', enabled: () => !!active }),
      item('Redo', () => command('redo'), { key: 'Ctrl+Y', enabled: () => !!active }),
      SEP,
      item('Cut', () => desk.editOp('cut'), { key: 'Ctrl+X' }),
      item('Copy', () => desk.editOp('copy'), { key: 'Ctrl+C' }),
      item('Paste', () => desk.editOp('paste'), { key: 'Ctrl+V' }),
      item('Paste without formatting', () => desk.editOp('pasteAndMatchStyle'), { key: 'Ctrl+Shift+V' }),
      SEP,
      item('Select all', () => desk.editOp('selectAll'), { key: 'Ctrl+A' }),
      SEP,
      item('Find', () => openFind(), { key: 'Ctrl+F', enabled: () => !!active }),
      item('Find and replace', () => openFind(true), { key: 'Ctrl+H', enabled: () => !!active && isEditable() }),
    ]) });

    const modes = active ? active.modes : [];
    menus.push({ name: 'View', items: merge('View', [
      ...(modes.length > 1 ? [sub('Mode', modes.map((m) => item(MODE_LABELS[m], () => setMode(active, m), { checked: () => active.mode === m }))), item('Switch mode', () => command('toggleMode'), { key: 'Ctrl+E' }), SEP] : []),
      item('Zoom in', () => zoomTab(active, 1), { key: 'Ctrl+=', enabled: () => canZoom(active) }),
      item('Zoom out', () => zoomTab(active, -1), { key: 'Ctrl+-', enabled: () => canZoom(active) }),
      item('Actual size (100%)', () => zoomTab(active, 0), { key: 'Ctrl+0', enabled: () => canZoom(active) }),
      SEP,
      item('Full screen', () => desk.fullscreen(), { key: 'F11' }),
      { label: 'Split view', sub: () => Split.menuItems(), enabled: () => !!active },
      item('Open in a new window', () => Popout.send(active, 'window'), { enabled: () => !!active }),
      item('Open side by side with this window', () => Popout.send(active, 'side'), { enabled: () => !!active }),
      { label: 'Move to window', sub: async () => { const w = await desk.windows(); return w.length ? w.map((x) => item(x.title || `Window ${x.id}`, () => Popout.send(active, x.id))) : [item('No other Doc Studio windows', null, { enabled: () => false })]; }, enabled: () => !!active },
      item('Arrange all windows side by side', () => desk.tileWindows()),
      item(compact ? 'Show the sidebar and menus (leave compact mode)' : 'Compact mode (hide the sidebar and menus)', () => toggleCompact(), { key: 'Ctrl+Shift+F' }),
      item(typeof Shell !== 'undefined' && !Shell.sidebarShown() ? 'Show the sidebar' : 'Hide the sidebar', () => Shell.toggleSidebar()),
      sub('Theme', ['system', 'light', 'dark'].map((t) => item(t[0].toUpperCase() + t.slice(1), () => desk.theme(t), { checked: async () => (await desk.theme()) === t }))),
      SEP,
      item('Next tab', () => command('nextTab', 1), { key: 'Ctrl+Tab', enabled: () => tabs.length > 1 }),
      item('Previous tab', () => command('nextTab', -1), { key: 'Ctrl+Shift+Tab', enabled: () => tabs.length > 1 }),
    ]) });

    if (editable || fromTools.Insert || extras.Insert) menus.push({ name: 'Insert', items: merge('Insert', []) });
    if (editable || fromTools.Format || extras.Format) menus.push({ name: 'Format', items: merge('Format', []) });
    if (k === 'sheet') menus.push({ name: 'Data', items: merge('Data', []) });
    if (k === 'pptx') menus.push({ name: 'Slide', items: merge('Slide', []) });
    if (extras.Arrange) menus.push({ name: 'Arrange', items: merge('Arrange', []) });
    if (k === 'pdf') menus.push({ name: 'Page', items: merge('Page', []) });

    menus.push({ name: 'Tools', items: merge('Tools', [
      item('Word count', () => showWordCount(active), { key: 'Ctrl+Shift+C', enabled: () => !!active && ['docx', 'odt', 'rtf', 'md', 'text', 'html', 'pptx', 'pdf', 'epub'].includes(active.kind) }),
      item('Spelling underlines', () => toggleSpelling(), { checked: () => spellingOn() }),
      SEP,
      item('Convert files…', () => Convert.batch()),
      item('Open backups folder', () => command('backups')),
    ]) });

    menus.push({ name: 'Help', items: [
      item('Search the menus', () => focusSearch(), { key: 'Alt+/' }),
      item('Keyboard shortcuts', () => showShortcuts(), { key: 'Ctrl+/' }),
      item('Supported file types', () => showHelp()),
      SEP,
      item('About Doc Studio', () => command('about')),
    ] });

    // Drop empty menus and tidy separators.
    for (const m of menus) m.items = tidy(m.items);
    return menus.filter((m) => m.items.length);
  }
  function tidy(items) {
    const out = [];
    for (const it of items) {
      if (it.sep && (!out.length || out[out.length - 1].sep)) continue;
      out.push(it.sub && Array.isArray(it.sub) ? { ...it, sub: tidy(it.sub) } : it);
    }
    while (out.length && out[out.length - 1].sep) out.pop();
    return out;
  }

  // ---------------------------------------------------------------- rendering

  async function render() {
    const host = $('#menubar');
    if (!host) return;
    const list = await build();
    names = list.map((m) => m.name);
    const row = $('#menu-names');
    row.replaceChildren(...list.map((m, i) => {
      const b = h('button', { class: 'menu-name', type: 'button', 'data-i': i, text: m.name, 'aria-haspopup': 'true' });
      b.addEventListener('mousedown', (e) => { e.preventDefault(); if (open && open.index === i) close(); else show(i); });
      b.addEventListener('mouseenter', () => { if (open && open.index !== i) show(i); });
      return b;
    }));
    const title = $('#doc-title');
    title.textContent = active ? active.name : '';
    title.title = active ? (active.path ? `${active.path}\nClick to rename` : 'Click to rename') : '';
    title.hidden = !active;
    $('#doc-kind').textContent = active ? (active.ext || '?').slice(0, 4).toUpperCase() : '';
    $('#doc-kind').className = active ? `tab-kind k-${active.kind}` : '';
    // The menu names show while editing; otherwise every menu is under More (and Alt+letter).
    host.hidden = compact || !(typeof Shell !== 'undefined' ? Shell.showMenubar() : true);
  }

  async function show(i) {
    close();
    const list = await build();
    const m = list[i];
    if (!m) return;
    let btn = $(`#menu-names [data-i="${i}"]`);
    if (!btn || !btn.offsetParent) btn = $('#ab-more');
    btn.classList.add('open');
    const panel = await panelFor(m.items);
    const r = btn.getBoundingClientRect();
    panel.style.left = `${r.left}px`;
    panel.style.top = `${r.bottom + 2}px`;
    document.body.append(panel);
    keepOnScreen(panel);
    open = { index: i, panel, btn };
    panel.querySelector('.menu-item:not(.disabled)')?.classList.add('hot');
  }
  function close() {
    document.querySelectorAll('.menu-panel:not(#menu-results)').forEach((p) => p.remove());
    document.querySelectorAll('.menu-name.open, [data-menu-anchor].open').forEach((b) => b.classList.remove('open'));
    open = null;
  }
  // A menu of any items under a button (Open ▾, Convert ▾, More ▾ ...). Clicking the button again closes it.
  async function popup(anchor, items) {
    const was = open && open.btn === anchor;
    close();
    if (was) return;
    const panel = await panelFor(tidy(items));
    const r = anchor.getBoundingClientRect();
    panel.style.left = `${r.left}px`;
    panel.style.top = `${r.bottom + 4}px`;
    document.body.append(panel);
    keepOnScreen(panel);
    anchor.classList.add('open');
    open = { index: -1, panel, btn: anchor };
    panel.querySelector('.menu-item:not(.disabled)')?.classList.add('hot');
  }
  function keepOnScreen(panel) {
    const r = panel.getBoundingClientRect();
    if (r.right > innerWidth - 4) panel.style.left = `${Math.max(4, innerWidth - r.width - 4)}px`;
    if (r.bottom > innerHeight - 4) { panel.style.maxHeight = `${innerHeight - r.top - 8}px`; panel.style.overflowY = 'auto'; }
  }

  async function panelFor(items, level) {
    const panel = h('div', { class: 'menu-panel', role: 'menu' });
    panel.dataset.level = level || 0;
    for (const it of items) {
      if (it.sep) { panel.append(h('div', { class: 'menu-sep' })); continue; }
      const enabled = it.enabled ? await it.enabled() : !!(it.run || it.sub);
      const checked = it.checked ? await it.checked() : false;
      const row = h('div', { class: `menu-item${enabled ? '' : ' disabled'}${it.sub ? ' has-sub' : ''}`, role: 'menuitem' },
        h('span', { class: 'menu-check', text: checked ? '✓' : '' }),
        h('span', { class: 'menu-label', text: it.label }),
        h('span', { class: 'menu-key', text: it.sub ? '›' : (it.key || '') }));
      row.addEventListener('mousedown', (e) => e.preventDefault());
      if (enabled && it.sub) {
        const openSub = async () => {
          if (row.classList.contains('sub-open')) return;
          panel.querySelectorAll('.sub-open').forEach((x) => x.classList.remove('sub-open'));
          removeDeeper(+panel.dataset.level);
          row.classList.add('sub-open');
          const subs = typeof it.sub === 'function' ? await it.sub() : it.sub;
          const sp = await panelFor(tidy(subs), +panel.dataset.level + 1);
          const r = row.getBoundingClientRect();
          sp.style.left = `${r.right - 2}px`;
          sp.style.top = `${r.top - 4}px`;
          document.body.append(sp);
          const sr = sp.getBoundingClientRect();
          if (sr.right > innerWidth - 4) sp.style.left = `${Math.max(4, r.left - sr.width + 2)}px`;
          keepOnScreen(sp);
          return sp;
        };
        row.addEventListener('mouseenter', openSub);
        row.addEventListener('click', openSub);
        row._openSub = openSub;
      } else {
        row.addEventListener('mouseenter', () => { panel.querySelectorAll('.sub-open').forEach((x) => x.classList.remove('sub-open')); removeDeeper(+panel.dataset.level); });
      }
      if (enabled && it.run) row.addEventListener('click', () => { close(); setTimeout(() => runItem(it), 0); });
      row.addEventListener('mousemove', () => { panel.querySelectorAll('.hot').forEach((x) => x.classList.remove('hot')); row.classList.add('hot'); });
      panel.append(row);
    }
    return panel;
  }
  function removeDeeper(level) { document.querySelectorAll('.menu-panel:not(#menu-results)').forEach((p) => { if (+p.dataset.level > level) p.remove(); }); }
  async function runItem(it) {
    try { await it.run(); } catch (err) { logError(err); desk.error(`Couldn't do "${it.label}"`, err.message); }
  }

  // Keyboard inside open menus.
  document.addEventListener('keydown', async (e) => {
    if (!open) return;
    const panels = [...document.querySelectorAll('.menu-panel:not(#menu-results)')];
    const panel = panels[panels.length - 1];
    const rows = [...panel.querySelectorAll('.menu-item:not(.disabled)')];
    let hot = panel.querySelector('.menu-item.hot');
    const move = (d) => { const i = rows.indexOf(hot); const n = rows[(i + d + rows.length) % rows.length]; if (hot) hot.classList.remove('hot'); n.classList.add('hot'); n.scrollIntoView({ block: 'nearest' }); };
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); if (panels.length > 1) { panel.remove(); } else close(); return; }
    if (e.key === 'ArrowDown') { e.preventDefault(); move(1); return; }
    if (e.key === 'ArrowUp') { e.preventDefault(); move(-1); return; }
    if (e.key === 'ArrowRight') {
      e.preventDefault();
      if (hot && hot._openSub) { const sp = await hot._openSub(); sp?.querySelector('.menu-item:not(.disabled)')?.classList.add('hot'); }
      else show((open.index + 1) % names.length);
      return;
    }
    if (e.key === 'ArrowLeft') { e.preventDefault(); if (panels.length > 1) panel.remove(); else show((open.index - 1 + names.length) % names.length); return; }
    if (e.key === 'Enter' && hot) { e.preventDefault(); hot.click(); }
  }, true);
  document.addEventListener('mousedown', (e) => { if (open && !e.target.closest('.menu-panel,.menu-name,[data-menu-anchor]')) close(); });
  window.addEventListener('blur', close);
  window.addEventListener('resize', close);
  // Alt+letter opens a menu, as in Google Docs.
  document.addEventListener('keydown', (e) => {
    if (!e.altKey || e.ctrlKey || e.metaKey || e.shiftKey || e.key.length !== 1) return;
    const i = names.findIndex((n) => n[0].toLowerCase() === e.key.toLowerCase());
    if (i < 0) return;
    if (active && active.sheetView && active.sheetView.editing) return;
    e.preventDefault();
    show(i);
  });

  // ---------------------------------------------------------------- search the menus

  async function allItems() {
    const out = [];
    const walk = async (items, path) => {
      for (const it of items) {
        if (it.sep) continue;
        if (it.sub) { await walk(typeof it.sub === 'function' ? await it.sub() : it.sub, [...path, it.label]); continue; }
        if (!it.run) continue;
        if (it.enabled && !(await it.enabled())) continue;
        out.push({ ...it, path: [...path, it.label] });
      }
    };
    for (const m of await build()) await walk(m.items, [m.name]);
    return out;
  }
  function focusSearch() { const s = $('#menu-search'); s.focus(); s.select(); }
  async function search(q) {
    const res = $('#menu-results');
    q = q.trim().toLowerCase();
    if (!q) { res.hidden = true; return; }
    const words = q.split(/\s+/);
    const docs = typeof Shell !== 'undefined' ? await Shell.searchDocs(q) : [];
    const hits = (await allItems()).map((it) => {
      const text = it.path.join(' ').toLowerCase();
      const label = it.label.toLowerCase();
      if (!words.every((w) => text.includes(w))) return null;
      return { it, score: (label.startsWith(q) ? 0 : label.includes(q) ? 1 : 2) + it.path.length * 0.01 };
    }).filter(Boolean).sort((a, b) => a.score - b.score).slice(0, docs.length ? 8 : 14);
    const rows = [];
    const pick = (fn) => (e) => { e.preventDefault(); res.hidden = true; $('#menu-search').value = ''; $('#menu-search').blur(); fn(); };
    if (docs.length) {
      rows.push(h('div', { class: 'menu-head', text: 'Documents' }));
      for (const d of docs.slice(0, 8)) {
        const row = h('div', { class: 'menu-item doc-hit' }, h('span', { class: `tab-kind k-${FORMATS.kindOf(d.name)}`, text: (FORMATS.extOf(d.name) || '?').slice(0, 4).toUpperCase() }), h('span', { class: 'menu-label', text: d.name }), h('span', { class: 'menu-path', text: d.where }));
        row.addEventListener('mousedown', pick(() => (d.tab ? activate(d.tab) : openPath(d.path))));
        rows.push(row);
      }
    }
    if (hits.length) {
      rows.push(h('div', { class: 'menu-head', text: 'Commands' }));
      for (const x of hits) {
        const row = h('div', { class: 'menu-item' }, h('span', { class: 'menu-label', text: x.it.label }), h('span', { class: 'menu-path', text: x.it.path.slice(0, -1).join(' › ') }), h('span', { class: 'menu-key', text: x.it.key || '' }));
        row.addEventListener('mousedown', pick(() => runItem(x.it)));
        rows.push(row);
      }
    }
    rows.find((x) => x.classList.contains('menu-item'))?.classList.add('hot');
    res.replaceChildren(...(rows.length ? rows : [h('div', { class: 'menu-empty', text: 'No matching documents or commands' })]));
    const r = ($('#app-search') || $('#menu-search')).getBoundingClientRect();
    res.style.left = `${r.left}px`;
    res.style.width = `${r.width}px`;
    res.style.top = `${r.bottom + 6}px`;
    res.hidden = false;
  }

  // ---------------------------------------------------------------- compact mode

  let compact = false;
  function toggleCompact() { compact = !compact; document.body.classList.toggle('compact', compact); if (typeof Shell !== 'undefined') Shell.paintSidebarBtn(); $('#menubar').hidden = compact || !(typeof Shell !== 'undefined' && Shell.showMenubar()); flash(compact ? 'Compact mode - Ctrl+Shift+F shows the sidebar and menus again' : 'Sidebar and menus shown'); }

  function wire() {
    const s = $('#menu-search');
    s.addEventListener('input', () => search(s.value));
    s.addEventListener('focus', () => { if (s.value) search(s.value); });
    s.addEventListener('blur', () => setTimeout(() => { $('#menu-results').hidden = true; }, 150));
    s.addEventListener('keydown', (e) => {
      const res = $('#menu-results');
      const rows = [...res.querySelectorAll('.menu-item')];
      const hot = res.querySelector('.hot');
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        if (!rows.length) return;
        const i = rows.indexOf(hot);
        const n = rows[(i + (e.key === 'ArrowDown' ? 1 : -1) + rows.length) % rows.length];
        hot?.classList.remove('hot'); n.classList.add('hot');
      }
      if (e.key === 'Enter' && hot) { e.preventDefault(); hot.dispatchEvent(new MouseEvent('mousedown')); }
      if (e.key === 'Escape') { s.value = ''; res.hidden = true; s.blur(); }
    });
    $('#doc-title').addEventListener('click', () => renameTab(active));
  }
  wire();

  return { render, close, focusSearch, toggleCompact, build, allItems, show, search, popup, tidy, get compact() { return compact; } };
})();
