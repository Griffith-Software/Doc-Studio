'use strict';
// The Doc Studio window: app bar (logo, search, light/dark, settings), action bar (Open, Save, Print, Edit, Annotate,
// Convert, More), the sidebar (Home, Documents, Recent, Starred, Cloud Storage and the tools), the page bar
// (page number, zoom, layout, full screen), page thumbnails on the right, and the Home / file browser pages.

const Shell = (() => {
  const I = (d, size) => `<svg viewBox="0 0 24 24" width="${size || 20}" height="${size || 20}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
  const ICON = {
    open: I('<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v1H3z" fill="currentColor" stroke="none"/><path d="M3 10h18l-2 9H5z" fill="currentColor" stroke="none" opacity=".85"/>', 22),
    caret: I('<path d="M6 9l6 6 6-6"/>', 18),
    save: I('<path d="M5 3h11l3 3v13a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z"/><path d="M7 3v6h9V3"/><rect x="7" y="13" width="10" height="6" rx="1"/>', 24),
    print: I('<path d="M6 9V3h12v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>', 24),
    edit: I('<path d="M4 5h11M9.5 5v14"/><path d="M17 5v14M15 5h4M15 19h4"/>', 24),
    annotate: I('<path d="M4 4h16v12H9l-5 4z"/><path d="M8 9h8M8 12h5"/>', 24),
    convert: I('<path d="M4 8h14l-3-3M20 16H6l3 3"/>', 24),
    more: I('<rect x="4" y="4" width="6" height="6" rx="1"/><rect x="14" y="4" width="6" height="6" rx="1"/><rect x="4" y="14" width="6" height="6" rx="1"/><rect x="14" y="14" width="6" height="6" rx="1"/>', 24),
    home: I('<path d="M3 11l9-8 9 8v9a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1z"/>', 24),
    homeFill: I('<path d="M3 11l9-8 9 8v9a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1z" fill="currentColor"/>', 24),
    folder: I('<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M3 10h18"/>', 24),
    clock: I('<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>', 24),
    star: I('<path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1-4.4-4.3 6.1-.9z"/>', 24),
    starFill: I('<path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1-4.4-4.3 6.1-.9z" fill="currentColor"/>', 24),
    cloud: I('<path d="M7 18a4 4 0 0 1-.5-8A6 6 0 0 1 18 9a4.5 4.5 0 0 1-.5 9z"/>', 24),
    ocr: I('<path d="M4 8V5a1 1 0 0 1 1-1h3M16 4h3a1 1 0 0 1 1 1v3M20 16v3a1 1 0 0 1-1 1h-3M8 20H5a1 1 0 0 1-1-1v-3"/><text x="12" y="15" font-size="7.5" font-weight="700" text-anchor="middle" fill="currentColor" stroke="none" font-family="Segoe UI, sans-serif">OCR</text>', 24),
    compare: I('<rect x="3" y="3" width="11" height="14" rx="1.5"/><path d="M10 7h9a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1z"/>', 24),
    merge: I('<rect x="3" y="3" width="18" height="18" rx="3"/><path d="M12 8v8M8 12h8"/>', 24),
    split: I('<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M8.2 7.8L20 19M8.2 16.2L20 5"/>', 24),
    scan: I('<path d="M3 15h18v4a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z"/><path d="M4 15 18.5 5.5"/><path d="M7 18h3"/>', 24),
    phone: I('<rect x="7" y="2.5" width="10" height="19" rx="2"/><path d="M11 18.5h2"/>', 24),
    compress: I('<path d="M12 3v12M7 10l5 5 5-5"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>', 24),
    panel: I('<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M9 4v16"/><path d="M5.5 8h1.5M5.5 11h1.5"/>', 22),
    gear: I('<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.6 1.6 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.6 1.6 0 0 0-1.8-.3 1.6 1.6 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.6 1.6 0 0 0-1-1.5 1.6 1.6 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.6 1.6 0 0 0 .3-1.8 1.6 1.6 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.6 1.6 0 0 0 1.5-1 1.6 1.6 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.6 1.6 0 0 0 1.8.3H9a1.6 1.6 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.6 1.6 0 0 0 1 1.5 1.6 1.6 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.6 1.6 0 0 0-.3 1.8V9a1.6 1.6 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.6 1.6 0 0 0-1.5 1z"/>', 24),
    sun: I('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>', 22),
    moon: I('<path d="M21 12.8A9 9 0 1 1 11.2 3 7 7 0 0 0 21 12.8z"/>', 22),
    prev: I('<path d="M15 18l-6-6 6-6"/>', 20),
    next: I('<path d="M9 18l6-6-6-6"/>', 20),
    minus: I('<path d="M5 12h14"/>', 22),
    plus: I('<path d="M12 5v14M5 12h14"/>', 22),
    single: I('<rect x="6" y="3" width="12" height="18" rx="1.5"/><path d="M9 8h6M9 12h6M9 16h4"/>', 22),
    two: I('<rect x="3" y="4" width="8" height="16" rx="1"/><rect x="13" y="4" width="8" height="16" rx="1"/>', 22),
    thumbs: I('<rect x="3" y="4" width="7" height="7" rx="1"/><rect x="3" y="13" width="7" height="7" rx="1"/><path d="M13 6h8M13 9h6M13 15h8M13 18h6"/>', 22),
    full: I('<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>', 22),
    up: I('<path d="M12 19V5M5 12l7-7 7 7"/>', 18),
    file: I('<path d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z"/><path d="M14 3v5h5"/>', 20),
  };

  let view = 'doc';          // 'doc' (documents open in tabs) or a page: home, documents, recent, starred, cloud
  let browseDir = null;      // folder shown on the Documents / Cloud pages
  let places = null;
  let starred = new Set();
  const pref = (k, d) => { try { const v = localStorage.getItem(`docstudio.${k}`); return v == null ? d : JSON.parse(v); } catch { return d; } };
  const setPref = (k, v) => { try { localStorage.setItem(`docstudio.${k}`, JSON.stringify(v)); } catch { /* not kept */ } };
  const lower = (p) => String(p || '').toLowerCase();
  const fileName = (p) => p.split(/[\\/]/).pop();
  const dirName = (p) => p.slice(0, Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/')));
  const run = (label, fn) => async () => { try { await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } };

  // ---------------------------------------------------------------- building the chrome

  function setup() {
    const ab = (id, icon, label, caret) => { const b = $(id); b.innerHTML = `${icon}<span>${label}</span>${caret ? ICON.caret : ''}`; return b; };
    $('#btn-open').innerHTML = `${ICON.open}<span>Open</span>`;
    $('#btn-open-more').innerHTML = ICON.caret;
    ab('#ab-save', ICON.save, 'Save');
    ab('#ab-print', ICON.print, 'Print');
    ab('#ab-edit', ICON.edit, 'Edit');
    ab('#ab-annotate', ICON.annotate, 'Annotate');
    ab('#ab-convert', ICON.convert, 'Convert', true);
    ab('#ab-more', ICON.more, 'More', true);
    $('#btn-settings').innerHTML = ICON.gear;
    $('#btn-sidebar').innerHTML = ICON.panel;
    $('#btn-sidebar').addEventListener('click', () => toggleSidebar());
    document.body.classList.toggle('no-sidebar', !pref('sidebar', true));
    paintSidebarBtn();
    $('#pg-prev').innerHTML = ICON.prev; $('#pg-next').innerHTML = ICON.next;
    $('#zm-out').innerHTML = ICON.minus; $('#zm-in').innerHTML = ICON.plus;
    $('#lay-single').innerHTML = ICON.single; $('#lay-two').innerHTML = ICON.two; $('#lay-thumbs').innerHTML = ICON.thumbs; $('#db-full').innerHTML = ICON.full;
    $('#zm-sel').replaceChildren(...ZOOMS.map((z) => h('option', { value: z, text: `${Math.round(z * 100)}%` })));

    const place = (id, icon, label, onIcon) => {
      const b = h('button', { class: 'side-item', type: 'button', 'data-page': id, title: label });
      b.innerHTML = `<span class="side-icon">${icon}</span><span class="side-label">${label}</span>`;
      b.dataset.icon = icon; b.dataset.onIcon = onIcon || icon;
      b.addEventListener('click', () => go(id));
      return b;
    };
    $('#side-places').replaceChildren(
      place('home', ICON.home, 'Home', ICON.homeFill), place('documents', ICON.folder, 'Documents'), place('recent', ICON.clock, 'Recent'),
      place('starred', ICON.star, 'Starred', ICON.starFill), place('cloud', ICON.cloud, 'Cloud Storage'));
    const tool = (id, icon, label, title) => {
      const b = h('button', { class: 'side-item', type: 'button', 'data-tool': id, title });
      b.innerHTML = `<span class="side-icon">${icon}</span><span class="side-label">${label}</span>`;
      b.addEventListener('click', run(label.toLowerCase(), () => Tools[id]()));
      return b;
    };
    $('#side-tools').replaceChildren(
      tool('scan', ICON.scan, 'Scan', 'Scan pages from a connected scanner'),
      tool('phone', ICON.phone, 'Phone', 'Send photos and files between your phone and this PC'),
      tool('convert', ICON.convert, 'Convert', 'Convert this document or a batch of files'),
      tool('ocr', ICON.ocr, 'OCR (Scan to Text)', 'Recognise the text in a scan, photo or PDF'),
      tool('compare', ICON.compare, 'Compare', 'Compare two documents'),
      tool('merge', ICON.merge, 'Merge', 'Merge PDFs, pictures and documents into one PDF'),
      tool('split', ICON.split, 'Split', 'Split a PDF into several files'),
      tool('compress', ICON.compress, 'Compress', 'Make a PDF, picture or Office file smaller'));

    // Action bar
    $('#btn-open').addEventListener('click', () => openDialog());
    $('#btn-open-more').addEventListener('click', run('open the menu', () => Menus.popup($('#btn-open-more'), openMenu())));
    $('#ab-save').addEventListener('click', run('save', () => active && save(active, false)));
    $('#ab-print').addEventListener('click', run('print', () => doPrint()));
    $('#ab-edit').addEventListener('click', run('switch to editing', () => toggleEdit()));
    $('#ab-annotate').addEventListener('click', run('annotate', () => annotate()));
    $('#ab-convert').addEventListener('click', run('open Convert', async () => Menus.popup($('#ab-convert'), await convertMenu())));
    $('#ab-more').addEventListener('click', run('open More', async () => Menus.popup($('#ab-more'), await moreMenu())));
    $('#btn-settings').addEventListener('click', run('open settings', () => settings()));
    $('#btn-theme').addEventListener('click', run('change the theme', () => setTheme(isDark() ? 'light' : 'dark')));
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', paintTheme);
    paintTheme();

    // Page bar
    $('#pg-prev').addEventListener('click', () => step(-1));
    $('#pg-next').addEventListener('click', () => step(1));
    $('#pg-num').addEventListener('keydown', (e) => { if (e.key === 'Enter') { const p = pagerFor(active); if (p) p.go(Math.max(0, Math.min(p.count() - 1, (+e.target.value || 1) - 1))); e.target.blur(); } });
    $('#pg-num').addEventListener('blur', () => updatePager());
    $('#zm-out').addEventListener('click', () => zoomTab(active, -1));
    $('#zm-in').addEventListener('click', () => zoomTab(active, 1));
    $('#zm-sel').addEventListener('change', (e) => zoomTab(active, 0, +e.target.value));
    $('#lay-single').addEventListener('click', () => setLayout('single'));
    $('#lay-two').addEventListener('click', () => setLayout('two'));
    $('#lay-thumbs').addEventListener('click', () => { setPref('thumbs', !pref('thumbs', true)); refresh(); });
    $('#db-full').addEventListener('click', () => desk.fullscreen());
    $('#db-star').addEventListener('click', run('star the document', async () => { if (active && active.path) { await toggleStar(active.path); refresh(); } }));

    // Narrow windows: the sidebar shows icons only.
    const fit = () => document.body.classList.toggle('narrow', innerWidth < 1100);
    addEventListener('resize', fit); fit();
    document.body.classList.toggle('no-status', !pref('statusBar', true));
    document.addEventListener('keydown', (e) => {
      if (e.key === 'F11') { e.preventDefault(); desk.fullscreen(); }
      if (e.ctrlKey && e.key === 'PageDown' && pagerFor(active)) { e.preventDefault(); step(1); }
      if (e.ctrlKey && e.key === 'PageUp' && pagerFor(active)) { e.preventDefault(); step(-1); }
    });
    loadStarred();
    desk.docIndex().then((f) => { indexCache = f; indexAt = Date.now(); }).catch(() => {});
  }

  // The sidebar can be hidden with the button at the top left (remembered), or with compact mode (Ctrl+Shift+F).
  const sidebarShown = () => !document.body.classList.contains('no-sidebar') && !document.body.classList.contains('compact');
  function toggleSidebar() {
    if (document.body.classList.contains('compact')) { Menus.toggleCompact(); document.body.classList.remove('no-sidebar'); setPref('sidebar', true); paintSidebarBtn(); return; }
    const hide = !document.body.classList.contains('no-sidebar');
    document.body.classList.toggle('no-sidebar', hide);
    setPref('sidebar', !hide);
    paintSidebarBtn();
  }
  function paintSidebarBtn() { const b = $('#btn-sidebar'); if (!b) return; const shown = sidebarShown(); b.title = shown ? 'Hide the sidebar' : 'Show the sidebar'; b.classList.toggle('off', !shown); b.setAttribute('aria-pressed', String(shown)); }
  async function loadStarred() { try { starred = new Set((await desk.starred()).map(lower)); } catch { starred = new Set(); } refresh(); }
  async function toggleStar(p) { const on = await desk.toggleStar(p); if (on) starred.add(lower(p)); else starred.delete(lower(p)); flash(on ? `Starred ${fileName(p)}` : `Removed ${fileName(p)} from Starred`); return on; }
  const isStarred = (p) => !!p && starred.has(lower(p));

  // ---------------------------------------------------------------- theme

  const isDark = () => matchMedia('(prefers-color-scheme: dark)').matches;
  async function setTheme(t) { await desk.theme(t); setTimeout(paintTheme, 50); }
  function paintTheme() { const b = $('#btn-theme'); b.innerHTML = isDark() ? ICON.sun : ICON.moon; b.title = isDark() ? 'Switch to light mode' : 'Switch to dark mode'; }

  // ---------------------------------------------------------------- views

  function go(page) {
    // Home is the workspace: from a document it shows the Home page, and from Home it goes back to the document.
    if (page === 'home') view = view === 'doc' && active ? 'home' : 'doc';
    else view = page;
    if (page === 'documents' && (!browseDir || browseKind !== 'documents')) { browseDir = null; browseKind = 'documents'; }
    if (page === 'cloud' && browseKind !== 'cloud') { browseDir = null; browseKind = 'cloud'; }
    Menus.close();
    refreshChrome();
  }
  let browseKind = null;
  function showDoc() { if (view !== 'doc') { view = 'doc'; } }
  const pageShown = () => view !== 'doc' || !active;

  function refresh() {
    const page = pageShown() ? (view === 'doc' ? 'home' : view) : null;
    document.body.classList.toggle('on-page', !!page);
    $('#welcome').hidden = page !== 'home';
    $('#browser').hidden = !page || page === 'home';
    if (page === 'home') renderHome();
    else if (page) renderBrowser(page);
    $('#tabs').classList.toggle('dim', !!page);
    document.querySelectorAll('#side-places .side-item').forEach((b) => {
      const on = b.dataset.page === (page || 'home');
      b.classList.toggle('on', on);
      b.querySelector('.side-icon').innerHTML = on ? b.dataset.onIcon : b.dataset.icon;
    });
    const doc = !page && !!active;
    $('#docbar').hidden = !doc;
    if (!doc) { $('#toolbar').hidden = true; $('#thumbs').hidden = true; }
    // Action bar
    const has = (x) => !!active && doc && x;
    $('#ab-save').disabled = !has(true);
    $('#ab-print').disabled = !has(true);
    $('#ab-edit').disabled = !has(active && active.modes.includes('edit'));
    $('#ab-edit').classList.toggle('on', !!(doc && active && ['edit', 'split'].includes(active.mode)));
    $('#ab-annotate').disabled = !has(active && ['pdf', 'docx', 'pptx', 'sheet', 'odt', 'rtf'].includes(active.kind));
    $('#ab-convert').disabled = false;
    for (const id of ['#db-split', '#db-popout']) { const b = document.querySelector(id); if (b) b.disabled = !has(true); }
    if (doc) updateDocbar();
  }
  const showMenubar = () => !!active && !pageShown() && (['edit', 'split'].includes(active.mode) || active.kind === 'sheet');

  // ---------------------------------------------------------------- action bar menus

  const NEW_TYPES = [['Markdown', 'md', 'Ctrl+Shift+N'], null, ['Document (.docx)', 'docx', 'Ctrl+N'], ['Spreadsheet (.xlsx)', 'xlsx'], ['Presentation (.pptx)', 'pptx'], ['PDF', 'pdf'], null,
    ['OpenDocument text', 'odt'], ['OpenDocument spreadsheet', 'ods'], ['OpenDocument presentation', 'odp'], ['Rich Text', 'rtf'], null,
    ['Plain text', 'txt'], ['Web page (HTML)', 'html'], ['CSV table', 'csv'], ['JSON', 'json']];
  const newItems = () => NEW_TYPES.map((x) => (x ? { label: x[0], key: x[2] || '', run: () => newDocument(x[1]) } : { sep: true }));
  function openMenu() {
    return [
      { label: 'Open…', key: 'Ctrl+O', run: () => openDialog() },
      { label: 'Open recent', sub: async () => { const r = await desk.recent(); return r.length ? r.map((f) => ({ label: f, run: () => openPath(f) })) : [{ label: 'No recent files', enabled: () => false }]; } },
      { label: 'New', sub: newItems() },
      { sep: true },
      ...newItems().filter((x) => !x.sep).slice(0, 5),
    ];
  }
  async function convertMenu() {
    const targets = active && !pageShown() ? await Convert.targets(active) : [];
    return [
      ...(active && !pageShown() ? [{ label: 'Convert to…', run: () => Convert.dialog(active) }, { sep: true }, ...targets.map((t) => ({ label: `${t.label} (.${t.ext})`, run: () => Convert.convertTab(active, t.ext) })), { sep: true }] : []),
      { label: 'Convert files…', run: () => Convert.batch() },
      { label: 'Export as PDF…', key: 'Ctrl+Shift+E', run: () => doExportPdf(), enabled: () => !!active && !pageShown() },
    ];
  }
  async function moreMenu() {
    const menus = await Menus.build();
    return [
      ...menus.map((m) => ({ label: m.name, sub: m.items })),
      { sep: true },
      { label: 'Theme', sub: ['system', 'light', 'dark'].map((t) => ({ label: t[0].toUpperCase() + t.slice(1), run: () => setTheme(t), checked: async () => (await desk.theme()) === t })) },
      { label: sidebarShown() ? 'Hide the sidebar' : 'Show the sidebar', run: () => toggleSidebar() },
      { label: Menus.compact ? 'Leave compact mode (show the sidebar and menus)' : 'Compact mode (hide the sidebar and menus)', key: 'Ctrl+Shift+F', run: () => Menus.toggleCompact() },
      { label: 'Settings…', run: () => settings() },
      { label: 'Default apps…', run: () => defaultApps() },
      { label: 'AI assistants…', run: () => aiAssistants() },
    ];
  }
  async function toggleEdit() {
    if (!active) return;
    showDoc();
    const editing = ['edit', 'split'].includes(active.mode);
    const target = editing ? (active.modes.includes('view') ? 'view' : active.modes.find((m) => m !== active.mode)) : (active.modes.includes('edit') ? 'edit' : null);
    if (target) await setMode(active, target);
  }
  // Annotate: the commenting and markup commands of the current document.
  async function annotate() {
    if (!active) return;
    if (active.kind === 'pdf') {
      if (active.mode !== 'edit') await setMode(active, 'edit');
      active.pdfEditor.setTool('highlight');
      if (active.setPdfTool) active.setPdfTool('highlight');
      flash('Annotate: select text to highlight, or pick another markup tool on the toolbar');
      return;
    }
    if (active.modes.includes('edit') && active.mode !== 'edit') await setMode(active, 'edit');
    const all = await Menus.allItems();
    // Commenting and markup commands only: not files (Open recent), new documents or window commands.
    const found = all.filter((it) => /comment|note|highlight|annotat|suggest|review/i.test(it.label) && !/footnote|endnote|speaker|notes? folder|markdown note/i.test(it.label)
      && !['File', 'Help'].includes(it.path[0]) && !it.path.includes('Split view') && !/[\\/]/.test(it.label));
    if (!found.length) { flash('Nothing to annotate in this kind of document.'); return; }
    Menus.popup($('#ab-annotate'), found.slice(0, 14).map((it) => ({ ...it, label: it.path.length > 2 ? `${it.label}` : it.label })));
  }

  // ---------------------------------------------------------------- page bar, pages and thumbnails

  // A pager knows the pages of the current view: count, current, go(i), the scrolling element and the pages.
  function pagerFor(tab) {
    if (!tab || pageShown()) return null;
    const dom = (scroller, sel, flow) => {
      if (!scroller) return null;
      const pages = () => [...scroller.querySelectorAll(sel)];
      return {
        scroller, pages, flow: flow && flow(),
        count: () => pages().length,
        current: () => currentOf(scroller, pages()),
        go: (i) => { const p = pages()[i]; if (p) scroller.scrollTop += p.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 12; },
      };
    };
    if (tab.kind === 'pdf' && tab.pdfEditor) {
      const ed = tab.pdfEditor;
      return { scroller: ed.main, builtInThumbs: true, flow: ed.main, pages: () => ed.pageEls, count: () => ed.pages.length, current: () => ed.current || 0, go: (i) => ed.pageEls[i] && ed.pageEls[i].scrollIntoView({ block: 'start' }) };
    }
    if (tab.kind === 'docx' && tab.mode === 'view' && tab.viewBox) return dom(tab.viewBox, 'section.docx', () => tab.viewBox.querySelector('.docx-wrapper'));
    if (tab.kind === 'pptx' && tab.mode === 'view' && tab.viewBox) return dom(tab.viewBox, '.pp-slide', () => tab.viewBox);
    if (tab.kind === 'pptx' && tab.mode === 'edit' && tab.pptEditor) {
      const ed = tab.pptEditor;
      return { builtInThumbs: true, noLayout: true, unit: 'slide', count: () => ed.deck.slides.length, current: () => ed.current, go: (i) => ed.go(i) };
    }
    if (tab.kind === 'sheet' && tab.sheetView) {
      const v = tab.sheetView;
      const vis = () => v.wb.sheets.map((s, i) => [s, i]).filter(([s]) => s.state === 'visible');
      return { builtInThumbs: true, noLayout: true, unit: 'sheet', count: () => vis().length, current: () => Math.max(0, vis().findIndex(([, i]) => i === v.sheetIndex)), go: (i) => { const x = vis()[i]; if (x) { v.showSheet(x[1]); setTimeout(updatePager, 0); } } };
    }
    return null;
  }
  function currentOf(scroller, pages) {
    const top = scroller.getBoundingClientRect().top + Math.min(160, scroller.clientHeight / 3);
    let cur = 0;
    pages.forEach((p, i) => { if (p.getBoundingClientRect().top <= top) cur = i; });
    return cur;
  }
  function step(d) { const p = pagerFor(active); if (p) p.go(Math.max(0, Math.min(p.count() - 1, p.current() + d))); setTimeout(updatePager, 60); }

  const watched = new WeakSet();
  let raf = 0;
  function updatePager() {
    const p = pagerFor(active);
    $('#pager').hidden = !p;
    if (!p) return;
    if (p.scroller && !watched.has(p.scroller)) {
      watched.add(p.scroller);
      p.scroller.addEventListener('scroll', () => { if (!raf) raf = requestAnimationFrame(() => { raf = 0; updatePager(); markThumb(); }); }, { passive: true });
    }
    const n = p.count(), c = Math.min(n - 1, p.current());
    if (document.activeElement !== $('#pg-num')) $('#pg-num').value = n ? String(c + 1) : '0';
    $('#pg-count').textContent = `/ ${n}`;
    $('#pg-num').title = p.unit === 'sheet' ? 'Sheet' : p.unit === 'slide' ? 'Slide' : 'Page';
    $('#pg-prev').disabled = c <= 0; $('#pg-next').disabled = c >= n - 1;
    $('#pg-prev').title = `Previous ${p.unit || 'page'} (Ctrl+Page Up)`; $('#pg-next').title = `Next ${p.unit || 'page'} (Ctrl+Page Down)`;
  }
  function setLayout(l) { if (!active) return; active.layout = l; setPref('layout', l); refresh(); }

  function updateDocbar() {
    const tab = active;
    updatePager();
    const z = canZoom(tab);
    $('#zoombox').hidden = !z;
    if (z) { const v = String(tab.zoom || 1); const sel = $('#zm-sel'); if (![...sel.options].some((o) => o.value === v)) sel.append(h('option', { value: v, text: `${Math.round(+v * 100)}%` })); sel.value = v; }
    const p = pagerFor(tab);
    const layout = tab.layout || pref('layout', 'single');
    const thumbsOn = pref('thumbs', true);
    // Layout and thumbnail buttons only where pages can be laid out (not sheets, or slides while editing).
    $('#layout-btns').hidden = !p || !!p.noLayout;
    $('#lay-single').classList.toggle('on', layout === 'single');
    $('#lay-two').classList.toggle('on', layout === 'two');
    $('#lay-two').disabled = $('#lay-single').disabled = !!(p && p.noLayout);
    $('#lay-thumbs').classList.toggle('on', thumbsOn);
    tab.el.classList.toggle('layout-two', !!p && !p.noLayout && layout === 'two');
    // PDF pages are drawn at a size: refit them when the layout changes.
    if (p && tab.kind === 'pdf' && tab.pdfEditor && tab.pdfFit !== layout) { tab.pdfFit = layout; requestAnimationFrame(() => { tab.pdfEditor.fitWidth(layout === 'two' ? 2 : 1); tab.zoom = 1; updatePager(); }); }
    if (p && p.flow) p.flow.classList.add('pager-flow');
    tab.el.classList.toggle('no-thumbs', !thumbsOn);
    $('#db-star').hidden = !tab.path;
    $('#db-star').innerHTML = isStarred(tab.path) ? ICON.starFill : ICON.star;
    $('#db-star').classList.toggle('starred', isStarred(tab.path));
    $('#db-star').title = isStarred(tab.path) ? 'Remove from Starred' : 'Star this document';
    // Thumbnails: our panel for pages drawn as HTML; PDFs, slides and sheets have their own.
    const own = !!p && !p.builtInThumbs && thumbsOn;
    $('#thumbs').hidden = !own;
    if (own) buildThumbs(tab, p);
  }

  // Thumbnails are scaled copies of the pages, made when they scroll into view.
  let thumbFor = null, thumbIO = null;
  function buildThumbs(tab, p) {
    const pages = p.pages();
    const sig = `${tab.id}|${tab.mode}|${pages.length}|${tab.viewStamp || 0}|${pages[0] ? pages[0].offsetWidth : 0}`;
    const box = $('#thumbs');
    if (thumbFor === sig && box.childElementCount) { markThumb(); return; }
    thumbFor = sig;
    if (thumbIO) thumbIO.disconnect();
    thumbIO = new IntersectionObserver((entries) => {
      for (const en of entries) {
        if (!en.isIntersecting || en.target.dataset.done) continue;
        en.target.dataset.done = '1';
        const pg = p.pages()[+en.target.dataset.i];
        if (!pg) continue;
        const w = pg.offsetWidth || 800;
        const scale = 118 / w;
        const clone = pg.cloneNode(true);
        clone.querySelectorAll('video, audio, iframe, [contenteditable]').forEach((x) => (x.tagName === 'VIDEO' || x.tagName === 'AUDIO' || x.tagName === 'IFRAME' ? x.remove() : x.removeAttribute('contenteditable')));
        clone.removeAttribute('id'); clone.querySelectorAll('[id]').forEach((x) => x.removeAttribute('id'));
        Object.assign(clone.style, { transform: `scale(${scale})`, transformOrigin: '0 0', margin: '0', position: 'absolute', left: '0', top: '0', width: `${w}px`, pointerEvents: 'none' });
        const holder = h('div', { class: `thumb-paper ${pg.matches('section.docx') ? 'docx-wrapper' : ''}` });
        holder.style.height = `${Math.round(pg.offsetHeight * scale)}px`;
        holder.append(clone);
        en.target.querySelector('.thumb-slot').replaceChildren(holder);
      }
    }, { root: box, rootMargin: '400px 0px' });
    box.replaceChildren(...pages.slice(0, 400).map((pg, i) => {
      const w = pg.offsetWidth || 800;
      const t = h('button', { class: 'thumb', type: 'button', 'data-i': i, title: `${p.unit === 'slide' || tab.kind === 'pptx' ? 'Slide' : 'Page'} ${i + 1}` },
        h('div', { class: 'thumb-slot', style: `height:${Math.round((pg.offsetHeight || 1000) * 118 / w)}px` }), h('span', { class: 'thumb-no', text: String(i + 1) }));
      t.addEventListener('click', () => { p.go(i); setTimeout(() => { updatePager(); markThumb(); }, 50); });
      thumbIO.observe(t);
      return t;
    }));
    markThumb();
  }
  function markThumb() {
    const p = pagerFor(active);
    const box = $('#thumbs');
    if (!p || box.hidden) return;
    const c = p.current();
    box.querySelectorAll('.thumb').forEach((t) => {
      const on = +t.dataset.i === c;
      if (on && !t.classList.contains('on')) t.scrollIntoView({ block: 'nearest' });
      t.classList.toggle('on', on);
    });
  }
  // Pages drawn again (edits, zoom): rebuild the thumbnails.
  function pagesChanged(tab) { tab.viewStamp = (tab.viewStamp || 0) + 1; if (tab === active) setTimeout(() => refresh(), 30); }

  // ---------------------------------------------------------------- Home

  const kindBadge = (name) => h('span', { class: `tab-kind k-${FORMATS.kindOf(name)}`, text: (FORMATS.extOf(name) || '?').slice(0, 4).toUpperCase() });
  let homeSeq = 0;
  async function renderHome() {
    const seq = ++homeSeq;
    const root = $('#welcome');
    const card = (label, ext, sub) => h('button', { class: 'new-card', type: 'button', onclick: () => newDocument(ext) }, h('span', { class: `new-icon k-${FORMATS.kindOf(`x.${ext}`)}`, text: ext.toUpperCase() }), h('b', { text: label }), h('small', { text: sub }));
    const more = h('button', { class: 'new-card more', type: 'button', 'data-menu-anchor': '' }, h('span', { class: 'new-icon', text: '+' }), h('b', { text: 'More types' }), h('small', { text: 'ODF, RTF, text, HTML, CSV, JSON' }));
    more.addEventListener('click', () => Menus.popup(more, newItems().slice(5)));
    const recent = await desk.recent();
    if (seq !== homeSeq) return;
    const list = (files, empty) => (files.length ? h('div', { class: 'file-list' }, ...files.slice(0, 8).map((f) => fileRow({ path: f, name: fileName(f) }, true))) : h('p', { class: 'muted empty', text: empty }));
    const starredList = [...(await desk.starred())];
    if (seq !== homeSeq) return;
    const toolTile = (id, icon, label, sub) => { const b = h('button', { class: 'tool-tile', type: 'button' }, h('span', { class: 'tile-icon' }), h('b', { text: label }), h('small', { text: sub })); b.firstChild.innerHTML = icon; b.addEventListener('click', run(label.toLowerCase(), () => Tools[id]())); return b; };
    root.replaceChildren(h('div', { class: 'home' },
      h('div', { class: 'home-hero' }, h('img', { src: 'assets/logo.png', alt: '', width: 56, height: 56 }), h('div', {}, h('h1', { text: 'Welcome to Doc Studio' }), h('p', { class: 'lede', text: 'Open, edit and convert documents, spreadsheets, presentations, PDFs and more, keeping their original formatting. Drop a file anywhere to open it.' }))),
      h('h2', { text: 'Start something new' }),
      h('div', { class: 'new-row' }, card('Markdown', 'md', 'AI chats, notes and docs'), card('Document', 'docx', 'Letters, reports, résumés'), card('Spreadsheet', 'xlsx', 'Live formulas and charts'), card('Presentation', 'pptx', 'Slides and slideshows'), card('PDF', 'pdf', 'Forms, notes, signatures'), more),
      h('div', { class: 'home-cols' },
        h('section', {}, h('div', { class: 'home-head' }, h('h2', { text: 'Recent' }), h('button', { class: 'link', text: 'See all', onclick: () => go('recent') })), list(recent, 'Files you open will appear here.')),
        h('section', {}, h('div', { class: 'home-head' }, h('h2', { text: 'Starred' }), h('button', { class: 'link', text: 'See all', onclick: () => go('starred') })), list(starredList, 'Star a document (☆ on the page bar or in a list) to keep it here.'))),
      h('h2', { text: 'Tools' }),
      h('div', { class: 'tool-row' }, toolTile('scan', ICON.scan, 'Scan', 'From a connected scanner'), toolTile('phone', ICON.phone, 'Phone', 'Photos and files over Wi-Fi'), toolTile('convert', ICON.convert, 'Convert', 'Any format to any format'), toolTile('ocr', ICON.ocr, 'OCR', 'Scan or photo to text'), toolTile('merge', ICON.merge, 'Merge', 'Many files into one PDF'), toolTile('split', ICON.split, 'Split', 'One PDF into several'), toolTile('compress', ICON.compress, 'Compress', 'Smaller PDFs and files'), toolTile('compare', ICON.compare, 'Compare', 'See what changed')),
      h('div', { class: 'formats' },
        h('div', {}, h('h3', { text: 'View & edit' }), h('p', { text: 'Markdown (.md, with a live preview), documents (.docx), spreadsheets (.xlsx, with live formulas), presentations (.pptx), PDF, OpenDocument text (.odt) and RTF, keeping their original formatting; text, HTML, CSV, JSON, XML and code files' })),
        h('div', {}, h('h3', { text: 'Through Office or LibreOffice' }), h('p', { text: 'Older and other formats (.doc, .xls, .xlsb, .ods, .ppt, .odp) open in the full editors and save back to their own format.' })),
        h('div', {}, h('h3', { text: 'View' }), h('p', { text: 'EPUB, images, audio and video' })),
        h('div', {}, h('h3', { text: 'Open a file' }), h('p', {}, h('button', { class: 'link', text: 'Browse your documents…', onclick: () => openDialog() }))))));
  }

  // ---------------------------------------------------------------- file lists (Documents, Recent, Starred, Cloud Storage)

  function when(ms) { if (!ms) return ''; const d = new Date(ms); const today = new Date(); return d.toDateString() === today.toDateString() ? d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : d.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' }); }
  function fileRow(f, compact) {
    const star = h('button', { class: `row-star${isStarred(f.path) ? ' on' : ''}`, type: 'button', title: isStarred(f.path) ? 'Remove from Starred' : 'Star' });
    star.innerHTML = isStarred(f.path) ? ICON.starFill : ICON.star;
    star.addEventListener('click', run('star the file', async (e) => { const on = await toggleStar(f.path); star.classList.toggle('on', on); star.innerHTML = on ? ICON.starFill : ICON.star; }));
    star.addEventListener('dblclick', (e) => e.stopPropagation());
    const row = h('div', { class: `file-row${compact ? ' compact' : ''}`, tabindex: '0', title: f.path },
      kindBadge(f.name), h('span', { class: 'file-name', text: f.name }), h('span', { class: 'file-where', text: dirName(f.path) }),
      ...(compact ? [] : [h('span', { class: 'file-when', text: when(f.mtime) }), h('span', { class: 'file-size', text: f.size != null ? formatBytes(f.size) : '' })]), star);
    const open = () => openPath(f.path);
    row.addEventListener('click', (e) => { if (!e.target.closest('.row-star')) open(); });
    row.addEventListener('keydown', (e) => { if (e.key === 'Enter') open(); });
    row.addEventListener('contextmenu', (e) => { e.preventDefault(); SheetView.showMenu(e.clientX, e.clientY, [['Open', '', open], ['Show in folder', '', () => desk.reveal(f.path)], [isStarred(f.path) ? 'Remove from Starred' : 'Star', '', () => toggleStar(f.path).then(() => refresh())]]); });
    return row;
  }
  function folderRow(d) {
    const row = h('div', { class: 'file-row folder', tabindex: '0', title: d.path }, h('span', { class: 'folder-icon' }), h('span', { class: 'file-name', text: d.name }), h('span', { class: 'file-where', text: '' }), h('span', { class: 'file-when', text: when(d.mtime) }), h('span', { class: 'file-size', text: '' }), h('span', {}));
    row.firstChild.innerHTML = ICON.folder;
    const open = () => { browseDir = d.path; refresh(); };
    row.addEventListener('click', open);
    row.addEventListener('keydown', (e) => { if (e.key === 'Enter') open(); });
    return row;
  }

  let browserSeq = 0;
  async function renderBrowser(page) {
    const seq = ++browserSeq;
    const root = $('#browser');
    places = places || await desk.places();
    const titles = { documents: 'Documents', recent: 'Recent', starred: 'Starred', cloud: 'Cloud Storage' };
    const filter = h('input', { class: 'browse-filter', type: 'search', placeholder: 'Filter by name', 'aria-label': 'Filter by name' });
    const head = h('div', { class: 'browse-head' }, h('h1', { text: titles[page] }), filter);
    const body = h('div', { class: 'browse-body' });
    root.replaceChildren(h('div', { class: 'browse' }, head, body));
    const applyFilter = () => { const q = filter.value.trim().toLowerCase(); body.querySelectorAll('.file-row:not(.list-head)').forEach((r) => { const n = r.querySelector('.file-name'); r.hidden = !!q && !!n && !n.textContent.toLowerCase().includes(q); }); };
    filter.addEventListener('input', applyFilter);
    const listHead = () => h('div', { class: 'file-row list-head' }, h('span', {}), h('span', { text: 'Name' }), h('span', { text: 'Folder' }), h('span', { text: 'Modified' }), h('span', { text: 'Size' }), h('span', {}));
    const withStats = async (paths) => Promise.all(paths.map(async (p) => { try { const st = await desk.stat(p); return { path: p, name: fileName(p), size: st && st.size, mtime: st && Date.parse(st.modified) }; } catch { return { path: p, name: fileName(p) }; } }));

    if (page === 'recent' || page === 'starred') {
      const files = await withStats(page === 'recent' ? await desk.recent() : await desk.starred());
      if (seq !== browserSeq) return;
      body.append(files.length ? h('div', { class: 'file-list' }, listHead(), ...files.map((f) => fileRow(f))) : h('p', { class: 'muted empty', text: page === 'recent' ? 'Files you open will appear here.' : 'Nothing starred yet. Star a document with ☆ on the page bar or next to any file.' }));
      return;
    }
    if (page === 'cloud' && !browseDir) {
      const clouds = places.cloud;
      body.append(clouds.length
        ? h('div', { class: 'cloud-grid' }, ...clouds.map((c) => { const b = h('button', { class: `cloud-card c-${c.kind}`, type: 'button' }, h('span', { class: 'tile-icon' }), h('b', { text: c.name }), h('small', { text: c.path })); b.firstChild.innerHTML = ICON.cloud; b.addEventListener('click', () => { browseDir = c.path; refresh(); }); return b; }))
        : h('div', { class: 'empty-cloud' }, h('p', { text: 'No cloud folders were found on this PC.' }), h('p', { class: 'muted', text: 'Install the desktop app of OneDrive, Google Drive, Dropbox, iCloud or Box and sign in; its folder appears here and files open and save straight to your cloud.' })));
      return;
    }
    const roots = page === 'cloud' ? places.cloud.map((c) => ({ name: c.name, path: c.path })) : [{ name: 'Documents', path: places.documents }, { name: 'Desktop', path: places.desktop }, { name: 'Downloads', path: places.downloads }, { name: 'Home', path: places.home }];
    const dir = browseDir || roots[0].path;
    // The closest root that contains this folder is the highlighted one.
    const within = roots.filter((r) => lower(dir) === lower(r.path) || lower(dir).startsWith(`${lower(r.path)}\\`)).sort((a, b) => b.path.length - a.path.length);
    const chips = h('div', { class: 'browse-chips' }, ...roots.map((r) => h('button', { class: `chip${within[0] === r ? ' on' : ''}`, type: 'button', text: r.name, onclick: () => { browseDir = r.path; refresh(); } })),
      h('button', { class: 'chip ghosty', type: 'button', text: 'Other folder…', onclick: run('choose a folder', async () => { const [d] = await desk.pick({ folder: true, title: 'Choose a folder' }); if (d) { browseDir = d; refresh(); } }) }));
    // Breadcrumbs from the root that contains this folder.
    const base = within[0] || { name: dir.slice(0, 3), path: dir.slice(0, 3) };
    const crumbs = h('div', { class: 'crumbs' });
    const rel = dir.slice(base.path.length).split(/[\\/]/).filter(Boolean);
    let acc = base.path;
    crumbs.append(h('button', { class: 'crumb', type: 'button', text: base.name, onclick: () => { browseDir = base.path; refresh(); } }));
    for (const part of rel) { acc = `${acc}\\${part}`; const target = acc; crumbs.append(h('span', { class: 'crumb-sep', text: '›' }), h('button', { class: 'crumb', type: 'button', text: part, onclick: () => { browseDir = target; refresh(); } })); }
    if (lower(dir) !== lower(base.path)) { const up = h('button', { class: 'crumb up', type: 'button', title: 'Up one folder', onclick: () => { browseDir = dirName(dir); refresh(); } }); up.innerHTML = ICON.up; crumbs.prepend(up); }
    head.after(chips, crumbs);
    const r = await desk.list(dir);
    if (seq !== browserSeq) return;
    if (r.error) { body.append(h('p', { class: 'muted empty', text: `This folder can't be read: ${r.error}` })); return; }
    const folders = r.items.filter((x) => x.dir).sort((a, b) => a.name.localeCompare(b.name));
    const files = r.items.filter((x) => !x.dir).sort((a, b) => b.mtime - a.mtime);
    body.append(folders.length || files.length ? h('div', { class: 'file-list' }, listHead(), ...folders.map(folderRow), ...files.map((f) => fileRow(f))) : h('p', { class: 'muted empty', text: 'No documents in this folder.' }));
    applyFilter();
  }

  // ---------------------------------------------------------------- search documents (the app bar's search also finds commands)

  let indexCache = null, indexAt = 0;
  async function searchDocs(q) {
    const words = q.toLowerCase().split(/\s+/).filter(Boolean);
    if (!words.length) return [];
    if (!indexCache || Date.now() - indexAt > 60000) { indexAt = Date.now(); desk.docIndex().then((f) => { indexCache = f; }).catch(() => {}); }
    const out = [], seen = new Set();
    const add = (name, path, tab, where) => {
      const key = lower(path || `tab:${tab && tab.id}`);
      if (seen.has(key) || !words.every((w) => name.toLowerCase().includes(w))) return;
      seen.add(key);
      out.push({ name, path, tab, where });
    };
    for (const t of tabs) add(t.name, t.path, t, 'Open');
    for (const f of await desk.recent()) add(fileName(f), f, null, 'Recent');
    for (const f of await desk.starred()) add(fileName(f), f, null, 'Starred');
    for (const f of indexCache || []) { add(fileName(f), f, null, dirName(f)); if (out.length > 40) break; }
    return out;
  }

  // ---------------------------------------------------------------- default apps
  // Windows lets only the owner choose default apps, and only File Explorer's "Open with" offers "Always". So this
  // lists every type Doc Studio reads with the app that opens it now, then, for each type the owner ticked, shows a
  // sample file in File Explorer. Choosing Open with > Doc Studio > Always there opens the sample in Doc Studio,
  // which tells this list that type is done (main.js keeps the sample from opening as a document).

  let pickedHandler = null;
  if (desk.onDefaultPicked) desk.onDefaultPicked((ext) => { if (pickedHandler) pickedHandler(ext); });

  async function defaultApps() {
    const box = h('div', { class: 'dialog defaults-dialog' }, h('h3', { text: 'Default apps' }));
    const back = h('div', { class: 'backdrop' }, box);
    let running = false, stop = false, skip = null;
    const close = () => { if (running) { stop = true; if (skip) skip(); return; } pickedHandler = null; back.remove(); };
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(); } });
    const note = h('div', { class: 'defaults-note', text: 'Checking which app opens each type of file…' });
    const list = h('div', { class: 'defaults-list' });
    const go = h('button', { type: 'button', text: 'Make Doc Studio the default', disabled: true });
    const done = h('button', { type: 'button', class: 'ghost', text: 'Close', onclick: close });
    const next = h('button', { type: 'button', class: 'ghost', text: 'Skip this type', hidden: true, onclick: () => { if (skip) skip(); } });
    const all = h('button', { type: 'button', class: 'ghost defaults-all', text: 'Tick all', disabled: true });
    box.append(note, list, h('div', { class: 'dialog-actions' }, all, h('span', { class: 'spacer' }), next, done, go));
    document.body.append(back);
    done.focus();
    let info;
    try { info = await desk.defaultsStatus(); } catch (err) { note.textContent = `Couldn't ask Windows: ${err.message}`; return; }
    const rows = {};
    const paint = (ext) => {
      const r = rows[ext], s = info.status[ext] || {};
      r.state.textContent = s.mine ? '✓ Doc Studio' :(s.app ? `Opens with ${s.app}` : 'No app chosen');
      r.row.classList.toggle('mine', !!s.mine);
      // Already Doc Studio: nothing to tick.
      r.check.disabled = isSet(s);
      if (isSet(s)) r.check.checked = false;
    };
    // Set: Doc Studio opens the type (each type's icon comes from the File Explorer icon handler, shell-icons.js).
    const isSet = (s) => !!(s && s.mine);
    const ticked = () => Object.keys(rows).filter((x) => rows[x].check.checked && !isSet(info.status[x]));
    const count = () => {
      const n = ticked().length, left = Object.values(info.status).filter((s) => !isSet(s)).length;
      go.disabled = running || !n;
      go.textContent = n ? `Make Doc Studio the default for ${n} type${n === 1 ? '' : 's'}` : 'Make Doc Studio the default';
      // Tick all / Untick all: toggles the types still to set.
      const open = Object.keys(rows).filter((x) => !isSet(info.status[x]));
      const allTicked = open.length && open.every((x) => rows[x].check.checked);
      all.textContent = allTicked ? 'Untick all' : 'Tick all';
      all.disabled = running || !open.length;
      const total = Object.keys(rows).length, opens = Object.values(info.status).filter((s) => s.mine).length;
      if (!running) {
        note.textContent = !left ? 'Doc Studio is the default app for every type it opens, each with its own icon.'
          : `Doc Studio opens ${opens} of ${total} types. Tick the ones you want and press the button. Windows only lets you choose a default yourself, so Doc Studio shows a sample file for each type in File Explorer and tells you what to click.`;
      }
    };
    for (const [group, types] of info.types) {
      list.append(h('div', { class: 'defaults-group', text: group }));
      for (const [ext, label] of types) {
        const check = h('input', { type: 'checkbox' });
        const state = h('span', { class: 'defaults-state' });
        const row = h('label', { class: 'defaults-row' }, check, h('b', { text: `.${ext}` }), h('span', { class: 'defaults-label', text: label }), state);
        rows[ext] = { row, check, state };
        check.checked = !isSet(info.status[ext]);
        check.addEventListener('change', count);
        list.append(row);
        paint(ext);
      }
    }
    // A sample opened in Doc Studio: ask Windows what opens that type now.
    let waiting = null;
    pickedHandler = async (ext) => {
      if (!rows[ext]) return;
      try { info.status[ext] = await desk.checkDefault(ext); } catch { return; }
      paint(ext); count();
      if (waiting) waiting(ext);
    };
    const steps = (ext, i, n, again) => {
      note.replaceChildren(
        h('b', { text: `${i + 1} of ${n}: .${ext} files` }),
        again ? h('p', { class: 'defaults-again', text: `Windows still opens .${ext} files with ${info.status[ext].app || 'another app'}. Make sure you press Always, not Just once.` }) : '',
        h('ol', null,
          h('li', { text: `In the File Explorer window that just opened, right-click "Choose the app for .${ext} files.${ext}".` }),
          h('li', { text: 'Choose Open with, then Choose another app.' }),
          h('li', { text: 'Pick Doc Studio and press Always.' })),
        h('p', { text: 'Doc Studio moves on to the next type by itself.' }));
    };
    all.addEventListener('click', () => {
      const open = Object.keys(rows).filter((x) => !isSet(info.status[x]));
      const tick = !open.every((x) => rows[x].check.checked);
      for (const x of open) rows[x].check.checked = tick;
      count();
    });
    go.addEventListener('click', async () => {
      const todo = ticked();
      running = true; stop = false; count();
      done.textContent = 'Stop';
      next.hidden = false;
      for (const [i, ext] of todo.entries()) {
        if (stop || !back.isConnected) break;
        if (isSet(info.status[ext])) continue;
        rows[ext].row.classList.add('busy');
        rows[ext].row.scrollIntoView({ block: 'nearest' });
        steps(ext, i, todo.length, false);
        try { await desk.revealDefaultSample(ext); } catch (err) { desk.error(`Couldn't show the sample .${ext} file`, err.message); break; }
        // Wait until this type opens with Doc Studio, or the owner skips or stops.
        for (;;) {
          const got = await new Promise((resolve) => { waiting = (x) => { if (x === ext) resolve('picked'); }; skip = () => resolve('skip'); });
          waiting = null; skip = null;
          if (got !== 'picked' || isSet(info.status[ext])) break;
          steps(ext, i, todo.length, true);
        }
        rows[ext].row.classList.remove('busy');
      }
      running = false;
      next.hidden = true;
      done.textContent = 'Close';
      count();
    });
    count();
  }

  // ---------------------------------------------------------------- AI assistants
  // Let Claude (Code and Desktop), Gemini CLI and OpenAI Codex use Doc Studio through its MCP server. The packages
  // ship in resources\ai (tools/make-ai-packages.js); ai/bridge.js answers them and enforces the file rules.

  async function aiAssistants() {
    const info = await desk.aiInfo();
    const box = h('div', { class: 'dialog ai-dialog' }, h('h3', { text: 'AI assistants' }));
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    const on = h('input', { type: 'checkbox' });
    on.checked = info.enabled;
    on.addEventListener('change', async () => { await desk.aiSetEnabled(on.checked); flash(on.checked ? 'AI assistants can use Doc Studio' : 'AI assistants are turned off'); });
    const copy = (text, what) => async () => { await navigator.clipboard.writeText(text); flash(`${what} copied`); };
    const code = (text) => h('pre', { class: 'ai-code', text });
    const found = (yes) => h('span', { class: `ai-found${yes ? ' yes' : ''}`, text: yes ? 'Installed on this PC' : 'Not found on this PC' });
    const card = (title, yes, ...kids) => h('section', { class: 'ai-card' }, h('div', { class: 'ai-card-head' }, h('b', { text: title }), found(yes)), ...kids);
    const q = (p) => `"${p}"`;
    const claudeCmds = `/plugin marketplace add ${q(info.dir)}\n/plugin install doc-studio@doc-studio`;
    const geminiCmd = `gemini extensions install ${q(info.gemini)}`;
    const codexToml = `[mcp_servers.doc_studio]\ncommand = '${info.exe}'\nargs = ['${info.server}']\nenv = { ELECTRON_RUN_AS_NODE = '1' }`;
    box.append(
      h('p', { class: 'muted ai-intro', text: 'AI assistants can read, create, convert, merge, split, compress, OCR, scan and edit pictures with Doc Studio. They only create new files: to change one of your files they copy it first; they may change files an AI made (asking you first for files from an earlier session); they never delete anything. Doc Studio sends nothing off this PC itself, but what an assistant reads goes to that assistant, which works under its own maker\'s terms. Off until you switch it on; while off, nothing listens for assistants.' }),
      h('label', { class: 'check-row ai-switch' }, on, h('span', { text: 'Let AI assistants use Doc Studio' })),
      info.ready ? '' : h('p', { class: 'tool-note tool-problem', text: 'The AI packages are missing from this copy of Doc Studio. Reinstall it.' }),
      card('Claude Code', info.apps.claudeCode, h('p', { class: 'ai-step', text: 'In Claude Code, run these two commands:' }), code(claudeCmds), h('div', { class: 'ai-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Copy the commands', onclick: copy(claudeCmds, 'Commands') }))),
      // Claude Desktop doesn't always register for .mcpb files, so show the file and Claude's own install steps.
      card('Claude Desktop', info.apps.claudeDesktop,
        h('ol', { class: 'ai-steps' },
          h('li', { text: 'Press Show the extension file: File Explorer opens with doc-studio.mcpb selected.' }),
          h('li', { text: 'In Claude Desktop, open Settings, then Extensions.' }),
          h('li', { text: 'Drag doc-studio.mcpb onto that page, or use Install extension (under Advanced settings) and choose it. Then press Install.' })),
        code(info.mcpb),
        h('div', { class: 'ai-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Show the extension file', onclick: async () => { try { await desk.aiInstallClaudeDesktop(); } catch (err) { desk.error("Couldn't show the extension", err.message); } } }))),
      card('Gemini CLI', info.apps.gemini, h('p', { class: 'ai-step', text: 'In a terminal, run:' }), code(geminiCmd), h('div', { class: 'ai-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Copy the command', onclick: copy(geminiCmd, 'Command') }))),
      card('OpenAI Codex', info.apps.codex, h('p', { class: 'ai-step', text: 'Add this to %USERPROFILE%\\.codex\\config.toml (create the file if it is not there), then restart Codex:' }), code(codexToml), h('div', { class: 'ai-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Copy the setup', onclick: copy(codexToml, 'Setup') }))),
      h('p', { class: 'muted ai-foot', text: `ChatGPT's app can only use tools on the internet, so it can't reach Doc Studio on this PC; Codex is OpenAI's way in. ${info.madeFiles ? `AI assistants have made ${info.madeFiles} file${info.madeFiles === 1 ? '' : 's'} with Doc Studio.` : ''}` }),
      h('div', { class: 'dialog-actions' }, h('button', { type: 'button', text: 'Done', onclick: close })));
    document.body.append(back);
    on.focus();
  }

  // ---------------------------------------------------------------- settings

  async function settings() {
    const theme = await desk.theme();
    let name = ''; try { name = localStorage.getItem('docstudio.author') || ''; } catch { /* none */ }
    const asked = DocTools.form('Settings', [
      { name: 'theme', label: 'Appearance', type: 'select', value: theme, options: [['system', 'Same as Windows'], ['light', 'Light'], ['dark', 'Dark']] },
      { name: 'author', label: 'Your name (on comments, notes and suggestions)', value: name },
      { name: 'thumbs', label: 'Show page thumbnails', type: 'checkbox', value: pref('thumbs', true) },
      { name: 'layout', label: 'Page layout', type: 'select', value: pref('layout', 'single'), options: [['single', 'Single page'], ['two', 'Two pages side by side']] },
      { name: 'statusBar', label: 'Show the status bar', type: 'checkbox', value: pref('statusBar', true) },
      { name: 'spell', label: 'Underline spelling mistakes', type: 'checkbox', value: typeof spellingOn === 'function' ? spellingOn() : true },
    ], 'Save');
    const form = [...document.querySelectorAll('.form-dialog')].pop();
    if (form) {
      form.querySelector('.dialog-actions').before(h('div', { class: 'settings-defaults' },
        h('span', null, h('b', { text: 'Default apps' }), h('small', { text: 'Open Markdown, PDF, documents, spreadsheets, presentations and text files with Doc Studio' })),
        h('button', { type: 'button', class: 'ghost', text: 'Make Doc Studio the default…', onclick: () => defaultApps() })));
      form.querySelector('.dialog-actions').before(h('div', { class: 'settings-defaults' },
        h('span', null, h('b', { text: 'AI assistants' }), h('small', { text: 'Let Claude, Gemini or Codex use Doc Studio' })),
        h('button', { type: 'button', class: 'ghost', text: 'Set up…', onclick: () => aiAssistants() })));
    }
    const v = await asked;
    if (!v) return;
    try { localStorage.setItem('docstudio.author', v.author.trim()); } catch { /* not kept */ }
    setPref('thumbs', !!v.thumbs); setPref('layout', v.layout); setPref('statusBar', !!v.statusBar);
    tabs.forEach((t) => { t.layout = v.layout; });
    document.body.classList.toggle('no-status', !v.statusBar);
    if (typeof spellingOn === 'function' && !!v.spell !== spellingOn()) toggleSpelling();
    if (v.theme !== theme) await setTheme(v.theme);
    refresh();
  }

  // Everything that changes which page is shown goes back to the document view.
  const activate0 = activate;
  activate = function (tab) { if (tab) view = 'doc'; return activate0(tab); }; // eslint-disable-line no-global-assign
  setup();

  return { sidebarShown, toggleSidebar, paintSidebarBtn, refresh, showMenubar, searchDocs, go, pagesChanged, pagerFor, updatePager, ICON, isStarred, toggleStar, get view() { return view; } };
})();

refreshChrome();
desk.appVersion().then((v) => { document.querySelector('#status-version').textContent = `Version ${v}`; }).catch(() => {});
