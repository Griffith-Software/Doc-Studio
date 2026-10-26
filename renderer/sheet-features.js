'use strict';
// More spreadsheet features, added to SheetView.View: insert/delete/hide rows and columns, row heights, the
// right-click menu, sort, filter, freeze panes, conditional formatting (drawn and created), and pictures,
// shapes and charts from the sheet's drawing.

(() => {
  const P = SheetView.View.prototype;
  const { key, addr, colName, parseRange, rangeRef, shiftFormula, XErr } = SheetModel;
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const EMU = 9525;

  // ======================================================================
  // Whole-workbook undo for structural changes: snapshot the package before and after.
  // ======================================================================

  P.snapshotOp = async function (label, fn) {
    if (this.editing) this.endEdit(true);
    const before = await this.wb.save();
    try { await fn(); } catch (err) { await this.reloadFrom(before, this.sheetIndex); throw err; }
    await this.wb.loadTables();
    this.calc.build();
    this.calc.recalc([], { all: true, volatile: false });
    const after = await this.wb.save();
    this.pushUndo({ snapshot: { before, after }, sheetIndex: this.sheetIndex, label });
    this.afterStructure();
    this.changed();
  };

  P.afterStructure = function () {
    this.cfInvalidate();
    this.layoutDirty = true;
    this.render(true);
    this.paintTabs();
    this.updateBars();
    this.loadDrawings();
  };

  P.reloadFrom = async function (bytes, sheetIndex) {
    const wb = await new SheetModel.Workbook().load(bytes);
    this.wb = wb;
    this.calc = new SheetCalc.Calc(wb);
    this.drawingCache = new Map();
    if (this.opts.onWorkbook) this.opts.onWorkbook(wb);
    const sel = { ...this.sel }, act = { ...this.active };
    this.showSheet(Math.min(sheetIndex, wb.sheets.length - 1));
    this.sel = sel; this.active = act; this.anchor = act;
    this.afterStructure();
  };

  const restore = P.restore;
  P.restore = function (entry, which) {
    if (entry.snapshot) { this.reloadFrom(entry.snapshot[which], entry.sheetIndex).then(() => this.changed()); return; }
    return restore.call(this, entry, which);
  };

  // ======================================================================
  // Rows and columns
  // ======================================================================

  P.selRows = function () { const g = this.expandMerges(this.sel); return { a: g.r1, n: Math.min(g.r2, SheetModel.MAX_ROW) - g.r1 + 1 }; };
  P.selCols = function () { const g = this.expandMerges(this.sel); return { a: g.c1, n: Math.min(g.c2, SheetModel.MAX_COL) - g.c1 + 1 }; };

  P.insertRows = function (below) {
    const { a, n } = this.selRows();
    if (n > 10000) return this.opts.error && this.opts.error('Select fewer rows to insert.');
    return this.snapshotOp('Insert rows', () => SheetStructure.shift(this.wb, this.sheet, 'row', below ? a + n : a, n));
  };
  P.deleteRows = function () {
    const { a, n } = this.selRows();
    return this.snapshotOp('Delete rows', () => SheetStructure.shift(this.wb, this.sheet, 'row', a, -Math.min(n, Math.max(this.sheet.maxRow, a) - a + 1)));
  };
  P.insertCols = function (right) {
    const { a, n } = this.selCols();
    if (n > 1000) return this.opts.error && this.opts.error('Select fewer columns to insert.');
    return this.snapshotOp('Insert columns', () => SheetStructure.shift(this.wb, this.sheet, 'col', right ? a + n : a, n));
  };
  P.deleteCols = function () {
    const { a, n } = this.selCols();
    return this.snapshotOp('Delete columns', () => SheetStructure.shift(this.wb, this.sheet, 'col', a, -Math.min(n, Math.max(this.sheet.maxCol, a) - a + 1)));
  };

  P.hideRows = function (hide) {
    const g = this.expandMerges(this.sel);
    const r2 = Math.min(g.r2, Math.max(this.sheet.maxRow, g.r1) + 1000);
    return this.snapshotOp(hide ? 'Hide rows' : 'Unhide rows', () => {
      for (let r = g.r1; r <= r2; r++) {
        const info = this.sheet.rows.get(r);
        if (hide || (info && info.hidden)) this.sheet.setRowHidden(r, hide);
      }
    });
  };
  P.hideCols = function (hide) {
    const g = this.expandMerges(this.sel);
    const c2 = Math.min(g.c2, Math.max(this.sheet.maxCol, g.c1) + 200);
    return this.snapshotOp(hide ? 'Hide columns' : 'Unhide columns', () => {
      for (let c = g.c1; c <= c2; c++) if (hide || this.sheet.isColHidden(c)) this.sheet.setColHidden(c, hide);
    });
  };

  P.rowEdgeNear = function (sy) {
    const r = this.rowAt(sy);
    if (Math.abs(this.rowTop[r + 1] - sy) <= 3) return r;
    if (r > 1 && Math.abs(this.rowTop[r] - sy) <= 3) return r - 1;
    return 0;
  };

  P.startRowResize = function (r, startY) {
    const startH = this.rowTop[r + 1] - this.rowTop[r];
    const guide = document.createElement('div');
    guide.className = 'sg-resize-guide sg-resize-guide-h';
    this.grid.append(guide);
    const place = (h) => { const y = 22 + (r > this.F.rows ? this.rowTop[r] - this.main.scrollTop : this.rowTop[r]) + h; guide.style.top = `${y}px`; };
    place(startH);
    let h = startH;
    const move = (ev) => { h = Math.max(0, startH + (ev.clientY - startY) / (this.zoom || 1)); place(h); };
    const up = () => {
      window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up);
      guide.remove();
      if (h === startH) return;
      const rows = (this.sel.c1 === 1 && this.sel.c2 === SheetModel.MAX_COL && r >= this.sel.r1 && r <= this.sel.r2) ? [this.sel.r1, this.sel.r2] : [r, r];
      this.snapshotOp('Row height', () => { for (let x = rows[0]; x <= rows[1]; x++) this.sheet.setRowHeight(x, Math.round(h * 0.75 * 4) / 4); });
    };
    window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
  };

  P.setRowHeightDialog = async function () {
    const g = this.expandMerges(this.sel);
    const v = await this.opts.ask('Row height', [{ name: 'h', label: 'Height (points)', value: String(this.sheet.rowHeightPt(this.active.r)), type: 'number' }]);
    if (!v || !(+v.h >= 0)) return;
    const r2 = Math.min(g.r2, Math.max(this.sheet.maxRow, g.r1));
    return this.snapshotOp('Row height', () => { for (let r = g.r1; r <= r2; r++) this.sheet.setRowHeight(r, Math.min(409, +v.h)); });
  };
  P.setColWidthDialog = async function () {
    const g = this.expandMerges(this.sel);
    const cur = this.sheet.colWidthChars(this.active.c);
    const shown = cur == null ? 8.43 : Math.max(0, Math.round((cur - 0.7109375) * 100) / 100);
    const v = await this.opts.ask('Column width', [{ name: 'w', label: 'Width (characters)', value: String(shown), type: 'number' }]);
    if (!v || !(+v.w >= 0)) return;
    const c2 = Math.min(g.c2, Math.max(this.sheet.maxCol, g.c1));
    const stored = Math.trunc(((+v.w * 7 + 5) / 7) * 256) / 256;
    return this.snapshotOp('Column width', () => { for (let c = g.c1; c <= c2; c++) this.sheet.setColWidth(c, Math.min(255, stored)); });
  };

  // ---------- freeze panes at the active cell (or unfreeze)

  P.toggleFreeze = function () {
    const s = this.sheet;
    const frozen = s.frozen.rows || s.frozen.cols;
    const rows = frozen ? 0 : this.active.r - 1;
    const cols = frozen ? 0 : this.active.c - 1;
    if (!frozen && !rows && !cols) { this.opts.error && this.opts.error('Select the cell below and to the right of what you want to keep visible.'); return; }
    return this.snapshotOp(frozen ? 'Unfreeze panes' : 'Freeze panes', () => {
      s.frozen = { rows, cols };
      const doc = s.doc;
      const root = doc.documentElement;
      const prefix = root.prefix ? root.prefix + ':' : '';
      let views = [...root.childNodes].find((x) => x.localName === 'sheetViews');
      if (!views) { views = doc.createElementNS(NS, prefix + 'sheetViews'); root.insertBefore(views, [...root.childNodes].find((x) => x.localName === 'sheetFormatPr' || x.localName === 'cols' || x.localName === 'sheetData')); }
      let view = [...views.childNodes].find((x) => x.localName === 'sheetView');
      if (!view) { view = doc.createElementNS(NS, prefix + 'sheetView'); view.setAttribute('workbookViewId', '0'); views.append(view); }
      for (const old of [...view.childNodes].filter((x) => x.localName === 'pane' || x.localName === 'selection')) view.removeChild(old);
      if (rows || cols) {
        const pane = doc.createElementNS(NS, prefix + 'pane');
        if (cols) pane.setAttribute('xSplit', cols);
        if (rows) pane.setAttribute('ySplit', rows);
        pane.setAttribute('topLeftCell', addr(rows + 1, cols + 1));
        const active = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
        pane.setAttribute('activePane', active);
        pane.setAttribute('state', 'frozen');
        view.insertBefore(pane, view.firstChild);
        const sel = doc.createElementNS(NS, prefix + 'selection');
        sel.setAttribute('pane', active); sel.setAttribute('activeCell', addr(rows + 1, cols + 1)); sel.setAttribute('sqref', addr(rows + 1, cols + 1));
        view.append(sel);
      }
      s.structureChanged = true;
    });
  };

  // ======================================================================
  // Right-click menu
  // ======================================================================

  P.onContextMenu = function (e) {
    const h = this.hit(e);
    if (!h.inColHdr && !h.inRowHdr) {
      const g = this.expandMerges(this.sel);
      if (!(h.r >= g.r1 && h.r <= g.r2 && h.c >= g.c1 && h.c <= g.c2)) this.select(h.r, h.c);
    } else if (h.inColHdr && !h.inRowHdr && !(h.c >= this.sel.c1 && h.c <= this.sel.c2 && this.sel.r2 === SheetModel.MAX_ROW)) this.selectCols(h.c, h.c);
    else if (h.inRowHdr && !h.inColHdr && !(h.r >= this.sel.r1 && h.r <= this.sel.r2 && this.sel.c2 === SheetModel.MAX_COL)) this.selectRows(h.r, h.r);
    const whole = this.sel.r1 === 1 && this.sel.r2 === SheetModel.MAX_ROW ? 'cols' : this.sel.c1 === 1 && this.sel.c2 === SheetModel.MAX_COL ? 'rows' : null;
    const run = (fn) => () => { Promise.resolve(fn()).catch((err) => { logError(err); this.opts.error && this.opts.error("Couldn't do that", err.message); }); };
    const items = [
      ['Cut', 'Ctrl+X', run(() => { this.grid.focus(); document.execCommand('cut'); })],
      ['Copy', 'Ctrl+C', run(() => { this.grid.focus(); document.execCommand('copy'); })],
      ['Paste', 'Ctrl+V', run(async () => { this.grid.focus(); const t = this.opts.readClipboard ? await this.opts.readClipboard() : ''; if (t) this.pasteText(t); })],
      null,
    ];
    if (whole !== 'cols') items.push(['Insert row above', '', run(() => this.insertRows(false))], ['Insert row below', '', run(() => this.insertRows(true))]);
    if (whole !== 'rows') items.push(['Insert column left', '', run(() => this.insertCols(false))], ['Insert column right', '', run(() => this.insertCols(true))]);
    if (whole !== 'cols') items.push([this.sel.r2 > this.sel.r1 ? 'Delete rows' : 'Delete row', '', run(() => this.deleteRows())]);
    if (whole !== 'rows') items.push([this.sel.c2 > this.sel.c1 ? 'Delete columns' : 'Delete column', '', run(() => this.deleteCols())]);
    items.push(['Clear contents', 'Delete', run(() => this.clearSelection())], null);
    items.push(['Sort A to Z', '', run(() => this.sortSelection(true))], ['Sort Z to A', '', run(() => this.sortSelection(false))]);
    items.push([this.sheet.autoFilter ? 'Remove filter' : 'Filter', '', run(() => this.toggleFilter())], null);
    if (whole === 'rows' || !whole) items.push(['Hide rows', '', run(() => this.hideRows(true))], ['Unhide rows', '', run(() => this.hideRows(false))], ['Row height…', '', run(() => this.setRowHeightDialog())]);
    if (whole === 'cols' || !whole) items.push(['Hide columns', '', run(() => this.hideCols(true))], ['Unhide columns', '', run(() => this.hideCols(false))], ['Column width…', '', run(() => this.setColWidthDialog())]);
    items.push(null, [this.sheet.frozen.rows || this.sheet.frozen.cols ? 'Unfreeze panes' : 'Freeze panes', '', run(() => this.toggleFreeze())]);
    showMenu(e.clientX, e.clientY, items);
  };

  function showMenu(x, y, items) {
    closeMenus();
    const menu = document.createElement('div');
    menu.className = 'sg-menu';
    menu.setAttribute('role', 'menu');
    for (const it of items) {
      if (!it) { if (menu.lastChild && !menu.lastChild.classList.contains('sep')) menu.append(Object.assign(document.createElement('div'), { className: 'sep' })); continue; }
      const b = document.createElement('button');
      b.setAttribute('role', 'menuitem');
      b.innerHTML = `<span></span><kbd></kbd>`;
      b.firstChild.textContent = it[0];
      b.lastChild.textContent = it[1] || '';
      b.addEventListener('mousedown', (ev) => ev.preventDefault());
      b.addEventListener('click', () => { closeMenus(); it[2](); });
      menu.append(b);
    }
    document.body.append(menu);
    const r = menu.getBoundingClientRect();
    menu.style.left = `${Math.min(x, innerWidth - r.width - 4)}px`;
    menu.style.top = `${Math.min(y, innerHeight - r.height - 4)}px`;
    setTimeout(() => {
      const away = (ev) => { if (!menu.contains(ev.target)) closeMenus(); };
      const esc = (ev) => { if (ev.key === 'Escape') closeMenus(); };
      document.addEventListener('mousedown', away, true);
      document.addEventListener('keydown', esc, true);
      menu._cleanup = () => { document.removeEventListener('mousedown', away, true); document.removeEventListener('keydown', esc, true); };
    });
    return menu;
  }
  function closeMenus() { document.querySelectorAll('.sg-menu, .sg-popup').forEach((m) => { if (m._cleanup) m._cleanup(); m.remove(); }); }
  SheetView.showMenu = showMenu;

  // Paste text from the menu (keyboard paste uses the clipboard event).
  P.pasteText = function (text) {
    const fake = { preventDefault() {}, clipboardData: { getData: () => text } };
    this.onPaste(fake);
  };

  // ======================================================================
  // Sort
  // ======================================================================

  // The block of data around the active cell (Excel's "current region").
  P.currentRegion = function () {
    const s = this.sheet;
    const filled = (r, c) => { const x = s.get(r, c); return !!x && ((x.value !== null && x.value !== undefined && x.value !== '') || x.formula); };
    let g = { r1: this.active.r, c1: this.active.c, r2: this.active.r, c2: this.active.c };
    for (let grew = true; grew;) {
      grew = false;
      const rowHas = (r) => { for (let c = Math.max(1, g.c1 - 1); c <= g.c2 + 1; c++) if (filled(r, c)) return true; return false; };
      const colHas = (c) => { for (let r = Math.max(1, g.r1 - 1); r <= g.r2 + 1; r++) if (filled(r, c)) return true; return false; };
      if (g.r1 > 1 && rowHas(g.r1 - 1)) { g.r1--; grew = true; }
      if (rowHas(g.r2 + 1)) { g.r2++; grew = true; }
      if (g.c1 > 1 && colHas(g.c1 - 1)) { g.c1--; grew = true; }
      if (colHas(g.c2 + 1)) { g.c2++; grew = true; }
    }
    return g;
  };

  function sortKey(v) {
    if (v === null || v === undefined || v === '') return [4, 0];
    if (typeof v === 'number') return [0, v];
    if (typeof v === 'string') return [1, v.toLowerCase()];
    if (typeof v === 'boolean') return [2, v ? 1 : 0];
    return [3, 0];
  }

  P.sortSelection = function (asc, byCol) {
    const single = this.sel.r1 === this.sel.r2 && this.sel.c1 === this.sel.c2;
    let g = single ? this.currentRegion() : this.effectiveRange();
    const col = byCol || this.active.c;
    const s = this.sheet;
    if (s.merges.some((m) => m.r2 >= g.r1 && m.r1 <= g.r2 && m.c2 >= g.c1 && m.c1 <= g.c2)) { this.opts.error && this.opts.error("Merged cells can't be sorted. Unmerge them first."); return; }
    // Header row: Excel's guess - the first row is text where the data below isn't, or it is bold.
    const top = s.get(g.r1, col), below = s.get(g.r1 + 1, col);
    const hasHeader = g.r2 > g.r1 && top && typeof top.value === 'string' && ((below && typeof below.value !== 'string') || this.wb.styles.font(top.s).bold);
    return this.sortRange({ ...g, r1: hasHeader ? g.r1 + 1 : g.r1 }, col, asc);
  };

  P.sortRange = function (g, col, asc) {
    const s = this.sheet;
    if (g.r2 <= g.r1) return;
    const rows = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const cells = [];
      for (let c = g.c1; c <= g.c2; c++) { const x = s.get(r, c); cells.push(x ? { value: x.value, formula: x.formula, s: x.s } : { value: null, formula: null, s: s.defaultStyleAt(r, c) }); }
      const k = s.get(r, col);
      rows.push({ r, cells, key: sortKey(k ? k.value : null) });
    }
    rows.sort((a, b) => {
      // Blanks always go last, whatever the direction (as in Excel).
      if (a.key[0] === 4 || b.key[0] === 4) return a.key[0] - b.key[0];
      const d = a.key[0] !== b.key[0] ? a.key[0] - b.key[0] : (a.key[1] < b.key[1] ? -1 : a.key[1] > b.key[1] ? 1 : 0);
      return (asc ? d : -d) || a.r - b.r;
    });
    const items = [];
    rows.forEach((row, i) => {
      const r = g.r1 + i;
      row.cells.forEach((x, j) => {
        const patch = x.formula ? { formula: shiftFormula(x.formula, r - row.r, 0), s: x.s } : { value: x.value, s: x.s };
        items.push({ r, c: g.c1 + j, patch });
      });
    });
    this.apply(items);
    this.opts.onStatus && this.opts.onStatus(`Sorted ${rows.length} rows by column ${colName(col)}`);
  };

  // ======================================================================
  // AutoFilter
  // ======================================================================

  // Filters on this sheet: the sheet's own AutoFilter and each table's built-in filter.
  P.filters = function () {
    const out = [];
    const s = this.sheet;
    if (s.autoFilter && s.autoFilter.ref) out.push({ kind: 'sheet', ref: s.autoFilter.ref, cols: s.autoFilter.cols });
    for (const t of this.wb.tables || []) if (t.sheet === s && t.autoFilter) out.push({ kind: 'table', table: t, ref: t.autoFilter.ref, cols: t.autoFilter.cols });
    return out;
  };
  P.filterAt = function (r, c) { return this.filters().find((f) => { const g = parseRange(f.ref); return g && r === g.r1 && c >= g.c1 && c <= g.c2; }) || null; };
  P.filterRange = function () { const f = this.filters().find((x) => x.kind === 'sheet'); return f ? parseRange(f.ref) : null; };
  P.filterButtonAt = function (r, c) { return !!this.filterAt(r, c); };
  P.filterActiveAt = function (r, c) { const f = this.filterAt(r, c); return !!(f && f.cols.has(c)); };

  P.toggleFilter = function () {
    const s = this.sheet;
    const t = this.wb.tableAt(s, this.active.r, this.active.c);
    if (t) {
      // Inside a table: switch the table's own filter buttons on or off (a sheet filter can't overlap a table).
      return this.snapshotOp(t.autoFilter ? 'Remove filter' : 'Filter', () => {
        const root = t.doc.documentElement;
        if (t.autoFilter) {
          const g = parseRange(t.autoFilter.ref);
          for (let r = g.r1 + 1; r <= g.r2; r++) { const info = s.rows.get(r); if (info && info.hidden) s.setRowHidden(r, false); }
          const el = [...root.childNodes].find((x) => x.localName === 'autoFilter');
          if (el) root.removeChild(el);
          t.autoFilter = null;
        } else {
          const el = t.doc.createElementNS(root.namespaceURI, (root.prefix ? root.prefix + ':' : '') + 'autoFilter');
          el.setAttribute('ref', rangeRef({ ...t.g, r2: t.g.r2 - t.totals }));
          const first = [...root.childNodes].find((x) => x.nodeType === 1);
          root.insertBefore(el, first || null);
          t.autoFilter = { ref: el.getAttribute('ref'), cols: new Map(), el };
        }
        this.wb.markPart(t.path);
      });
    }
    if (s.autoFilter) {
      return this.snapshotOp('Remove filter', () => {
        const g = parseRange(s.autoFilter.ref);
        for (let r = g.r1 + 1; r <= g.r2; r++) { const info = s.rows.get(r); if (info && info.hidden) s.setRowHidden(r, false); }
        s.autoFilter = null;
        s.filterChanged = true;
      });
    }
    const single = this.sel.r1 === this.sel.r2 && this.sel.c1 === this.sel.c2;
    const g = single ? this.currentRegion() : this.effectiveRange();
    if (g.r2 === g.r1 && single) { this.opts.error && this.opts.error('Select a cell in a table of data first.'); return; }
    if ((this.wb.tables || []).some((tb) => tb.sheet === s && tb.g.r2 >= g.r1 && tb.g.r1 <= g.r2 && tb.g.c2 >= g.c1 && tb.g.c1 <= g.c2)) {
      this.opts.error && this.opts.error('That range overlaps a table. Click inside the table to use its own filter.');
      return;
    }
    return this.snapshotOp('Filter', () => { s.autoFilter = { ref: rangeRef(g), cols: new Map() }; s.filterChanged = true; });
  };

  P.cellText = function (r, c) { return SheetView.display(this.sheet.get(r, c), this.wb.styles, this.wb.date1904).text; };

  P.openFilterMenu = function (r, c, btn) {
    const f = this.filterAt(r, c);
    if (!f) return;
    let g = parseRange(f.ref);
    // A sheet filter grows to cover data added below it (tables have fixed extents).
    if (f.kind === 'sheet') {
      const save = this.active; this.active = { r: g.r1, c: g.c1 };
      const region = this.currentRegion();
      this.active = save;
      if (region.r2 > g.r2) g = { ...g, r2: region.r2 };
    }
    const values = new Map();
    let blanks = false;
    for (let rr = g.r1 + 1; rr <= g.r2; rr++) { const t = this.cellText(rr, c); if (t === '') blanks = true; else values.set(t, (values.get(t) || 0) + 1); }
    const cur = f.cols.get(c);
    const keys = [...values.keys()].sort((a, b) => { const na = parseFloat(a.replace(/[^0-9.-]/g, '')), nb = parseFloat(b.replace(/[^0-9.-]/g, '')); return !isNaN(na) && !isNaN(nb) && /\d/.test(a) && /\d/.test(b) ? na - nb : a.localeCompare(b); });
    closeMenus();
    const pop = document.createElement('div');
    pop.className = 'sg-popup';
    const mk = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
    const sortA = mk('button', 'link-btn', 'Sort A to Z'), sortZ = mk('button', 'link-btn', 'Sort Z to A');
    const search = mk('input'); search.placeholder = 'Search'; search.type = 'search';
    const list = mk('div', 'sg-filter-list');
    const boxes = [];
    const addBox = (label, val, checked) => {
      const l = mk('label');
      const cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = checked; cb.value = val;
      l.append(cb, document.createTextNode(label));
      list.append(l);
      boxes.push({ cb, l, label: label.toLowerCase(), val });
      return cb;
    };
    const all = addBox('(Select all)', '__all', true);
    for (const k of keys) addBox(k, k, !cur || cur.values.has(k));
    if (blanks) addBox('(Blanks)', '__blank', !cur || cur.blanks);
    all.checked = boxes.slice(1).every((b) => b.cb.checked);
    all.addEventListener('change', () => boxes.slice(1).forEach((b) => { if (b.l.style.display !== 'none') b.cb.checked = all.checked; }));
    search.addEventListener('input', () => { const q = search.value.toLowerCase(); boxes.slice(1).forEach((b) => { b.l.style.display = b.label.includes(q) ? '' : 'none'; }); });
    const ok = mk('button', '', 'OK'), cancel = mk('button', 'ghost', 'Cancel'), clear = mk('button', 'link-btn', 'Clear filter');
    const actions = mk('div', 'sg-popup-actions'); actions.append(clear, cancel, ok);
    pop.append(sortA, sortZ, mk('hr'), search, list, actions);
    document.body.append(pop);
    const rect = btn.getBoundingClientRect();
    pop.style.left = `${Math.min(rect.left, innerWidth - 270)}px`;
    pop.style.top = `${Math.min(rect.bottom + 2, innerHeight - pop.offsetHeight - 6)}px`;
    search.focus();
    const done = () => closeMenus();
    sortA.onclick = () => { done(); this.sortRange({ ...g, r1: g.r1 + 1 }, c, true); };
    sortZ.onclick = () => { done(); this.sortRange({ ...g, r1: g.r1 + 1 }, c, false); };
    cancel.onclick = done;
    clear.onclick = () => { done(); this.applyFilter(g, c, null, f); };
    ok.onclick = () => {
      done();
      const picked = boxes.slice(1).filter((b) => b.cb.checked);
      const everything = picked.length === boxes.length - 1;
      this.applyFilter(g, c, everything ? null : { values: new Set(picked.filter((b) => b.val !== '__blank').map((b) => b.val)), blanks: picked.some((b) => b.val === '__blank') }, f);
    };
    setTimeout(() => {
      const away = (ev) => { if (!pop.contains(ev.target)) closeMenus(); };
      document.addEventListener('mousedown', away, true);
      pop._cleanup = () => document.removeEventListener('mousedown', away, true);
    });
  };

  // Apply a column filter (null clears it) and show/hide the rows. `f` defaults to the sheet filter.
  P.applyFilter = function (g, c, filter, f) {
    const s = this.sheet;
    f = f || this.filters().find((x) => x.kind === 'sheet') || this.filterAt(g.r1, c);
    return this.snapshotOp('Filter', () => {
      const t = f.kind === 'table' ? this.wb.tables.find((x) => x.path === f.table.path) : null;
      const cols = t ? t.autoFilter.cols : s.autoFilter.cols;
      if (filter) cols.set(c, filter); else cols.delete(c);
      for (let r = g.r1 + 1; r <= g.r2; r++) {
        let show = true;
        for (const [fc, ff] of cols) {
          const text = this.cellText(r, fc);
          if (text === '' ? !ff.blanks : !ff.values.has(text)) { show = false; break; }
        }
        const info = s.rows.get(r);
        if (!!(info && info.hidden) !== !show) s.setRowHidden(r, !show);
      }
      if (t) {
        // The table's own <autoFilter> carries the criteria.
        const doc = t.doc;
        const el = t.autoFilter.el || [...doc.documentElement.childNodes].find((x) => x.localName === 'autoFilter');
        const p = doc.documentElement.prefix ? doc.documentElement.prefix + ':' : '';
        const make = (n, attrs) => { const e = doc.createElementNS(doc.documentElement.namespaceURI, p + n); for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v); return e; };
        for (const old of [...el.childNodes].filter((x) => x.localName === 'filterColumn')) el.removeChild(old);
        const fg = parseRange(el.getAttribute('ref'));
        const before = [...el.childNodes].find((x) => x.nodeType === 1) || null;
        for (const [fc, ff] of [...cols].sort((a, b) => a[0] - b[0])) {
          const col = make('filterColumn', { colId: fc - fg.c1 });
          const fl = make('filters', ff.blanks ? { blank: '1' } : {});
          for (const v of ff.values) fl.append(make('filter', { val: v }));
          col.append(fl);
          el.insertBefore(col, before);
        }
        this.wb.markPart(t.path);
      } else {
        s.autoFilter.ref = rangeRef(g);
        s.filterChanged = true;
      }
      s.structureChanged = true;
    });
  };

  // ======================================================================
  // Conditional formatting: draw the file's rules, and add new ones
  // ======================================================================

  P.cfInvalidate = function () { this.cfCache = new Map(); this.cfRulesFor = null; };

  const changed = P.changed;
  P.changed = function () { this.cfInvalidate(); this.scheduleChartRefresh(); return changed.call(this); };

  function dxfStyles(wb) {
    if (wb._dxfs) return wb._dxfs;
    const doc = wb.styles.doc;
    const out = [];
    const dxfs = [...doc.documentElement.childNodes].find((x) => x.localName === 'dxfs');
    for (const d of dxfs ? [...dxfs.childNodes].filter((x) => x.localName === 'dxf') : []) {
      const st = {};
      const font = [...d.childNodes].find((x) => x.localName === 'font');
      if (font) for (const f of font.childNodes) {
        if (f.nodeType !== 1) continue;
        const v = f.getAttribute('val');
        if (f.localName === 'b') st.bold = v !== '0' && v !== 'false';
        if (f.localName === 'i') st.italic = v !== '0' && v !== 'false';
        if (f.localName === 'strike') st.strike = v !== '0' && v !== 'false';
        if (f.localName === 'u') st.underline = (v || 'single') !== 'none';
        if (f.localName === 'color') st.color = wb.styles.color(f);
      }
      const fill = [...d.childNodes].find((x) => x.localName === 'fill');
      const pf = fill && [...fill.childNodes].find((x) => x.localName === 'patternFill');
      if (pf) {
        const bg = [...pf.childNodes].find((x) => x.localName === 'bgColor');
        const fg = [...pf.childNodes].find((x) => x.localName === 'fgColor');
        st.fill = wb.styles.color(bg) || wb.styles.color(fg);
      }
      const nf = [...d.childNodes].find((x) => x.localName === 'numFmt');
      if (nf) st.numFmt = nf.getAttribute('formatCode');
      out.push(st);
    }
    wb._dxfs = out;
    return out;
  }

  P.loadCfRules = function () {
    const s = this.sheet;
    const rules = [];
    if (!s.doc) return rules;
    const dx = dxfStyles(this.wb);
    for (const cfEl of [...s.doc.documentElement.childNodes].filter((x) => x.localName === 'conditionalFormatting')) {
      const ranges = (cfEl.getAttribute('sqref') || '').split(/\s+/).map(parseRange).filter(Boolean);
      if (!ranges.length) continue;
      for (const r of [...cfEl.childNodes].filter((x) => x.localName === 'cfRule')) {
        const at = (n, d) => (r.hasAttribute(n) ? r.getAttribute(n) : d);
        const kidsOf = (el, name) => [...el.childNodes].filter((x) => x.localName === name);
        const rule = {
          ranges, type: at('type'), op: at('operator'), priority: +at('priority', 9999), stop: at('stopIfTrue') === '1',
          dxf: r.hasAttribute('dxfId') ? dx[+r.getAttribute('dxfId')] : null, text: at('text'),
          rank: +at('rank', 10), percent: at('percent') === '1', bottom: at('bottom') === '1',
          above: at('aboveAverage', '1') !== '0', equal: at('equalAverage') === '1', stdDev: +at('stdDev', 0),
          period: at('timePeriod'), formulas: kidsOf(r, 'formula').map((f) => f.textContent),
        };
        const scale = kidsOf(r, 'colorScale')[0], bar = kidsOf(r, 'dataBar')[0], icons = kidsOf(r, 'iconSet')[0];
        const cfvo = (el) => kidsOf(el, 'cfvo').map((v) => ({ type: v.getAttribute('type'), val: v.getAttribute('val'), gte: v.getAttribute('gte') !== '0' }));
        if (scale) { rule.cfvo = cfvo(scale); rule.colors = kidsOf(scale, 'color').map((x) => this.wb.styles.color(x) || '#FFFFFF'); }
        if (bar) { rule.cfvo = cfvo(bar); rule.color = this.wb.styles.color(kidsOf(bar, 'color')[0]) || '#638EC6'; }
        if (icons) { rule.cfvo = cfvo(icons); rule.iconSet = icons.getAttribute('iconSet') || '3TrafficLights1'; rule.reverse = icons.getAttribute('reverse') === '1'; rule.showValue = icons.getAttribute('showValue') !== '0'; }
        rules.push(rule);
      }
    }
    rules.sort((a, b) => a.priority - b.priority);
    return rules;
  };

  P.ruleValues = function (rule) {
    if (rule.values) return rule.values;
    const out = [];
    for (const g of rule.ranges) {
      const r2 = Math.min(g.r2, this.sheet.maxRow), c2 = Math.min(g.c2, this.sheet.maxCol);
      for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) { const x = this.sheet.get(r, c); if (x) out.push(x.value); }
    }
    rule.values = out;
    return out;
  };

  function percentile(sorted, p) {
    if (!sorted.length) return 0;
    const i = (sorted.length - 1) * p;
    const lo = Math.floor(i), hi = Math.ceil(i);
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (i - lo);
  }

  P.cfvoValue = function (rule, v, nums) {
    const sorted = nums.slice().sort((a, b) => a - b);
    const min = sorted[0] ?? 0, max = sorted[sorted.length - 1] ?? 0;
    switch (v.type) {
      case 'min': return min;
      case 'max': return max;
      case 'num': return +v.val;
      case 'percent': return min + (max - min) * (+v.val / 100);
      case 'percentile': return percentile(sorted, +v.val / 100);
      case 'formula': { const x = this.calc.engine.evaluate(v.val, this.sheet.name, 1, 1); return typeof x === 'number' ? x : 0; }
      case 'autoMin': return Math.min(0, min);
      case 'autoMax': return Math.max(0, max);
      default: return min;
    }
  };

  const mix = (a, b, t) => {
    const p = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
    const x = p(a), y = p(b);
    return '#' + x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('');
  };

  const ICONS = {
    Arrows: [['↓', '#C0504D'], ['→', '#F2B200'], ['↑', '#3B9C3B']],
    ArrowsGray: [['↓', '#7F7F7F'], ['→', '#7F7F7F'], ['↑', '#7F7F7F']],
    TrafficLights: [['●', '#C0504D'], ['●', '#F2B200'], ['●', '#3B9C3B']],
    Symbols: [['✖', '#C0504D'], ['!', '#F2B200'], ['✔', '#3B9C3B']],
    Flags: [['⚑', '#C0504D'], ['⚑', '#F2B200'], ['⚑', '#3B9C3B']],
    Stars: [['☆', '#BFBFBF'], ['★', '#F2B200'], ['★', '#F2B200']],
  };
  function iconFor(set, idx, count) {
    const fam = Object.keys(ICONS).find((k) => set.includes(k.replace('Gray', ''))) || 'TrafficLights';
    const pal = ICONS[set.includes('Gray') ? 'ArrowsGray' : fam];
    const i = count <= 3 ? idx : Math.round((idx / (count - 1)) * 2);
    const [char, color] = pal[Math.max(0, Math.min(2, i))];
    return { char, color };
  }

  P.ruleMatches = function (rule, r, c, cell) {
    const v = cell ? cell.value : null;
    const g0 = rule.ranges[0];
    const evalF = (f) => {
      const shifted = shiftFormula(f, r - g0.r1, c - g0.c1);
      const x = this.calc.engine.evaluate(this.calc.prep(shifted, this.sheet, r, c), this.sheet.name, r, c);
      return this.calc.engine.isError(x) ? null : Array.isArray(x) ? (x[0] ? x[0][0] : null) : x;
    };
    const num = (x) => (typeof x === 'number' ? x : typeof x === 'boolean' ? +x : x === null || x === undefined || x === '' ? 0 : x);
    const cmp = (a, b) => { a = num(a); b = num(b); if (typeof a === 'string' || typeof b === 'string') { a = String(a).toLowerCase(); b = String(b).toLowerCase(); } return a < b ? -1 : a > b ? 1 : 0; };
    const s = v === null || v === undefined ? '' : String(v instanceof XErr ? v.code : v);
    switch (rule.type) {
      case 'cellIs': {
        if (v instanceof XErr) return false;
        const f = rule.formulas.map(evalF);
        switch (rule.op) {
          case 'equal': return cmp(v, f[0]) === 0;
          case 'notEqual': return cmp(v, f[0]) !== 0;
          case 'greaterThan': return cmp(v, f[0]) > 0;
          case 'lessThan': return cmp(v, f[0]) < 0;
          case 'greaterThanOrEqual': return cmp(v, f[0]) >= 0;
          case 'lessThanOrEqual': return cmp(v, f[0]) <= 0;
          case 'between': { const [a, b] = [f[0], f[1]].sort((x, y) => cmp(x, y)); return cmp(v, a) >= 0 && cmp(v, b) <= 0; }
          case 'notBetween': { const [a, b] = [f[0], f[1]].sort((x, y) => cmp(x, y)); return cmp(v, a) < 0 || cmp(v, b) > 0; }
          default: return false;
        }
      }
      case 'expression': { const x = evalF(rule.formulas[0] || 'FALSE'); return x === true || (typeof x === 'number' && x !== 0); }
      case 'containsText': return s.toLowerCase().includes((rule.text || '').toLowerCase());
      case 'notContainsText': return !s.toLowerCase().includes((rule.text || '').toLowerCase());
      case 'beginsWith': return s.toLowerCase().startsWith((rule.text || '').toLowerCase());
      case 'endsWith': return s.toLowerCase().endsWith((rule.text || '').toLowerCase());
      case 'containsBlanks': return s.trim() === '';
      case 'notContainsBlanks': return s.trim() !== '';
      case 'containsErrors': return v instanceof XErr;
      case 'notContainsErrors': return !(v instanceof XErr);
      case 'duplicateValues': case 'uniqueValues': {
        if (s === '') return false;
        if (!rule.counts) { rule.counts = new Map(); for (const x of this.ruleValues(rule)) { if (x === null || x === '') continue; const k = String(x).toLowerCase(); rule.counts.set(k, (rule.counts.get(k) || 0) + 1); } }
        const n = rule.counts.get(s.toLowerCase()) || 0;
        return rule.type === 'duplicateValues' ? n > 1 : n === 1;
      }
      case 'top10': {
        if (typeof v !== 'number') return false;
        if (!rule.cut) {
          const nums = this.ruleValues(rule).filter((x) => typeof x === 'number').sort((a, b) => (rule.bottom ? a - b : b - a));
          const n = Math.max(1, rule.percent ? Math.floor(nums.length * rule.rank / 100) : rule.rank);
          rule.cut = nums[Math.min(n, nums.length) - 1];
        }
        return rule.bottom ? v <= rule.cut : v >= rule.cut;
      }
      case 'aboveAverage': {
        if (typeof v !== 'number') return false;
        if (rule.avg === undefined) {
          const nums = this.ruleValues(rule).filter((x) => typeof x === 'number');
          rule.avg = nums.reduce((a, b) => a + b, 0) / (nums.length || 1);
          rule.sd = Math.sqrt(nums.reduce((a, b) => a + (b - rule.avg) ** 2, 0) / Math.max(1, nums.length - 1));
        }
        const t = rule.avg + (rule.above ? 1 : -1) * rule.stdDev * rule.sd;
        return rule.above ? (rule.equal ? v >= t : v > t) : (rule.equal ? v <= t : v < t);
      }
      case 'timePeriod': {
        if (typeof v !== 'number') return false;
        const today = Math.floor(this.wb.dateToSerial(new Date(Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()))));
        const d = Math.floor(v);
        const dow = (x) => (((x - 1) % 7) + 7) % 7; // Sunday = 0 for serials since 1900-01-01 (a Sunday in Excel's calendar)
        const date = (x) => new Date(Date.UTC(1899, 11, 30) + x * 86400000);
        const ym = (x) => { const t = date(x); return t.getUTCFullYear() * 12 + t.getUTCMonth(); };
        switch (rule.period) {
          case 'today': return d === today;
          case 'yesterday': return d === today - 1;
          case 'tomorrow': return d === today + 1;
          case 'last7Days': return d <= today && d > today - 7;
          case 'thisWeek': return d - dow(d) === today - dow(today);
          case 'lastWeek': return d - dow(d) === today - dow(today) - 7;
          case 'nextWeek': return d - dow(d) === today - dow(today) + 7;
          case 'thisMonth': return ym(d) === ym(today);
          case 'lastMonth': return ym(d) === ym(today) - 1;
          case 'nextMonth': return ym(d) === ym(today) + 1;
          default: return false;
        }
      }
      default: return false;
    }
  };

  P.cfFor = function (r, c, cell) {
    const k = key(r, c);
    if (!this.cfCache) this.cfCache = new Map();
    if (this.cfCache.has(k)) return this.cfCache.get(k);
    if (!this.cfRulesFor || this.cfRulesFor.sheet !== this.sheet) this.cfRulesFor = { sheet: this.sheet, rules: this.loadCfRules() };
    let out = null;
    for (const rule of this.cfRulesFor.rules) {
      if (!rule.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)) continue;
      const v = cell ? cell.value : null;
      if (rule.type === 'colorScale' || rule.type === 'dataBar' || rule.type === 'iconSet') {
        if (typeof v !== 'number') continue;
        const nums = this.ruleValues(rule).filter((x) => typeof x === 'number');
        const th = rule.cfvo.map((x) => this.cfvoValue(rule, x, nums));
        out = out || {};
        if (rule.type === 'colorScale' && !out.fill) {
          if (rule.colors.length === 2 || th.length === 2) { const t = th[1] === th[0] ? 0 : (v - th[0]) / (th[1] - th[0]); out.fill = mix(rule.colors[0], rule.colors[1], Math.max(0, Math.min(1, t))); }
          else { const t = v <= th[1] ? (th[1] === th[0] ? 0 : (v - th[0]) / (th[1] - th[0])) : (th[2] === th[1] ? 1 : (v - th[1]) / (th[2] - th[1])); out.fill = v <= th[1] ? mix(rule.colors[0], rule.colors[1], Math.max(0, Math.min(1, t))) : mix(rule.colors[1], rule.colors[2], Math.max(0, Math.min(1, t))); }
        } else if (rule.type === 'dataBar' && !out.bar) {
          const lo = th[0], hi = th[1];
          out.bar = { pct: hi === lo ? 100 : 10 + 90 * Math.max(0, Math.min(1, (v - lo) / (hi - lo))), color: rule.color };
        } else if (rule.type === 'iconSet' && !out.icon) {
          let idx = 0;
          for (let i = 1; i < th.length; i++) if (rule.cfvo[i].gte ? v >= th[i] : v > th[i]) idx = i;
          if (rule.reverse) idx = th.length - 1 - idx;
          out.icon = iconFor(rule.iconSet, idx, th.length);
        }
        if (rule.stop) break;
        continue;
      }
      if (!this.ruleMatches(rule, r, c, cell)) continue;
      if (rule.dxf) { out = out || {}; for (const [p, val] of Object.entries(rule.dxf)) if (out[p] === undefined && val !== null && val !== undefined) out[p] = val; }
      if (rule.stop) break;
    }
    this.cfCache.set(k, out);
    return out;
  };

  // ---- creating rules

  P.addCfRule = function (spec) {
    const g = this.effectiveRange();
    const sqref = rangeRef(g);
    return this.snapshotOp('Conditional formatting', () => {
      const s = this.sheet;
      const doc = s.doc;
      const root = doc.documentElement;
      const prefix = root.prefix ? root.prefix + ':' : '';
      const el = (n, attrs) => { const e = doc.createElementNS(NS, prefix + n); for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v); return e; };
      const prios = [...doc.getElementsByTagNameNS(NS, 'cfRule')].map((x) => +x.getAttribute('priority') || 0);
      const cf = el('conditionalFormatting', { sqref });
      const rule = el('cfRule', { type: spec.type, priority: Math.max(0, ...prios) + 1 });
      if (spec.dxf) rule.setAttribute('dxfId', this.addDxf(spec.dxf));
      for (const [k, v] of Object.entries(spec.attrs || {})) rule.setAttribute(k, v);
      for (const f of spec.formulas || []) { const x = el('formula'); x.textContent = f; rule.append(x); }
      if (spec.child) rule.append(spec.child(el));
      cf.append(rule);
      const order = ['conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
      const existing = [...root.childNodes].filter((x) => x.localName === 'conditionalFormatting');
      const after = existing.length ? existing[existing.length - 1].nextSibling : [...root.childNodes].find((x) => x.nodeType === 1 && order.includes(x.localName));
      root.insertBefore(cf, after || null);
      s.structureChanged = true;
    });
  };

  P.addDxf = function (st) {
    const sdoc = this.wb.styles.doc;
    const root = sdoc.documentElement;
    const el = (n, attrs) => { const e = sdoc.createElementNS(root.namespaceURI, (root.prefix ? root.prefix + ':' : '') + n); for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v); return e; };
    let dxfs = [...root.childNodes].find((x) => x.localName === 'dxfs');
    if (!dxfs) {
      dxfs = el('dxfs', { count: 0 });
      const before = [...root.childNodes].find((x) => ['tableStyles', 'colors', 'extLst'].includes(x.localName));
      root.insertBefore(dxfs, before || null);
    }
    const d = el('dxf');
    if (st.color || st.bold) { const f = el('font'); if (st.bold) f.append(el('b')); if (st.color) f.append(el('color', { rgb: 'FF' + st.color.replace('#', '').toUpperCase() })); d.append(f); }
    if (st.fill) { const f = el('fill'); const p = el('patternFill'); p.append(el('bgColor', { rgb: 'FF' + st.fill.replace('#', '').toUpperCase() })); f.append(p); d.append(f); }
    dxfs.append(d);
    const n = [...dxfs.childNodes].filter((x) => x.localName === 'dxf').length;
    dxfs.setAttribute('count', n);
    this.wb.styles.dirty = true;
    this.wb._dxfs = null;
    return n - 1;
  };

  P.clearCfRules = function (wholeSheet) {
    const g = this.effectiveRange();
    return this.snapshotOp('Clear rules', () => {
      const s = this.sheet;
      const root = s.doc.documentElement;
      for (const cf of [...root.childNodes].filter((x) => x.localName === 'conditionalFormatting')) {
        if (wholeSheet) { root.removeChild(cf); continue; }
        const ranges = (cf.getAttribute('sqref') || '').split(/\s+/).filter(Boolean);
        const keep = ranges.filter((ref) => { const r = parseRange(ref); return !(r && r.r1 >= g.r1 && r.r2 <= g.r2 && r.c1 >= g.c1 && r.c2 <= g.c2); });
        if (!keep.length) root.removeChild(cf); else cf.setAttribute('sqref', keep.join(' '));
      }
      s.structureChanged = true;
    });
  };

  P.cfPreset = async function (kind) {
    const tl = addr(this.effectiveRange().r1, this.effectiveRange().c1);
    const red = { fill: '#FFC7CE', color: '#9C0006' };
    const ask = (title, fields) => this.opts.ask(title, fields);
    switch (kind) {
      case 'gt': case 'lt': case 'eq': {
        const v = await ask(kind === 'gt' ? 'Highlight cells greater than' : kind === 'lt' ? 'Highlight cells less than' : 'Highlight cells equal to', [{ name: 'v', label: 'Value or formula', value: '' }]);
        if (!v || v.v === '') return;
        const val = v.v.startsWith('=') ? v.v.slice(1) : isNaN(+v.v) ? `"${v.v.replace(/"/g, '""')}"` : v.v;
        return this.addCfRule({ type: 'cellIs', attrs: { operator: { gt: 'greaterThan', lt: 'lessThan', eq: 'equal' }[kind] }, formulas: [val], dxf: red });
      }
      case 'between': {
        const v = await ask('Highlight cells between', [{ name: 'a', label: 'From', value: '' }, { name: 'b', label: 'To', value: '' }]);
        if (!v || v.a === '' || v.b === '') return;
        return this.addCfRule({ type: 'cellIs', attrs: { operator: 'between' }, formulas: [v.a, v.b], dxf: red });
      }
      case 'text': {
        const v = await ask('Highlight cells containing text', [{ name: 't', label: 'Text', value: '' }]);
        if (!v || !v.t) return;
        const t = v.t.replace(/"/g, '""');
        return this.addCfRule({ type: 'containsText', attrs: { operator: 'containsText', text: v.t }, formulas: [`NOT(ISERROR(SEARCH("${t}",${tl})))`], dxf: red });
      }
      case 'dup': return this.addCfRule({ type: 'duplicateValues', dxf: red });
      case 'top10': return this.addCfRule({ type: 'top10', attrs: { rank: 10 }, dxf: { fill: '#C6EFCE', color: '#006100' } });
      case 'bar': return this.addCfRule({ type: 'dataBar', child: (el) => { const b = el('dataBar'); b.append(el('cfvo', { type: 'min' }), el('cfvo', { type: 'max' }), el('color', { rgb: 'FF638EC6' })); return b; } });
      case 'scale': return this.addCfRule({ type: 'colorScale', child: (el) => { const b = el('colorScale'); b.append(el('cfvo', { type: 'min' }), el('cfvo', { type: 'percentile', val: '50' }), el('cfvo', { type: 'max' }), el('color', { rgb: 'FFF8696B' }), el('color', { rgb: 'FFFFEB84' }), el('color', { rgb: 'FF63BE7B' })); return b; } });
      case 'icons': return this.addCfRule({ type: 'iconSet', child: (el) => { const b = el('iconSet', { iconSet: '3Arrows' }); b.append(el('cfvo', { type: 'percent', val: '0' }), el('cfvo', { type: 'percent', val: '33' }), el('cfvo', { type: 'percent', val: '67' })); return b; } });
      case 'clear': return this.clearCfRules(false);
      case 'clearAll': return this.clearCfRules(true);
      default: return undefined;
    }
  };

  // ======================================================================
  // Excel's built-in table styles (TableStyleLight/Medium/Dark n), approximated from the theme colours
  // ======================================================================

  function tint(hex, t) { // t > 0 lightens towards white
    const v = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
    return '#' + v.map((x) => Math.round(t >= 0 ? x + (255 - x) * t : x * (1 + t)).toString(16).padStart(2, '0')).join('');
  }

  P.tableLook = function (t) {
    if (t._look !== undefined) return t._look;
    const info = [...t.doc.documentElement.childNodes].find((x) => x.localName === 'tableStyleInfo');
    const name = info && info.getAttribute('name');
    const m = name && /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
    if (!m) { t._look = null; return null; }
    const n = +m[2];
    const idx = (n - 1) % 7;
    const block = Math.floor((n - 1) / 7);
    const theme = this.wb.theme;
    const accent = idx === 0 ? '#' + (theme[1] || '000000') : '#' + (theme[3 + idx] || '4472C4');
    const flag = (a, d) => (info.hasAttribute(a) ? info.getAttribute(a) === '1' : d);
    const look = { rows: flag('showRowStripes', true), cols: flag('showColumnStripes', false), first: flag('showFirstColumn', false), last: flag('showLastColumn', false) };
    const base = idx === 0 ? '#808080' : accent;
    if (m[1] === 'Light') {
      Object.assign(look, block === 1
        ? { headFill: base, headColor: '#FFFFFF', band: null, border: base }
        : block === 2 ? { headFill: null, headColor: null, band: tint(base, 0.8), border: base, grid: true }
          : { headFill: null, headColor: null, band: tint(base, 0.8), border: base, headLine: true });
    } else if (m[1] === 'Medium') {
      Object.assign(look, block === 0 ? { headFill: idx === 0 ? '#000000' : accent, headColor: '#FFFFFF', band: tint(base, 0.8), border: tint(base, 0.4) }
        : block === 1 ? { headFill: idx === 0 ? '#000000' : accent, headColor: '#FFFFFF', band: tint(base, 0.6), band2: tint(base, 0.8) }
          : block === 2 ? { headFill: idx === 0 ? '#000000' : accent, headColor: '#FFFFFF', band: '#D9D9D9', band2: '#EDEDED' }
            : { headFill: tint(base, 0.6), headColor: '#000000', band: tint(base, 0.8), border: base });
    } else {
      Object.assign(look, { headFill: '#000000', headColor: '#FFFFFF', band: tint(base, -0.25), band2: base, color: '#FFFFFF' });
    }
    t._look = look;
    return look;
  };

  // Fill/font the table style gives a cell (the cell's own formatting is drawn on top of this).
  P.tableStyleFor = function (r, c) {
    const t = (this.wb.tables || []).find((x) => x.sheet === this.sheet && r >= x.g.r1 && r <= x.g.r2 && c >= x.g.c1 && c <= x.g.c2);
    if (!t) return null;
    const look = this.tableLook(t);
    if (!look) return null;
    const out = {};
    if (t.header && r === t.g.r1) {
      if (look.headFill) out.fill = look.headFill;
      if (look.headColor) out.color = look.headColor;
      out.bold = true;
      if (look.border || look.headLine) out.bottom = look.border || look.headFill;
      return out;
    }
    if (t.totals && r === t.g.r2) { out.bold = true; out.top = look.border || look.headFill || '#000'; return out; }
    const i = r - t.g.r1 - t.header;
    if (look.rows && i % 2 === 0 && look.band) out.fill = look.band;
    else if (look.rows && look.band2) out.fill = look.band2;
    if (look.cols && (c - t.g.c1) % 2 === 0 && look.band) out.fill = look.band;
    if ((look.first && c === t.g.c1) || (look.last && c === t.g.c2)) out.bold = true;
    if (look.color) out.color = look.color;
    if (look.border) out.bottom = look.border;
    return out;
  };

  // ======================================================================
  // Drawings: pictures, shapes and charts anchored to cells
  // ======================================================================

  P.loadDrawings = async function () {
    const s = this.sheet;
    if (!this.drawingCache) this.drawingCache = new Map();
    for (const d of this.drawings || []) if (d.chart) d.chart.destroy();
    this.drawings = [];
    if (s.unsupported || s.isNew) { this.render(true); return; }
    const token = (this.drawingToken = (this.drawingToken || 0) + 1);
    const path = await this.wb.drawingPath(s);
    if (!path || token !== this.drawingToken) return;
    const doc = await this.wb.part(path);
    const rels = await this.wb.rels(path);
    if (!doc || token !== this.drawingToken) return;
    const anchors = [...doc.documentElement.childNodes].filter((x) => x.nodeType === 1 && /Anchor$/.test(x.localName));
    const out = [];
    for (const a of anchors) {
      const pos = parseAnchor(a);
      if (!pos) continue;
      const content = [...a.childNodes].find((x) => x.nodeType === 1 && ['pic', 'graphicFrame', 'sp', 'grpSp', 'cxnSp'].includes(x.localName));
      if (!content) continue;
      const item = { pos, el: document.createElement('div') };
      item.el.className = 'sg-drawing';
      if (content.localName === 'pic') {
        const blip = content.getElementsByTagNameNS(A, 'blip')[0];
        const rid = blip && (blip.getAttributeNS(RNS, 'embed') || blip.getAttribute('r:embed'));
        const rel = rid && rels.list.find((r) => r.id === rid);
        if (rel && rel.path) {
          const bytes = await this.wb.partBytes(rel.path);
          if (bytes) {
            const ext = rel.path.split('.').pop().toLowerCase();
            const type = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' }[ext] || 'image/png';
            const img = document.createElement('img');
            img.src = URL.createObjectURL(new Blob([bytes], { type }));
            img.draggable = false;
            item.el.append(img);
          }
        }
      } else if (content.localName === 'graphicFrame') {
        const ch = content.getElementsByTagNameNS(C, 'chart')[0];
        const rid = ch && (ch.getAttributeNS(RNS, 'id') || ch.getAttribute('r:id'));
        const rel = rid && rels.list.find((r) => r.id === rid);
        if (rel && rel.path) {
          const cdoc = await this.wb.part(rel.path);
          if (cdoc) { item.chartDoc = cdoc; item.el.classList.add('sg-chart'); }
        }
      } else if (content.localName === 'sp') {
        item.el.classList.add('sg-shape');
        const fill = content.getElementsByTagNameNS(A, 'solidFill')[0];
        const color = fill && this.drawingColor(fill);
        if (color) item.el.style.background = color;
        const ln = content.getElementsByTagNameNS(A, 'ln')[0];
        if (ln && !ln.getElementsByTagNameNS(A, 'noFill').length) item.el.style.border = `1px solid ${(ln.getElementsByTagNameNS(A, 'solidFill')[0] && this.drawingColor(ln.getElementsByTagNameNS(A, 'solidFill')[0])) || '#41719C'}`;
        const paras = [...content.getElementsByTagNameNS(A, 'p')].map((p) => [...p.getElementsByTagNameNS(A, 't')].map((t) => t.textContent).join(''));
        if (paras.some((t) => t)) { const t = document.createElement('div'); t.className = 'sg-shape-text'; t.textContent = paras.join('\n'); item.el.append(t); }
        if (!color && !paras.some((t) => t)) continue;
      } else continue;
      out.push(item);
    }
    if (token !== this.drawingToken) return;
    this.drawings = out;
    this.render(true);
    for (const d of out) if (d.chartDoc) this.drawChart(d);
  };

  function parseAnchor(a) {
    const n = (el, name) => { const x = el && [...el.childNodes].find((k) => k.localName === name); return x ? +x.textContent : 0; };
    const pt = (el) => ({ col: n(el, 'col'), colOff: n(el, 'colOff'), row: n(el, 'row'), rowOff: n(el, 'rowOff') });
    const child = (name) => [...a.childNodes].find((k) => k.localName === name);
    const ext = child('ext');
    if (a.localName === 'twoCellAnchor') return { from: pt(child('from')), to: pt(child('to')) };
    if (a.localName === 'oneCellAnchor') return { from: pt(child('from')), cx: ext ? +ext.getAttribute('cx') : 0, cy: ext ? +ext.getAttribute('cy') : 0 };
    if (a.localName === 'absoluteAnchor') { const p = child('pos'); return { x: p ? +p.getAttribute('x') : 0, y: p ? +p.getAttribute('y') : 0, cx: ext ? +ext.getAttribute('cx') : 0, cy: ext ? +ext.getAttribute('cy') : 0 }; }
    return null;
  }

  P.drawingRect = function (pos) {
    const X = (p) => this.colLeft[Math.min(p.col + 1, this.nCols + 1)] + p.colOff / EMU;
    const Y = (p) => this.rowTop[Math.min(p.row + 1, this.nRows + 1)] + p.rowOff / EMU;
    if (pos.to) { const x = X(pos.from), y = Y(pos.from); return { x, y, w: Math.max(4, X(pos.to) - x), h: Math.max(4, Y(pos.to) - y) }; }
    if (pos.from) return { x: X(pos.from), y: Y(pos.from), w: pos.cx / EMU, h: pos.cy / EMU };
    return { x: pos.x / EMU, y: pos.y / EMU, w: pos.cx / EMU, h: pos.cy / EMU };
  };

  P.attachDrawings = function (layer, offX, offY) {
    if (layer !== this.mainLayer || !this.drawings || !this.drawings.length) return;
    for (const d of this.drawings) {
      const r = this.drawingRect(d.pos);
      Object.assign(d.el.style, { left: `${r.x - offX}px`, top: `${r.y - offY}px`, width: `${r.w}px`, height: `${r.h}px` });
      layer.append(d.el);
      if (d.chart && (Math.abs(d.chart.width - r.w) > 1 || Math.abs(d.chart.height - r.h) > 1)) d.chart.resize(r.w, r.h);
    }
  };

  P.drawingColor = function (el) {
    const srgb = el.getElementsByTagNameNS(A, 'srgbClr')[0];
    const scheme = el.getElementsByTagNameNS(A, 'schemeClr')[0];
    let hex = srgb ? srgb.getAttribute('val') : null;
    const node = srgb || scheme;
    if (!hex && scheme) {
      const names = { lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11 };
      hex = this.wb.theme[names[scheme.getAttribute('val')]] || null;
    }
    if (!hex) return null;
    let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    const mod = (name) => { const m = node && [...node.childNodes].find((x) => x.localName === name); return m ? +m.getAttribute('val') / 100000 : null; };
    const lumMod = mod('lumMod'), lumOff = mod('lumOff'), shade = mod('shade'), tint = mod('tint');
    if (lumMod !== null || lumOff !== null) { const f = (v) => Math.min(255, Math.max(0, v * (lumMod ?? 1) + 255 * (lumOff ?? 0))); [r, g, b] = [r, g, b].map(f); }
    if (shade !== null) [r, g, b] = [r, g, b].map((v) => v * shade);
    if (tint !== null) [r, g, b] = [r, g, b].map((v) => v + (255 - v) * (1 - tint));
    return '#' + [r, g, b].map((v) => Math.round(v).toString(16).padStart(2, '0')).join('');
  };

  // Values for a chart reference like 'Sheet 1'!$B$2:$B$9 from the live workbook (falls back to the cache).
  P.refValues = function (f, numeric, visibleOnly) {
    if (!f) return null;
    const t = SheetModel.scanRefs(f).filter((x) => x.kind === 'ref');
    if (!t.length) return null;
    const sheet = this.wb.sheet(t[0].sheet || this.sheet.name);
    if (!sheet) return null;
    const a = t[0], b = t[1] || t[0];
    const out = [];
    for (let r = Math.min(a.row, b.row); r <= Math.max(a.row, b.row); r++) {
      if (visibleOnly && (sheet.rows.get(r) || {}).hidden) continue;
      for (let c = Math.min(a.col, b.col); c <= Math.max(a.col, b.col); c++) {
        if (visibleOnly && sheet.isColHidden(c)) continue;
        const cell = sheet.get(r, c);
        if (numeric) out.push(cell && typeof cell.value === 'number' ? cell.value : null);
        else out.push(SheetView.display(cell, this.wb.styles, this.wb.date1904).text);
      }
    }
    return out;
  };

  function cachePts(el) {
    if (!el) return null;
    const cache = [...el.getElementsByTagNameNS(C, 'numCache'), ...el.getElementsByTagNameNS(C, 'strCache')][0];
    if (!cache) {
      const lit = [...el.getElementsByTagNameNS(C, 'numLit'), ...el.getElementsByTagNameNS(C, 'strLit')][0];
      if (!lit) return null;
      return [...lit.getElementsByTagNameNS(C, 'pt')].map((p) => p.textContent);
    }
    const count = +([...cache.childNodes].find((x) => x.localName === 'ptCount') || { getAttribute: () => 0 }).getAttribute('val');
    const out = new Array(count).fill(null);
    for (const p of cache.getElementsByTagNameNS(C, 'pt')) out[+p.getAttribute('idx')] = p.textContent;
    return out;
  }

  P.chartSpec = function (doc) {
    const plot = doc.getElementsByTagNameNS(C, 'plotArea')[0];
    if (!plot) return null;
    const typesEls = [...plot.childNodes].filter((x) => x.nodeType === 1 && /Chart$/.test(x.localName));
    const accents = [4, 5, 6, 7, 8, 9].map((i) => '#' + (this.wb.theme[i] || '4472C4'));
    const datasets = [];
    let labels = null;
    let kind = null, horizontal = false, stacked = null;
    let si = 0;
    const kidsN = (el, n) => [...el.childNodes].filter((x) => x.localName === n);
    // Excel leaves hidden rows/columns out of charts unless "show data in hidden rows" is on.
    const pv = doc.getElementsByTagNameNS(C, 'plotVisOnly')[0];
    const vis = !pv || pv.getAttribute('val') !== '0';
    for (const t of typesEls) {
      const ln = t.localName.replace(/3D/, '');
      const type = { barChart: 'bar', lineChart: 'line', areaChart: 'line', pieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter', radarChart: 'radar', bubbleChart: 'scatter', ofPieChart: 'pie', stockChart: 'line' }[ln] || 'bar';
      if (!kind) kind = type;
      const dir = kidsN(t, 'barDir')[0];
      if (dir && dir.getAttribute('val') === 'bar') horizontal = true;
      const grouping = kidsN(t, 'grouping')[0];
      if (grouping && /stacked/i.test(grouping.getAttribute('val'))) stacked = grouping.getAttribute('val');
      const vary = kidsN(t, 'varyColors')[0];
      for (const ser of kidsN(t, 'ser')) {
        const tx = kidsN(ser, 'tx')[0];
        const nameRef = tx && tx.getElementsByTagNameNS(C, 'f')[0];
        const name = (nameRef && this.refValues(nameRef.textContent, false)?.[0]) || (tx && (cachePts(tx) || [])[0]) || (tx && tx.getElementsByTagNameNS(C, 'v')[0]?.textContent) || `Series ${si + 1}`;
        const cat = kidsN(ser, 'cat')[0] || kidsN(ser, 'xVal')[0];
        const val = kidsN(ser, 'val')[0] || kidsN(ser, 'yVal')[0];
        const catF = cat && cat.getElementsByTagNameNS(C, 'f')[0];
        const valF = val && val.getElementsByTagNameNS(C, 'f')[0];
        const vals = (valF && this.refValues(valF.textContent, true, vis)) || (cachePts(val) || []).map((x) => (x === null ? null : +x));
        const cats = (catF && this.refValues(catF.textContent, type === 'scatter', vis)) || cachePts(cat) || vals.map((_, i) => String(i + 1));
        if (!labels && type !== 'scatter') labels = cats;
        const spPr = kidsN(ser, 'spPr')[0];
        const solid = spPr && spPr.getElementsByTagNameNS(A, 'solidFill')[0];
        const color = (solid && this.drawingColor(solid)) || accents[si % accents.length];
        const ds = { label: name, data: type === 'scatter' ? vals.map((y, i) => ({ x: +cats[i], y })) : vals, backgroundColor: color, borderColor: color, borderWidth: type === 'line' || type === 'radar' ? 2 : 1, type: type === kind ? undefined : type };
        if (ln === 'areaChart') { ds.fill = true; ds.backgroundColor = color + '99'; }
        if (type === 'pie' || type === 'doughnut' || (vary && vary.getAttribute('val') === '1' && type === 'bar' && kidsN(t, 'ser').length === 1)) {
          ds.backgroundColor = vals.map((_, i) => {
            const dpt = kidsN(ser, 'dPt').find((d) => +kidsN(d, 'idx')[0]?.getAttribute('val') === i);
            const f2 = dpt && dpt.getElementsByTagNameNS(A, 'solidFill')[0];
            return (f2 && this.drawingColor(f2)) || accents[i % accents.length];
          });
          ds.borderColor = '#fff';
        }
        if (type === 'line' && kidsN(ser, 'marker')[0]?.getElementsByTagNameNS(C, 'symbol')[0]?.getAttribute('val') === 'none') ds.pointRadius = 0;
        datasets.push(ds);
        si++;
      }
    }
    const titleEl = doc.getElementsByTagNameNS(C, 'title')[0];
    const autoDeleted = doc.getElementsByTagNameNS(C, 'autoTitleDeleted')[0]?.getAttribute('val') === '1';
    let title = titleEl ? [...titleEl.getElementsByTagNameNS(A, 't')].map((t) => t.textContent).join('') : '';
    if (titleEl && !title) { const f = titleEl.getElementsByTagNameNS(C, 'f')[0]; if (f) title = this.refValues(f.textContent, false)?.[0] || ''; }
    if (!title && !titleEl && !autoDeleted && datasets.length === 1 && kind !== 'scatter') title = datasets[0].label;
    const legendEl = doc.getElementsByTagNameNS(C, 'legend')[0];
    const lp = legendEl && legendEl.getElementsByTagNameNS(C, 'legendPos')[0];
    const legend = legendEl ? ({ b: 'bottom', t: 'top', l: 'left', r: 'right', tr: 'right' }[lp ? lp.getAttribute('val') : 'r'] || 'right') : null;
    const holeEl = doc.getElementsByTagNameNS(C, 'holeSize')[0];
    return { kind: kind || 'bar', horizontal, stacked, labels: labels || [], datasets, title, legend, hole: holeEl ? +holeEl.getAttribute('val') : 50 };
  };

  P.chartConfig = function (spec) {
    const pie = spec.kind === 'pie' || spec.kind === 'doughnut';
    const stacked = !!spec.stacked;
    const scales = pie || spec.kind === 'radar' ? undefined : {
      x: { stacked, grid: { color: '#e5e5e5' }, ticks: { color: '#595959', font: { size: 11 } }, ...(spec.kind === 'scatter' ? { type: 'linear' } : {}) },
      y: { stacked, grid: { color: '#e5e5e5' }, ticks: { color: '#595959', font: { size: 11 } }, beginAtZero: true, max: spec.stacked === 'percentStacked' ? undefined : undefined },
    };
    return {
      type: spec.kind,
      data: { labels: spec.labels, datasets: spec.datasets },
      options: {
        animation: false, responsive: false, maintainAspectRatio: false, indexAxis: spec.horizontal ? 'y' : 'x',
        cutout: spec.kind === 'doughnut' ? `${spec.hole}%` : undefined,
        plugins: {
          title: { display: !!spec.title, text: spec.title, color: '#404040', font: { size: 14, weight: 'normal' } },
          legend: { display: !!spec.legend, position: spec.legend || 'right', labels: { color: '#404040', boxWidth: 12, font: { size: 11 } } },
          tooltip: { enabled: true },
        },
        scales,
      },
    };
  };

  P.drawChart = function (d) {
    const spec = this.chartSpec(d.chartDoc);
    if (!spec || typeof Chart === 'undefined') return;
    const r = this.drawingRect(d.pos);
    d.el.replaceChildren();
    const canvas = document.createElement('canvas');
    canvas.width = r.w; canvas.height = r.h;
    d.el.append(canvas);
    try { d.chart = new Chart(canvas, this.chartConfig(spec)); } catch (err) { console.error(err); d.el.textContent = 'Chart'; }
  };

  // Charts follow edits to the cells they plot.
  P.scheduleChartRefresh = function () {
    if (!this.drawings || !this.drawings.some((d) => d.chart)) return;
    clearTimeout(this.chartTimer);
    this.chartTimer = setTimeout(() => {
      for (const d of this.drawings) {
        if (!d.chart) continue;
        const spec = this.chartSpec(d.chartDoc);
        if (!spec) continue;
        d.chart.data.labels = spec.labels;
        d.chart.data.datasets = spec.datasets;
        d.chart.update('none');
      }
    }, 150);
  };

  // Load drawings whenever a sheet is shown.
  const showSheet = P.showSheet;
  P.showSheet = function (i) {
    const r = showSheet.call(this, i);
    this.cfInvalidate();
    this.loadDrawings();
    return r;
  };
})();
