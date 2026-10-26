'use strict';
// Spreadsheet editor UI for a SheetModel workbook: virtualised grid with the file's own formatting, Excel-style
// selection and editing, formula bar, formatting toolbar, undo, copy/paste and sheet tabs.

const SheetView = (() => {
  const { key, addr, colName, parseAddr, parseRange, rangeRef, XErr, shiftFormula, scanRefs, quoteSheet } = SheetModel;
  const HEADER_W = 46;
  const HEADER_H = 22;
  const BUFFER_ROWS = 12;
  const BUFFER_COLS = 4;
  const ACCENT = '#217346';

  const colPx = (w) => (w == null ? 64 : w === 0 ? 0 : Math.trunc(((256 * w + 18) / 256) * 7));
  const pxToWidth = (px) => { const chars = Math.max(0, Math.trunc(((px - 5) / 7) * 100 + 0.5) / 100); return Math.trunc(((chars * 7 + 5) / 7) * 256) / 256; };
  const rowPx = (pt) => Math.round(pt * 4 / 3);

  const BORDER_CSS = {
    thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', hair: '1px dotted',
    double: '3px double', mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dotted', mediumDashDotDot: '2px dotted', slantDashDot: '2px dashed',
  };
  const FORMAT_COLORS = { red: '#FF0000', blue: '#0000FF', green: '#00FF00', yellow: '#FFFF00', magenta: '#FF00FF', cyan: '#00FFFF', black: '#000000', white: '#FFFFFF' };

  // ---------- value display

  function formatSections(fmt) {
    const out = []; let cur = ''; let q = false; let b = false;
    for (const ch of fmt) {
      if (ch === '"') q = !q;
      if (!q && ch === '[') b = true;
      if (!q && ch === ']') b = false;
      if (ch === ';' && !q && !b) { out.push(cur); cur = ''; } else cur += ch;
    }
    out.push(cur);
    return out;
  }

  function display(cell, styles, date1904, fmtOverride) {
    if (!cell) return { text: '', kind: 'empty' };
    const v = cell.value;
    if (v === null || v === undefined || v === '') return { text: '', kind: 'empty' };
    if (v instanceof XErr) return { text: v.code, kind: 'error' };
    if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', kind: 'bool' };
    const fmt = fmtOverride || styles.numFmt(cell.s);
    if (typeof v === 'string') {
      if (fmt.includes('@') && fmt !== '@') { try { return { text: XLSX.SSF.format(fmt, v), kind: 'text' }; } catch { /* fall through */ } }
      return { text: v, kind: 'text' };
    }
    let text;
    try { text = XLSX.SSF.format(fmt, v, { date1904 }); } catch { text = String(v); }
    let color = null;
    const secs = formatSections(fmt);
    const sec = secs.length > 1 && v < 0 ? secs[1] : secs.length > 2 && v === 0 ? secs[2] : secs[0];
    const m = /\[(red|blue|green|yellow|magenta|cyan|black|white)\]/i.exec(sec || '');
    if (m) color = FORMAT_COLORS[m[1].toLowerCase()];
    return { text, kind: 'number', color };
  }

  // What the formula bar shows for a cell.
  function rawText(cell, styles, date1904) {
    if (!cell) return '';
    if (cell.formula) return '=' + cell.formula;
    const v = cell.value;
    if (v === null || v === undefined) return '';
    if (v instanceof XErr) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') {
      const fmt = styles.numFmt(cell.s);
      if (XLSX.SSF.is_date && XLSX.SSF.is_date(fmt)) {
        const hasDate = /[dmy]/i.test(fmt.replace(/\[[^\]]*\]/g, '').replace(/"[^"]*"/g, '')) && !/^\[?h/i.test(fmt);
        const hasTime = /[hs]/i.test(fmt.replace(/"[^"]*"/g, ''));
        const f = hasDate && hasTime ? 'm/d/yyyy h:mm:ss AM/PM' : hasTime && !hasDate ? 'h:mm:ss AM/PM' : 'm/d/yyyy';
        try { return XLSX.SSF.format(f, v, { date1904 }); } catch { return String(v); }
      }
      if (/%/.test(fmt) && !/"[^"]*%[^"]*"/.test(fmt)) return `${+(v * 100).toPrecision(15)}%`;
      return String(+v.toPrecision(15));
    }
    return String(v);
  }

  // ---------- typed input -> value (+ a number format when the cell is General)

  const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
  function serial(y, m, d, hh, mi, ss, date1904) {
    const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
    return (Date.UTC(y, m - 1, d, hh || 0, mi || 0, ss || 0) - epoch) / 86400000;
  }
  function parseTime(s) {
    const m = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([ap]\.?m\.?)?$/i.exec(s.trim());
    if (!m) return null;
    let h = +m[1];
    if (m[4]) { const pm = /^p/i.test(m[4]); if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12; }
    if (h > 23 || +m[2] > 59) return null;
    return { h, mi: +m[2], s: m[3] ? +m[3] : 0, secs: !!m[3] };
  }
  function parseDate(s) {
    s = s.trim();
    let m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s))) return { y: +m[1], m: +m[2], d: +m[3] };
    if ((m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(s))) { let y = +m[3]; if (y < 100) y += y < 30 ? 2000 : 1900; return { y, m: +m[1], d: +m[2] }; }
    if ((m = /^(\d{1,2})\/(\d{1,2})$/.exec(s))) return { y: new Date().getFullYear(), m: +m[1], d: +m[2] };
    if ((m = /^(\d{1,2})[- ]([A-Za-z]{3,9})[- ,]*(\d{2,4})?$/.exec(s))) { const mo = MONTHS.indexOf(m[2].slice(0, 3).toLowerCase()); if (mo >= 0) { let y = m[3] ? +m[3] : new Date().getFullYear(); if (y < 100) y += 2000; return { y, m: mo + 1, d: +m[1] }; } }
    if ((m = /^([A-Za-z]{3,9})\.? (\d{1,2}),? (\d{4})$/.exec(s))) { const mo = MONTHS.indexOf(m[1].slice(0, 3).toLowerCase()); if (mo >= 0) return { y: +m[3], m: mo + 1, d: +m[2] }; }
    return null;
  }

  function parseInput(text, fmtIsGeneral, date1904) {
    if (text === '') return { value: null };
    if (text.startsWith('=') && text.length > 1) return { formula: text.slice(1) };
    if (text.startsWith("'")) return { value: text.slice(1) };
    const t = text.trim();
    if (/^(true|false)$/i.test(t)) return { value: /^true$/i.test(t) };
    if (/^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)$/i.test(t)) return { value: new XErr(t.toUpperCase()) };
    let m = /^([+-])?\(?([$€£¥])?\s*([+-])?((?:\d{1,3}(?:,\d{3})+|\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?)\s*(%)?\)?$/.exec(t);
    if (m && m[4] && /\d/.test(m[4])) {
      let n = parseFloat(m[4].replace(/,/g, ''));
      if (m[1] === '-' || m[3] === '-' || (/^\(.*\)$/.test(t))) n = -n;
      let fmt = null;
      if (m[5]) { n /= 100; fmt = m[4].includes('.') ? '0.00%' : '0%'; }
      else if (m[2]) fmt = `${m[2] === '$' ? '\\$' : m[2]}#,##0${m[4].includes('.') ? '.00' : ''}`;
      else if (m[4].includes(',')) fmt = m[4].includes('.') ? '#,##0.00' : '#,##0';
      return { value: n, numFmt: fmtIsGeneral ? fmt : null };
    }
    // Dates and times.
    const parts = t.split(/\s+(?=\d{1,2}:)/);
    const d = parseDate(parts[0]);
    const tm = parts[1] ? parseTime(parts[1]) : (!d ? parseTime(t) : null);
    if (d && d.m >= 1 && d.m <= 12 && d.d >= 1 && d.d <= 31) {
      const v = serial(d.y, d.m, d.d, tm && tm.h, tm && tm.mi, tm && tm.s, date1904);
      return { value: v, numFmt: fmtIsGeneral ? (tm ? 'm/d/yyyy h:mm' : 'm/d/yyyy') : null };
    }
    if (tm && !d) return { value: (tm.h * 3600 + tm.mi * 60 + tm.s) / 86400, numFmt: fmtIsGeneral ? (tm.secs ? 'h:mm:ss' : 'h:mm') : null };
    return { value: text };
  }

  // ======================================================================

  class View {
    constructor(host, wb, opts) {
      this.host = host;
      this.wb = wb;
      this.opts = opts || {};
      this.calc = new SheetCalc.Calc(wb);
      this.undoStack = [];
      this.redoStack = [];
      this.sheetIndex = Math.max(0, wb.sheets.findIndex((s) => !s.unsupported && s.state === 'visible'));
      this.sel = { r1: 1, c1: 1, r2: 1, c2: 1 };
      this.active = { r: 1, c: 1 };
      this.anchor = { r: 1, c: 1 };
      this.measure = document.createElement('canvas').getContext('2d');
      this.buildDom();
      const initial = this.calc.initial();
      if (initial.length) this.opts.onInitialCalc && this.opts.onInitialCalc(initial.length);
      this.showSheet(this.sheetIndex);
    }

    get sheet() { return this.wb.sheets[this.sheetIndex]; }

    // ---------- DOM

    buildDom() {
      const el = (tag, cls, attrs) => { const e = document.createElement(tag); if (cls) e.className = cls; for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v); return e; };
      this.root = el('div', 'sheet-app');
      // Formula bar
      this.nameBox = el('input', 'sheet-namebox', { 'aria-label': 'Cell address', spellcheck: 'false' });
      this.fInput = el('textarea', 'sheet-finput', { 'aria-label': 'Formula bar', spellcheck: 'false', rows: '1' });
      const fx = el('span', 'sheet-fx'); fx.textContent = 'fx';
      this.fbar = el('div', 'sheet-fbar'); this.fbar.append(this.nameBox, fx, this.fInput);
      // Grid
      this.grid = el('div', 'sheet-grid2', { tabindex: '0', role: 'grid', 'aria-label': 'Spreadsheet' });
      this.main = el('div', 'sg-main');
      this.spacer = el('div', 'sg-spacer');
      this.mainLayer = el('div', 'sg-layer');
      this.main.append(this.spacer, this.mainLayer);
      this.top = el('div', 'sg-pane sg-top'); this.topLayer = el('div', 'sg-layer'); this.top.append(this.topLayer);
      this.left = el('div', 'sg-pane sg-left'); this.leftLayer = el('div', 'sg-layer'); this.left.append(this.leftLayer);
      this.corner = el('div', 'sg-pane sg-corner'); this.cornerLayer = el('div', 'sg-layer'); this.corner.append(this.cornerLayer);
      this.colHdr = el('div', 'sg-hdr sg-colhdr'); this.colHdrFrozen = el('div', 'sg-hdr-frozen'); this.colHdrMain = el('div', 'sg-hdr-main'); this.colHdrInner = el('div', 'sg-layer');
      this.colHdrMain.append(this.colHdrInner); this.colHdr.append(this.colHdrFrozen, this.colHdrMain);
      this.rowHdr = el('div', 'sg-hdr sg-rowhdr'); this.rowHdrFrozen = el('div', 'sg-hdr-frozen'); this.rowHdrMain = el('div', 'sg-hdr-main'); this.rowHdrInner = el('div', 'sg-layer');
      this.rowHdrMain.append(this.rowHdrInner); this.rowHdr.append(this.rowHdrFrozen, this.rowHdrMain);
      this.selAll = el('div', 'sg-selall', { title: 'Select all' });
      this.editor = el('textarea', 'sg-editor', { spellcheck: 'false' });
      this.editor.hidden = true;
      this.grid.append(this.main, this.top, this.left, this.corner, this.colHdr, this.rowHdr, this.selAll, this.editor);
      // Sheet tabs
      this.tabsBar = el('div', 'sheet-tabs2');
      this.root.append(this.fbar, this.grid, this.tabsBar);
      this.host.replaceChildren(this.root);

      this.main.addEventListener('scroll', () => this.onScroll());
      this.grid.addEventListener('mousedown', (e) => this.onMouseDown(e));
      this.grid.addEventListener('dblclick', (e) => this.onDblClick(e));
      this.grid.addEventListener('keydown', (e) => this.onKey(e));
      // Clipboard events go to the document when nothing is selected, so listen there and act when the grid has focus.
      const mine = () => this.root.isConnected && document.activeElement === this.grid;
      this.clipHandlers = {
        copy: (e) => { if (mine()) this.onCopy(e); },
        cut: (e) => { if (mine()) { this.onCopy(e); if (this.clip) this.clip.cut = true; } },
        paste: (e) => { if (mine()) this.onPaste(e); },
      };
      for (const [ev, fn] of Object.entries(this.clipHandlers)) document.addEventListener(ev, fn);
      this.grid.addEventListener('mousemove', (e) => this.onHover(e));
      this.grid.addEventListener('contextmenu', (e) => { e.preventDefault(); if (this.onContextMenu) this.onContextMenu(e); });
      this.editor.addEventListener('keydown', (e) => this.onEditorKey(e));
      this.editor.addEventListener('input', () => { this.fInput.value = this.editor.value; this.pendingRef = null; this.autoSizeEditor(); });
      this.fInput.addEventListener('focus', () => { if (!this.editing) this.beginEdit(null, 'bar'); });
      this.fInput.addEventListener('input', () => { if (this.editing) { this.editor.value = this.fInput.value; this.pendingRef = null; } });
      this.fInput.addEventListener('keydown', (e) => this.onEditorKey(e, true));
      this.nameBox.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); this.goTo(this.nameBox.value); this.grid.focus(); }
        if (e.key === 'Escape') { this.updateBars(); this.grid.focus(); }
      });
      this.resizeObs = new ResizeObserver(() => this.render(true));
      this.resizeObs.observe(this.grid);
    }

    // ---------- geometry

    layout() {
      const s = this.sheet;
      this.F = { rows: Math.min(s.frozen.rows || 0, 200), cols: Math.min(s.frozen.cols || 0, 50) };
      const vh = this.main.clientHeight || 600;
      const vw = this.main.clientWidth || 900;
      const needRows = Math.max(s.maxRow + 60, Math.ceil(((this.main.scrollTop || 0) + vh * 2) / 20), 100);
      const needCols = Math.max(s.maxCol + 12, 26, Math.ceil(((this.main.scrollLeft || 0) + vw * 2) / 64));
      if (this.layoutFor === s && this.nRows >= needRows && this.nCols >= needCols && !this.layoutDirty) return;
      this.nRows = Math.min(needRows + 200, SheetModel.MAX_ROW);
      this.nCols = Math.min(needCols + 10, SheetModel.MAX_COL);
      this.rowTop = new Float64Array(this.nRows + 2);
      for (let r = 1; r <= this.nRows; r++) this.rowTop[r + 1] = this.rowTop[r] + rowPx(s.rowHeightPt(r));
      this.colLeft = new Float64Array(this.nCols + 2);
      this.colLeft[1] = 0;
      for (let c = 1; c <= this.nCols; c++) this.colLeft[c + 1] = this.colLeft[c] + colPx(s.colWidthChars(c));
      this.frozenH = this.rowTop[this.F.rows + 1];
      this.frozenW = this.colLeft[this.F.cols + 1];
      this.layoutFor = s;
      this.layoutDirty = false;
      const W = HEADER_W, H = HEADER_H;
      Object.assign(this.main.style, { top: `${H + this.frozenH}px`, left: `${W + this.frozenW}px` });
      Object.assign(this.top.style, { top: `${H}px`, left: `${W + this.frozenW}px`, height: `${this.frozenH}px`, display: this.F.rows ? '' : 'none' });
      Object.assign(this.left.style, { top: `${H + this.frozenH}px`, left: `${W}px`, width: `${this.frozenW}px`, display: this.F.cols ? '' : 'none' });
      Object.assign(this.corner.style, { top: `${H}px`, left: `${W}px`, width: `${this.frozenW}px`, height: `${this.frozenH}px`, display: this.F.rows && this.F.cols ? '' : 'none' });
      Object.assign(this.colHdrFrozen.style, { width: `${this.frozenW}px` });
      Object.assign(this.colHdrMain.style, { left: `${this.frozenW}px` });
      Object.assign(this.rowHdrFrozen.style, { height: `${this.frozenH}px` });
      Object.assign(this.rowHdrMain.style, { top: `${this.frozenH}px` });
      this.spacer.style.width = `${this.colLeft[this.nCols + 1] - this.frozenW}px`;
      this.spacer.style.height = `${this.rowTop[this.nRows + 1] - this.frozenH}px`;
      this.rendered = null;
    }

    rowAt(y) { // y in sheet coordinates
      let lo = 1, hi = this.nRows;
      while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (this.rowTop[mid] <= y) lo = mid; else hi = mid - 1; }
      return lo;
    }
    colAt(x) {
      let lo = 1, hi = this.nCols;
      while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (this.colLeft[mid] <= x) lo = mid; else hi = mid - 1; }
      return lo;
    }

    // ---------- rendering

    onScroll() {
      if (this.scrollRaf) return;
      this.scrollRaf = requestAnimationFrame(() => { this.scrollRaf = null; this.render(); });
    }

    render(force) {
      if (!this.sheet || this.sheet.unsupported) return;
      this.layout();
      const st = this.main.scrollTop, sl = this.main.scrollLeft;
      const vh = this.main.clientHeight, vw = this.main.clientWidth;
      const yA = st + this.frozenH, yB = st + this.frozenH + vh;
      const xA = sl + this.frozenW, xB = sl + this.frozenW + vw;
      const r1 = Math.max(this.F.rows + 1, this.rowAt(yA)), r2 = Math.min(this.nRows, this.rowAt(yB));
      const c1 = Math.max(this.F.cols + 1, this.colAt(xA)), c2 = Math.min(this.nCols, this.colAt(xB));
      const rd = this.rendered;
      const inside = rd && r1 >= rd.r1 && r2 <= rd.r2 && c1 >= rd.c1 && c2 <= rd.c2;
      if (force || !inside) {
        const R1 = Math.max(this.F.rows + 1, r1 - BUFFER_ROWS), R2 = Math.min(this.nRows, r2 + BUFFER_ROWS);
        const C1 = Math.max(this.F.cols + 1, c1 - BUFFER_COLS), C2 = Math.min(this.nCols, c2 + BUFFER_COLS);
        this.rendered = { r1: R1, r2: R2, c1: C1, c2: C2 };
        this.paintRegion(this.mainLayer, R1, R2, C1, C2, this.frozenW, this.frozenH);
        if (this.F.rows) this.paintRegion(this.topLayer, 1, this.F.rows, C1, C2, this.frozenW, 0);
        else this.topLayer.replaceChildren();
        if (this.F.cols) this.paintRegion(this.leftLayer, R1, R2, 1, this.F.cols, 0, this.frozenH);
        else this.leftLayer.replaceChildren();
        if (this.F.rows && this.F.cols) this.paintRegion(this.cornerLayer, 1, this.F.rows, 1, this.F.cols, 0, 0);
        else this.cornerLayer.replaceChildren();
        this.paintHeaders(R1, R2, C1, C2);
      }
      this.topLayer.style.transform = `translateX(${-sl}px)`;
      this.leftLayer.style.transform = `translateY(${-st}px)`;
      this.colHdrInner.style.transform = `translateX(${-sl}px)`;
      this.rowHdrInner.style.transform = `translateY(${-st}px)`;
      this.paintSelection();
      if (this.editing && this.editing.where === 'cell') this.positionEditor();
    }

    paintHeaders(R1, R2, C1, C2) {
      const sel = this.sel;
      const mk = (text, x, y, w, h, on) => {
        const d = document.createElement('div');
        d.className = `sg-h${on ? ' on' : ''}`;
        d.style.cssText = `left:${x}px;top:${y}px;width:${w}px;height:${h}px`;
        d.textContent = text;
        return d;
      };
      const cols = [], colsF = [], rows = [], rowsF = [];
      for (let c = 1; c <= this.F.cols; c++) colsF.push(mk(colName(c), this.colLeft[c], 0, this.colLeft[c + 1] - this.colLeft[c], HEADER_H, c >= sel.c1 && c <= sel.c2));
      for (let c = C1; c <= C2; c++) { if (this.colLeft[c + 1] === this.colLeft[c]) continue; const d = mk(colName(c), this.colLeft[c] - this.frozenW, 0, this.colLeft[c + 1] - this.colLeft[c], HEADER_H, c >= sel.c1 && c <= sel.c2); d.dataset.col = c; cols.push(d); }
      for (let r = 1; r <= this.F.rows; r++) rowsF.push(mk(r, 0, this.rowTop[r], HEADER_W, this.rowTop[r + 1] - this.rowTop[r], r >= sel.r1 && r <= sel.r2));
      for (let r = R1; r <= R2; r++) if (this.rowTop[r + 1] > this.rowTop[r]) rows.push(mk(r, 0, this.rowTop[r] - this.frozenH, HEADER_W, this.rowTop[r + 1] - this.rowTop[r], r >= sel.r1 && r <= sel.r2));
      this.colHdrFrozen.replaceChildren(...colsF);
      this.colHdrInner.replaceChildren(...cols);
      this.rowHdrFrozen.replaceChildren(...rowsF);
      this.rowHdrInner.replaceChildren(...rows);
      this.hdrRange = { R1, R2, C1, C2 };
    }

    paintRegion(layer, R1, R2, C1, C2, offX, offY) {
      const s = this.sheet;
      const styles = this.wb.styles;
      const frag = document.createDocumentFragment();
      const grid = s.showGrid;
      const merges = s.merges.filter((m) => m.r2 >= R1 && m.r1 <= R2 && m.c2 >= C1 && m.c1 <= C2);
      const covered = new Set();
      for (const m of merges) for (let r = Math.max(m.r1, R1); r <= Math.min(m.r2, R2); r++) for (let c = Math.max(m.c1, C1); c <= Math.min(m.c2, C2); c++) covered.add(key(r, c));
      const styleAt = (r, c) => { const cell = s.get(r, c); return cell ? cell.s : s.defaultStyleAt(r, c); };
      const borderCss = (b) => (b ? `${BORDER_CSS[b.style] || '1px solid'} ${b.color || '#000'}` : null);
      const filled = (r, c) => { const cell = s.get(r, c); return !!cell && cell.value !== null && cell.value !== undefined && cell.value !== ''; };

      const drawCell = (r, c, w, h, mergeTo) => {
        const cell = s.get(r, c);
        const si = cell ? cell.s : s.defaultStyleAt(r, c);
        const css = styles.css(si);
        let cf = this.cfFor ? this.cfFor(r, c, cell) : null;
        // Table styles sit under the cell's own formatting; conditional formats sit on top.
        const ts = this.tableStyleFor ? this.tableStyleFor(r, c) : null;
        if (ts) {
          const x = this.wb.styles.xf(si);
          const own = x.fillId > 1 ? {} : { fill: ts.fill };
          const fontOwn = x.fontId && css.font_.color && css.font_.color !== '#000000' ? {} : { color: ts.color };
          cf = { ...own, ...fontOwn, bold: ts.bold || undefined, ...(cf || {}) };
        }
        const d = document.createElement('div');
        d.className = 'sc';
        let width = w;
        let disp = display(cell, styles, this.wb.date1904, cf && cf.numFmt);
        // Extras drawn by sheet-tools: checkboxes, dropdown arrows, note markers, formulas shown as text.
        if (this.decorateCell) disp = this.decorateCell(r, c, cell, disp, d) || disp;
        // Borders: each cell draws its right and bottom edges (merged with the neighbour's left/top), plus its
        // left/top edges at the first column/row, so shared edges are drawn once.
        const rightN = mergeTo ? mergeTo.c2 + 1 : c + 1;
        const bottomN = mergeTo ? mergeTo.r2 + 1 : r + 1;
        const right = css.border.right || styles.css(styleAt(r, rightN)).border.left;
        const bottom = css.border.bottom || styles.css(styleAt(bottomN, c)).border.top;
        let st = `left:${this.colLeft[c] - offX}px;top:${this.rowTop[r] - offY}px;height:${h}px;`;
        const fillColor = (cf && cf.fill) || css.fill;
        if (fillColor) st += `background:${fillColor};`;
        st += right ? `border-right:${borderCss(right)};` : grid ? 'border-right:1px solid #e1e1e1;' : '';
        const tsBottom = ts && ts.bottom && !css.border.bottom ? `1px solid ${ts.bottom}` : null;
        st += bottom ? `border-bottom:${borderCss(bottom)};` : tsBottom ? `border-bottom:${tsBottom};` : grid ? 'border-bottom:1px solid #e1e1e1;' : '';
        if (ts && ts.top && !css.border.top) st += `border-top:2px double ${ts.top};`;
        if (css.border.left && (c === 1 || !styles.css(styleAt(r, c - 1)).border.right)) st += `border-left:${borderCss(css.border.left)};`;
        if (css.border.top && (r === 1 || !styles.css(styleAt(r - 1, c)).border.bottom)) st += `border-top:${borderCss(css.border.top)};`;
        if (disp.text) {
          let h2 = css.h;
          if (h2 === 'general') h2 = disp.kind === 'number' ? 'right' : disp.kind === 'bool' || disp.kind === 'error' ? 'center' : 'left';
          if (h2 === 'centerContinuous') h2 = 'center';
          if (h2 === 'fill' || h2 === 'distributed') h2 = h2 === 'fill' ? 'left' : 'justify';
          let text = disp.text;
          // Text spills into empty neighbours to the right, like Excel.
          if (!css.wrap && disp.kind === 'text' && (h2 === 'left') && !mergeTo) {
            this.measure.font = css.font;
            const need = this.measure.measureText(text).width + 6 + css.indent * 9;
            let cc = c + 1;
            while (width < need && cc <= this.nCols && cc <= c + 40 && !filled(r, cc) && !s.mergeAt(r, cc)) { width += this.colLeft[cc + 1] - this.colLeft[cc]; cc++; }
            if (width > w) st += `z-index:1;${fillColor ? '' : 'background:#fff;'}`;
          }
          let squeeze = 1;
          if (disp.kind === 'number' && !css.wrap) {
            // Browsers draw Calibri a little wider than Excel (which snaps glyphs to whole pixels), so allow a
            // small tolerance and squeeze slightly before falling back to ####.
            this.measure.font = css.font;
            const tw = this.measure.measureText(text).width;
            const room = width - 6; // padding + border
            if (tw > room * 1.12) text = '#'.repeat(Math.max(1, Math.floor(room / Math.max(1, this.measure.measureText('#').width))));
            else if (tw > room) squeeze = room / tw;
          }
          let font = css.font;
          if (cf && (cf.bold !== undefined || cf.italic !== undefined)) {
            const f = css.font_;
            font = `${(cf.italic ?? f.italic) ? 'italic ' : ''}${(cf.bold ?? f.bold) ? '700 ' : ''}${(f.size * 4 / 3).toFixed(2)}px "${f.name}", Calibri, Carlito, sans-serif`;
          }
          st += `font:${font};`;
          const color = (cf && cf.color) || disp.color || css.color;
          if (color) st += `color:${color};`;
          const deco = cf && (cf.underline !== undefined || cf.strike !== undefined)
            ? [(cf.underline ?? css.font_.underline) ? 'underline' : '', (cf.strike ?? css.font_.strike) ? 'line-through' : ''].filter(Boolean).join(' ')
            : css.decoration;
          if (deco) st += `text-decoration:${deco};`;
          st += `justify-content:${h2 === 'right' ? 'flex-end' : h2 === 'center' ? 'center' : 'flex-start'};text-align:${h2};`;
          st += `align-items:${css.v === 'top' ? 'flex-start' : css.v === 'center' ? 'center' : 'flex-end'};`;
          if (css.wrap) st += 'white-space:pre-wrap;';
          if (css.indent) st += `padding-${h2 === 'right' ? 'right' : 'left'}:${css.indent * 9 + 2}px;`;
          const span = document.createElement('span');
          if (squeeze < 1) { span.style.transform = `scaleX(${squeeze.toFixed(3)})`; span.style.transformOrigin = h2 === 'right' ? 'right' : h2 === 'center' ? 'center' : 'left'; span.style.overflow = 'visible'; }
          if (css.vert === 'superscript') span.style.verticalAlign = 'super';
          if (css.vert === 'subscript') span.style.verticalAlign = 'sub';
          span.textContent = text;
          // Text rotation: 1-90 = counter-clockwise, 91-180 = clockwise (n - 90), 255 = stacked letters.
          if (css.rotation) {
            if (css.rotation === 255) { span.style.writingMode = 'vertical-rl'; span.style.textOrientation = 'upright'; }
            else { span.style.display = 'inline-block'; span.style.whiteSpace = 'nowrap'; span.style.transform = `rotate(${css.rotation <= 90 ? -css.rotation : css.rotation - 90}deg)`; }
          }
          if (cf && cf.icon) { const ic = document.createElement('i'); ic.className = 'cf-icon'; ic.textContent = cf.icon.char; ic.style.color = cf.icon.color; d.append(ic); }
          d.append(span);
          if (disp.kind === 'error') d.title = text;
        }
        if (cf && cf.bar) {
          const bar = document.createElement('i');
          bar.className = 'cf-bar';
          bar.style.cssText = `width:${Math.max(0, Math.min(100, cf.bar.pct)).toFixed(1)}%;background:linear-gradient(90deg, ${cf.bar.color}, ${cf.bar.color}33)`;
          d.prepend(bar);
        }
        if (this.filterButtonAt && this.filterButtonAt(r, c)) {
          const fb = document.createElement('b');
          fb.className = `sg-filter-btn${this.filterActiveAt(r, c) ? ' on' : ''}`;
          fb.dataset.r = r; fb.dataset.c = c;
          fb.title = 'Sort and filter';
          d.append(fb);
        }
        st += `width:${width}px;`;
        d.style.cssText = st;
        frag.append(d);
      };

      for (let r = R1; r <= R2; r++) {
        const h = this.rowTop[r + 1] - this.rowTop[r];
        if (!h) continue;
        for (let c = C1; c <= C2; c++) {
          if (covered.has(key(r, c))) continue;
          const w = this.colLeft[c + 1] - this.colLeft[c];
          if (!w) continue;
          drawCell(r, c, w, h, null);
        }
      }
      for (const m of merges) {
        if (m.c1 > this.nCols || m.r1 > this.nRows) continue;
        const w = this.colLeft[Math.min(m.c2, this.nCols) + 1] - this.colLeft[m.c1];
        const h = this.rowTop[Math.min(m.r2, this.nRows) + 1] - this.rowTop[m.r1];
        drawCell(m.r1, m.c1, w, h, m);
      }
      layer.replaceChildren(frag);
      if (this.attachDrawings) this.attachDrawings(layer, offX, offY);
    }

    // Selection and active-cell boxes, drawn in each pane the selection touches.
    paintSelection() {
      for (const box of this.selBoxes || []) box.remove();
      this.selBoxes = [];
      const g = this.expandMerges(this.sel);
      const a = this.sheet.mergeAt(this.active.r, this.active.c) || { r1: this.active.r, c1: this.active.c, r2: this.active.r, c2: this.active.c };
      const panes = [
        [this.mainLayer, this.frozenW, this.frozenH, this.F.rows + 1, Infinity, this.F.cols + 1, Infinity],
        [this.topLayer, this.frozenW, 0, 1, this.F.rows, this.F.cols + 1, Infinity],
        [this.leftLayer, 0, this.frozenH, this.F.rows + 1, Infinity, 1, this.F.cols],
        [this.cornerLayer, 0, 0, 1, this.F.rows, 1, this.F.cols],
      ];
      const lastR = Math.min(g.r2, this.nRows), lastC = Math.min(g.c2, this.nCols);
      for (const [layer, ox, oy, pr1, pr2, pc1, pc2] of panes) {
        if (pr2 < pr1 || pc2 < pc1) continue;
        const add = (rg, cls) => {
          const r1 = Math.max(rg.r1, pr1), r2 = Math.min(rg.r2, pr2, this.nRows), c1 = Math.max(rg.c1, pc1), c2 = Math.min(rg.c2, pc2, this.nCols);
          if (r1 > r2 || c1 > c2) return;
          const d = document.createElement('div');
          d.className = cls;
          d.style.cssText = `left:${this.colLeft[c1] - ox - 1}px;top:${this.rowTop[r1] - oy - 1}px;width:${this.colLeft[c2 + 1] - this.colLeft[c1] + 1}px;height:${this.rowTop[r2 + 1] - this.rowTop[r1] + 1}px`;
          layer.append(d);
          this.selBoxes.push(d);
        };
        if (g.r1 !== g.r2 || g.c1 !== g.c2) add({ r1: g.r1, r2: lastR, c1: g.c1, c2: lastC }, 'sg-sel');
        add(a, 'sg-active');
      }
      // Header highlight
      for (const d of this.colHdrInner.children) d.classList.toggle('on', +d.dataset.col >= g.c1 && +d.dataset.col <= g.c2);
      [...this.rowHdrInner.children].forEach((d) => { const r = +d.textContent; d.classList.toggle('on', r >= g.r1 && r <= g.r2); });
    }

    expandMerges(g) {
      const s = this.sheet;
      let out = { ...g };
      for (let changed = true; changed;) {
        changed = false;
        for (const m of s.merges) {
          if (m.r2 < out.r1 || m.r1 > out.r2 || m.c2 < out.c1 || m.c1 > out.c2) continue;
          const n = { r1: Math.min(out.r1, m.r1), c1: Math.min(out.c1, m.c1), r2: Math.max(out.r2, m.r2), c2: Math.max(out.c2, m.c2) };
          if (n.r1 !== out.r1 || n.c1 !== out.c1 || n.r2 !== out.r2 || n.c2 !== out.c2) { out = n; changed = true; }
        }
      }
      return out;
    }

    // ---------- sheets

    showSheet(i) {
      this.endEdit(false);
      this.sheetIndex = i;
      this.layoutDirty = true;
      this.layoutFor = null;
      this.main.scrollTop = 0; this.main.scrollLeft = 0;
      const s = this.sheet;
      if (s.unsupported) {
        this.mainLayer.replaceChildren(Object.assign(document.createElement('div'), { className: 'sg-note', textContent: 'This sheet is a chart or dialog sheet. It is kept in the file but Doc Studio shows worksheets only.' }));
        this.paintTabs();
        return;
      }
      const sv = s.doc && s.doc.getElementsByTagNameNS(s.doc.documentElement.namespaceURI, 'selection')[0];
      const start = sv && sv.getAttribute('activeCell') ? parseAddr(sv.getAttribute('activeCell')) : null;
      this.select(start ? start.r : 1, start ? start.c : 1);
      this.render(true);
      this.paintTabs();
      this.updateBars();
    }

    paintTabs() {
      const bar = this.tabsBar;
      bar.replaceChildren();
      this.wb.sheets.forEach((s, i) => {
        if (s.state !== 'visible') return;
        const b = document.createElement('button');
        b.className = `sg-tab${i === this.sheetIndex ? ' active' : ''}`;
        b.textContent = s.name;
        b.title = 'Double-click to rename';
        b.addEventListener('click', () => { if (i !== this.sheetIndex) this.showSheet(i); });
        b.addEventListener('dblclick', () => this.renameSheet(i));
        bar.append(b);
      });
      const add = document.createElement('button');
      add.className = 'sg-tab sg-add';
      add.textContent = '+';
      add.title = 'New sheet';
      add.addEventListener('click', () => this.addSheet());
      bar.append(add);
    }

    addSheet() {
      let n = this.wb.sheets.length + 1;
      while (this.wb.sheet(`Sheet${n}`)) n++;
      this.wb.addSheet(`Sheet${n}`);
      this.changed();
      this.showSheet(this.wb.sheets.length - 1);
    }

    async renameSheet(i) {
      const s = this.wb.sheets[i];
      const v = this.opts.ask ? await this.opts.ask('Rename sheet', [{ name: 'name', label: 'Sheet name', value: s.name }]) : null;
      if (!v) return;
      const name = v.name.trim();
      if (!name || name === s.name) return;
      if (/[\\/?*[\]:]/.test(name) || name.length > 31) { this.opts.error && this.opts.error('Sheet names can be up to 31 characters and can\'t contain \\ / ? * [ ] :'); return; }
      if (this.wb.sheet(name)) { this.opts.error && this.opts.error(`There is already a sheet called ${name}.`); return; }
      this.wb.renameSheet(s, name);
      this.calc.build();
      this.changed();
      this.paintTabs();
      this.updateBars();
    }

    // ---------- selection

    select(r, c, extend) {
      r = Math.max(1, Math.min(SheetModel.MAX_ROW, r));
      c = Math.max(1, Math.min(SheetModel.MAX_COL, c));
      if (extend) {
        this.sel = { r1: Math.min(this.anchor.r, r), c1: Math.min(this.anchor.c, c), r2: Math.max(this.anchor.r, r), c2: Math.max(this.anchor.c, c) };
        this.cursor = { r, c };
      } else {
        const m = this.sheet.mergeAt(r, c);
        if (m) { r = m.r1; c = m.c1; }
        this.active = { r, c };
        this.anchor = { r, c };
        this.cursor = { r, c };
        this.sel = { r1: r, c1: c, r2: r, c2: c };
      }
      this.updateBars();
      if (this.rowTop) { this.ensureVisible(this.cursor.r, this.cursor.c); this.paintSelection(); this.paintHeaderHighlight(); }
      if (this.opts.onSelection) this.opts.onSelection();
    }

    paintHeaderHighlight() {
      const g = this.expandMerges(this.sel);
      for (const d of this.colHdrFrozen.children) { const c = SheetModel.colNum(d.textContent); d.classList.toggle('on', c >= g.c1 && c <= g.c2); }
      for (const d of this.rowHdrFrozen.children) { const r = +d.textContent; d.classList.toggle('on', r >= g.r1 && r <= g.r2); }
    }

    ensureVisible(r, c) {
      this.layout();
      if (r > this.nRows - 5 || c > this.nCols - 3) { this.layoutDirty = true; this.layout(); }
      const m = this.main;
      if (r > this.F.rows) {
        const y1 = this.rowTop[r] - this.frozenH, y2 = this.rowTop[Math.min(r + 1, this.nRows + 1)] - this.frozenH;
        if (y1 < m.scrollTop) m.scrollTop = y1;
        else if (y2 > m.scrollTop + m.clientHeight) m.scrollTop = y2 - m.clientHeight;
      }
      if (c > this.F.cols) {
        const x1 = this.colLeft[c] - this.frozenW, x2 = this.colLeft[Math.min(c + 1, this.nCols + 1)] - this.frozenW;
        if (x1 < m.scrollLeft) m.scrollLeft = x1;
        else if (x2 > m.scrollLeft + m.clientWidth) m.scrollLeft = Math.min(x1, x2 - m.clientWidth);
      }
      this.render();
    }

    goTo(text) {
      const t = text.trim();
      let sheet = this.sheet, ref = t;
      const bang = t.lastIndexOf('!');
      if (bang > 0) { const sn = t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'"); const s = this.wb.sheet(sn); if (s) { sheet = s; ref = t.slice(bang + 1); } }
      const g = parseRange(ref) || (this.wb.resolveName(t, this.sheet.name) && (() => { const n = this.wb.resolveName(t, this.sheet.name); return n.from ? { r1: n.from.row, c1: n.from.col, r2: n.to.row, c2: n.to.col } : { r1: n.row, c1: n.col, r2: n.row, c2: n.col }; })());
      if (!g) { this.updateBars(); return; }
      const idx = this.wb.sheets.indexOf(sheet);
      if (idx !== this.sheetIndex) this.showSheet(idx);
      this.select(g.r1, g.c1);
      if (g.r2 !== g.r1 || g.c2 !== g.c1) { this.anchor = { r: g.r1, c: g.c1 }; this.select(g.r2, g.c2, true); }
    }

    updateBars() {
      const g = this.sel;
      this.nameBox.value = g.r1 === g.r2 && g.c1 === g.c2 ? addr(this.active.r, this.active.c) : (this.sel.r2 - this.sel.r1 + 1) + 'R × ' + (this.sel.c2 - this.sel.c1 + 1) + 'C';
      if (document.activeElement !== this.nameBox) this.nameBox.value = addr(this.active.r, this.active.c);
      if (!this.editing) this.fInput.value = rawText(this.cellOrMaster(this.active.r, this.active.c), this.wb.styles, this.wb.date1904);
    }

    cellOrMaster(r, c) {
      const cell = this.sheet.get(r, c);
      if (cell && cell.arrayMaster && !cell.formula) {
        const { r: mr, c: mc } = SheetModel.unkey(cell.arrayMaster);
        const m = this.sheet.get(mr, mc);
        return m ? { ...m, value: cell.value } : cell;
      }
      return cell;
    }

    // Sum / Average / Count of the selection, for the status bar.
    selectionStats() {
      const g = this.sel;
      if (g.r1 === g.r2 && g.c1 === g.c2) return '';
      let sum = 0, count = 0, nums = 0;
      const s = this.sheet;
      const r2 = Math.min(g.r2, s.maxRow), c2 = Math.min(g.c2, s.maxCol);
      if ((r2 - g.r1 + 1) * (c2 - g.c1 + 1) > 500000) return '';
      for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) {
        const cell = s.get(r, c);
        if (!cell || cell.value === null || cell.value === '') continue;
        count++;
        if (typeof cell.value === 'number') { sum += cell.value; nums++; }
      }
      if (!count) return '';
      const f = (n) => (+n.toPrecision(12)).toLocaleString(undefined, { maximumFractionDigits: 10 });
      return nums ? `Average: ${f(sum / nums)}   Count: ${count}   Sum: ${f(sum)}` : `Count: ${count}`;
    }

    // ---------- mouse

    hit(e) {
      const rect = this.grid.getBoundingClientRect();
      // With zoom, screen pixels are divided back into the grid's own pixels.
      const z = this.zoom || 1;
      const x = (e.clientX - rect.left) / z, y = (e.clientY - rect.top) / z;
      const inColHdr = y < HEADER_H, inRowHdr = x < HEADER_W;
      const gx = x - HEADER_W, gy = y - HEADER_H;
      const sx = gx < this.frozenW ? gx : gx + this.main.scrollLeft;
      const sy = gy < this.frozenH ? gy : gy + this.main.scrollTop;
      return { x, y, inColHdr, inRowHdr, r: this.rowAt(Math.max(0, sy)), c: this.colAt(Math.max(0, sx)), sx, sy };
    }

    onHover(e) {
      if (this.dragging) return;
      const h = this.hit(e);
      const near = h.inColHdr && !h.inRowHdr && this.colEdgeNear(h.sx);
      const nearRow = h.inRowHdr && !h.inColHdr && this.rowEdgeNear && this.rowEdgeNear(h.sy);
      this.grid.style.cursor = near ? 'col-resize' : nearRow ? 'row-resize' : '';
    }

    colEdgeNear(sx) {
      const c = this.colAt(sx);
      if (Math.abs(this.colLeft[c + 1] - sx) <= 3) return c;
      if (c > 1 && Math.abs(this.colLeft[c] - sx) <= 3) return c - 1;
      return 0;
    }

    onMouseDown(e) {
      if (e.target === this.editor || e.button !== 0) return;
      const sbW = this.main.offsetWidth - this.main.clientWidth;
      const rect = this.main.getBoundingClientRect();
      if (e.clientX > rect.right - sbW || e.clientY > rect.bottom - (this.main.offsetHeight - this.main.clientHeight)) return; // scrollbars
      const h = this.hit(e);
      // Click a cell while typing a formula: insert its reference.
      if (this.editing && !h.inColHdr && !h.inRowHdr && this.canInsertRef()) {
        e.preventDefault();
        this.insertRef(h.r, h.c, h.r, h.c);
        this.dragging = { mode: 'ref', r: h.r, c: h.c };
        const move = (ev) => { const k = this.hit(ev); this.insertRef(this.dragging.r, this.dragging.c, k.r, k.c); };
        const up = () => { this.dragging = null; window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
        window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
        return;
      }
      if (this.editing) this.endEdit(true);
      e.preventDefault();
      this.grid.focus({ preventScroll: true });
      const fbtn = e.target.closest && e.target.closest('.sg-filter-btn');
      if (fbtn && this.openFilterMenu) { this.openFilterMenu(+fbtn.dataset.r, +fbtn.dataset.c, fbtn); return; }
      if (h.inColHdr && h.inRowHdr) { this.anchor = { r: 1, c: 1 }; this.active = { r: 1, c: 1 }; this.sel = { r1: 1, c1: 1, r2: SheetModel.MAX_ROW, c2: SheetModel.MAX_COL }; this.updateBars(); this.paintSelection(); this.opts.onSelection && this.opts.onSelection(); return; }
      if (h.inColHdr) {
        const edge = this.colEdgeNear(h.sx);
        if (edge) { this.startColResize(edge, e.clientX); return; }
        this.selectCols(e.shiftKey ? this.anchor.c : h.c, h.c, e.shiftKey);
        const move = (ev) => { const k = this.hit(ev); this.selectCols(this.anchor.c, k.c, true); };
        const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
        window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
        return;
      }
      if (h.inRowHdr && this.rowEdgeNear) {
        const edge = this.rowEdgeNear(h.sy);
        if (edge) { this.startRowResize(edge, e.clientY); return; }
      }
      if (h.inRowHdr) {
        this.selectRows(e.shiftKey ? this.anchor.r : h.r, h.r, e.shiftKey);
        const move = (ev) => { const k = this.hit(ev); this.selectRows(this.anchor.r, k.r, true); };
        const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
        window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
        return;
      }
      this.select(h.r, h.c, e.shiftKey);
      this.dragging = { mode: 'sel' };
      const move = (ev) => {
        const k = this.hit(ev);
        if (k.r !== this.cursor.r || k.c !== this.cursor.c) this.select(k.r, k.c, true);
      };
      const up = () => { this.dragging = null; window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
      window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
    }

    selectCols(c1, c2, keepAnchor) {
      if (!keepAnchor) { this.anchor = { r: 1, c: c1 }; this.active = { r: 1, c: c1 }; }
      this.sel = { r1: 1, r2: SheetModel.MAX_ROW, c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
      this.cursor = { r: 1, c: c2 };
      this.updateBars(); this.paintSelection(); this.paintHeaderHighlight();
      this.opts.onSelection && this.opts.onSelection();
    }
    selectRows(r1, r2, keepAnchor) {
      if (!keepAnchor) { this.anchor = { r: r1, c: 1 }; this.active = { r: r1, c: 1 }; }
      this.sel = { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: 1, c2: SheetModel.MAX_COL };
      this.cursor = { r: r2, c: 1 };
      this.updateBars(); this.paintSelection(); this.paintHeaderHighlight();
      this.opts.onSelection && this.opts.onSelection();
    }

    startColResize(c, startX) {
      const startW = this.colLeft[c + 1] - this.colLeft[c];
      const guide = document.createElement('div');
      guide.className = 'sg-resize-guide';
      this.grid.append(guide);
      const place = (w) => { const x = HEADER_W + (c > this.F.cols ? this.colLeft[c] - this.main.scrollLeft : this.colLeft[c]) + w; guide.style.left = `${x}px`; };
      place(startW);
      let w = startW;
      const move = (ev) => { w = Math.max(0, startW + (ev.clientX - startX) / (this.zoom || 1)); place(w); };
      const up = () => {
        window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up);
        guide.remove();
        if (w === startW) return;
        const cols = (this.sel.r1 === 1 && this.sel.r2 === SheetModel.MAX_ROW && c >= this.sel.c1 && c <= this.sel.c2) ? range(this.sel.c1, this.sel.c2) : [c];
        this.recordCols(() => cols.forEach((cc) => this.sheet.setColWidth(cc, pxToWidth(w))));
      };
      window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
    }

    recordCols(fn) {
      const s = this.sheet;
      const before = s.cols.map((x) => ({ ...x }));
      fn();
      const after = s.cols.map((x) => ({ ...x }));
      this.pushUndo({ sheet: s, cols: { before, after } });
      this.layoutDirty = true;
      this.render(true);
      this.changed();
    }

    onDblClick(e) {
      const h = this.hit(e);
      if (h.inColHdr && !h.inRowHdr) {
        const edge = this.colEdgeNear(h.sx);
        if (edge) this.autoFitCol(edge);
        return;
      }
      if (h.inColHdr || h.inRowHdr) return;
      this.beginEdit(null, 'cell');
    }

    autoFitCol(c) {
      const s = this.sheet;
      let max = 0;
      for (let r = 1; r <= s.maxRow; r++) {
        const cell = s.get(r, c);
        if (!cell || cell.value === null) continue;
        const css = this.wb.styles.css(cell.s);
        this.measure.font = css.font;
        max = Math.max(max, this.measure.measureText(display(cell, this.wb.styles, this.wb.date1904).text).width);
      }
      if (max) this.recordCols(() => s.setColWidth(c, pxToWidth(Math.ceil(max) + 10)));
    }

    // ---------- keyboard

    onKey(e) {
      // Keys typed in the in-cell editor bubble up here; the editor handles its own keys.
      if (this.editing || e.target === this.editor) return;
      const k = e.key;
      const ctrl = e.ctrlKey || e.metaKey;
      const move = (dr, dc) => {
        e.preventDefault();
        const from = e.shiftKey ? this.cursor : this.active;
        let r = from.r, c = from.c;
        if (ctrl) [r, c] = this.jump(r, c, dr, dc);
        else {
          const m = this.sheet.mergeAt(r, c);
          if (m && !e.shiftKey) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; }
          r += dr; c += dc;
        }
        this.select(r, c, e.shiftKey);
      };
      if (k === 'ArrowDown') return move(1, 0);
      if (k === 'ArrowUp') return move(-1, 0);
      if (k === 'ArrowRight') return move(0, 1);
      if (k === 'ArrowLeft') return move(0, -1);
      if (k === 'PageDown' || k === 'PageUp') {
        e.preventDefault();
        const rows = Math.max(1, Math.floor(this.main.clientHeight / 20) - 1);
        this.main.scrollTop += (k === 'PageDown' ? 1 : -1) * rows * 20;
        this.select(Math.max(1, this.active.r + (k === 'PageDown' ? rows : -rows)), this.active.c, e.shiftKey);
        return;
      }
      if (k === 'Home') { e.preventDefault(); return this.select(ctrl ? 1 : this.active.r, 1, e.shiftKey); }
      if (k === 'End' && ctrl) { e.preventDefault(); return this.select(Math.max(1, this.sheet.maxRow), Math.max(1, this.sheet.maxCol), e.shiftKey); }
      if (k === 'Tab') { e.preventDefault(); return this.moveAfterCommit(0, e.shiftKey ? -1 : 1); }
      if (k === 'Enter') { e.preventDefault(); return this.moveAfterCommit(e.shiftKey ? -1 : 1, 0); }
      if (k === 'F2') { e.preventDefault(); return this.beginEdit(null, 'cell'); }
      if (k === 'F9') { e.preventDefault(); return this.recalcAll(); }
      if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); if (k === 'Backspace') return this.beginEdit('', 'cell'); return this.clearSelection(); }
      if (ctrl) {
        const l = k.toLowerCase();
        if (l === 'z') { e.preventDefault(); return e.shiftKey ? this.redo() : this.undo(); }
        if (l === 'y') { e.preventDefault(); return this.redo(); }
        if (l === 'a') { e.preventDefault(); this.anchor = { r: 1, c: 1 }; this.sel = { r1: 1, c1: 1, r2: SheetModel.MAX_ROW, c2: SheetModel.MAX_COL }; this.updateBars(); this.paintSelection(); return; }
        if (l === 'b') { e.preventDefault(); return this.toggleFont('bold'); }
        if (l === 'i') { e.preventDefault(); return this.toggleFont('italic'); }
        if (l === 'u') { e.preventDefault(); return this.toggleFont('underline'); }
        if (l === 'd') { e.preventDefault(); return this.fill('down'); }
        if (l === 'r') { e.preventDefault(); return this.fill('right'); }
        if (l === 'c' || l === 'x' || l === 'v') return; // native copy/cut/paste events
        return;
      }
      if (k === 'Escape') { if (this.clip) { this.clip = null; this.paintSelection(); } return; }
      if (k.length === 1 && !e.altKey) { e.preventDefault(); this.beginEdit(k, 'cell', true); }
    }

    // Ctrl+Arrow: to the edge of the current block of data, or to the next filled cell.
    jump(r, c, dr, dc) {
      const s = this.sheet;
      const filled = (rr, cc) => { const x = s.get(rr, cc); return !!x && x.value !== null && x.value !== ''; };
      const maxR = Math.max(s.maxRow, r), maxC = Math.max(s.maxCol, c);
      const ok = (rr, cc) => rr >= 1 && cc >= 1 && rr <= maxR && cc <= maxC;
      let rr = r, cc = c;
      if (filled(r, c) && ok(r + dr, c + dc) && filled(r + dr, c + dc)) {
        while (ok(rr + dr, cc + dc) && filled(rr + dr, cc + dc)) { rr += dr; cc += dc; }
        return [rr, cc];
      }
      rr += dr; cc += dc;
      while (ok(rr, cc) && !filled(rr, cc)) { rr += dr; cc += dc; }
      if (ok(rr, cc)) return [rr, cc];
      return [dr < 0 ? 1 : dr > 0 ? maxR : r, dc < 0 ? 1 : dc > 0 ? maxC : c];
    }

    moveAfterCommit(dr, dc) {
      const g = this.sel;
      if (g.r1 === g.r2 && g.c1 === g.c2) {
        const m = this.sheet.mergeAt(this.active.r, this.active.c);
        let r = this.active.r, c = this.active.c;
        if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; }
        this.select(r + dr, c + dc);
        return;
      }
      // Within a selection, Enter/Tab move inside it.
      let { r, c } = this.active;
      c += dc; r += dr;
      if (c > g.c2) { c = g.c1; r++; } if (c < g.c1) { c = g.c2; r--; }
      if (r > g.r2) { r = g.r1; c = Math.min(g.c2, c + (dr ? 1 : 0)); if (c > g.c2) c = g.c1; } if (r < g.r1) r = g.r2;
      this.active = { r, c };
      this.updateBars(); this.paintSelection(); this.ensureVisible(r, c);
    }

    // ---------- editing

    beginEdit(initial, where, typed) {
      const { r, c } = this.active;
      const cell = this.sheet.get(r, c);
      if (cell && cell.arrayMaster && !cell.formula) { this.opts.error && this.opts.error("You can't change part of an array formula."); return; }
      const raw = rawText(cell, this.wb.styles, this.wb.date1904);
      this.editing = { r, c, where, typed: !!typed, original: raw };
      if (where === 'bar') { this.pendingRef = null; return; }
      const css = this.wb.styles.css(cell ? cell.s : 0);
      this.editor.style.font = css.font;
      this.editor.style.color = css.color || '#000';
      this.editor.style.background = css.fill || '#fff';
      this.editor.style.textAlign = css.h === 'right' || css.h === 'center' ? css.h : 'left';
      this.editor.value = initial != null ? initial : raw;
      this.fInput.value = this.editor.value;
      this.editor.hidden = false;
      this.positionEditor();
      this.editor.focus();
      const end = this.editor.value.length;
      this.editor.setSelectionRange(end, end);
      this.pendingRef = null;
    }

    positionEditor() {
      const { r, c } = this.editing;
      const m = this.sheet.mergeAt(r, c) || { r1: r, c1: c, r2: r, c2: c };
      const x = HEADER_W + (c > this.F.cols ? this.colLeft[m.c1] - this.main.scrollLeft : this.colLeft[m.c1]);
      const y = HEADER_H + (r > this.F.rows ? this.rowTop[m.r1] - this.main.scrollTop : this.rowTop[m.r1]);
      const w = this.colLeft[m.c2 + 1] - this.colLeft[m.c1];
      const hgt = this.rowTop[m.r2 + 1] - this.rowTop[m.r1];
      Object.assign(this.editor.style, { left: `${x}px`, top: `${y}px`, minWidth: `${w + 1}px`, width: `${w + 1}px`, minHeight: `${hgt + 1}px`, height: `${hgt + 1}px` });
      this.autoSizeEditor();
    }

    autoSizeEditor() {
      if (this.editor.hidden) return;
      this.measure.font = this.editor.style.font;
      const lines = this.editor.value.split('\n');
      const widest = Math.max(...lines.map((l) => this.measure.measureText(l).width)) + 12;
      const min = parseFloat(this.editor.style.minWidth) || 60;
      this.editor.style.width = `${Math.min(Math.max(min, widest), this.grid.clientWidth - parseFloat(this.editor.style.left) - 4)}px`;
      this.editor.style.height = 'auto';
      this.editor.style.height = `${Math.max(parseFloat(this.editor.style.minHeight) || 20, this.editor.scrollHeight)}px`;
    }

    onEditorKey(e, fromBar) {
      if (!this.editing && fromBar) {
        if (e.key === 'Enter') { e.preventDefault(); this.grid.focus(); }
        return;
      }
      const k = e.key;
      const input = fromBar ? this.fInput : this.editor;
      if (k === 'Enter' && e.altKey) {
        e.preventDefault();
        const { selectionStart: s, selectionEnd: en, value } = input;
        input.value = value.slice(0, s) + '\n' + value.slice(en);
        input.setSelectionRange(s + 1, s + 1);
        input.dispatchEvent(new Event('input'));
        return;
      }
      if (k === 'Enter' || k === 'Tab') {
        e.preventDefault();
        if (fromBar) this.editor.value = this.fInput.value;
        if (this.endEdit(true) === false) return;
        this.grid.focus({ preventScroll: true });
        this.moveAfterCommit(k === 'Enter' ? (e.shiftKey ? -1 : 1) : 0, k === 'Tab' ? (e.shiftKey ? -1 : 1) : 0);
        return;
      }
      if (k === 'Escape') { e.preventDefault(); this.endEdit(false); this.grid.focus({ preventScroll: true }); return; }
      if (!fromBar && this.editing.typed && /^Arrow/.test(k) && !this.isFormulaRefPosition()) {
        e.preventDefault();
        if (this.endEdit(true) === false) return;
        this.grid.focus({ preventScroll: true });
        const d = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
        this.select(this.active.r + d[0], this.active.c + d[1]);
      }
      if (k === 'F4' && (fromBar ? this.fInput : this.editor).value.startsWith('=')) { e.preventDefault(); this.cycleAbsolute(input); }
    }

    isFormulaRefPosition() { return false; }

    canInsertRef() {
      const input = this.editing && this.editing.where === 'bar' ? this.fInput : this.editor;
      const v = input.value;
      if (!v.startsWith('=')) return false;
      if (this.pendingRef) return true;
      const before = v.slice(0, input.selectionStart).trimEnd();
      return /[=(,+\-*/^&<>:;]$/.test(before);
    }

    insertRef(r1, c1, r2, c2) {
      const input = this.editing.where === 'bar' ? this.fInput : this.editor;
      const g = { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
      const sheetPrefix = this.editing.sheet && this.editing.sheet !== this.sheet ? quoteSheet(this.sheet.name) + '!' : '';
      const ref = sheetPrefix + rangeRef(g);
      const v = input.value;
      let start = input.selectionStart, end = input.selectionEnd;
      if (this.pendingRef) { start = this.pendingRef.start; end = this.pendingRef.end; }
      input.value = v.slice(0, start) + ref + v.slice(end);
      this.pendingRef = { start, end: start + ref.length };
      input.setSelectionRange(start + ref.length, start + ref.length);
      if (input === this.editor) this.fInput.value = input.value; else this.editor.value = input.value;
      this.autoSizeEditor();
      this.refHighlight = g;
      input.focus();
    }

    cycleAbsolute(input) {
      const v = input.value;
      const pos = input.selectionStart;
      const tok = scanRefs(v.slice(1)).find((t) => t.kind === 'ref' && pos - 1 >= t.refStart && pos - 1 <= t.end);
      if (!tok) return;
      const state = (tok.colAbs ? 2 : 0) + (tok.rowAbs ? 1 : 0);
      const next = { 0: [true, true], 3: [false, true], 1: [true, false], 2: [false, false] }[state];
      const text = `${next[0] ? '$' : ''}${colName(tok.col)}${next[1] ? '$' : ''}${tok.row}`;
      const s = tok.refStart + 1, e = tok.end + 1;
      input.value = v.slice(0, s) + text + v.slice(e);
      input.setSelectionRange(s + text.length, s + text.length);
      input.dispatchEvent(new Event('input'));
    }

    // Commit (or cancel) the edit. Returns false if the input was rejected.
    endEdit(commit) {
      if (!this.editing) return true;
      const ed = this.editing;
      const text = ed.where === 'bar' ? this.fInput.value : this.editor.value;
      if (commit && text !== ed.original) {
        if (text.startsWith('=') && text.length > 1) {
          const probe = this.calc.engine.evaluate(this.calc.prep(text.slice(1), this.sheet, ed.r, ed.c), this.sheet.name, ed.r, ed.c);
          if (probe && probe.error === '#ERROR!') { this.opts.error && this.opts.error('There is a problem with this formula.'); return false; }
        }
        this.editing = null;
        this.editor.hidden = true;
        this.setFromText(ed.r, ed.c, text);
      } else {
        this.editing = null;
        this.editor.hidden = true;
      }
      this.pendingRef = null;
      this.updateBars();
      return true;
    }

    setFromText(r, c, text) {
      if (r < 1 || c < 1) return; // rows and columns are 1-based
      const s = this.sheet;
      const cell = s.get(r, c);
      const fmt = this.wb.styles.numFmt(cell ? cell.s : s.defaultStyleAt(r, c));
      const p = parseInput(text, fmt === 'General', this.wb.date1904);
      this.apply([{ r, c, patch: p }]);
    }

    // Apply a batch of cell changes as one undo step. Each item: { r, c, patch: { value?, formula?, numFmt?, s? } }
    apply(items, sheet) {
      sheet = sheet || this.sheet;
      const undo = [];
      const touched = [];
      for (const { r, c, patch } of items) {
        const prior = sheet.get(r, c);
        if (prior && prior.spillFrom && ('value' in patch || 'formula' in patch)) { const a = SheetModel.unkey(prior.spillFrom); touched.push({ sheet, r: a.r, c: a.c }); }
        const before = snapshot(prior);
        const p = {};
        if ('formula' in patch) {
          p.formula = patch.formula || null;
          p.formulaChanged = true;
          p.dynamic = !!patch.formula;
          if (!patch.formula) p.value = 'value' in patch ? patch.value : null;
        } else if ('value' in patch) { p.value = patch.value; p.formula = null; p.formulaChanged = !!(sheet.get(r, c) && sheet.get(r, c).formula); }
        if (patch.numFmt) { const cur = sheet.get(r, c); p.s = this.wb.styles.derive(cur ? cur.s : sheet.defaultStyleAt(r, c), { numFmt: patch.numFmt }); }
        if ('s' in patch) p.s = patch.s;
        const cell = sheet.setCell(r, c, p);
        if ('formula' in p || before.formula) { this.calc.formulaChanged(sheet, cell); this.wb.formulasChanged = true; }
        undo.push({ r, c, before, after: snapshot(cell) });
        touched.push({ sheet, r, c });
      }
      this.calc.recalc(touched);
      if (this.calc.cycles && this.calc.cycles.size) {
        const first = this.calc.nodes.get([...this.calc.cycles][0]);
        this.opts.onStatus && this.opts.onStatus(`Circular reference${first ? ` at ${first.sheet.name}!${addr(first.cell.r, first.cell.c)}` : ''} - those cells show 0`);
      }
      this.pushUndo({ sheet, cells: undo });
      this.render(true);
      this.updateBars();
      this.changed();
    }

    pushUndo(entry) {
      this.undoStack.push(entry);
      if (this.undoStack.length > 200) this.undoStack.shift();
      this.redoStack = [];
    }

    restore(entry, which) {
      const s = entry.sheet;
      const idx = this.wb.sheets.indexOf(s);
      if (idx !== this.sheetIndex) this.showSheet(idx);
      if (entry.cells) {
        const touched = [];
        for (const it of entry.cells) {
          const snap = it[which];
          const cell = s.setCell(it.r, it.c, { value: snap.value, formula: snap.formula, s: snap.s, formulaChanged: true });
          if (snap.array) cell.array = snap.array;
          this.calc.formulaChanged(s, cell);
          touched.push({ sheet: s, r: it.r, c: it.c });
        }
        this.calc.recalc(touched);
        const first = entry.cells[0];
        if (first) this.select(first.r, first.c);
      }
      if (entry.cols) { s.cols = entry.cols[which].map((x) => ({ ...x })); s.colsChanged = true; this.layoutDirty = true; }
      if (entry.merges) { s.merges = entry.merges[which].map((x) => ({ ...x })); s.mergesChanged = true; }
      this.render(true);
      this.updateBars();
      this.changed();
    }
    undo() { const e = this.undoStack.pop(); if (!e) return; this.restore(e, 'before'); this.redoStack.push(e); }
    redo() { const e = this.redoStack.pop(); if (!e) return; this.restore(e, 'after'); this.undoStack.push(e); }

    recalcAll() {
      const n = this.calc.recalc([], { all: true }).length;
      if (n) this.changed();
      this.render(true);
      this.updateBars();
      this.opts.onStatus && this.opts.onStatus(n ? `Recalculated - ${n} value(s) changed` : 'Recalculated - no changes');
    }

    // The selection, except that huge selections (whole rows/columns) are trimmed to the data.
    effectiveRange() {
      const g = this.expandMerges(this.sel);
      const s = this.sheet;
      if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) <= 200000) return g;
      return { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, Math.max(s.maxRow, g.r1)), c2: Math.min(g.c2, Math.max(s.maxCol, g.c1)) };
    }

    selectedCells() {
      const { r1, c1, r2, c2 } = this.effectiveRange();
      const g = { r1, c1 };
      const out = [];
      for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) out.push({ r, c });
      return out;
    }

    clearSelection() {
      const items = this.selectedCells().filter(({ r, c }) => { const x = this.sheet.get(r, c); return x && (x.value !== null || x.formula); });
      if (items.some(({ r, c }) => { const x = this.sheet.get(r, c); return x.arrayMaster && !x.formula; })) { this.opts.error && this.opts.error("You can't change part of an array formula."); return; }
      if (items.length) this.apply(items.map(({ r, c }) => ({ r, c, patch: { value: null, formula: null } })));
    }

    // Ctrl+D / Ctrl+R: copy the first row/column of the selection across it, adjusting relative references.
    fill(dir) {
      const g = this.sel;
      const items = [];
      const s = this.sheet;
      if (dir === 'down') {
        const srcR = g.r1 === g.r2 ? g.r1 - 1 : g.r1;
        if (srcR < 1) return;
        for (let c = g.c1; c <= Math.min(g.c2, Math.max(s.maxCol, g.c1)); c++) {
          const src = s.get(srcR, c);
          for (let r = srcR + 1; r <= Math.min(g.r2, srcR + 100000); r++) items.push({ r, c, patch: copyPatch(src, r - srcR, 0) });
        }
      } else {
        const srcC = g.c1 === g.c2 ? g.c1 - 1 : g.c1;
        if (srcC < 1) return;
        for (let r = g.r1; r <= Math.min(g.r2, Math.max(s.maxRow, g.r1)); r++) {
          const src = s.get(r, srcC);
          for (let c = srcC + 1; c <= g.c2; c++) items.push({ r, c, patch: copyPatch(src, 0, c - srcC) });
        }
      }
      if (items.length) this.apply(items);
    }

    // ---------- clipboard

    onCopy(e) {
      if (this.editing) return;
      e.preventDefault();
      const g = this.expandMerges(this.sel);
      const s = this.sheet;
      const r2 = Math.min(g.r2, Math.max(s.maxRow, g.r1)), c2 = Math.min(g.c2, Math.max(s.maxCol, g.c1));
      const lines = [];
      const cells = [];
      for (let r = g.r1; r <= r2; r++) {
        const row = [];
        for (let c = g.c1; c <= c2; c++) {
          const cell = s.get(r, c);
          row.push(display(cell, this.wb.styles, this.wb.date1904).text.replace(/\t|\r?\n/g, ' '));
          cells.push({ dr: r - g.r1, dc: c - g.c1, snap: snapshot(cell) });
        }
        lines.push(row.join('\t'));
      }
      const text = lines.join('\r\n');
      e.clipboardData.setData('text/plain', text);
      this.clip = { text, cells, sheet: s, g: { r1: g.r1, c1: g.c1, r2, c2 }, cut: false };
      this.opts.onStatus && this.opts.onStatus('Copied - select a cell and paste');
    }

    onPaste(e) {
      if (this.editing) return;
      e.preventDefault();
      const text = e.clipboardData.getData('text/plain');
      if (!text) return;
      const { r: r0, c: c0 } = this.active;
      const items = [];
      if (this.clip && text.replace(/\r/g, '') === this.clip.text.replace(/\r/g, '')) {
        // Our own copy: paste formulas (with references adjusted) and formatting.
        const dr = r0 - this.clip.g.r1, dc = c0 - this.clip.g.c1;
        for (const { dr: r, dc: c, snap } of this.clip.cells) {
          const patch = snap.formula ? { formula: shiftFormula(snap.formula, dr, dc) } : { value: snap.value };
          if (this.clip.sheet === this.sheet || this.wb.sheets.includes(this.clip.sheet)) patch.s = snap.s;
          items.push({ r: r0 + r, c: c0 + c, patch });
        }
        if (this.clip.cut) {
          const dest = new Set(items.map((x) => key(x.r, x.c)));
          for (let r = this.clip.g.r1; r <= this.clip.g.r2; r++) for (let c = this.clip.g.c1; c <= this.clip.g.c2; c++) {
            if (this.clip.sheet === this.sheet && dest.has(key(r, c))) continue;
            if (this.clip.sheet === this.sheet) items.push({ r, c, patch: { value: null, formula: null, s: 0 } });
          }
          this.clip = null;
        }
      } else {
        const rows = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
        rows.forEach((row, i) => row.forEach((t, j) => {
          const cell = this.sheet.get(r0 + i, c0 + j);
          const fmt = this.wb.styles.numFmt(cell ? cell.s : 0);
          items.push({ r: r0 + i, c: c0 + j, patch: parseInput(t, fmt === 'General', this.wb.date1904) });
        }));
      }
      if (!items.length) return;
      this.apply(items);
      const last = items.reduce((m, x) => ({ r: Math.max(m.r, x.r), c: Math.max(m.c, x.c) }), { r: r0, c: c0 });
      this.anchor = { r: r0, c: c0 };
      this.select(last.r, last.c, true);
    }

    // ---------- formatting

    activeCss() { const cell = this.sheet.get(this.active.r, this.active.c); return this.wb.styles.css(cell ? cell.s : this.sheet.defaultStyleAt(this.active.r, this.active.c)); }

    styleSelection(change) {
      const cells = this.selectedCells();
      if (cells.length > 200000) { this.opts.error && this.opts.error('That selection is too large to format in one go.'); return; }
      const s = this.sheet;
      const items = cells.map(({ r, c }) => {
        const cur = s.get(r, c);
        const base = cur ? cur.s : s.defaultStyleAt(r, c);
        const ch = typeof change === 'function' ? change(r, c) : change;
        return { r, c, patch: { s: this.wb.styles.derive(base, ch) } };
      });
      this.apply(items);
    }

    toggleFont(prop) {
      const on = this.activeCss().font_[prop];
      this.styleSelection({ font: { [prop]: !on } });
    }

    borders(preset) {
      const g = this.effectiveRange();
      const T = 'thin';
      this.styleSelection((r, c) => {
        if (preset === 'none') return { border: { left: null, right: null, top: null, bottom: null } };
        if (preset === 'all') return { border: { left: T, right: T, top: T, bottom: T } };
        if (preset === 'bottom') return r === g.r2 ? { border: { bottom: T } } : {};
        const b = {};
        if (r === g.r1) b.top = T; if (r === g.r2) b.bottom = T;
        if (c === g.c1) b.left = T; if (c === g.c2) b.right = T;
        return Object.keys(b).length ? { border: b } : {};
      });
    }

    changeDecimals(delta) {
      const cell = this.sheet.get(this.active.r, this.active.c);
      let fmt = this.wb.styles.numFmt(cell ? cell.s : 0);
      if (fmt === 'General') {
        const v = cell && typeof cell.value === 'number' ? cell.value : 0;
        const dec = (String(v).split('.')[1] || '').length;
        const n = Math.max(0, dec + delta);
        fmt = n ? `0.${'0'.repeat(n)}` : '0';
      } else {
        const secs = formatSections(fmt).map((sec) => {
          const m = /(0|#)(\.(0+))?(?![^"]*"(?:[^"]*"[^"]*")*[^"]*$)/.exec(sec);
          if (!m) return sec;
          const cur = m[3] ? m[3].length : 0;
          const n = Math.max(0, cur + delta);
          const repl = m[1] + (n ? '.' + '0'.repeat(n) : '');
          return sec.slice(0, m.index) + repl + sec.slice(m.index + m[0].length);
        });
        fmt = secs.join(';');
      }
      this.styleSelection({ numFmt: fmt });
    }

    toggleMerge() {
      const g = this.sel;
      const s = this.sheet;
      const before = s.merges.map((m) => ({ ...m }));
      const hit = s.merges.filter((m) => m.r2 >= g.r1 && m.r1 <= g.r2 && m.c2 >= g.c1 && m.c1 <= g.c2);
      if (hit.length) s.merges = s.merges.filter((m) => !hit.includes(m));
      else {
        if (g.r1 === g.r2 && g.c1 === g.c2) return;
        if (g.r2 - g.r1 > 1000 || g.c2 - g.c1 > 200) { this.opts.error && this.opts.error('That selection is too large to merge.'); return; }
        // Excel keeps only the top-left value.
        const others = [];
        for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
          if (r === g.r1 && c === g.c1) continue;
          const x = s.get(r, c);
          if (x && (x.value !== null || x.formula)) others.push({ r, c, patch: { value: null, formula: null } });
        }
        if (others.length) this.apply(others);
        s.merges.push({ ...g });
        const cur = s.get(g.r1, g.c1);
        this.apply([{ r: g.r1, c: g.c1, patch: { s: this.wb.styles.derive(cur ? cur.s : 0, { align: { horizontal: 'center' } }) } }]);
      }
      s.mergesChanged = true;
      this.pushUndo({ sheet: s, merges: { before, after: s.merges.map((m) => ({ ...m })) } });
      this.select(g.r1, g.c1);
      this.render(true);
      this.changed();
    }

    autoSum() {
      const { r, c } = this.active;
      const s = this.sheet;
      const isNum = (rr, cc) => { const x = s.get(rr, cc); return x && typeof x.value === 'number'; };
      let ref = '';
      if (isNum(r - 1, c)) { let r1 = r - 1; while (r1 > 1 && isNum(r1 - 1, c)) r1--; ref = rangeRef({ r1, c1: c, r2: r - 1, c2: c }); }
      else if (isNum(r, c - 1)) { let c1 = c - 1; while (c1 > 1 && isNum(r, c1 - 1)) c1--; ref = rangeRef({ r1: r, c1, r2: r, c2: c - 1 }); }
      this.beginEdit(`=SUM(${ref})`, 'cell');
      if (!ref) { const p = this.editor.value.length - 1; this.editor.setSelectionRange(p, p); }
    }

    // ---------- find (whole current sheet, in display text)

    // Cells whose shown text or formula matches (re: the find bar's pattern, honouring match case / regex).
    findAll(q, re) {
      const needle = q.toLowerCase();
      const out = [];
      if (!needle) return out;
      const test = (s) => { if (!re) return s.toLowerCase().includes(needle); re.lastIndex = 0; return re.test(s); };
      for (const cell of this.sheet.cells.values()) {
        const t = display(cell, this.wb.styles, this.wb.date1904).text;
        if (test(t) || (cell.formula && test(cell.formula))) out.push({ r: cell.r, c: cell.c });
      }
      return out.sort((a, b) => a.r - b.r || a.c - b.c);
    }

    // ---------- printing: the used range as a styled table

    printable() {
      const s = this.sheet;
      const styles = this.wb.styles;
      const table = document.createElement('table');
      table.className = 'sheet-print';
      const R = Math.min(s.maxRow, 5000), C = Math.min(s.maxCol, 60);
      const covered = new Set();
      for (const m of s.merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(key(r, c));
      const colgroup = document.createElement('colgroup');
      for (let c = 1; c <= C; c++) { const col = document.createElement('col'); col.style.width = `${colPx(s.colWidthChars(c))}px`; colgroup.append(col); }
      table.append(colgroup);
      for (let r = 1; r <= R; r++) {
        const tr = document.createElement('tr');
        for (let c = 1; c <= C; c++) {
          if (covered.has(key(r, c))) continue;
          const cell = s.get(r, c);
          const css = styles.css(cell ? cell.s : 0);
          const td = document.createElement('td');
          const m = s.merges.find((x) => x.r1 === r && x.c1 === c);
          if (m) { td.rowSpan = m.r2 - m.r1 + 1; td.colSpan = m.c2 - m.c1 + 1; }
          const disp = display(cell, styles, this.wb.date1904);
          td.textContent = disp.text;
          let h = css.h === 'general' ? (disp.kind === 'number' ? 'right' : 'left') : css.h;
          td.style.cssText = `font:${css.font};${css.color ? `color:${css.color};` : ''}${css.fill ? `background:${css.fill};` : ''}text-align:${h};${css.wrap ? 'white-space:pre-wrap;' : ''}`;
          for (const side of ['left', 'right', 'top', 'bottom']) if (css.border[side]) td.style[`border${side[0].toUpperCase()}${side.slice(1)}`] = `${BORDER_CSS[css.border[side].style] || '1px solid'} ${css.border[side].color}`;
          tr.append(td);
        }
        table.append(tr);
      }
      return table;
    }

    changed() { this.opts.onChange && this.opts.onChange(); }

    destroy() {
      this.resizeObs.disconnect();
      for (const [ev, fn] of Object.entries(this.clipHandlers)) document.removeEventListener(ev, fn);
    }
  }

  function snapshot(cell) {
    if (!cell) return { value: null, formula: null, s: 0 };
    return { value: cell.value, formula: cell.formula || null, s: cell.s || 0, array: cell.array ? { ...cell.array } : undefined };
  }
  function copyPatch(src, dr, dc) {
    if (!src) return { value: null, formula: null, s: 0 };
    return src.formula ? { formula: shiftFormula(src.formula, dr, dc), s: src.s } : { value: src.value, s: src.s };
  }
  function range(a, b) { const out = []; for (let i = a; i <= b; i++) out.push(i); return out; }

  return { View, parseInput, display, rawText };
})();
