'use strict';
// More Google Sheets features (added to SheetView.View): paste special, insert/delete cells with shifting,
// group rows/columns (outline), filter views and slicers (kept in the workbook's custom properties), protected
// ranges, workbook theme, drawings, and copying a sheet to another open workbook.

(() => {
  const P = SheetView.View.prototype;
  const { key, unkey, addr, colName, parseRange, rangeRef, scanRefs, shiftFormula, quoteSheet } = SheetModel;
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const status = (v, m) => v.opts.onStatus && v.opts.onStatus(m);

  // ======================================================================
  // Paste special (from Doc Studio's own copy)
  // ======================================================================

  // mode: 'values' | 'formats' | 'formulas' | 'transpose' | 'widths'
  P.pasteSpecial = function (mode) {
    const clip = this.clip;
    if (!clip) throw new Error('Copy some cells in Doc Studio first (Ctrl+C).');
    const { r: r0, c: c0 } = this.active;
    const items = [];
    const src = clip.sheet;
    if (mode === 'widths') {
      return this.snapshotOp('Paste column widths', () => {
        for (let c = clip.g.c1; c <= clip.g.c2; c++) { const w = src.colWidthChars(c); if (w != null) this.sheet.setColWidth(c0 + c - clip.g.c1, w); }
      });
    }
    for (const { dr, dc, snap } of clip.cells) {
      const r = mode === 'transpose' ? r0 + dc : r0 + dr;
      const c = mode === 'transpose' ? c0 + dr : c0 + dc;
      const orig = src.get(clip.g.r1 + dr, clip.g.c1 + dc);
      if (mode === 'values') items.push({ r, c, patch: { value: orig ? orig.value : snap.value, formula: null } });
      else if (mode === 'formats') items.push({ r, c, patch: { s: snap.s } });
      else if (mode === 'formulas') items.push({ r, c, patch: snap.formula ? { formula: shiftFormula(snap.formula, r - (clip.g.r1 + dr), c - (clip.g.c1 + dc)) } : { value: snap.value } });
      else if (mode === 'transpose') items.push({ r, c, patch: snap.formula ? { formula: shiftFormula(snap.formula, r - (clip.g.r1 + dr), c - (clip.g.c1 + dc)), s: snap.s } : { value: snap.value, formula: null, s: snap.s } });
    }
    if (items.length) this.apply(items);
  };

  // ======================================================================
  // Insert / delete cells, shifting their neighbours
  // ======================================================================

  // Move references that point into the strip of shifted cells. A range whose far end is in the strip grows or
  // shrinks with it; references entirely inside deleted cells become #REF!.
  function shiftRefs(f, own, target, axis, strip, delta, deleted) {
    const toks = scanRefs(f);
    if (!toks.length) return f;
    let out = '', last = 0;
    const along = (r, c) => (axis === 'col' ? c : r);
    const across = (r, c) => (axis === 'col' ? r : c);
    const band = axis === 'col' ? [strip.r1, strip.r2] : [strip.c1, strip.c2];
    const delA = deleted && (axis === 'col' ? [deleted.c1, deleted.c2] : [deleted.r1, deleted.r2]);
    const inDel = (p) => delA && p >= delA[0] && p <= delA[1];
    // New position along the shift axis for one endpoint (isEnd: the far end of a range).
    const moveP = (p, isEnd) => {
      if (inDel(p)) return isEnd ? delA[0] - 1 : delA[0];
      return p >= strip.from ? p + delta : p;
    };
    const fmt = (t, r, c) => `${t.colAbs ? '$' : ''}${colName(c)}${t.rowAbs ? '$' : ''}${r}`;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      if (t.kind !== 'ref') continue;
      const nxt = toks[i + 1];
      const pair = nxt && nxt.kind === 'ref' && !nxt.sheet && f.slice(t.end, nxt.refStart) === ':';
      if ((t.sheet || own).toLowerCase() !== target.toLowerCase()) { if (pair) i++; continue; }
      const ends = pair ? [t, nxt] : [t];
      // Only references lying wholly within the strip's rows (or columns) move.
      const lo = Math.min(...ends.map((x) => across(x.row, x.col))), hi = Math.max(...ends.map((x) => across(x.row, x.col)));
      if (lo < band[0] || hi > band[1]) { if (pair) i++; continue; }
      const ps = ends.map((x) => along(x.row, x.col));
      let np;
      if (pair) {
        const a = Math.min(...ps), b = Math.max(...ps);
        if (inDel(a) && inDel(b)) np = null;
        else { const na = moveP(a, false), nb = moveP(b, true); np = ps[0] <= ps[1] ? [na, nb] : [nb, na]; }
      } else np = inDel(ps[0]) ? null : [moveP(ps[0], false)];
      if (np === null) {
        out += f.slice(last, t.refStart) + '#REF!';
        last = pair ? nxt.end : t.end;
        if (pair) i++;
        continue;
      }
      if (np.some((p, k) => p !== ps[k])) {
        const pos = ends.map((x, k) => (axis === 'col' ? [x.row, np[k]] : [np[k], x.col]));
        out += f.slice(last, t.refStart) + fmt(t, pos[0][0], pos[0][1]) + (pair ? `:${fmt(nxt, pos[1][0], pos[1][1])}` : '');
        last = pair ? nxt.end : t.end;
      }
      if (pair) i++;
    }
    return out + f.slice(last);
  }

  // dir: 'right' | 'down' (insert) or 'left' | 'up' (delete)
  P.shiftCells = function (dir) {
    const g = this.expandMerges(this.sel);
    const s = this.sheet;
    const n = dir === 'right' || dir === 'left' ? g.c2 - g.c1 + 1 : g.r2 - g.r1 + 1;
    if (s.merges.some((m) => (dir === 'right' || dir === 'left' ? m.r2 >= g.r1 && m.r1 <= g.r2 && m.c2 >= g.c1 : m.c2 >= g.c1 && m.c1 <= g.c2 && m.r2 >= g.r1))) throw new Error('Unmerge the cells in the way first.');
    const axis = dir === 'right' || dir === 'left' ? 'col' : 'row';
    const insert = dir === 'right' || dir === 'down';
    const delta = insert ? n : -n;
    // Cells that move: in the selection's rows (or columns) from the selection onwards (after it, for deletes).
    const strip = axis === 'col' ? { r1: g.r1, r2: g.r2, from: insert ? g.c1 : g.c2 + 1 } : { c1: g.c1, c2: g.c2, from: insert ? g.r1 : g.r2 + 1 };
    const deleted = insert ? null : g;
    return this.snapshotOp(insert ? 'Insert cells' : 'Delete cells', () => {
      const moving = [];
      for (const cell of s.cells.values()) {
        const inRows = axis === 'col' ? cell.r >= g.r1 && cell.r <= g.r2 : cell.c >= g.c1 && cell.c <= g.c2;
        if (!inRows) continue;
        const pos = axis === 'col' ? cell.c : cell.r;
        if (pos >= strip.from || (!insert && pos >= (axis === 'col' ? g.c1 : g.r1))) moving.push(cell);
      }
      for (const c of moving) if (c.shared != null) s.unshare(c.shared);
      const snaps = moving.map((c) => ({ r: c.r, c: c.c, value: c.value, formula: c.formula, s: c.s }));
      for (const c of moving) s.setCell(c.r, c.c, { value: null, formula: null, s: 0 });
      for (const x of snaps) {
        const pos = axis === 'col' ? x.c : x.r;
        if (!insert && pos <= (axis === 'col' ? g.c2 : g.r2)) continue; // deleted cells
        const r = axis === 'row' ? x.r + delta : x.r, c = axis === 'col' ? x.c + delta : x.c;
        if (r < 1 || c < 1) continue;
        const cell = s.setCell(r, c, { value: x.value, formula: x.formula, s: x.s });
        if (x.formula) cell.formulaChanged = true;
      }
      // Formulas everywhere that point at moved (or deleted) cells.
      for (const sh of this.wb.sheets) {
        for (const [k, cell] of sh.cells) {
          if (!cell.formula) continue;
          const nf = shiftRefs(cell.formula, sh.name, s.name, axis, strip, delta, deleted);
          if (nf !== cell.formula) { if (cell.shared != null) sh.unshare(cell.shared); cell.formula = nf; cell.formulaChanged = true; sh.changed.add(k); }
        }
      }
      this.wb.formulasChanged = true;
    });
  };

  // ======================================================================
  // Group rows / columns (outline levels), collapse and expand
  // ======================================================================

  P.groupSel = function (axis, add) {
    const g = this.expandMerges(this.sel);
    const s = this.sheet;
    return this.snapshotOp(add ? `Group ${axis === 'row' ? 'rows' : 'columns'}` : `Ungroup ${axis === 'row' ? 'rows' : 'columns'}`, () => {
      if (axis === 'row') {
        for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 20000); r++) {
          const info = { ...(s.rows.get(r) || {}) };
          info.outline = Math.max(0, Math.min(7, (info.outline || 0) + (add ? 1 : -1)));
          if (!info.outline && !add) info.hidden = false;
          s.rows.set(r, info); s.rowsChanged.add(r);
        }
      } else {
        for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 1000); c++) {
          // Give the column its own <col> entry (keeping its width and hidden state), then change its level.
          const cur = s.cols.find((x) => c >= x.min && c <= x.max);
          s.setColWidth(c, cur && cur.width != null ? cur.width : (s.defaultColWidth || 9.140625));
          for (const col of s.cols) if (c >= col.min && c <= col.max) { col.outline = Math.max(0, Math.min(7, (col.outline || 0) + (add ? 1 : -1))); if (!col.outline && !add) col.hidden = false; col.el = null; }
        }
      }
      // Excel's outline buttons need the highest level in sheetFormatPr.
      const root = s.doc.documentElement;
      const fmt = [...root.childNodes].find((x) => x.localName === 'sheetFormatPr');
      if (fmt) {
        const maxRow = Math.max(0, ...[...s.rows.values()].map((i) => i.outline || 0));
        const maxCol = Math.max(0, ...s.cols.map((c) => c.outline || 0));
        if (maxRow) fmt.setAttribute('outlineLevelRow', maxRow); else fmt.removeAttribute('outlineLevelRow');
        if (maxCol) fmt.setAttribute('outlineLevelCol', maxCol); else fmt.removeAttribute('outlineLevelCol');
      }
      s.structureChanged = true;
    });
  };
  // The group (contiguous rows/columns with an outline level) around the active cell.
  function groupAround(view, axis) {
    const s = view.sheet;
    const level = (i) => (axis === 'row' ? (s.rows.get(i) || {}).outline || 0 : (s.cols.find((c) => i >= c.min && i <= c.max) || {}).outline || 0);
    const at = axis === 'row' ? view.active.r : view.active.c;
    if (!level(at)) return null;
    let a = at, b = at;
    while (a > 1 && level(a - 1)) a--;
    while (level(b + 1)) b++;
    return { a, b };
  }
  P.collapseGroup = function (axis, collapse) {
    const grp = groupAround(this, axis);
    if (!grp) throw new Error(`Select a cell in a group of ${axis === 'row' ? 'rows' : 'columns'} first (Data > Group).`);
    const s = this.sheet;
    return this.snapshotOp(collapse ? 'Collapse group' : 'Expand group', () => {
      for (let i = grp.a; i <= grp.b; i++) { if (axis === 'row') s.setRowHidden(i, collapse); else s.setColHidden(i, collapse); }
      if (axis === 'row') { const info = { ...(s.rows.get(grp.b + 1) || {}) }; info.collapsed = collapse; s.rows.set(grp.b + 1, info); s.rowsChanged.add(grp.b + 1); }
    });
  };

  // ======================================================================
  // Saved settings in the workbook: filter views, slicers, protected ranges
  // (docProps/custom.xml property "DocStudio.Sheets", ignored by other programs)
  // ======================================================================

  const CUSTOM = 'docProps/custom.xml';
  P.ddProps = async function () {
    if (this.wb._ddProps) return this.wb._ddProps;
    let props = { filterViews: [], slicers: [], protected: [] };
    const f = this.wb.zip.file(CUSTOM);
    if (f) {
      const d = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
      const p = [...d.getElementsByTagName('property')].find((x) => x.getAttribute('name') === 'DocStudio.Sheets');
      if (p) try { props = { ...props, ...JSON.parse(p.textContent) }; } catch { /* ignore */ }
    }
    this.wb._ddProps = props;
    return props;
  };
  P.saveDdProps = async function (props, label) {
    const json = JSON.stringify(props).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    await this.packageOp(label, async (pkg) => {
      let d = await pkg.doc(CUSTOM);
      if (!d) {
        d = new DOMParser().parseFromString('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>', 'application/xml');
        pkg.put(CUSTOM, d);
        const rels = await pkg.doc('_rels/.rels');
        const el = rels.createElementNS('http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship');
        const used = new Set([...rels.documentElement.children].map((x) => x.getAttribute('Id')));
        let n = 1; while (used.has(`rId${n}`)) n++;
        el.setAttribute('Id', `rId${n}`); el.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'); el.setAttribute('Target', CUSTOM);
        rels.documentElement.append(el);
        await pkg.override(CUSTOM, 'application/vnd.openxmlformats-officedocument.custom-properties+xml');
      }
      const root = d.documentElement;
      for (const p of [...root.childNodes].filter((x) => x.nodeType === 1 && x.getAttribute('name') === 'DocStudio.Sheets')) p.remove();
      const pids = [...root.childNodes].filter((x) => x.nodeType === 1).map((x) => +x.getAttribute('pid') || 1);
      const p = new DOMParser().parseFromString(`<property xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${Math.max(1, ...pids) + 1}" name="DocStudio.Sheets"><vt:lpwstr>${json}</vt:lpwstr></property>`, 'application/xml').documentElement;
      root.append(d.importNode(p, true));
    });
    this.wb._ddProps = props;
  };

  // ---------- filter views: named, saved filter settings for the sheet filter

  P.saveFilterView = async function (name) {
    const s = this.sheet;
    if (!s.autoFilter) throw new Error('Create a filter first (Data > Create a filter), set it up, then save it as a view.');
    const props = await this.ddProps();
    const view = { name, sheet: s.name, ref: s.autoFilter.ref, cols: [...s.autoFilter.cols].map(([c, f]) => [c, { values: [...f.values], blanks: !!f.blanks }]) };
    props.filterViews = props.filterViews.filter((v) => !(v.name === name && v.sheet === s.name)).concat(view);
    await this.saveDdProps(props, 'Save filter view');
  };
  P.applyFilterView = async function (name) {
    const props = await this.ddProps();
    const v = props.filterViews.find((x) => x.name === name && x.sheet === this.sheet.name);
    if (!v) throw new Error('That filter view is no longer here.');
    const s = this.sheet;
    const g = parseRange(v.ref);
    return this.snapshotOp(`Filter view: ${name}`, () => {
      s.autoFilter = { ref: v.ref, cols: new Map(v.cols.map(([c, f]) => [c, { values: new Set(f.values), blanks: f.blanks }])) };
      s.filterChanged = true;
      for (let r = g.r1 + 1; r <= g.r2; r++) {
        let show = true;
        for (const [fc, ff] of s.autoFilter.cols) { const t = this.cellText(r, fc); if (t === '' ? !ff.blanks : !ff.values.has(t)) { show = false; break; } }
        const info = s.rows.get(r);
        if (!!(info && info.hidden) !== !show) s.setRowHidden(r, !show);
      }
      s.structureChanged = true;
    });
  };
  P.deleteFilterView = async function (name) {
    const props = await this.ddProps();
    props.filterViews = props.filterViews.filter((v) => !(v.name === name && v.sheet === this.sheet.name));
    await this.saveDdProps(props, 'Delete filter view');
  };

  // ---------- slicers: a panel of buttons for one column of the sheet filter or a table

  P.addSlicer = async function () {
    const f = this.filters()[0] && (this.filterAt(parseRange(this.filters()[0].ref).r1, this.active.c) || this.filters()[0]);
    if (!f) throw new Error('Create a filter or a table first, then add a slicer for one of its columns.');
    const g = parseRange(f.ref);
    const c = this.active.c >= g.c1 && this.active.c <= g.c2 ? this.active.c : g.c1;
    const props = await this.ddProps();
    props.slicers = props.slicers.filter((x) => !(x.sheet === this.sheet.name && x.col === c)).concat({ sheet: this.sheet.name, col: c, ref: f.ref, table: f.kind === 'table' ? f.table.name : null });
    await this.saveDdProps(props, 'Add slicer');
    this.showSlicers();
  };
  P.showSlicers = async function () {
    const props = await this.ddProps();
    if (this.slicerBox) this.slicerBox.remove();
    const mine = props.slicers.filter((x) => x.sheet === this.sheet.name);
    if (!mine.length) return;
    const box = h('div', { class: 'sg-slicers' });
    for (const sl of mine) {
      const f = this.filters().find((x) => (sl.table ? x.kind === 'table' && x.table.name === sl.table : x.kind === 'sheet'));
      if (!f) continue;
      const g = parseRange(f.ref);
      const values = [...new Set(Array.from({ length: g.r2 - g.r1 }, (_, i) => this.cellText(g.r1 + 1 + i, sl.col)).filter((t) => t !== ''))].sort();
      const cur = f.cols.get(sl.col);
      const panel = h('div', { class: 'sg-slicer' }, h('div', { class: 'sg-slicer-head' }, h('b', { text: this.cellText(g.r1, sl.col) || colName(sl.col) }), h('button', { class: 'ghost', title: 'Clear', text: '⨯', onclick: () => this.applyFilter(g, sl.col, null, f).then(() => this.showSlicers()) }), h('button', { class: 'ghost', title: 'Remove slicer', text: '🗑', onclick: async () => { props.slicers = props.slicers.filter((x) => x !== sl); await this.saveDdProps(props, 'Remove slicer'); this.showSlicers(); } })));
      for (const v of values) {
        const on = !cur || cur.values.has(v);
        // Click shows just this value; Ctrl+click adds or removes it.
        panel.append(h('button', { class: `sg-slicer-item${on ? ' on' : ''}`, text: v, onclick: (e) => {
          let set;
          if (e.ctrlKey) { set = new Set(cur ? cur.values : values); if (set.has(v)) set.delete(v); else set.add(v); } else set = new Set([v]);
          this.applyFilter(g, sl.col, set.size ? { values: set, blanks: false } : null, f).then(() => this.showSlicers()).catch((err) => desk.error("Couldn't filter", err.message));
        } }));
      }
      box.append(panel);
    }
    if (getComputedStyle(this.host).position === 'static') this.host.style.position = 'relative';
    this.host.append(box);
    this.slicerBox = box;
  };

  // ---------- protected ranges (Doc Studio keeps them from being edited)

  P.protectRange = async function () {
    const g = this.effectiveRange();
    const props = await this.ddProps();
    props.protected = props.protected.filter((x) => !(x.sheet === this.sheet.name && x.ref === rangeRef(g))).concat({ sheet: this.sheet.name, ref: rangeRef(g) });
    await this.saveDdProps(props, 'Protect range');
    status(this, `Protected ${rangeRef(g)}`);
  };
  P.unprotectRange = async function () {
    const props = await this.ddProps();
    const { r, c } = this.active;
    props.protected = props.protected.filter((x) => { if (x.sheet !== this.sheet.name) return true; const g = parseRange(x.ref); return !(r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2); });
    await this.saveDdProps(props, 'Unprotect range');
  };
  const apply = P.apply;
  P.apply = function (items, sheet) {
    const props = this.wb._ddProps;
    const sh = sheet || this.sheet;
    if (props && props.protected.length) {
      const prot = props.protected.filter((x) => x.sheet === sh.name).map((x) => parseRange(x.ref));
      const hit = items.find((it) => prot.some((g) => it.r >= g.r1 && it.r <= g.r2 && it.c >= g.c1 && it.c <= g.c2));
      if (hit) { (this.opts.error || desk.error)(`${addr(hit.r, hit.c)} is in a protected range.`, 'Use Data > Protected ranges to unprotect it first.'); return; }
    }
    return apply.call(this, items, sheet);
  };

  // Load the saved settings when a workbook is shown (slicers appear, protection applies).
  const showSheet = P.showSheet;
  P.showSheet = function (i) {
    const out = showSheet.call(this, i);
    this.ddProps().then(() => this.showSlicers()).catch(() => {});
    return out;
  };

  // ======================================================================
  // Workbook theme (colours and fonts)
  // ======================================================================

  const THEMES = {
    Office: ['44546A', 'E7E6E6', ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'], 'Calibri Light', 'Calibri'],
    Ocean: ['17406D', 'DBEFF9', ['0F6FC6', '009DD9', '0BD0D9', '10CF9B', '7CCA62', 'A5C249'], 'Segoe UI Semibold', 'Segoe UI'],
    Forest: ['2C3C24', 'E5EADB', ['549E39', '8AB833', 'C0CF3A', '029676', '4AB5C4', '0989B1'], 'Georgia', 'Georgia'],
    Sunset: ['4E3B30', 'FBEEC9', ['F07F09', '9F2936', '1B587C', '4E8542', '604878', 'C19859'], 'Trebuchet MS', 'Trebuchet MS'],
    Berry: ['3F1D5A', 'EDE3F2', ['8E44AD', 'C0392B', 'D35400', '16A085', '2980B9', '7F8C8D'], 'Century Gothic', 'Century Gothic'],
    Mono: ['262626', 'EDEDED', ['595959', '7F7F7F', 'A5A5A5', '3F3F3F', '262626', 'BFBFBF'], 'Arial', 'Arial'],
  };
  SheetView.THEMES = Object.keys(THEMES);
  P.applyTheme = function (name) {
    const t = THEMES[name];
    if (!t) return;
    const wbPath = this.wb.wbPath;
    return this.packageOp(`Theme: ${name}`, async (pkg) => {
      const rels = await pkg.rels(wbPath);
      const rel = [...rels.documentElement.children].find((r) => r.getAttribute('Type').endsWith('/theme'));
      if (!rel) throw new Error('This workbook has no theme part.');
      const path = wbPath.replace(/[^/]+$/, '') + rel.getAttribute('Target');
      const d = await pkg.doc(path);
      const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
      const scheme = d.getElementsByTagNameNS(A, 'clrScheme')[0];
      const set = (n, v) => { const el = [...scheme.children].find((x) => x.localName === n); if (!el) return; while (el.firstChild) el.firstChild.remove(); const c = d.createElementNS(A, 'a:srgbClr'); c.setAttribute('val', v); el.append(c); };
      set('dk2', t[0]); set('lt2', t[1]);
      t[2].forEach((c, i) => set(`accent${i + 1}`, c));
      scheme.setAttribute('name', `Doc Studio ${name}`);
      const major = d.getElementsByTagNameNS(A, 'majorFont')[0], minor = d.getElementsByTagNameNS(A, 'minorFont')[0];
      const latin = (f) => f && [...f.children].find((x) => x.localName === 'latin');
      if (latin(major)) latin(major).setAttribute('typeface', t[3]);
      if (latin(minor)) latin(minor).setAttribute('typeface', t[4]);
      // Cells that use the theme's body font follow it (fonts marked with scheme="minor").
      const stylesRel = [...rels.documentElement.children].find((r) => r.getAttribute('Type').endsWith('/styles'));
      if (stylesRel) {
        const sp = wbPath.replace(/[^/]+$/, '') + stylesRel.getAttribute('Target');
        const sd = await pkg.doc(sp);
        for (const f of sd.getElementsByTagNameNS(NS, 'font')) {
          const sch = [...f.children].find((x) => x.localName === 'scheme');
          const nm = [...f.children].find((x) => x.localName === 'name');
          if (sch && nm) nm.setAttribute('val', sch.getAttribute('val') === 'major' ? t[3] : t[4]);
        }
      }
    });
  };

  // ======================================================================
  // Copy a sheet into another open workbook
  // ======================================================================

  P.copySheetTo = function (targetView) {
    const src = this.sheet, sw = this.wb, tw = targetView.wb;
    return targetView.snapshotOp('Copy sheet', () => {
      let name = src.name, n = 2;
      while (tw.sheet(name)) name = `${src.name} (${n++})`;
      const dst = tw.addSheet(name);
      const map = new Map();
      // Styles are rebuilt in the target from what they look like.
      const styleFor = (si) => {
        if (!si) return 0;
        if (map.has(si)) return map.get(si);
        const st = sw.styles, f = st.font(si), x = st.xf(si), b = st.border(si);
        const change = {
          font: { name: f.name, size: f.size, bold: !!f.bold, italic: !!f.italic, underline: !!f.underline, strike: !!f.strike, color: f.color && /^#[0-9a-f]{6}$/i.test(f.color) ? f.color : null },
          numFmt: st.numFmt(si),
          align: { horizontal: x.align.horizontal || null, vertical: x.align.vertical || null, wrap: !!x.align.wrap, textRotation: x.align.rotation ? String(x.align.rotation) : null },
        };
        const fill = st.fill(si);
        if (fill && /^#[0-9a-f]{6}$/i.test(fill)) change.fill = fill;
        const bd = {};
        for (const side of ['left', 'right', 'top', 'bottom']) if (b[side]) bd[side] = b[side].style;
        if (Object.keys(bd).length) change.border = bd;
        const id = tw.styles.derive(0, change);
        tw.styles.dirty = true;
        map.set(si, id);
        return id;
      };
      for (const cell of src.cells.values()) {
        if (cell.arrayMaster && !cell.formula) continue;
        dst.setCell(cell.r, cell.c, { value: cell.value, formula: cell.formula, s: styleFor(cell.s) });
        if (cell.formula) dst.get(cell.r, cell.c).formulaChanged = true;
      }
      for (const col of src.cols) if (col.width != null) for (let c = col.min; c <= Math.min(col.max, col.min + 300); c++) dst.setColWidth(c, col.width);
      for (const [r, info] of src.rows) if (info.ht != null) dst.setRowHeight(r, info.ht);
      dst.merges = src.merges.map((m) => ({ ...m }));
      dst.mergesChanged = !!dst.merges.length;
      tw.formulasChanged = true;
    }).then(() => targetView.showSheet(targetView.wb.sheets.length - 1));
  };

  // ======================================================================
  // Menus added to the workbook editor
  // ======================================================================

  SheetView.sheetMenus2 = function (tab, view) {
    const run = (label, fn) => async () => { try { if (view.editing) view.endEdit(true); await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } };
    const it = (label, fn, o) => ({ label, run: run(label.toLowerCase(), fn), ...(o || {}) });
    const sub = (label, items) => ({ label, sub: items });
    const SEP = { sep: true };
    return {
      Edit: [sub('Delete cells', [it('Delete cells and shift left', () => view.shiftCells('left')), it('Delete cells and shift up', () => view.shiftCells('up'))]), sub('Paste special', [it('Values only', () => view.pasteSpecial('values'), { key: 'Ctrl+Shift+V' }), it('Format only', () => view.pasteSpecial('formats')), it('Formulas only', () => view.pasteSpecial('formulas')), it('Transposed', () => view.pasteSpecial('transpose')), it('Column widths only', () => view.pasteSpecial('widths'))])],
      Insert: [sub('Cells', [it('Insert cells and shift right', () => view.shiftCells('right')), it('Insert cells and shift down', () => view.shiftCells('down'))]),
        it('Drawing…', async () => { const p = await DocsExtra.drawing('Drawing'); if (p) await view.insertPicture(p.bytes, 'png', p.width, p.height); })],
      Format: [sub('Theme', SheetView.THEMES.map((n) => it(n, () => view.applyTheme(n))))],
      Data: [
        SEP,
        sub('Filter views', [
          it('Save current filter as a view…', async () => { const v = await DocTools.form('Save filter view', [{ name: 'n', label: 'Name', value: 'Filter view 1' }], 'Save'); if (v && v.n.trim()) await view.saveFilterView(v.n.trim()); }),
          { label: 'Apply a view', sub: async () => { const p = await view.ddProps(); const l = p.filterViews.filter((x) => x.sheet === view.sheet.name); return l.length ? l.map((x) => it(x.name, () => view.applyFilterView(x.name))) : [{ label: 'No saved views', run: null, enabled: () => false }]; } },
          { label: 'Delete a view', sub: async () => { const p = await view.ddProps(); const l = p.filterViews.filter((x) => x.sheet === view.sheet.name); return l.length ? l.map((x) => it(x.name, () => view.deleteFilterView(x.name))) : [{ label: 'No saved views', run: null, enabled: () => false }]; } },
        ]),
        it('Add a slicer', () => view.addSlicer()),
        sub('Protected ranges', [it('Protect the selected range', () => view.protectRange()), it('Unprotect the range at the cursor', () => view.unprotectRange())]),
        SEP,
        sub('Group', [it('Group rows', () => view.groupSel('row', true), { key: 'Alt+Shift+→' }), it('Ungroup rows', () => view.groupSel('row', false), { key: 'Alt+Shift+←' }), it('Group columns', () => view.groupSel('col', true)), it('Ungroup columns', () => view.groupSel('col', false)), SEP,
          it('Collapse row group', () => view.collapseGroup('row', true)), it('Expand row group', () => view.collapseGroup('row', false)), it('Collapse column group', () => view.collapseGroup('col', true)), it('Expand column group', () => view.collapseGroup('col', false))]),
      ],
      Tools: [{ label: 'Copy this sheet to', sub: () => {
        const others = tabs.filter((t) => t !== tab && t.kind === 'sheet' && t.sheetView);
        return others.length ? others.map((t) => it(t.name, () => view.copySheetTo(t.sheetView))) : [{ label: 'Open another workbook first', run: null, enabled: () => false }];
      } }],
    };
  };
  const baseMenus = SheetView.sheetMenus;
  SheetView.sheetMenus = function (tab, view) {
    const m = baseMenus(tab, view);
    for (const [k, list] of Object.entries(SheetView.sheetMenus2(tab, view))) m[k] = [...(m[k] || []), ...(list[0] && list[0].sep ? [] : [{ sep: true }]), ...list];
    return m;
  };

  // Shortcuts: Alt+Shift+→/← group/ungroup rows; Ctrl+Shift+V paste values.
  document.addEventListener('keydown', (e) => {
    const v = active && active.sheetView;
    if (!v || v.editing || !v.grid.contains(document.activeElement)) return;
    const go = (fn) => { e.preventDefault(); e.stopPropagation(); Promise.resolve(fn()).catch((err) => desk.error("Couldn't do that", err.message)); };
    if (e.altKey && e.shiftKey && e.key === 'ArrowRight') go(() => v.groupSel('row', true));
    else if (e.altKey && e.shiftKey && e.key === 'ArrowLeft') go(() => v.groupSel('row', false));
    else if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'v' && v.clip) go(() => v.pasteSpecial('values'));
  }, true);
})();
