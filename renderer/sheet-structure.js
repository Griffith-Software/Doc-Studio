'use strict';
// Structural sheet edits: insert/delete rows and columns, keeping every reference in the workbook correct
// (formulas on all sheets, defined names, merges, widths/heights, conditional formatting, data validation,
// hyperlinks, autofilters, tables and chart series).

const SheetStructure = (() => {
  const { key, addr, colName, colNum, parseRange, rangeRef, scanRefs, MAX_ROW, MAX_COL } = SheetModel;

  // ---------- reference arithmetic

  // Adjust one position or range on an axis. Returns null when everything it pointed at was deleted.
  function adjustSpan(a, b, at, delta) {
    if (delta > 0) return [a >= at ? a + delta : a, b >= at ? b + delta : b];
    const n = -delta, end = at + n - 1;
    const na = a < at ? a : a > end ? a - n : at;
    const nb = b < at ? b : b > end ? b - n : at - 1;
    return nb < na ? null : [na, nb];
  }

  // Rewrite references to `target` sheet inside formula text `f` written on sheet `own`.
  function adjustFormula(f, own, target, axis, at, delta) {
    const toks = scanRefs(f);
    if (!toks.length) return f;
    const tgt = target.toLowerCase();
    const sheetOf = (t) => (t.sheet || own).toLowerCase();
    let out = '';
    let last = 0;
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i];
      const nxt = toks[i + 1];
      // Ranges: "A1:B2" (two refs joined by ':'), "A:C" and "1:5" (paired tokens).
      const isPair = nxt && f.slice(t.end, nxt.refStart) === ':' && !nxt.sheet && (nxt.kind === t.kind || (t.kind === 'ref' && nxt.kind === 'ref'));
      if (sheetOf(t) !== tgt) { if (isPair) i++; continue; }
      out += f.slice(last, t.refStart);
      if (t.kind === 'ref') {
        const r2 = isPair ? nxt.row : t.row, c2 = isPair ? nxt.col : t.col;
        let rows = [t.row, r2], cols = [t.col, c2];
        if (axis === 'row') rows = adjustSpan(Math.min(t.row, r2), Math.max(t.row, r2), at, delta);
        else cols = adjustSpan(Math.min(t.col, c2), Math.max(t.col, c2), at, delta);
        if (!rows || !cols) { out += '#REF!'; last = isPair ? nxt.end : t.end; if (isPair) i++; continue; }
        const one = (abs1, c, abs2, r) => `${abs1 ? '$' : ''}${colName(c)}${abs2 ? '$' : ''}${r}`;
        if (isPair) {
          const swappedR = axis === 'row' && t.row > r2, swappedC = axis === 'col' && t.col > c2;
          const [ra, rb] = swappedR ? [rows[1], rows[0]] : rows;
          const [ca, cb] = swappedC ? [cols[1], cols[0]] : cols;
          out += one(t.colAbs, axis === 'col' ? ca : t.col, t.rowAbs, axis === 'row' ? ra : t.row) + ':' + one(nxt.colAbs, axis === 'col' ? cb : nxt.col, nxt.rowAbs, axis === 'row' ? rb : nxt.row);
          last = nxt.end; i++;
        } else {
          out += one(t.colAbs, cols[0], t.rowAbs, rows[0]);
          last = t.end;
        }
      } else if (t.kind === 'col') {
        const c2 = nxt ? nxt.col : t.col;
        const cols = axis === 'col' ? adjustSpan(Math.min(t.col, c2), Math.max(t.col, c2), at, delta) : [t.col, c2];
        if (!cols) out += '#REF!';
        else out += `${t.colAbs ? '$' : ''}${colName(cols[0])}:${nxt && nxt.colAbs ? '$' : ''}${colName(cols[1])}`;
        last = nxt ? nxt.end : t.end; if (nxt) i++;
      } else {
        const r2 = nxt ? nxt.row : t.row;
        const rows = axis === 'row' ? adjustSpan(Math.min(t.row, r2), Math.max(t.row, r2), at, delta) : [t.row, r2];
        if (!rows) out += '#REF!';
        else out += `${t.rowAbs ? '$' : ''}${rows[0]}:${nxt && nxt.rowAbs ? '$' : ''}${rows[1]}`;
        last = nxt ? nxt.end : t.end; if (nxt) i++;
      }
    }
    return out + f.slice(last);
  }

  // "A1:B2 D4" style lists (sqref). Ranges that vanish are dropped; returns '' when nothing is left.
  function adjustSqref(list, axis, at, delta) {
    return list.split(/\s+/).filter(Boolean).map((part) => {
      const g = parseRange(part);
      if (!g) return part;
      const rows = axis === 'row' ? adjustSpan(g.r1, g.r2, at, delta) : [g.r1, g.r2];
      const cols = axis === 'col' ? adjustSpan(g.c1, g.c2, at, delta) : [g.c1, g.c2];
      if (!rows || !cols) return null;
      return rangeRef({ r1: rows[0], r2: Math.min(rows[1], MAX_ROW), c1: cols[0], c2: Math.min(cols[1], MAX_COL) });
    }).filter(Boolean).join(' ');
  }

  // ---------- the operation

  // axis 'row' | 'col'; delta > 0 inserts `delta` rows/cols before `at`; delta < 0 deletes -delta starting at `at`.
  async function shift(wb, sheet, axis, at, delta) {
    const target = sheet.name;
    const n = Math.abs(delta);
    const inDeleted = (p) => delta < 0 && p >= at && p < at + n;
    const moved = (p) => (p >= at ? (delta > 0 ? p + n : p - n) : p);

    // 1. Shared formulas in this sheet become plain formulas (their positions are about to change).
    for (const [k, cell] of sheet.cells) if (cell.shared != null) { delete cell.shared; cell.formulaChanged = true; sheet.changed.add(k); }

    // 2. Formulas everywhere that point at this sheet.
    for (const s of wb.sheets) {
      if (s.unsupported) continue;
      for (const [k, cell] of s.cells) {
        if (!cell.formula) continue;
        const nf = adjustFormula(cell.formula, s.name, target, axis, at, delta);
        if (nf !== cell.formula) { if (cell.shared != null) s.unshare(cell.shared); cell.formula = nf; cell.formulaChanged = true; s.changed.add(k); }
      }
    }
    for (const d of wb.definedNames) {
      const nf = adjustFormula(d.ref, target, target, axis, at, delta);
      if (nf !== d.ref) { d.ref = nf; d.el.textContent = nf; wb.wbChanged = true; }
    }

    // 3. Cells move (their XML elements move with them).
    const pos = (cell) => (axis === 'row' ? cell.r : cell.c);
    const next = new Map();
    for (const cell of sheet.cells.values()) {
      if (inDeleted(pos(cell))) { if (cell.el && cell.el.parentNode) cell.el.parentNode.removeChild(cell.el); continue; }
      if (axis === 'row') cell.r = moved(cell.r); else cell.c = moved(cell.c);
      if (cell.el) cell.el.setAttribute('r', addr(cell.r, cell.c));
      if (cell.array) {
        const rows = axis === 'row' ? adjustSpan(cell.array.r1, cell.array.r2, at, delta) : [cell.array.r1, cell.array.r2];
        const cols = axis === 'col' ? adjustSpan(cell.array.c1, cell.array.c2, at, delta) : [cell.array.c1, cell.array.c2];
        if (rows && cols) cell.array = { r1: rows[0], r2: rows[1], c1: cols[0], c2: cols[1] };
        cell.formulaChanged = true;
      }
      next.set(key(cell.r, cell.c), cell);
    }
    sheet.cells = next;
    for (const cell of sheet.cells.values()) {
      if (!cell.arrayMaster) continue;
      const master = [...sheet.cells.values()].find((m) => m.array && cell.r >= m.array.r1 && cell.r <= m.array.r2 && cell.c >= m.array.c1 && cell.c <= m.array.c2);
      if (master) cell.arrayMaster = key(master.r, master.c); else delete cell.arrayMaster;
    }
    // Every moved formula's cached position changed: rewrite them on save.
    for (const [k, cell] of sheet.cells) if (cell.formula || cell.formulaChanged) sheet.changed.add(k);
    sheet.changed = new Set([...sheet.changed].filter((k) => sheet.cells.has(k)));

    // 4. Row elements / row info.
    if (axis === 'row') {
      const rowEls = new Map();
      for (const [r, el] of sheet.rowEls) {
        if (inDeleted(r)) { if (el.parentNode) el.parentNode.removeChild(el); continue; }
        const nr = moved(r);
        el.setAttribute('r', nr);
        el.removeAttribute('spans');
        rowEls.set(nr, el);
      }
      sheet.rowEls = rowEls;
      const rows = new Map();
      for (const [r, info] of sheet.rows) if (!inDeleted(r)) rows.set(moved(r), info);
      sheet.rows = rows;
      if (sheet.autoHt) { const a = new Map(); for (const [r, h] of sheet.autoHt) if (!inDeleted(r)) a.set(moved(r), h); sheet.autoHt = a; }
    } else {
      for (const el of sheet.rowEls.values()) el.removeAttribute('spans');
      const cols = [];
      for (const col of sheet.cols) {
        if (delta > 0) {
          // A width range containing the insert point grows (new columns take the formatting to their left).
          if (col.min >= at) cols.push({ ...col, min: col.min + n, max: col.max + n, el: null });
          else if (col.max >= at) cols.push({ ...col, max: col.max + n, el: null });
          else cols.push(col);
        } else {
          const span = adjustSpan(col.min, col.max, at, delta);
          if (span) cols.push({ ...col, min: span[0], max: span[1], el: null });
        }
      }
      sheet.cols = cols;
      sheet.colsChanged = true;
    }

    // 5. Merges.
    sheet.merges = sheet.merges.map((m) => {
      const rows = axis === 'row' ? adjustSpan(m.r1, m.r2, at, delta) : [m.r1, m.r2];
      const cols = axis === 'col' ? adjustSpan(m.c1, m.c2, at, delta) : [m.c1, m.c2];
      if (!rows || !cols || (rows[0] === rows[1] && cols[0] === cols[1])) return null;
      return { r1: rows[0], r2: rows[1], c1: cols[0], c2: cols[1] };
    }).filter(Boolean);
    sheet.mergesChanged = true;
    if (sheet.autoFilter) {
      const g = parseRange(adjustSqref(sheet.autoFilter.ref, axis, at, delta) || '');
      sheet.autoFilter.ref = g ? rangeRef(g) : null;
      if (axis === 'col') {
        const cols = new Map();
        for (const [c, v] of sheet.autoFilter.cols) if (!inDeleted(c)) cols.set(moved(c), v);
        sheet.autoFilter.cols = cols;
      }
      sheet.filterChanged = true;
    }

    // 6. Everything else in the worksheet XML that names cells: conditional formatting, data validation,
    //    hyperlinks, autoFilter, and their Excel 2010 extension copies.
    for (const s of wb.sheets) {
      if (s.unsupported || !s.doc) continue;
      const self = s === sheet;
      walkRefs(s.doc.documentElement, s, target, self, axis, at, delta);
    }

    // 7. Tables on this sheet and chart series anywhere.
    await adjustTables(wb, sheet, axis, at, delta);
    await adjustCharts(wb, target, axis, at, delta);

    // 8. Extents.
    sheet.maxRow = 0; sheet.maxCol = 0;
    for (const c of sheet.cells.values()) { if (c.r > sheet.maxRow) sheet.maxRow = c.r; if (c.c > sheet.maxCol) sheet.maxCol = c.c; }
    sheet.structureChanged = true;
    wb.formulasChanged = true;
  }

  const SKIP = new Set(['sheetData', 'cols', 'mergeCells', 'dimension', 'sheetViews', 'autoFilter']);
  function walkRefs(el, sheet, target, self, axis, at, delta) {
    for (const ch of [...el.childNodes]) {
      if (ch.nodeType !== 1) continue;
      if (SKIP.has(ch.localName) && ch.parentNode === sheet.doc.documentElement) continue;
      if (self) {
        for (const attr of ['sqref', 'ref']) {
          if (!ch.hasAttribute(attr)) continue;
          const v = adjustSqref(ch.getAttribute(attr), axis, at, delta);
          if (!v) { ch.parentNode.removeChild(ch); break; } // the whole rule/validation/link was deleted
          ch.setAttribute(attr, v);
        }
        if (!ch.parentNode) continue;
        if (ch.localName === 'sqref' && ch.childElementCount === 0) {
          const v = adjustSqref(ch.textContent, axis, at, delta);
          if (!v) { const holder = ch.parentNode; holder.parentNode && holder.parentNode.removeChild(holder); continue; }
          ch.textContent = v;
        }
      }
      if (['formula', 'formula1', 'formula2', 'f'].includes(ch.localName) && ch.childElementCount === 0 && ch.parentNode && ch.parentNode.localName !== 'c') {
        const nf = adjustFormula(ch.textContent, sheet.name, target, axis, at, delta);
        if (nf !== ch.textContent) ch.textContent = nf;
        continue;
      }
      walkRefs(ch, sheet, target, self, axis, at, delta);
    }
  }

  async function adjustTables(wb, sheet, axis, at, delta) {
    const rels = await wb.sheetRels(sheet);
    for (const rel of rels.list.filter((r) => r.type.endsWith('/table'))) {
      const doc = await wb.part(rel.path);
      if (!doc) continue;
      const t = doc.documentElement;
      const before = parseRange(t.getAttribute('ref'));
      const ref = adjustSqref(t.getAttribute('ref'), axis, at, delta);
      if (!ref) continue; // table fully deleted: Excel would remove it; leave the part (rare)
      t.setAttribute('ref', ref);
      for (const af of t.getElementsByTagNameNS(t.namespaceURI, 'autoFilter')) af.setAttribute('ref', adjustSqref(af.getAttribute('ref'), axis, at, delta) || ref);
      const after = parseRange(ref);
      if (axis === 'col' && before && after) {
        const colsEl = t.getElementsByTagNameNS(t.namespaceURI, 'tableColumns')[0];
        const list = [...colsEl.children];
        if (delta > 0 && at > before.c1 && at <= before.c2 + 1) {
          // New columns inside a table get names, ids and header cells, as Excel requires.
          let id = Math.max(...list.map((c) => +c.getAttribute('id'))) + 1;
          const insertBefore = list[at - before.c1] || null;
          for (let i = 0; i < Math.abs(delta); i++) {
            let n = list.length + i + 1, name = `Column${n}`;
            while (list.some((c) => c.getAttribute('name') === name)) name = `Column${++n}`;
            const tc = doc.createElementNS(t.namespaceURI, (t.prefix ? t.prefix + ':' : '') + 'tableColumn');
            tc.setAttribute('id', id++); tc.setAttribute('name', name);
            colsEl.insertBefore(tc, insertBefore);
            list.push(tc);
            const hr = after.r1, hc = at + i;
            sheet.setCell(hr, hc, { value: name });
          }
        } else if (delta < 0) {
          const kill = [];
          for (let c = at; c < at - delta; c++) if (c >= before.c1 && c <= before.c2) kill.push(list[c - before.c1]);
          for (const k of kill) if (k) colsEl.removeChild(k);
        }
        colsEl.setAttribute('count', colsEl.children.length);
      }
      wb.markPart(rel.path);
    }
  }

  async function adjustCharts(wb, target, axis, at, delta) {
    for (const s of wb.sheets) {
      if (s.unsupported) continue;
      for (const chartPath of await wb.chartPaths(s)) {
        const doc = await wb.part(chartPath);
        if (!doc) continue;
        let changed = false;
        for (const f of doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/chart', 'f')) {
          const nf = adjustFormula(f.textContent, target, target, axis, at, delta);
          if (nf !== f.textContent) { f.textContent = nf; changed = true; }
        }
        if (changed) wb.markPart(chartPath);
      }
    }
  }

  return { shift, adjustFormula, adjustSqref };
})();
