'use strict';
// Recalculation for a SheetModel workbook: a dependency graph over formula cells, built with the engine's
// dependency parser, so an edit recalculates only the cells that depend on it (in the right order).

const SheetCalc = (() => {
  const { key, XErr } = SheetModel;
  const VOLATILE = /\b(NOW|TODAY|RAND|RANDBETWEEN|INDIRECT|OFFSET|CELL|INFO)\s*\(/i;

  class Calc {
    constructor(wb) {
      this.wb = wb;
      this.engine = DDFormula.createEngine({
        cell: (sheet, r, c) => this.cellValue(sheet, r, c),
        range: (sheet, r1, c1, r2, c2) => this.rangeValues(sheet, r1, c1, r2, c2),
        name: (name, sheet) => this.wb.resolveName(name, sheet),
        functions: { _XREF: (book, sheet, ref) => this.externalValue(book.value, sheet.value, ref.value) },
      });
      this.build();
    }

    cellValue(sheetName, r, c) {
      const s = this.wb.sheet(sheetName);
      if (!s) return this.engine.error('#REF!');
      const cell = s.get(r, c);
      if (!cell) return null;
      const v = cell.value;
      if (v instanceof XErr) return this.engine.error(v.code);
      return v === undefined ? null : v;
    }

    rangeValues(sheetName, r1, c1, r2, c2) {
      const s = this.wb.sheet(sheetName);
      if (!s) throw this.engine.error('#REF!');
      // Whole-column/row references stop at the used range.
      r2 = Math.min(r2, Math.max(s.maxRow, r1));
      c2 = Math.min(c2, Math.max(s.maxCol, c1));
      const out = [];
      for (let r = r1; r <= r2; r++) {
        const row = new Array(c2 - c1 + 1).fill(null);
        for (let c = c1; c <= c2; c++) {
          const cell = s.cells.get(key(r, c));
          if (cell) { const v = cell.value; row[c - c1] = v instanceof XErr ? this.engine.error(v.code) : (v === undefined ? null : v); }
        }
        out.push(row);
      }
      return out;
    }

    // ---- external workbooks: [n]Sheet!A1 becomes _XREF(n,"Sheet","A1"), answered from Excel's cached values.

    // Excel-only syntax becomes plain references the engine understands: _xlfn./_xlws. prefixes, table
    // ("structured") references, and links to other workbooks.
    prep(f, sheet, r, c) {
      if (!f) return f;
      if (f.indexOf('_xl') >= 0) f = f.replace(/_xlfn\.|_xlws\.|_xlpm\./g, '');
      if (f.indexOf('[') < 0) return f;
      f = this.structured(f, sheet, r, c);
      if (f.indexOf('[') < 0) return f;
      return f.replace(/(?:'\[(\d+)\]((?:[^']|'')+)'|\[(\d+)\]([A-Za-z0-9_.]+))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)/g,
        (m, b1, s1, b2, s2, ref) => `_XREF(${b1 || b2},"${(s1 ? s1.replace(/''/g, "'") : s2).replace(/"/g, '""')}","${ref.replace(/\$/g, '')}")`);
    }

    structured(f, sheet, r, c) {
      const tables = this.wb.tables || [];
      if (!tables.length) return f;
      const here = sheet && r ? this.wb.tableAt(sheet, r, c) : null;
      const q = (s) => SheetModel.quoteSheet(s);
      const A = SheetModel.addr;
      const area = (t, spec) => {
        const g = t.g;
        const dataR1 = g.r1 + t.header, dataR2 = g.r2 - t.totals;
        let c1 = g.c1, c2 = g.c2, r1 = dataR1, r2 = dataR2;
        const items = [];
        const parts = spec.match(/\[[^\[\]]*\]/g) || [`[${spec}]`];
        let colsSeen = [];
        for (const p of parts) {
          const v = p.slice(1, -1).trim();
          if (/^#all$/i.test(v)) items.push('all');
          else if (/^#data$/i.test(v)) items.push('data');
          else if (/^#headers$/i.test(v)) items.push('headers');
          else if (/^#totals$/i.test(v)) items.push('totals');
          else if (/^#this row$/i.test(v)) items.push('row');
          else colsSeen.push(v.replace(/^@/, '').replace(/'(.)/g, '$1'));
        }
        if (/^@/.test(spec.trim())) { items.push('row'); colsSeen = [spec.trim().slice(1).replace(/^\[|\]$/g, '')]; }
        if (colsSeen.length) {
          const idx = colsSeen.map((n) => t.cols.findIndex((x) => x.toLowerCase() === n.toLowerCase()));
          if (idx.some((i) => i < 0)) return '#REF!';
          c1 = g.c1 + Math.min(...idx); c2 = g.c1 + Math.max(...idx);
        }
        if (items.includes('all')) { r1 = g.r1; r2 = g.r2; }
        if (items.includes('headers')) { r1 = g.r1; r2 = items.includes('data') ? dataR2 : g.r1; }
        if (items.includes('totals')) { r2 = g.r2; if (!items.includes('data')) r1 = g.r2; }
        if (items.includes('row')) { if (!r) return '#REF!'; r1 = r2 = r; }
        const prefix = t.sheet === sheet ? '' : q(t.sheet.name) + '!';
        return r1 === r2 && c1 === c2 ? prefix + A(r1, c1) : `${prefix}${A(r1, c1)}:${A(r2, c2)}`;
      };
      // Table[...] (with nested brackets) and bare [@Col] / [[#This Row],[Col]] inside a table.
      let out = '';
      let i = 0;
      while (i < f.length) {
        const ch = f[i];
        if (ch === '"') { let j = i + 1; while (j < f.length && !(f[j] === '"' && f[j + 1] !== '"')) j += f[j] === '"' ? 2 : 1; out += f.slice(i, j + 1); i = j + 1; continue; }
        const m = /^([A-Za-z_\\][A-Za-z0-9_.]*)?\[/.exec(f.slice(i));
        if (m && (m[1] || here) && !(m[1] && /^\d/.test(m[1]))) {
          const t = m[1] ? tables.find((x) => x.name.toLowerCase() === m[1].toLowerCase()) : here;
          // External workbook references like [1]Sheet!A1 are not table references.
          if (!t && (!m[1] || /^\[\d+\]/.test(f.slice(i + (m[1] || '').length)))) { out += ch; i++; continue; }
          if (!t) { out += ch; i++; continue; }
          let j = i + (m[1] ? m[1].length : 0), depth = 0;
          for (; j < f.length; j++) { if (f[j] === '[') depth++; else if (f[j] === ']') { depth--; if (!depth) break; } }
          const spec = f.slice(i + (m[1] ? m[1].length : 0) + 1, j);
          out += area(t, spec);
          i = j + 1;
          continue;
        }
        out += ch; i++;
      }
      return out;
    }

    externalValue(book, sheetName, ref) {
      const ext = this.wb.externals[book - 1];
      if (!ext) throw this.engine.error('#REF!');
      const si = ext.sheets.findIndex((n) => n.toLowerCase() === String(sheetName).toLowerCase());
      const data = ext.data[si] || new Map();
      const g = SheetModel.parseRange(ref);
      const get = (r, c) => { const v = data.get(key(r, c)); return v instanceof XErr ? this.engine.error(v.code) : (v === undefined ? null : v); };
      if (g.r1 === g.r2 && g.c1 === g.c2) return get(g.r1, g.c1);
      const out = [];
      for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(get(r, c)); out.push(row); }
      return out;
    }

    // ---- graph

    nodeId(sheet, cell) { return `${sheet.name}\u0000${key(cell.r, cell.c)}`; }

    build() {
      this.nodes = new Map();     // id -> { sheet, cell, deps }
      this.points = new Map();    // "sheet\0key" -> Set(node ids)
      this.ranges = [];           // { sheet, r1, c1, r2, c2, id }
      this.volatile = new Set();
      for (const s of this.wb.sheets) {
        if (s.unsupported) continue;
        for (const cell of s.cells.values()) if (cell.formula && !cell.arrayMaster) this.add(s, cell);
      }
    }

    add(sheet, cell) {
      const id = this.nodeId(sheet, cell);
      const deps = this.engine.dependencies(this.prep(cell.formula, sheet, cell.r, cell.c), sheet.name, cell.r, cell.c);
      const node = { id, sheet, cell, deps };
      this.nodes.set(id, node);
      for (const d of deps) {
        const sh = (d.sheet || sheet.name).toLowerCase();
        if (d.from) this.ranges.push({ sheet: sh, r1: d.from.row, c1: d.from.col, r2: d.to.row, c2: d.to.col, id });
        else {
          const pk = `${sh}\u0000${key(d.row, d.col)}`;
          if (!this.points.has(pk)) this.points.set(pk, new Set());
          this.points.get(pk).add(id);
        }
      }
      if (VOLATILE.test(cell.formula)) this.volatile.add(id);
    }

    remove(sheet, cell) {
      const id = this.nodeId(sheet, cell);
      if (!this.nodes.has(id)) return;
      this.nodes.delete(id);
      for (const set of this.points.values()) set.delete(id);
      this.ranges = this.ranges.filter((r) => r.id !== id);
      this.volatile.delete(id);
    }

    // The formula of a cell changed (or appeared/disappeared).
    formulaChanged(sheet, cell) {
      this.remove(sheet, cell);
      if (!cell.formula && cell.spill) this.clearSpill(sheet, cell, null);
      if (cell.formula && !cell.arrayMaster) this.add(sheet, cell);
    }

    dependentsOf(sheetName, r, c) {
      const sh = sheetName.toLowerCase();
      const out = new Set(this.points.get(`${sh}\u0000${key(r, c)}`) || []);
      for (const g of this.ranges) if (g.sheet === sh && r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) out.add(g.id);
      return out;
    }

    // Recalculate everything that depends on the changed cells (plus volatile formulas). Returns the set of
    // [sheet, cell] pairs whose value changed, for repainting and saving.
    recalc(changed, opts) {
      const dirty = new Set(opts && opts.all ? this.nodes.keys() : []);
      const queue = [];
      for (const ch of changed || []) {
        const id = `${ch.sheet.name}\u0000${key(ch.r, ch.c)}`;
        if (this.nodes.has(id)) dirty.add(id);
        queue.push([ch.sheet.name, ch.r, ch.c]);
      }
      if (!opts || opts.volatile !== false) for (const id of this.volatile) dirty.add(id);
      for (const id of dirty) { const n = this.nodes.get(id); queue.push([n.sheet.name, n.cell.r, n.cell.c]); }
      while (queue.length) {
        const [sh, r, c] = queue.pop();
        for (const id of this.dependentsOf(sh, r, c)) {
          if (dirty.has(id)) continue;
          dirty.add(id);
          const n = this.nodes.get(id);
          queue.push([n.sheet.name, n.cell.r, n.cell.c]);
          if (n.cell.array) for (let rr = n.cell.array.r1; rr <= n.cell.array.r2; rr++) for (let cc = n.cell.array.c1; cc <= n.cell.array.c2; cc++) queue.push([n.sheet.name, rr, cc]);
        }
      }
      return this.evaluate(dirty);
    }

    // Formulas without a cached value (files written by other programs) get calculated when opened.
    initial() {
      const missing = [...this.nodes.values()].filter((n) => !n.cell.cached).map((n) => ({ sheet: n.sheet, r: n.cell.r, c: n.cell.c }));
      return this.recalc(missing, { volatile: false });
    }

    evaluate(dirty) {
      const state = new Map(); // id -> 1 visiting, 2 done
      const order = [];
      const cycles = new Set();
      const depsInDirty = (node) => {
        const out = [];
        for (const d of node.deps) {
          const sh = (d.sheet || node.sheet.name).toLowerCase();
          if (d.from) {
            for (const id of dirty) {
              const n = this.nodes.get(id);
              if (n.sheet.name.toLowerCase() === sh && n.cell.r >= d.from.row && n.cell.r <= d.to.row && n.cell.c >= d.from.col && n.cell.c <= d.to.col) out.push(id);
            }
          } else {
            const s = this.wb.sheet(sh);
            if (!s) continue;
            const id = `${s.name}\u0000${key(d.row, d.col)}`;
            if (dirty.has(id)) out.push(id);
            const target = s.cells.get(key(d.row, d.col));
            if (target && target.arrayMaster) { const mid = `${s.name}\u0000${target.arrayMaster}`; if (dirty.has(mid)) out.push(mid); }
          }
        }
        return out;
      };
      // Depth-first ordering with an explicit stack (long chains such as running totals would overflow recursion).
      for (const start of dirty) {
        if (state.get(start) === 2) continue;
        const stack = [[start, null]];
        while (stack.length) {
          const top = stack[stack.length - 1];
          const [id] = top;
          if (top[1] === null) {
            if (state.get(id) === 2) { stack.pop(); continue; }
            state.set(id, 1);
            top[1] = depsInDirty(this.nodes.get(id));
          }
          let pushed = false;
          while (top[1].length) {
            const d = top[1].pop();
            if (d === id) { cycles.add(id); continue; }
            const st = state.get(d);
            if (st === 1) { cycles.add(d); continue; }
            if (st === 2) continue;
            stack.push([d, null]);
            pushed = true;
            break;
          }
          if (!pushed) { state.set(id, 2); order.push(id); stack.pop(); }
        }
      }
      const updated = [];
      for (const id of order) {
        const n = this.nodes.get(id);
        const { cell, sheet } = n;
        let v;
        if (cycles.has(id)) v = 0;
        // Only dynamic-array formulas (typed here, or flagged in the file) may return several values; older
        // formulas keep Excel's implicit single-value behaviour.
        else v = this.engine.evaluate(this.prep(cell.formula, sheet, cell.r, cell.c), sheet.name, cell.r, cell.c, !!(cell.array || cell.dynamic));
        if (cell.array && Array.isArray(v)) {
          const g = cell.array;
          for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
            const row = v[r - g.r1];
            const x = row ? row[c - g.c1] : undefined;
            const target = r === cell.r && c === cell.c ? cell : sheet.cells.get(key(r, c));
            if (!target) continue;
            if (this.setValue(target, x === undefined ? new XErr('#N/A') : this.normalize(x))) { sheet.changed.add(key(r, c)); updated.push([sheet, target]); }
          }
          continue;
        }
        // Several values from a normal formula spill into the cells below/right (dynamic arrays).
        if (Array.isArray(v) && (v.length > 1 || (v[0] && v[0].length > 1))) {
          const res = this.spill(sheet, cell, v);
          for (const u of res.updated) updated.push(u);
          continue;
        }
        if (Array.isArray(v)) v = v[0] ? v[0][0] : 0;
        if (cell.spill) for (const u of this.clearSpill(sheet, cell, null)) updated.push(u);
        if (this.setValue(cell, this.normalize(v))) { sheet.changed.add(key(cell.r, cell.c)); updated.push([sheet, cell]); }
      }
      // Cells that read spilled values need a pass of their own.
      const spilled = updated.filter(([s, c]) => c.spillFrom && !dirty.has(this.nodeId(s, c)));
      if (spilled.length && !this.inSpillPass) {
        this.inSpillPass = true;
        try { for (const u of this.recalc(spilled.map(([s, c]) => ({ sheet: s, r: c.r, c: c.c })), { volatile: false })) updated.push(u); } finally { this.inSpillPass = false; }
      }
      this.cycles = cycles;
      return updated;
    }

    spill(sheet, cell, arr) {
      const updated = [];
      const anchor = key(cell.r, cell.c);
      const rows = arr.length, cols = Math.max(...arr.map((r) => r.length));
      const g = { r1: cell.r, c1: cell.c, r2: cell.r + rows - 1, c2: cell.c + cols - 1 };
      let blocked = false;
      for (let r = g.r1; r <= g.r2 && !blocked; r++) for (let c = g.c1; c <= g.c2; c++) {
        if (r === cell.r && c === cell.c) continue;
        const x = sheet.get(r, c);
        if (x && x.spillFrom !== anchor && (x.formula || (x.value !== null && x.value !== undefined && x.value !== '') || x.arrayMaster)) { blocked = true; break; }
      }
      for (const u of this.clearSpill(sheet, cell, blocked ? null : g)) updated.push(u);
      if (blocked) {
        if (this.setValue(cell, new XErr('#SPILL!'))) { sheet.changed.add(anchor); updated.push([sheet, cell]); }
        return { updated };
      }
      cell.spill = g;
      sheet.changed.add(anchor);
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        const x = arr[r - g.r1] ? arr[r - g.r1][c - g.c1] : undefined;
        const val = this.normalize(x === undefined ? null : x);
        let target = sheet.get(r, c);
        if (!target) { target = sheet.setCell(r, c, { value: null, spillFrom: anchor }); }
        if (target !== cell) target.spillFrom = anchor;
        if (this.setValue(target, val) || target === cell) { sheet.changed.add(key(r, c)); updated.push([sheet, target]); }
      }
      return { updated };
    }

    // Remove spilled values that are outside the new spill range (or all of them).
    clearSpill(sheet, cell, keep) {
      const out = [];
      const old = cell.spill;
      if (!old) return out;
      const anchor = key(cell.r, cell.c);
      for (let r = old.r1; r <= old.r2; r++) for (let c = old.c1; c <= old.c2; c++) {
        if (r === cell.r && c === cell.c) continue;
        if (keep && r >= keep.r1 && r <= keep.r2 && c >= keep.c1 && c <= keep.c2) continue;
        const x = sheet.get(r, c);
        if (!x || x.spillFrom !== anchor) continue;
        delete x.spillFrom;
        x.value = null;
        sheet.changed.add(key(r, c));
        out.push([sheet, x]);
      }
      if (!keep) delete cell.spill;
      return out;
    }

    normalize(v) {
      if (this.engine.isError(v)) return new XErr(v.error);
      if (v === null || v === undefined) return 0;
      if (typeof v === 'object' && !(v instanceof XErr)) return new XErr('#VALUE!');
      if (typeof v === 'number' && !Number.isFinite(v)) return new XErr('#NUM!');
      return v;
    }

    setValue(cell, v) {
      const old = cell.value;
      const same = old === v || (old instanceof XErr && v instanceof XErr && old.code === v.code);
      cell.value = v;
      cell.cached = true;
      return !same;
    }
  }

  return { Calc };
})();
