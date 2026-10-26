'use strict';
// Excel workbook model: reads an .xlsx package into sheets, cells and styles, and saves by patching only the
// parts that changed (edited cells, new styles, column widths, merges, new sheets). Everything else in the
// package - charts, pictures, macros, pivot tables, conditional formatting, data validation - is untouched.

const SheetModel = (() => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const MAX_ROW = 1048576;
  const MAX_COL = 16384;
  const KEYW = 16385;

  // ---------- addresses

  const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colNum = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
  const addr = (r, c) => colName(c) + r;
  const parseAddr = (a) => { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a.trim()); return m ? { r: +m[2], c: colNum(m[1]) } : null; };
  const parseRange = (ref) => {
    const [a, b] = ref.split(':');
    const p = parseAddr(a); const q = b ? parseAddr(b) : p;
    if (!p || !q) return null;
    return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) };
  };
  const rangeRef = (g) => (g.r1 === g.r2 && g.c1 === g.c2 ? addr(g.r1, g.c1) : `${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`);
  const key = (r, c) => r * KEYW + c;
  const unkey = (k) => ({ r: Math.floor(k / KEYW), c: k % KEYW });

  class XErr { constructor(code) { this.code = code; } toString() { return this.code; } }

  // ---------- XML helpers

  const parseXml = (t) => new DOMParser().parseFromString(t, 'application/xml');
  const serialize = (doc) => {
    const s = new XMLSerializer().serializeToString(doc);
    return s.startsWith('<?xml') ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
  };
  const kids = (n, name) => [...n.childNodes].filter((c) => c.nodeType === 1 && (!name || c.localName === name));
  const child = (n, name) => kids(n, name)[0] || null;
  const text = (n) => (n ? n.textContent : '');
  function makeEl(doc, name, attrs) {
    const root = doc.documentElement;
    const el = doc.createElementNS(root.namespaceURI, (root.prefix ? root.prefix + ':' : '') + name);
    for (const [k, v] of Object.entries(attrs || {})) if (v !== undefined && v !== null) el.setAttribute(k, String(v));
    return el;
  }
  function insertOrdered(parent, el, order) {
    const old = child(parent, el.localName);
    if (old) { parent.replaceChild(el, old); return el; }
    const rank = order.indexOf(el.localName);
    const after = kids(parent).find((c) => order.indexOf(c.localName) > rank);
    parent.insertBefore(el, after || null);
    return el;
  }
  const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const WB_ORDER = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
  const SS_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
  const FONT_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];

  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const p of target.split('/')) { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); }
    return parts.join('/');
  }
  async function readRels(zip, partPath) {
    const relsPath = partPath.replace(/([^/]+)$/, '_rels/$1.rels');
    const f = zip.file(relsPath);
    const doc = f ? parseXml(await f.async('string')) : parseXml(`<Relationships xmlns="${PKG}"/>`);
    const list = [...doc.getElementsByTagNameNS(PKG, 'Relationship')].map((r) => ({
      id: r.getAttribute('Id'), type: r.getAttribute('Type'), target: r.getAttribute('Target'), mode: r.getAttribute('TargetMode'),
      path: r.getAttribute('TargetMode') === 'External' ? null : resolvePath(partPath, r.getAttribute('Target')), el: r,
    }));
    return { relsPath, doc, list };
  }

  // ---------- colours

  const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];

  function applyTint(hex, tint) {
    if (!tint) return hex;
    let r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0, l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
      h /= 6;
    }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    if (s === 0) r = g = b = l;
    else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3); }
    return [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  // ---------- styles

  class Styles {
    constructor(doc, theme) {
      this.doc = doc;
      this.theme = theme;
      const root = doc.documentElement;
      this.numFmts = new Map();
      const nf = child(root, 'numFmts');
      if (nf) for (const n of kids(nf, 'numFmt')) this.numFmts.set(+n.getAttribute('numFmtId'), n.getAttribute('formatCode'));
      this.fonts = kids(child(root, 'fonts') || root, 'font');
      this.fills = kids(child(root, 'fills') || root, 'fill');
      this.borders = kids(child(root, 'borders') || root, 'border');
      this.xfs = kids(child(root, 'cellXfs') || root, 'xf');
      this.cssCache = new Map();
      this.deriveCache = new Map();
      this.dirty = false;
    }

    color(el) {
      if (!el) return null;
      if (el.getAttribute('auto') === '1') return null;
      let hex = null;
      if (el.hasAttribute('rgb')) hex = el.getAttribute('rgb').slice(-6);
      else if (el.hasAttribute('theme')) hex = this.theme[+el.getAttribute('theme')] || null;
      else if (el.hasAttribute('indexed')) { const i = +el.getAttribute('indexed'); hex = i === 64 || i === 65 ? null : INDEXED[i] || null; }
      if (!hex) return null;
      return '#' + applyTint(hex.toUpperCase(), parseFloat(el.getAttribute('tint') || '0'));
    }

    xf(i) {
      const x = this.xfs[i] || this.xfs[0];
      if (!x) return { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, align: {} };
      const al = child(x, 'alignment');
      return {
        numFmtId: +(x.getAttribute('numFmtId') || 0), fontId: +(x.getAttribute('fontId') || 0), fillId: +(x.getAttribute('fillId') || 0), borderId: +(x.getAttribute('borderId') || 0),
        align: al ? { horizontal: al.getAttribute('horizontal'), vertical: al.getAttribute('vertical'), wrap: al.getAttribute('wrapText') === '1' || al.getAttribute('wrapText') === 'true', indent: +(al.getAttribute('indent') || 0), rotation: +(al.getAttribute('textRotation') || 0) } : {},
      };
    }

    numFmt(i) {
      const id = this.xf(i).numFmtId;
      if (this.numFmts.has(id)) return this.numFmts.get(id);
      const table = XLSX.SSF.get_table ? XLSX.SSF.get_table() : {};
      return table[id] || 'General';
    }

    font(i) {
      const f = this.fonts[this.xf(i).fontId] || this.fonts[0];
      const get = (n) => f && child(f, n);
      const flag = (n) => { const e = get(n); return !!e && e.getAttribute('val') !== '0' && e.getAttribute('val') !== 'false'; };
      const u = get('u');
      return {
        name: get('name') ? get('name').getAttribute('val') : 'Calibri',
        size: get('sz') ? +get('sz').getAttribute('val') : 11,
        bold: flag('b'), italic: flag('i'), strike: flag('strike'),
        underline: u ? (u.getAttribute('val') || 'single') !== 'none' : false,
        color: this.color(get('color')),
        vert: get('vertAlign') ? get('vertAlign').getAttribute('val') : null,
      };
    }

    fill(i) {
      const f = this.fills[this.xf(i).fillId];
      const p = f && child(f, 'patternFill');
      if (p && p.getAttribute('patternType') && p.getAttribute('patternType') !== 'none') {
        if (p.getAttribute('patternType') === 'solid') return this.color(child(p, 'fgColor')) || null;
        return this.color(child(p, 'fgColor')) || this.color(child(p, 'bgColor'));
      }
      const g = f && child(f, 'gradientFill');
      if (g) { const stop = child(g, 'stop'); return stop ? this.color(child(stop, 'color')) : null; }
      return null;
    }

    border(i) {
      const b = this.borders[this.xf(i).borderId];
      const out = {};
      if (!b) return out;
      for (const side of ['left', 'right', 'top', 'bottom']) {
        const e = child(b, side) || child(b, side === 'left' ? 'start' : side === 'right' ? 'end' : '');
        const st = e && e.getAttribute('style');
        if (st && st !== 'none') out[side] = { style: st, color: this.color(child(e, 'color')) || '#000000' };
      }
      return out;
    }

    // CSS for rendering a cell with style index i.
    css(i) {
      if (this.cssCache.has(i)) return this.cssCache.get(i);
      const f = this.font(i);
      const x = this.xf(i);
      const css = {
        font: `${f.italic ? 'italic ' : ''}${f.bold ? '700 ' : ''}${(f.size * 4 / 3).toFixed(2)}px "${f.name}", Calibri, Carlito, sans-serif`,
        color: f.color, decoration: [f.underline ? 'underline' : '', f.strike ? 'line-through' : ''].filter(Boolean).join(' '),
        vert: f.vert, fill: this.fill(i), border: this.border(i),
        h: x.align.horizontal || 'general', v: x.align.vertical || 'bottom', wrap: !!x.align.wrap, indent: x.align.indent || 0, rotation: x.align.rotation || 0,
        fmt: this.numFmt(i), font_: f,
      };
      this.cssCache.set(i, css);
      return css;
    }

    // ---- writing new styles: clone the xf, swap in new font/fill/border/numFmt, reuse identical entries.
    derive(i, change) {
      const ck = `${i}|${JSON.stringify(change)}`;
      if (this.deriveCache.has(ck)) return this.deriveCache.get(ck);
      const doc = this.doc;
      const root = doc.documentElement;
      const base = (this.xfs[i] || this.xfs[0]).cloneNode(true);
      if (change.font) {
        const f = (this.fonts[+(base.getAttribute('fontId') || 0)] || this.fonts[0]).cloneNode(true);
        const set = (name, attrs) => { const old = child(f, name); if (old) f.removeChild(old); if (attrs) insertOrdered(f, makeEl(doc, name, attrs), FONT_ORDER); };
        const ch = change.font;
        if ('bold' in ch) set('b', ch.bold ? {} : null);
        if ('italic' in ch) set('i', ch.italic ? {} : null);
        if ('underline' in ch) set('u', ch.underline ? {} : null);
        if ('strike' in ch) set('strike', ch.strike ? {} : null);
        if ('size' in ch) set('sz', { val: ch.size });
        if ('color' in ch) set('color', ch.color ? { rgb: 'FF' + ch.color.replace('#', '').toUpperCase() } : null);
        if ('name' in ch) { set('name', { val: ch.name }); set('scheme', null); }
        base.setAttribute('fontId', this.reuse('fonts', this.fonts, f));
        base.setAttribute('applyFont', '1');
      }
      if ('fill' in change) {
        const f = makeEl(doc, 'fill');
        const p = makeEl(doc, 'patternFill', { patternType: change.fill ? 'solid' : 'none' });
        if (change.fill) { p.append(makeEl(doc, 'fgColor', { rgb: 'FF' + change.fill.replace('#', '').toUpperCase() })); p.append(makeEl(doc, 'bgColor', { indexed: 64 })); }
        f.append(p);
        base.setAttribute('fillId', change.fill ? this.reuse('fills', this.fills, f) : 0);
        base.setAttribute('applyFill', '1');
      }
      if (change.border) {
        // Sides not mentioned keep their current style.
        const cur = this.border(i);
        const b = makeEl(doc, 'border');
        for (const side of ['left', 'right', 'top', 'bottom']) {
          const s = side in change.border ? change.border[side] : (cur[side] && cur[side].style);
          const e = makeEl(doc, side, s ? { style: s } : {});
          if (s) e.append(makeEl(doc, 'color', { indexed: 64 }));
          b.append(e);
        }
        b.append(makeEl(doc, 'diagonal'));
        base.setAttribute('borderId', this.reuse('borders', this.borders, b));
        base.setAttribute('applyBorder', '1');
      }
      if ('numFmt' in change) {
        base.setAttribute('numFmtId', this.numFmtId(change.numFmt));
        base.setAttribute('applyNumberFormat', '1');
      }
      if (change.align) {
        let al = child(base, 'alignment');
        if (!al) { al = makeEl(doc, 'alignment'); base.insertBefore(al, base.firstChild); }
        for (const [k, v] of Object.entries(change.align)) {
          const attr = k === 'wrap' ? 'wrapText' : k;
          if (v === null || v === false || v === '') al.removeAttribute(attr); else al.setAttribute(attr, v === true ? '1' : v);
        }
        if (!al.attributes.length) base.removeChild(al);
        base.setAttribute('applyAlignment', '1');
      }
      const idx = this.reuse('cellXfs', this.xfs, base);
      this.deriveCache.set(ck, idx);
      return idx;
    }

    reuse(containerName, list, el) {
      const s = new XMLSerializer().serializeToString(el);
      const found = list.findIndex((x) => new XMLSerializer().serializeToString(x) === s);
      if (found >= 0) return found;
      const root = this.doc.documentElement;
      let container = child(root, containerName);
      if (!container) container = insertOrdered(root, makeEl(this.doc, containerName), SS_ORDER);
      container.append(el);
      list.push(el);
      container.setAttribute('count', String(list.length));
      this.dirty = true;
      this.cssCache.clear();
      return list.length - 1;
    }

    numFmtId(code) {
      const table = XLSX.SSF.get_table ? XLSX.SSF.get_table() : {};
      for (const [id, c] of Object.entries(table)) if (c === code && +id < 164) return +id;
      for (const [id, c] of this.numFmts) if (c === code) return id;
      const id = Math.max(163, ...this.numFmts.keys()) + 1;
      const root = this.doc.documentElement;
      let nf = child(root, 'numFmts');
      if (!nf) nf = insertOrdered(root, makeEl(this.doc, 'numFmts'), SS_ORDER);
      nf.append(makeEl(this.doc, 'numFmt', { numFmtId: id, formatCode: code }));
      nf.setAttribute('count', String(kids(nf, 'numFmt').length));
      this.numFmts.set(id, code);
      this.dirty = true;
      return id;
    }
  }

  // ---------- formula text: find references (shift for copy/fill/shared formulas, rename sheets)

  // Tokens: { kind: 'ref'|'col'|'row'|'other', start, end, sheet?, ... }
  function scanRefs(f) {
    const out = [];
    let i = 0;
    const n = f.length;
    let pendingSheet = null; // {start, name}
    while (i < n) {
      const ch = f[i];
      if (ch === '"') { i++; while (i < n) { if (f[i] === '"') { if (f[i + 1] === '"') { i += 2; continue; } i++; break; } i++; } pendingSheet = null; continue; }
      if (ch === "'") {
        const s = i; i++;
        let name = '';
        while (i < n) { if (f[i] === "'") { if (f[i + 1] === "'") { name += "'"; i += 2; continue; } i++; break; } name += f[i++]; }
        if (f[i] === '!') { pendingSheet = { start: s, name, quoted: true }; i++; continue; }
        continue;
      }
      if (ch === '[') { let d = 0; while (i < n) { if (f[i] === '[') d++; if (f[i] === ']') { d--; if (!d) { i++; break; } } i++; } continue; }
      if (/[A-Za-z0-9_$.\\]/.test(ch)) {
        const s = i;
        while (i < n && /[A-Za-z0-9_$.\\]/.test(f[i])) i++;
        const word = f.slice(s, i);
        if (f[i] === '!') { pendingSheet = { start: s, name: word, quoted: false }; i++; continue; }
        if (f[i] === '(') { pendingSheet = null; continue; }
        const sheet = pendingSheet; pendingSheet = null;
        const tokStart = sheet ? sheet.start : s;
        let m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(word);
        if (m && colNum(m[2]) <= MAX_COL && +m[4] >= 1 && +m[4] <= MAX_ROW) {
          out.push({ kind: 'ref', start: tokStart, refStart: s, end: i, sheet: sheet && sheet.name, quoted: sheet && sheet.quoted, colAbs: !!m[1], col: colNum(m[2]), rowAbs: !!m[3], row: +m[4] });
          continue;
        }
        // Whole columns (A:C) and rows (1:5).
        if (f[i] === ':') {
          const m2 = /^(\$?)([A-Za-z]{1,3})$/.exec(word);
          const r2 = /^(\$?)(\d+)$/.exec(word);
          const j = i + 1;
          let k = j; while (k < n && /[A-Za-z0-9$]/.test(f[k])) k++;
          const word2 = f.slice(j, k);
          if (m2) { const mm = /^(\$?)([A-Za-z]{1,3})$/.exec(word2); if (mm) { out.push({ kind: 'col', start: tokStart, refStart: s, end: i, sheet: sheet && sheet.name, quoted: sheet && sheet.quoted, colAbs: !!m2[1], col: colNum(m2[2]) }); out.push({ kind: 'col', start: j, refStart: j, end: k, colAbs: !!mm[1], col: colNum(mm[2]), second: true }); i = k; continue; } }
          if (r2) { const rr = /^(\$?)(\d+)$/.exec(word2); if (rr) { out.push({ kind: 'row', start: tokStart, refStart: s, end: i, sheet: sheet && sheet.name, quoted: sheet && sheet.quoted, rowAbs: !!r2[1], row: +r2[2] }); out.push({ kind: 'row', start: j, refStart: j, end: k, rowAbs: !!rr[1], row: +rr[2], second: true }); i = k; continue; } }
        }
        continue;
      }
      pendingSheet = null;
      i++;
    }
    return out;
  }

  const quoteSheet = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseAddr(name) ? name : `'${name.replace(/'/g, "''")}'`);

  // Move relative references by (dr, dc); references pushed off the sheet become #REF!.
  function shiftFormula(f, dr, dc) {
    if (!dr && !dc) return f;
    const toks = scanRefs(f);
    let out = '';
    let last = 0;
    for (const t of toks) {
      out += f.slice(last, t.refStart);
      if (t.kind === 'ref') {
        const r = t.rowAbs ? t.row : t.row + dr;
        const c = t.colAbs ? t.col : t.col + dc;
        out += r < 1 || c < 1 || r > MAX_ROW || c > MAX_COL ? '#REF!' : `${t.colAbs ? '$' : ''}${colName(c)}${t.rowAbs ? '$' : ''}${r}`;
      } else if (t.kind === 'col') {
        const c = t.colAbs ? t.col : t.col + dc;
        out += c < 1 ? '#REF!' : `${t.colAbs ? '$' : ''}${colName(c)}`;
      } else {
        const r = t.rowAbs ? t.row : t.row + dr;
        out += r < 1 ? '#REF!' : `${t.rowAbs ? '$' : ''}${r}`;
      }
      last = t.end;
    }
    return out + f.slice(last);
  }

  function renameSheetInFormula(f, oldName, newName) {
    const toks = scanRefs(f).filter((t) => t.sheet && t.sheet.toLowerCase() === oldName.toLowerCase());
    if (!toks.length) return f;
    let out = '';
    let last = 0;
    for (const t of toks) { out += f.slice(last, t.start) + quoteSheet(newName) + '!'; last = t.refStart; }
    return out + f.slice(last);
  }

  // ---------- sheets

  class Sheet {
    constructor(wb, info) {
      Object.assign(this, info); // name, sheetId, rid, path, state
      this.wb = wb;
      this.cells = new Map();
      this.rows = new Map();
      this.cols = [];
      this.merges = [];
      this.frozen = { rows: 0, cols: 0 };
      this.showGrid = true;
      this.defaultRowHeight = 15;
      this.defaultColWidth = null;
      this.maxRow = 0;
      this.maxCol = 0;
      this.changed = new Set();      // cell keys whose content or style changed
      this.colsChanged = false;
      this.mergesChanged = false;
      this.rowsChanged = new Set();  // rows whose height/hidden state changed
      this.autoFilter = null;        // { ref, cols: Map(absCol -> { values: Set, blanks }) }
      this.filterChanged = false;
      this.structureChanged = false;
    }

    get(r, c) { return this.cells.get(key(r, c)); }

    async parse(xmlText) {
      this.doc = parseXml(xmlText);
      const root = this.doc.documentElement;
      const sv = root.getElementsByTagNameNS(NS, 'sheetView')[0];
      if (sv) {
        if (sv.getAttribute('showGridLines') === '0' || sv.getAttribute('showGridLines') === 'false') this.showGrid = false;
        const pane = child(sv, 'pane');
        if (pane && (pane.getAttribute('state') === 'frozen' || pane.getAttribute('state') === 'frozenSplit')) {
          this.frozen = { rows: +(pane.getAttribute('ySplit') || 0), cols: +(pane.getAttribute('xSplit') || 0) };
        }
      }
      const fmt = child(root, 'sheetFormatPr');
      if (fmt) {
        if (fmt.getAttribute('defaultRowHeight')) this.defaultRowHeight = +fmt.getAttribute('defaultRowHeight');
        if (fmt.getAttribute('defaultColWidth')) this.defaultColWidth = +fmt.getAttribute('defaultColWidth');
      }
      const cols = child(root, 'cols');
      if (cols) for (const c of kids(cols, 'col')) {
        this.cols.push({ min: +c.getAttribute('min'), max: +c.getAttribute('max'), width: c.hasAttribute('width') ? +c.getAttribute('width') : null, hidden: c.getAttribute('hidden') === '1' || c.getAttribute('hidden') === 'true', style: c.hasAttribute('style') ? +c.getAttribute('style') : null, outline: +(c.getAttribute('outlineLevel') || 0), el: c });
      }
      const af = child(root, 'autoFilter');
      if (af && af.getAttribute('ref')) {
        const g = parseRange(af.getAttribute('ref'));
        const cols = new Map();
        for (const fc of kids(af, 'filterColumn')) {
          const f = child(fc, 'filters');
          if (!f) continue;
          cols.set(g.c1 + +fc.getAttribute('colId'), { values: new Set(kids(f, 'filter').map((x) => x.getAttribute('val'))), blanks: f.getAttribute('blank') === '1' || f.getAttribute('blank') === 'true' });
        }
        this.autoFilter = { ref: af.getAttribute('ref'), cols };
      }
      const mc = child(root, 'mergeCells');
      if (mc) for (const m of kids(mc, 'mergeCell')) { const g = parseRange(m.getAttribute('ref')); if (g) this.merges.push(g); }

      const sd = child(root, 'sheetData');
      this.sheetData = sd;
      this.rowEls = new Map();
      const sharedMasters = new Map();
      const sharedChildren = [];
      let rIdx = 0;
      for (const row of kids(sd, 'row')) {
        rIdx = row.hasAttribute('r') ? +row.getAttribute('r') : rIdx + 1;
        this.rowEls.set(rIdx, row);
        const info = {};
        if (row.hasAttribute('ht')) info.ht = +row.getAttribute('ht');
        if (row.getAttribute('hidden') === '1' || row.getAttribute('hidden') === 'true') info.hidden = true;
        if (row.getAttribute('customFormat') === '1' && row.hasAttribute('s')) info.style = +row.getAttribute('s');
        if (row.hasAttribute('outlineLevel')) info.outline = +row.getAttribute('outlineLevel');
        if (Object.keys(info).length) this.rows.set(rIdx, info);
        let cIdx = 0;
        for (const c of kids(row, 'c')) {
          let r = rIdx;
          if (c.hasAttribute('r')) { const p = parseAddr(c.getAttribute('r')); r = p.r; cIdx = p.c; } else cIdx++;
          const cell = { r, c: cIdx, s: +(c.getAttribute('s') || 0), el: c, value: null, formula: null };
          const t = c.getAttribute('t') || 'n';
          const v = child(c, 'v');
          const f = child(c, 'f');
          if (t === 's') cell.value = v ? this.wb.sst[+text(v)] ?? '' : '';
          else if (t === 'inlineStr') cell.value = textOfRich(child(c, 'is'));
          else if (t === 'b') cell.value = text(v) === '1' || text(v) === 'true';
          else if (t === 'e') cell.value = new XErr(text(v));
          else if (t === 'str') cell.value = text(v);
          else if (t === 'd') cell.value = v ? this.wb.dateToSerial(new Date(text(v))) : null;
          else cell.value = v && text(v) !== '' ? +text(v) : null;
          cell.cached = !!v && text(v) !== '';
          if (c.hasAttribute('cm')) cell.dynamic = true;
          if (f) {
            const ft = f.getAttribute('t');
            const ftext = text(f);
            if (ft === 'shared') {
              const si = f.getAttribute('si');
              cell.shared = si;
              if (ftext) { cell.formula = ftext; sharedMasters.set(si, cell); } else sharedChildren.push(cell);
            } else if (ft === 'array') {
              cell.formula = ftext;
              cell.array = f.getAttribute('ref') ? parseRange(f.getAttribute('ref')) : { r1: r, c1: cIdx, r2: r, c2: cIdx };
            } else if (ft !== 'dataTable') cell.formula = ftext;
          }
          if (cell.value === null && !cell.formula && !cell.s) continue;
          this.cells.set(key(r, cIdx), cell);
          if (r > this.maxRow) this.maxRow = r;
          if (cIdx > this.maxCol) this.maxCol = cIdx;
        }
      }
      for (const cell of sharedChildren) {
        const m = sharedMasters.get(cell.shared);
        if (m) cell.formula = shiftFormula(m.formula, cell.r - m.r, cell.c - m.c);
      }
      // Cells covered by an array formula point at their master.
      for (const cell of this.cells.values()) {
        if (!cell.array) continue;
        for (let r = cell.array.r1; r <= cell.array.r2; r++) for (let c = cell.array.c1; c <= cell.array.c2; c++) {
          const k = key(r, c);
          if (r === cell.r && c === cell.c) continue;
          const other = this.cells.get(k) || { r, c, s: 0, value: null, formula: null };
          other.arrayMaster = key(cell.r, cell.c);
          this.cells.set(k, other);
        }
      }
      for (const m of this.merges) { if (m.r2 > this.maxRow) this.maxRow = m.r2; if (m.c2 > this.maxCol) this.maxCol = m.c2; }
      // Rows without a saved height grow to fit their largest font, as Excel does.
      this.autoHt = new Map();
      for (const cell of this.cells.values()) this.noteFont(cell);
    }

    noteFont(cell) {
      if (!this.autoHt || !cell.s || cell.value === null || cell.value === undefined) return;
      const size = this.wb.styles.font(cell.s).size;
      if (size <= 11) return;
      const ht = Math.round(size * 1.32 * 4) / 4;
      if (ht > (this.autoHt.get(cell.r) || 0)) this.autoHt.set(cell.r, ht);
    }

    colWidthChars(c) {
      for (const col of this.cols) if (c >= col.min && c <= col.max) { if (col.hidden) return 0; if (col.width != null) return col.width; }
      return this.defaultColWidth; // null = Excel's default (64 px)
    }
    rowHeightPt(r) {
      const info = this.rows.get(r);
      if (info) { if (info.hidden) return 0; if (info.ht != null) return info.ht; }
      const auto = this.autoHt && this.autoHt.get(r);
      return auto && auto > this.defaultRowHeight ? auto : this.defaultRowHeight;
    }
    mergeAt(r, c) { return this.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || null; }

    // ---- edits (the view calls these; undo is the view's job)
    setCell(r, c, patch) {
      const k = key(r, c);
      let cell = this.cells.get(k);
      if (!cell) { cell = { r, c, s: this.defaultStyleAt(r, c), value: null, formula: null }; this.cells.set(k, cell); }
      if (cell.shared != null && ('formula' in patch || 'value' in patch)) this.unshare(cell.shared);
      // Typing into a spilled cell takes it over (the spilling formula then shows #SPILL!).
      if (cell.spillFrom && ('formula' in patch || 'value' in patch) && !patch.spillFrom) delete cell.spillFrom;
      Object.assign(cell, patch);
      if ('formula' in patch && !patch.formula) { delete cell.array; }
      if ('s' in patch || 'value' in patch) this.noteFont(cell);
      this.changed.add(k);
      if (r > this.maxRow) this.maxRow = r;
      if (c > this.maxCol) this.maxCol = c;
      return cell;
    }
    defaultStyleAt(r, c) {
      const row = this.rows.get(r);
      if (row && row.style != null) return row.style;
      for (const col of this.cols) if (c >= col.min && c <= col.max && col.style != null) return col.style;
      return 0;
    }
    unshare(si) {
      for (const [k, cell] of this.cells) if (cell.shared === si) { delete cell.shared; this.changed.add(k); cell.formulaChanged = true; }
      this.wb.formulasChanged = true;
    }
    setColWidth(c, widthChars) {
      // Split any <col> range that covers c so only this column changes.
      const out = [];
      let found = false;
      for (const col of this.cols) {
        if (c < col.min || c > col.max) { out.push(col); continue; }
        if (col.min < c) out.push({ ...col, max: c - 1, el: null });
        out.push({ ...col, min: c, max: c, width: widthChars, custom: true, el: null });
        if (col.max > c) out.push({ ...col, min: c + 1, el: null });
        found = true;
      }
      if (!found) out.push({ min: c, max: c, width: widthChars, custom: true, hidden: false, style: null });
      this.cols = out.sort((a, b) => a.min - b.min);
      this.colsChanged = true;
    }

    setRowHeight(r, pt) { const info = { ...(this.rows.get(r) || {}) }; if (pt == null) delete info.ht; else info.ht = pt; info.custom = pt != null; this.rows.set(r, info); this.rowsChanged.add(r); }
    setRowHidden(r, hidden) { const info = { ...(this.rows.get(r) || {}) }; info.hidden = !!hidden; this.rows.set(r, info); this.rowsChanged.add(r); }
    setColHidden(c, hidden) {
      const cur = this.colWidthChars(c);
      if (cur === 0 || cur == null) this.setColWidth(c, this.cols.find((x) => c >= x.min && c <= x.max && x.width) ? this.cols.find((x) => c >= x.min && c <= x.max).width : (this.defaultColWidth || 9.140625));
      else this.setColWidth(c, cur);
      for (const col of this.cols) if (c >= col.min && c <= col.max) { col.hidden = !!hidden; col.el = null; }
      this.colsChanged = true;
    }
    isColHidden(c) { return this.cols.some((col) => c >= col.min && c <= col.max && col.hidden); }

    // ---- write back
    save(sst) {
      const doc = this.doc;
      const root = doc.documentElement;
      if (!this.changed.size && !this.colsChanged && !this.mergesChanged && !this.rowsChanged.size && !this.filterChanged && !this.structureChanged) return false;
      const sd = this.sheetData;
      const rowEl = (r) => {
        let el = this.rowEls.get(r);
        if (el) return el;
        el = makeEl(doc, 'row', { r });
        const next = [...this.rowEls.keys()].filter((x) => x > r).sort((a, b) => a - b)[0];
        sd.insertBefore(el, next ? this.rowEls.get(next) : null);
        this.rowEls.set(r, el);
        return el;
      };
      for (const k of this.changed) {
        const cell = this.cells.get(k);
        if (!cell) continue;
        const row = rowEl(cell.r);
        row.removeAttribute('spans');
        const empty = cell.value === null && !cell.formula && !cell.s;
        if (empty && !cell.arrayMaster && !cell.spillFrom) {
          if (cell.el && cell.el.parentNode) cell.el.parentNode.removeChild(cell.el);
          cell.el = null;
          this.cells.delete(k);
          continue;
        }
        let c = cell.el;
        if (!c) {
          c = makeEl(doc, 'c', { r: addr(cell.r, cell.c) });
          const after = kids(row, 'c').find((x) => { const p = parseAddr(x.getAttribute('r') || ''); return p && p.c > cell.c; });
          row.insertBefore(c, after || null);
          cell.el = c;
        }
        if (cell.s) c.setAttribute('s', cell.s); else c.removeAttribute('s');
        // Formula: keep shared/array structure unless it was changed.
        let f = kids(c, 'f')[0];
        if (cell.formula && !cell.arrayMaster) {
          const keepShared = f && f.getAttribute('t') === 'shared' && cell.shared != null && !cell.formulaChanged;
          if (!keepShared) {
            if (!f) f = makeEl(doc, 'f');
            for (const a of ['t', 'si', 'ref', 'aca', 'ca', 'dt2D', 'dtr', 'del1', 'del2', 'r1', 'r2', 'bx']) f.removeAttribute(a);
            // Spilled results are saved as an array formula over the spill range.
            const arr = cell.array || cell.spill;
            if (arr) { f.setAttribute('t', 'array'); f.setAttribute('ref', rangeRef(arr)); }
            f.textContent = cell.formula;
          }
          if (!f.parentNode) c.insertBefore(f, c.firstChild);
          if (cell.formulaChanged) c.removeAttribute('cm');
        } else if (f) c.removeChild(f);
        for (const v of kids(c, 'v').concat(kids(c, 'is'))) c.removeChild(v);
        const val = cell.value;
        c.removeAttribute('t');
        if (val === null || val === undefined) { /* no value */ }
        else if (val instanceof XErr) { c.setAttribute('t', 'e'); c.append(textEl(doc, 'v', val.code)); }
        else if (typeof val === 'boolean') { c.setAttribute('t', 'b'); c.append(textEl(doc, 'v', val ? '1' : '0')); }
        else if (typeof val === 'number') c.append(textEl(doc, 'v', String(val)));
        else if (cell.formula || cell.arrayMaster || cell.spillFrom) { c.setAttribute('t', 'str'); c.append(textEl(doc, 'v', String(val))); }
        else { c.setAttribute('t', 's'); c.append(textEl(doc, 'v', String(sst.index(String(val))))); }
      }
      for (const r of this.rowsChanged) {
        const info = this.rows.get(r) || {};
        const el = rowEl(r);
        if (info.hidden) el.setAttribute('hidden', '1'); else el.removeAttribute('hidden');
        if ('outline' in info) { if (info.outline) el.setAttribute('outlineLevel', info.outline); else el.removeAttribute('outlineLevel'); }
        if ('collapsed' in info) { if (info.collapsed) el.setAttribute('collapsed', '1'); else el.removeAttribute('collapsed'); }
        if (info.ht != null) { el.setAttribute('ht', info.ht); el.setAttribute('customHeight', '1'); }
        else if (info.custom === false) { el.removeAttribute('ht'); el.removeAttribute('customHeight'); }
      }
      if (this.filterChanged) {
        const old = child(root, 'autoFilter');
        if (this.autoFilter && this.autoFilter.ref) {
          const g = parseRange(this.autoFilter.ref);
          const af = makeEl(doc, 'autoFilter', { ref: this.autoFilter.ref });
          for (const [c, f] of [...this.autoFilter.cols].sort((a, b) => a[0] - b[0])) {
            if (c < g.c1 || c > g.c2) continue;
            const fc = makeEl(doc, 'filterColumn', { colId: c - g.c1 });
            const fl = makeEl(doc, 'filters', f.blanks ? { blank: '1' } : {});
            for (const v of f.values) fl.append(makeEl(doc, 'filter', { val: v }));
            fc.append(fl);
            af.append(fc);
          }
          insertOrdered(root, af, WS_ORDER);
        } else if (old) root.removeChild(old);
      }
      // Dimension
      let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
      for (const cell of this.cells.values()) { if (cell.r < r1) r1 = cell.r; if (cell.c < c1) c1 = cell.c; if (cell.r > r2) r2 = cell.r; if (cell.c > c2) c2 = cell.c; }
      const dim = child(root, 'dimension') || insertOrdered(root, makeEl(doc, 'dimension'), WS_ORDER);
      dim.setAttribute('ref', r2 ? rangeRef({ r1, c1, r2, c2 }) : 'A1');
      if (this.colsChanged) {
        if (this.cols.length) {
          const cols = makeEl(doc, 'cols');
          for (const col of this.cols) {
            const el = col.el ? col.el.cloneNode(true) : makeEl(doc, 'col');
            el.setAttribute('min', col.min); el.setAttribute('max', col.max);
            if (col.width != null) el.setAttribute('width', col.width);
            if (col.custom) el.setAttribute('customWidth', '1');
            if (col.style != null) el.setAttribute('style', col.style);
            if (col.hidden) el.setAttribute('hidden', '1'); else el.removeAttribute('hidden');
            if ('outline' in col) { if (col.outline) el.setAttribute('outlineLevel', col.outline); else el.removeAttribute('outlineLevel'); }
            cols.append(el);
          }
          insertOrdered(root, cols, WS_ORDER);
        } else { const old = child(root, 'cols'); if (old) root.removeChild(old); }
      }
      if (this.mergesChanged) {
        const old = child(root, 'mergeCells');
        if (this.merges.length) {
          const mc = makeEl(doc, 'mergeCells', { count: this.merges.length });
          for (const m of this.merges) mc.append(makeEl(doc, 'mergeCell', { ref: rangeRef(m) }));
          insertOrdered(root, mc, WS_ORDER);
        } else if (old) root.removeChild(old);
      }
      return true;
    }
  }

  function textEl(doc, name, value) { const e = makeEl(doc, name); e.textContent = value; return e; }
  function textOfRich(is) {
    if (!is) return '';
    const ts = [...is.getElementsByTagNameNS(NS, 't')].filter((t) => !(t.parentNode && t.parentNode.localName === 'rPh'));
    return ts.map((t) => t.textContent).join('');
  }

  // ---------- shared strings

  class SharedStrings {
    constructor(doc) {
      this.doc = doc;
      this.list = doc ? kids(doc.documentElement, 'si').map(textOfRich) : [];
      this.lookup = new Map();
      this.list.forEach((s, i) => { if (!this.lookup.has(s)) this.lookup.set(s, i); });
      this.dirty = false;
    }
    index(s) {
      if (this.lookup.has(s)) return this.lookup.get(s);
      if (!this.doc) this.doc = parseXml(`<sst xmlns="${NS}" count="0" uniqueCount="0"/>`);
      const si = makeEl(this.doc, 'si');
      const t = textEl(this.doc, 't', s);
      if (/^\s|\s$|\n/.test(s)) t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      si.append(t);
      this.doc.documentElement.append(si);
      this.list.push(s);
      this.lookup.set(s, this.list.length - 1);
      this.dirty = true;
      return this.list.length - 1;
    }
  }

  // ======================================================================

  class Workbook {
    async load(bytes) {
      this.bytes = bytes;
      const zip = this.zip = await JSZip.loadAsync(bytes);
      const pkgRels = await readRels(zip, '');
      const main = pkgRels.list.find((r) => /\/officeDocument$/.test(r.type));
      this.wbPath = main ? main.path : 'xl/workbook.xml';
      this.wbDoc = parseXml(await zip.file(this.wbPath).async('string'));
      this.wbRels = await readRels(zip, this.wbPath);
      const wbRoot = this.wbDoc.documentElement;
      const pr = child(wbRoot, 'workbookPr');
      this.date1904 = !!pr && (pr.getAttribute('date1904') === '1' || pr.getAttribute('date1904') === 'true');
      const relOf = (type) => this.wbRels.list.find((r) => r.type.endsWith('/' + type));
      // Theme colours in Excel's index order: lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink.
      this.theme = [];
      const themeRel = relOf('theme');
      if (themeRel && zip.file(themeRel.path)) {
        const t = parseXml(await zip.file(themeRel.path).async('string'));
        const scheme = t.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'clrScheme')[0];
        if (scheme) {
          const val = (name) => { const e = kids(scheme, name)[0]; if (!e) return null; const c = kids(e)[0]; return c ? (c.getAttribute('lastClr') || c.getAttribute('val')) : null; };
          this.theme = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(val);
        }
      }
      const sstRel = relOf('sharedStrings');
      this.sstPath = sstRel ? sstRel.path : null;
      this.sstObj = new SharedStrings(sstRel && zip.file(sstRel.path) ? parseXml(await zip.file(sstRel.path).async('string')) : null);
      this.sst = this.sstObj.list;
      const stylesRel = relOf('styles');
      this.stylesPath = stylesRel ? stylesRel.path : null;
      this.styles = new Styles(stylesRel && zip.file(stylesRel.path) ? parseXml(await zip.file(stylesRel.path).async('string')) : parseXml(BLANK_STYLES), this.theme);
      this.definedNames = [...wbRoot.getElementsByTagNameNS(NS, 'definedName')].map((d) => ({ name: d.getAttribute('name'), local: d.hasAttribute('localSheetId') ? +d.getAttribute('localSheetId') : null, ref: d.textContent, el: d }));
      // External workbook links: [n] in formulas is the n-th externalReference; Excel caches their values.
      this.externals = [];
      for (const ref of wbRoot.getElementsByTagNameNS(NS, 'externalReference')) {
        const rel = this.wbRels.list.find((r) => r.id === (ref.getAttributeNS(RNS, 'id') || ref.getAttribute('r:id')));
        const ext = { sheets: [], data: [] };
        if (rel && zip.file(rel.path)) {
          const x = parseXml(await zip.file(rel.path).async('string'));
          ext.sheets = [...x.getElementsByTagNameNS(NS, 'sheetName')].map((n) => n.getAttribute('val'));
          for (const sd of x.getElementsByTagNameNS(NS, 'sheetData')) {
            const m = new Map();
            for (const cel of sd.getElementsByTagNameNS(NS, 'cell')) {
              const p = parseAddr(cel.getAttribute('r'));
              const v = child(cel, 'v');
              const t = cel.getAttribute('t');
              if (!p || !v) continue;
              m.set(key(p.r, p.c), t === 's' || t === 'str' ? text(v) : t === 'b' ? text(v) === '1' : t === 'e' ? new XErr(text(v)) : +text(v));
            }
            ext.data[+sd.getAttribute('sheetId')] = m;
          }
        }
        this.externals.push(ext);
      }
      this.sheets = [];
      for (const s of wbRoot.getElementsByTagNameNS(NS, 'sheet')) {
        const rid = s.getAttributeNS(RNS, 'id') || s.getAttribute('r:id');
        const rel = this.wbRels.list.find((r) => r.id === rid);
        const info = { name: s.getAttribute('name'), sheetId: +s.getAttribute('sheetId'), rid, path: rel && rel.path, state: s.getAttribute('state') || 'visible', el: s, kind: rel && rel.type.endsWith('/worksheet') ? 'worksheet' : 'other' };
        const sheet = new Sheet(this, info);
        if (info.kind === 'worksheet' && zip.file(info.path)) await sheet.parse(await zip.file(info.path).async('string'));
        else sheet.unsupported = true;
        this.sheets.push(sheet);
      }
      this.formulasChanged = false;
      this.wbChanged = false;
      this.newParts = {};
      this.partDocs = new Map();
      this.dirtyParts = new Set();
      this.relsCache = new Map();
      await this.loadTables();
      return this;
    }

    async loadTables() {
      this.tables = [];
      for (const s of this.sheets) {
        if (s.unsupported || s.isNew) continue;
        const rels = await this.rels(s.path);
        for (const rel of rels.list.filter((r) => r.type.endsWith('/table'))) {
          const doc = await this.part(rel.path);
          if (!doc) continue;
          const t = doc.documentElement;
          const g = parseRange(t.getAttribute('ref') || '');
          if (!g) continue;
          const cols = [...t.getElementsByTagNameNS(NS, 'tableColumn')].map((c) => c.getAttribute('name'));
          const afEl = [...t.childNodes].find((x) => x.localName === 'autoFilter');
          let autoFilter = null;
          if (afEl) {
            const fg = parseRange(afEl.getAttribute('ref') || t.getAttribute('ref'));
            const fcols = new Map();
            for (const fc of kids(afEl, 'filterColumn')) {
              const f = child(fc, 'filters');
              if (f) fcols.set(fg.c1 + +fc.getAttribute('colId'), { values: new Set(kids(f, 'filter').map((x) => x.getAttribute('val'))), blanks: f.getAttribute('blank') === '1' });
            }
            autoFilter = { ref: afEl.getAttribute('ref') || t.getAttribute('ref'), cols: fcols, el: afEl };
          }
          this.tables.push({
            autoFilter, doc,
            name: t.getAttribute('displayName') || t.getAttribute('name'), sheet: s, g, cols,
            header: t.getAttribute('headerRowCount') === '0' ? 0 : 1, totals: +(t.getAttribute('totalsRowCount') || 0),
            path: rel.path,
          });
        }
      }
    }
    tableAt(sheet, r, c) { return this.tables.find((t) => t.sheet === sheet && r >= t.g.r1 && r <= t.g.r2 && c >= t.g.c1 && c <= t.g.c2); }

    // Other package parts (tables, drawings, charts): parsed on demand, written back only if marked.
    async part(path) {
      if (this.partDocs.has(path)) return this.partDocs.get(path);
      const f = this.zip.file(path);
      const doc = f ? parseXml(await f.async('string')) : null;
      this.partDocs.set(path, doc);
      return doc;
    }
    markPart(path) { this.dirtyParts.add(path); }
    async partBytes(path) { const f = this.zip.file(path); return f ? f.async('arraybuffer') : null; }
    async rels(partPath) {
      if (!this.relsCache.has(partPath)) this.relsCache.set(partPath, await readRels(this.zip, partPath));
      return this.relsCache.get(partPath);
    }
    async sheetRels(sheet) { return sheet.isNew ? { list: [] } : this.rels(sheet.path); }
    async drawingPath(sheet) {
      const r = (await this.sheetRels(sheet)).list.find((x) => x.type.endsWith('/drawing'));
      return r ? r.path : null;
    }
    async chartPaths(sheet) {
      const d = await this.drawingPath(sheet);
      if (!d) return [];
      return (await this.rels(d)).list.filter((x) => x.type.endsWith('/chart')).map((x) => x.path);
    }

    sheet(name) { const n = String(name).toLowerCase(); return this.sheets.find((s) => s.name.toLowerCase() === n); }

    dateToSerial(d) {
      const epoch = this.date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
      return (d.getTime() - d.getTimezoneOffset() * 0 - epoch) / 86400000;
    }

    // Defined name -> reference for the formula engine.
    resolveName(name, sheetName) {
      const n = name.toLowerCase();
      const sheetIdx = this.sheets.findIndex((s) => s.name.toLowerCase() === String(sheetName).toLowerCase());
      const d = this.definedNames.find((x) => x.name.toLowerCase() === n && x.local === sheetIdx) || this.definedNames.find((x) => x.name.toLowerCase() === n && x.local == null);
      if (!d) return undefined;
      const toks = scanRefs(d.ref).filter((t) => t.kind === 'ref');
      if (!toks.length) return undefined;
      const sh = toks[0].sheet || sheetName;
      if (toks.length >= 2 && d.ref.includes(':')) return { sheet: sh, from: { row: Math.min(toks[0].row, toks[1].row), col: Math.min(toks[0].col, toks[1].col) }, to: { row: Math.max(toks[0].row, toks[1].row), col: Math.max(toks[0].col, toks[1].col) } };
      return { sheet: sh, row: toks[0].row, col: toks[0].col };
    }

    addSheet(name) {
      const doc = this.wbDoc;
      const ids = this.sheets.map((s) => s.sheetId);
      const sheetId = Math.max(0, ...ids) + 1;
      let n = sheetId;
      const dir = this.wbPath.replace(/[^/]+$/, 'worksheets/');
      while (this.zip.file(`${dir}sheet${n}.xml`) || this.newParts[`${dir}sheet${n}.xml`]) n++;
      const path = `${dir}sheet${n}.xml`;
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="${NS}" xmlns:r="${RNS}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`;
      this.newParts[path] = xml;
      const rid = this.addRel(this.wbRels, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', path.slice(this.wbPath.replace(/[^/]+$/, '').length));
      const sheetsEl = child(doc.documentElement, 'sheets');
      const el = makeEl(doc, 'sheet', { name, sheetId });
      el.setAttributeNS(RNS, 'r:id', rid);
      sheetsEl.append(el);
      this.contentOverrides = this.contentOverrides || [];
      this.contentOverrides.push({ part: '/' + path, type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' });
      const sheet = new Sheet(this, { name, sheetId, rid, path, state: 'visible', el, kind: 'worksheet' });
      sheet.doc = parseXml(xml);
      sheet.sheetData = sheet.doc.getElementsByTagNameNS(NS, 'sheetData')[0];
      sheet.rowEls = new Map();
      sheet.isNew = true;
      this.sheets.push(sheet);
      this.wbChanged = true;
      return sheet;
    }

    renameSheet(sheet, newName) {
      const old = sheet.name;
      sheet.name = newName;
      sheet.el.setAttribute('name', newName);
      for (const s of this.sheets) for (const [k, cell] of s.cells) {
        if (!cell.formula) continue;
        const nf = renameSheetInFormula(cell.formula, old, newName);
        if (nf !== cell.formula) { if (cell.shared != null) s.unshare(cell.shared); cell.formula = nf; cell.formulaChanged = true; s.changed.add(k); }
      }
      for (const d of this.definedNames) { const nf = renameSheetInFormula(d.ref, old, newName); if (nf !== d.ref) { d.ref = nf; d.el.textContent = nf; } }
      this.wbChanged = true;
      this.formulasChanged = true;
    }

    // Excel keeps a hidden _xlnm._FilterDatabase name for every autofilter.
    syncFilterNames() {
      const root = this.wbDoc.documentElement;
      this.sheets.forEach((s, i) => {
        if (!s.filterChanged) return;
        let d = this.definedNames.find((x) => x.name === '_xlnm._FilterDatabase' && x.local === i);
        if (s.autoFilter && s.autoFilter.ref) {
          const g = parseRange(s.autoFilter.ref);
          const ref = `${quoteSheet(s.name)}!$${colName(g.c1)}$${g.r1}:$${colName(g.c2)}$${g.r2}`;
          if (!d) {
            let dn = child(root, 'definedNames');
            if (!dn) dn = insertOrdered(root, makeEl(this.wbDoc, 'definedNames'), WB_ORDER);
            const el = makeEl(this.wbDoc, 'definedName', { name: '_xlnm._FilterDatabase', localSheetId: i, hidden: '1' });
            dn.append(el);
            d = { name: '_xlnm._FilterDatabase', local: i, ref, el };
            this.definedNames.push(d);
          }
          d.ref = ref; d.el.textContent = ref;
        } else if (d) {
          d.el.parentNode.removeChild(d.el);
          this.definedNames = this.definedNames.filter((x) => x !== d);
        }
        this.wbChanged = true;
      });
    }

    addRel(rels, type, target) {
      const used = new Set(rels.list.map((r) => r.id));
      let n = 1; while (used.has(`rId${n}`)) n++;
      const el = rels.doc.createElementNS(PKG, 'Relationship');
      el.setAttribute('Id', `rId${n}`); el.setAttribute('Type', type); el.setAttribute('Target', target);
      rels.doc.documentElement.append(el);
      rels.list.push({ id: `rId${n}`, type, target, el });
      rels.changed = true;
      return `rId${n}`;
    }

    async save() {
      const zip = await JSZip.loadAsync(this.bytes);
      let anyFormula = false;
      for (const s of this.sheets) {
        if (s.unsupported) continue;
        for (const k of s.changed) { const c = s.cells.get(k); if (c && (c.formula || c.formulaChanged)) anyFormula = true; }
        if (s.save(this.sstObj) || s.isNew) zip.file(s.path, serialize(s.doc));
      }
      if (this.sstObj.dirty) {
        const root = this.sstObj.doc.documentElement;
        root.setAttribute('uniqueCount', String(kids(root, 'si').length));
        root.setAttribute('count', String(Math.max(+(root.getAttribute('count') || 0), kids(root, 'si').length)));
        if (!this.sstPath) {
          this.sstPath = this.wbPath.replace(/[^/]+$/, 'sharedStrings.xml');
          this.addRel(this.wbRels, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', 'sharedStrings.xml');
          this.contentOverrides = this.contentOverrides || [];
          this.contentOverrides.push({ part: '/' + this.sstPath, type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml' });
        }
        zip.file(this.sstPath, serialize(this.sstObj.doc));
      }
      if (this.styles.dirty) zip.file(this.stylesPath, serialize(this.styles.doc));
      // Formulas changed: Excel's calculation chain is now stale (it would report a repair), so drop it and
      // ask Excel to recalculate everything when the file is next opened.
      if (anyFormula || this.formulasChanged) {
        const calcRel = this.wbRels.list.find((r) => r.type.endsWith('/calcChain'));
        if (calcRel) {
          zip.remove(calcRel.path);
          calcRel.el.parentNode.removeChild(calcRel.el);
          this.wbRels.list = this.wbRels.list.filter((r) => r !== calcRel);
          this.wbRels.changed = true;
          this.removeOverrides = ['/' + calcRel.path];
        }
        const root = this.wbDoc.documentElement;
        const calcPr = child(root, 'calcPr') || insertOrdered(root, makeEl(this.wbDoc, 'calcPr', { calcId: '191029' }), WB_ORDER);
        calcPr.setAttribute('fullCalcOnLoad', '1');
        this.wbChanged = true;
      }
      for (const p of this.dirtyParts) { const d = this.partDocs.get(p); if (d) zip.file(p, serialize(d)); }
      this.syncFilterNames();
      if (this.wbChanged) zip.file(this.wbPath, serialize(this.wbDoc));
      if (this.wbRels.changed) zip.file(this.wbRels.relsPath, serialize(this.wbRels.doc));
      for (const [p, data] of Object.entries(this.newParts)) if (!this.sheets.some((s) => s.path === p)) zip.file(p, data);
      if (this.contentOverrides || this.removeOverrides) {
        const ct = parseXml(await zip.file('[Content_Types].xml').async('string'));
        const root = ct.documentElement;
        for (const o of this.contentOverrides || []) {
          if ([...root.getElementsByTagNameNS(CT_NS, 'Override')].some((e) => e.getAttribute('PartName') === o.part)) continue;
          const e = ct.createElementNS(CT_NS, 'Override');
          e.setAttribute('PartName', o.part); e.setAttribute('ContentType', o.type);
          root.append(e);
        }
        for (const p of this.removeOverrides || []) for (const e of [...root.getElementsByTagNameNS(CT_NS, 'Override')]) if (e.getAttribute('PartName') === p) root.removeChild(e);
        zip.file('[Content_Types].xml', serialize(ct));
      }
      const out = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      // The saved package is the new baseline; the in-memory XML already matches it.
      this.bytes = out;
      this.zip = await JSZip.loadAsync(out);
      this.newParts = {};
      this.contentOverrides = null;
      this.removeOverrides = null;
      this.wbChanged = false;
      this.formulasChanged = false;
      this.wbRels.changed = false;
      this.sstObj.dirty = false;
      this.styles.dirty = false;
      this.dirtyParts.clear();
      for (const s of this.sheets) {
        s.changed.clear(); s.colsChanged = false; s.mergesChanged = false; s.isNew = false;
        s.rowsChanged.clear(); s.filterChanged = false; s.structureChanged = false;
        for (const c of s.cells.values()) delete c.formulaChanged;
      }
      return out;
    }
  }

  // ---------- a new, empty workbook (what Excel itself writes for Book1, minus optional parts)

  const BLANK_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;

  const BLANK_THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

  async function blankBytes() {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG}"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
    zip.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Doc Studio</Application></Properties>`);
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    zip.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Doc Studio</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS}" xmlns:r="${RNS}"><workbookPr defaultThemeVersion="166925"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`);
    zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG}"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
    zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS}" xmlns:r="${RNS}"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
    zip.file('xl/styles.xml', BLANK_STYLES);
    zip.file('xl/theme/theme1.xml', BLANK_THEME);
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }

  return {
    Workbook, XErr, colName, colNum, addr, parseAddr, parseRange, rangeRef, key, unkey, shiftFormula, scanRefs, quoteSheet,
    blankBytes, MAX_ROW, MAX_COL,
  };
})();
