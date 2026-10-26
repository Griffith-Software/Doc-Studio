'use strict';
// Formula engine bundle (built to renderer/vendor/formula.js by `npm run bundle`).
// fast-formula-parser (MIT) parses and evaluates; functions it lacks come from @formulajs/formulajs (MIT) or are
// implemented here. Exposes window.DDFormula.

const FormulaParser = require('fast-formula-parser');
const { DepParser, FormulaError } = FormulaParser;
const formulajs = require('@formulajs/formulajs');
const { Infix } = require('fast-formula-parser/formulas/operators');

const { FormulaHelpers, Types } = require('fast-formula-parser/formulas/helpers');

// Empty cells come through as null. Excel treats a blank as "" where text is expected, 0 for numbers and FALSE
// for logicals; the parser would otherwise turn null into "[object Object]".
const accept = FormulaHelpers.accept.bind(FormulaHelpers);
FormulaHelpers.accept = (param, type, ...rest) => {
  const t = Array.isArray(type) ? type[0] : type;
  if (param && typeof param === 'object' && !Array.isArray(param) && param.value === null && !param.omitted) {
    const blank = t === Types.STRING ? '' : t === Types.BOOLEAN ? false : t === Types.NUMBER || t === Types.NUMBER_NO_BOOLEAN ? 0 : param.value;
    param = { ...param, value: blank };
  }
  return accept(param, type, ...rest);
};

// Excel compares text without regard to case ("a" = "A" is TRUE), and a blank equals "", 0 and FALSE.
const compareOp = Infix.compareOp.bind(Infix);
Infix.compareOp = (value1, infix, value2, isArray1, isArray2) => {
  const low = (v) => (typeof v === 'string' ? v.toLowerCase() : Array.isArray(v) ? v.map((r) => (Array.isArray(r) ? r.map(low) : low(r))) : v);
  const blank = (v, other) => (v !== null && v !== undefined ? v : typeof other === 'string' ? '' : typeof other === 'boolean' ? false : 0);
  const a = blank(value1, value2);
  const b = blank(value2, value1);
  return compareOp(low(a), infix, low(b), isArray1, isArray2);
};

const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#CALC!', '#SPILL!'];

function toFormulaError(e) {
  if (e instanceof FormulaError) return e;
  const msg = e && (e.message || e.toString());
  if (ERRORS.includes(msg)) return new FormulaError(msg);
  return FormulaError.VALUE;
}

// formulajs functions take plain values (ranges as 2D arrays) and return values or Error objects.
function wrapFormulajs(fn) {
  return (...args) => {
    const vals = args.map((a) => (a.omitted ? undefined : a.value));
    let r;
    try { r = fn(...vals); } catch (e) { throw toFormulaError(e); }
    if (r instanceof Error) throw toFormulaError(r);
    if (r === undefined || r === null) return '';
    if (typeof r === 'number' && !Number.isFinite(r)) throw FormulaError.NUM;
    if (Array.isArray(r) && !Array.isArray(r[0])) return [r];
    return r;
  };
}

function formulajsByName() {
  const out = {};
  for (const [k, v] of Object.entries(formulajs)) {
    if (typeof v === 'function') out[k] = v;
    else if (v && typeof v === 'object') for (const [s, f] of Object.entries(v)) if (typeof f === 'function') out[`${k}.${s}`] = f;
  }
  return out;
}

const flat = (v) => (Array.isArray(v) ? v.flat(Infinity) : [v]);
const isBlank = (v) => v === undefined || v === null || v === '';

function colToNum(s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }

// Parse "Sheet!A1", "'My Sheet'!A1:B5", "A1" into a reference object for INDIRECT.
function parseRefText(text, defaultSheet) {
  const m = String(text).trim().match(/^(?:(?:'((?:[^']|'')+)'|([^!:'"]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/);
  if (!m) throw FormulaError.REF;
  const sheet = m[1] ? m[1].replace(/''/g, "'") : (m[2] || defaultSheet);
  const a = { row: +m[4], col: colToNum(m[3]) };
  if (!m[5]) return { sheet, row: a.row, col: a.col };
  const b = { row: +m[6], col: colToNum(m[5]) };
  return { sheet, from: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) }, to: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) } };
}

const CUSTOM = {
  'STDEV': 'STDEV.S', 'VAR': 'VAR.S', 'MODE': 'MODE.SNGL', 'RANK': 'RANK.EQ', 'STDEVP': 'STDEV.P', 'VARP': 'VAR.P',
  'PERCENTILE': 'PERCENTILE.INC', 'QUARTILE': 'QUARTILE.INC', 'NORMDIST': 'NORM.DIST', 'NORMINV': 'NORM.INV',
};

const OWN = {
  XLOOKUP(lookup, lookupArr, returnArr, ifNotFound, matchMode, searchMode) {
    const key = lookup.value;
    const keys = flat(lookupArr.value);
    const rows = Array.isArray(returnArr.value) ? returnArr.value : [[returnArr.value]];
    const vertical = rows.length === keys.length;
    const mm = matchMode && !matchMode.omitted ? matchMode.value : 0;
    const reverse = searchMode && !searchMode.omitted && searchMode.value < 0;
    const order = keys.map((_, i) => i);
    if (reverse) order.reverse();
    const eq = (a, b) => (typeof a === 'string' && typeof b === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b);
    let hit = order.find((i) => eq(keys[i], key));
    if (hit === undefined && (mm === -1 || mm === 1)) {
      let best;
      for (const i of order) {
        const k = keys[i];
        if (typeof k !== typeof key) continue;
        if (mm === -1 && k < key && (best === undefined || k > keys[best])) best = i;
        if (mm === 1 && k > key && (best === undefined || k < keys[best])) best = i;
      }
      hit = best;
    }
    if (hit === undefined && mm === 2) {
      const re = new RegExp('^' + String(key).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
      hit = order.find((i) => re.test(String(keys[i])));
    }
    if (hit === undefined) {
      if (ifNotFound && !ifNotFound.omitted) return ifNotFound.value;
      throw FormulaError.NA;
    }
    if (vertical) return rows[hit].length === 1 ? rows[hit][0] : [rows[hit]];
    return rows.length === 1 ? rows[0][hit] : rows.map((r) => [r[hit]]);
  },
  MATCH(lookup, arr, type) {
    const key = lookup.value;
    const list = flat(arr.value);
    const t = type && !type.omitted ? type.value : 1;
    const low = (v) => (typeof v === 'string' ? v.toLowerCase() : v);
    if (t === 0) {
      if (typeof key === 'string' && /[*?]/.test(key)) {
        const re = new RegExp('^' + String(key).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
        const i = list.findIndex((v) => re.test(String(v)));
        if (i >= 0) return i + 1;
        throw FormulaError.NA;
      }
      const i = list.findIndex((v) => low(v) === low(key));
      if (i >= 0) return i + 1;
      throw FormulaError.NA;
    }
    // Sorted search: largest value <= key (type 1) or smallest value >= key (type -1).
    let best = -1;
    for (let i = 0; i < list.length; i++) {
      const v = list[i];
      if (typeof v !== typeof key || isBlank(v)) continue;
      if (t > 0 ? low(v) <= low(key) : low(v) >= low(key)) best = i; else break;
    }
    if (best < 0) throw FormulaError.NA;
    return best + 1;
  },
  FILTER(arr, include, ifEmpty) {
    const rows = Array.isArray(arr.value) ? arr.value : [[arr.value]];
    const inc = flat(include.value);
    const out = rows.filter((_, i) => inc[i] === true || (typeof inc[i] === 'number' && inc[i] !== 0));
    if (!out.length) { if (ifEmpty && !ifEmpty.omitted) return ifEmpty.value; throw new FormulaError('#CALC!'); }
    return out;
  },
  SEQUENCE(rows, cols, start, step) {
    const r = rows.value, c = cols && !cols.omitted ? cols.value : 1;
    const s = start && !start.omitted ? start.value : 1, st = step && !step.omitted ? step.value : 1;
    const out = [];
    for (let i = 0; i < r; i++) { const row = []; for (let j = 0; j < c; j++) row.push(s + (i * c + j) * st); out.push(row); }
    return out;
  },
  // Shows its label (or the address); Doc Studio opens the link on Ctrl+click.
  HYPERLINK(link, label) { return label && !label.omitted ? label.value : link.value; },
  CONCATENATE(...args) { return args.map((a) => flat(a.value).map((v) => (isBlank(v) ? '' : v === true ? 'TRUE' : v === false ? 'FALSE' : String(v))).join('')).join(''); },
};

// Functions that need the raw reference (context, ...rawArgs).
const OWN_CONTEXT = {
  // The parser's IF tests the reference object, not its value; Excel treats blank/0 as FALSE and text as #VALUE!.
  IF(ctx, test, ifTrue, ifFalse) {
    const v = ctx.utils.extractRefValue(test).val;
    const truth = (x) => {
      if (x instanceof FormulaError) throw x;
      if (x === null || x === undefined || x === '') return false;
      if (typeof x === 'boolean') return x;
      if (typeof x === 'number') return x !== 0;
      if (typeof x === 'string') { const u = x.toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; }
      throw FormulaError.VALUE;
    };
    const pick = (cond) => {
      const branch = cond ? ifTrue : ifFalse;
      if (branch === undefined) return cond ? 0 : false;
      if (branch === null) return 0;
      const val = ctx.utils.extractRefValue(branch).val;
      return val === null || val === undefined ? 0 : val;
    };
    if (Array.isArray(v)) return v.map((row) => row.map((x) => { const b = pick(truth(x)); return Array.isArray(b) ? b[0][0] : b; }));
    return pick(truth(v));
  },
  INDIRECT(ctx, refText, a1) {
    const text = ctx.utils.extractRefValue(refText).val;
    void a1;
    return { ref: parseRefText(text, ctx.position.sheet) };
  },
  OFFSET(ctx, ref, rows, cols, height, width) {
    const r = ref && ref.ref;
    if (!r) throw FormulaError.VALUE;
    const num = (a, d) => (a == null ? d : +ctx.utils.extractRefValue(a).val || 0);
    const from = r.from ? { ...r.from } : { row: r.row, col: r.col };
    const to = r.to ? { ...r.to } : { row: r.row, col: r.col };
    const dr = num(rows, 0), dc = num(cols, 0);
    const h = num(height, to.row - from.row + 1), w = num(width, to.col - from.col + 1);
    const row = from.row + dr, col = from.col + dc;
    if (row < 1 || col < 1 || h < 1 || w < 1) throw FormulaError.REF;
    const sheet = r.sheet || ctx.position.sheet;
    if (h === 1 && w === 1) return { ref: { sheet, row, col } };
    return { ref: { sheet, from: { row, col }, to: { row: row + h - 1, col: col + w - 1 } } };
  },
};

// callbacks: { cell(sheet,row,col) -> value | FormulaError | undefined, range(sheet,r1,c1,r2,c2) -> 2D array,
//              name(name, sheet) -> ref | undefined, sheets() -> [names] }
function createEngine(callbacks) {
  const fnames = formulajsByName();
  const functions = { ...OWN };
  const parser = new FormulaParser({
    functions,
    functionsNeedContext: OWN_CONTEXT,
    onCell: ({ sheet, row, col }) => callbacks.cell(sheet, row, col),
    onRange: (ref) => callbacks.range(ref.sheet, ref.from.row, ref.from.col, ref.to.row, ref.to.col),
    onVariable: (name, sheet) => callbacks.name(name, sheet),
  });
  // INDIRECT/OFFSET need the unevaluated reference.
  for (const n of ['INDIRECT', 'OFFSET']) if (!parser.funsNeedContextAndNoDataRetrieve.includes(n)) parser.funsNeedContextAndNoDataRetrieve.push(n);
  // Fill gaps from formulajs: names the parser lacks, or has only as "not implemented" stubs (probed once).
  const notImplemented = (e) => {
    for (let d = e; d; d = d.details) if (/not implemented/i.test(String(d.message || ''))) return true;
    return false;
  };
  const has = (n) => {
    if (!parser.functions[n]) return false;
    try { parser.parse(`${n}(1)`, { sheet: '__probe__', row: 1, col: 1 }); return true; } catch (e) { return !notImplemented(e); }
  };
  for (const [name, fn] of Object.entries(fnames)) if (!has(name) && !OWN[name] && !OWN_CONTEXT[name]) parser.functions[name] = wrapFormulajs(fn);
  for (const [alias, target] of Object.entries(CUSTOM)) if (!has(alias) && parser.functions[target]) parser.functions[alias] = parser.functions[target];
  for (const [name, fn] of Object.entries(callbacks.functions || {})) parser.functions[name.toUpperCase()] = fn;

  const deps = new DepParser({ onVariable: (name, sheet) => callbacks.name(name, sheet) });
  return {
    // Evaluate one formula (text without '='). Returns a value, a 2D array (array results) or a FormulaError.
    evaluate(formula, sheet, row, col, allowArray) {
      try {
        return parser.parse(formula, { sheet, row, col }, !!allowArray);
      } catch (e) {
        // The parser wraps errors in #ERROR!; surface the Excel error underneath (#NAME?, #VALUE!...).
        let d = e;
        while (d && d.details instanceof FormulaError) d = d.details;
        if (d instanceof FormulaError && ERRORS.includes(d.error)) return d;
        if (d && /not implemented/i.test(String(d.message))) return FormulaError.NAME;
        return d instanceof FormulaError && d.error !== '#ERROR!' ? d : FormulaError.VALUE;
      }
    },
    // References a formula reads: [{sheet,row,col}] or [{sheet,from,to}].
    dependencies(formula, sheet, row, col) {
      try { return deps.parse(formula, { sheet, row, col }); } catch { return []; }
    },
    names: () => Object.keys(parser.functions).filter((n) => /^[A-Z][A-Z0-9.]*$/.test(n)).sort(),
    isError: (v) => v instanceof FormulaError,
    error: (code) => new FormulaError(code),
    FormulaError,
  };
}

module.exports = { createEngine, FormulaError };
