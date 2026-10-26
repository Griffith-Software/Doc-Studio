'use strict';
// Rich Text (.rtf) editor for PCs without Word or LibreOffice.
//
// The file is drawn by rtf.js and edited in place. Saving writes RTF from what is on screen: fonts, sizes,
// colours, bold/italic/underline/strike, super/subscript, highlight, alignment, indents, spacing, line breaks and
// pictures, with the original page size and margins. A file that was not changed is saved byte for byte.

const RtfEditor = (() => {
  class Editor {
    constructor(host, opts) { this.host = host; this.opts = opts || {}; }

    async load(bytes) {
      this.bytes = bytes;
      const els = await render(bytes);
      const page = document.createElement('article');
      page.className = 'paper rtf-view rtf-edit';
      page.contentEditable = this.opts.readOnly ? 'false' : 'true';
      els.forEach((el) => page.append(el));
      if (!page.textContent.trim() && !page.querySelector('img')) page.replaceChildren(Object.assign(document.createElement('div'), { innerHTML: '<br>' }));
      page.addEventListener('input', () => { this.edited = true; this.opts.onChange && this.opts.onChange(); });
      page.addEventListener('paste', (e) => { e.preventDefault(); const t = e.clipboardData.getData('text/plain'); if (t) document.execCommand('insertText', false, t); });
      this.article = page;
      this.host.replaceChildren(page);
      // Page setup from the original header (kept on save).
      // (Word puts them after its long style sheet, so look through the whole file; the first of each wins.)
      const all = new TextDecoder('latin1').decode(new Uint8Array(bytes));
      const words = new Map();
      for (const m of all.matchAll(/\\(paperw|paperh|margl|margr|margt|margb|landscape)(-?\d*)(?![a-z])/g)) if (!words.has(m[1])) words.set(m[1], m[0]);
      this.pageWords = [...words.values()].join('');
      this.edited = false;
      return this;
    }

    get stage() { return this.article; }
    selectedParagraphs() {
      const sel = getSelection();
      if (!sel.rangeCount) return [];
      const r = sel.getRangeAt(0);
      return [...this.article.children].filter((p) => r.intersectsNode(p));
    }
    align(ps, value) { for (const p of ps) p.style.textAlign = value; this.edited = true; this.opts.onChange && this.opts.onChange(); }
    async insertImage(bytes, ext, w, h) {
      let bin = '';
      const u8 = new Uint8Array(bytes);
      for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode(...u8.subarray(i, i + 0x8000));
      const b64 = btoa(bin);
      const max = this.article.clientWidth - 40;
      const s = Math.min(1, max / w);
      document.execCommand('insertHTML', false, `<img src="data:image/${ext === 'jpg' ? 'jpeg' : ext};base64,${b64}" style="width:${Math.round(w * s)}px;height:${Math.round(h * s)}px">`);
      this.edited = true;
      this.opts.onChange && this.opts.onChange();
    }

    async save() {
      if (!this.edited) return this.bytes;
      const out = await writeRtf(this.article, this.pageWords);
      this.bytes = out;
      this.edited = false;
      return out;
    }
    destroy() { this.host.replaceChildren(); }
  }

  // rtf.js can't always look up the code page for \'hh character escapes (files written by Word list fonts in many
  // character sets). When it fails, draw again with those escapes turned into Unicode escapes (Windows-1252).
  async function render(bytes) {
    RTFJS.loggingEnabled(false); WMFJS.loggingEnabled(false); EMFJS.loggingEnabled(false);
    try { return await new RTFJS.Document(bytes).render(); } catch (err) {
      const cp = new TextDecoder('windows-1252');
      const text = new TextDecoder('latin1').decode(new Uint8Array(bytes))
        .replace(/\\'([0-9a-fA-F]{2})/g, (m, hx) => { const c = cp.decode(Uint8Array.of(parseInt(hx, 16))).charCodeAt(0); return `\\u${c > 32767 ? c - 65536 : c}?`; });
      const out = new Uint8Array(text.length);
      for (let i = 0; i < text.length; i++) out[i] = text.charCodeAt(i) & 0xff;
      return new RTFJS.Document(out.buffer).render();
    }
  }

  // ---------------------------------------------------------------- HTML -> RTF

  const esc = (s) => {
    let out = '';
    for (const ch of s) {
      const c = ch.codePointAt(0);
      if (ch === '\\' || ch === '{' || ch === '}') out += '\\' + ch;
      else if (ch === '\t') out += '\\tab ';
      else if (c < 0x80) out += ch;
      else if (c <= 0xffff) out += `\\u${c > 32767 ? c - 65536 : c}?`;
      else { const hi = Math.floor((c - 0x10000) / 0x400) + 0xd800, lo = ((c - 0x10000) % 0x400) + 0xdc00; out += `\\u${hi - 65536}?\\u${lo - 65536}?`; }
    }
    return out;
  };
  const twips = (pxv) => Math.round((parseFloat(pxv) || 0) * 15);
  const rgb = (c) => { const m = /rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/.exec(c || ''); return m && (m[4] === undefined || +m[4] > 0) ? [+m[1], +m[2], +m[3]] : null; };

  async function writeRtf(root, pageWords) {
    const fonts = [], colors = [];
    const fontIdx = (f) => { const name = (f || 'Calibri').split(',')[0].replace(/["']/g, '').trim() || 'Calibri'; let i = fonts.indexOf(name); if (i < 0) { fonts.push(name); i = fonts.length - 1; } return i; };
    const colorIdx = (c) => { const v = rgb(c); if (!v) return 0; const k = v.join(','); let i = colors.indexOf(k); if (i < 0) { colors.push(k); i = colors.length - 1; } return i + 1; };
    fontIdx('Calibri');
    const body = [];
    const blocks = [];
    // Block = an element whose children are inline content. Tables and list items are their own kinds.
    const isBlock = (el) => /^(DIV|P|H[1-6]|LI|BLOCKQUOTE|PRE|TABLE|UL|OL)$/.test(el.tagName) || getComputedStyle(el).display === 'block';
    const isList = (el) => el.nodeType === 1 && (el.tagName === 'UL' || el.tagName === 'OL');
    const collect = (el) => {
      if (el.tagName === 'TABLE') { blocks.push({ table: el }); return; }
      if (el.tagName === 'LI') {
        // The item's own text (paragraphs inside it included) with its marker; nested lists follow as their own items.
        blocks.push({ li: el, nodes: [...el.childNodes].filter((c) => !isList(c)), parent: el });
        for (const c of el.children) if (isList(c)) collect(c);
        return;
      }
      const hasBlockKid = [...el.children].some(isBlock);
      if (!hasBlockKid) { blocks.push(el); return; }
      // Mixed content: inline runs between blocks become their own paragraph.
      let run = [];
      const flushRun = () => { if (run.some((n) => (n.textContent || '').trim() || n.nodeName === 'IMG')) blocks.push({ nodes: run, parent: el }); run = []; };
      for (const c of el.childNodes) {
        if (c.nodeType === 1 && isBlock(c)) { flushRun(); collect(c); } else run.push(c);
      }
      flushRun();
    };
    for (const c of root.childNodes) {
      if (c.nodeType === 1 && isBlock(c)) collect(c);
      else if ((c.textContent || '').trim()) blocks.push({ nodes: [c], parent: root });
    }

    // Text runs with their character formatting.
    const runs = async (nodes, el) => {
      let line = '';
      const emit = async (n) => {
        if (n.nodeType === 3) {
          const text = n.data.replace(/\u00a0/g, ' ').replace(/\n/g, ' ');
          if (!text) return;
          const s = getComputedStyle(n.parentElement);
          let fmt = `\\f${fontIdx(s.fontFamily)}\\fs${Math.round(parseFloat(s.fontSize) * 0.75 * 2)}`;
          const w = s.fontWeight;
          if (w === 'bold' || +w >= 600) fmt += '\\b';
          if (s.fontStyle === 'italic') fmt += '\\i';
          // Decorations come from any ancestor inside the paragraph.
          let u = false, st = false, sup = null, bg = null;
          for (let p = n.parentElement; p && p !== root.parentElement; p = p.parentElement) {
            const ps = getComputedStyle(p);
            const d = ps.textDecorationLine || '';
            if (/underline/.test(d)) u = true;
            if (/line-through/.test(d)) st = true;
            if (!sup && (ps.verticalAlign === 'super' || ps.verticalAlign === 'sub')) sup = ps.verticalAlign;
            if (!bg && rgb(ps.backgroundColor) && p !== root) bg = ps.backgroundColor;
            if (p === el) break;
          }
          if (u) fmt += '\\ul';
          if (st) fmt += '\\strike';
          if (sup) fmt += sup === 'super' ? '\\super' : '\\sub';
          const ci = colorIdx(s.color);
          if (ci) fmt += `\\cf${ci}`;
          if (bg) fmt += `\\highlight${colorIdx(bg)}\\cb${colorIdx(bg)}`;
          line += `{${fmt} ${esc(text)}}`;
          return;
        }
        if (n.nodeType !== 1) return;
        if (n.tagName === 'BR') { line += '\\line '; return; }
        if (n.tagName === 'IMG') { line += await pict(n); return; }
        if (n.tagName === 'INPUT' && n.type === 'checkbox') { line += n.checked || n.hasAttribute('checked') ? '\\u9745?\\tab ' : '\\u9744?\\tab '; return; }
        // A paragraph inside a list item or table cell continues on a new line.
        const blockish = /^(P|DIV|H[1-6])$/.test(n.tagName);
        if (blockish && line) line += '\\line ';
        for (const c of n.childNodes) await emit(c);
      };
      for (const n of nodes) await emit(n);
      return line;
    };
    const paraProps = (el) => {
      const cs = getComputedStyle(el);
      let pard = '';
      const al = cs.textAlign;
      pard += al === 'center' ? '\\qc' : al === 'right' || al === 'end' ? '\\qr' : al === 'justify' ? '\\qj' : '\\ql';
      const sb = twips(cs.marginTop), sa = twips(cs.marginBottom);
      if (sb) pard += `\\sb${sb}`;
      if (sa) pard += `\\sa${sa}`;
      const lh = parseFloat(cs.lineHeight), fsz = parseFloat(cs.fontSize);
      if (lh && fsz && lh / fsz > 1.05) pard += `\\sl${Math.round(lh / fsz * 240)}\\slmult1`;
      return { pard, cs };
    };

    for (const b of blocks) {
      if (b.table) {
        // Table rows: equal-width bordered cells across the text width (6.5in); merged cells keep their span.
        const rows = [...b.table.rows];
        const cols = Math.max(1, ...rows.map((r) => [...r.cells].reduce((n, c) => n + (c.colSpan || 1), 0)));
        const unit = Math.floor(9360 / cols);
        for (const r of rows) {
          let row = '\\trowd\\trgaph108\\trleft0';
          let x = 0;
          const border = '\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10';
          for (const c of r.cells) {
            x += unit * (c.colSpan || 1);
            const bg = rgb(getComputedStyle(c).backgroundColor);
            row += `${border}${bg ? `\\clcbpat${colorIdx(getComputedStyle(c).backgroundColor)}` : ''}\\cellx${x}`;
          }
          for (const c of r.cells) {
            const { pard } = paraProps(c);
            row += `\\pard\\plain\\intbl${pard} ${await runs([...c.childNodes], c)}\\cell`;
          }
          body.push(row + '\\row');
        }
        continue;
      }
      if (b.li) {
        // List item: bullet or number, hanging indent per nesting level.
        const list = b.li.parentElement;
        let depth = 0;
        for (let p = list; p && p !== root; p = p.parentElement) if (isList(p)) depth++;
        let marker = '\\u8226?';
        if (list && list.tagName === 'OL') {
          const items = [...list.children].filter((c) => c.tagName === 'LI');
          const n = (list.start || 1) + items.indexOf(b.li);
          const type = (getComputedStyle(list).listStyleType || 'decimal');
          marker = esc(type === 'lower-alpha' ? String.fromCharCode(96 + n) : type === 'upper-alpha' ? String.fromCharCode(64 + n) : type === 'lower-roman' ? toRoman(n).toLowerCase() : type === 'upper-roman' ? toRoman(n) : String(n)) + '.';
        } else if (depth === 2) marker = '\\u9702?';
        else if (depth >= 3) marker = '\\u9642?';
        const { pard } = paraProps(b.li);
        const li = 360 + 360 * depth;
        body.push(`\\pard\\plain${pard}\\fi-360\\li${li}\\tx${li} ${marker}\\tab ${await runs(b.nodes, b.li)}\\par`);
        continue;
      }
      const el = b.nodes ? b.parent : b;
      const { pard, cs } = paraProps(el);
      let props = '\\pard\\plain' + pard;
      const li = twips(cs.marginLeft) + twips(cs.paddingLeft), ri = twips(cs.marginRight) + twips(cs.paddingRight), fi = twips(cs.textIndent);
      if (li) props += `\\li${li}`;
      if (ri) props += `\\ri${ri}`;
      if (fi) props += `\\fi${fi}`;
      body.push(`${props} ${await runs(b.nodes || [...el.childNodes], el)}\\par`);
    }
    const head = `{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl${fonts.map((f, i) => `{\\f${i}\\fnil ${esc(f)};}`).join('')}}{\\colortbl;${colors.map((c) => { const [r, g, b] = c.split(','); return `\\red${r}\\green${g}\\blue${b};`; }).join('')}}{\\*\\generator Doc Studio;}\\viewkind4\\uc1${pageWords || ''}\r\n`;
    return new TextEncoder().encode(head + body.join('\r\n') + '\r\n}').buffer;
  }

  function toRoman(n) {
    const map = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    let out = '';
    for (const [v, r] of map) while (n >= v) { out += r; n -= v; }
    return out;
  }

  // Pictures as PNG (\pngblip), sized like on screen.
  async function pict(img) {
    try {
      if (!img.complete) await img.decode();
      const c = document.createElement('canvas');
      c.width = img.naturalWidth; c.height = img.naturalHeight;
      c.getContext('2d').drawImage(img, 0, 0);
      const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
      const bytes = new Uint8Array(await blob.arrayBuffer());
      let hexs = '';
      for (let i = 0; i < bytes.length; i++) { hexs += bytes[i].toString(16).padStart(2, '0'); if (i % 64 === 63) hexs += '\r\n'; }
      const w = img.width || img.naturalWidth, h = img.height || img.naturalHeight;
      return `{\\pict\\pngblip\\picw${img.naturalWidth}\\pich${img.naturalHeight}\\picwgoal${twips(w)}\\pichgoal${twips(h)}\r\n${hexs}}`;
    } catch { return ''; }
  }

  function blank() {
    return new TextEncoder().encode('{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\fnil Calibri;}}{\\*\\generator Doc Studio;}\\viewkind4\\uc1\\paperw12240\\paperh15840\\margl1440\\margr1440\\margt1440\\margb1440\r\n\\pard\\sa160\\sl259\\slmult1\\f0\\fs22\\par\r\n}').buffer;
  }

  return { Editor, blank, writeRtf, render };
})();
