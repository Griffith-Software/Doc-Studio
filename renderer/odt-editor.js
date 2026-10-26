'use strict';
// OpenDocument text (.odt) editor that keeps the original file.
//
// content.xml is drawn as HTML with the document's own styles (paragraph/text style chains, default styles, font
// faces, list styles, tables, pictures, page size). Every paragraph remembers its XML node. Saving rebuilds only the
// paragraphs whose text or formatting changed: their runs reuse the original span styles, and new formatting goes
// into new automatic styles. Paragraphs created with Enter are copies of the paragraph they came from; deleted
// paragraphs are removed. Fields, notes, bookmarks, frames and anything else inside a paragraph are carried over
// untouched. Every other part of the package (styles.xml, meta, settings, pictures) is written back as it was.

const OdtEditor = (() => {
  const NS = {
    office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    xlink: 'http://www.w3.org/1999/xlink',
    svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
    manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
    loext: 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
  };
  const PREFIX = { [NS.office]: 'office', [NS.text]: 'text', [NS.style]: 'style', [NS.fo]: 'fo', [NS.table]: 'table', [NS.draw]: 'draw', [NS.xlink]: 'xlink', [NS.svg]: 'svg' };
  const parseXml = (t) => new DOMParser().parseFromString(t, 'application/xml');
  const serialize = (doc) => { const s = new XMLSerializer().serializeToString(doc); return s.startsWith('<?xml') ? s : '<?xml version="1.0" encoding="UTF-8"?>\n' + s; };
  const kids = (el, ns, name) => (el ? [...el.childNodes].filter((c) => c.nodeType === 1 && (!ns || c.namespaceURI === ns) && (!name || c.localName === name)) : []);
  const first = (el, ns, name) => kids(el, ns, name)[0] || null;
  const attr = (el, ns, name) => (el && el.hasAttributeNS(ns, name) ? el.getAttributeNS(ns, name) : null);
  const mk = (doc, ns, name, attrs) => { const e = doc.createElementNS(ns, `${PREFIX[ns]}:${name}`); for (const [k, v] of Object.entries(attrs || {})) { const [p, n] = k.split(':'); e.setAttributeNS(NS[p], k, v); void n; } return e; };
  const isPara = (n) => n && n.nodeType === 1 && n.namespaceURI === NS.text && (n.localName === 'p' || n.localName === 'h');

  // Lengths: "2cm", "0.5in", "12pt", "10mm", "1pc", "96px" -> CSS px.
  function px(v) {
    if (!v) return null;
    const m = /^(-?[\d.]+)(cm|mm|in|pt|pc|px)?$/.exec(String(v).trim());
    if (!m) return null;
    const n = +m[1];
    return { cm: 37.795, mm: 3.7795, in: 96, pt: 4 / 3, pc: 16, px: 1 }[m[2] || 'px'] * n;
  }
  const pxs = (v) => { const p = px(v); return p == null ? null : `${+p.toFixed(2)}px`; };

  // ---------------------------------------------------------------- styles

  class Styles {
    constructor(stylesDoc, contentDoc) {
      this.map = new Map();      // `${family}:${name}` -> { el, auto }
      this.defaults = new Map(); // family -> el
      this.lists = new Map();    // list style name -> el
      this.fonts = new Map();    // font face name -> CSS family
      for (const doc of [stylesDoc, contentDoc]) {
        if (!doc) continue;
        for (const ff of doc.getElementsByTagNameNS(NS.style, 'font-face')) this.fonts.set(attr(ff, NS.style, 'name'), attr(ff, NS.svg, 'font-family') || attr(ff, NS.style, 'name'));
        for (const holder of [...doc.getElementsByTagNameNS(NS.office, 'styles'), ...doc.getElementsByTagNameNS(NS.office, 'automatic-styles')]) {
          const auto = holder.localName === 'automatic-styles';
          for (const st of kids(holder)) {
            if (st.namespaceURI === NS.style && st.localName === 'style') this.map.set(`${attr(st, NS.style, 'family')}:${attr(st, NS.style, 'name')}`, { el: st, auto, doc });
            else if (st.namespaceURI === NS.style && st.localName === 'default-style') this.defaults.set(attr(st, NS.style, 'family'), st);
            else if (st.namespaceURI === NS.text && st.localName === 'list-style') this.lists.set(attr(st, NS.style, 'name'), st);
          }
        }
      }
    }
    get(family, name) { return name ? this.map.get(`${family}:${name}`) || null : null; }
    // Style elements from the most general to the style itself.
    chain(family, name, withDefault) {
      const out = [];
      let st = this.get(family, name);
      const seen = new Set();
      while (st && !seen.has(st.el)) { seen.add(st.el); out.unshift(st.el); st = this.get(family, attr(st.el, NS.style, 'parent-style-name')); }
      if (withDefault && this.defaults.has(family)) out.unshift(this.defaults.get(family));
      return out;
    }
    // Merged attributes of one kind of properties ("text-properties", "paragraph-properties"...) along a chain.
    props(chain, kind) {
      const out = {};
      for (const st of chain) {
        const p = first(st, NS.style, kind);
        if (!p) continue;
        for (const a of p.attributes) {
          if (a.localName === 'font-size' && /%$/.test(a.value) && out['font-size'] && /pt$/.test(out['font-size'])) out['font-size'] = `${parseFloat(out['font-size']) * parseFloat(a.value) / 100}pt`;
          else out[a.localName === 'font-name' && a.namespaceURI === NS.style ? 'font-name' : a.localName] = a.value;
          if (a.namespaceURI === NS.fo && a.localName === 'font-family') out['font-family'] = a.value;
        }
        if (kind === 'paragraph-properties') { const tabs = first(p, NS.style, 'tab-stops'); if (tabs) out.__tabs = tabs; }
      }
      return out;
    }
    textCss(t) {
      const css = {};
      const fam = t['font-name'] ? this.fonts.get(t['font-name']) || t['font-name'] : t['font-family'];
      if (fam) {
        const name = fam.replace(/^'|'$/g, '');
        const generic = /serif|times|georgia|cambria|garamond|book|roman|minion/i.test(name) && !/sans/i.test(name) ? 'serif' : /mono|courier|consol/i.test(name) ? 'monospace' : 'sans-serif';
        css['font-family'] = `${name.split(',')[0].includes(' ') ? `'${name}'` : name}, ${generic}`;
      }
      if (t['font-size']) css['font-size'] = /%$/.test(t['font-size']) ? `${parseFloat(t['font-size']) / 100}em` : t['font-size'];
      if (t['font-weight']) css['font-weight'] = t['font-weight'];
      if (t['font-style']) css['font-style'] = t['font-style'];
      const deco = [];
      if (t['text-underline-style'] && t['text-underline-style'] !== 'none') deco.push('underline');
      if (t['text-line-through-style'] && t['text-line-through-style'] !== 'none') deco.push('line-through');
      if (t['text-underline-style'] || t['text-line-through-style']) css['text-decoration'] = deco.join(' ') || 'none';
      if (t.color && t['use-window-font-color'] !== 'true') css.color = t.color;
      if (t['background-color'] && t['background-color'] !== 'transparent') css['background-color'] = t['background-color'];
      if (t['text-position']) { const v = t['text-position'].split(' ')[0]; css['vertical-align'] = v === 'super' || parseFloat(v) > 0 ? 'super' : v === 'sub' || parseFloat(v) < 0 ? 'sub' : 'baseline'; if (css['vertical-align'] !== 'baseline') css['font-size'] = css['font-size'] || '0.58em'; }
      if (t['font-variant'] === 'small-caps') css['font-variant'] = 'small-caps';
      if (t['text-transform'] && t['text-transform'] !== 'none') css['text-transform'] = t['text-transform'];
      if (t['letter-spacing'] && t['letter-spacing'] !== 'normal') css['letter-spacing'] = pxs(t['letter-spacing']);
      return css;
    }
    paraCss(p) {
      const css = {};
      const align = p['text-align'];
      if (align) css['text-align'] = { start: 'left', end: 'right', left: 'left', right: 'right', center: 'center', justify: 'justify' }[align] || 'left';
      for (const side of ['top', 'bottom', 'left', 'right']) { const v = pxs(p[`margin-${side}`]); if (v) css[`margin-${side}`] = v; }
      if (p['text-indent']) css['text-indent'] = pxs(p['text-indent']);
      if (p['line-height']) css['line-height'] = /%$/.test(p['line-height']) ? `${parseFloat(p['line-height']) / 100 * 1.15}` : pxs(p['line-height']);
      if (p['background-color'] && p['background-color'] !== 'transparent') css['background-color'] = p['background-color'];
      for (const side of ['', '-top', '-bottom', '-left', '-right']) { const b = p[`border${side}`]; if (b && b !== 'none') css[`border${side}`] = b.replace(/([\d.]+)(cm|mm|in|pt)/g, (m) => pxs(m)); }
      for (const side of ['', '-top', '-bottom', '-left', '-right']) { const v = pxs(p[`padding${side}`]); if (v) css[`padding${side}`] = v; }
      return css;
    }
  }
  const cssText = (o) => Object.entries(o).filter(([, v]) => v != null && v !== '').map(([k, v]) => `${k}:${v}`).join(';');

  // ---------------------------------------------------------------- the editor

  class Editor {
    constructor(host, opts) {
      this.host = host;
      this.opts = opts || {};
    }

    async load(bytes) {
      this.bytes = bytes;
      this.zip = await JSZip.loadAsync(bytes);
      const read = async (p) => (this.zip.file(p) ? parseXml(await this.zip.file(p).async('string')) : null);
      this.content = await read('content.xml');
      this.stylesDoc = await read('styles.xml');
      if (!this.content) throw new Error('This file has no content.xml - it is not an OpenDocument text file.');
      this.styles = new Styles(this.stylesDoc, this.content);
      this.urls = [];
      this.newFiles = new Map();
      this.paras = [];     // XML paragraph nodes in document order
      this.keeps = [];     // XML nodes carried over untouched (fields, frames, notes...)
      this.original = [];  // innerHTML of each paragraph as first drawn
      this.autoSeq = 0;
      this.autoText = new Map();
      await this.render();
      return this;
    }

    // Page size and margins from the first master page.
    pageLayout() {
      const st = this.stylesDoc;
      const mp = st && st.getElementsByTagNameNS(NS.style, 'master-page')[0];
      const name = mp && attr(mp, NS.style, 'page-layout-name');
      const pl = name && [...st.getElementsByTagNameNS(NS.style, 'page-layout')].find((x) => attr(x, NS.style, 'name') === name);
      const pp = pl && first(pl, NS.style, 'page-layout-properties');
      return {
        width: px(attr(pp, NS.fo, 'page-width')) || 816, height: px(attr(pp, NS.fo, 'page-height')) || 1056,
        top: px(attr(pp, NS.fo, 'margin-top')) ?? 76, bottom: px(attr(pp, NS.fo, 'margin-bottom')) ?? 76,
        left: px(attr(pp, NS.fo, 'margin-left')) ?? 76, right: px(attr(pp, NS.fo, 'margin-right')) ?? 76,
        master: mp,
      };
    }

    async render() {
      const pl = this.pageLayout();
      const page = document.createElement('div');
      page.className = 'odt-page';
      page.style.cssText = `width:${pl.width}px;min-height:${pl.height}px;padding:${pl.top}px ${pl.right}px ${pl.bottom}px ${pl.left}px`;
      // Default paragraph look (default style + "Standard").
      const base = this.styles.chain('paragraph', 'Standard', true);
      page.style.cssText += ';' + cssText(this.styles.textCss(this.styles.props(base, 'text-properties')));
      const hf = async (kind) => {
        const el = pl.master && first(pl.master, NS.style, kind);
        if (!el || attr(el, NS.style, 'display') === 'false') return null;
        const box = document.createElement('div');
        box.className = `odt-${kind}`;
        box.contentEditable = 'false';
        for (const c of kids(el)) { const x = await this.block(c, true); if (x) box.append(x); }
        return box;
      };
      const header = await hf('header');
      const article = document.createElement('article');
      article.className = 'odt-body';
      article.contentEditable = this.opts.readOnly ? 'false' : 'true';
      article.spellcheck = true;
      const body = this.content.getElementsByTagNameNS(NS.office, 'text')[0];
      if (body) for (const c of kids(body)) { const x = await this.block(c); if (x) article.append(x); }
      if (!article.querySelector('[data-op]') && !this.opts.readOnly) {
        // An empty body still needs one paragraph to type into.
        const p = mk(this.content, NS.text, 'p', { 'text:style-name': 'Standard' });
        body.append(p);
        article.append(await this.block(p));
      }
      const footer = await hf('footer');
      page.append(...[header, article, footer].filter(Boolean));
      this.page = page;
      this.article = article;
      this.host.replaceChildren(page);
      this.original = this.paras.map((_, i) => { const d = article.querySelector(`[data-op="${i}"]`); return d ? d.innerHTML + '|' + d.getAttribute('style') : null; });
      article.addEventListener('input', () => this.opts.onChange && this.opts.onChange());
      article.addEventListener('paste', (e) => {
        // Paste as plain text so foreign HTML never ends up in the document.
        e.preventDefault();
        const t = e.clipboardData.getData('text/plain');
        if (t) document.execCommand('insertText', false, t);
      });
    }

    // ---- block-level XML -> HTML

    async block(n, readOnly) {
      if (n.nodeType !== 1) return null;
      const ns = n.namespaceURI, ln = n.localName;
      if (isPara(n)) return this.paragraph(n, readOnly);
      if (ns === NS.text && ln === 'list') return this.list(n, readOnly, 0);
      if (ns === NS.table && ln === 'table') return this.table(n, readOnly);
      if (ns === NS.text && (ln === 'section' || ln === 'index-body')) {
        const d = document.createElement('div');
        for (const c of kids(n)) { const x = await this.block(c, readOnly); if (x) d.append(x); }
        return d;
      }
      if (ns === NS.text && /^(table-of-content|alphabetical-index|illustration-index|bibliography|user-index|object-index|table-index)$/.test(ln)) {
        // Generated indexes are shown but not edited (the office program regenerates them).
        const d = document.createElement('div');
        d.className = 'odt-index';
        d.contentEditable = 'false';
        const ib = first(n, NS.text, 'index-body');
        if (ib) for (const c of kids(ib)) { const x = await this.block(c, true); if (x) d.append(x); }
        return d;
      }
      if (ns === NS.draw && ln === 'frame') { const s = document.createElement('div'); const f = await this.frame(n); if (f) s.append(f); return s; }
      return null;
    }

    async paragraph(n, readOnly) {
      const S = this.styles;
      const styleName = attr(n, NS.text, 'style-name');
      const isH = n.localName === 'h';
      const chain = S.chain('paragraph', styleName, true);
      const tp = S.props(chain, 'text-properties');
      const pp = S.props(chain, 'paragraph-properties');
      const el = document.createElement(isH ? `h${Math.min(6, +(attr(n, NS.text, 'outline-level') || 1))}` : 'p');
      el.className = 'odt-p';
      el.style.cssText = cssText({ ...S.paraCss(pp), ...S.textCss(tp) });
      if (isH) { if (!tp['font-size']) el.style.fontSize = ''; if (!tp['font-weight']) el.style.fontWeight = ''; }
      const brk = chain.map((s) => attr(first(s, NS.style, 'paragraph-properties'), NS.fo, 'break-before')).filter(Boolean).pop();
      if (brk === 'page') el.classList.add('odt-break');
      if (!readOnly) { el.dataset.op = this.paras.length; this.paras.push(n); }
      await this.inline(n, el, readOnly);
      if (!el.childNodes.length) el.append(document.createElement('br'));
      return el;
    }

    async inline(n, el, readOnly) {
      for (const c of n.childNodes) {
        if (c.nodeType === 3) { el.append(document.createTextNode(c.data.replace(/[\r\n\t ]+/g, ' '))); continue; }
        if (c.nodeType !== 1) continue;
        const ns = c.namespaceURI, ln = c.localName;
        if (ns === NS.text && ln === 'span') {
          const span = document.createElement('span');
          const sn = attr(c, NS.text, 'style-name');
          if (sn) span.dataset.sn = sn;
          span.style.cssText = cssText(this.styles.textCss(this.styles.props(this.styles.chain('text', sn), 'text-properties')));
          span.dataset.os = span.style.cssText;
          await this.inline(c, span, readOnly);
          el.append(span);
        } else if (ns === NS.text && ln === 's') el.append(document.createTextNode(' '.repeat(+(attr(c, NS.text, 'c') || 1))));
        else if (ns === NS.text && ln === 'tab') el.append(document.createTextNode('\t'));
        else if (ns === NS.text && ln === 'line-break') el.append(document.createElement('br'));
        else if (ns === NS.text && ln === 'soft-page-break') continue;
        else if (ns === NS.text && ln === 'a') {
          const a = document.createElement('a');
          a.href = attr(c, NS.xlink, 'href') || '#';
          a.dataset.keepA = this.keeps.length;
          this.keeps.push(c);
          await this.inline(c, a, readOnly);
          el.append(a);
        } else {
          // Anything else is shown read-only and saved exactly as it was.
          const keep = document.createElement('span');
          keep.className = 'odt-keep';
          keep.contentEditable = 'false';
          keep.dataset.keep = this.keeps.length;
          this.keeps.push(c);
          if (ns === NS.draw && ln === 'frame') { const f = await this.frame(c); if (f) keep.append(f); }
          else if (ns === NS.text && ln === 'note') {
            const cit = first(c, NS.text, 'note-citation');
            keep.append(Object.assign(document.createElement('sup'), { textContent: cit ? cit.textContent : '*', title: c.textContent }));
          } else if (ns === NS.office && ln === 'annotation') { keep.classList.add('odt-comment'); keep.title = c.textContent; keep.textContent = '💬'; }
          else keep.textContent = c.textContent; // fields (page number, date, references...) show their current value
          el.append(keep);
        }
      }
    }

    async list(n, readOnly, level, styleName) {
      const sn = attr(n, NS.text, 'style-name') || styleName;
      const ls = sn && this.styles.lists.get(sn);
      const lvl = ls && kids(ls).find((x) => +(attr(x, NS.text, 'level') || 0) === level + 1);
      const numbered = lvl && lvl.localName === 'list-level-style-number';
      const el = document.createElement(numbered ? 'ol' : 'ul');
      el.className = 'odt-list';
      if (lvl) {
        if (numbered) {
          const f = attr(lvl, NS.style, 'num-format');
          el.style.listStyleType = { 1: 'decimal', a: 'lower-alpha', A: 'upper-alpha', i: 'lower-roman', I: 'upper-roman', '': 'none' }[f ?? '1'] || 'decimal';
          const start = attr(lvl, NS.text, 'start-value');
          if (start) el.start = +start;
        } else if (lvl.localName === 'list-level-style-bullet') {
          // Bullets in the Symbol/Wingdings private range (as Word writes them) are shown as their usual glyphs.
          const raw = attr(lvl, NS.text, 'bullet-char') || '•';
          const code = raw.charCodeAt(0);
          const ch = code >= 0xf000 && code <= 0xf0ff ? ({ 0xf0b7: '•', 0xf0a7: '▪', 0xf0a8: '◆', 0xf0d8: '➢', 0xf0fc: '✓', 0xf076: '❖', 0xf06e: '■', 0xf06f: '□', 0xf0a1: '○' }[code] || '•') : raw;
          el.style.listStyleType = `"${ch.replace(/["\\]/g, '')}  "`;
        } else el.style.listStyleType = 'none';
        const lp = first(lvl, NS.style, 'list-level-properties');
        const la = lp && first(lp, NS.style, 'list-level-label-alignment');
        const ml = px(attr(la, NS.fo, 'margin-left')) ?? px(attr(lp, NS.text, 'space-before'));
        if (ml != null) el.style.paddingLeft = `${Math.max(18, ml)}px`;
      }
      if (attr(n, NS.text, 'continue-numbering') === 'true' && this.lastOl && numbered) el.start = this.lastOl.count + 1;
      let count = 0;
      for (const item of kids(n, NS.text)) {
        if (item.localName !== 'list-item' && item.localName !== 'list-header') continue;
        const li = document.createElement('li');
        if (item.localName === 'list-header') li.style.listStyleType = 'none';
        for (const c of kids(item)) {
          if (c.namespaceURI === NS.text && c.localName === 'list') { li.append(await this.list(c, readOnly, level + 1, sn)); continue; }
          const x = await this.block(c, readOnly);
          if (x) li.append(x);
        }
        el.append(li);
        count++;
      }
      if (numbered && level === 0) this.lastOl = { count: (el.start || 1) - 1 + count };
      return el;
    }

    async table(n, readOnly) {
      const S = this.styles;
      const t = document.createElement('table');
      t.className = 'odt-table';
      const tp = S.props(S.chain('table', attr(n, NS.table, 'style-name')), 'table-properties');
      if (tp.width) t.style.width = pxs(tp.width);
      const colgroup = document.createElement('colgroup');
      for (const c of n.getElementsByTagNameNS(NS.table, 'table-column')) {
        if (c.parentNode !== n && c.parentNode.parentNode !== n) continue;
        const cp = S.props(S.chain('table-column', attr(c, NS.table, 'style-name')), 'table-column-properties');
        for (let i = 0; i < +(attr(c, NS.table, 'number-columns-repeated') || 1) && i < 200; i++) {
          const col = document.createElement('col');
          if (cp['column-width']) col.style.width = pxs(cp['column-width']);
          colgroup.append(col);
        }
      }
      t.append(colgroup);
      const rows = [];
      const collect = (el) => { for (const c of kids(el, NS.table)) { if (c.localName === 'table-row') rows.push(c); else if (/^table-(header-rows|rows|row-group)$/.test(c.localName)) collect(c); } };
      collect(n);
      const tbody = document.createElement('tbody');
      for (const r of rows) {
        const tr = document.createElement('tr');
        for (const c of kids(r, NS.table)) {
          if (c.localName === 'covered-table-cell') continue;
          if (c.localName !== 'table-cell') continue;
          const td = document.createElement('td');
          const cp = S.props(S.chain('table-cell', attr(c, NS.table, 'style-name')), 'table-cell-properties');
          td.style.cssText = cssText(S.paraCss(cp));
          if (cp['background-color'] && cp['background-color'] !== 'transparent') td.style.backgroundColor = cp['background-color'];
          const cs = attr(c, NS.table, 'number-columns-spanned'), rs = attr(c, NS.table, 'number-rows-spanned');
          if (cs > 1) td.colSpan = +cs;
          if (rs > 1) td.rowSpan = +rs;
          for (const x of kids(c)) { const b = await this.block(x, readOnly); if (b) td.append(b); }
          tr.append(td);
        }
        tbody.append(tr);
      }
      t.append(tbody);
      return t;
    }

    async frame(n) {
      const img = first(n, NS.draw, 'image');
      const w = pxs(attr(n, NS.svg, 'width')), hh = pxs(attr(n, NS.svg, 'height'));
      if (img) {
        const href = attr(img, NS.xlink, 'href');
        const bytes = href && (this.newFiles.get(href) || (this.zip.file(href) && await this.zip.file(href).async('uint8array')));
        if (!bytes) return null;
        const ext = href.split('.').pop().toLowerCase();
        const url = URL.createObjectURL(new Blob([bytes], { type: ext === 'svg' ? 'image/svg+xml' : ext === 'jpg' ? 'image/jpeg' : `image/${ext}` }));
        this.urls.push(url);
        const el = document.createElement('img');
        el.src = url;
        el.className = 'odt-img';
        if (w) el.style.width = w;
        if (hh) el.style.height = hh;
        return el;
      }
      const tb = first(n, NS.draw, 'text-box');
      if (tb) {
        const d = document.createElement('span');
        d.className = 'odt-textbox';
        if (w) d.style.width = w;
        for (const c of kids(tb)) { const x = await this.block(c, true); if (x) d.append(x); }
        return d;
      }
      return null;
    }

    // ---------------------------------------------------------------- toolbar helpers

    get stage() { return this.page; }
    paragraphStyles() {
      const out = [];
      for (const [k, v] of this.styles.map) {
        if (!k.startsWith('paragraph:') || v.auto) continue;
        const name = attr(v.el, NS.style, 'name');
        const display = attr(v.el, NS.style, 'display-name') || name.replace(/_20_/g, ' ');
        if (/^(Header|Footer|Table_20_Contents|Contents|Index|Frame_20_contents|Footnote|Endnote|Caption|Illustration|Drawing|Figure|Table_20_Heading|Header_20_and_20_Footer|Sender|Addressee|Marginalia)/.test(name)) continue;
        out.push({ id: name, name: display });
      }
      return out;
    }
    selectedParagraphs() {
      const sel = getSelection();
      if (!sel.rangeCount) return [];
      const r = sel.getRangeAt(0);
      return [...this.article.querySelectorAll('.odt-p')].filter((p) => r.intersectsNode(p));
    }
    caretParagraph() { const s = getSelection(); const n = s.anchorNode; const el = n && (n.nodeType === 1 ? n : n.parentElement); return el && el.closest('.odt-p'); }
    styleOf(p) { if (p.dataset.newStyle) return p.dataset.newStyle; const n = this.paras[+p.dataset.op]; const sn = n && attr(n, NS.text, 'style-name'); const st = this.styles.get('paragraph', sn); return st && st.auto ? attr(st.el, NS.style, 'parent-style-name') : sn; }
    setParagraphStyle(ps, id) {
      const S = this.styles;
      const chain = S.chain('paragraph', id, true);
      for (const p of ps) {
        p.dataset.newStyle = id;
        p.style.cssText = cssText({ ...S.paraCss(S.props(chain, 'paragraph-properties')), ...S.textCss(S.props(chain, 'text-properties')) });
        delete p.dataset.align;
      }
      this.opts.onChange && this.opts.onChange();
    }
    align(ps, value) {
      for (const p of ps) { p.style.textAlign = value; p.dataset.align = value; }
      this.opts.onChange && this.opts.onChange();
    }
    async insertImage(bytes, ext, w, h) {
      let n = 1;
      while (this.zip.file(`Pictures/docstudio${n}.${ext}`) || this.newFiles.has(`Pictures/docstudio${n}.${ext}`)) n++;
      const href = `Pictures/docstudio${n}.${ext}`;
      this.newFiles.set(href, new Uint8Array(bytes));
      const pl = this.pageLayout();
      const maxW = pl.width - pl.left - pl.right;
      const s = Math.min(1, maxW / w);
      const cm = (v) => `${(v * s / 37.795).toFixed(3)}cm`;
      const frame = mk(this.content, NS.draw, 'frame', { 'draw:name': `Image${Date.now() % 100000}`, 'text:anchor-type': 'as-char', 'svg:width': cm(w), 'svg:height': cm(h), 'draw:z-index': '0' });
      frame.append(mk(this.content, NS.draw, 'image', { 'xlink:href': href, 'xlink:type': 'simple', 'xlink:show': 'embed', 'xlink:actuate': 'onLoad' }));
      const keep = document.createElement('span');
      keep.className = 'odt-keep';
      keep.contentEditable = 'false';
      keep.dataset.keep = this.keeps.length;
      this.keeps.push(frame);
      keep.append(await this.frame(frame));
      // Insert at the caret, always inside a paragraph (pictures are anchored "as character").
      const sel = getSelection();
      let range = sel.rangeCount ? sel.getRangeAt(0) : null;
      const host = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement);
      const para = host && host.closest && host.closest('.odt-p');
      if (!para || !this.article.contains(para)) {
        const last = [...this.article.querySelectorAll('.odt-p')].pop();
        range = document.createRange();
        range.selectNodeContents(last);
        range.collapse(false);
      }
      range.deleteContents();
      range.insertNode(keep);
      range.setStartAfter(keep);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
      this.opts.onChange && this.opts.onChange();
    }

    // ---------------------------------------------------------------- saving

    async save() {
      this.applyEdits();
      const zip = this.zip;
      zip.file('content.xml', serialize(this.content));
      if (this.newFiles.size) {
        const man = parseXml(await zip.file('META-INF/manifest.xml').async('string'));
        const root = man.documentElement;
        for (const [p, b] of this.newFiles) {
          zip.file(p, b);
          if (![...root.children].some((e) => e.getAttributeNS(NS.manifest, 'full-path') === p)) {
            const e = man.createElementNS(NS.manifest, 'manifest:file-entry');
            e.setAttributeNS(NS.manifest, 'manifest:full-path', p);
            e.setAttributeNS(NS.manifest, 'manifest:media-type', p.endsWith('.png') ? 'image/png' : p.endsWith('.gif') ? 'image/gif' : p.endsWith('.bmp') ? 'image/bmp' : 'image/jpeg');
            root.append(e);
          }
        }
        zip.file('META-INF/manifest.xml', serialize(man));
      }
      // The mimetype entry must stay first and uncompressed.
      const mime = zip.file('mimetype') ? await zip.file('mimetype').async('string') : 'application/vnd.oasis.opendocument.text';
      zip.file('mimetype', mime, { compression: 'STORE' });
      const out = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
      this.newFiles.clear();
      // Start again from the saved file so later saves compare against what is on disk now.
      this.urls.forEach((u) => URL.revokeObjectURL(u));
      const editable = this.article.contentEditable;
      const scroll = this.host.scrollTop;
      await this.load(out);
      this.article.contentEditable = editable;
      this.host.scrollTop = scroll;
      return out;
    }

    applyEdits() {
      const doms = [...this.article.querySelectorAll('.odt-p')].filter((p) => !p.closest('[contenteditable="false"]'));
      // Group the DOM paragraphs by the XML paragraph they came from; unknown blocks follow the one before them.
      const groups = new Map();
      let last = null;
      for (const d of doms) {
        const op = d.dataset.op !== undefined ? +d.dataset.op : last;
        if (op === null) continue;
        if (!groups.has(op)) groups.set(op, []);
        groups.get(op).push(d);
        last = op;
      }
      this.paras.forEach((xp, i) => {
        const list = groups.get(i) || [];
        if (!list.length) { this.removePara(xp); return; }
        const [firstDom, ...extra] = list;
        const sig = firstDom.innerHTML + '|' + firstDom.getAttribute('style');
        if (sig !== this.original[i] || firstDom.dataset.newStyle || firstDom.dataset.align) this.rebuild(xp, firstDom);
        let after = xp;
        for (const d of extra) {
          const np = xp.cloneNode(false);
          for (const a of [...np.attributes]) if (/^(id|xml:id)$/.test(a.name) || a.localName === 'id') np.removeAttributeNode(a);
          this.rebuild(np, d);
          const liFirst = firstDom.closest('li'), liThis = d.closest('li');
          const xli = xp.parentNode && xp.parentNode.localName === 'list-item' ? xp.parentNode : null;
          if (xli && liThis && liThis !== liFirst) {
            // A new bullet: a new list item after the current one.
            const item = xli.cloneNode(false);
            item.append(np);
            const anchor = after.parentNode && after.parentNode.localName === 'list-item' ? after.parentNode : xli;
            anchor.after(item);
            after = np;
          } else {
            after.after(np);
            after = np;
          }
        }
      });
    }

    removePara(xp) {
      const parent = xp.parentNode;
      if (!parent) return;
      parent.removeChild(xp);
      // Keep containers valid: an emptied list item goes (and an emptied list), an emptied table cell keeps one paragraph.
      if (parent.localName === 'list-item' && !kids(parent).length) {
        const list = parent.parentNode;
        list.removeChild(parent);
        if (!kids(list, NS.text).some((c) => c.localName === 'list-item' || c.localName === 'list-header') && list.parentNode) {
          const lp = list.parentNode;
          lp.removeChild(list);
          if (lp.localName === 'list-item' && !kids(lp).length && lp.parentNode) lp.parentNode.removeChild(lp);
        }
      } else if (parent.localName === 'table-cell' && !kids(parent).length) {
        parent.append(mk(this.content, NS.text, 'p', { 'text:style-name': attr(xp, NS.text, 'style-name') || 'Standard' }));
      }
    }

    rebuild(xp, dom) {
      const doc = this.content;
      // Paragraph style: chosen from the list, or the original with a new alignment.
      if (dom.dataset.newStyle) xp.setAttributeNS(NS.text, 'text:style-name', dom.dataset.newStyle);
      if (dom.dataset.align) {
        const base = attr(xp, NS.text, 'style-name') || 'Standard';
        xp.setAttributeNS(NS.text, 'text:style-name', this.autoPara(base, { 'fo:text-align': { left: 'start', right: 'end', center: 'center', justify: 'justify' }[dom.dataset.align] || 'start' }));
      }
      while (xp.firstChild) xp.removeChild(xp.firstChild);
      const out = { root: dom, atStart: true, prevSpace: false };
      this.emitChildren(dom, xp, null, {}, out);
    }

    emitChildren(domParent, xmlParent, spanStyle, fmt, out) {
      for (const c of domParent.childNodes) {
        if (c.nodeType === 3) { this.emitText(c.data, xmlParent, spanStyle, fmt, out); continue; }
        if (c.nodeType !== 1) continue;
        if (c.dataset.keep !== undefined) { xmlParent.append(this.keeps[+c.dataset.keep].cloneNode(true)); out.prevSpace = false; out.atStart = false; continue; }
        if (c.tagName === 'BR') { if (c === domParent.lastChild && domParent === out.root) continue; xmlParent.append(mk(this.content, NS.text, 'line-break')); out.prevSpace = false; continue; }
        if (c.tagName === 'A') {
          const orig = c.dataset.keepA !== undefined ? this.keeps[+c.dataset.keepA] : null;
          const a = orig ? orig.cloneNode(false) : mk(this.content, NS.text, 'a', { 'xlink:type': 'simple' });
          a.setAttributeNS(NS.xlink, 'xlink:href', c.getAttribute('href') || '');
          xmlParent.append(a);
          this.emitChildren(c, a, spanStyle, fmt, out);
          continue;
        }
        // Span: an original span (data-sn) sets the base style; anything else adds formatting.
        let sn = spanStyle, f = fmt;
        if (c.dataset.sn !== undefined || c.dataset.os !== undefined) {
          sn = c.dataset.sn || null;
          f = { ...fmt, ...diffCss(c.dataset.os || '', c.style) };
        } else f = { ...fmt, ...formatOf(c) };
        this.emitChildren(c, xmlParent, sn, f, out);
      }
    }

    emitText(text, xmlParent, spanStyle, fmt, out) {
      if (!text) return;
      const clean = Object.fromEntries(Object.entries(fmt).filter(([, v]) => v !== undefined));
      const style = Object.keys(clean).length ? this.autoTextStyle(spanStyle, clean) : spanStyle;
      let target = xmlParent;
      if (style) {
        const lastEl = xmlParent.lastChild;
        if (lastEl && lastEl.nodeType === 1 && lastEl.namespaceURI === NS.text && lastEl.localName === 'span' && attr(lastEl, NS.text, 'style-name') === style) target = lastEl;
        else { target = mk(this.content, NS.text, 'span', { 'text:style-name': style }); xmlParent.append(target); }
      }
      // Whitespace: ODF collapses runs of spaces, so extra spaces are written as <text:s/>, tabs as <text:tab/>.
      const doc = this.content;
      let buf = '';
      const flush = () => { if (buf) { target.append(doc.createTextNode(buf)); buf = ''; } };
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (ch === '\t') { flush(); target.append(mk(doc, NS.text, 'tab')); out.prevSpace = false; out.atStart = false; continue; }
        if (ch === '\n') { flush(); target.append(mk(doc, NS.text, 'line-break')); out.prevSpace = false; continue; }
        if (ch === ' ' || ch === '\u00a0') {
          let j = i; while (j < text.length && (text[j] === ' ' || text[j] === '\u00a0')) j++;
          let count = j - i;
          if (!out.atStart && !out.prevSpace) { buf += ' '; count--; }
          if (count > 0) { flush(); const s = mk(doc, NS.text, 's'); if (count > 1) s.setAttributeNS(NS.text, 'text:c', String(count)); target.append(s); }
          out.prevSpace = true;
          i = j - 1;
          continue;
        }
        buf += ch;
        out.prevSpace = false; out.atStart = false;
      }
      flush();
    }

    automaticStyles() {
      let as = this.content.getElementsByTagNameNS(NS.office, 'automatic-styles')[0];
      if (!as) { as = mk(this.content, NS.office, 'automatic-styles'); const body = this.content.getElementsByTagNameNS(NS.office, 'body')[0]; body.parentNode.insertBefore(as, body); }
      return as;
    }
    newName(prefix) {
      let name;
      do name = `${prefix}${++this.autoSeq}`; while (this.styles.map.has(`text:${name}`) || this.styles.map.has(`paragraph:${name}`));
      return name;
    }
    ensureFont(name) {
      const doc = this.content;
      let decls = doc.getElementsByTagNameNS(NS.office, 'font-face-decls')[0];
      if (!decls) { decls = mk(doc, NS.office, 'font-face-decls'); doc.documentElement.insertBefore(decls, this.automaticStyles()); }
      if (this.styles.fonts.has(name)) return;
      decls.append(mk(doc, NS.style, 'font-face', { 'style:name': name, 'svg:font-family': name.includes(' ') ? `'${name}'` : name }));
      this.styles.fonts.set(name, name);
    }

    // An automatic text style = the span's original style plus the new formatting.
    autoTextStyle(base, fmt) {
      const key = `${base || ''}|${JSON.stringify(fmt)}`;
      if (this.autoText.has(key)) return this.autoText.get(key);
      const doc = this.content;
      const name = this.newName('DDT');
      const st = mk(doc, NS.style, 'style', { 'style:name': name, 'style:family': 'text' });
      const tp = mk(doc, NS.style, 'text-properties');
      const b = this.styles.get('text', base);
      if (b && b.auto) {
        const parent = attr(b.el, NS.style, 'parent-style-name');
        if (parent) st.setAttributeNS(NS.style, 'style:parent-style-name', parent);
        const otp = first(b.el, NS.style, 'text-properties');
        if (otp) for (const a of otp.attributes) tp.setAttributeNS(a.namespaceURI, a.name, a.value);
      } else if (b) st.setAttributeNS(NS.style, 'style:parent-style-name', base);
      const set = (k, v) => { const [p] = k.split(':'); tp.setAttributeNS(NS[p], k, v); };
      if (fmt.b !== undefined) { set('fo:font-weight', fmt.b ? 'bold' : 'normal'); set('style:font-weight-asian', fmt.b ? 'bold' : 'normal'); set('style:font-weight-complex', fmt.b ? 'bold' : 'normal'); }
      if (fmt.i !== undefined) { set('fo:font-style', fmt.i ? 'italic' : 'normal'); set('style:font-style-asian', fmt.i ? 'italic' : 'normal'); set('style:font-style-complex', fmt.i ? 'italic' : 'normal'); }
      if (fmt.u !== undefined) { set('style:text-underline-style', fmt.u ? 'solid' : 'none'); if (fmt.u) { set('style:text-underline-width', 'auto'); set('style:text-underline-color', 'font-color'); } }
      if (fmt.strike !== undefined) { set('style:text-line-through-style', fmt.strike ? 'solid' : 'none'); if (fmt.strike) set('style:text-line-through-type', 'single'); }
      if (fmt.color) set('fo:color', fmt.color);
      if (fmt.bg) set('fo:background-color', fmt.bg);
      if (fmt.size) { set('fo:font-size', fmt.size); set('style:font-size-asian', fmt.size); set('style:font-size-complex', fmt.size); }
      if (fmt.font) { this.ensureFont(fmt.font); set('style:font-name', fmt.font); }
      if (fmt.pos) set('style:text-position', fmt.pos === 'super' ? 'super 58%' : 'sub 58%');
      st.append(tp);
      this.automaticStyles().append(st);
      this.styles.map.set(`text:${name}`, { el: st, auto: true, doc });
      this.autoText.set(key, name);
      return name;
    }

    autoPara(base, props) {
      const doc = this.content;
      const name = this.newName('DDP');
      const st = mk(doc, NS.style, 'style', { 'style:name': name, 'style:family': 'paragraph' });
      const b = this.styles.get('paragraph', base);
      const pp = mk(doc, NS.style, 'paragraph-properties');
      if (b && b.auto) {
        const parent = attr(b.el, NS.style, 'parent-style-name');
        if (parent) st.setAttributeNS(NS.style, 'style:parent-style-name', parent);
        const lsn = attr(b.el, NS.style, 'list-style-name');
        if (lsn) st.setAttributeNS(NS.style, 'style:list-style-name', lsn);
        const mpn = attr(b.el, NS.style, 'master-page-name');
        if (mpn) st.setAttributeNS(NS.style, 'style:master-page-name', mpn);
        for (const c of kids(b.el)) { if (c.localName === 'paragraph-properties') { for (const a of c.attributes) pp.setAttributeNS(a.namespaceURI, a.name, a.value); for (const x of kids(c)) pp.append(x.cloneNode(true)); } else st.append(c.cloneNode(true)); }
      } else st.setAttributeNS(NS.style, 'style:parent-style-name', base);
      for (const [k, v] of Object.entries(props)) { const [p] = k.split(':'); pp.setAttributeNS(NS[p], k, v); }
      st.insertBefore(pp, st.firstChild);
      this.automaticStyles().append(st);
      this.styles.map.set(`paragraph:${name}`, { el: st, auto: true, doc });
      return name;
    }

    destroy() { (this.urls || []).forEach((u) => URL.revokeObjectURL(u)); this.host.replaceChildren(); }
  }

  // ---------------------------------------------------------------- formatting read from the edited DOM

  function formatOf(el) {
    const f = {};
    const t = el.tagName;
    if (t === 'B' || t === 'STRONG') f.b = true;
    if (t === 'I' || t === 'EM') f.i = true;
    if (t === 'U') f.u = true;
    if (t === 'S' || t === 'STRIKE') f.strike = true;
    if (t === 'SUP') f.pos = 'super';
    if (t === 'SUB') f.pos = 'sub';
    if (t === 'FONT') { if (el.color) f.color = el.color; if (el.face) f.font = el.face; }
    return { ...f, ...cssFormat(el.style) };
  }
  function cssFormat(st) {
    const f = {};
    if (st.fontWeight) f.b = st.fontWeight === 'bold' || +st.fontWeight >= 600;
    if (st.fontStyle) f.i = st.fontStyle === 'italic';
    const deco = st.textDecorationLine || st.textDecoration;
    if (deco) { if (/underline/.test(deco)) f.u = true; if (/line-through/.test(deco)) f.strike = true; if (deco === 'none') { f.u = false; f.strike = false; } }
    if (st.color) f.color = hex(st.color);
    if (st.backgroundColor) f.bg = hex(st.backgroundColor);
    if (st.fontSize && /pt$/.test(st.fontSize)) f.size = st.fontSize;
    else if (st.fontSize && /px$/.test(st.fontSize)) f.size = `${Math.round(parseFloat(st.fontSize) * 0.75 * 2) / 2}pt`;
    if (st.fontFamily) f.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    if (st.verticalAlign === 'super' || st.verticalAlign === 'sub') f.pos = st.verticalAlign;
    return f;
  }
  // Formatting that differs from what an original span was drawn with.
  function diffCss(origText, st) {
    const o = document.createElement('span').style;
    o.cssText = origText;
    const now = cssFormat(st), was = cssFormat(o);
    const out = {};
    for (const k of new Set([...Object.keys(now), ...Object.keys(was)])) if (JSON.stringify(now[k]) !== JSON.stringify(was[k])) out[k] = now[k] === undefined ? undefined : now[k];
    return out;
  }
  function hex(c) {
    const m = /rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(c);
    return m ? '#' + [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('') : c;
  }

  // ---------------------------------------------------------------- a new document

  async function blank() {
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
    const o = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"';
    const now = new Date().toISOString().replace(/\.\d+Z$/, '');
    const heading = (n, size, extra) => `<style:style style:name="Heading_20_${n}" style:display-name="Heading ${n}" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="${n}" style:class="text"><style:paragraph-properties fo:margin-top="${n === 1 ? '0.423cm' : '0.353cm'}" fo:margin-bottom="0.212cm"/><style:text-properties fo:font-size="${size}" fo:font-weight="bold" style:font-size-asian="${size}" style:font-weight-asian="bold" style:font-size-complex="${size}" style:font-weight-complex="bold"${extra || ''}/></style:style>`;
    zip.file('styles.xml', `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${o} office:version="1.3"><office:font-face-decls><style:font-face style:name="Calibri" svg:font-family="Calibri" style:font-family-generic="swiss"/><style:font-face style:name="Calibri Light" svg:font-family="'Calibri Light'" style:font-family-generic="swiss"/></office:font-face-decls><office:styles><style:default-style style:family="paragraph"><style:paragraph-properties fo:hyphenation-ladder-count="no-limit" style:writing-mode="page"/><style:text-properties style:font-name="Calibri" fo:font-size="11pt" fo:language="en" fo:country="US" style:font-size-asian="11pt" style:font-size-complex="11pt"/></style:default-style><style:style style:name="Standard" style:family="paragraph" style:class="text"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.282cm" fo:line-height="108%"/></style:style><style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard" style:class="text"/><style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text"><style:paragraph-properties fo:keep-with-next="always"/><style:text-properties style:font-name="Calibri Light" fo:color="#2f5496"/></style:style>${heading(1, '16pt')}${heading(2, '13pt')}${heading(3, '12pt', ' fo:color="#1f3763"')}<style:style style:name="Title" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="chapter"><style:paragraph-properties fo:margin-bottom="0cm" fo:line-height="100%"/><style:text-properties style:font-name="Calibri Light" fo:font-size="28pt" style:font-size-asian="28pt" style:font-size-complex="28pt"/></style:style><style:style style:name="Subtitle" style:family="paragraph" style:parent-style-name="Standard" style:class="chapter"><style:text-properties fo:color="#5a5a5a" fo:letter-spacing="0.026cm"/></style:style><style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Standard" style:class="html"><style:paragraph-properties fo:margin-left="1.524cm" fo:margin-right="1.524cm" fo:text-align="center"/><style:text-properties fo:font-style="italic" fo:color="#404040"/></style:style><style:style style:name="Strong_20_Emphasis" style:display-name="Strong Emphasis" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style><style:style style:name="Emphasis" style:family="text"><style:text-properties fo:font-style="italic"/></style:style></office:styles><office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21.59cm" fo:page-height="27.94cm" style:print-orientation="portrait" fo:margin-top="2.54cm" fo:margin-bottom="2.54cm" fo:margin-left="2.54cm" fo:margin-right="2.54cm"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="pm1"/></office:master-styles></office:document-styles>`);
    zip.file('content.xml', `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${o} office:version="1.3"><office:font-face-decls/><office:automatic-styles/><office:body><office:text><text:p text:style-name="Standard"/></office:text></office:body></office:document-content>`);
    zip.file('meta.xml', `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta ${o} office:version="1.3"><office:meta><meta:generator>Doc Studio</meta:generator><meta:initial-creator>Doc Studio</meta:initial-creator><meta:creation-date>${now}</meta:creation-date><dc:date>${now}</dc:date></office:meta></office:document-meta>`);
    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>`);
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }


  // ---------------------------------------------------------------- HTML -> a new ODT (for conversions)
  // Headings, paragraphs, bold/italic/underline/strike/code/sub/sup, links, line breaks, bullet and numbered
  // lists (nested), tables, block quotes, preformatted text, rules and pictures (data: URLs).

  async function fromHtml(html) {
    const zip = await JSZip.loadAsync(await blank());
    const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
    const x = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const auto = [];
    const textStyles = new Map();
    const pictures = [];
    let tableNo = 0;
    const textStyle = (f) => {
      const key = JSON.stringify(f);
      if (textStyles.has(key)) return textStyles.get(key);
      const name = `T${textStyles.size + 1}`;
      const a = [];
      if (f.b) a.push('fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"');
      if (f.i) a.push('fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"');
      if (f.u) a.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
      if (f.s) a.push('style:text-line-through-style="solid" style:text-line-through-type="single"');
      if (f.code) a.push('style:font-name="Consolas" fo:background-color="#f2f2f2"');
      if (f.sup) a.push('style:text-position="super 58%"');
      if (f.sub) a.push('style:text-position="sub 58%"');
      if (f.link) a.push('fo:color="#0563c1" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
      auto.push(`<style:style style:name="${name}" style:family="text"><style:text-properties ${a.join(' ')}/></style:style>`);
      textStyles.set(key, name);
      return name;
    };
    const inline = async (node, f) => {
      let out = '';
      for (const c of node.childNodes) {
        if (c.nodeType === 3) {
          const t = c.data.replace(/\s+/g, ' ');
          if (!t) continue;
          const body = x(t).replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`);
          const has = Object.values(f).some(Boolean);
          out += has ? `<text:span text:style-name="${textStyle(f)}">${body}</text:span>` : body;
          continue;
        }
        if (c.nodeType !== 1) continue;
        const t = c.tagName;
        if (t === 'BR') { out += '<text:line-break/>'; continue; }
        if (t === 'IMG') { out += await picture(c); continue; }
        if (t === 'A' && c.getAttribute('href')) { out += `<text:a xlink:type="simple" xlink:href="${x(c.getAttribute('href'))}">${await inline(c, { ...f, link: true })}</text:a>`; continue; }
        const g = { ...f };
        if (t === 'B' || t === 'STRONG') g.b = true;
        if (t === 'I' || t === 'EM') g.i = true;
        if (t === 'U' || t === 'INS') g.u = true;
        if (t === 'S' || t === 'DEL' || t === 'STRIKE') g.s = true;
        if (t === 'CODE' || t === 'KBD') g.code = true;
        if (t === 'SUP') g.sup = true;
        if (t === 'SUB') g.sub = true;
        out += await inline(c, g);
      }
      return out;
    };
    const picture = async (img) => {
      const m = /^data:image\/(png|jpe?g|gif|bmp|svg\+xml|webp);base64,(.*)$/.exec(img.getAttribute('src') || '');
      if (!m) return '';
      const ext = m[1] === 'svg+xml' ? 'svg' : m[1].replace('jpeg', 'jpg');
      const name = `Pictures/image${pictures.length + 1}.${ext}`;
      let w = +img.getAttribute('width') || 0, hgt = +img.getAttribute('height') || 0;
      if (!w || !hgt) { try { const i = new Image(); i.src = img.getAttribute('src'); await i.decode(); w = i.naturalWidth; hgt = i.naturalHeight; } catch { w = 300; hgt = 200; } }
      const s = Math.min(1, 624 / w); // fit the text width (6.5in at 96 dpi)
      pictures.push({ name, data: m[2], type: `image/${m[1]}` });
      return `<draw:frame draw:name="Image${pictures.length}" text:anchor-type="as-char" svg:width="${(w * s / 37.795).toFixed(3)}cm" svg:height="${(hgt * s / 37.795).toFixed(3)}cm" draw:z-index="0"><draw:image xlink:href="${name}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
    };
    const list = async (el, depth) => {
      let out = `<text:list${depth === 0 ? ` text:style-name="${el.tagName === 'OL' ? 'LNum' : 'LBul'}"` : ''}>`;
      for (const li of el.children) {
        if (li.tagName !== 'LI') continue;
        out += '<text:list-item>';
        const own = document.createElement('div');
        const nested = [];
        for (const c of [...li.childNodes]) { if (c.nodeType === 1 && (c.tagName === 'UL' || c.tagName === 'OL')) nested.push(c); else own.append(c.cloneNode(true)); }
        // Task-list checkboxes become ☐ / ☑.
        own.querySelectorAll('input[type="checkbox"]').forEach((cb) => cb.replaceWith(cb.checked || cb.hasAttribute('checked') ? '☑ ' : '☐ '));
        const blocks = [...own.children].filter((c) => c.tagName === 'P');
        if (blocks.length) for (const p of blocks) out += `<text:p text:style-name="ListP">${await inline(p, {})}</text:p>`;
        else out += `<text:p text:style-name="ListP">${await inline(own, {})}</text:p>`;
        for (const n of nested) out += await list(n, depth + 1);
        out += '</text:list-item>';
      }
      return `${out}</text:list>`;
    };
    const table = async (el) => {
      const rows = [...el.querySelectorAll(':scope > tr, :scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr')];
      const cols = Math.max(1, ...rows.map((r) => [...r.children].reduce((n, c) => n + (+c.getAttribute('colspan') || 1), 0)));
      const name = `Table${++tableNo}`;
      // Explicit column widths (Word needs them to lay out merged cells).
      auto.push(`<style:style style:name="${name}C" style:family="table-column"><style:table-column-properties style:column-width="${(16.51 / cols).toFixed(3)}cm"/></style:style>`);
      let out = `<table:table table:name="${name}" table:style-name="TblW">${`<table:table-column table:style-name="${name}C"/>`.repeat(cols)}`;
      for (const r of rows) {
        out += '<table:table-row>';
        let n = 0;
        for (const c of r.children) {
          const span = +c.getAttribute('colspan') || 1;
          const head = c.tagName === 'TH';
          out += `<table:table-cell table:style-name="Cell"${span > 1 ? ` table:number-columns-spanned="${span}"` : ''}><text:p text:style-name="${head ? 'TableHead' : 'Standard'}">${await inline(c, head ? { b: true } : {})}</text:p></table:table-cell>`;
          for (let k = 1; k < span; k++) out += '<table:covered-table-cell/>';
          n += span;
        }
        for (; n < cols; n++) out += '<table:table-cell table:style-name="Cell"><text:p/></table:table-cell>';
        out += '</table:table-row>';
      }
      return `${out}</table:table>`;
    };
    const blocks = async (parent) => {
      let out = '';
      let loose = document.createElement('div');
      const flushLoose = async () => { if (loose.textContent.trim() || loose.querySelector('img')) out += `<text:p text:style-name="Standard">${await inline(loose, {})}</text:p>`; loose = document.createElement('div'); };
      for (const c of [...parent.childNodes]) {
        const t = c.nodeType === 1 ? c.tagName : '';
        if (!/^(H[1-6]|P|UL|OL|TABLE|PRE|BLOCKQUOTE|HR|DIV|SECTION|ARTICLE|FIGURE|HEADER|FOOTER|MAIN|NAV)$/.test(t)) { loose.append(c.cloneNode(true)); continue; }
        await flushLoose();
        if (/^H[1-6]$/.test(t)) { const lvl = +t[1]; out += `<text:h text:style-name="Heading_20_${Math.min(3, lvl)}" text:outline-level="${lvl}">${await inline(c, {})}</text:h>`; }
        else if (t === 'P') out += `<text:p text:style-name="Standard">${await inline(c, {})}</text:p>`;
        else if (t === 'UL' || t === 'OL') out += await list(c, 0);
        else if (t === 'TABLE') out += await table(c);
        else if (t === 'PRE') for (const line of c.textContent.replace(/\n$/, '').split('\n')) out += `<text:p text:style-name="Code">${x(line).replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`).replace(/^ /, '<text:s/>')}</text:p>`;
        else if (t === 'BLOCKQUOTE') out += (await blocks(c)).replace(/text:style-name="Standard"/g, 'text:style-name="Quotations"');
        else if (t === 'HR') out += '<text:p text:style-name="Rule"/>';
        else out += await blocks(c);
      }
      await flushLoose();
      return out;
    };
    const body = await blocks(doc.body) || '<text:p text:style-name="Standard"/>';
    auto.push(
      '<style:style style:name="Code" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:background-color="#f4f4f4"/><style:text-properties style:font-name="Consolas" fo:font-size="10pt"/></style:style>',
      '<style:style style:name="Rule" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:border-bottom="0.5pt solid #999999" fo:padding="0cm"/></style:style>',
      '<style:style style:name="TableHead" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-bottom="0cm"/></style:style>',
      '<style:style style:name="TblW" style:family="table"><style:table-properties style:width="16.51cm" table:align="margins"/></style:style>',
      '<style:style style:name="Cell" style:family="table-cell"><style:table-cell-properties fo:padding="0.1cm" fo:border="0.5pt solid #999999"/></style:style>',
      '<text:list-style style:name="LBul">' + [1, 2, 3, 4, 5].map((l) => `<text:list-level-style-bullet text:level="${l}" text:bullet-char="${['•', '◦', '▪', '•', '◦'][l - 1]}"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${(0.635 * l + 0.635).toFixed(3)}cm" fo:text-indent="-0.635cm" fo:margin-left="${(0.635 * l + 0.635).toFixed(3)}cm"/></style:list-level-properties></text:list-level-style-bullet>`).join('') + '</text:list-style>',
      '<text:list-style style:name="LNum">' + [1, 2, 3, 4, 5].map((l) => `<text:list-level-style-number text:level="${l}" style:num-suffix="." style:num-format="${['1', 'a', 'i', '1', 'a'][l - 1]}"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${(0.635 * l + 0.635).toFixed(3)}cm" fo:text-indent="-0.635cm" fo:margin-left="${(0.635 * l + 0.635).toFixed(3)}cm"/></style:list-level-properties></text:list-level-style-number>`).join('') + '</text:list-style>',
      '<style:style style:name="ListP" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-bottom="0.1cm"/></style:style>',
    );
    let content = await zip.file('content.xml').async('string');
    content = content
      .replace('<office:font-face-decls/>', '<office:font-face-decls><style:font-face style:name="Consolas" svg:font-family="Consolas" style:font-family-generic="modern" style:font-pitch="fixed"/></office:font-face-decls>')
      .replace('<office:automatic-styles/>', `<office:automatic-styles>${auto.join('')}</office:automatic-styles>`)
      .replace('<office:text><text:p text:style-name="Standard"/></office:text>', `<office:text>${body}</office:text>`);
    zip.file('content.xml', content);
    if (pictures.length) {
      let man = await zip.file('META-INF/manifest.xml').async('string');
      for (const p of pictures) {
        zip.file(p.name, p.data, { base64: true });
        man = man.replace('</manifest:manifest>', `<manifest:file-entry manifest:full-path="${p.name}" manifest:media-type="${p.type}"/></manifest:manifest>`);
      }
      zip.file('META-INF/manifest.xml', man);
    }
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }

  return { Editor, blank, fromHtml, NS };
})();
