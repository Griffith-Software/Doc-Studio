'use strict';
// Formatting-preserving Word editor.
//
// The document is shown with docx-preview (the same renderer as View mode) and edited in place. On save, each
// edited XML part (the body, headers, footers, footnotes, endnotes) is patched: paragraphs whose rendering didn't
// change are kept byte-for-byte, edited paragraphs are rebuilt from their original paragraph/run properties plus
// only the formatting the user changed, and everything else in the package is untouched.
//
// Mapping rendered HTML back to XML: before rendering, marked copies of the parts get a bookmark at the start of
// every paragraph and an invisible private-use text tag at the start of every run. After rendering, the tags are
// read into data-dd-p ("part:index") / data-dd-r (global run id) attributes and removed.

const DocxEditor = (() => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XMLNS_XML = 'http://www.w3.org/XML/1998/namespace';
  const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
  const REL_NUMBERING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';
  const TAG_OPEN = '';
  const TAG_CLOSE = '';
  const TAG_RE = /([-]+)/g;
  const PART_BASE = 1000000; // run ids are partIndex * PART_BASE + local index
  const WRAPPERS = new Set(['hyperlink', 'ins', 'del', 'smartTag', 'customXml', 'sdt', 'moveFrom', 'moveTo', 'dir', 'bdo']);
  const SIMPLE_RUN_CHILDREN = new Set(['rPr', 't', 'tab', 'br', 'cr', 'lastRenderedPageBreak']);
  const EXCLUDE_ANCESTORS = new Set(['txbxContent', 'AlternateContent', 'drawing', 'pict', 'object', 'Fallback', 'Choice']);
  const PART_TYPES = { header: 'header', footer: 'footer', footnotes: 'footnotes', endnotes: 'endnotes' };

  // Element order inside w:rPr and w:pPr (Word rejects out-of-order children).
  const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
  const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

  const isW = (n, name) => n && n.nodeType === 1 && n.namespaceURI === W && (!name || n.localName === name);
  const kids = (n) => [...n.childNodes].filter((c) => c.nodeType === 1);
  const child = (n, name) => kids(n).find((c) => isW(c, name)) || null;
  const wEl = (doc, name, attrs) => {
    const el = doc.createElementNS(W, `w:${name}`);
    for (const [k, v] of Object.entries(attrs || {})) el.setAttributeNS(W, `w:${k}`, String(v));
    return el;
  };
  const encodeId = (n) => TAG_OPEN + [...n.toString(10)].map((d) => String.fromCharCode(0xe010 + +d)).join('') + TAG_CLOSE;
  const decodeId = (s) => +[...s].map((c) => c.charCodeAt(0) - 0xe010).join('');
  const parseXml = (text) => new DOMParser().parseFromString(text, 'application/xml');
  const serialize = (doc) => {
    const s = new XMLSerializer().serializeToString(doc);
    return s.startsWith('<?xml') ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
  };
  const partOf = (key) => +String(key).split(':')[0];

  function setOrdered(parent, el, order) {
    const old = child(parent, el.localName);
    if (old) { parent.replaceChild(el, old); return; }
    const rank = order.indexOf(el.localName);
    const after = kids(parent).find((c) => order.indexOf(c.localName) > rank);
    parent.insertBefore(el, after || null);
  }
  function removeChild(parent, name) { const c = child(parent, name); if (c) parent.removeChild(c); }

  // ---------- indexing: every paragraph of a part, and every run in it, in document order

  function partRoot(doc, kind) {
    return kind === 'body' ? doc.getElementsByTagNameNS(W, 'body')[0] : doc.documentElement;
  }

  function partParagraphs(doc, kind) {
    const root = partRoot(doc, kind);
    return [...root.getElementsByTagNameNS(W, 'p')].filter((p) => {
      for (let a = p.parentNode; a && a !== root; a = a.parentNode) {
        if (EXCLUDE_ANCESTORS.has(a.localName)) return false;
        // Separator "notes" are Word's own lines between text and notes; leave them alone.
        if ((a.localName === 'footnote' || a.localName === 'endnote') && a.getAttributeNS(W, 'type')) return false;
      }
      return true;
    });
  }

  // Flatten a paragraph into runs and other nodes ("leaves"), each with the wrapper elements around it.
  function flatten(p) {
    const items = [];
    const rec = (node, chain) => {
      for (const c of kids(node)) {
        if (node === p && isW(c, 'pPr')) continue;
        if (isW(c, 'r')) items.push({ type: 'run', node: c, chain });
        else if (c.namespaceURI === W && WRAPPERS.has(c.localName)) {
          if (c.localName === 'sdt') { const content = child(c, 'sdtContent'); if (content) rec(content, [...chain, c]); else items.push({ type: 'leaf', node: c, chain }); }
          else rec(c, [...chain, c]);
        } else items.push({ type: 'leaf', node: c, chain });
      }
    };
    rec(p, []);
    items.forEach((it, i) => { it.pos = i; });
    return items;
  }

  function index(doc, kind, pi) {
    const paras = partParagraphs(doc, kind);
    const runs = [];
    const pinfo = paras.map((p, i) => {
      const items = flatten(p);
      items.forEach((it) => {
        if (it.type !== 'run') return;
        it.id = pi * PART_BASE + runs.length;
        it.pkey = `${pi}:${i}`;
        it.simple = kids(it.node).every((c) => isW(c) && SIMPLE_RUN_CHILDREN.has(c.localName) && !(c.localName === 'br' && c.getAttributeNS(W, 'type') && c.getAttributeNS(W, 'type') !== 'textWrapping'));
        runs.push(it);
      });
      return { node: p, items };
    });
    let tables = [];
    if (kind === 'body') {
      const body = partRoot(doc, kind);
      tables = [...body.getElementsByTagNameNS(W, 'tbl')].filter((t) => {
        for (let a = t.parentNode; a && a !== body; a = a.parentNode) if (EXCLUDE_ANCESTORS.has(a.localName)) return false;
        return true;
      });
    }
    return { paras, pinfo, runs, tables };
  }

  // ---------- styles.xml: paragraph styles for the style picker

  function readStyles(stylesXml) {
    const out = { paragraph: [], byId: {}, defaultPara: null };
    if (!stylesXml) return out;
    for (const s of stylesXml.getElementsByTagNameNS(W, 'style')) {
      if (s.getAttributeNS(W, 'type') !== 'paragraph') continue;
      const id = s.getAttributeNS(W, 'styleId');
      const nameEl = child(s, 'name');
      const info = {
        id,
        name: nameEl ? nameEl.getAttributeNS(W, 'val') : id,
        next: child(s, 'next') ? child(s, 'next').getAttributeNS(W, 'val') : null,
        quick: !!child(s, 'qFormat'),
        hidden: !!child(s, 'semiHidden') || !!child(s, 'hidden'),
      };
      if (s.getAttributeNS(W, 'default') === '1' || s.getAttributeNS(W, 'default') === 'true') out.defaultPara = id;
      out.byId[id] = info;
      out.paragraph.push(info);
    }
    return out;
  }

  // ---------- reading the user's formatting off the DOM

  function inlineStyle(el) {
    const s = el.style;
    return {
      fontWeight: s.fontWeight, fontStyle: s.fontStyle, textDecoration: s.textDecorationLine || s.textDecoration,
      color: s.color, backgroundColor: s.backgroundColor, fontSize: s.fontSize, fontFamily: s.fontFamily, verticalAlign: s.verticalAlign,
    };
  }

  // Translate CSS declarations into run-property intents. Empty string = "back to the style's default".
  function cssToFormat(st, out, onlyChangedFrom) {
    const take = (key, prop, value) => {
      if (out[key] !== undefined) return;
      if (onlyChangedFrom && (onlyChangedFrom[prop] || '') === (st[prop] || '')) return;
      out[key] = value;
    };
    if (st.fontWeight || onlyChangedFrom) {
      const w = st.fontWeight;
      take('bold', 'fontWeight', w ? (w === 'bold' || w === 'bolder' || +w >= 600) : '');
    }
    if (st.fontStyle || onlyChangedFrom) take('italic', 'fontStyle', st.fontStyle ? st.fontStyle === 'italic' || st.fontStyle === 'oblique' : '');
    if (st.textDecoration || onlyChangedFrom) {
      const d = st.textDecoration || '';
      if (!onlyChangedFrom || (onlyChangedFrom.textDecoration || '') !== d) {
        if (out.underline === undefined) out.underline = d ? d.includes('underline') : '';
        if (out.strike === undefined) out.strike = d ? d.includes('line-through') : '';
      }
    }
    if (st.color || onlyChangedFrom) take('color', 'color', st.color ? cssColorToHex(st.color) : '');
    if (st.backgroundColor || onlyChangedFrom) take('highlight', 'backgroundColor', st.backgroundColor ? cssColorToHex(st.backgroundColor) : '');
    if (st.fontSize || onlyChangedFrom) take('size', 'fontSize', st.fontSize ? cssSizeToHalfPoints(st.fontSize) : '');
    if (st.fontFamily || onlyChangedFrom) take('font', 'fontFamily', st.fontFamily ? st.fontFamily.split(',')[0].replace(/["']/g, '').trim() : '');
    if (st.verticalAlign || onlyChangedFrom) take('vert', 'verticalAlign', st.verticalAlign === 'super' ? 'superscript' : st.verticalAlign === 'sub' ? 'subscript' : st.verticalAlign ? 'baseline' : '');
  }

  function cssColorToHex(c) {
    if (!c || c === 'transparent' || c === 'initial' || c === 'inherit') return 'transparent';
    const m = c.match(/rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/);
    if (m) {
      if (m[4] !== undefined && +m[4] === 0) return 'transparent';
      return [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase();
    }
    if (/^#[0-9a-f]{6}$/i.test(c)) return c.slice(1).toUpperCase();
    if (/^#[0-9a-f]{3}$/i.test(c)) return c.slice(1).split('').map((x) => x + x).join('').toUpperCase();
    const probe = document.createElement('span');
    probe.style.color = c;
    document.body.append(probe);
    const rgb = getComputedStyle(probe).color;
    probe.remove();
    return rgb.startsWith('rgb') ? cssColorToHex(rgb) : '000000';
  }

  function cssSizeToHalfPoints(s) {
    const named = { 'xx-small': 7.5, 'x-small': 10, small: 13, medium: 16, large: 18, 'x-large': 24, 'xx-large': 32, 'xxx-large': 48 };
    let px;
    if (named[s]) px = named[s];
    else if (s.endsWith('pt')) return Math.round(parseFloat(s) * 2);
    else if (s.endsWith('px')) px = parseFloat(s);
    else return '';
    return Math.max(2, Math.round(px * 0.75 * 2));
  }

  function tagToFormat(el, out) {
    const t = el.tagName;
    const set = (k, v) => { if (out[k] === undefined) out[k] = v; };
    if (t === 'B' || t === 'STRONG') set('bold', true);
    if (t === 'I' || t === 'EM') set('italic', true);
    if (t === 'U') set('underline', true);
    if (t === 'S' || t === 'STRIKE' || t === 'DEL') set('strike', true);
    if (t === 'SUP') set('vert', 'superscript');
    if (t === 'SUB') set('vert', 'subscript');
    if (t === 'FONT') {
      if (el.getAttribute('color')) set('color', cssColorToHex(el.getAttribute('color')));
      if (el.getAttribute('face')) set('font', el.getAttribute('face').split(',')[0].trim());
    }
  }

  function applyFormat(doc, rPr, fmt) {
    const onOff = (name, v) => {
      if (v === '') { removeChild(rPr, name); return; }
      setOrdered(rPr, v ? wEl(doc, name) : wEl(doc, name, { val: '0' }), RPR_ORDER);
    };
    if (fmt.bold !== undefined) { onOff('b', fmt.bold); onOff('bCs', fmt.bold); }
    if (fmt.italic !== undefined) { onOff('i', fmt.italic); onOff('iCs', fmt.italic); }
    if (fmt.strike !== undefined) onOff('strike', fmt.strike);
    if (fmt.underline !== undefined) {
      if (fmt.underline === '') removeChild(rPr, 'u');
      else setOrdered(rPr, wEl(doc, 'u', { val: fmt.underline ? 'single' : 'none' }), RPR_ORDER);
    }
    if (fmt.color !== undefined) {
      if (fmt.color === '' || fmt.color === 'transparent') removeChild(rPr, 'color');
      else setOrdered(rPr, wEl(doc, 'color', { val: fmt.color }), RPR_ORDER);
    }
    if (fmt.highlight !== undefined) {
      removeChild(rPr, 'highlight');
      if (fmt.highlight === '' || fmt.highlight === 'transparent') removeChild(rPr, 'shd');
      else setOrdered(rPr, wEl(doc, 'shd', { val: 'clear', color: 'auto', fill: fmt.highlight }), RPR_ORDER);
    }
    if (fmt.size !== undefined) {
      if (fmt.size === '') { removeChild(rPr, 'sz'); removeChild(rPr, 'szCs'); }
      else { setOrdered(rPr, wEl(doc, 'sz', { val: fmt.size }), RPR_ORDER); setOrdered(rPr, wEl(doc, 'szCs', { val: fmt.size }), RPR_ORDER); }
    }
    if (fmt.font !== undefined) {
      if (fmt.font === '') removeChild(rPr, 'rFonts');
      else setOrdered(rPr, wEl(doc, 'rFonts', { ascii: fmt.font, hAnsi: fmt.font, cs: fmt.font, eastAsia: fmt.font }), RPR_ORDER);
    }
    if (fmt.vert !== undefined) {
      if (fmt.vert === '' || fmt.vert === 'baseline') removeChild(rPr, 'vertAlign');
      else setOrdered(rPr, wEl(doc, 'vertAlign', { val: fmt.vert }), RPR_ORDER);
    }
  }

  // Pasted HTML keeps simple formatting (bold, italic, colours, fonts, links, paragraphs) and nothing else.
  const PASTE_CSS = ['color', 'background-color', 'font-weight', 'font-style', 'text-decoration', 'text-decoration-line', 'font-size', 'font-family', 'vertical-align'];
  function cleanPastedHtml(html) {
    const clean = DOMPurify.sanitize(html, {
      ALLOWED_TAGS: ['b', 'strong', 'i', 'em', 'u', 's', 'strike', 'del', 'sub', 'sup', 'span', 'font', 'a', 'p', 'br', 'div', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'td', 'th', 'tr', 'table', 'tbody', 'thead', 'ul', 'ol', 'blockquote', 'pre', 'code'],
      ALLOWED_ATTR: ['style', 'href', 'color', 'face'],
      ALLOW_DATA_ATTR: false,
      RETURN_DOM: true,
    });
    for (const el of clean.querySelectorAll('[style]')) {
      const keep = PASTE_CSS.map((p) => (el.style.getPropertyValue(p) ? `${p}:${el.style.getPropertyValue(p)}` : '')).filter(Boolean).join(';');
      if (keep) el.setAttribute('style', keep); else el.removeAttribute('style');
    }
    for (const a of clean.querySelectorAll('a[href]')) if (!/^(https?:|mailto:)/i.test(a.getAttribute('href'))) a.removeAttribute('href');
    // Lists, headings and table cells become plain paragraphs (their content keeps its inline formatting).
    for (const el of [...clean.querySelectorAll('li, h1, h2, h3, h4, h5, h6, blockquote, pre, div, td, th')]) {
      const p = document.createElement('p');
      const isHead = /^H\d$/.test(el.tagName);
      if (el.tagName === 'LI') p.append(document.createTextNode(el.parentElement && el.parentElement.tagName === 'OL' ? `${[...el.parentElement.children].indexOf(el) + 1}. ` : '• '));
      while (el.firstChild) p.append(el.firstChild);
      if (isHead) { const b = document.createElement('b'); while (p.firstChild) b.append(p.firstChild); p.append(b); }
      el.replaceWith(p);
    }
    for (const el of [...clean.querySelectorAll('ul, ol, table, tbody, thead, tr')]) { while (el.firstChild) el.before(el.firstChild); el.remove(); }
    for (const p of [...clean.querySelectorAll('p p')]) { while (p.firstChild) p.before(p.firstChild); p.remove(); }
    return clean.innerHTML;
  }

  // ======================================================================

  class Editor {
    constructor(host, opts) {
      this.host = host;          // element the editable document is rendered into
      this.opts = opts || {};    // { onChange }
      this.history = [];         // package snapshots before structural edits (lists, tables, pictures...)
      this.future = [];
      this.nativeDepth = 0;      // edits the browser can undo itself since the last re-render
    }

    // Load (or reload) from package bytes. `restore` = { scrollTop, caretPara } to put the user back where they were.
    async load(bytes, restore) {
      this.bytes = bytes;
      this.zip = await JSZip.loadAsync(bytes);
      this.docPath = await mainDocumentPath(this.zip);
      this.relsPath = this.docPath.replace(/([^/]+)$/, '_rels/$1.rels');
      const stylesPath = this.docPath.replace(/[^/]+$/, 'styles.xml');
      this.styles = readStyles(this.zip.file(stylesPath) ? parseXml(await this.zip.file(stylesPath).async('string')) : null);

      // Parts: the body first, then headers, footers and notes the document links to.
      this.parts = [{ kind: 'body', path: this.docPath }];
      const relsFile = this.zip.file(this.relsPath);
      if (relsFile) {
        const rels = parseXml(await relsFile.async('string'));
        for (const r of rels.getElementsByTagNameNS(PKG_REL, 'Relationship')) {
          const type = (r.getAttribute('Type') || '').split('/').pop();
          if (!PART_TYPES[type] || r.getAttribute('TargetMode') === 'External') continue;
          const path = resolve(this.docPath, r.getAttribute('Target'));
          if (this.zip.file(path) && !this.parts.some((p) => p.path === path)) this.parts.push({ kind: PART_TYPES[type], path });
        }
      }
      for (const [pi, part] of this.parts.entries()) {
        part.pi = pi;
        part.doc = parseXml(await this.zip.file(part.path).async('string'));
        if (part.doc.getElementsByTagName('parsererror').length) throw new Error(`The document part ${part.path} could not be read.`);
        part.relsPath = part.path.replace(/([^/]+)$/, '_rels/$1.rels');
        part.idx = index(part.doc, part.kind, pi);
      }
      this.doc = this.parts[0].doc;
      this.idx = this.parts[0].idx;

      // Marked copies for rendering.
      const mzip = await JSZip.loadAsync(bytes);
      for (const part of this.parts) {
        const marked = part.doc.cloneNode(true);
        const midx = index(marked, part.kind, part.pi);
        midx.pinfo.forEach((pinf, i) => {
          const bm = wEl(marked, 'bookmarkStart', { id: 900000000 + i, name: `__ddp${part.pi}_${i}` });
          const pPr = child(pinf.node, 'pPr');
          pinf.node.insertBefore(bm, pPr ? pPr.nextSibling : pinf.node.firstChild);
        });
        midx.runs.forEach((r) => {
          const t = wEl(marked, 't');
          t.textContent = encodeId(r.id);
          const rPr = child(r.node, 'rPr');
          r.node.insertBefore(t, rPr ? rPr.nextSibling : r.node.firstChild);
        });
        mzip.file(part.path, serialize(marked));
      }
      const markedBytes = await mzip.generateAsync({ type: 'arraybuffer', compression: 'STORE' });

      const stage = document.createElement('div');
      stage.className = 'docx-view docx-edit-view';
      stage.style.visibility = 'hidden';
      this.host.replaceChildren(stage);
      await docx.renderAsync(markedBytes, stage, null, {
        className: 'docx', inWrapper: true, breakPages: false, ignoreLastRenderedPageBreak: true,
        renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true,
        renderChanges: true, renderComments: true, useBase64URL: true, experimental: true,
      });
      this.stage = stage;
      this.inheritRunStyles();
      this.tagRendering();
      stage.style.visibility = '';
      this.wire();
      this.nativeDepth = 0;
      this.drawPageMarkers();
      if (this.opts.onLoad) this.opts.onLoad();
      if (restore) {
        this.host.parentElement.scrollTop = restore.scrollTop || 0;
        if (restore.caretPara != null) {
          const p = stage.querySelector(`p[data-dd-p="${restore.caretPara}"]`);
          if (p) { const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); p.closest('[contenteditable="true"]')?.focus(); getSelection().removeAllRanges(); getSelection().addRange(r); }
        }
      }
    }

    runAt(id) {
      const part = this.parts[Math.floor(id / PART_BASE)];
      return part ? part.idx.runs[id % PART_BASE] : undefined;
    }

    // docx-preview styles text through "<selector> span" rules, so text typed straight into a paragraph (an empty
    // paragraph has no spans) would get no document formatting. Copy the font rules up to the paragraph/document
    // level so such text inherits them; real spans still match their own, more specific rules.
    inheritRunStyles() {
      const props = ['font-family', 'font-size', 'color', 'font-weight', 'font-style', 'text-decoration', 'letter-spacing', 'text-transform'];
      const extra = [];
      for (const st of this.stage.querySelectorAll('style')) {
        let rules;
        try { rules = st.sheet.cssRules; } catch { continue; }
        for (const rule of rules) {
          if (!rule.selectorText) continue;
          const sels = rule.selectorText.split(',').map((x) => x.trim()).filter((x) => / span$/.test(x)).map((x) => x.replace(/ span$/, ''));
          if (!sels.length) continue;
          const decl = props.map((p) => (rule.style.getPropertyValue(p) ? `${p}: ${rule.style.getPropertyValue(p)};` : '')).join('');
          if (decl) extra.push(`${sels.join(', ')} { ${decl} }`);
        }
      }
      if (extra.length) this.stage.append(h('style', { text: extra.join('\n') }));
    }

    tagRendering() {
      const stage = this.stage;
      this.renderedParas = new Set();
      this.renderedRuns = new Set();
      stage.querySelectorAll('span[id^="__ddp"]').forEach((m) => {
        const mm = /^__ddp(\d+)_(\d+)$/.exec(m.id);
        const p = m.closest('p');
        if (mm && p && p.dataset.ddP === undefined) { p.dataset.ddP = `${mm[1]}:${mm[2]}`; this.renderedParas.add(p.dataset.ddP); }
        m.remove();
      });
      const walker = document.createTreeWalker(stage, NodeFilter.SHOW_TEXT);
      const tagged = [];
      for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.data.includes(TAG_OPEN)) tagged.push(n);
      for (const n of tagged) {
        let el = n.parentElement;
        if (el && (el.tagName === 'SUP' || el.tagName === 'SUB') && el.parentElement && el.parentElement.tagName === 'SPAN') el = el.parentElement;
        n.data.replace(TAG_RE, (_, digits) => {
          const id = decodeId(digits);
          if (el && el.dataset.ddR === undefined) { el.dataset.ddR = id; this.renderedRuns.add(id); }
          return '';
        });
        n.data = n.data.replace(TAG_RE, '');
        if (!n.data) n.remove();
      }
      stage.querySelectorAll('[data-dd-r]').forEach((span) => {
        const run = this.runAt(+span.dataset.ddR);
        if (run && (!run.simple || isDeleted(run))) span.contentEditable = 'false';
      });
      stage.querySelectorAll('del').forEach((d) => { d.contentEditable = 'false'; });
      stage.querySelectorAll('span[class*="tab-stop"]').forEach((t) => { t.dataset.ddTab = '1'; t.contentEditable = 'false'; });
      stage.querySelectorAll('a').forEach((a) => {
        const r = a.querySelector('[data-dd-r]');
        const run = r && this.runAt(+r.dataset.ddR);
        const link = run && [...run.chain].reverse().find((w) => w.localName === 'hyperlink');
        if (link) a.dataset.ddH = run.chain.indexOf(link) + ':' + r.dataset.ddR;
        a.addEventListener('click', (e) => { if (!e.ctrlKey) e.preventDefault(); });
      });
      // Body tables -> index of the w:tbl.
      this.renderedTables = new Set();
      this.idx.tables.forEach((tbl, ti) => {
        const firstP = tbl.getElementsByTagNameNS(W, 'p')[0];
        const i = firstP ? this.idx.paras.indexOf(firstP) : -1;
        if (i < 0) return;
        let depth = 0;
        for (let a = firstP.parentNode; a && a !== tbl; a = a.parentNode) if (isW(a, 'tbl')) depth++;
        const domP = stage.querySelector(`article p[data-dd-p="0:${i}"]`);
        let t = domP && domP.closest('table');
        for (let d = 0; t && d < depth; d++) t = t.parentElement.closest('table');
        if (t) { t.dataset.ddT = ti; this.renderedTables.add(ti); }
      });
      // docx-preview can list the same note twice (a reference run with its own superscript formatting is drawn
      // twice); keep one copy so it is neither shown nor saved twice.
      const seenNotes = new Set();
      for (const li of [...stage.querySelectorAll('ol > li')]) {
        const keys = [...li.querySelectorAll('p[data-dd-p]')].map((p) => p.dataset.ddP);
        if (!keys.length) continue;
        const k = keys.join(',');
        if (seenNotes.has(k)) li.remove(); else seenNotes.add(k);
      }
      // Editable areas: the body, headers, footers and the notes lists.
      const areas = new Set([...stage.querySelectorAll('section.docx > article, section.docx > header, section.docx > footer')]);
      stage.querySelectorAll('p[data-dd-p]').forEach((p) => { const pi = partOf(p.dataset.ddP); const kind = this.parts[pi] && this.parts[pi].kind; if (kind === 'footnotes' || kind === 'endnotes') { const ol = p.closest('ol'); if (ol) areas.add(ol); } });
      this.areas = [...areas];
      this.articles = this.areas.filter((a) => a.tagName === 'ARTICLE');
      for (const a of this.areas) {
        a.contentEditable = 'true';
        a.spellcheck = true;
        if (a.tagName !== 'ARTICLE') { a.classList.add('dd-hf'); a.title = a.tagName === 'OL' ? 'Notes' : a.tagName === 'HEADER' ? 'Header' : 'Footer'; }
      }
      // Mark every element that exists now, so formatting wrappers added while editing can be told apart.
      for (const a of this.areas) a.querySelectorAll('*').forEach((el) => { el.dataset.ddO = '1'; });
      // Baselines for change detection.
      this.paraSig = new Map();
      this.paraStyle0 = new Map();
      stage.querySelectorAll('p[data-dd-p]').forEach((p) => {
        if (this.paraSig.has(p.dataset.ddP)) return;
        this.paraSig.set(p.dataset.ddP, p.outerHTML);
        this.paraStyle0.set(p.dataset.ddP, { textAlign: p.style.textAlign, marginLeft: p.style.marginLeft });
      });
      this.runStyle0 = new Map();
      stage.querySelectorAll('[data-dd-r]').forEach((s) => { if (!this.runStyle0.has(+s.dataset.ddR)) this.runStyle0.set(+s.dataset.ddR, inlineStyle(s)); });
    }

    wire() {
      const onInput = (e) => {
        if (e.inputType === 'historyUndo') this.nativeDepth = Math.max(0, this.nativeDepth - 1);
        else if (e.inputType === 'historyRedo') this.nativeDepth++;
        else { this.nativeDepth++; this.future = []; }
        if (e.inputType === 'insertParagraph') this.applyNextStyle();
        clearTimeout(this.markerTimer);
        this.markerTimer = setTimeout(() => this.drawPageMarkers(), 300);
        if (this.opts.onChange) this.opts.onChange();
      };
      for (const a of this.areas) {
        a.addEventListener('input', onInput);
        a.addEventListener('paste', (e) => {
          e.preventDefault();
          const html = e.clipboardData.getData('text/html');
          const text = e.clipboardData.getData('text/plain');
          // Formatted paste keeps bold/italic/colours/links; Ctrl+Shift+V (or plain text) takes the formatting where it lands.
          if (html && !this.plainPaste) document.execCommand('insertHTML', false, cleanPastedHtml(html));
          else if (text) document.execCommand('insertText', false, text);
          this.plainPaste = false;
        });
        a.addEventListener('drop', (e) => { if (e.dataTransfer && e.dataTransfer.types.includes('Files')) return; e.preventDefault(); });
        a.addEventListener('beforeinput', (e) => { if (this.suggesting) this.suggestInput(e); });
        a.addEventListener('keydown', (e) => {
          if (e.key === 'Tab' && !e.ctrlKey) { e.preventDefault(); this.insertTab(); }
          if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'v') this.plainPaste = true;
        });
      }
    }

    insertTab() {
      const span = document.createElement('span');
      span.dataset.ddTab = '1';
      span.contentEditable = 'false';
      span.className = 'docx-tab-stop';
      span.textContent = ' ';
      const sel = getSelection();
      if (!sel.rangeCount) return;
      const r = sel.getRangeAt(0);
      r.deleteContents();
      r.insertNode(span);
      r.setStartAfter(span); r.collapse(true);
      sel.removeAllRanges(); sel.addRange(r);
      if (this.opts.onChange) this.opts.onChange();
    }

    // After Enter at the end of e.g. a heading, Word switches to the style's "next" style.
    applyNextStyle() {
      const p = this.caretParagraph();
      if (!p) return;
      const prev = p.previousElementSibling;
      if (!prev || prev.dataset.ddP !== p.dataset.ddP || p.textContent.trim()) return;
      const cur = this.styleOf(p);
      const info = cur && this.styles.byId[cur];
      if (info && info.next && info.next !== cur) this.setParagraphStyle([p], info.next);
    }

    // ---------- Suggesting mode: typing is an insertion, deleting marks text as deleted (both tracked changes)

    suggestInput(e) {
      const t = e.inputType;
      const sel = getSelection();
      if (!sel.rangeCount) return;
      const isDelete = /^delete/.test(t);
      const isInsert = t === 'insertText' || t === 'insertReplacementText' || t === 'insertFromPaste' || t === 'insertFromDrop';
      if (!isDelete && !isInsert) return; // paragraph breaks and formatting go through unchanged
      e.preventDefault();
      let range = sel.getRangeAt(0);
      if (isDelete && range.collapsed) {
        const dir = /Backward/.test(t) ? 'backward' : 'forward';
        const unit = /Word/.test(t) ? 'word' : /Line|Soft/.test(t) ? 'lineboundary' : 'character';
        sel.modify('extend', dir, unit);
        range = sel.getRangeAt(0);
        if (range.collapsed) return;
      }
      const backward = /Backward/.test(t);
      if (!range.collapsed) this.markDeleted(range);
      if (isInsert) {
        const text = e.data != null ? e.data : (e.dataTransfer ? e.dataTransfer.getData('text/plain') : '');
        if (text) this.insertSuggestion(text);
      } else {
        const r = sel.getRangeAt(0);
        r.collapse(backward);
        sel.removeAllRanges(); sel.addRange(r);
      }
      clearTimeout(this.markerTimer);
      this.markerTimer = setTimeout(() => this.drawPageMarkers(), 300);
      this.nativeDepth++;
      if (this.opts.onChange) this.opts.onChange();
    }

    // Wrap the text in a range as deleted (text this session suggested inserting is simply removed).
    markDeleted(range) {
      const root = range.commonAncestorContainer.nodeType === 1 ? range.commonAncestorContainer : range.commonAncestorContainer.parentNode;
      const nodes = [];
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) if (range.intersectsNode(n) && n.data.length) nodes.push(n);
      if (root.nodeType === 3) nodes.push(root);
      let last = null;
      for (const n of nodes) {
        const a = n === range.startContainer ? range.startOffset : 0;
        const b = n === range.endContainer ? range.endOffset : n.data.length;
        if (b <= a) continue;
        const host = n.parentElement;
        if (host.closest('.dd-sugg-del, del, [contenteditable="false"]')) { last = host.closest('.dd-sugg-del, del, [contenteditable="false"]'); continue; }
        let mid = n;
        if (a > 0) mid = n.splitText(a);
        if (b - a < mid.data.length) mid.splitText(b - a);
        if (host.closest('.dd-sugg-ins')) { mid.remove(); continue; }
        const span = document.createElement('span');
        span.className = 'dd-sugg-del';
        span.contentEditable = 'false';
        mid.replaceWith(span);
        span.append(mid);
        last = span;
      }
      const sel = getSelection();
      if (last && last.isConnected) { const r = document.createRange(); r.setStartAfter(last); r.collapse(true); sel.removeAllRanges(); sel.addRange(r); }
      // Remove insertion spans that are now empty.
      for (const sp of root.querySelectorAll ? root.querySelectorAll('.dd-sugg-ins') : []) if (!sp.textContent) sp.remove();
    }

    // Insert text as a suggestion at the caret, continuing an insertion span the caret is in.
    insertSuggestion(text) {
      const sel = getSelection();
      const r = sel.getRangeAt(0);
      const host = r.startContainer.nodeType === 3 ? r.startContainer.parentElement : r.startContainer;
      const ins = host && host.closest && host.closest('.dd-sugg-ins');
      if (ins && r.startContainer.nodeType === 3) {
        const n = r.startContainer;
        n.insertData(r.startOffset, text);
        const nr = document.createRange(); nr.setStart(n, r.startOffset + text.length); nr.collapse(true);
        sel.removeAllRanges(); sel.addRange(nr);
        return;
      }
      const span = document.createElement('span');
      span.className = 'dd-sugg-ins';
      const tn = document.createTextNode(text);
      span.append(tn);
      r.insertNode(span);
      const nr = document.createRange(); nr.setStart(tn, text.length); nr.collapse(true);
      sel.removeAllRanges(); sel.addRange(nr);
    }

    // ---------- undo that reaches across re-renders (lists, tables, pictures, page breaks)

    canUndo() { return this.nativeDepth > 0 || this.history.length > 0; }
    async undo() {
      if (this.nativeDepth > 0) { document.execCommand('undo'); return; }
      const prev = this.history.pop();
      if (!prev) return;
      const now = (await this.build()).bytes;
      this.future.push({ bytes: now, scrollTop: this.host.parentElement.scrollTop });
      await this.load(prev.bytes, prev);
      if (this.opts.onChange) this.opts.onChange();
    }
    async redo() {
      if (!this.future.length) { document.execCommand('redo'); return; }
      const next = this.future.pop();
      const now = (await this.build()).bytes;
      this.history.push({ bytes: now, scrollTop: this.host.parentElement.scrollTop });
      await this.load(next.bytes, next);
      if (this.opts.onChange) this.opts.onChange();
    }

    // ---------- page markers in the continuous edit view (where pages would break, by height)

    drawPageMarkers() {
      if (!this.stage) return;
      this.stage.querySelectorAll('.dd-page-marker').forEach((m) => m.remove());
      for (const section of this.stage.querySelectorAll('section.docx')) {
        const article = section.querySelector(':scope > article');
        if (!article) continue;
        // Text in columns flows side by side, so heights don't tell where pages end.
        if ((parseInt(getComputedStyle(article).columnCount, 10) || 1) > 1) continue;
        const cs = getComputedStyle(section);
        const pageH = parseFloat(section.style.minHeight) || 0;
        if (!pageH) continue;
        const header = section.querySelector(':scope > header');
        const footer = section.querySelector(':scope > footer');
        const usable = pageH - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom) - (header ? header.offsetHeight : 0) - (footer ? footer.offsetHeight : 0);
        if (usable < 100) continue;
        let used = 0, page = 1;
        for (const block of article.children) {
          const h = block.offsetHeight + parseFloat(getComputedStyle(block).marginTop) + parseFloat(getComputedStyle(block).marginBottom);
          if (used > 0 && used + h > usable) {
            page++;
            used = 0;
            const m = document.createElement('div');
            m.className = 'dd-page-marker';
            m.contentEditable = 'false';
            m.dataset.page = `Page ${page}`;
            m.style.top = `${block.offsetTop - 6}px`;
            section.append(m);
          }
          used += h;
        }
      }
    }

    // ---------- commands used by the toolbar

    caretParagraph() {
      const sel = getSelection();
      if (!sel.rangeCount) return null;
      const n = sel.anchorNode;
      const el = n && (n.nodeType === 1 ? n : n.parentElement);
      const p = el && el.closest('p');
      return p && this.stage.contains(p) ? p : null;
    }

    selectedParagraphs() {
      const sel = getSelection();
      if (!sel.rangeCount) return [];
      const r = sel.getRangeAt(0);
      const all = this.areas.flatMap((a) => [...a.querySelectorAll('p')]).filter((p) => !p.closest('[data-dd-r]'));
      const hit = all.filter((p) => r.intersectsNode(p));
      if (hit.length) return hit;
      const p = this.caretParagraph();
      return p ? [p] : [];
    }

    paraNode(key) {
      const [pi, i] = String(key).split(':').map(Number);
      const part = this.parts[pi];
      return part ? part.idx.paras[i] : null;
    }

    styleOf(p) {
      if (p.dataset.ddStyle !== undefined) return p.dataset.ddStyle;
      const node = p.dataset.ddP !== undefined ? this.paraNode(p.dataset.ddP) : null;
      const ps = node && child(child(node, 'pPr') || node, 'pStyle');
      return ps ? ps.getAttributeNS(W, 'val') : (this.styles.defaultPara || '');
    }

    setParagraphStyle(ps, styleId) {
      const cls = (id) => `docx_${(id || '').replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase()}`;
      ps.forEach((p) => {
        const old = this.styleOf(p);
        if (old) p.classList.remove(cls(old));
        if (styleId) p.classList.add(cls(styleId));
        p.dataset.ddStyle = styleId;
      });
      if (this.opts.onChange) this.opts.onChange();
    }

    indent(ps, steps) {
      ps.forEach((p) => {
        const n = (+(p.dataset.ddIndent || 0)) + steps;
        p.dataset.ddIndent = n;
        const base = parseFloat(this.paraStyle0.get(p.dataset.ddP)?.marginLeft || '0') || 0;
        p.style.marginLeft = `${Math.max(0, base + n * 48)}px`;
      });
      if (this.opts.onChange) this.opts.onChange();
    }

    // Structural edits are applied to the XML and the document re-rendered (and can be undone as a whole).
    async structural(fn, prepare, opts) {
      const caret = this.caretParagraph();
      if (fn && !(opts && opts.anyPart) && caret && caret.dataset.ddP && partOf(caret.dataset.ddP) !== 0) throw new Error('Tables, pictures and page breaks can only be added to the main text, not headers, footers or notes.');
      if (caret) caret.dataset.ddCaret = '1';
      const scrollTop = this.host.parentElement.scrollTop;
      const beforeBytes = (await this.build({ keepCaret: true })).bytes;
      if (prepare) await prepare();
      const built = await this.build({ structural: fn });
      this.history.push({ bytes: beforeBytes, scrollTop, caretPara: built.caretKey });
      this.future = [];
      await this.load(built.bytes, { scrollTop, caretPara: built.caretKey });
      if (this.opts.onChange) this.opts.onChange();
    }

    toggleList(kind) {
      const ps = this.selectedParagraphs();
      if (!ps.length) return;
      ps.forEach((p) => { p.dataset.ddList = kind; });
      return this.structural(null);
    }

    insertTable(rows, cols) {
      return this.structural((ctx) => {
        const { doc, caretNode } = ctx;
        const sectPr = doc.getElementsByTagNameNS(W, 'sectPr');
        const pgW = sectPr.length ? +(child(sectPr[sectPr.length - 1], 'pgSz')?.getAttributeNS(W, 'w') || 12240) : 12240;
        const mar = sectPr.length ? child(sectPr[sectPr.length - 1], 'pgMar') : null;
        const width = pgW - (mar ? (+mar.getAttributeNS(W, 'left') || 1440) + (+mar.getAttributeNS(W, 'right') || 1440) : 2880);
        const colW = Math.floor(width / cols);
        const tbl = wEl(doc, 'tbl');
        const tblPr = wEl(doc, 'tblPr');
        if (this.styles.byId.TableGrid) tblPr.append(wEl(doc, 'tblStyle', { val: 'TableGrid' }));
        tblPr.append(wEl(doc, 'tblW', { w: 0, type: 'auto' }));
        const borders = wEl(doc, 'tblBorders');
        for (const side of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) borders.append(wEl(doc, side, { val: 'single', sz: 4, space: 0, color: 'auto' }));
        tblPr.append(borders);
        tblPr.append(wEl(doc, 'tblLook', { val: '04A0', firstRow: 1, lastRow: 0, firstColumn: 1, lastColumn: 0, noHBand: 0, noVBand: 1 }));
        tbl.append(tblPr);
        const grid = wEl(doc, 'tblGrid');
        for (let c = 0; c < cols; c++) grid.append(wEl(doc, 'gridCol', { w: colW }));
        tbl.append(grid);
        for (let r = 0; r < rows; r++) {
          const tr = wEl(doc, 'tr');
          for (let c = 0; c < cols; c++) {
            const tc = wEl(doc, 'tc');
            const tcPr = wEl(doc, 'tcPr');
            tcPr.append(wEl(doc, 'tcW', { w: colW, type: 'dxa' }));
            tc.append(tcPr, wEl(doc, 'p'));
            tr.append(tc);
          }
          tbl.append(tr);
        }
        insertAfterBlock(caretNode, tbl, doc);
      });
    }

    insertPageBreak() {
      return this.structural(({ doc, caretNode }) => {
        const p = wEl(doc, 'p');
        const r = wEl(doc, 'r');
        r.append(wEl(doc, 'br', { type: 'page' }));
        p.append(r);
        insertAfterBlock(caretNode, p, doc);
      });
    }

    insertImage(bytes, ext, widthPx, heightPx) {
      return this.structural(async (ctx) => {
        const { doc, caretNode } = ctx;
        const name = await ctx.addMedia(bytes, ext);
        const rid = ctx.addRel(REL_IMAGE, `media/${name}`);
        const maxW = 624;
        const scale = widthPx > maxW ? maxW / widthPx : 1;
        const cx = Math.round(widthPx * scale * 9525);
        const cy = Math.round(heightPx * scale * 9525);
        const id = 5000 + Math.floor(Math.random() * 100000);
        const xml = `<w:p xmlns:w="${W}"><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="${R}" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
        insertAfterBlock(caretNode, doc.importNode(parseXml(xml).documentElement, true), doc);
      });
    }

    // ======================================================================
    // Building the new package
    // ======================================================================

    // DOM paragraphs of one part, grouped by the source paragraph they came from (Enter clones the attributes).
    // A header/footer used by several sections is rendered several times; the copy the user changed wins.
    domGroups(part) {
      const pi = part.pi;
      const containers = new Map();
      for (const area of this.areas) {
        let lastId = null;
        for (const p of area.querySelectorAll('p')) {
          if (p.closest('[data-dd-r]')) continue;
          const id = p.dataset.ddP !== undefined ? p.dataset.ddP : lastId;
          if (id == null) continue;
          lastId = id;
          if (partOf(id) !== pi) continue;
          if (!containers.has(area)) containers.set(area, []);
          containers.get(area).push({ p, id });
        }
      }
      let chosen = [...containers.keys()];
      if (part.kind !== 'body' && chosen.length > 1) {
        const changed = chosen.find((area) => containers.get(area).some(({ p, id }) => !this.isClean(p, id)) || containers.get(area).length !== new Set(containers.get(area).map((x) => x.id)).size);
        chosen = [changed || chosen[0]];
      }
      const groups = new Map();
      for (const area of chosen) for (const { p, id } of containers.get(area)) {
        if (!groups.has(id)) groups.set(id, []);
        groups.get(id).push(p);
      }
      return { groups, domPs: chosen.flatMap((a) => containers.get(a).map((x) => x.p)) };
    }

    async build(opts) {
      opts = opts || {};
      const out = await JSZip.loadAsync(this.bytes);
      const extraFiles = {};
      let ctTypes = null;
      let caretNode = null;
      let bodyDoc = null;
      let bodyOutputs = null;
      let bodyRels = null;
      const builtParts = new Map();
      for (const part of this.parts) {
        const res = await this.buildPart(part, extraFiles);
        builtParts.set(part.path, res);
        if (res.caretNode) caretNode = res.caretNode;
        if (part.kind === 'body') { bodyDoc = res.doc; bodyOutputs = res; bodyRels = res.rels; }
      }
      const otherDocs = new Map(); // other XML parts read or created by a structural edit
      const doc = bodyDoc;
      const body = doc.getElementsByTagNameNS(W, 'body')[0];
      // Structural edit (insert table/image/break) at the caret.
      if (opts.structural) {
        const sctx = {
          doc, caretNode: caretNode && caretNode.ownerDocument === doc ? caretNode : lastBodyBlock(body),
          addRel: (type, target) => bodyRels.add(type, target),
          addMedia: async (bytes, ext) => {
            const dir = this.docPath.replace(/[^/]+$/, 'media/');
            let n = 1;
            while (this.zip.file(`${dir}image${n}.${ext}`) || extraFiles[`${dir}image${n}.${ext}`]) n++;
            extraFiles[`${dir}image${n}.${ext}`] = bytes;
            ctTypes = ctTypes || parseXml(await this.zip.file('[Content_Types].xml').async('string'));
            const root = ctTypes.documentElement;
            if (![...root.children].some((d) => d.localName === 'Default' && (d.getAttribute('Extension') || '').toLowerCase() === ext)) {
              const d = ctTypes.createElementNS(root.namespaceURI, 'Default');
              d.setAttribute('Extension', ext);
              d.setAttribute('ContentType', ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`);
              root.insertBefore(d, root.firstChild);
            }
            return `image${n}.${ext}`;
          },
        };
        const ct = async () => (ctTypes = ctTypes || parseXml(await this.zip.file('[Content_Types].xml').async('string')));
        Object.assign(sctx, {
          W, R, wEl, child, kids, isW, setOrdered, removeChild, PPR_ORDER, RPR_ORDER, parseXml, insertAfterBlock, styles: this.styles,
          // An XML part by path: an editable part already built above, or any other part (read once, written back).
          partDoc: async (path) => {
            if (builtParts.has(path)) { builtParts.get(path).changed = true; return builtParts.get(path).doc; }
            if (!otherDocs.has(path)) otherDocs.set(path, this.zip.file(path) ? parseXml(await this.zip.file(path).async('string')) : null);
            return otherDocs.get(path);
          },
          partRels: (path) => (builtParts.has(path) ? builtParts.get(path).rels : null),
          // A new part linked from the document, with its content type.
          createPart: async (name, relType, contentType, xml) => {
            const dir = this.docPath.replace(/[^/]+$/, '');
            const path = dir + name;
            otherDocs.set(path, parseXml(xml));
            const rid = bodyRels.add(relType, name);
            const c = await ct();
            const o = c.createElementNS(c.documentElement.namespaceURI, 'Override');
            o.setAttribute('PartName', `/${path}`);
            o.setAttribute('ContentType', contentType);
            c.documentElement.append(o);
            return { path, rid, doc: otherDocs.get(path) };
          },
          relTarget: (type) => { const r = [...bodyRels.doc.documentElement.children].find((x) => x.getAttribute('Type') === type); return r ? resolve(this.docPath, r.getAttribute('Target')) : null; },
          docDir: this.docPath.replace(/[^/]+$/, ''),
          bodyRelsDoc: () => bodyRels.doc,
          knownPaths: () => [...Object.keys(this.zip.files), ...otherDocs.keys(), ...Object.keys(extraFiles)],
          sectPrs: () => [...doc.getElementsByTagNameNS(W, 'sectPr')],
          // Replace the inline marker character `ch` (typed at the caret) with nodes; the run around it is split.
          atMarker: (ch, make) => {
            const docs = [doc, ...[...builtParts.values()].map((r) => r.doc)];
            for (const d of docs) {
              for (const t of [...d.getElementsByTagNameNS(W, 't')]) {
                const i = t.textContent.indexOf(ch);
                if (i < 0) continue;
                const r = t.parentNode;
                const rPr = child(r, 'rPr');
                const before = r.cloneNode(false), after = r.cloneNode(false);
                let seen = false;
                for (const c of [...r.childNodes]) {
                  if (c === t) {
                    seen = true;
                    const tb = t.cloneNode(false); tb.textContent = t.textContent.slice(0, i);
                    const ta = t.cloneNode(false); ta.textContent = t.textContent.slice(i + 1);
                    if (tb.textContent) before.append(tb);
                    if (ta.textContent) after.append(ta);
                    continue;
                  }
                  if (isW(c, 'rPr')) { before.append(c.cloneNode(true)); after.append(c.cloneNode(true)); continue; }
                  (seen ? after : before).append(c.cloneNode(true));
                }
                const has = (x) => kids(x).some((k) => !isW(k, 'rPr'));
                const nodes = [].concat(make(d, rPr ? rPr.cloneNode(true) : null) || []);
                const parent = r.parentNode;
                if (has(before)) parent.insertBefore(before, r);
                for (const n of nodes) parent.insertBefore(n, r);
                if (has(after)) parent.insertBefore(after, r);
                parent.removeChild(r);
                if (d !== doc) for (const res of builtParts.values()) if (res.doc === d) res.changed = true;
                return d;
              }
            }
            return null;
          },
        });
        await opts.structural(sctx);
      }
      for (const part of this.parts) {
        const res = builtParts.get(part.path);
        if (part.kind === 'body' || !res.changed) continue;
        out.file(part.path, serialize(res.doc));
        if (res.rels.changed) out.file(part.relsPath, serialize(res.rels.doc));
      }
      for (const [path, d] of otherDocs) if (d) out.file(path, serialize(d));
      // Lists toggled on paragraphs (any part).
      const listed = [...this.stage.querySelectorAll('p[data-dd-list]')];
      if (listed.length) {
        const numbering = await this.ensureNumbering((t, x) => bodyRels.add(t, x), extraFiles, () => ctTypes, (v) => { ctTypes = v; });
        for (const domP of listed) {
          const node = this.outputFor(domP, bodyOutputs);
          if (!node) continue;
          const pdoc = node.ownerDocument;
          const pPr = node.firstChild && isW(node.firstChild, 'pPr') ? node.firstChild : (() => { const x = wEl(pdoc, 'pPr'); node.insertBefore(x, node.firstChild); return x; })();
          const cur = child(pPr, 'numPr');
          const curId = cur && child(cur, 'numId') ? child(cur, 'numId').getAttributeNS(W, 'val') : null;
          const wantId = numbering[domP.dataset.ddList];
          if (curId != null && String(curId) === String(wantId)) { removeChild(pPr, 'numPr'); continue; }
          const numPr = wEl(pdoc, 'numPr');
          numPr.append(wEl(pdoc, 'ilvl', { val: 0 }), wEl(pdoc, 'numId', { val: wantId }));
          setOrdered(pPr, numPr, PPR_ORDER);
          if (this.styles.byId.ListParagraph && !child(pPr, 'pStyle')) setOrdered(pPr, wEl(pdoc, 'pStyle', { val: 'ListParagraph' }), PPR_ORDER);
        }
      }
      let caretKey = null;
      if (caretNode && caretNode.ownerDocument === doc) {
        const at = partParagraphs(doc, 'body').indexOf(caretNode);
        if (at >= 0) caretKey = `0:${at}`;
      }
      if (!opts.keepCaret) this.stage.querySelectorAll('[data-dd-caret]').forEach((p) => delete p.dataset.ddCaret);
      out.file(this.docPath, serialize(doc));
      if (bodyRels.changed) out.file(this.relsPath, serialize(bodyRels.doc));
      if (ctTypes) out.file('[Content_Types].xml', serialize(ctTypes));
      for (const [p, data] of Object.entries(extraFiles)) out.file(p, data);
      const bytes = await out.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      return { bytes, caretKey };
    }

    async relsFor(part) {
      const f = this.zip.file(part.relsPath);
      const doc = f ? parseXml(await f.async('string')) : parseXml(`<Relationships xmlns="${PKG_REL}"/>`);
      const rels = {
        doc, changed: false,
        add: (type, target, external) => {
          const used = new Set([...doc.documentElement.children].map((r) => r.getAttribute('Id')));
          let n = used.size + 1;
          while (used.has(`rId${n}`)) n++;
          const rel = doc.createElementNS(PKG_REL, 'Relationship');
          rel.setAttribute('Id', `rId${n}`);
          rel.setAttribute('Type', type);
          rel.setAttribute('Target', target);
          if (external) rel.setAttribute('TargetMode', 'External');
          doc.documentElement.append(rel);
          rels.changed = true;
          return `rId${n}`;
        },
      };
      return rels;
    }

    async buildPart(part, extraFiles) {
      const pi = part.pi;
      const doc = part.doc.cloneNode(true);
      const idx = index(doc, part.kind, pi);
      const rels = await this.relsFor(part);
      const linkRels = new Map();
      const linkRel = (url) => { if (!linkRels.has(url)) linkRels.set(url, rels.add(REL_HYPERLINK, url, true)); return linkRels.get(url); };
      const { groups } = this.domGroups(part);
      const flushed = new Set();          // "part:para:pos" already written
      const touched = new Map();          // paragraph key -> output paragraph that received its runs
      const outputs = new Map();          // paragraph key -> [nodes]
      let caretNode = null;
      let changed = false;
      const run = (id) => (Math.floor(id / PART_BASE) === pi ? idx.runs[id % PART_BASE] : null);
      const ctx = { doc, idx, flushed, touched, linkRel, run, pi };
      for (const [key, dom] of groups) {
        if (!this.renderedParas.has(key)) continue;
        const i = +key.split(':')[1];
        const pNode = idx.paras[i];
        if (!pNode) continue;
        const hasSect = !!(child(pNode, 'pPr') && child(child(pNode, 'pPr'), 'sectPr'));
        const list = [];
        dom.forEach((domP, k) => {
          const first = k === 0;
          const last = k === dom.length - 1;
          let node;
          if (first && dom.length === 1 && this.isClean(domP, key)) node = pNode;
          else { node = this.buildParagraph(domP, key, { first, keepSect: hasSect && last }, ctx); changed = true; }
          if (domP.dataset.ddCaret) caretNode = node;
          list.push(node);
        });
        outputs.set(key, list);
      }
      const finishLeaves = (key, target) => {
        const i = +key.split(':')[1];
        idx.pinfo[i].items.forEach((it) => {
          const fk = `${key}:${it.pos}`;
          if (flushed.has(fk)) return;
          if (it.type === 'run' && this.renderedRuns.has(it.id)) return; // deleted by the user
          flushed.add(fk);
          appendWithChain(target, it.node.cloneNode(true), it.chain, doc);
        });
      };
      let lastOut = null;
      idx.paras.forEach((pNode, i) => {
        const key = `${pi}:${i}`;
        if (!this.renderedParas.has(key)) { lastOut = pNode; return; }
        const outList = outputs.get(key);
        const parent = pNode.parentNode;
        if (outList && outList.length === 1 && outList[0] === pNode) {
          idx.pinfo[i].items.forEach((it) => flushed.add(`${key}:${it.pos}`));
          lastOut = pNode;
          return;
        }
        changed = true;
        if (outList && outList.length) {
          for (const n of outList) parent.insertBefore(n, pNode);
          finishLeaves(key, outList[outList.length - 1]);
          parent.removeChild(pNode);
          lastOut = outList[outList.length - 1];
          return;
        }
        // Deleted paragraph.
        const sect = child(pNode, 'pPr') && child(child(pNode, 'pPr'), 'sectPr');
        if (sect) {
          const keep = wEl(doc, 'p');
          keep.append(child(pNode, 'pPr').cloneNode(true));
          parent.insertBefore(keep, pNode);
          finishLeaves(key, keep);
          lastOut = keep;
        } else {
          const target = touched.get(key) || (lastOut && isW(lastOut, 'p') ? lastOut : null);
          if (target) finishLeaves(key, target);
          else {
            const stub = wEl(doc, 'p');
            if (child(pNode, 'pPr')) stub.append(child(pNode, 'pPr').cloneNode(true));
            finishLeaves(key, stub);
            if (stub.childNodes.length > (child(pNode, 'pPr') ? 1 : 0)) { parent.insertBefore(stub, pNode); lastOut = stub; }
          }
        }
        parent.removeChild(pNode);
      });
      if (part.kind === 'body') {
        const domTables = new Set([...this.stage.querySelectorAll('table[data-dd-t]')].map((t) => +t.dataset.ddT));
        idx.tables.forEach((tbl, ti) => {
          if (this.renderedTables.has(ti) && !domTables.has(ti) && tbl.parentNode) { tbl.parentNode.removeChild(tbl); changed = true; }
        });
      }
      // Every table cell, header, footer and note must keep at least one paragraph.
      for (const tc of [...doc.getElementsByTagNameNS(W, 'tc')]) {
        const last = kids(tc).filter((c) => !isW(c, 'tcPr')).pop();
        if (!last || !isW(last, 'p')) tc.append(wEl(doc, 'p'));
      }
      const root = partRoot(doc, part.kind);
      if (part.kind === 'body') {
        if (!kids(root).some((c) => isW(c, 'p') || isW(c, 'tbl'))) root.insertBefore(wEl(doc, 'p'), child(root, 'sectPr'));
      } else if (part.kind === 'header' || part.kind === 'footer') {
        if (!kids(root).some((c) => isW(c, 'p') || isW(c, 'tbl'))) root.append(wEl(doc, 'p'));
      } else {
        for (const note of kids(root)) if (!kids(note).some((c) => isW(c, 'p'))) note.append(wEl(doc, 'p'));
      }
      return { doc, rels, outputs, caretNode, changed };
    }

    outputFor(domP, res) {
      const id = domP.dataset.ddP;
      if (!id || partOf(id) !== 0 || !res || !res.outputs.has(id)) return null;
      const same = [...this.stage.querySelectorAll(`article p[data-dd-p="${id}"]`)].filter((p) => !p.closest('[data-dd-r]'));
      return res.outputs.get(id)[same.indexOf(domP)] || null;
    }

    isClean(domP, key) {
      return domP.outerHTML === this.paraSig.get(key) && !domP.dataset.ddStyle && !domP.dataset.ddList && !domP.dataset.ddIndent && !domP.dataset.ddSpacing && !domP.dataset.ddInd && !domP.dataset.ddBorder;
    }

    // Build a new w:p for an edited DOM paragraph that came from source paragraph `key`.
    buildParagraph(domP, key, how, ctx) {
      const { doc, idx, flushed, touched } = ctx;
      const i = +key.split(':')[1];
      const src = idx.paras[i];
      const p = wEl(doc, 'p');
      for (const a of src.attributes) {
        if (!how.first && /paraId|textId/.test(a.localName)) continue;
        p.setAttributeNode(a.cloneNode());
      }
      const srcPPr = child(src, 'pPr');
      const pPr = srcPPr ? srcPPr.cloneNode(true) : wEl(doc, 'pPr');
      if (!how.keepSect) removeChild(pPr, 'sectPr');
      if (domP.dataset.ddStyle !== undefined) {
        if (domP.dataset.ddStyle) setOrdered(pPr, wEl(doc, 'pStyle', { val: domP.dataset.ddStyle }), PPR_ORDER);
        else removeChild(pPr, 'pStyle');
      }
      const base = this.paraStyle0.get(key) || {};
      if ((domP.style.textAlign || '') !== (base.textAlign || '')) {
        const a = domP.style.textAlign;
        if (!a) removeChild(pPr, 'jc');
        else setOrdered(pPr, wEl(doc, 'jc', { val: { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', justify: 'both' }[a] || 'left' }), PPR_ORDER);
      }
      if (domP.dataset.ddIndent && +domP.dataset.ddIndent) {
        const ind = child(pPr, 'ind') ? child(pPr, 'ind').cloneNode(true) : wEl(doc, 'ind');
        const left = +(ind.getAttributeNS(W, 'left') || ind.getAttributeNS(W, 'start') || 0);
        ind.removeAttributeNS(W, 'start');
        ind.setAttributeNS(W, 'w:left', String(Math.max(0, left + 720 * +domP.dataset.ddIndent)));
        setOrdered(pPr, ind, PPR_ORDER);
      }
      if (domP.dataset.ddSpacing) {
        const sp = JSON.parse(domP.dataset.ddSpacing);
        const el = child(pPr, 'spacing') ? child(pPr, 'spacing').cloneNode(true) : wEl(doc, 'spacing');
        if (sp.line != null) { el.setAttributeNS(W, 'w:line', String(Math.round(sp.line * 240))); el.setAttributeNS(W, 'w:lineRule', 'auto'); }
        if (sp.before != null) { el.setAttributeNS(W, 'w:before', String(Math.round(sp.before))); el.removeAttributeNS(W, 'beforeAutospacing'); }
        if (sp.after != null) { el.setAttributeNS(W, 'w:after', String(Math.round(sp.after))); el.removeAttributeNS(W, 'afterAutospacing'); }
        setOrdered(pPr, el, PPR_ORDER);
        if (sp.keepNext != null) { if (sp.keepNext) setOrdered(pPr, wEl(doc, 'keepNext'), PPR_ORDER); else removeChild(pPr, 'keepNext'); }
      }
      // Exact indents from the ruler: { left, right, firstLine } in twips (firstLine < 0 = hanging).
      if (domP.dataset.ddInd) {
        const v = JSON.parse(domP.dataset.ddInd);
        const ind = child(pPr, 'ind') ? child(pPr, 'ind').cloneNode(true) : wEl(doc, 'ind');
        for (const a of ['start', 'end', 'firstLine', 'hanging', 'left', 'right']) ind.removeAttributeNS(W, a);
        if (v.left) ind.setAttributeNS(W, 'w:left', String(Math.round(v.left)));
        if (v.right) ind.setAttributeNS(W, 'w:right', String(Math.round(v.right)));
        if (v.firstLine > 0) ind.setAttributeNS(W, 'w:firstLine', String(Math.round(v.firstLine)));
        if (v.firstLine < 0) ind.setAttributeNS(W, 'w:hanging', String(Math.round(-v.firstLine)));
        if (ind.attributes.length) setOrdered(pPr, ind, PPR_ORDER); else removeChild(pPr, 'ind');
      }
      // Paragraph borders and shading: { border: { color, size (eighths of a point) } | null, shade: 'RRGGBB' | null }.
      if (domP.dataset.ddBorder) {
        const v = JSON.parse(domP.dataset.ddBorder);
        if ('border' in v) {
          removeChild(pPr, 'pBdr');
          if (v.border) {
            const b = wEl(doc, 'pBdr');
            for (const side of ['top', 'left', 'bottom', 'right']) b.append(wEl(doc, side, { val: 'single', sz: v.border.size || 6, space: 4, color: (v.border.color || '000000').replace('#', '') }));
            setOrdered(pPr, b, PPR_ORDER);
          }
        }
        if ('shade' in v) {
          removeChild(pPr, 'shd');
          if (v.shade) setOrdered(pPr, wEl(doc, 'shd', { val: 'clear', color: 'auto', fill: v.shade.replace('#', '') }), PPR_ORDER);
        }
      }
      if (pPr.childNodes.length) p.append(pPr);

      const items = this.readParagraph(domP, ctx);
      const out = { stack: [] };
      const emit = (node, chain) => {
        let k = 0;
        while (k < out.stack.length && k < chain.length && out.stack[k].key === chain[k]) k++;
        out.stack.length = k;
        for (; k < chain.length; k++) {
          const w = chain[k];
          const parentEl = k ? out.stack[k - 1].inner : p;
          const made = makeWrapper(w, doc, ctx);
          parentEl.append(made.outer);
          out.stack.push({ key: w, inner: made.inner });
        }
        (out.stack.length ? out.stack[out.stack.length - 1].inner : p).append(node);
      };
      const flushBefore = (run) => {
        const pk = run.pkey;
        const pinf = idx.pinfo[+pk.split(':')[1]];
        for (const it of pinf.items) {
          if (it.pos >= run.pos) break;
          const fk = `${pk}:${it.pos}`;
          if (flushed.has(fk)) continue;
          if (it.type === 'run' && this.renderedRuns.has(it.id)) { flushed.add(fk); continue; }
          flushed.add(fk);
          emit(it.node.cloneNode(true), it.chain);
        }
        flushed.add(`${pk}:${run.pos}`);
        if (pk !== key) touched.set(pk, p);
      };

      let group = null;
      const endGroup = () => {
        if (!group) return;
        const run = group.run;
        const r = wEl(doc, 'r');
        if (run) for (const a of run.node.attributes) r.setAttributeNode(a.cloneNode());
        const baseRPr = run ? child(run.node, 'rPr') : (srcPPr && child(srcPPr, 'rPr'));
        const rPr = baseRPr ? cloneAs(baseRPr) : wEl(doc, 'rPr');
        applyFormat(doc, rPr, group.fmt);
        if (rPr.childNodes.length) r.append(rPr);
        for (const piece of group.pieces) {
          if (piece.k === 'tab') r.append(wEl(doc, 'tab'));
          else if (piece.k === 'br') r.append(wEl(doc, 'br'));
          else {
            const t = wEl(doc, group.del ? 'delText' : 't');
            t.setAttributeNS(XMLNS_XML, 'xml:space', 'preserve');
            t.textContent = piece.text;
            r.append(t);
          }
        }
        emit(r, group.chain);
        group = null;
      };

      for (const it of items) {
        const run = it.run != null ? ctx.run(it.run) : null;
        if (run && !flushed.has(`${run.pkey}:${run.pos}`)) { endGroup(); flushBefore(run); }
        if (it.k === 'opaque') { endGroup(); if (run) emit(run.node.cloneNode(true), run.chain); continue; }
        const chain = this.chainFor(run, it.link, ctx);
        if (it.change) chain.push(`chg:${it.change}`);
        const fmtKey = JSON.stringify(it.fmt);
        if (group && group.run === run && group.fmtKey === fmtKey && sameChain(group.chain, chain)) group.pieces.push(it);
        else { endGroup(); group = { run, fmt: it.fmt, fmtKey, chain, pieces: [it], del: it.change === 'del' }; }
      }
      endGroup();
      if (how.first) {
        for (const it of idx.pinfo[i].items) {
          const fk = `${key}:${it.pos}`;
          if (flushed.has(fk)) continue;
          if (it.type === 'run' && this.renderedRuns.has(it.id)) continue;
          flushed.add(fk);
          emit(it.node.cloneNode(true), it.chain);
        }
      }
      return p;
    }

    chainFor(run, link, ctx) {
      const base = run ? run.chain.filter((w) => w.localName !== 'hyperlink') : [];
      if (!link) return base;
      if (link.orig) {
        const [chainIdx, runId] = link.orig;
        const r = ctx.run(runId);
        const w = r && r.chain[chainIdx];
        return w ? [...base, w] : base;
      }
      return [...base, `new:${link.url}`];
    }

    // Read a DOM paragraph into pieces: { k: 't'|'tab'|'br'|'opaque', run, text, fmt, link }.
    readParagraph(domP, ctx) {
      const items = [];
      let ctxRun = null;
      const linkOf = (node) => {
        const a = (node.nodeType === 1 ? node : node.parentElement).closest('a');
        if (!a || !domP.contains(a)) return null;
        if (a.dataset.ddH) return { orig: a.dataset.ddH.split(':').map(Number) };
        const href = a.getAttribute('href');
        return href ? { url: href } : null;
      };
      const fmtOf = (textNode) => {
        const fmt = {};
        for (let el = textNode.parentElement; el && el !== domP; el = el.parentElement) {
          if (el.dataset.ddR !== undefined) {
            cssToFormat(inlineStyle(el), fmt, this.runStyle0.get(+el.dataset.ddR) || {});
            break;
          }
          if (!el.dataset.ddO) { tagToFormat(el, fmt); cssToFormat(inlineStyle(el), fmt, null); }
        }
        for (const k of Object.keys(fmt)) if (fmt[k] === undefined) delete fmt[k];
        return fmt;
      };
      const walk = (node) => {
        for (const c of node.childNodes) {
          if (c.nodeType === 3) {
            const text = c.data.replace(/ /g, ' ').replace(/[-]/g, '');
            const sugg = c.parentElement.closest('.dd-sugg-ins, .dd-sugg-del');
            if (text) items.push({ k: 't', run: ctxRun, text, fmt: fmtOf(c), link: linkOf(c), change: sugg && domP.contains(sugg) ? (sugg.classList.contains('dd-sugg-del') ? 'del' : 'ins') : null });
            continue;
          }
          if (c.nodeType !== 1) continue;
          if (c.classList.contains('dd-page-marker')) continue;
          if (c.dataset.ddR !== undefined) {
            const id = +c.dataset.ddR;
            const run = ctx.run(id);
            if (!run) { ctxRun = null; walk(c); continue; } // text moved in from another part: keep the text
            if (!run.simple || isDeleted(run)) { items.push({ k: 'opaque', run: id }); ctxRun = id; continue; }
            ctxRun = id;
            walk(c);
            continue;
          }
          if (c.dataset.ddTab) { items.push({ k: 'tab', run: ctxRun, fmt: {}, link: linkOf(c) }); continue; }
          if (c.tagName === 'BR') { items.push({ k: 'br', run: ctxRun, fmt: {}, link: linkOf(c) }); continue; }
          if (c.tagName === 'IMG' || c.tagName === 'svg') continue;
          walk(c);
        }
      };
      walk(domP);
      if (items.length && items[items.length - 1].k === 'br') items.pop();
      for (let k = 0; k < items.length; k++) {
        if ((items[k].k === 'tab' || items[k].k === 'br') && k > 0 && items[k - 1].run === items[k].run && items[k - 1].fmt) items[k].fmt = items[k - 1].fmt;
      }
      return items;
    }

    async ensureNumbering(addRel, extraFiles, getCt, setCt) {
      const dir = this.docPath.replace(/[^/]+$/, '');
      const path = `${dir}numbering.xml`;
      let xml = extraFiles[path] ? parseXml(extraFiles[path]) : (this.zip.file(path) ? parseXml(await this.zip.file(path).async('string')) : null);
      const created = !xml;
      if (!xml) xml = parseXml(`<w:numbering xmlns:w="${W}"/>`);
      const root = xml.documentElement;
      const abstracts = [...root.getElementsByTagNameNS(W, 'abstractNum')];
      const nums = [...root.getElementsByTagNameNS(W, 'num')];
      const fmtOfAbstract = (a) => { const l = [...a.getElementsByTagNameNS(W, 'lvl')].find((x) => x.getAttributeNS(W, 'ilvl') === '0'); const f = l && child(l, 'numFmt'); return f ? f.getAttributeNS(W, 'val') : null; };
      const result = {};
      const textOfAbstract = (a) => { const l = [...a.getElementsByTagNameNS(W, 'lvl')].find((x) => x.getAttributeNS(W, 'ilvl') === '0'); const t = l && child(l, 'lvlText'); return t ? t.getAttributeNS(W, 'val') : null; };
      const wanted = new Set([...this.stage.querySelectorAll('p[data-dd-list]')].map((p) => p.dataset.ddList));
      for (const [kind, def] of Object.entries(LIST_KINDS)) {
        if (kind !== 'bullet' && kind !== 'number' && !wanted.has(kind)) continue;
        const fmt = def.fmt;
        const abs = abstracts.find((a) => fmtOfAbstract(a) === fmt && (def.first == null || textOfAbstract(a) === def.first || (kind === 'bullet' && ['', '', '•'].includes(textOfAbstract(a)))));
        const num = abs && nums.find((n) => child(n, 'abstractNumId') && child(n, 'abstractNumId').getAttributeNS(W, 'val') === abs.getAttributeNS(W, 'abstractNumId'));
        if (num) { result[kind] = num.getAttributeNS(W, 'numId'); continue; }
        const absId = Math.max(-1, ...abstracts.map((a) => +a.getAttributeNS(W, 'abstractNumId'))) + 1;
        const numId = Math.max(0, ...nums.map((n) => +n.getAttributeNS(W, 'numId'))) + 1;
        const esc = (x) => x.replace(/&/g, '&amp;').replace(/"/g, '&quot;');
        const a = parseXml(`<w:abstractNum xmlns:w="${W}" w:abstractNumId="${absId}"><w:multiLevelType w:val="hybridMultilevel"/>${
          [0, 1, 2, 3, 4, 5, 6, 7, 8].map((l) => { const lv = def.level(l); return `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${lv.fmt}"/><w:lvlText w:val="${esc(lv.text)}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (l + 1)}" w:hanging="360"/></w:pPr>${lv.font ? `<w:rPr><w:rFonts w:ascii="${lv.font}" w:hAnsi="${lv.font}" w:hint="default"/></w:rPr>` : ''}</w:lvl>`; }).join('')
        }</w:abstractNum>`).documentElement;
        const firstNum = root.getElementsByTagNameNS(W, 'num')[0];
        root.insertBefore(xml.importNode(a, true), firstNum || null);
        const n = parseXml(`<w:num xmlns:w="${W}" w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num>`).documentElement;
        root.append(xml.importNode(n, true));
        abstracts.push(a); nums.push(n);
        result[kind] = String(numId);
      }
      extraFiles[path] = serialize(xml);
      if (created) {
        addRel(REL_NUMBERING, 'numbering.xml');
        const ct = getCt() || parseXml(await this.zip.file('[Content_Types].xml').async('string'));
        const r = ct.documentElement;
        const o = ct.createElementNS(r.namespaceURI, 'Override');
        o.setAttribute('PartName', `/${path}`);
        o.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
        r.append(o);
        setCt(ct);
      }
      return result;
    }

    destroy() { this.host.replaceChildren(); }
  }

  // ---------- helpers

  function cloneAs(el) {
    const c = el.cloneNode(true);
    // A paragraph mark's rPr can carry revision marks that aren't valid in a run's rPr.
    for (const k of kids(c)) if (['ins', 'del', 'moveFrom', 'moveTo'].includes(k.localName)) c.removeChild(k);
    return c;
  }

  // Tracked-change ids must be unique in the part.
  function nextChangeId(doc) {
    if (doc.__ddChangeId == null) {
      let max = 0;
      for (const n of doc.getElementsByTagNameNS(W, '*')) { const v = n.getAttributeNS(W, 'id'); if (v && /^(ins|del|rPrChange|pPrChange|moveFrom|moveTo)$/.test(n.localName) && +v > max) max = +v; }
      doc.__ddChangeId = Math.max(max, 9000);
    }
    return ++doc.__ddChangeId;
  }
  function changeAuthor() { try { return (localStorage.getItem('docstudio.author') || '').trim() || 'Doc Studio'; } catch { return 'Doc Studio'; } }
  // A run inside a tracked deletion (its text is w:delText).
  function isDeleted(run) { return run.chain.some((w) => w.localName === 'del' || w.localName === 'moveFrom'); }

  // List styles: Word bullets (Symbol \uF0B7 shown as •) and numbering formats.
  const sub = (l, list) => list[l % list.length];
  const LIST_KINDS = {
    bullet: { fmt: 'bullet', first: '•', level: (l) => (l % 3 === 0 ? { fmt: 'bullet', text: '•' } : { fmt: 'bullet', text: sub(l, ['', '◦', '▪']) }) },
    number: { fmt: 'decimal', first: '%1.', level: (l) => ({ fmt: sub(l, ['decimal', 'lowerLetter', 'lowerRoman']), text: `%${l + 1}.` }) },
    arrow: { fmt: 'bullet', first: '➢', level: (l) => ({ fmt: 'bullet', text: sub(l, ['➢', '◦', '▪']) }) },
    check: { fmt: 'bullet', first: '✓', level: (l) => ({ fmt: 'bullet', text: sub(l, ['✓', '◦', '▪']) }) },
    square: { fmt: 'bullet', first: '▪', level: (l) => ({ fmt: 'bullet', text: sub(l, ['▪', '◦', '•']) }) },
    dash: { fmt: 'bullet', first: '–', level: (l) => ({ fmt: 'bullet', text: sub(l, ['–', '◦', '▪']) }) },
    star: { fmt: 'bullet', first: '★', level: (l) => ({ fmt: 'bullet', text: sub(l, ['★', '◦', '▪']) }) },
    paren: { fmt: 'decimal', first: '%1)', level: (l) => ({ fmt: sub(l, ['decimal', 'lowerLetter', 'lowerRoman']), text: `%${l + 1})` }) },
    letter: { fmt: 'lowerLetter', first: '%1.', level: (l) => ({ fmt: sub(l, ['lowerLetter', 'lowerRoman', 'decimal']), text: `%${l + 1}.` }) },
    LETTER: { fmt: 'upperLetter', first: '%1.', level: (l) => ({ fmt: sub(l, ['upperLetter', 'decimal', 'lowerLetter']), text: `%${l + 1}.` }) },
    roman: { fmt: 'lowerRoman', first: '%1.', level: (l) => ({ fmt: sub(l, ['lowerRoman', 'lowerLetter', 'decimal']), text: `%${l + 1}.` }) },
    ROMAN: { fmt: 'upperRoman', first: '%1.', level: (l) => ({ fmt: sub(l, ['upperRoman', 'upperLetter', 'decimal']), text: `%${l + 1}.` }) },
  };

  function sameChain(a, b) { return a.length === b.length && a.every((x, i) => x === b[i]); }

  function makeWrapper(w, doc, ctx) {
    if (typeof w === 'string' && w.startsWith('chg:')) {
      // A tracked change made in Suggesting mode.
      const el = wEl(doc, w.slice(4), { id: nextChangeId(doc), author: changeAuthor(), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z') });
      return { outer: el, inner: el };
    }
    if (typeof w === 'string' && w.startsWith('new:')) {
      const el = wEl(doc, 'hyperlink', { history: 1 });
      el.setAttributeNS(R, 'r:id', ctx.linkRel(w.slice(4)));
      return { outer: el, inner: el };
    }
    const outer = w.cloneNode(false);
    if (w.localName === 'sdt') {
      for (const c of kids(w)) if (!isW(c, 'sdtContent')) outer.append(c.cloneNode(true));
      const content = wEl(doc, 'sdtContent');
      outer.append(content);
      return { outer, inner: content };
    }
    return { outer, inner: outer };
  }

  function appendWithChain(target, node, chain, doc) {
    let parent = target;
    for (const w of chain) { const m = makeWrapper(w, doc, { linkRel: () => '' }); parent.append(m.outer); parent = m.inner; }
    parent.append(node);
  }

  function insertAfterBlock(caretNode, block, doc) {
    let anchor = caretNode;
    const body = doc.getElementsByTagNameNS(W, 'body')[0];
    if (!anchor) anchor = lastBodyBlock(body);
    const parent = anchor ? anchor.parentNode : body;
    parent.insertBefore(block, anchor ? anchor.nextSibling : child(body, 'sectPr'));
    if (isW(block, 'tbl')) parent.insertBefore(wEl(doc, 'p'), block.nextSibling);
  }

  function lastBodyBlock(body) {
    const blocks = kids(body).filter((c) => isW(c, 'p') || isW(c, 'tbl'));
    return blocks[blocks.length - 1] || null;
  }

  function resolve(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const p of target.split('/')) { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); }
    return parts.join('/');
  }

  async function mainDocumentPath(zip) {
    const rels = zip.file('_rels/.rels');
    if (rels) {
      const x = parseXml(await rels.async('string'));
      for (const r of x.getElementsByTagName('Relationship')) {
        if (/\/officeDocument$/.test(r.getAttribute('Type'))) return r.getAttribute('Target').replace(/^\//, '');
      }
    }
    return 'word/document.xml';
  }

  return { Editor, cleanPastedHtml, W, R, wEl, child, kids, isW, parseXml, serialize, LIST_KINDS };
})();
