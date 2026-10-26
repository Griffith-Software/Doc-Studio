'use strict';
// PowerPoint renderer + editor.
//
// Slides are drawn from their own XML with the full inheritance chain (slide -> layout -> master -> theme):
// placeholder positions and text styles, theme colours/fonts, backgrounds, preset and custom shapes, pictures,
// tables, groups, connectors, charts and SmartArt drawings. Editing patches only what changed: edited text
// bodies are rebuilt from their own paragraph/run properties, moved shapes get an explicit position, and slide
// operations update presentation.xml, relationships and content types. Every other part is untouched.

const PptxEditor = (() => {
  const NS = {
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    pkg: 'http://schemas.openxmlformats.org/package/2006/relationships',
    ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
    c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  };
  const REL = {
    slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
    layout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
    image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    notes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
  };
  const EMU_PT = 12700;
  const parseXml = (t) => new DOMParser().parseFromString(t, 'application/xml');
  const serialize = (doc) => { const s = new XMLSerializer().serializeToString(doc); return s.startsWith('<?xml') ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s; };
  const kids = (el, name) => (el ? [...el.childNodes].filter((c) => c.nodeType === 1 && (!name || c.localName === name)) : []);
  const first = (el, name) => kids(el, name)[0] || null;
  const path1 = (el, ...names) => { let x = el; for (const n of names) { x = first(x, n); if (!x) return null; } return x; };
  const num = (el, attr, d) => (el && el.hasAttribute(attr) ? +el.getAttribute(attr) : d);
  const aEl = (doc, name, attrs) => { const e = doc.createElementNS(NS.a, `a:${name}`); for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v); return e; };
  const pEl = (doc, name, attrs) => { const e = doc.createElementNS(NS.p, `p:${name}`); for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v); return e; };

  function resolve(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const p of target.split('/')) { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); }
    return parts.join('/');
  }
  function relative(fromPart, toPath) {
    const a = fromPart.split('/'); a.pop();
    const b = toPath.split('/');
    while (a.length && b.length && a[0] === b[0]) { a.shift(); b.shift(); }
    return '../'.repeat(a.length) + b.join('/');
  }

  // ---------------------------------------------------------------- colours

  function hexToHsl(hex) {
    let r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0; const l = (max + min) / 2;
    if (max !== min) { const d = max - min; s = l > 0.5 ? d / (2 - max - min) : d / (max + min); h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
    return [h, s, l];
  }
  function hslToHex([h, s, l]) {
    const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    let r, g, b;
    if (s === 0) r = g = b = l; else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s; const p = 2 * l - q; r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3); }
    return [r, g, b].map((x) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  const PRESET = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500' };

  // ---------------------------------------------------------------- the deck

  class Deck {
    async load(bytes) {
      this.bytes = bytes;
      this.zip = await JSZip.loadAsync(bytes);
      this.docs = new Map();      // path -> parsed XML
      this.dirty = new Set();     // paths to write back
      this.removed = new Set();
      this.newFiles = new Map();  // path -> bytes
      const pkgRels = await this.rels('');
      const main = pkgRels.find((r) => /\/officeDocument$/.test(r.type));
      this.presPath = main ? main.path : 'ppt/presentation.xml';
      this.pres = await this.doc(this.presPath);
      const sz = this.pres.getElementsByTagNameNS(NS.p, 'sldSz')[0];
      this.cx = num(sz, 'cx', 12192000); this.cy = num(sz, 'cy', 6858000);
      this.defaultTextStyle = this.pres.getElementsByTagNameNS(NS.p, 'defaultTextStyle')[0] || null;
      await this.loadSlides();
      return this;
    }

    async doc(path) {
      if (this.docs.has(path)) return this.docs.get(path);
      const f = this.zip.file(path);
      const d = f ? parseXml(await f.async('string')) : null;
      this.docs.set(path, d);
      return d;
    }
    async rels(part) {
      const relsPath = part ? part.replace(/([^/]+)$/, '_rels/$1.rels') : '_rels/.rels';
      let d = this.docs.get(relsPath);
      if (!d) { const f = this.zip.file(relsPath); d = f ? parseXml(await f.async('string')) : parseXml(`<Relationships xmlns="${NS.pkg}"/>`); this.docs.set(relsPath, d); }
      return [...d.getElementsByTagNameNS(NS.pkg, 'Relationship')].map((r) => ({ id: r.getAttribute('Id'), type: r.getAttribute('Type'), target: r.getAttribute('Target'), mode: r.getAttribute('TargetMode'), path: r.getAttribute('TargetMode') === 'External' ? null : resolve(part, r.getAttribute('Target')), el: r }));
    }
    relsDoc(part) { return this.docs.get(part.replace(/([^/]+)$/, '_rels/$1.rels')); }
    addRel(part, type, target, external) {
      const relsPath = part.replace(/([^/]+)$/, '_rels/$1.rels');
      let d = this.docs.get(relsPath);
      if (!d) { d = parseXml(`<Relationships xmlns="${NS.pkg}"/>`); this.docs.set(relsPath, d); }
      const used = new Set([...d.documentElement.children].map((x) => x.getAttribute('Id')));
      let n = 1; while (used.has(`rId${n}`)) n++;
      const el = d.createElementNS(NS.pkg, 'Relationship');
      el.setAttribute('Id', `rId${n}`); el.setAttribute('Type', type); el.setAttribute('Target', target);
      if (external) el.setAttribute('TargetMode', 'External');
      d.documentElement.append(el);
      this.dirty.add(relsPath);
      return `rId${n}`;
    }
    async bytesOf(path) { if (this.newFiles.has(path)) return this.newFiles.get(path); const f = this.zip.file(path); return f ? f.async('uint8array') : null; }

    async loadSlides() {
      const presRels = await this.rels(this.presPath);
      this.slides = [];
      for (const s of this.pres.getElementsByTagNameNS(NS.p, 'sldId')) {
        const rid = s.getAttributeNS(NS.r, 'id') || s.getAttribute('r:id');
        const rel = presRels.find((r) => r.id === rid);
        if (!rel || !rel.path) continue;
        this.slides.push(await this.loadSlide(rel.path, s));
      }
    }

    async loadSlide(path, sldIdEl) {
      const doc = await this.doc(path);
      const rels = await this.rels(path);
      const layoutPath = (rels.find((r) => r.type === REL.layout) || {}).path;
      const layout = layoutPath ? await this.doc(layoutPath) : null;
      const layoutRels = layoutPath ? await this.rels(layoutPath) : [];
      const masterPath = (layoutRels.find((r) => r.type.endsWith('/slideMaster')) || {}).path;
      const master = masterPath ? await this.doc(masterPath) : null;
      const masterRels = masterPath ? await this.rels(masterPath) : [];
      const themePath = (masterRels.find((r) => r.type.endsWith('/theme')) || {}).path;
      const theme = themePath ? await this.doc(themePath) : null;
      const notesPath = (rels.find((r) => r.type === REL.notes) || {}).path;
      return { path, doc, rels, layoutPath, layout, layoutRels, masterPath, master, masterRels, theme, themePath, notesPath, sldIdEl };
    }

    // ---- theme
    themeInfo(slide) {
      if (slide.themeCache) return slide.themeCache;
      const t = slide.theme;
      const scheme = t && t.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
      const colors = {};
      if (scheme) for (const c of kids(scheme)) { const v = kids(c)[0]; if (v) colors[c.localName] = (v.getAttribute('lastClr') || v.getAttribute('val') || '000000').toUpperCase(); }
      const fontOf = (name) => { const f = t && t.getElementsByTagNameNS(NS.a, name)[0]; const l = f && first(f, 'latin'); return l ? l.getAttribute('typeface') : 'Calibri'; };
      // Colour map: master's clrMap, overridden by the layout or slide.
      const map = {};
      const cm = slide.master && slide.master.getElementsByTagNameNS(NS.p, 'clrMap')[0];
      if (cm) for (const a of cm.attributes) map[a.name] = a.value;
      for (const d of [slide.layout, slide.doc]) {
        const ovr = d && d.getElementsByTagNameNS(NS.p, 'clrMapOvr')[0];
        const om = ovr && ovr.getElementsByTagNameNS(NS.a, 'overrideClrMapping')[0];
        if (om) for (const a of om.attributes) map[a.name] = a.value;
      }
      const fmt = t && t.getElementsByTagNameNS(NS.a, 'fmtScheme')[0];
      slide.themeCache = {
        colors, map, major: fontOf('majorFont'), minor: fontOf('minorFont'),
        fills: fmt ? kids(first(fmt, 'fillStyleLst')) : [], lines: fmt ? kids(first(fmt, 'lnStyleLst')) : [], bgFills: fmt ? kids(first(fmt, 'bgFillStyleLst')) : [],
      };
      return slide.themeCache;
    }

    // Colour element container (e.g. a:solidFill) -> { hex, alpha }.
    color(container, slide, ph) {
      if (!container) return null;
      const c = kids(container).find((x) => /Clr$/.test(x.localName));
      if (!c) return null;
      const th = this.themeInfo(slide);
      let hex = '000000';
      if (c.localName === 'srgbClr') hex = c.getAttribute('val');
      else if (c.localName === 'sysClr') hex = c.getAttribute('lastClr') || '000000';
      else if (c.localName === 'prstClr') hex = PRESET[c.getAttribute('val')] || '000000';
      else if (c.localName === 'schemeClr') {
        let v = c.getAttribute('val');
        if (v === 'phClr') return ph || null;
        if (th.map[v]) v = th.map[v];
        hex = th.colors[v] || ({ bg1: th.colors.lt1, tx1: th.colors.dk1, bg2: th.colors.lt2, tx2: th.colors.dk2 }[v]) || '000000';
      } else if (c.localName === 'scrgbClr') hex = ['r', 'g', 'b'].map((k) => Math.round(num(c, k, 0) / 100000 * 255).toString(16).padStart(2, '0')).join('');
      let hsl = hexToHsl(hex.toUpperCase());
      let alpha = 1;
      for (const m of kids(c)) {
        const v = num(m, 'val', 0) / 100000;
        if (m.localName === 'lumMod') hsl[2] *= v;
        else if (m.localName === 'lumOff') hsl[2] += v;
        else if (m.localName === 'tint') hsl[2] = hsl[2] + (1 - hsl[2]) * (1 - v);
        else if (m.localName === 'shade') hsl[2] *= v;
        else if (m.localName === 'satMod') hsl[1] *= v;
        else if (m.localName === 'alpha') alpha = v;
        hsl[2] = Math.max(0, Math.min(1, hsl[2])); hsl[1] = Math.max(0, Math.min(1, hsl[1]));
      }
      return { hex: '#' + hslToHex(hsl), alpha };
    }
    css(col) { if (!col) return null; if (col.alpha >= 1) return col.hex; const v = [1, 3, 5].map((i) => parseInt(col.hex.slice(i, i + 2), 16)); return `rgba(${v.join(',')},${col.alpha.toFixed(3)})`; }

    // Fill of an spPr-like element: solid, gradient, picture, none. Returns CSS background or null; undefined = not specified.
    async fillCss(spPr, slide, rels, ph) {
      if (!spPr) return undefined;
      if (first(spPr, 'noFill')) return 'transparent';
      const solid = first(spPr, 'solidFill');
      if (solid) return this.css(this.color(solid, slide, ph));
      const grad = first(spPr, 'gradFill');
      if (grad) {
        const stops = kids(first(grad, 'gsLst'), 'gs').map((g) => `${this.css(this.color(g, slide, ph))} ${num(g, 'pos', 0) / 1000}%`);
        const lin = first(grad, 'lin');
        const ang = lin ? num(lin, 'ang', 0) / 60000 + 90 : 180;
        if (first(grad, 'path')) return `radial-gradient(circle, ${stops.join(', ')})`;
        return `linear-gradient(${ang}deg, ${stops.join(', ')})`;
      }
      const blip = first(spPr, 'blipFill');
      if (blip) {
        const url = await this.imageUrl(first(blip, 'blip'), rels);
        if (url) return `center / cover no-repeat url("${url}")`;
      }
      return undefined;
    }

    async imageUrl(blip, rels) {
      if (!blip) return null;
      const rid = blip.getAttributeNS(NS.r, 'embed') || blip.getAttribute('r:embed');
      const rel = rid && rels.find((r) => r.id === rid);
      if (!rel || !rel.path) return null;
      this.urlCache = this.urlCache || new Map();
      if (this.urlCache.has(rel.path)) return this.urlCache.get(rel.path);
      const bytes = await this.bytesOf(rel.path);
      if (!bytes) return null;
      const ext = rel.path.split('.').pop().toLowerCase();
      const type = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf' }[ext] || 'application/octet-stream';
      const url = URL.createObjectURL(new Blob([bytes], { type }));
      this.urlCache.set(rel.path, url);
      return url;
    }

    // ---- placeholders: find the layout and master shapes a slide placeholder inherits from
    phOf(sp) {
      const nv = kids(sp).find((x) => /^nv/.test(x.localName));
      const ph = nv && path1(nv, 'nvPr', 'ph');
      if (!ph) return null;
      return { type: ph.getAttribute('type') || 'obj', idx: ph.getAttribute('idx') };
    }
    findPh(doc, ph) {
      if (!doc || !ph) return null;
      const tree = doc.getElementsByTagNameNS(NS.p, 'spTree')[0];
      const all = tree ? [...tree.getElementsByTagNameNS(NS.p, 'sp')] : [];
      const norm = (t) => ({ ctrTitle: 'title', subTitle: 'body', obj: 'body' }[t] || t);
      return (ph.idx != null && all.find((s) => { const p = this.phOf(s); return p && p.idx === ph.idx; }))
        || all.find((s) => { const p = this.phOf(s); return p && p.type === ph.type; })
        || all.find((s) => { const p = this.phOf(s); return p && norm(p.type) === norm(ph.type); }) || null;
    }
    chain(sp, slide, layer) {
      // [own, layout, master] shape elements for inheritance (layout/master shapes inherit from master only).
      const ph = this.phOf(sp);
      if (!ph) return [sp];
      if (layer === 'master') return [sp];
      if (layer === 'layout') return [sp, this.findPh(slide.master, ph)].filter(Boolean);
      return [sp, this.findPh(slide.layout, ph), this.findPh(slide.master, ph)].filter(Boolean);
    }
  }

  // ---------------------------------------------------------------- geometry (SVG paths, in shape pixels)

  function presetPath(name, w, h, adj) {
    const a = (i, d) => (adj[i] != null ? adj[i] : d) / 100000;
    const m = Math.min(w, h);
    switch (name) {
      case 'rect': case 'flowChartProcess': case 'textBox': return `M0 0H${w}V${h}H0Z`;
      case 'roundRect': case 'flowChartAlternateProcess': { const r = a(0, 16667) * m; return `M${r} 0H${w - r}A${r} ${r} 0 0 1 ${w} ${r}V${h - r}A${r} ${r} 0 0 1 ${w - r} ${h}H${r}A${r} ${r} 0 0 1 0 ${h - r}V${r}A${r} ${r} 0 0 1 ${r} 0Z`; }
      case 'ellipse': case 'flowChartConnector': return `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 0 0 ${h / 2}Z`;
      case 'triangle': { const x = a(0, 50000) * w; return `M${x} 0L${w} ${h}H0Z`; }
      case 'rtTriangle': return `M0 0L${w} ${h}H0Z`;
      case 'diamond': case 'flowChartDecision': return `M${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}L0 ${h / 2}Z`;
      case 'parallelogram': { const x = a(0, 25000) * m; return `M${x} 0H${w}L${w - x} ${h}H0Z`; }
      case 'trapezoid': { const x = a(0, 25000) * m; return `M${x} 0H${w - x}L${w} ${h}H0Z`; }
      case 'pentagon': case 'homePlate': { const x = a(0, 50000) * m; return `M0 0H${w - x}L${w} ${h / 2}L${w - x} ${h}H0Z`; }
      case 'chevron': { const x = a(0, 50000) * m; return `M0 0H${w - x}L${w} ${h / 2}L${w - x} ${h}H0L${x} ${h / 2}Z`; }
      case 'hexagon': { const x = a(0, 25000) * m; return `M${x} 0H${w - x}L${w} ${h / 2}L${w - x} ${h}H${x}L0 ${h / 2}Z`; }
      case 'octagon': { const x = a(0, 29289) * m; return `M${x} 0H${w - x}L${w} ${x}V${h - x}L${w - x} ${h}H${x}L0 ${h - x}V${x}Z`; }
      case 'rightArrow': { const t = a(0, 50000) * h, hx = a(1, 50000) * m; return `M0 ${(h - t) / 2}H${w - hx}V0L${w} ${h / 2}L${w - hx} ${h}V${(h + t) / 2}H0Z`; }
      case 'leftArrow': { const t = a(0, 50000) * h, hx = a(1, 50000) * m; return `M${w} ${(h - t) / 2}H${hx}V0L0 ${h / 2}L${hx} ${h}V${(h + t) / 2}H${w}Z`; }
      case 'upArrow': { const t = a(0, 50000) * w, hy = a(1, 50000) * m; return `M${(w - t) / 2} ${h}V${hy}H0L${w / 2} 0L${w} ${hy}H${(w + t) / 2}V${h}Z`; }
      case 'downArrow': { const t = a(0, 50000) * w, hy = a(1, 50000) * m; return `M${(w - t) / 2} 0V${h - hy}H0L${w / 2} ${h}L${w} ${h - hy}H${(w + t) / 2}V0Z`; }
      case 'star5': { const pts = []; for (let i = 0; i < 10; i++) { const r = i % 2 ? 0.38 : 0.5; const ang = -Math.PI / 2 + i * Math.PI / 5; pts.push(`${w / 2 + Math.cos(ang) * r * w} ${h / 2 + Math.sin(ang) * r * h}`); } return `M${pts.join('L')}Z`; }
      case 'plus': { const x = a(0, 25000) * m; return `M${x} 0H${w - x}V${x}H${w}V${h - x}H${w - x}V${h}H${x}V${h - x}H0V${x}H${x}Z`; }
      case 'line': case 'straightConnector1': case 'bentConnector3': case 'curvedConnector3': return `M0 0L${w} ${h}`;
      case 'snip1Rect': { const x = a(0, 16667) * m; return `M0 0H${w - x}L${w} ${x}V${h}H0Z`; }
      case 'round2SameRect': { const r = a(0, 16667) * m; return `M${r} 0H${w - r}A${r} ${r} 0 0 1 ${w} ${r}V${h}H0V${r}A${r} ${r} 0 0 1 ${r} 0Z`; }
      case 'flowChartTerminator': { const r = h / 2; return `M${r} 0H${w - r}A${r} ${r} 0 0 1 ${w - r} ${h}H${r}A${r} ${r} 0 0 1 ${r} 0Z`; }
      case 'wedgeRectCallout': case 'wedgeRoundRectCallout': return `M0 0H${w}V${h * 0.8}H${w * 0.4}L${w * 0.2} ${h}V${h * 0.8}H0Z`;
      case 'cloud': case 'cloudCallout': return `M${w * 0.2} ${h * 0.4}A${w * 0.2} ${h * 0.25} 0 0 1 ${w * 0.5} ${h * 0.15}A${w * 0.2} ${h * 0.25} 0 0 1 ${w * 0.85} ${h * 0.35}A${w * 0.18} ${h * 0.25} 0 0 1 ${w * 0.8} ${h * 0.85}A${w * 0.3} ${h * 0.2} 0 0 1 ${w * 0.2} ${h * 0.85}A${w * 0.18} ${h * 0.22} 0 0 1 ${w * 0.2} ${h * 0.4}Z`;
      case 'donut': { const r = a(0, 25000) * m; return `${presetPath('ellipse', w, h, [])}M${r} ${h / 2}A${w / 2 - r} ${h / 2 - r} 0 1 1 ${w - r} ${h / 2}A${w / 2 - r} ${h / 2 - r} 0 1 1 ${r} ${h / 2}Z`; }
      case 'heart': return `M${w / 2} ${h * 0.25}C${w * 0.5} 0 0 0 0 ${h * 0.3}C0 ${h * 0.6} ${w * 0.5} ${h * 0.8} ${w / 2} ${h}C${w * 0.5} ${h * 0.8} ${w} ${h * 0.6} ${w} ${h * 0.3}C${w} 0 ${w * 0.5} 0 ${w / 2} ${h * 0.25}Z`;
      default: return `M0 0H${w}V${h}H0Z`;
    }
  }

  function customPath(custGeom, w, h) {
    let d = '';
    for (const path of kids(path1(custGeom, 'pathLst'), 'path')) {
      const pw = num(path, 'w', 0) || w, ph = num(path, 'h', 0) || h;
      const sx = w / pw, sy = h / ph;
      let cx = 0, cy = 0;
      const pt = (el) => { const p = first(el, 'pt'); return [num(p, 'x', 0) * sx, num(p, 'y', 0) * sy]; };
      for (const c of kids(path)) {
        if (c.localName === 'moveTo') { [cx, cy] = pt(c); d += `M${cx} ${cy}`; }
        else if (c.localName === 'lnTo') { [cx, cy] = pt(c); d += `L${cx} ${cy}`; }
        else if (c.localName === 'cubicBezTo') { const p = kids(c, 'pt').map((q) => `${num(q, 'x', 0) * sx} ${num(q, 'y', 0) * sy}`); d += `C${p.join(' ')}`; const last = kids(c, 'pt').pop(); cx = num(last, 'x', 0) * sx; cy = num(last, 'y', 0) * sy; }
        else if (c.localName === 'quadBezTo') { const p = kids(c, 'pt').map((q) => `${num(q, 'x', 0) * sx} ${num(q, 'y', 0) * sy}`); d += `Q${p.join(' ')}`; const last = kids(c, 'pt').pop(); cx = num(last, 'x', 0) * sx; cy = num(last, 'y', 0) * sy; }
        else if (c.localName === 'arcTo') {
          const wR = num(c, 'wR', 0) * sx, hR = num(c, 'hR', 0) * sy, st = num(c, 'stAng', 0) / 60000 * Math.PI / 180, sw = num(c, 'swAng', 0) / 60000 * Math.PI / 180;
          const ox = cx - wR * Math.cos(st), oy = cy - hR * Math.sin(st);
          const ex = ox + wR * Math.cos(st + sw), ey = oy + hR * Math.sin(st + sw);
          d += `A${wR} ${hR} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${ex} ${ey}`;
          cx = ex; cy = ey;
        } else if (c.localName === 'close') d += 'Z';
      }
    }
    return d;
  }

  // ======================================================================
  // Rendering
  // ======================================================================

  class Renderer {
    constructor(deck, opts) { this.deck = deck; this.opts = opts || {}; }

    async slide(slide, widthPx, editable) {
      const deck = this.deck;
      const k = widthPx / deck.cx; // px per EMU
      const el = document.createElement('div');
      el.className = 'pp-slide';
      el.style.width = `${widthPx}px`;
      el.style.height = `${deck.cy * k}px`;
      el.style.setProperty('--pp-k', k);
      this.k = k;
      this.editable = editable;
      // Background: slide -> layout -> master.
      let bg = null;
      for (const d of [slide.doc, slide.layout, slide.master]) {
        const b = d && d.getElementsByTagNameNS(NS.p, 'bg')[0];
        if (!b) continue;
        const bgPr = first(b, 'bgPr');
        const rels = d === slide.doc ? slide.rels : d === slide.layout ? slide.layoutRels : slide.masterRels;
        if (bgPr) bg = await deck.fillCss(bgPr, slide, rels);
        else {
          const ref = first(b, 'bgRef');
          if (ref) {
            const idx = num(ref, 'idx', 0);
            const th = deck.themeInfo(slide);
            const style = idx >= 1001 ? th.bgFills[idx - 1001] : th.fills[idx - 1];
            const phc = deck.color(ref, slide);
            // Theme fill styles use phClr, which takes the colour given by the reference.
            if (style) bg = await deck.fillCss({ childNodes: [style] }, slide, rels, phc);
            if (!bg && phc) bg = deck.css(phc);
          }
        }
        if (bg) break;
      }
      el.style.background = bg || '#FFFFFF';
      // Master and layout decorations (non-placeholder shapes), then the slide's shapes.
      const show = (d) => !d || !d.documentElement.hasAttribute('showMasterSp') || d.documentElement.getAttribute('showMasterSp') !== '0';
      if (show(slide.doc) && show(slide.layout) && slide.master) await this.tree(el, slide.master.getElementsByTagNameNS(NS.p, 'spTree')[0], slide, slide.masterRels, 'master');
      if (show(slide.doc) && slide.layout) await this.tree(el, slide.layout.getElementsByTagNameNS(NS.p, 'spTree')[0], slide, slide.layoutRels, 'layout');
      await this.tree(el, slide.doc.getElementsByTagNameNS(NS.p, 'spTree')[0], slide, slide.rels, 'slide');
      return el;
    }

    async tree(host, tree, slide, rels, layer, group) {
      if (!tree) return;
      let i = 0;
      for (const node of kids(tree)) {
        const ln = node.localName;
        if (!['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'AlternateContent'].includes(ln)) continue;
        const shapeNode = ln === 'AlternateContent' ? (first(first(node, 'Choice'), 'sp') || first(first(node, 'Fallback'), 'sp') || first(first(node, 'Fallback'), 'pic')) : node;
        if (!shapeNode) continue;
        if (layer !== 'slide' && this.deck.phOf(shapeNode)) continue; // layout/master placeholders are templates, not content
        const el = await this.shape(shapeNode, slide, rels, layer, group);
        if (el) {
          if (layer === 'slide' && !group) { el.dataset.idx = i; el.classList.add('pp-own'); }
          host.append(el);
        }
        i++;
      }
    }

    xfrm(chain) {
      for (const s of chain) {
        const spPr = kids(s).find((x) => x.localName === 'spPr' || x.localName === 'grpSpPr' || x.localName === 'xfrm');
        const x = spPr && (spPr.localName === 'xfrm' ? spPr : first(spPr, 'xfrm'));
        if (x && first(x, 'off')) return { x: num(first(x, 'off'), 'x', 0), y: num(first(x, 'off'), 'y', 0), w: num(first(x, 'ext'), 'cx', 0), h: num(first(x, 'ext'), 'cy', 0), rot: num(x, 'rot', 0) / 60000, flipH: x.getAttribute('flipH') === '1', flipV: x.getAttribute('flipV') === '1', el: x };
      }
      return null;
    }

    place(el, xf, group) {
      const k = this.k;
      let { x, y, w, h } = xf;
      if (group) { x = group.ox + (x - group.chx) * group.sx; y = group.oy + (y - group.chy) * group.sy; w *= group.sx; h *= group.sy; }
      Object.assign(el.style, { left: `${x * k}px`, top: `${y * k}px`, width: `${w * k}px`, height: `${h * k}px` });
      const t = [];
      if (xf.rot) t.push(`rotate(${xf.rot}deg)`);
      if (xf.flipH) t.push('scaleX(-1)');
      if (xf.flipV) t.push('scaleY(-1)');
      if (t.length) el.style.transform = t.join(' ');
      return { w: w * k, h: h * k };
    }

    async shape(node, slide, rels, layer, group) {
      const deck = this.deck;
      const ln = node.localName;
      const chain = ln === 'sp' ? deck.chain(node, slide, layer) : [node];
      const xf = this.xfrm(chain);
      if (ln === 'grpSp') {
        if (!xf) return null;
        const g = first(first(node, 'grpSpPr'), 'xfrm');
        const chOff = first(g, 'chOff'), chExt = first(g, 'chExt');
        const sub = { ox: xf.x, oy: xf.y, chx: num(chOff, 'x', 0), chy: num(chOff, 'y', 0), sx: xf.w / (num(chExt, 'cx', 0) || xf.w || 1), sy: xf.h / (num(chExt, 'cy', 0) || xf.h || 1) };
        if (group) { sub.ox = group.ox + (xf.x - group.chx) * group.sx; sub.oy = group.oy + (xf.y - group.chy) * group.sy; sub.sx *= group.sx; sub.sy *= group.sy; }
        const frag = document.createElement('div');
        frag.className = 'pp-group';
        await this.tree(frag, node, slide, rels, layer, sub);
        // Group children are positioned against the slide, so unwrap them into a zero-offset container.
        frag.style.cssText = 'position:absolute;left:0;top:0;width:0;height:0;';
        if (xf.rot) frag.style.transform = `rotate(${xf.rot}deg)`;
        return frag;
      }
      if (!xf) return null;
      const el = document.createElement('div');
      el.className = `pp-shape pp-${ln}`;
      const box = this.place(el, xf, group);
      const spPr = first(node, 'spPr');
      const style = first(node, 'style');
      const th = deck.themeInfo(slide);
      if (ln === 'pic') {
        const blipFill = first(node, 'blipFill');
        const url = await deck.imageUrl(first(blipFill, 'blip'), rels);
        if (url) {
          const img = document.createElement('img');
          img.src = url; img.draggable = false;
          const src = first(blipFill, 'srcRect');
          if (src) {
            const l = num(src, 'l', 0) / 100000, t = num(src, 't', 0) / 100000, r = num(src, 'r', 0) / 100000, b = num(src, 'b', 0) / 100000;
            Object.assign(img.style, { position: 'absolute', width: `${100 / (1 - l - r)}%`, height: `${100 / (1 - t - b)}%`, left: `${-l / (1 - l - r) * 100}%`, top: `${-t / (1 - t - b) * 100}%`, maxWidth: 'none' });
          }
          el.append(img);
          const geom = first(spPr, 'prstGeom');
          if (geom && geom.getAttribute('prst') === 'ellipse') el.style.borderRadius = '50%';
        }
        return el;
      }
      if (ln === 'graphicFrame') {
        const data = node.getElementsByTagNameNS(NS.a, 'graphicData')[0];
        const tbl = data && first(data, 'tbl');
        if (tbl) { el.append(await this.table(tbl, slide, box)); el.classList.add('pp-table'); return el; }
        const chart = data && data.getElementsByTagNameNS(NS.c, 'chart')[0];
        if (chart) {
          const rid = chart.getAttributeNS(NS.r, 'id') || chart.getAttribute('r:id');
          const rel = rels.find((r) => r.id === rid);
          if (rel && rel.path) { const cdoc = await deck.doc(rel.path); if (cdoc) this.drawChart(el, cdoc, slide, box); }
          return el;
        }
        const dgm = data && data.getAttribute('uri') && /diagram/.test(data.getAttribute('uri'));
        if (dgm) {
          const drawingRel = rels.find((r) => /diagramDrawing$/.test(r.type));
          const ddoc = drawingRel && drawingRel.path ? await deck.doc(drawingRel.path) : null;
          const dtree = ddoc && ddoc.getElementsByTagNameNS('http://schemas.microsoft.com/office/drawing/2008/diagram', 'spTree')[0];
          if (dtree) {
            const sub = { ox: xf.x, oy: xf.y, chx: 0, chy: 0, sx: 1, sy: 1 };
            const holder = document.createElement('div');
            holder.style.cssText = 'position:absolute;left:0;top:0;width:0;height:0;';
            await this.tree(holder, dtree, slide, rels.concat(await deck.rels(drawingRel.path)), 'diagram', sub);
            return holder;
          }
        }
        el.classList.add('pp-unknown');
        return el;
      }
      // Shape (sp / cxnSp): geometry, fill, line, then text.
      const geomEl = chain.map((s) => first(first(s, 'spPr'), 'prstGeom') || first(first(s, 'spPr'), 'custGeom')).find(Boolean);
      let fill;
      for (const s of chain) { fill = await deck.fillCss(first(s, 'spPr'), slide, rels); if (fill !== undefined) break; }
      const refColor = (name) => { const r = style && first(style, name); return r ? deck.color(r, slide) : null; };
      if (fill === undefined && style) {
        const fr = first(style, 'fillRef');
        const idx = num(fr, 'idx', 0);
        if (idx > 0) { const st = th.fills[idx - 1]; const phc = refColor('fillRef'); fill = (st && await deck.fillCss({ childNodes: [st] }, slide, rels, phc)) || deck.css(phc); }
      }
      let line = null;
      for (const s of chain) {
        const lnEl = first(first(s, 'spPr'), 'ln');
        if (!lnEl) continue;
        if (first(lnEl, 'noFill')) { line = { none: true }; break; }
        const c = first(lnEl, 'solidFill');
        line = { color: c ? deck.css(deck.color(c, slide)) : null, w: num(lnEl, 'w', 9525), dash: first(lnEl, 'prstDash') ? first(lnEl, 'prstDash').getAttribute('val') : 'solid', head: first(lnEl, 'headEnd'), tail: first(lnEl, 'tailEnd') };
        break;
      }
      if (style && (!line || (!line.none && !line.color))) {
        const lr = first(style, 'lnRef');
        const idx = num(lr, 'idx', 0);
        if (idx > 0) { const st = th.lines[idx - 1]; const phc = refColor('lnRef'); line = { ...(line || {}), color: (line && line.color) || deck.css(phc), w: (line && line.w) || num(st, 'w', 12700) }; }
      }
      const isLine = ln === 'cxnSp' || (geomEl && /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(geomEl.getAttribute('prst') || ''));
      if (geomEl || fill || line) {
        const svgNS = 'http://www.w3.org/2000/svg';
        const svg = document.createElementNS(svgNS, 'svg');
        svg.setAttribute('class', 'pp-geom');
        const W = Math.max(1, box.w), H = Math.max(1, box.h);
        svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
        svg.setAttribute('preserveAspectRatio', 'none');
        const path = document.createElementNS(svgNS, 'path');
        const adj = geomEl ? kids(first(geomEl, 'avLst'), 'gd').map((g) => +((g.getAttribute('fmla') || '').split(' ')[1] || 0)) : [];
        const d = geomEl && geomEl.localName === 'custGeom' ? customPath(geomEl, W, H) : presetPath(geomEl ? geomEl.getAttribute('prst') : 'rect', W, H, adj);
        path.setAttribute('d', d);
        const fillCss = isLine ? 'none' : (fill && fill !== 'transparent' ? fill : 'none');
        if (/gradient|url\(/.test(fillCss)) { el.style.background = fillCss; el.style.clipPath = `path('${d}')`; path.setAttribute('fill', 'none'); }
        else path.setAttribute('fill', fillCss);
        if (line && !line.none && (line.color || isLine)) {
          path.setAttribute('stroke', line.color || '#000');
          path.setAttribute('stroke-width', Math.max(0.75, (line.w || 9525) * this.k));
          if (line.dash && line.dash !== 'solid') path.setAttribute('stroke-dasharray', { dash: '6 4', sysDash: '4 3', dot: '1 3', sysDot: '1 2', dashDot: '6 3 1 3', lgDash: '10 4' }[line.dash] || '4 3');
          if (line.tail && line.tail.getAttribute('type') && line.tail.getAttribute('type') !== 'none') path.setAttribute('marker-end', 'url(#pp-arrow)');
          path.setAttribute('vector-effect', 'non-scaling-stroke');
        }
        svg.append(path);
        if (line && line.tail && line.tail.getAttribute('type') && line.tail.getAttribute('type') !== 'none') {
          svg.insertAdjacentHTML('afterbegin', `<defs><marker id="pp-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="${line.color || '#000'}"/></marker></defs>`);
        }
        el.append(svg);
      }
      const tx = first(node, 'txBody');
      const fontColor = style && first(style, 'fontRef') ? deck.color(first(style, 'fontRef'), slide) : null;
      const fontRefIdx = style && first(style, 'fontRef') ? first(style, 'fontRef').getAttribute('idx') : null;
      if (tx || (layer === 'slide' && this.editable && this.deck.phOf(node))) {
        el.append(this.text(tx, chain, slide, layer, { color: fontColor, font: fontRefIdx === 'major' ? th.major : fontRefIdx === 'minor' ? th.minor : null }, node));
      }
      return el;
    }

    // ---- text

    levelProps(chain, slide, layer, lvl, node) {
      // Most general first: presentation default -> master txStyles -> master ph -> layout ph -> own lstStyle.
      const deck = this.deck;
      const ph = deck.phOf(node);
      const out = [];
      const lvlName = `lvl${lvl + 1}pPr`;
      if (!ph && deck.defaultTextStyle) { const d = first(deck.defaultTextStyle, lvlName); if (d) out.push(d); }
      const tx = slide.master && slide.master.getElementsByTagNameNS(NS.p, 'txStyles')[0];
      if (tx) {
        const kind = !ph ? 'otherStyle' : /title|ctrTitle/.test(ph.type) ? 'titleStyle' : /body|subTitle|obj/.test(ph.type) ? 'bodyStyle' : 'otherStyle';
        const st = first(tx, kind);
        const d = st && first(st, lvlName);
        if (d) out.push(d);
      }
      for (const s of [...chain].reverse()) {
        const ls = path1(s, 'txBody', 'lstStyle');
        const d = ls && first(ls, lvlName);
        if (d) out.push(d);
      }
      return out;
    }

    text(tx, chain, slide, layer, styleRef, node) {
      const deck = this.deck;
      const th = deck.themeInfo(slide);
      const k = this.k;
      const box = document.createElement('div');
      box.className = 'pp-text';
      // Body properties: own -> layout -> master.
      const bodies = chain.map((s) => path1(s, 'txBody', 'bodyPr')).filter(Boolean);
      const bp = (attr, d) => { for (const b of bodies) if (b.hasAttribute(attr)) return b.getAttribute(attr); return d; };
      const ins = (n, d) => +bp(n, d) * k;
      box.style.padding = `${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px`;
      const anchor = bp('anchor', 't');
      box.style.justifyContent = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
      if (bp('wrap', 'square') === 'none') box.style.whiteSpace = 'nowrap';
      const vert = bp('vert', 'horz');
      if (vert === 'vert' || vert === 'eaVert') box.style.writingMode = 'vertical-rl';
      if (vert === 'vert270') { box.style.writingMode = 'vertical-rl'; box.style.transform = 'rotate(180deg)'; }
      let fontScale = 1, spcRed = 0;
      for (const b of bodies) { const na = first(b, 'normAutofit'); if (na) { fontScale = num(na, 'fontScale', 100000) / 100000; spcRed = num(na, 'lnSpcReduction', 0) / 100000; break; } }
      const ph = deck.phOf(node);
      const paras = tx ? kids(tx, 'p') : [];
      const autoNum = {};
      const empty = !paras.some((p) => p.textContent.length);
      paras.forEach((p, pi) => {
        const pPr = first(p, 'pPr');
        const lvl = num(pPr, 'lvl', 0);
        const levels = this.levelProps(chain, slide, layer, lvl, node);
        const all = [...levels, pPr].filter(Boolean);
        const attr = (name, d) => { for (let i = all.length - 1; i >= 0; i--) if (all[i].hasAttribute(name)) return all[i].getAttribute(name); return d; };
        const childOf = (name) => { for (let i = all.length - 1; i >= 0; i--) { const c = first(all[i], name); if (c) return c; } return null; };
        const defR = levels.map((l) => first(l, 'defRPr')).filter(Boolean);
        const div = document.createElement('div');
        div.className = 'pp-p';
        div.dataset.p = pi;
        const algn = attr('algn', 'l');
        div.style.textAlign = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[algn] || 'left';
        const marL = +attr('marL', 0) * k, indent = +attr('indent', 0) * k;
        div.style.paddingLeft = `${marL}px`;
        div.style.textIndent = `${indent}px`;
        const spc = (name) => { const c = childOf(name); if (!c) return null; const pct = first(c, 'spcPct'), pts = first(c, 'spcPts'); return pct ? { pct: num(pct, 'val', 100000) / 100000 } : pts ? { pt: num(pts, 'val', 0) / 100 } : null; };
        const lnSpc = spc('lnSpc');
        const sizeOf = (rPr) => { for (const r of [rPr, ...[...defR].reverse()]) if (r && r.hasAttribute('sz')) return +r.getAttribute('sz') / 100; return 18; };
        const firstR = first(first(p, 'r'), 'rPr') || first(p, 'endParaRPr');
        const baseSize = sizeOf(firstR) * fontScale;
        const pxPt = k * EMU_PT;
        if (lnSpc && lnSpc.pct) div.style.lineHeight = `${Math.max(0.6, 1.2 * lnSpc.pct * (1 - spcRed))}`;
        else if (lnSpc && lnSpc.pt) div.style.lineHeight = `${lnSpc.pt * pxPt}px`;
        else div.style.lineHeight = `${1.2 * (1 - spcRed)}`;
        const bef = spc('spcBef'), aft = spc('spcAft');
        if (bef) div.style.marginTop = `${bef.pt != null ? bef.pt * pxPt : bef.pct * baseSize * pxPt}px`;
        if (aft) div.style.marginBottom = `${aft.pt != null ? aft.pt * pxPt : aft.pct * baseSize * pxPt}px`;
        // Bullets
        const hasText = p.textContent.length > 0;
        const buNone = childOf('buNone'), buChar = childOf('buChar'), buAuto = childOf('buAutoNum');
        const lastBu = [buNone, buChar, buAuto].filter(Boolean).sort((a, b) => all.indexOf(a.parentNode) - all.indexOf(b.parentNode)).pop();
        if (hasText && lastBu && lastBu !== buNone) {
          const bu = document.createElement('span');
          bu.className = 'pp-bullet';
          bu.contentEditable = 'false';
          if (lastBu.localName === 'buChar') bu.textContent = lastBu.getAttribute('char') || '•';
          else { autoNum[lvl] = (autoNum[lvl] || num(lastBu, 'startAt', 1) - 1) + 1; bu.textContent = formatAuto(lastBu.getAttribute('type'), autoNum[lvl]); }
          const buFont = childOf('buFont');
          if (buFont && buFont.getAttribute('typeface')) bu.style.fontFamily = `"${buFont.getAttribute('typeface')}"`;
          const buClr = childOf('buClr');
          if (buClr) bu.style.color = deck.css(deck.color(buClr, slide));
          const buSz = childOf('buSzPct');
          bu.style.fontSize = `${baseSize * pxPt * (buSz ? num(buSz, 'val', 100000) / 100000 : 1)}px`;
          bu.style.minWidth = `${Math.max(0, -indent)}px`;
          div.append(bu);
        }
        if (!hasText) for (const k2 of Object.keys(autoNum)) if (+k2 >= lvl) delete autoNum[k2];
        let ri = 0;
        for (const c of kids(p)) {
          if (c.localName === 'r' || c.localName === 'fld') {
            const span = document.createElement('span');
            span.className = 'pp-r';
            span.dataset.r = ri;
            this.runStyle(span, first(c, 'rPr'), defR, slide, th, styleRef, fontScale, ph);
            let t = first(c, 't') ? first(c, 't').textContent : '';
            if (c.localName === 'fld') {
              span.contentEditable = 'false';
              span.dataset.fld = '1';
              if ((c.getAttribute('type') || '').startsWith('slidenum')) t = String(this.opts.slideNumber || t);
            }
            span.textContent = t;
            div.append(span);
          } else if (c.localName === 'br') {
            const br = document.createElement('br');
            br.dataset.r = ri;
            div.append(br);
          }
          ri++;
        }
        // Bullets take the colour and font size of the paragraph's first run unless they have their own.
        const bullet = div.querySelector('.pp-bullet');
        const firstRun = div.querySelector('.pp-r');
        if (bullet && firstRun && !bullet.style.color) bullet.style.color = firstRun.style.color;
        // Empty paragraphs keep their height.
        const endR = first(p, 'endParaRPr');
        if (!hasText) { const s = document.createElement('span'); s.className = 'pp-r pp-empty'; this.runStyle(s, endR, defR, slide, th, styleRef, fontScale, ph); s.innerHTML = '<br>'; div.append(s); }
        box.append(div);
      });
      // Placeholder prompts while editing (not saved).
      if (empty && ph && this.editable && layer === 'slide') {
        box.dataset.prompt = ph.type === 'subTitle' ? 'Click to add subtitle' : /title/i.test(ph.type) ? 'Click to add title' : /pic/.test(ph.type) ? 'Picture placeholder' : 'Click to add text';
        box.classList.add('pp-prompt');
      }
      return box;
    }

    runStyle(span, rPr, defR, slide, th, styleRef, fontScale, ph) {
      const deck = this.deck;
      const all = [...defR, rPr].filter(Boolean);
      const attr = (name, d) => { for (let i = all.length - 1; i >= 0; i--) if (all[i].hasAttribute(name)) return all[i].getAttribute(name); return d; };
      const childOf = (name) => { for (let i = all.length - 1; i >= 0; i--) { const c = first(all[i], name); if (c) return c; } return null; };
      const pxPt = this.k * EMU_PT;
      const sz = +attr('sz', 1800) / 100 * fontScale;
      span.style.fontSize = `${sz * pxPt}px`;
      if (attr('b', '0') === '1') span.style.fontWeight = '700';
      if (attr('i', '0') === '1') span.style.fontStyle = 'italic';
      const u = attr('u', 'none'), st = attr('strike', 'noStrike');
      const deco = [u !== 'none' ? 'underline' : '', st !== 'noStrike' ? 'line-through' : ''].filter(Boolean).join(' ');
      if (deco) span.style.textDecoration = deco;
      // The run's own colour wins, then the shape's style colour, then inherited list-style defaults.
      const ownFill = rPr && first(rPr, 'solidFill');
      const inherited = childOf('solidFill');
      const color = ownFill ? deck.color(ownFill, slide) : styleRef.color || (inherited ? deck.color(inherited, slide) : null);
      if (color) span.style.color = deck.css(color);
      else {
        // Text defaults to the theme's tx1 colour.
        const th2 = deck.themeInfo(slide);
        const map = th2.map.tx1 || 'dk1';
        span.style.color = '#' + (th2.colors[map] || '000000');
      }
      const latin = childOf('latin');
      let face = latin ? latin.getAttribute('typeface') : null;
      if (!face) face = styleRef.font || (ph && /title|ctrTitle/.test(ph.type) ? th.major : th.minor);
      if (face === '+mj-lt') face = th.major;
      if (face === '+mn-lt') face = th.minor;
      span.style.fontFamily = `"${face}", "${th.minor}", Calibri, sans-serif`;
      const base = +attr('baseline', 0);
      if (base) { span.style.verticalAlign = base > 0 ? 'super' : 'sub'; span.style.fontSize = `${sz * pxPt * 0.7}px`; }
      const cap = attr('cap', 'none');
      if (cap === 'all') span.style.textTransform = 'uppercase';
      if (cap === 'small') span.style.fontVariant = 'small-caps';
      const spc = +attr('spc', 0);
      if (spc) span.style.letterSpacing = `${spc / 100 * pxPt}px`;
      const hl = childOf('highlight');
      if (hl) span.style.background = deck.css(deck.color(hl, slide));
    }

    async table(tbl, slide, box) {
      const deck = this.deck;
      const k = this.k;
      const th = deck.themeInfo(slide);
      const table = document.createElement('table');
      const tblPr = first(tbl, 'tblPr');
      const firstRow = tblPr && tblPr.getAttribute('firstRow') === '1', bandRow = tblPr && tblPr.getAttribute('bandRow') === '1';
      const styleId = tblPr && first(tblPr, 'tableStyleId') ? first(tblPr, 'tableStyleId').textContent : null;
      const accent = '#' + (th.colors.accent1 || '4472C4');
      const tint = (hex, t) => '#' + [1, 3, 5].map((i) => Math.round(parseInt(hex.slice(i, i + 2), 16) + (255 - parseInt(hex.slice(i, i + 2), 16)) * t).toString(16).padStart(2, '0')).join('');
      const styled = !!styleId;
      const cols = kids(first(tbl, 'tblGrid'), 'gridCol').map((g) => num(g, 'w', 0) * k);
      const cg = document.createElement('colgroup');
      for (const w of cols) { const c = document.createElement('col'); c.style.width = `${w}px`; cg.append(c); }
      table.append(cg);
      kids(tbl, 'tr').forEach((tr, ri) => {
        const row = document.createElement('tr');
        row.style.height = `${num(tr, 'h', 0) * k}px`;
        for (const tc of kids(tr, 'tc')) {
          if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') continue;
          const td = document.createElement('td');
          if (tc.hasAttribute('gridSpan')) td.colSpan = +tc.getAttribute('gridSpan');
          if (tc.hasAttribute('rowSpan')) td.rowSpan = +tc.getAttribute('rowSpan');
          const tcPr = first(tc, 'tcPr');
          const header = firstRow && ri === 0;
          let bg = tcPr ? (first(tcPr, 'noFill') ? 'transparent' : first(tcPr, 'solidFill') ? deck.css(deck.color(first(tcPr, 'solidFill'), slide)) : null) : null;
          if (!bg && styled) bg = header ? accent : bandRow ? (ri % 2 ? tint(accent, 0.8) : tint(accent, 0.9)) : tint(accent, 0.9);
          if (bg) td.style.background = bg;
          if (styled) td.style.border = '1px solid #fff';
          for (const [side, name] of [['Left', 'lnL'], ['Right', 'lnR'], ['Top', 'lnT'], ['Bottom', 'lnB']]) {
            const l = tcPr && first(tcPr, name);
            if (l && !first(l, 'noFill') && first(l, 'solidFill')) td.style[`border${side}`] = `${Math.max(1, num(l, 'w', 12700) * k)}px solid ${deck.css(deck.color(first(l, 'solidFill'), slide))}`;
          }
          const textEl = this.text(first(tc, 'txBody'), [tc], slide, 'slide', { color: header && styled ? { hex: '#FFFFFF', alpha: 1 } : null, font: null }, tc);
          if (header && styled) textEl.querySelectorAll('.pp-r').forEach((s) => { s.style.fontWeight = '700'; });
          const mar = (n, d) => num(tcPr, n, d) * k;
          textEl.style.padding = `${mar('marT', 45720)}px ${mar('marR', 91440)}px ${mar('marB', 45720)}px ${mar('marL', 91440)}px`;
          const anchor = tcPr && tcPr.getAttribute('anchor');
          textEl.style.justifyContent = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
          td.append(textEl);
          row.append(td);
        }
        table.append(row);
      });
      table.style.width = `${box.w}px`;
      return table;
    }

    drawChart(el, cdoc, slide) {
      if (typeof Chart === 'undefined' || !SheetView || !SheetView.View) return;
      const P = SheetView.View.prototype;
      const theme = this.deck.themeInfo(slide);
      const fake = {
        wb: { theme: ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((n) => theme.colors[n]) },
        refValues: () => null,
      };
      fake.drawingColor = P.drawingColor.bind(fake);
      const spec = P.chartSpec.call(fake, cdoc);
      if (!spec) return;
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(10, parseFloat(el.style.width)); canvas.height = Math.max(10, parseFloat(el.style.height));
      el.append(canvas);
      el.classList.add('pp-chart');
      requestAnimationFrame(() => { try { new Chart(canvas, P.chartConfig.call(fake, spec)); } catch (e) { console.warn(e); } });
    }
  }

  function formatAuto(type, n) {
    const roman = (x) => { const m = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; for (const [v, r] of m) while (x >= v) { s += r; x -= v; } return s; };
    const alpha = (x) => String.fromCharCode(96 + ((x - 1) % 26) + 1);
    const t = type || 'arabicPeriod';
    let core = String(n);
    if (/^alphaLc/.test(t)) core = alpha(n); else if (/^alphaUc/.test(t)) core = alpha(n).toUpperCase();
    else if (/^romanLc/.test(t)) core = roman(n); else if (/^romanUc/.test(t)) core = roman(n).toUpperCase();
    if (/ParenBoth$/.test(t)) return `(${core})`;
    if (/ParenR$/.test(t)) return `${core})`;
    if (/Plain$/.test(t)) return core;
    return `${core}.`;
  }

  return { Deck, Renderer, NS, REL, parseXml, serialize, kids, first, path1, num, aEl, pEl, relative, resolve };
})();
