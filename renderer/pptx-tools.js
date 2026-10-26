'use strict';
// Google Slides-style features for the presentation editor (added to PptxUI.Editor):
//   Present: full-screen slideshow with transitions and click animations; presenter view (notes, next slide, timer)
//   Arrange: order, align, distribute, centre on page, rotate/flip, group/ungroup; smart guides while dragging
//   Insert:  table, line, arrow, word art, slide numbers, links; table cell text editing
//   Format:  fill, border colour/weight/dash, shadow, size & rotation
//   Slide:   background, apply layout, transition, skip slide, theme colours and fonts
//   Animate: appear / fade in / fly in on click
//   View:    grid view, zoom

(() => {
  const E = PptxUI.Editor.prototype;
  const { NS, REL, parseXml, serialize, kids, first, aEl, pEl } = PptxEditor;
  const SHAPES = new Set(['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'AlternateContent']);
  // Import an XML fragment; the namespace declarations it was parsed with are already in scope in the target.
  const imp = (doc, xml) => { const n = doc.importNode(parseXml(xml).documentElement, true); for (const a of [...n.attributes]) if (a.name.startsWith('xmlns')) n.removeAttributeNode(a); return n; };
  const NSD = `xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"`;
  const EMU_IN = 914400;
  const hex = (c) => c.replace('#', '').toUpperCase();
  const list = (ed, slide) => kids(ed.tree(slide)).filter((x) => SHAPES.has(x.localName));
  const cur = (ed) => ed.deck.slides[ed.current];

  // ======================================================================
  // Selection helpers
  // ======================================================================

  E.selection = function () {
    const set = this.multi && this.multi.size ? [...this.multi] : this.selected ? [this.selected] : [];
    return set.filter((s) => s.isConnected);
  };
  E.selectedNodes = function () {
    const slide = cur(this);
    return this.selection().map((el) => ({ el, node: list(this, slide)[+el.dataset.idx], idx: +el.dataset.idx })).filter((x) => x.node);
  };
  // Shape boxes in EMU from what is drawn (works for placeholders that inherit their position).
  E.boxOf = function (el) { return this.shapeBox(el); };
  E.redrawKeep = async function (indices) {
    await this.renderCurrent();
    this.refreshThumb(this.current);
    const els = indices.map((i) => this.slideEl.querySelector(`.pp-own[data-idx="${i}"]`)).filter(Boolean);
    els.forEach((el, i) => this.select(el, i > 0));
  };
  const target = (node) => (node.localName === 'AlternateContent' ? (first(first(node, 'Choice'), 'sp') || node) : node);

  // ======================================================================
  // Arrange
  // ======================================================================

  E.order = async function (how) {
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select a shape first.');
    this.snapshot();
    const slide = cur(this);
    const tree = this.tree(slide);
    // Front/forward work from the top-most selected shape down, back/backward from the bottom-most up,
    // so the selected shapes keep their order among themselves.
    const ordered = [...sel].sort((a, b) => a.idx - b.idx);
    for (const { node } of (how === 'front' || how === 'forward' ? [...ordered].reverse() : ordered)) {
      const shapes = list(this, slide);
      const i = shapes.indexOf(node);
      if (how === 'front') tree.append(node);
      else if (how === 'back') { if (shapes[0] !== node) tree.insertBefore(node, shapes[0]); }
      else if (how === 'forward' && shapes[i + 1] && !sel.some((x) => x.node === shapes[i + 1])) shapes[i + 1].after(node);
      else if (how === 'backward' && shapes[i - 1] && !sel.some((x) => x.node === shapes[i - 1])) tree.insertBefore(node, shapes[i - 1]);
    }
    if (how === 'front') for (const { node } of ordered) tree.append(node);
    this.markSlide(slide);
    this.changed();
    const after = list(this, slide);
    await this.redrawKeep(sel.map((x) => after.indexOf(x.node)));
  };

  // Set a shape's box (EMU), keeping rotation/flip attributes.
  function place(ed, node, box) { ed.setXfrm(target(node), box); }

  E.align = async function (how) {
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select shapes to align first.');
    this.snapshot();
    const boxes = sel.map((x) => this.boxOf(x.el));
    // One shape aligns to the slide; several align to each other.
    const ref = sel.length === 1 ? { x: 0, y: 0, w: this.deck.cx, h: this.deck.cy } : {
      x: Math.min(...boxes.map((b) => b.x)), y: Math.min(...boxes.map((b) => b.y)),
      w: Math.max(...boxes.map((b) => b.x + b.w)) - Math.min(...boxes.map((b) => b.x)), h: Math.max(...boxes.map((b) => b.y + b.h)) - Math.min(...boxes.map((b) => b.y)),
    };
    sel.forEach((x, i) => {
      const b = { ...boxes[i] };
      if (how === 'left') b.x = ref.x;
      if (how === 'center') b.x = ref.x + (ref.w - b.w) / 2;
      if (how === 'right') b.x = ref.x + ref.w - b.w;
      if (how === 'top') b.y = ref.y;
      if (how === 'middle') b.y = ref.y + (ref.h - b.h) / 2;
      if (how === 'bottom') b.y = ref.y + ref.h - b.h;
      place(this, x.node, b);
    });
    this.markSlide(cur(this));
    this.changed();
    await this.redrawKeep(sel.map((x) => x.idx));
  };
  E.centerOnPage = async function (axis) {
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select a shape first.');
    this.snapshot();
    for (const x of sel) {
      const b = this.boxOf(x.el);
      if (axis !== 'v') b.x = (this.deck.cx - b.w) / 2;
      if (axis !== 'h') b.y = (this.deck.cy - b.h) / 2;
      place(this, x.node, b);
    }
    this.markSlide(cur(this));
    this.changed();
    await this.redrawKeep(sel.map((x) => x.idx));
  };
  E.distribute = async function (axis) {
    const sel = this.selectedNodes();
    if (sel.length < 3) throw new Error('Select three or more shapes to distribute.');
    this.snapshot();
    const items = sel.map((x) => ({ ...x, b: this.boxOf(x.el) })).sort((a, b) => (axis === 'h' ? a.b.x - b.b.x : a.b.y - b.b.y));
    const start = axis === 'h' ? items[0].b.x : items[0].b.y;
    const lastB = items[items.length - 1].b;
    const end = axis === 'h' ? lastB.x + lastB.w : lastB.y + lastB.h;
    const total = items.reduce((s, x) => s + (axis === 'h' ? x.b.w : x.b.h), 0);
    const gap = (end - start - total) / (items.length - 1);
    let pos = start;
    for (const x of items) { if (axis === 'h') x.b.x = pos; else x.b.y = pos; pos += (axis === 'h' ? x.b.w : x.b.h) + gap; place(this, x.node, x.b); }
    this.markSlide(cur(this));
    this.changed();
    await this.redrawKeep(sel.map((x) => x.idx));
  };
  E.rotate = async function (how) {
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select a shape first.');
    this.snapshot();
    for (const x of sel) {
      place(this, x.node, this.boxOf(x.el));
      const xf = this.xfrmEl(target(x.node), true);
      if (how === 'cw' || how === 'ccw') { const r = (+(xf.getAttribute('rot') || 0) + (how === 'cw' ? 5400000 : -5400000) + 21600000) % 21600000; if (r) xf.setAttribute('rot', r); else xf.removeAttribute('rot'); }
      if (how === 'flipH') { if (xf.getAttribute('flipH') === '1') xf.removeAttribute('flipH'); else xf.setAttribute('flipH', '1'); }
      if (how === 'flipV') { if (xf.getAttribute('flipV') === '1') xf.removeAttribute('flipV'); else xf.setAttribute('flipV', '1'); }
    }
    this.markSlide(cur(this));
    this.changed();
    await this.redrawKeep(sel.map((x) => x.idx));
  };

  E.group = async function () {
    const sel = this.selectedNodes();
    if (sel.length < 2) throw new Error('Select two or more shapes to group (Shift+click to add shapes).');
    if (sel.some((x) => this.deck.phOf(target(x.node)))) throw new Error('Placeholders (title, text boxes from the layout) can\'t be grouped.');
    this.snapshot();
    const slide = cur(this);
    const doc = slide.doc;
    const boxes = sel.map((x) => this.boxOf(x.el));
    sel.forEach((x, i) => place(this, x.node, boxes[i]));
    const x1 = Math.min(...boxes.map((b) => b.x)), y1 = Math.min(...boxes.map((b) => b.y));
    const x2 = Math.max(...boxes.map((b) => b.x + b.w)), y2 = Math.max(...boxes.map((b) => b.y + b.h));
    const id = this.nextId(slide);
    const g = imp(doc, `<p:grpSp ${NSD}><p:nvGrpSpPr><p:cNvPr id="${id}" name="Group ${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${Math.round(x1)}" y="${Math.round(y1)}"/><a:ext cx="${Math.round(x2 - x1)}" cy="${Math.round(y2 - y1)}"/><a:chOff x="${Math.round(x1)}" y="${Math.round(y1)}"/><a:chExt cx="${Math.round(x2 - x1)}" cy="${Math.round(y2 - y1)}"/></a:xfrm></p:grpSpPr></p:grpSp>`);
    const ordered = [...sel].sort((a, b) => a.idx - b.idx);
    const topmost = ordered[ordered.length - 1].node;
    topmost.after(g);
    for (const x of ordered) g.append(x.node);
    this.markSlide(slide);
    this.changed();
    await this.redrawKeep([list(this, slide).indexOf(g)]);
  };
  E.ungroup = async function () {
    const sel = this.selectedNodes().filter((x) => x.node.localName === 'grpSp');
    if (!sel.length) throw new Error('Select a group first.');
    this.snapshot();
    const slide = cur(this);
    for (const { node } of sel) {
      const xf = first(first(node, 'grpSpPr'), 'xfrm');
      const n = (el, a, d) => (el && el.hasAttribute(a) ? +el.getAttribute(a) : d);
      const off = first(xf, 'off'), ext = first(xf, 'ext'), co = first(xf, 'chOff'), ce = first(xf, 'chExt');
      const sx = n(ext, 'cx', 1) / (n(ce, 'cx', 0) || n(ext, 'cx', 1)), sy = n(ext, 'cy', 1) / (n(ce, 'cy', 0) || n(ext, 'cy', 1));
      for (const c of kids(node).filter((k) => SHAPES.has(k.localName))) {
        const cx = this.xfrmEl(target(c), false);
        const o = cx && first(cx, 'off'), e = cx && first(cx, 'ext');
        if (o && e) {
          o.setAttribute('x', Math.round(n(off, 'x', 0) + (n(o, 'x', 0) - n(co, 'x', 0)) * sx));
          o.setAttribute('y', Math.round(n(off, 'y', 0) + (n(o, 'y', 0) - n(co, 'y', 0)) * sy));
          e.setAttribute('cx', Math.round(n(e, 'cx', 0) * sx));
          e.setAttribute('cy', Math.round(n(e, 'cy', 0) * sy));
        }
        node.before(c);
      }
      node.remove();
    }
    this.markSlide(slide);
    this.changed();
    await this.redrawKeep([]);
  };

  // ---------- smart guides while dragging

  E.snapTo = function (shape, l, t) {
    const k = this.k, T = 6;
    const w = parseFloat(shape.style.width), hgt = parseFloat(shape.style.height);
    const W = this.deck.cx * k, H = this.deck.cy * k;
    const xs = [0, W / 2, W], ys = [0, H / 2, H];
    for (const o of this.slideEl.querySelectorAll('.pp-own')) {
      if (o === shape || !o.style.left) continue;
      const ol = parseFloat(o.style.left), ot = parseFloat(o.style.top), ow = parseFloat(o.style.width), oh = parseFloat(o.style.height);
      if (!ow || !oh) continue;
      xs.push(ol, ol + ow / 2, ol + ow); ys.push(ot, ot + oh / 2, ot + oh);
    }
    const best = (pos, size, lines) => {
      let out = null;
      for (const [edge, off] of [[pos, 0], [pos + size / 2, size / 2], [pos + size, size]]) {
        for (const g of lines) { const d = Math.abs(edge - g); if (d <= T && (!out || d < out.d)) out = { d, v: g - off, line: g }; }
      }
      return out;
    };
    const bx = best(l, w, xs), by = best(t, hgt, ys);
    this.clearGuides();
    if (bx) { l = bx.v; this.slideEl.append(Object.assign(document.createElement('i'), { className: 'ppe-guide v', style: `left:${bx.line}px` })); }
    if (by) { t = by.v; this.slideEl.append(Object.assign(document.createElement('i'), { className: 'ppe-guide h', style: `top:${by.line}px` })); }
    return [l, t];
  };
  E.clearGuides = function () { if (this.slideEl) this.slideEl.querySelectorAll('.ppe-guide').forEach((g) => g.remove()); };

  // ======================================================================
  // Insert
  // ======================================================================

  E.insertLine = async function (arrow) {
    const slide = cur(this);
    this.snapshot();
    const id = this.nextId(slide);
    const cx = this.deck.cx, cy = this.deck.cy;
    const xml = `<p:cxnSp ${NSD}><p:nvCxnSpPr><p:cNvPr id="${id}" name="${arrow ? 'Arrow' : 'Line'} ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="${Math.round(cx * 0.35)}" y="${Math.round(cy * 0.5)}"/><a:ext cx="${Math.round(cx * 0.3)}" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="28575"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>${arrow ? '<a:tailEnd type="triangle" w="med" len="med"/>' : ''}</a:ln></p:spPr></p:cxnSp>`;
    this.tree(slide).append(imp(slide.doc, xml));
    this.markSlide(slide);
    await this.afterInsert();
  };
  E.insertWordArt = async function (text) {
    const slide = cur(this);
    this.snapshot();
    const id = this.nextId(slide);
    const cx = this.deck.cx, cy = this.deck.cy;
    const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const xml = `<p:sp ${NSD}><p:nvSpPr><p:cNvPr id="${id}" name="Text Art ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(cx * 0.15)}" y="${Math.round(cy * 0.38)}"/><a:ext cx="${Math.round(cx * 0.7)}" cy="${Math.round(cy * 0.2)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="6000" b="1"><a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill></a:ln><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="000000"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst></a:rPr><a:t>${esc(text || 'Text art')}</a:t></a:r></a:p></p:txBody></p:sp>`;
    this.tree(slide).append(imp(slide.doc, xml));
    this.markSlide(slide);
    await this.afterInsert();
  };
  E.insertTable = async function (rows, cols) {
    const slide = cur(this);
    this.snapshot();
    const id = this.nextId(slide);
    const cx = this.deck.cx, cy = this.deck.cy;
    const w = Math.round(cx * 0.8), colW = Math.floor(w / cols), rowH = 370840;
    const cell = '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr/></a:tc>';
    const xml = `<p:graphicFrame ${NSD}><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${Math.round(cx * 0.1)}" y="${Math.round(cy * 0.25)}"/><a:ext cx="${colW * cols}" cy="${rowH * rows}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${`<a:gridCol w="${colW}"/>`.repeat(cols)}</a:tblGrid>${`<a:tr h="${rowH}">${cell.repeat(cols)}</a:tr>`.repeat(rows)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
    this.tree(slide).append(imp(slide.doc, xml));
    this.markSlide(slide);
    await this.afterInsert();
  };
  // Double-click a table cell to edit its text (saved into that cell's a:txBody).
  E.editTableCell = function (shape, node, e) {
    const td = e.target.closest('td');
    if (!td) return;
    const tbl = node.getElementsByTagNameNS(NS.a, 'tbl')[0];
    const tr = kids(tbl, 'tr')[td.parentElement.rowIndex];
    const tcs = tr ? kids(tr, 'tc').filter((tc) => tc.getAttribute('hMerge') !== '1' && tc.getAttribute('vMerge') !== '1') : [];
    const tc = tcs[[...td.parentElement.children].indexOf(td)];
    const text = td.querySelector('.pp-text');
    if (!tc || !text) return;
    this.snapshot();
    this.editingNode = tc;
    this.editingEl = text;
    this.editingShape = shape;
    text.contentEditable = 'true';
    text.classList.add('editing');
    text.querySelectorAll('*').forEach((x) => { x.dataset.o = '1'; });
    text.querySelectorAll('.pp-r').forEach((x) => { x.dataset.os = x.style.cssText; });
    if (!text.querySelector('.pp-p')) { const p = document.createElement('div'); p.className = 'pp-p'; p.dataset.p = '0'; p.innerHTML = '<br>'; text.append(p); }
    text.focus();
    if (document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(e.clientX, e.clientY); if (r && text.contains(r.startContainer)) { getSelection().removeAllRanges(); getSelection().addRange(r); } }
    text.addEventListener('input', () => this.changed());
    if (this.opts.onEditing) this.opts.onEditing(true);
  };

  // Slide numbers: a slide-number placeholder (from the layout) on every slide, or none.
  E.toggleSlideNumbers = async function () {
    const deck = this.deck;
    const has = deck.slides.some((s) => [...s.doc.getElementsByTagNameNS(NS.p, 'ph')].some((p) => p.getAttribute('type') === 'sldNum'));
    this.snapshot();
    for (const s of deck.slides) {
      const tree = this.tree(s);
      if (has) {
        for (const sp of [...s.doc.getElementsByTagNameNS(NS.p, 'sp')]) { const ph = deck.phOf(sp); if (ph && ph.type === 'sldNum') sp.remove(); }
      } else {
        const lph = s.layout && [...s.layout.getElementsByTagNameNS(NS.p, 'sp')].find((sp) => { const ph = deck.phOf(sp); return ph && ph.type === 'sldNum'; });
        const idx = lph ? deck.phOf(lph).idx : null;
        const id = this.nextId(s);
        const pos = lph ? '' : `<p:spPr><a:xfrm><a:off x="${Math.round(deck.cx * 0.82)}" y="${Math.round(deck.cy * 0.9)}"/><a:ext cx="${Math.round(deck.cx * 0.14)}" cy="${Math.round(deck.cy * 0.07)}"/></a:xfrm></p:spPr>`;
        tree.append(imp(s.doc, `<p:sp ${NSD}><p:nvSpPr><p:cNvPr id="${id}" name="Slide Number ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter"${idx != null ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr>${pos || '<p:spPr/>'}<p:txBody><a:bodyPr/><a:lstStyle/><a:p>${lph ? '' : '<a:pPr algn="r"/>'}<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`));
      }
      this.markSlide(s);
    }
    this.changed();
    await this.renderAll();
    return !has;
  };

  // Link on the selected text (while editing) or on the selected shape.
  E.insertLink = async function (url) {
    if (!/^(https?:|mailto:)/i.test(url)) throw new Error('Links must start with https://, http:// or mailto:');
    const slide = cur(this);
    if (this.editingEl) {
      document.execCommand('createLink', false, url);
      this.changed();
      return;
    }
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select a shape, or some text while editing it.');
    this.snapshot();
    const rid = this.deck.addRel(slide.path, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', url, true);
    slide.rels = await this.deck.rels(slide.path);
    for (const { node } of sel) {
      const c = first(first(target(node), 'nvSpPr') || first(target(node), 'nvPicPr') || first(target(node), 'nvCxnSpPr') || node, 'cNvPr');
      if (!c) continue;
      for (const old of kids(c, 'hlinkClick')) old.remove();
      const h = aEl(slide.doc, 'hlinkClick');
      h.setAttributeNS(NS.r, 'r:id', rid);
      c.prepend(h);
    }
    this.markSlide(slide);
    this.changed();
  };

  // ======================================================================
  // Format: fill, border, shadow, size & rotation
  // ======================================================================

  function spPrOf(ed, node) { const n = target(node); return ed.xfrmEl(n, true) && first(n, n.localName === 'grpSp' ? 'grpSpPr' : 'spPr'); }
  const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
  const AFTER_FILL = ['ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'];
  E.setFill = async function (color) {
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select a shape first.');
    this.snapshot();
    const slide = cur(this);
    for (const { node } of sel) {
      if (node.localName === 'graphicFrame') continue;
      const pr = spPrOf(this, node);
      for (const f of kids(pr).filter((c) => FILLS.includes(c.localName))) f.remove();
      const fill = color ? imp(slide.doc, `<a:solidFill xmlns:a="${NS.a}"><a:srgbClr val="${hex(color)}"/></a:solidFill>`) : aEl(slide.doc, 'noFill');
      pr.insertBefore(fill, kids(pr).find((c) => AFTER_FILL.includes(c.localName)) || null);
    }
    this.markSlide(slide);
    this.changed();
    await this.redrawKeep(sel.map((x) => x.idx));
  };
  // border: { color, width (pt), dash, none }
  E.setBorder = async function (border) {
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select a shape first.');
    this.snapshot();
    const slide = cur(this);
    for (const { node } of sel) {
      if (node.localName === 'graphicFrame') continue;
      const pr = spPrOf(this, node);
      let ln = first(pr, 'ln');
      if (!ln) { ln = aEl(slide.doc, 'ln'); pr.insertBefore(ln, kids(pr).find((c) => AFTER_FILL.slice(1).includes(c.localName)) || null); }
      if (border.none) { for (const c of kids(ln)) if (FILLS.includes(c.localName)) c.remove(); ln.prepend(aEl(slide.doc, 'noFill')); continue; }
      if (border.color) {
        for (const c of kids(ln)) if (FILLS.includes(c.localName)) c.remove();
        ln.prepend(imp(slide.doc, `<a:solidFill xmlns:a="${NS.a}"><a:srgbClr val="${hex(border.color)}"/></a:solidFill>`));
      }
      if (border.width) ln.setAttribute('w', Math.round(border.width * 12700));
      if (border.dash) {
        for (const c of kids(ln, 'prstDash')) c.remove();
        const d = aEl(slide.doc, 'prstDash', { val: border.dash });
        const fill = kids(ln).find((c) => FILLS.includes(c.localName));
        if (fill) fill.after(d); else ln.prepend(d);
      }
    }
    this.markSlide(slide);
    this.changed();
    await this.redrawKeep(sel.map((x) => x.idx));
  };
  E.toggleShadow = async function () {
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select a shape first.');
    this.snapshot();
    const slide = cur(this);
    for (const { node } of sel) {
      if (node.localName === 'graphicFrame') continue;
      const pr = spPrOf(this, node);
      const eff = first(pr, 'effectLst');
      if (eff && first(eff, 'outerShdw')) { eff.remove(); continue; }
      if (eff) eff.remove();
      const e = imp(slide.doc, `<a:effectLst xmlns:a="${NS.a}"><a:outerShdw blurRad="76200" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`);
      pr.insertBefore(e, kids(pr).find((c) => ['scene3d', 'sp3d', 'extLst'].includes(c.localName)) || null);
    }
    this.markSlide(slide);
    this.changed();
    await this.redrawKeep(sel.map((x) => x.idx));
  };
  E.sizeInfo = function () {
    const sel = this.selectedNodes()[0];
    if (!sel) return null;
    const b = this.boxOf(sel.el);
    const xf = this.xfrmEl(target(sel.node), false);
    return { x: b.x / EMU_IN, y: b.y / EMU_IN, w: b.w / EMU_IN, h: b.h / EMU_IN, rot: xf && xf.getAttribute('rot') ? +xf.getAttribute('rot') / 60000 : 0 };
  };
  E.setSize = async function (v) {
    const sel = this.selectedNodes()[0];
    if (!sel) throw new Error('Select a shape first.');
    this.snapshot();
    place(this, sel.node, { x: v.x * EMU_IN, y: v.y * EMU_IN, w: v.w * EMU_IN, h: v.h * EMU_IN });
    const xf = this.xfrmEl(target(sel.node), true);
    const r = Math.round(((v.rot % 360) + 360) % 360 * 60000);
    if (r) xf.setAttribute('rot', r); else xf.removeAttribute('rot');
    this.markSlide(cur(this));
    this.changed();
    await this.redrawKeep([sel.idx]);
  };

  // ======================================================================
  // Slide: background, layout, transition, skip, theme
  // ======================================================================

  function setBackground(slide, color) {
    const cSld = slide.doc.getElementsByTagNameNS(NS.p, 'cSld')[0];
    for (const b of kids(cSld, 'bg')) b.remove();
    if (!color) return;
    cSld.insertBefore(imp(slide.doc, `<p:bg ${NSD}><p:bgPr><a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`), cSld.firstChild);
  }
  E.setBackground = async function (color, all) {
    this.snapshot();
    for (const s of all ? this.deck.slides : [cur(this)]) { setBackground(s, color); this.markSlide(s); }
    this.changed();
    await this.renderAll();
  };
  E.applyLayout = async function (layoutPath) {
    const slide = cur(this);
    this.snapshot();
    const rels = this.deck.relsDoc(slide.path);
    const rel = [...rels.documentElement.children].find((r) => r.getAttribute('Type') === REL.layout);
    rel.setAttribute('Target', PptxEditor.relative(slide.path, layoutPath));
    this.deck.dirty.add(slide.path.replace(/([^/]+)$/, '_rels/$1.rels'));
    this.markSlide(slide);
    await this.reloadSlides(this.current);
  };
  const TRANSITIONS = {
    none: '', fade: '<p:fade/>', push: '<p:push dir="u"/>', wipe: '<p:wipe dir="r"/>', cover: '<p:cover dir="l"/>', split: '<p:split orient="vert" dir="out"/>',
    zoom: '<p:zoom/>', dissolve: '<p:dissolve/>', cut: '<p:cut/>',
  };
  E.transitionOf = function (slide) {
    const t = (slide || cur(this)).doc.getElementsByTagNameNS(NS.p, 'transition')[0];
    if (!t) return { type: 'none', speed: 'med' };
    const kind = kids(t).find((c) => c.namespaceURI === NS.p);
    return { type: kind ? kind.localName : 'fade', speed: t.getAttribute('spd') || 'med', dir: kind && kind.getAttribute('dir') };
  };
  E.setTransition = async function (type, speed, all) {
    this.snapshot();
    for (const s of all ? this.deck.slides : [cur(this)]) {
      const root = s.doc.documentElement;
      for (const t of kids(root, 'transition')) t.remove();
      for (const ac of kids(root, 'AlternateContent')) if (ac.getElementsByTagNameNS(NS.p, 'transition').length) ac.remove();
      if (type && type !== 'none') {
        const t = imp(s.doc, `<p:transition ${NSD} spd="${speed || 'med'}">${TRANSITIONS[type] || '<p:fade/>'}</p:transition>`);
        const after = kids(root).find((c) => ['timing', 'extLst'].includes(c.localName));
        root.insertBefore(t, after || null);
      }
      this.markSlide(s);
    }
    this.changed();
    await this.renderAll();
  };
  E.toggleSkip = async function (i) {
    const s = this.deck.slides[i ?? this.current];
    this.snapshot();
    const root = s.doc.documentElement;
    if (root.getAttribute('show') === '0') root.removeAttribute('show'); else root.setAttribute('show', '0');
    this.markSlide(s);
    this.changed();
    await this.renderThumbs();
  };
  E.isSkipped = function (i) { const s = this.deck.slides[i ?? this.current]; return !!s && s.doc.documentElement.getAttribute('show') === '0'; };

  const THEMES = {
    Office: { dk2: '44546A', lt2: 'E7E6E6', accents: ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'], fonts: ['Calibri Light', 'Calibri'] },
    Ocean: { dk2: '17406D', lt2: 'DBEFF9', accents: ['0F6FC6', '009DD9', '0BD0D9', '10CF9B', '7CCA62', 'A5C249'], fonts: ['Segoe UI Semibold', 'Segoe UI'] },
    Forest: { dk2: '2C3C24', lt2: 'E5EADB', accents: ['549E39', '8AB833', 'C0CF3A', '029676', '4AB5C4', '0989B1'], fonts: ['Georgia', 'Georgia'] },
    Sunset: { dk2: '4E3B30', lt2: 'FBEEC9', accents: ['F07F09', '9F2936', '1B587C', '4E8542', '604878', 'C19859'], fonts: ['Trebuchet MS', 'Trebuchet MS'] },
    Berry: { dk2: '3F1D5A', lt2: 'EDE3F2', accents: ['8E44AD', 'C0392B', 'D35400', '16A085', '2980B9', '7F8C8D'], fonts: ['Century Gothic', 'Century Gothic'] },
    Mono: { dk2: '262626', lt2: 'EDEDED', accents: ['595959', '7F7F7F', 'A5A5A5', '3F3F3F', '262626', 'BFBFBF'], fonts: ['Arial', 'Arial'] },
    Night: { dark: true, dk2: '1F2A44', lt2: 'C9D3E6', accents: ['4FC3F7', 'FFB74D', 'AED581', 'F06292', '9575CD', '4DB6AC'], fonts: ['Segoe UI Light', 'Segoe UI'] },
  };
  E.themeNames = () => Object.keys(THEMES);
  E.applyTheme = async function (name) {
    const t = THEMES[name];
    if (!t) return;
    this.snapshot();
    const deck = this.deck;
    const paths = [...new Set(deck.slides.map((s) => s.themePath).filter(Boolean))];
    for (const p of paths) {
      const doc = await deck.doc(p);
      const scheme = doc.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
      const set = (name2, val) => { const el = first(scheme, name2); if (!el) return; while (el.firstChild) el.firstChild.remove(); el.append(imp(doc, `<a:srgbClr xmlns:a="${NS.a}" val="${val}"/>`)); };
      if (t.dark) { set('dk1', 'FFFFFF'); set('lt1', '111827'); } else { const dk1 = first(scheme, 'dk1'); if (dk1 && !first(dk1, 'sysClr')) set('dk1', '000000'); const lt1 = first(scheme, 'lt1'); if (lt1 && !first(lt1, 'sysClr')) set('lt1', 'FFFFFF'); }
      set('dk2', t.dk2); set('lt2', t.lt2);
      t.accents.forEach((c, i) => set(`accent${i + 1}`, c));
      scheme.setAttribute('name', `Doc Studio ${name}`);
      const fs = doc.getElementsByTagNameNS(NS.a, 'fontScheme')[0];
      if (fs) {
        const major = first(first(fs, 'majorFont'), 'latin'), minor = first(first(fs, 'minorFont'), 'latin');
        if (major) major.setAttribute('typeface', t.fonts[0]);
        if (minor) minor.setAttribute('typeface', t.fonts[1]);
      }
      deck.dirty.add(p);
    }
    this.changed();
    await deck.loadSlides();
    await this.renderAll();
  };

  // ======================================================================
  // Animations (entrance on click): appear, fade in, fly in from the bottom
  // ======================================================================

  function spidOf(node) { const c = target(node).getElementsByTagNameNS(NS.p, 'cNvPr')[0]; return c ? c.getAttribute('id') : null; }
  // The click steps of a slide: [[{ spid, effect }]] (effects in the same step start together).
  E.animationsOf = function (slide) {
    const timing = slide.doc.getElementsByTagNameNS(NS.p, 'timing')[0];
    if (!timing) return [];
    const main = [...timing.getElementsByTagNameNS(NS.p, 'cTn')].find((c) => c.getAttribute('nodeType') === 'mainSeq');
    if (!main) return [];
    const steps = [];
    for (const par of kids(first(main, 'childTnLst'), 'par')) {
      const effects = [];
      for (const ctn of par.getElementsByTagNameNS(NS.p, 'cTn')) {
        if (ctn.getAttribute('presetClass') !== 'entr') continue;
        const tgt = ctn.getElementsByTagNameNS(NS.p, 'spTgt')[0];
        if (!tgt) continue;
        const id = +ctn.getAttribute('presetID');
        effects.push({ spid: tgt.getAttribute('spid'), effect: id === 10 ? 'fade' : id === 2 ? 'fly' : id === 1 ? 'appear' : 'fade' });
      }
      if (effects.length) steps.push(effects);
    }
    return steps;
  };
  function timingXml(steps) {
    let id = 3;
    const effectXml = (e) => {
      const tgt = `<p:tgtEl><p:spTgt spid="${e.spid}"/></p:tgtEl>`;
      const show = `<p:set><p:cBhvr><p:cTn id="${id++}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`;
      let extra = '';
      if (e.effect === 'fade') extra = `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${id++}" dur="500"/>${tgt}</p:cBhvr></p:animEffect>`;
      if (e.effect === 'fly') extra = ['ppt_x', 'ppt_y'].map((a) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${id++}" dur="500" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${a}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${a === 'ppt_y' ? '1+#ppt_h/2' : '#ppt_x'}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#${a}"/></p:val></p:tav></p:tavLst></p:anim>`).join('');
      const preset = e.effect === 'fade' ? 'presetID="10" presetClass="entr" presetSubtype="0"' : e.effect === 'fly' ? 'presetID="2" presetClass="entr" presetSubtype="4"' : 'presetID="1" presetClass="entr" presetSubtype="0"';
      return { preset, body: show + extra };
    };
    const pars = steps.map((step) => {
      const outer = id++, inner = id++;
      const effects = step.map((e, k) => { const eid = id++; const x = effectXml(e); return `<p:par><p:cTn id="${eid}" ${x.preset} fill="hold"${e.text ? ' grpId="0"' : ''} nodeType="${k ? 'withEffect' : 'clickEffect'}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${x.body}</p:childTnLst></p:cTn></p:par>`; }).join('');
      return `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${inner}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
    }).join('');
    const blds = [...new Set(steps.flat().filter((e) => e.text).map((e) => e.spid))].map((s) => `<p:bldP spid="${s}" grpId="0"/>`).join('');
    return `<p:timing ${NSD}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${pars}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>${blds ? `<p:bldLst>${blds}</p:bldLst>` : ''}</p:timing>`;
  }
  // effect: 'appear'|'fade'|'fly'|null (remove) for the selected shapes, each on its own click.
  E.setAnimation = async function (effect) {
    const sel = this.selectedNodes();
    if (!sel.length) throw new Error('Select a shape to animate first.');
    const slide = cur(this);
    this.snapshot();
    const ids = new Set(sel.map((x) => spidOf(x.node)));
    let steps = this.animationsOf(slide).map((st) => st.filter((e) => !ids.has(e.spid))).filter((st) => st.length);
    if (effect) for (const x of sel) steps.push([{ spid: spidOf(x.node), effect, text: !!first(target(x.node), 'txBody') }]);
    // Shapes with text need a build entry; keep that flag for existing steps too.
    const shapes = list(this, slide);
    steps = steps.map((st) => st.map((e) => ({ ...e, text: e.text ?? shapes.some((n) => spidOf(n) === e.spid && first(target(n), 'txBody')) })));
    const root = slide.doc.documentElement;
    for (const t of kids(root, 'timing')) t.remove();
    if (steps.length) {
      const after = kids(root).find((c) => c.localName === 'extLst');
      root.insertBefore(imp(slide.doc, timingXml(steps)), after || null);
    }
    this.markSlide(slide);
    this.changed();
    await this.redrawKeep(sel.map((x) => x.idx));
  };

  // ======================================================================
  // Slideshow and presenter view
  // ======================================================================

  E.present = async function (opts) {
    opts = opts || {};
    this.commitText();
    const deck = this.deck;
    const order = deck.slides.map((s, i) => i).filter((i) => !this.isSkipped(i) || i === (opts.from ?? this.current));
    let pos = Math.max(0, order.indexOf(opts.from ?? this.current));
    if (opts.fromStart) pos = 0;
    const show = h('div', { class: `ppe-show${opts.presenter ? ' presenter' : ''}` });
    const stage = h('div', { class: 'ppe-show-stage' });
    show.append(stage);
    let side = null, notesEl = null, nextEl = null, timerEl = null, counter = null;
    if (opts.presenter) {
      notesEl = h('div', { class: 'ppe-show-notes' });
      nextEl = h('div', { class: 'ppe-show-next' });
      timerEl = h('div', { class: 'ppe-show-timer', text: '00:00' });
      counter = h('div', { class: 'ppe-show-count' });
      side = h('aside', { class: 'ppe-show-side' }, h('div', { class: 'ppe-show-head' }, timerEl, counter), h('b', { text: 'Next' }), nextEl, h('b', { text: 'Speaker notes' }), notesEl, h('div', { class: 'ppe-show-help', text: '→ / Space next · ← back · B black · Esc end' }));
      show.append(side);
    }
    document.body.append(show);
    const wasFull = await desk.fullscreen();
    const started = Date.now();
    const tick = timerEl ? setInterval(() => { const s = Math.floor((Date.now() - started) / 1000); timerEl.textContent = `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`; }, 500) : null;
    let steps = [], step = 0, black = false, cleaned = false;
    const spidEls = new Map();
    const render = async (dir) => {
      const i = order[pos];
      const s = deck.slides[i];
      await new Promise((r) => requestAnimationFrame(r));
      const area = stage.getBoundingClientRect();
      const width = Math.min(area.width, area.height * deck.cx / deck.cy);
      const el = await new PptxEditor.Renderer(deck, { slideNumber: i + 1 }).slide(s, width, false);
      // Shapes that animate in start hidden.
      steps = this.animationsOf(s);
      step = dir < 0 ? steps.length : 0;
      spidEls.clear();
      const shapes = list(this, s);
      el.querySelectorAll('.pp-own').forEach((x) => { const n = shapes[+x.dataset.idx]; if (n) spidEls.set(spidOf(n), x); });
      if (dir >= 0) for (const st of steps) for (const e of st) { const x = spidEls.get(e.spid); if (x) x.classList.add('ppe-anim-hidden'); }
      const tr = this.transitionOf(s);
      el.classList.add('ppe-show-slide');
      if (dir > 0 && tr.type !== 'none' && tr.type !== 'cut') { el.style.animationDuration = { slow: '1s', med: '0.7s', fast: '0.4s' }[tr.speed] || '0.7s'; el.classList.add(`tr-${tr.type}`); }
      stage.replaceChildren(el);
      if (opts.presenter) {
        notesEl.textContent = (await this.getNotes(i)) || 'No speaker notes for this slide.';
        counter.textContent = `Slide ${pos + 1} of ${order.length}`;
        const ni = order[pos + 1];
        nextEl.replaceChildren(ni != null ? await new PptxEditor.Renderer(deck, { slideNumber: ni + 1 }).slide(deck.slides[ni], 260, false) : h('p', { class: 'muted', text: 'End of slideshow' }));
      }
    };
    const playStep = () => {
      const st = steps[step++];
      for (const e of st) {
        const x = spidEls.get(e.spid);
        if (!x) continue;
        x.classList.remove('ppe-anim-hidden');
        x.classList.add(`ppe-anim-${e.effect}`);
      }
    };
    const next = async () => {
      if (black) { black = false; show.classList.remove('black'); return; }
      if (step < steps.length) { playStep(); return; }
      if (pos < order.length - 1) { pos++; await render(1); } else end();
    };
    const prev = async () => { if (pos > 0) { pos--; await render(-1); } };
    const end = () => {
      if (cleaned) return;
      cleaned = true;
      clearInterval(tick);
      document.removeEventListener('keydown', onKey, true);
      show.remove();
      if (wasFull) desk.fullscreen();
      this.go(order[pos]);
    };
    let typed = '';
    const onKey = (e) => {
      e.preventDefault(); e.stopPropagation();
      const k = e.key;
      if (k === 'Escape') end();
      else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n', 'N'].includes(k) && !(k === 'Enter' && typed)) next();
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) prev();
      else if (k === 'Home') { pos = 0; render(0); } else if (k === 'End') { pos = order.length - 1; render(0); }
      else if (k === 'b' || k === 'B' || k === '.') { black = !black; show.classList.toggle('black', black); }
      else if (/^\d$/.test(k)) typed += k;
      else if (k === 'Enter' && typed) { const n = Math.max(1, Math.min(order.length, +typed)); typed = ''; pos = n - 1; render(0); }
    };
    document.addEventListener('keydown', onKey, true);
    stage.addEventListener('click', () => next());
    stage.addEventListener('contextmenu', (e) => { e.preventDefault(); prev(); });
    await render(0);
    this.showing = { next, prev, end, el: show, get pos() { return pos; }, get step() { return step; } };
  };

  // ======================================================================
  // Grid view (all slides at a glance)
  // ======================================================================

  E.toggleGridView = async function () {
    if (this.gridEl) { this.gridEl.remove(); this.gridEl = null; return; }
    const grid = h('div', { class: 'ppe-grid' });
    this.gridEl = grid;
    this.stageWrap.append(grid);
    for (const [i, s] of this.deck.slides.entries()) {
      const card = h('div', { class: `ppe-grid-card${this.isSkipped(i) ? ' skipped' : ''}`, title: `Slide ${i + 1}` });
      card.append(await new PptxEditor.Renderer(this.deck, { slideNumber: i + 1 }).slide(s, 240, false), h('span', { text: String(i + 1) }));
      card.addEventListener('click', () => { this.toggleGridView(); this.go(i); });
      grid.append(card);
    }
  };

  // Skipped slides are dimmed in the side panel.
  const renderThumbs = E.renderThumbs;
  E.renderThumbs = async function () {
    await renderThumbs.call(this);
    [...this.side.children].forEach((t, i) => t.classList.toggle('skipped', this.isSkipped(i)));
  };

  // ======================================================================
  // Menus for the presentation editor
  // ======================================================================

  PptxUI.menus = function (tab, ed) {
    const run = (label, fn) => async () => { try { await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } };
    const it = (label, fn, o) => ({ label, run: run(label.toLowerCase(), fn), ...(o || {}) });
    const sub = (label, items) => ({ label, sub: items });
    const SEP = { sep: true };
    const editing = { enabled: () => tab.mode === 'edit' && !!tab.pptEditor };
    const ask = DocTools.form;
    return {
      View: [
        it('Present', () => PptxUI.present(tab, {}), { key: 'Ctrl+F5' }),
        it('Present from beginning', () => PptxUI.present(tab, { fromStart: true }), { key: 'Ctrl+Shift+F5' }),
        it('Presenter view', () => PptxUI.present(tab, { presenter: true })),
        SEP,
        it('Grid view', () => ed().toggleGridView(), { ...editing, checked: () => !!(tab.pptEditor && tab.pptEditor.gridEl) }),
      ],
      Insert: [
        it('Table…', () => new Promise((resolve) => DocTools.tableGrid($('#menu-names'), async (r, c) => { await ed().insertTable(r, c); resolve(); })), editing),
        sub('Line', [it('Line', () => ed().insertLine(false), editing), it('Arrow', () => ed().insertLine(true), editing)]),
        it('Text art…', async () => { const v = await ask('Text art', [{ name: 't', label: 'Text', value: '' }], 'Insert'); if (v && v.t) await ed().insertWordArt(v.t); }, editing),
        it('Special characters…', () => { if (!ed().editingEl) throw new Error('Double-click a text box to edit its text first.'); DocTools.specialCharacters(); }, editing),
        it('Link…', async () => { const v = await ask('Insert link', [{ name: 'url', label: 'Link', value: 'https://' }], 'Insert'); if (v) await ed().insertLink(v.url.trim()); }, { ...editing, key: 'Ctrl+K' }),
        SEP,
        it('Slide numbers (on/off)', async () => { const on = await ed().toggleSlideNumbers(); flash(on ? 'Slide numbers added' : 'Slide numbers removed'); }, editing),
        sub('Animation', [it('Appear', () => ed().setAnimation('appear'), editing), it('Fade in', () => ed().setAnimation('fade'), editing), it('Fly in from bottom', () => ed().setAnimation('fly'), editing), it('Remove animation', () => ed().setAnimation(null), editing)]),
      ],
      Format: [
        sub('Fill colour', [...[['Blue', '#4472C4'], ['Orange', '#ED7D31'], ['Green', '#70AD47'], ['Red', '#E53935'], ['Yellow', '#FFC000'], ['Grey', '#A5A5A5'], ['White', '#FFFFFF'], ['Black', '#000000']].map(([n, c]) => it(n, () => ed().setFill(c), editing)), it('Custom…', async () => { const v = await ask('Fill colour', [{ name: 'c', label: 'Colour', type: 'color', value: '#4472c4' }], 'Apply'); if (v) await ed().setFill(v.c); }, editing), it('No fill', () => ed().setFill(null), editing)]),
        sub('Border', [it('Colour…', async () => { const v = await ask('Border colour', [{ name: 'c', label: 'Colour', type: 'color', value: '#000000' }], 'Apply'); if (v) await ed().setBorder({ color: v.c }); }, editing),
          ...[1, 2, 3, 4, 6, 8].map((w) => it(`Weight ${w} pt`, () => ed().setBorder({ width: w }), editing)),
          ...[['solid', 'Solid'], ['dash', 'Dashed'], ['sysDot', 'Dotted'], ['dashDot', 'Dash dot']].map(([d, n]) => it(n, () => ed().setBorder({ dash: d }), editing)),
          it('No border', () => ed().setBorder({ none: true }), editing)]),
        it('Drop shadow (on/off)', () => ed().toggleShadow(), editing),
        it('Size & rotation…', async () => {
          const cur0 = ed().sizeInfo();
          if (!cur0) throw new Error('Select a shape first.');
          const v = await ask('Size & rotation (inches, degrees)', [{ name: 'x', label: 'Left', type: 'number', step: '0.01', value: cur0.x.toFixed(2) }, { name: 'y', label: 'Top', type: 'number', step: '0.01', value: cur0.y.toFixed(2) }, { name: 'w', label: 'Width', type: 'number', step: '0.01', value: cur0.w.toFixed(2) }, { name: 'h', label: 'Height', type: 'number', step: '0.01', value: cur0.h.toFixed(2) }, { name: 'rot', label: 'Rotation (degrees)', type: 'number', step: '1', value: String(Math.round(cur0.rot)) }], 'Apply');
          if (v) await ed().setSize({ x: +v.x, y: +v.y, w: Math.max(0.05, +v.w), h: Math.max(0.05, +v.h), rot: +v.rot || 0 });
        }, editing),
      ],
      Slide: [
        it('Skip slide (on/off)', () => ed().toggleSkip(), { ...editing, checked: () => !!(tab.pptEditor && tab.pptEditor.isSkipped()) }),
        sub('Change background', [it('Colour…', async () => { const v = await ask('Slide background', [{ name: 'c', label: 'Colour', type: 'color', value: '#ffffff' }, { name: 'all', label: 'Apply to all slides', type: 'checkbox', value: false }], 'Apply'); if (v) await ed().setBackground(v.c, v.all); }, editing), it('Reset to the layout\'s background', () => ed().setBackground(null, false), editing)]),
        { label: 'Apply layout', sub: async () => (await ed().layouts()).map((l) => it(l.name, () => ed().applyLayout(l.path), editing)) },
        it('Transition…', async () => {
          const t = ed().transitionOf();
          const v = await ask('Transition', [{ name: 'type', label: 'Effect', type: 'select', value: t.type in { none: 1, fade: 1, push: 1, wipe: 1, cover: 1, split: 1, zoom: 1, dissolve: 1, cut: 1 } ? t.type : 'fade', options: [['none', 'None'], ['fade', 'Fade'], ['push', 'Push'], ['wipe', 'Wipe'], ['cover', 'Cover'], ['split', 'Split'], ['zoom', 'Zoom'], ['dissolve', 'Dissolve'], ['cut', 'Cut']] }, { name: 'speed', label: 'Speed', type: 'select', value: t.speed, options: [['slow', 'Slow'], ['med', 'Medium'], ['fast', 'Fast']] }, { name: 'all', label: 'Apply to all slides', type: 'checkbox', value: false }], 'Apply');
          if (v) await ed().setTransition(v.type, v.speed, v.all);
        }, editing),
        sub('Change theme', PptxUI.Editor.prototype.themeNames().map((n) => it(n, () => ed().applyTheme(n), editing))),
      ],
      Arrange: [
        sub('Order', [it('Bring to front', () => ed().order('front'), { ...editing, key: 'Ctrl+Shift+↑' }), it('Bring forward', () => ed().order('forward'), { ...editing, key: 'Ctrl+↑' }), it('Send backward', () => ed().order('backward'), { ...editing, key: 'Ctrl+↓' }), it('Send to back', () => ed().order('back'), { ...editing, key: 'Ctrl+Shift+↓' })]),
        sub('Align', [it('Left', () => ed().align('left'), editing), it('Center', () => ed().align('center'), editing), it('Right', () => ed().align('right'), editing), SEP, it('Top', () => ed().align('top'), editing), it('Middle', () => ed().align('middle'), editing), it('Bottom', () => ed().align('bottom'), editing)]),
        sub('Distribute', [it('Horizontally', () => ed().distribute('h'), editing), it('Vertically', () => ed().distribute('v'), editing)]),
        sub('Center on page', [it('Horizontally', () => ed().centerOnPage('h'), editing), it('Vertically', () => ed().centerOnPage('v'), editing)]),
        sub('Rotate', [it('Rotate clockwise 90°', () => ed().rotate('cw'), editing), it('Rotate counter-clockwise 90°', () => ed().rotate('ccw'), editing), it('Flip horizontally', () => ed().rotate('flipH'), editing), it('Flip vertically', () => ed().rotate('flipV'), editing)]),
        SEP,
        it('Group', () => ed().group(), { ...editing, key: 'Ctrl+Alt+G' }),
        it('Ungroup', () => ed().ungroup(), { ...editing, key: 'Ctrl+Alt+Shift+G' }),
      ],
    };
  };

  // Present works from View too: the editor (which owns the deck) is prepared first.
  PptxUI.present = async function (tab, opts) {
    if (!tab.pptEditor) await buildPptxEditor(tab);
    return tab.pptEditor.present(opts);
  };

  // Keyboard: present, order, group.
  document.addEventListener('keydown', (e) => {
    if (!active || active.kind !== 'pptx' || document.querySelector('.ppe-show')) return;
    const run = (fn) => { e.preventDefault(); Promise.resolve(fn()).catch((err) => desk.error("Couldn't do that", err.message)); };
    if (e.key === 'F5' && e.ctrlKey) { run(() => PptxUI.present(active, e.shiftKey ? { fromStart: true } : {})); return; }
    const ed = active.pptEditor;
    if (!ed) return;
    const a = document.activeElement;
    const typing = a && (a.isContentEditable || ['INPUT', 'TEXTAREA', 'SELECT'].includes(a.tagName));
    if (typing || active.mode !== 'edit') return;
    if (e.ctrlKey && e.key === 'ArrowUp') run(() => ed.order(e.shiftKey ? 'front' : 'forward'));
    else if (e.ctrlKey && e.key === 'ArrowDown') run(() => ed.order(e.shiftKey ? 'back' : 'backward'));
    else if (e.ctrlKey && e.altKey && e.key.toLowerCase() === 'g') run(() => (e.shiftKey ? ed.ungroup() : ed.group()));
  }, true);
})();
