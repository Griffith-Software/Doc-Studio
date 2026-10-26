'use strict';
// PowerPoint editing UI on top of PptxEditor.Deck/Renderer.

const PptxUI = (() => {
  const { NS, REL, parseXml, serialize, kids, first, path1, num, aEl, pEl, relative } = PptxEditor;
  const SHAPE_TYPES = new Set(['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'AlternateContent']);

  class Editor {
    constructor(host, opts) {
      this.host = host;
      this.opts = opts || {};
      this.current = 0;
      this.history = [];
      this.future = [];
    }

    async load(bytes) {
      this.deck = await new PptxEditor.Deck().load(bytes);
      this.buildDom();
      await this.renderAll();
      return this;
    }

    // ---------------------------------------------------------------- layout

    buildDom() {
      const el = (tag, cls) => { const e = document.createElement(tag); if (cls) e.className = cls; return e; };
      this.root = el('div', 'ppe');
      this.side = el('div', 'ppe-side');
      this.stageWrap = el('div', 'ppe-main');
      this.stage = el('div', 'ppe-stage');
      this.notes = el('textarea', 'ppe-notes');
      this.notes.placeholder = 'Speaker notes';
      this.notes.addEventListener('change', () => this.setNotes(this.current, this.notes.value));
      this.stageWrap.append(this.stage, this.notes);
      this.root.append(this.side, this.stageWrap);
      this.host.replaceChildren(this.root);
      this.resizeObs = new ResizeObserver(() => { clearTimeout(this.rt); this.rt = setTimeout(() => this.renderCurrent(), 120); });
      this.resizeObs.observe(this.stageWrap);
      document.addEventListener('keydown', this.keyHandler = (e) => this.onKey(e));
    }

    async renderAll() {
      await this.renderThumbs();
      await this.renderCurrent();
    }

    async renderThumbs() {
      const deck = this.deck;
      this.side.replaceChildren();
      for (const [i, s] of deck.slides.entries()) {
        const t = document.createElement('div');
        t.className = `ppe-thumb${i === this.current ? ' on' : ''}`;
        t.draggable = true;
        const n = document.createElement('span');
        n.textContent = i + 1;
        const slideEl = await new PptxEditor.Renderer(deck, { slideNumber: i + 1 }).slide(s, 150, false);
        t.append(n, slideEl);
        t.addEventListener('click', () => this.go(i));
        t.addEventListener('dragstart', (e) => e.dataTransfer.setData('text/x-slide', String(i)));
        t.addEventListener('dragover', (e) => { if ([...e.dataTransfer.types].includes('text/x-slide')) { e.preventDefault(); t.classList.add('drop'); } });
        t.addEventListener('dragleave', () => t.classList.remove('drop'));
        t.addEventListener('drop', (e) => { e.preventDefault(); t.classList.remove('drop'); const from = +e.dataTransfer.getData('text/x-slide'); if (!isNaN(from) && from !== i) this.moveSlide(from, i); });
        t.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          this.go(i);
          SheetView.showMenu(e.clientX, e.clientY, [
            ['New slide', '', () => this.addSlideDialog()],
            ['Duplicate slide', '', () => this.duplicateSlide(i)],
            ['Delete slide', '', () => this.deleteSlide(i)],
          ]);
        });
        this.side.append(t);
      }
    }

    async go(i) {
      this.commitText();
      this.current = Math.max(0, Math.min(i, this.deck.slides.length - 1));
      this.side.querySelectorAll('.ppe-thumb').forEach((t, k) => t.classList.toggle('on', k === this.current));
      await this.renderCurrent();
    }

    async renderCurrent() {
      const deck = this.deck;
      const slide = deck.slides[this.current];
      if (!slide) { this.stage.replaceChildren(); return; }
      const availW = Math.max(300, this.stageWrap.clientWidth - 48);
      const availH = Math.max(200, this.stageWrap.clientHeight - 140);
      const width = Math.min(availW, availH * deck.cx / deck.cy, 1600) * (this.zoom || 1);
      const seq = this.renderSeq = (this.renderSeq || 0) + 1;
      const el = await new PptxEditor.Renderer(deck, { slideNumber: this.current + 1 }).slide(slide, width, true);
      if (seq !== this.renderSeq) return; // a newer redraw started meanwhile
      this.slideEl = el;
      this.k = width / deck.cx;
      this.stage.replaceChildren(el);
      this.wireSlide(el);
      const notes = await this.getNotes(this.current);
      if (document.activeElement !== this.notes) this.notes.value = notes;
      this.selected = null;
      this.multi = new Set();
      if (this.afterRender) this.afterRender(el);
    }

    // ---------------------------------------------------------------- shapes in the XML

    tree(slide) { return slide.doc.getElementsByTagNameNS(NS.p, 'spTree')[0]; }
    nodeFor(idx, slide) {
      const list = kids(this.tree(slide)).filter((x) => SHAPE_TYPES.has(x.localName));
      const n = list[idx];
      if (n && n.localName === 'AlternateContent') return first(first(n, 'Choice'), 'sp') || first(first(n, 'Fallback'), 'sp') || n;
      return n;
    }
    nextId(slide) {
      let max = 1;
      for (const c of slide.doc.getElementsByTagNameNS(NS.p, 'cNvPr')) max = Math.max(max, +c.getAttribute('id') || 0);
      return max + 1;
    }
    markSlide(slide) { this.deck.dirty.add(slide.path); }

    // ---------------------------------------------------------------- selection, move, resize

    wireSlide(el) {
      el.addEventListener('pointerdown', (e) => {
        const shape = e.target.closest('.pp-own');
        if (this.editingEl && this.editingEl.contains(e.target)) return;
        // The first click outside the text being edited only finishes editing (the slide is redrawn).
        if (this.editingEl) { this.commitText(); return; }
        if (!shape) { this.select(null); return; }
        const add = e.shiftKey || e.ctrlKey;
        this.select(shape, add);
        if (e.button === 0 && !add) this.drag(shape, e);
      });
      el.addEventListener('dblclick', (e) => {
        const shape = e.target.closest('.pp-own');
        if (!shape) return;
        const node = this.nodeFor(+shape.dataset.idx, this.deck.slides[this.current]);
        if (node && (node.localName === 'sp')) this.editText(shape, e);
        else if (node && node.localName === 'graphicFrame' && this.editTableCell) this.editTableCell(shape, node, e);
      });
    }

    // Shift/Ctrl+click adds to the selection (used by Arrange: align, distribute, group).
    select(shape, add) {
      this.multi = this.multi || new Set();
      if (add && shape) {
        if (this.selected && this.selected !== shape) this.multi.add(this.selected);
        if (this.multi.has(shape)) { this.multi.delete(shape); shape.classList.remove('sel'); return; }
        this.multi.add(shape);
        shape.classList.add('sel');
        this.selected = shape;
        this.slideEl.querySelectorAll('.ppe-handle').forEach((h) => h.remove());
        return;
      }
      this.multi = new Set();
      this.slideEl.querySelectorAll('.pp-own.sel').forEach((s) => s.classList.remove('sel'));
      this.slideEl.querySelectorAll('.ppe-handle').forEach((h) => h.remove());
      this.selected = shape;
      if (!shape) return;
      this.multi.add(shape);
      shape.classList.add('sel');
      shape.dataset.l0 = parseFloat(shape.style.left) || 0;
      shape.dataset.t0 = parseFloat(shape.style.top) || 0;
      if (shape.classList.contains('pp-group')) return; // groups move but don't resize
      const handle = document.createElement('b');
      handle.className = 'ppe-handle';
      handle.addEventListener('pointerdown', (e) => { e.stopPropagation(); e.preventDefault(); this.resize(shape, e); });
      shape.append(handle);
    }

    xfrmEl(node, create) {
      // Shapes keep their transform in spPr/a:xfrm (graphic frames: p:xfrm, groups: grpSpPr/a:xfrm).
      const doc = node.ownerDocument;
      if (node.localName === 'graphicFrame') {
        let x = first(node, 'xfrm');
        if (!x && create) { x = pEl(doc, 'xfrm'); node.insertBefore(x, first(node, 'graphic')); }
        return x;
      }
      const prName = node.localName === 'grpSp' ? 'grpSpPr' : 'spPr';
      let pr = first(node, prName);
      if (!pr && create) { pr = pEl(doc, prName); const after = kids(node).find((c) => /^nv/.test(c.localName)); node.insertBefore(pr, after ? after.nextSibling : node.firstChild); }
      let x = pr && first(pr, 'xfrm');
      if (!x && create) { x = aEl(doc, 'xfrm'); pr.insertBefore(x, pr.firstChild); }
      return x;
    }

    setXfrm(node, box) {
      const x = this.xfrmEl(node, true);
      const doc = node.ownerDocument;
      let off = first(x, 'off'), ext = first(x, 'ext');
      if (!off) { off = aEl(doc, 'off'); x.prepend(off); }
      if (!ext) { ext = aEl(doc, 'ext'); off.after(ext); }
      off.setAttribute('x', Math.round(box.x)); off.setAttribute('y', Math.round(box.y));
      ext.setAttribute('cx', Math.round(box.w)); ext.setAttribute('cy', Math.round(box.h));
      if (node.localName === 'grpSp') {
        // A moved group keeps its child coordinates.
        if (!first(x, 'chOff')) { const co = aEl(doc, 'chOff', { x: Math.round(box.x), y: Math.round(box.y) }); const ce = aEl(doc, 'chExt', { cx: Math.round(box.w), cy: Math.round(box.h) }); x.append(co, ce); }
      }
    }

    shapeBox(shape) {
      const k = this.k;
      return { x: parseFloat(shape.style.left) / k, y: parseFloat(shape.style.top) / k, w: parseFloat(shape.style.width) / k, h: parseFloat(shape.style.height) / k };
    }

    drag(shape, e) {
      const sx = e.clientX, sy = e.clientY;
      const l0 = parseFloat(shape.style.left), t0 = parseFloat(shape.style.top);
      let moved = false;
      const move = (ev) => {
        const dx = ev.clientX - sx, dy = ev.clientY - sy;
        if (!moved && Math.abs(dx) + Math.abs(dy) < 3) return;
        if (!moved) this.snapshot();
        moved = true;
        let l = l0 + dx, t = t0 + dy;
        // Smart guides: snap to the slide's edges/centre and other shapes (Alt turns snapping off).
        if (this.snapTo && !ev.altKey) [l, t] = this.snapTo(shape, l, t);
        shape.style.left = `${l}px`; shape.style.top = `${t}px`;
      };
      const up = () => {
        window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
        if (this.clearGuides) this.clearGuides();
        if (!moved) return;
        this.applyBox(shape);
      };
      window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
    }

    resize(shape, e) {
      const sx = e.clientX, sy = e.clientY;
      const w0 = parseFloat(shape.style.width), h0 = parseFloat(shape.style.height);
      const node = this.nodeFor(+shape.dataset.idx, this.deck.slides[this.current]);
      const keep = node && node.localName === 'pic';
      this.snapshot();
      const move = (ev) => {
        let w = Math.max(8, w0 + ev.clientX - sx), h = Math.max(8, h0 + ev.clientY - sy);
        if (keep && !ev.shiftKey) h = w * h0 / w0;
        shape.style.width = `${w}px`; shape.style.height = `${h}px`;
      };
      const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); this.applyBox(shape, true); };
      window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
    }

    async applyBox(shape, rerender) {
      const slide = this.deck.slides[this.current];
      const node = this.nodeFor(+shape.dataset.idx, slide);
      if (!node) return;
      if (node.localName === 'grpSp') {
        // Groups: shift the group's offset; children keep their own coordinates.
        const x = this.xfrmEl(node, false);
        const off = x && first(x, 'off');
        if (off) {
          const dx = (parseFloat(shape.style.left) - +shape.dataset.l0) / this.k, dy = (parseFloat(shape.style.top) - +shape.dataset.t0) / this.k;
          off.setAttribute('x', Math.round(num(off, 'x', 0) + dx)); off.setAttribute('y', Math.round(num(off, 'y', 0) + dy));
        }
        shape.dataset.l0 = parseFloat(shape.style.left); shape.dataset.t0 = parseFloat(shape.style.top);
      } else this.setXfrm(node, this.shapeBox(shape));
      this.markSlide(slide);
      this.changed();
      if (rerender) { await this.renderCurrent(); this.refreshThumb(this.current); }
      else this.refreshThumb(this.current);
    }

    async deleteSelected() {
      if (!this.selected) return;
      const slide = this.deck.slides[this.current];
      const list = kids(this.tree(slide)).filter((x) => SHAPE_TYPES.has(x.localName));
      const node = list[+this.selected.dataset.idx];
      if (!node) return;
      this.snapshot();
      node.parentNode.removeChild(node);
      this.markSlide(slide);
      this.changed();
      await this.renderCurrent();
      this.refreshThumb(this.current);
    }

    onKey(e) {
      if (!this.root.isConnected || !this.root.offsetParent) return;
      const a = document.activeElement;
      if (a && (a.isContentEditable || ['INPUT', 'TEXTAREA', 'SELECT'].includes(a.tagName))) return;
      if ((e.key === 'Delete' || e.key === 'Backspace') && this.selected) { e.preventDefault(); this.deleteSelected(); }
      if (e.key === 'PageDown' || (e.key === 'ArrowDown' && !this.selected)) { e.preventDefault(); this.go(this.current + 1); }
      if (e.key === 'PageUp' || (e.key === 'ArrowUp' && !this.selected)) { e.preventDefault(); this.go(this.current - 1); }
      if (this.selected && /^Arrow/.test(e.key)) {
        e.preventDefault();
        const d = e.shiftKey ? 10 : 1;
        const s = this.selected;
        if (!this.nudging) { this.snapshot(); this.nudging = true; setTimeout(() => { this.nudging = false; }, 600); }
        if (e.key === 'ArrowLeft') s.style.left = `${parseFloat(s.style.left) - d}px`;
        if (e.key === 'ArrowRight') s.style.left = `${parseFloat(s.style.left) + d}px`;
        if (e.key === 'ArrowUp') s.style.top = `${parseFloat(s.style.top) - d}px`;
        if (e.key === 'ArrowDown') s.style.top = `${parseFloat(s.style.top) + d}px`;
        this.applyBox(s);
      }
    }

    // ---------------------------------------------------------------- text editing

    editText(shape, e) {
      const text = shape.querySelector(':scope > .pp-text');
      if (!text) return;
      this.snapshot();
      this.editingEl = text;
      this.editingShape = shape;
      text.contentEditable = 'true';
      text.classList.add('editing');
      if (text.classList.contains('pp-prompt')) {
        // Start an empty placeholder with one paragraph.
        if (!text.querySelector('.pp-p')) { const p = document.createElement('div'); p.className = 'pp-p'; p.innerHTML = '<br>'; text.append(p); }
      }
      text.querySelectorAll('[data-dd-o]').forEach(() => {});
      text.querySelectorAll('*').forEach((x) => { x.dataset.o = '1'; });
      text.querySelectorAll('.pp-r').forEach((x) => { x.dataset.os = x.style.cssText; });
      text.focus();
      if (e && document.caretRangeFromPoint) {
        const r = document.caretRangeFromPoint(e.clientX, e.clientY);
        if (r && text.contains(r.startContainer)) { getSelection().removeAllRanges(); getSelection().addRange(r); }
      }
      text.addEventListener('input', this.inputHandler = () => this.changed());
      text.addEventListener('paste', (ev) => { ev.preventDefault(); const t = ev.clipboardData.getData('text/plain'); if (t) document.execCommand('insertText', false, t); });
      if (this.opts.onEditing) this.opts.onEditing(true);
    }

    commitText() {
      const text = this.editingEl;
      if (!text) return;
      const shape = this.editingShape;
      this.editingEl = null;
      this.editingShape = null;
      text.contentEditable = 'false';
      text.classList.remove('editing');
      if (this.opts.onEditing) this.opts.onEditing(false);
      const slide = this.deck.slides[this.current];
      const node = this.editingNode || this.nodeFor(+shape.dataset.idx, slide);
      this.editingNode = null;
      if (!node) return;
      this.writeText(node, text);
      this.markSlide(slide);
      // Redraw so the text shows with its real (inherited) formatting; keep the shape selected.
      const idx = shape.dataset.idx, cur = this.current;
      this.rendering = this.renderCurrent().then(() => {
        const again = this.slideEl.querySelector(`.pp-own[data-idx="${idx}"]`);
        if (again && cur === this.current) this.select(again);
      });
      this.refreshThumb(this.current);
    }

    // Rebuild a:p elements from the edited DOM, reusing each paragraph's pPr and each run's rPr.
    writeText(node, text) {
      const doc = node.ownerDocument;
      let tx = first(node, 'txBody');
      if (!tx) {
        tx = pEl(doc, 'txBody');
        tx.append(aEl(doc, 'bodyPr'), aEl(doc, 'lstStyle'));
        node.append(tx);
      }
      const srcPs = kids(tx, 'p');
      for (const p of srcPs) tx.removeChild(p);
      const domPs = [...text.querySelectorAll(':scope > .pp-p')];
      let lastIdx = 0;
      for (const dp of domPs) {
        const pi = dp.dataset.p !== undefined ? +dp.dataset.p : lastIdx;
        lastIdx = pi;
        const src = srcPs[pi] || srcPs[srcPs.length - 1] || null;
        const p = aEl(doc, 'p');
        const pPr = src && first(src, 'pPr');
        if (pPr) p.append(pPr.cloneNode(true));
        const align = dp.style.textAlign;
        if (align && dp.dataset.alignChanged) {
          const pp = first(p, 'pPr') || (() => { const x = aEl(doc, 'pPr'); p.prepend(x); return x; })();
          pp.setAttribute('algn', { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[align] || 'l');
        }
        const srcRuns = src ? kids(src).filter((c) => ['r', 'br', 'fld'].includes(c.localName)) : [];
        const runRPr = (ri) => { const r = srcRuns[ri]; return r ? first(r, 'rPr') : null; };
        const endPr = src && first(src, 'endParaRPr');
        let ctxRun = null;
        const pieces = [];
        const walk = (n) => {
          for (const c of n.childNodes) {
            if (c.nodeType === 3) { const t = c.data.replace(/ /g, ' '); if (t) pieces.push({ t, run: ctxRun, fmt: fmtOf(c, dp) }); continue; }
            if (c.nodeType !== 1) continue;
            if (c.classList.contains('pp-bullet')) continue;
            if (c.dataset.fld && c.dataset.r !== undefined) { pieces.push({ fld: srcRuns[+c.dataset.r] }); continue; }
            if (c.tagName === 'BR') { if (c.parentElement.classList.contains('pp-empty') || c === dp.lastChild) continue; pieces.push({ br: true, run: ctxRun }); continue; }
            if (c.classList.contains('pp-r') && c.dataset.r !== undefined) { ctxRun = +c.dataset.r; walk(c); continue; }
            walk(c);
          }
        };
        walk(dp);
        let group = null;
        const flush = () => {
          if (!group) return;
          const r = aEl(doc, 'r');
          const base = group.run != null ? runRPr(group.run) : endPr;
          const rPr = base ? cloneRPr(base, doc) : aEl(doc, 'rPr', { lang: 'en-US' });
          applyRunFormat(rPr, group.fmt, doc, this.k);
          r.append(rPr);
          const t = aEl(doc, 't');
          t.textContent = group.t;
          r.append(t);
          p.append(r);
          group = null;
        };
        for (const pc of pieces) {
          if (pc.fld) { flush(); if (pc.fld) p.append(pc.fld.cloneNode(true)); continue; }
          if (pc.br) { flush(); const br = aEl(doc, 'br'); const base = pc.run != null ? runRPr(pc.run) : endPr; if (base) br.append(cloneRPr(base, doc)); p.append(br); continue; }
          const key = JSON.stringify(pc.fmt);
          if (group && group.run === pc.run && group.key === key) group.t += pc.t;
          else { flush(); group = { run: pc.run, key, fmt: pc.fmt, t: pc.t }; }
        }
        flush();
        if (endPr) p.append(endPr.cloneNode(true));
        else if (!p.getElementsByTagNameNS(NS.a, 'r').length) p.append(aEl(doc, 'endParaRPr', { lang: 'en-US' }));
        tx.append(p);
      }
      if (!kids(tx, 'p').length) tx.append(aEl(doc, 'p'));
      this.changed();
    }

    // Formatting commands while editing text (applied with execCommand, read back on commit).
    format(cmd, value) {
      if (!this.editingEl) return false;
      document.execCommand('styleWithCSS', false, true);
      if (cmd === 'size') {
        document.execCommand('fontSize', false, '7');
        this.editingEl.querySelectorAll('[style*="xxx-large"], font[size="7"]').forEach((el) => { el.style.fontSize = `${value * this.k * 12700}px`; el.removeAttribute('size'); });
      } else if (cmd === 'align') {
        const sel = getSelection();
        const ps = [...this.editingEl.querySelectorAll('.pp-p')].filter((p) => sel.rangeCount && sel.getRangeAt(0).intersectsNode(p));
        for (const p of ps) { p.style.textAlign = value; p.dataset.alignChanged = '1'; }
      } else document.execCommand(cmd, false, value);
      this.changed();
      return true;
    }

    // ---------------------------------------------------------------- inserting

    async insertTextBox() {
      const slide = this.deck.slides[this.current];
      this.snapshot();
      const doc = slide.doc;
      const id = this.nextId(slide);
      const cx = this.deck.cx, cy = this.deck.cy;
      const xml = `<p:sp xmlns:p="${NS.p}" xmlns:a="${NS.a}"><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(cx * 0.3)}" y="${Math.round(cy * 0.4)}"/><a:ext cx="${Math.round(cx * 0.4)}" cy="${Math.round(cy * 0.1)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Text</a:t></a:r></a:p></p:txBody></p:sp>`;
      this.tree(slide).append(doc.importNode(parseXml(xml).documentElement, true));
      this.markSlide(slide);
      await this.afterInsert();
    }

    async insertShape(prst) {
      const slide = this.deck.slides[this.current];
      this.snapshot();
      const doc = slide.doc;
      const id = this.nextId(slide);
      const cx = this.deck.cx, cy = this.deck.cy;
      const xml = `<p:sp xmlns:p="${NS.p}" xmlns:a="${NS.a}"><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(cx * 0.35)}" y="${Math.round(cy * 0.35)}"/><a:ext cx="${Math.round(cx * 0.25)}" cy="${Math.round(cy * 0.2)}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
      this.tree(slide).append(doc.importNode(parseXml(xml).documentElement, true));
      this.markSlide(slide);
      await this.afterInsert();
    }

    async insertPicture(bytes, ext, wPx, hPx) {
      const deck = this.deck;
      const slide = deck.slides[this.current];
      this.snapshot();
      let n = 1;
      while (deck.zip.file(`ppt/media/image${n}.${ext}`) || deck.newFiles.has(`ppt/media/image${n}.${ext}`)) n++;
      const media = `ppt/media/image${n}.${ext}`;
      deck.newFiles.set(media, new Uint8Array(bytes));
      await this.ensureDefault(ext, ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`);
      const rid = deck.addRel(slide.path, REL.image, relative(slide.path, media));
      slide.rels = await deck.rels(slide.path);
      const maxW = deck.cx * 0.6, maxH = deck.cy * 0.6;
      let w = wPx * 9525, h = hPx * 9525;
      const s = Math.min(1, maxW / w, maxH / h);
      w *= s; h *= s;
      const id = this.nextId(slide);
      const xml = `<p:pic xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round((deck.cx - w) / 2)}" y="${Math.round((deck.cy - h) / 2)}"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      this.tree(slide).append(slide.doc.importNode(parseXml(xml).documentElement, true));
      this.markSlide(slide);
      await this.afterInsert();
    }

    async afterInsert() {
      this.changed();
      await this.renderCurrent();
      this.refreshThumb(this.current);
      const all = this.slideEl.querySelectorAll('.pp-own');
      if (all.length) this.select(all[all.length - 1]);
    }

    async ensureDefault(ext, type) {
      const ct = await this.deck.doc('[Content_Types].xml');
      const root = ct.documentElement;
      if ([...root.children].some((d) => d.localName === 'Default' && (d.getAttribute('Extension') || '').toLowerCase() === ext)) return;
      const d = ct.createElementNS(NS.ct, 'Default');
      d.setAttribute('Extension', ext); d.setAttribute('ContentType', type);
      root.insertBefore(d, root.firstChild);
      this.deck.dirty.add('[Content_Types].xml');
    }
    async addOverride(part, type) {
      const ct = await this.deck.doc('[Content_Types].xml');
      const root = ct.documentElement;
      if ([...root.children].some((d) => d.getAttribute('PartName') === '/' + part)) return;
      const o = ct.createElementNS(NS.ct, 'Override');
      o.setAttribute('PartName', '/' + part); o.setAttribute('ContentType', type);
      root.append(o);
      this.deck.dirty.add('[Content_Types].xml');
    }
    async removeOverride(part) {
      const ct = await this.deck.doc('[Content_Types].xml');
      for (const o of [...ct.documentElement.children]) if (o.getAttribute('PartName') === '/' + part) o.remove();
      this.deck.dirty.add('[Content_Types].xml');
    }

    // ---------------------------------------------------------------- slides

    async layouts() {
      const slide = this.deck.slides[this.current] || this.deck.slides[0];
      const out = [];
      for (const r of slide.masterRels.filter((x) => x.type === REL.layout)) {
        const d = await this.deck.doc(r.path);
        const name = d && d.getElementsByTagNameNS(NS.p, 'cSld')[0] ? d.getElementsByTagNameNS(NS.p, 'cSld')[0].getAttribute('name') : r.path;
        out.push({ path: r.path, name: name || r.path.split('/').pop() });
      }
      return out;
    }

    async addSlideDialog() {
      const layouts = await this.layouts();
      if (!layouts.length) return this.addSlide(null);
      const items = layouts.map((l) => [l.name, '', () => this.addSlide(l.path)]);
      const r = this.side.getBoundingClientRect();
      SheetView.showMenu(r.right - 20, r.top + 40, items);
    }

    newSlidePath() {
      let n = this.deck.slides.length + 1;
      while (this.deck.zip.file(`ppt/slides/slide${n}.xml`) || this.deck.docs.has(`ppt/slides/slide${n}.xml`) && this.deck.docs.get(`ppt/slides/slide${n}.xml`)) n++;
      return `ppt/slides/slide${n}.xml`;
    }

    async registerSlide(path, at) {
      const deck = this.deck;
      const rid = deck.addRel(deck.presPath, REL.slide, relative(deck.presPath, path));
      const lst = deck.pres.getElementsByTagNameNS(NS.p, 'sldIdLst')[0];
      let id = 255;
      for (const s of lst.children) id = Math.max(id, +s.getAttribute('id'));
      const sldId = pEl(deck.pres, 'sldId', { id: id + 1 });
      sldId.setAttributeNS(NS.r, 'r:id', rid);
      lst.insertBefore(sldId, lst.children[at] || null);
      deck.dirty.add(deck.presPath);
      await this.addOverride(path, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
    }

    async addSlide(layoutPath) {
      const deck = this.deck;
      this.commitText();
      this.snapshot();
      const ref = deck.slides[this.current] || deck.slides[0];
      layoutPath = layoutPath || ref.layoutPath;
      const layout = await deck.doc(layoutPath);
      const path = this.newSlidePath();
      // Empty placeholders for the layout's title/body/picture placeholders (not date/footer/number).
      let id = 2;
      const phs = [];
      for (const sp of layout.getElementsByTagNameNS(NS.p, 'sp')) {
        const ph = deck.phOf(sp);
        if (!ph || ['dt', 'ftr', 'sldNum'].includes(ph.type)) continue;
        const name = (path1(sp, 'nvSpPr', 'cNvPr') || { getAttribute: () => 'Placeholder' }).getAttribute('name');
        phs.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${escapeXml(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${ph.type !== 'obj' ? ` type="${ph.type}"` : ''}${ph.idx != null ? ` idx="${ph.idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`);
      }
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${phs.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
      deck.docs.set(path, parseXml(xml));
      deck.dirty.add(path);
      deck.docs.delete(path.replace(/([^/]+)$/, '_rels/$1.rels'));
      deck.addRel(path, REL.layout, relative(path, layoutPath));
      await this.registerSlide(path, this.current + 1);
      await this.reloadSlides(this.current + 1);
    }

    async duplicateSlide(i) {
      const deck = this.deck;
      this.commitText();
      this.snapshot();
      const src = deck.slides[i];
      const path = this.newSlidePath();
      deck.docs.set(path, parseXml(serialize(src.doc)));
      deck.dirty.add(path);
      // Same relationships (layout, pictures, charts...), except speaker notes, which belong to the original.
      const relsPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
      const relsDoc = parseXml(`<Relationships xmlns="${NS.pkg}"/>`);
      for (const r of src.rels) {
        if (r.type === REL.notes) continue;
        const el = relsDoc.createElementNS(NS.pkg, 'Relationship');
        el.setAttribute('Id', r.id); el.setAttribute('Type', r.type);
        el.setAttribute('Target', r.mode === 'External' ? r.target : relative(path, r.path));
        if (r.mode) el.setAttribute('TargetMode', r.mode);
        relsDoc.documentElement.append(el);
      }
      deck.docs.set(relsPath, relsDoc);
      deck.dirty.add(relsPath);
      await this.registerSlide(path, i + 1);
      await this.reloadSlides(i + 1);
    }

    async deleteSlide(i) {
      const deck = this.deck;
      if (deck.slides.length < 2) { this.opts.error && this.opts.error("A presentation needs at least one slide, so the last slide can't be deleted."); return; }
      this.commitText();
      this.snapshot();
      const s = deck.slides[i];
      const rid = s.sldIdEl.getAttributeNS(NS.r, 'id') || s.sldIdEl.getAttribute('r:id');
      s.sldIdEl.remove();
      const presRels = deck.relsDoc(deck.presPath);
      for (const r of [...presRels.documentElement.children]) if (r.getAttribute('Id') === rid) r.remove();
      deck.dirty.add(deck.presPath);
      deck.dirty.add(deck.presPath.replace(/([^/]+)$/, '_rels/$1.rels'));
      // Remove the slide part (and its notes) from the package.
      for (const p of [s.path, s.notesPath].filter(Boolean)) {
        deck.removed.add(p); deck.removed.add(p.replace(/([^/]+)$/, '_rels/$1.rels'));
        deck.dirty.delete(p);
        await this.removeOverride(p);
      }
      await this.reloadSlides(Math.min(i, deck.slides.length - 2));
    }

    async moveSlide(from, to) {
      const deck = this.deck;
      this.commitText();
      this.snapshot();
      const lst = deck.pres.getElementsByTagNameNS(NS.p, 'sldIdLst')[0];
      const el = deck.slides[from].sldIdEl;
      el.remove();
      lst.insertBefore(el, lst.children[to] || null);
      deck.dirty.add(deck.presPath);
      await this.reloadSlides(to);
    }

    async reloadSlides(current) {
      await this.deck.loadSlides();
      this.current = Math.max(0, Math.min(current, this.deck.slides.length - 1));
      this.changed();
      await this.renderAll();
    }

    async refreshThumb(i) {
      const t = this.side.children[i];
      const s = this.deck.slides[i];
      if (!t || !s) return;
      const el = await new PptxEditor.Renderer(this.deck, { slideNumber: i + 1 }).slide(s, 150, false);
      const old = t.querySelector('.pp-slide');
      if (old) old.replaceWith(el);
    }

    // ---------------------------------------------------------------- speaker notes

    async notesBody(i, create) {
      const deck = this.deck;
      const s = deck.slides[i];
      if (!s) return null;
      if (!s.notesPath && create) {
        const presRels = await deck.rels(deck.presPath);
        const nm = presRels.find((r) => r.type === REL.notesMaster);
        if (!nm) return null;
        let n = 1;
        while (deck.zip.file(`ppt/notesSlides/notesSlide${n}.xml`) || deck.docs.get(`ppt/notesSlides/notesSlide${n}.xml`)) n++;
        const path = `ppt/notesSlides/notesSlide${n}.xml`;
        const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:notes xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
        deck.docs.set(path, parseXml(xml));
        deck.dirty.add(path);
        deck.addRel(path, REL.notesMaster, relative(path, nm.path));
        deck.addRel(path, REL.slide, relative(path, s.path));
        deck.addRel(s.path, REL.notes, relative(s.path, path));
        await this.addOverride(path, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml');
        s.notesPath = path;
        s.rels = await deck.rels(s.path);
      }
      if (!s.notesPath) return null;
      const doc = await deck.doc(s.notesPath);
      const body = [...doc.getElementsByTagNameNS(NS.p, 'sp')].find((sp) => { const ph = deck.phOf(sp); return ph && ph.type === 'body'; });
      return body ? { doc, body, path: s.notesPath } : null;
    }

    async getNotes(i) {
      const n = await this.notesBody(i, false);
      if (!n) return '';
      return kids(first(n.body, 'txBody'), 'p').map((p) => [...p.getElementsByTagNameNS(NS.a, 't')].map((t) => t.textContent).join('')).join('\n');
    }

    async setNotes(i, text) {
      if (text === await this.getNotes(i)) return;
      this.snapshot();
      const n = await this.notesBody(i, true);
      if (!n) { this.opts.error && this.opts.error("This presentation has no notes master, so notes can't be added."); return; }
      const doc = n.doc;
      let tx = first(n.body, 'txBody');
      if (!tx) { tx = pEl(doc, 'txBody'); tx.append(aEl(doc, 'bodyPr'), aEl(doc, 'lstStyle')); n.body.append(tx); }
      const old = kids(tx, 'p');
      const rPr = old.map((p) => first(first(p, 'r'), 'rPr')).find(Boolean);
      old.forEach((p) => p.remove());
      for (const line of text.split('\n')) {
        const p = aEl(doc, 'p');
        if (line) { const r = aEl(doc, 'r'); r.append(rPr ? rPr.cloneNode(true) : aEl(doc, 'rPr', { lang: 'en-US' })); const t = aEl(doc, 't'); t.textContent = line; r.append(t); p.append(r); }
        else p.append(aEl(doc, 'endParaRPr', { lang: 'en-US' }));
        tx.append(p);
      }
      this.deck.dirty.add(n.path);
      this.changed();
    }

    // ---------------------------------------------------------------- undo (XML snapshots)

    snapshot() {
      const deck = this.deck;
      const docs = new Map();
      for (const [p, d] of deck.docs) docs.set(p, d ? serialize(d) : null);
      this.history.push({ docs, dirty: new Set(deck.dirty), removed: new Set(deck.removed), newFiles: new Map(deck.newFiles), current: this.current });
      if (this.history.length > 50) this.history.shift();
      this.future = [];
    }
    stateNow() {
      const deck = this.deck;
      const docs = new Map();
      for (const [p, d] of deck.docs) docs.set(p, d ? serialize(d) : null);
      return { docs, dirty: new Set(deck.dirty), removed: new Set(deck.removed), newFiles: new Map(deck.newFiles), current: this.current };
    }
    async restoreState(st) {
      const deck = this.deck;
      deck.docs = new Map([...st.docs].map(([p, s]) => [p, s ? parseXml(s) : null]));
      deck.dirty = new Set(st.dirty); deck.removed = new Set(st.removed); deck.newFiles = new Map(st.newFiles);
      deck.pres = deck.docs.get(deck.presPath);
      await deck.loadSlides();
      this.current = Math.min(st.current, deck.slides.length - 1);
      this.changed();
      await this.renderAll();
    }
    async undo() { this.commitText(); const st = this.history.pop(); if (!st) return; this.future.push(this.stateNow()); await this.restoreState(st); }
    async redo() { const st = this.future.pop(); if (!st) return; this.history.push(this.stateNow()); await this.restoreState(st); }

    changed() { if (this.opts.onChange) this.opts.onChange(); }

    // ---------------------------------------------------------------- saving

    async save() {
      this.commitText();
      const deck = this.deck;
      const zip = await JSZip.loadAsync(deck.bytes);
      for (const p of deck.dirty) { const d = deck.docs.get(p); if (d) zip.file(p, serialize(d)); }
      for (const [p, d] of deck.docs) if (/_rels\//.test(p) && d && deck.dirty.has(p)) zip.file(p, serialize(d));
      for (const [p, b] of deck.newFiles) zip.file(p, b);
      for (const p of deck.removed) if (zip.file(p)) zip.remove(p);
      const out = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      deck.bytes = out;
      deck.zip = await JSZip.loadAsync(out);
      deck.dirty.clear(); deck.removed.clear(); deck.newFiles.clear();
      return out;
    }

    destroy() {
      document.removeEventListener('keydown', this.keyHandler);
      if (this.resizeObs) this.resizeObs.disconnect();
      this.host.replaceChildren();
    }
  }

  // ---------------------------------------------------------------- run formatting from the DOM

  function cloneRPr(rPr, doc) {
    const c = rPr.cloneNode(true);
    c.removeAttribute('dirty');
    c.removeAttribute('err');
    return doc.importNode ? c : c;
  }

  function fmtOf(textNode, para) {
    const fmt = {};
    for (let el = textNode.parentElement; el && el !== para; el = el.parentElement) {
      const st = el.style;
      const orig = el.dataset.os !== undefined ? parseCss(el.dataset.os) : el.dataset.o ? null : {};
      if (orig === null) continue; // an original element other than a run span: no changes there
      const changed = (prop) => (orig[prop] || '') !== (st[prop] || '');
      const set = (k, v) => { if (fmt[k] === undefined) fmt[k] = v; };
      if (el.tagName === 'B' || el.tagName === 'STRONG') set('b', true);
      if (el.tagName === 'I' || el.tagName === 'EM') set('i', true);
      if (el.tagName === 'U') set('u', true);
      if (el.tagName === 'S' || el.tagName === 'STRIKE') set('strike', true);
      if (st.fontWeight && changed('fontWeight')) set('b', st.fontWeight === 'bold' || +st.fontWeight >= 600);
      if (st.fontStyle && changed('fontStyle')) set('i', st.fontStyle === 'italic');
      const deco = st.textDecorationLine || st.textDecoration;
      if (deco !== undefined && changed('textDecoration')) { set('u', /underline/.test(deco || '')); set('strike', /line-through/.test(deco || '')); }
      if (st.color && changed('color')) set('color', st.color);
      if (st.fontSize && changed('fontSize')) set('sizePx', parseFloat(st.fontSize));
      if (st.fontFamily && changed('fontFamily')) set('font', st.fontFamily.split(',')[0].replace(/["']/g, '').trim());
      if (el.dataset.os !== undefined) break; // reached the run span
    }
    return fmt;
  }
  function parseCss(text) { const d = document.createElement('span'); d.style.cssText = text; const st = d.style; return { fontWeight: st.fontWeight, fontStyle: st.fontStyle, textDecoration: st.textDecorationLine || st.textDecoration, color: st.color, fontSize: st.fontSize, fontFamily: st.fontFamily }; }

  function applyRunFormat(rPr, fmt, doc, k) {
    if (fmt.b !== undefined) rPr.setAttribute('b', fmt.b ? '1' : '0');
    if (fmt.i !== undefined) rPr.setAttribute('i', fmt.i ? '1' : '0');
    if (fmt.u !== undefined) rPr.setAttribute('u', fmt.u ? 'sng' : 'none');
    if (fmt.strike !== undefined) rPr.setAttribute('strike', fmt.strike ? 'sngStrike' : 'noStrike');
    if (fmt.sizePx !== undefined && k) rPr.setAttribute("sz", String(Math.round(fmt.sizePx / (k * 12700) * 100)));
    if (fmt.color !== undefined) {
      const m = /rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(fmt.color);
      const hex = m ? [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase() : fmt.color.replace('#', '').toUpperCase();
      for (const f of [...rPr.childNodes].filter((x) => ['solidFill', 'gradFill', 'noFill'].includes(x.localName))) f.remove();
      const fill = aEl(doc, 'solidFill');
      fill.append(aEl(doc, 'srgbClr', { val: hex }));
      // a:rPr children order: ln, fills, effects, highlight, ..., latin, ea, cs, sym, hlinkClick...
      const after = [...rPr.childNodes].find((x) => !['ln'].includes(x.localName));
      rPr.insertBefore(fill, after || null);
    }
    if (fmt.font !== undefined) {
      let latin = [...rPr.childNodes].find((x) => x.localName === 'latin');
      if (!latin) {
        latin = aEl(doc, 'latin');
        const before = [...rPr.childNodes].find((x) => ['ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'].includes(x.localName));
        rPr.insertBefore(latin, before || null);
      }
      latin.setAttribute('typeface', fmt.font);
    }
  }

  function escapeXml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); }

  async function blank() {
    const res = await fetch('templates/blank.pptx');
    const buf = await res.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const core = await zip.file('docProps/core.xml').async('string');
    zip.file('docProps/core.xml', core.replace('</cp:coreProperties>', `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`));
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }

  return { Editor, blank };
})();
