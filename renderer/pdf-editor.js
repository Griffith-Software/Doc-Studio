'use strict';
// PDF editor: pages drawn with pdf.js, changes written with pdf-lib.
//
// - Markup (highlight, freehand, rectangle/ellipse/line, sticky notes) becomes real PDF annotations with their own
//   appearance streams, so other PDF apps show them and can edit or delete them.
// - Added content (text, pictures, signatures, white-out, replaced text) is drawn into the page.
// - Form fields are filled through the PDF's own AcroForm.
// - Page operations (rotate, delete, insert, reorder, insert pages from another PDF) rewrite the page tree.
// Everything else in the file is carried over by pdf-lib unchanged.

const PdfEditor = (() => {
  const { pdfjs, PDFLib, fontkit } = DDPdf;
  const { PDFDocument, rgb, degrees, StandardFonts, PDFName, PDFHexString } = PDFLib;
  // Class names are minified in the bundle, so field types are found with instanceof.
  const fieldType = (f) => (f instanceof PDFLib.PDFTextField ? 'PDFTextField' : f instanceof PDFLib.PDFCheckBox ? 'PDFCheckBox' : f instanceof PDFLib.PDFRadioGroup ? 'PDFRadioGroup' : f instanceof PDFLib.PDFDropdown ? 'PDFDropdown' : f instanceof PDFLib.PDFOptionList ? 'PDFOptionList' : f instanceof PDFLib.PDFButton ? 'PDFButton' : f instanceof PDFLib.PDFSignature ? 'PDFSignature' : 'other');

  // Windows font files for fonts commonly found in PDFs: [regular, bold, italic, bold italic].
  const FONT_FILES = {
    arial: ['arial.ttf', 'arialbd.ttf', 'ariali.ttf', 'arialbi.ttf'],
    helvetica: ['arial.ttf', 'arialbd.ttf', 'ariali.ttf', 'arialbi.ttf'],
    calibri: ['calibri.ttf', 'calibrib.ttf', 'calibrii.ttf', 'calibriz.ttf'],
    'times new roman': ['times.ttf', 'timesbd.ttf', 'timesi.ttf', 'timesbi.ttf'],
    times: ['times.ttf', 'timesbd.ttf', 'timesi.ttf', 'timesbi.ttf'],
    georgia: ['georgia.ttf', 'georgiab.ttf', 'georgiai.ttf', 'georgiaz.ttf'],
    verdana: ['verdana.ttf', 'verdanab.ttf', 'verdanai.ttf', 'verdanaz.ttf'],
    tahoma: ['tahoma.ttf', 'tahomabd.ttf', 'tahoma.ttf', 'tahomabd.ttf'],
    'segoe ui': ['segoeui.ttf', 'segoeuib.ttf', 'segoeuii.ttf', 'segoeuiz.ttf'],
    'courier new': ['cour.ttf', 'courbd.ttf', 'couri.ttf', 'courbi.ttf'],
    courier: ['cour.ttf', 'courbd.ttf', 'couri.ttf', 'courbi.ttf'],
    'trebuchet ms': ['trebuc.ttf', 'trebucbd.ttf', 'trebucit.ttf', 'trebucbi.ttf'],
    garamond: ['GARA.TTF', 'GARABD.TTF', 'GARAIT.TTF', 'GARABD.TTF'],
    'book antiqua': ['BKANT.TTF', 'ANTQUAB.TTF', 'ANTQUAI.TTF', 'ANTQUABI.TTF'],
    'century gothic': ['GOTHIC.TTF', 'GOTHICB.TTF', 'GOTHICI.TTF', 'GOTHICBI.TTF'],
    'comic sans ms': ['comic.ttf', 'comicbd.ttf', 'comici.ttf', 'comicz.ttf'],
    consolas: ['consola.ttf', 'consolab.ttf', 'consolai.ttf', 'consolaz.ttf'],
    'segoe script': ['segoesc.ttf', 'segoescb.ttf', 'segoesc.ttf', 'segoescb.ttf'],
  };
  const FONT_CHOICES = ['Helvetica', 'Arial', 'Calibri', 'Times New Roman', 'Georgia', 'Verdana', 'Segoe UI', 'Courier New'];

  // "ABCDEF+Arial-BoldItalicMT" -> { family: 'arial', bold: true, italic: true }
  function parseFontName(name) {
    let n = String(name || '').replace(/^[A-Z]{6}\+/, '');
    const bold = /bold|black|heavy|semibold|demi/i.test(n);
    const italic = /italic|oblique/i.test(n);
    n = n.replace(/[-,](bold|italic|oblique|regular|roman|mt|ps|psmt|light|medium|black|semibold|condensed)+.*$/i, '').replace(/(MT|PS|PSMT)$/, '').replace(/([a-z])([A-Z])/g, '$1 $2').trim().toLowerCase();
    if (n === 'times new roman' || n.startsWith('times')) n = 'times new roman';
    if (n.startsWith('helvetica')) n = 'helvetica';
    if (n.startsWith('courier')) n = 'courier new';
    return { family: n, bold, italic };
  }

  const hexToRgb = (hex) => { const v = hex.replace('#', ''); return [0, 2, 4].map((i) => parseInt(v.slice(i, i + 2), 16) / 255); };
  const fmt = (n) => (Math.round(n * 1000) / 1000).toString();
  const clone = (x) => JSON.parse(JSON.stringify(x, (k, v) => (v instanceof Uint8Array ? { __bytes: Array.from(v) } : v)), (k, v) => (v && v.__bytes ? new Uint8Array(v.__bytes) : v));

  class Editor {
    constructor(host, opts) {
      this.host = host;
      this.opts = opts || {};
      this.scale = 1.25;
      this.tool = 'select';
      this.color = '#d32f2f';
      this.hiColor = '#ffeb3b';
      this.fontSize = 12;
      this.lineWidth = 2;
      this.font = 'Helvetica';
      this.items = [];        // pending edits, PDF coordinates
      this.formValues = new Map();
      this.history = [];
      this.future = [];
      this.fontCache = new Map();
    }

    // ---------------------------------------------------------------- loading

    async load(bytes, keepScroll) {
      const scrollTop = keepScroll ? this.main && this.main.scrollTop : 0;
      this.bytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
      if (this.task) this.task.destroy();
      this.task = pdfjs.getDocument({ data: this.bytes.slice(), isEvalSupported: false, disableFontFace: false, useSystemFonts: true, fontExtraProperties: true });
      this.pdf = await this.task.promise;
      this.lib = await PDFDocument.load(this.bytes, { ignoreEncryption: true, updateMetadata: false });
      this.encrypted = this.lib.isEncrypted;
      this.pages = [];
      for (let i = 1; i <= this.pdf.numPages; i++) {
        const page = await this.pdf.getPage(i);
        this.pages.push({ index: i - 1, page, rotation: page.rotate });
      }
      this.readForm();
      this.buildDom();
      if (keepScroll) this.main.scrollTop = scrollTop || 0;
      return this;
    }

    readForm() {
      this.fields = [];
      let form;
      try { form = this.lib.getForm(); } catch { return; }
      const pageRefs = this.lib.getPages().map((p) => p.ref);
      for (const field of form.getFields()) {
        const type = fieldType(field);
        let widgets = [];
        try { widgets = field.acroField.getWidgets(); } catch { continue; }
        widgets.forEach((w, wi) => {
          const rect = w.getRectangle();
          let pi = -1;
          const P = w.P();
          if (P) pi = pageRefs.findIndex((r) => r === P || (r.objectNumber === P.objectNumber));
          if (pi < 0) pi = this.lib.getPages().findIndex((pg) => { const annots = pg.node.Annots(); return annots && annots.asArray().some((a) => a === w.dict || (this.lib.context.lookup(a) === w.dict)); });
          if (pi < 0) return;
          let value;
          try {
            if (type === 'PDFTextField') value = field.getText() || '';
            else if (type === 'PDFCheckBox') value = field.isChecked();
            else if (type === 'PDFRadioGroup') value = field.getSelected() || '';
            else if (type === 'PDFDropdown' || type === 'PDFOptionList') value = (field.getSelected() || [])[0] || '';
          } catch { value = ''; }
          let option = null;
          if (type === 'PDFRadioGroup') { try { option = field.getOptions()[wi]; } catch { /* unknown */ } }
          this.fields.push({
            name: field.getName(), type, page: pi, rect: [rect.x, rect.y, rect.x + rect.width, rect.y + rect.height], value, option,
            options: type === 'PDFDropdown' || type === 'PDFOptionList' ? (() => { try { return field.getOptions(); } catch { return []; } })() : null,
            multiline: type === 'PDFTextField' ? (() => { try { return field.isMultiline(); } catch { return false; } })() : false,
            readOnly: (() => { try { return field.isReadOnly(); } catch { return false; } })(),
          });
        });
      }
    }

    // ---------------------------------------------------------------- DOM

    buildDom() {
      const el = (tag, cls) => { const e = document.createElement(tag); if (cls) e.className = cls; return e; };
      this.root = el('div', 'pdfe');
      this.side = el('div', 'pdfe-side');
      this.main = el('div', 'pdfe-main');
      this.root.append(this.side, this.main);
      this.host.replaceChildren(this.root);
      this.pageEls = [];
      const io = new IntersectionObserver((entries) => {
        for (const en of entries) if (en.isIntersecting) this.renderPage(+en.target.dataset.index);
      }, { root: this.main, rootMargin: '600px 0px' });
      this.io = io;
      for (const p of this.pages) {
        const vp = p.page.getViewport({ scale: this.scale });
        const wrap = el('div', 'pdfe-page');
        wrap.dataset.index = p.index;
        wrap.style.width = `${vp.width}px`;
        wrap.style.height = `${vp.height}px`;
        wrap.style.setProperty('--scale-factor', this.scale);
        wrap.style.setProperty('--total-scale-factor', this.scale);
        wrap.style.setProperty('--user-unit', 1);
        const label = el('div', 'pdfe-pageno');
        label.textContent = `${p.index + 1} / ${this.pages.length}`;
        wrap.append(label);
        this.main.append(wrap);
        this.pageEls.push(wrap);
        io.observe(wrap);
      }
      this.buildThumbs();
      this.main.addEventListener('scroll', () => this.updateCurrentPage());
      this.main.addEventListener('mouseup', (e) => this.onTextMouseUp(e));
      document.addEventListener('keydown', this.keyHandler = (e) => this.onKey(e));
    }

    buildThumbs() {
      this.side.replaceChildren();
      this.pages.forEach((p) => {
        const t = document.createElement('div');
        t.className = 'pdfe-thumb';
        t.draggable = true;
        t.dataset.index = p.index;
        const c = document.createElement('canvas');
        const vp = p.page.getViewport({ scale: 0.18 });
        c.width = vp.width; c.height = vp.height;
        p.page.render({ canvas: c, canvasContext: c.getContext('2d'), viewport: vp }).promise.catch(() => {});
        const n = document.createElement('span');
        n.textContent = p.index + 1;
        t.append(c, n);
        t.addEventListener('click', () => this.pageEls[p.index].scrollIntoView({ block: 'start' }));
        t.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/x-pdf-page', String(p.index)); t.classList.add('dragging'); });
        t.addEventListener('dragend', () => t.classList.remove('dragging'));
        t.addEventListener('dragover', (e) => { if ([...e.dataTransfer.types].includes('text/x-pdf-page')) { e.preventDefault(); t.classList.add('drop'); } });
        t.addEventListener('dragleave', () => t.classList.remove('drop'));
        t.addEventListener('drop', (e) => {
          e.preventDefault(); e.stopPropagation();
          t.classList.remove('drop');
          const from = +e.dataTransfer.getData('text/x-pdf-page');
          if (!isNaN(from) && from !== p.index && !this.readOnly) this.movePage(from, p.index);
        });
        t.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          if (!SheetView.showMenu || this.readOnly) return;
          this.current = p.index;
          SheetView.showMenu(e.clientX, e.clientY, [
            ['Rotate right', '', () => this.rotatePage(p.index, 90)],
            ['Rotate left', '', () => this.rotatePage(p.index, -90)],
            ['Insert blank page after', '', () => this.insertBlank(p.index + 1)],
            ['Insert pages from a PDF…', '', () => this.opts.pickPdf && this.opts.pickPdf().then((b) => b && this.insertFrom(b, p.index + 1))],
            null,
            ['Delete page', '', () => this.deletePage(p.index)],
          ]);
        });
        this.side.append(t);
      });
      this.updateCurrentPage();
    }

    updateCurrentPage() {
      const mid = this.main.scrollTop + this.main.clientHeight / 3;
      let cur = 0;
      for (const w of this.pageEls) if (w.offsetTop <= mid) cur = +w.dataset.index;
      this.current = cur;
      this.side.querySelectorAll('.pdfe-thumb').forEach((t) => t.classList.toggle('on', +t.dataset.index === cur));
    }

    viewport(i) { return this.pages[i].page.getViewport({ scale: this.scale }); }

    async renderPage(i) {
      const p = this.pages[i];
      const wrap = this.pageEls[i];
      if (!p || !wrap || wrap.dataset.rendered) return;
      wrap.dataset.rendered = '1';
      const vp = this.viewport(i);
      const dpr = window.devicePixelRatio || 1;
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(vp.width * dpr); canvas.height = Math.floor(vp.height * dpr);
      canvas.style.width = `${vp.width}px`; canvas.style.height = `${vp.height}px`;
      canvas.className = 'pdfe-canvas';
      const text = document.createElement('div');
      text.className = 'textLayer';
      const overlay = document.createElement('div');
      overlay.className = 'pdfe-overlay';
      const forms = document.createElement('div');
      forms.className = 'pdfe-forms';
      wrap.prepend(canvas);
      wrap.append(text, forms, overlay);
      p.canvas = canvas; p.text = text; p.overlay = overlay; p.forms = forms;
      await p.page.render({ canvas, canvasContext: canvas.getContext('2d'), viewport: vp, transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null, annotationMode: pdfjs.AnnotationMode.ENABLE }).promise;
      const tc = await p.page.getTextContent();
      p.textContent = tc;
      const layer = new pdfjs.TextLayer({ textContentSource: tc, container: text, viewport: vp });
      await layer.render();
      // Map text divs back to their text items (for "edit text").
      const items = tc.items.filter((it) => it.str !== undefined);
      let k = 0;
      for (const div of layer.textDivs) {
        while (k < items.length && items[k].str !== div.textContent) k++;
        if (k < items.length) div.dataset.item = tc.items.indexOf(items[k++]);
      }
      this.wireOverlay(i);
      this.drawForms(i);
      this.drawItems(i);
      this.applyToolClass();
    }

    rerender() {
      for (const p of this.pages) {
        const wrap = this.pageEls[p.index];
        delete wrap.dataset.rendered;
        wrap.querySelectorAll('canvas, .textLayer, .pdfe-overlay, .pdfe-forms').forEach((x) => x.remove());
        const vp = this.viewport(p.index);
        wrap.style.width = `${vp.width}px`; wrap.style.height = `${vp.height}px`;
        wrap.style.setProperty('--scale-factor', this.scale);
        wrap.style.setProperty('--total-scale-factor', this.scale);
        this.io.unobserve(wrap); this.io.observe(wrap);
      }
    }

    setZoom(scale) {
      const ratio = this.main.scrollTop / Math.max(1, this.main.scrollHeight);
      this.scale = Math.max(0.3, Math.min(4, scale));
      this.rerender();
      requestAnimationFrame(() => { this.main.scrollTop = ratio * this.main.scrollHeight; });
    }
    // Fit one page (or two side by side) to the width, no larger than 140%.
    fitWidth(n) {
      const p = this.pages[this.current || 0];
      if (!p) return;
      n = n || 1;
      const w = p.page.getViewport({ scale: 1 }).width;
      this.setZoom(Math.min(1.4, (this.main.clientWidth - 72 - 20 * (n - 1)) / (w * n)));
    }

    // ---------------------------------------------------------------- tools

    setTool(t) { if (this.readOnly && t !== 'select') return; this.tool = t; this.applyToolClass(); this.select(null); }
    // View mode: the same pages, nothing editable (form fields and markup can't be changed).
    setReadOnly(on) {
      this.readOnly = !!on;
      if (on) { this.placing = null; this.tool = 'select'; this.applyToolClass(); this.select(null); }
      if (this.root) this.root.classList.toggle('readonly', this.readOnly);
      this.side.querySelectorAll('.pdfe-thumb').forEach((t) => { t.draggable = !this.readOnly; });
    }
    // Every page as a picture (for printing), current edits included.
    async printImages(box) {
      const bytes = this.items.length || this.formValues.size || this.dirtyPages ? await this.build() : this.bytes;
      const task = pdfjs.getDocument({ data: new Uint8Array(bytes).slice(), isEvalSupported: false });
      const doc = await task.promise;
      const out = [];
      for (let i = 1; i <= doc.numPages; i++) {
        const page = await doc.getPage(i);
        const vp = page.getViewport({ scale: 2 });
        const c = document.createElement('canvas');
        c.width = vp.width; c.height = vp.height;
        await page.render({ canvas: c, canvasContext: c.getContext('2d'), viewport: vp, annotationMode: pdfjs.AnnotationMode.ENABLE_FORMS }).promise;
        const img = document.createElement('img');
        img.src = c.toDataURL('image/png');
        img.className = 'pdf-print-page';
        out.push(img);
      }
      task.destroy();
      box.replaceChildren(...out);
    }
    applyToolClass() {
      for (const w of this.pageEls) {
        w.dataset.tool = this.tool;
      }
    }

    toPdf(i, vx, vy) { return this.viewport(i).convertToPdfPoint(vx, vy); }
    toView(i, x, y) { return this.viewport(i).convertToViewportPoint(x, y); }
    rectToView(i, r) {
      const [x1, y1] = this.toView(i, r[0], r[1]);
      const [x2, y2] = this.toView(i, r[2], r[3]);
      return { left: Math.min(x1, x2), top: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) };
    }

    localPoint(i, e) {
      const r = this.pageEls[i].getBoundingClientRect();
      return [e.clientX - r.left, e.clientY - r.top];
    }

    wireOverlay(i) {
      const ov = this.pages[i].overlay;
      ov.addEventListener('pointerdown', (e) => this.onPointerDown(i, e));
      this.pages[i].text.addEventListener('click', (e) => { if (this.tool === 'edittext') this.startEditText(i, e); });
    }

    pushHistory() {
      this.history.push({ bytes: this.bytes, items: clone(this.items), forms: new Map(this.formValues) });
      if (this.history.length > 60) this.history.shift();
      this.future = [];
    }
    changed() { if (this.opts.onChange) this.opts.onChange(); }

    addItem(item) {
      this.pushHistory();
      item.id = `i${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
      this.items.push(item);
      this.drawItems(item.page);
      this.changed();
      return item;
    }

    onPointerDown(i, e) {
      if (e.button !== 0 || this.readOnly) return;
      const t = this.tool;
      const [vx, vy] = this.localPoint(i, e);
      if (t === 'select') {
        if (!e.target.closest('.pdfe-item')) this.select(null);
        return;
      }
      if (t === 'text') {
        e.preventDefault();
        const [x, y] = this.toPdf(i, vx, vy + this.fontSize * this.scale);
        const item = this.addItem({ type: 'text', page: i, x, y, size: this.fontSize, color: this.color, font: this.font, text: '' });
        this.editTextItem(item, true);
        return;
      }
      if (t === 'note') {
        e.preventDefault();
        const [x, y] = this.toPdf(i, vx, vy);
        this.opts.ask('Sticky note', [{ name: 'text', label: 'Note', value: '' }]).then((v) => {
          if (v && v.text.trim()) this.addItem({ type: 'note', page: i, x, y, text: v.text.trim(), color: '#ffd54f' });
        });
        return;
      }
      if (t === 'place' && this.placing) {
        e.preventDefault();
        const pl = this.placing;
        this.placing = null;
        const w = pl.w * this.scale, h = pl.h * this.scale;
        const [x1, y1] = this.toPdf(i, vx - w / 2, vy + h / 2);
        const [x2, y2] = this.toPdf(i, vx + w / 2, vy - h / 2);
        const item = this.addItem({ type: 'image', page: i, rect: [Math.min(x1, x2), Math.min(y1, y2), Math.max(x1, x2), Math.max(y1, y2)], png: pl.png, kind: pl.kind });
        this.setTool('select');
        this.select(item);
        if (this.opts.onToolChange) this.opts.onToolChange('select');
        return;
      }
      if (['draw', 'rect', 'ellipse', 'line', 'whiteout', 'highlightbox'].includes(t)) {
        e.preventDefault();
        const ov = this.pages[i].overlay;
        ov.setPointerCapture(e.pointerId);
        const pts = [[vx, vy]];
        const ghost = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        ghost.classList.add('pdfe-ghost');
        ov.append(ghost);
        const draw = () => {
          const [a, b] = [pts[0], pts[pts.length - 1]];
          const stroke = t === 'whiteout' ? '#999' : t === 'highlightbox' ? this.hiColor : this.color;
          if (t === 'draw') ghost.innerHTML = `<polyline points="${pts.map((p) => p.join(',')).join(' ')}" fill="none" stroke="${this.color}" stroke-width="${this.lineWidth * this.scale}" stroke-linecap="round" stroke-linejoin="round"/>`;
          else if (t === 'line') ghost.innerHTML = `<line x1="${a[0]}" y1="${a[1]}" x2="${b[0]}" y2="${b[1]}" stroke="${stroke}" stroke-width="${this.lineWidth * this.scale}"/>`;
          else if (t === 'ellipse') ghost.innerHTML = `<ellipse cx="${(a[0] + b[0]) / 2}" cy="${(a[1] + b[1]) / 2}" rx="${Math.abs(b[0] - a[0]) / 2}" ry="${Math.abs(b[1] - a[1]) / 2}" fill="none" stroke="${stroke}" stroke-width="${this.lineWidth * this.scale}"/>`;
          else ghost.innerHTML = `<rect x="${Math.min(a[0], b[0])}" y="${Math.min(a[1], b[1])}" width="${Math.abs(b[0] - a[0])}" height="${Math.abs(b[1] - a[1])}" fill="${t === 'whiteout' ? '#fff' : t === 'highlightbox' ? this.hiColor + '66' : 'none'}" stroke="${stroke}" stroke-width="${t === 'rect' ? this.lineWidth * this.scale : 1}"/>`;
        };
        const move = (ev) => { const p = this.localPoint(i, ev); if (t === 'draw') pts.push(p); else pts[1] = p; draw(); };
        const up = () => {
          ov.removeEventListener('pointermove', move); ov.removeEventListener('pointerup', up);
          ghost.remove();
          const a = pts[0], b = pts[pts.length - 1];
          if (t === 'draw') {
            if (pts.length < 2) return;
            this.addItem({ type: 'ink', page: i, paths: [pts.map((p) => this.toPdf(i, p[0], p[1]))], color: this.color, width: this.lineWidth });
            return;
          }
          if (Math.abs(b[0] - a[0]) < 3 && Math.abs(b[1] - a[1]) < 3) return;
          const p1 = this.toPdf(i, a[0], a[1]), p2 = this.toPdf(i, b[0], b[1]);
          const rect = [Math.min(p1[0], p2[0]), Math.min(p1[1], p2[1]), Math.max(p1[0], p2[0]), Math.max(p1[1], p2[1])];
          if (t === 'line') this.addItem({ type: 'line', page: i, a: p1, b: p2, color: this.color, width: this.lineWidth });
          else if (t === 'whiteout') this.addItem({ type: 'whiteout', page: i, rect });
          else if (t === 'highlightbox') this.addItem({ type: 'highlight', page: i, quads: [rect], color: this.hiColor });
          else this.addItem({ type: t, page: i, rect, color: this.color, width: this.lineWidth });
        };
        ov.addEventListener('pointermove', move); ov.addEventListener('pointerup', up);
        draw();
      }
    }

    // Highlight: select text with the mouse, then it is highlighted.
    onTextMouseUp() {
      if (this.tool !== 'highlight') return;
      const sel = getSelection();
      if (!sel.rangeCount || sel.isCollapsed) return;
      const range = sel.getRangeAt(0);
      const byPage = new Map();
      for (const r of range.getClientRects()) {
        if (r.width < 1 || r.height < 1) continue;
        const wrap = this.pageEls.find((w) => { const b = w.getBoundingClientRect(); return r.left < b.right && r.right > b.left && r.top < b.bottom && r.bottom > b.top; });
        if (!wrap) continue;
        const i = +wrap.dataset.index;
        const b = wrap.getBoundingClientRect();
        const p1 = this.toPdf(i, r.left - b.left, r.top - b.top), p2 = this.toPdf(i, r.right - b.left, r.bottom - b.top);
        if (!byPage.has(i)) byPage.set(i, []);
        byPage.get(i).push([Math.min(p1[0], p2[0]), Math.min(p1[1], p2[1]), Math.max(p1[0], p2[0]), Math.max(p1[1], p2[1])]);
      }
      const text = sel.toString();
      sel.removeAllRanges();
      for (const [i, quads] of byPage) this.addItem({ type: 'highlight', page: i, quads: mergeQuads(quads), color: this.hiColor, text });
    }

    // Edit existing text: the clicked text run becomes an editable box; on save the original is covered and the
    // new text drawn in the closest matching Windows font.
    startEditText(i, e) {
      const div = e.target.closest('.textLayer > span, .textLayer span');
      if (!div || div.dataset.item === undefined) return;
      const p = this.pages[i];
      const item = p.textContent.items[+div.dataset.item];
      if (!item || !item.str.trim()) return;
      const existing = this.items.find((x) => x.type === 'replace' && x.page === i && x.itemIndex === +div.dataset.item);
      if (existing) { this.editTextItem(existing, true); return; }
      const [a, b, , , ex, ey] = item.transform;
      const size = Math.hypot(a, b) || 10;
      const style = p.textContent.styles[item.fontName] || {};
      let fontName = '';
      try { if (p.page.commonObjs.has(item.fontName)) fontName = p.page.commonObjs.get(item.fontName).name || ''; } catch { /* not loaded */ }
      if (!fontName && style.fontFamily && !/^(sans-serif|serif|monospace)$/.test(style.fontFamily)) fontName = style.fontFamily;
      const desc = (style.descent || -0.2) * size;
      const asc = (style.ascent || 0.8) * size;
      const rect = [ex - 0.5, ey + desc - 0.5, ex + item.width + 0.5, ey + asc + 0.5];
      const color = sampleTextColor(p.canvas, this.rectToView(i, rect));
      const r = this.addItem({ type: 'replace', page: i, itemIndex: +div.dataset.item, rect, x: ex, y: ey, size, text: item.str, original: item.str, fontName, color, font: fontLabel(fontName) });
      this.editTextItem(r, true);
    }

    // ---------------------------------------------------------------- drawing pending items

    drawItems(i) {
      const p = this.pages[i];
      if (!p || !p.overlay) return;
      p.overlay.querySelectorAll('.pdfe-item').forEach((x) => x.remove());
      for (const it of this.items.filter((x) => x.page === i)) p.overlay.append(this.itemEl(it));
    }

    itemEl(it) {
      const i = it.page;
      const el = document.createElement('div');
      el.className = `pdfe-item pdfe-${it.type}`;
      el.dataset.id = it.id;
      const s = this.scale;
      const place = (box) => Object.assign(el.style, { left: `${box.left}px`, top: `${box.top}px`, width: `${box.width}px`, height: `${box.height}px` });
      if (it.type === 'text' || it.type === 'replace') {
        const lines = it.text.split('\n');
        const [vx, vy] = this.toView(i, it.x, it.y);
        if (it.type === 'replace') { const box = this.rectToView(i, it.rect); place({ ...box, width: Math.max(box.width, measure(it.text, it.size * s, it.font) + 6) }); el.style.background = '#fff'; }
        else Object.assign(el.style, { left: `${vx}px`, top: `${vy - it.size * s}px`, minWidth: '40px' });
        el.style.font = `${it.size * s}px ${cssFont(it.font, it.fontName)}`;
        el.style.color = it.color;
        el.style.lineHeight = '1.2';
        el.textContent = lines.join('\n') || ' ';
        el.addEventListener('dblclick', () => this.editTextItem(it, true));
      } else if (it.type === 'image') {
        place(this.rectToView(i, it.rect));
        const img = document.createElement('img');
        img.src = URL.createObjectURL(new Blob([it.png], { type: 'image/png' }));
        img.draggable = false;
        el.append(img);
      } else if (it.type === 'whiteout') {
        place(this.rectToView(i, it.rect));
      } else if (it.type === 'highlight') {
        el.style.left = '0'; el.style.top = '0';
        for (const q of it.quads) {
          const b = this.rectToView(i, q);
          const d = document.createElement('i');
          Object.assign(d.style, { left: `${b.left}px`, top: `${b.top}px`, width: `${b.width}px`, height: `${b.height}px`, background: it.color });
          el.append(d);
        }
      } else if (it.type === 'note') {
        const [vx, vy] = this.toView(i, it.x, it.y);
        Object.assign(el.style, { left: `${vx}px`, top: `${vy}px` });
        el.title = it.text;
        el.textContent = '✎';
      } else {
        const svgNS = 'http://www.w3.org/2000/svg';
        const svg = document.createElementNS(svgNS, 'svg');
        let box;
        if (it.type === 'ink') {
          const all = it.paths.flat().map(([x, y]) => this.toView(i, x, y));
          box = bounds(all, it.width * s + 2);
          svg.innerHTML = it.paths.map((path) => `<polyline points="${path.map(([x, y]) => { const [a, b] = this.toView(i, x, y); return `${a - box.left},${b - box.top}`; }).join(' ')}" fill="none" stroke="${it.color}" stroke-width="${it.width * s}" stroke-linecap="round" stroke-linejoin="round"/>`).join('');
        } else if (it.type === 'line') {
          const a = this.toView(i, it.a[0], it.a[1]), b = this.toView(i, it.b[0], it.b[1]);
          box = bounds([a, b], it.width * s + 2);
          svg.innerHTML = `<line x1="${a[0] - box.left}" y1="${a[1] - box.top}" x2="${b[0] - box.left}" y2="${b[1] - box.top}" stroke="${it.color}" stroke-width="${it.width * s}"/>`;
        } else {
          box = this.rectToView(i, it.rect);
          const w = it.width * s;
          svg.innerHTML = it.type === 'ellipse'
            ? `<ellipse cx="${box.width / 2}" cy="${box.height / 2}" rx="${Math.max(0, box.width / 2 - w / 2)}" ry="${Math.max(0, box.height / 2 - w / 2)}" fill="none" stroke="${it.color}" stroke-width="${w}"/>`
            : `<rect x="${w / 2}" y="${w / 2}" width="${Math.max(0, box.width - w)}" height="${Math.max(0, box.height - w)}" fill="none" stroke="${it.color}" stroke-width="${w}"/>`;
        }
        svg.setAttribute('width', box.width); svg.setAttribute('height', box.height);
        place(box);
        el.append(svg);
      }
      el.addEventListener('pointerdown', (e) => {
        if (this.tool !== 'select' || el.isContentEditable) return;
        e.stopPropagation(); e.preventDefault();
        this.select(it);
        this.dragItem(it, e);
      });
      if (this.selected === it.id) el.classList.add('sel');
      if (it.type === 'image' || it.type === 'whiteout' || it.type === 'rect' || it.type === 'ellipse') {
        const handle = document.createElement('b');
        handle.className = 'pdfe-handle';
        handle.addEventListener('pointerdown', (e) => { if (this.tool !== 'select') return; e.stopPropagation(); e.preventDefault(); this.resizeItem(it, e); });
        el.append(handle);
      }
      return el;
    }

    select(it) {
      this.selected = it ? it.id : null;
      for (const w of this.pageEls) w.querySelectorAll('.pdfe-item').forEach((x) => x.classList.toggle('sel', x.dataset.id === this.selected));
      if (this.opts.onSelect) this.opts.onSelect(it);
    }

    dragItem(it, e) {
      const i = it.page;
      const start = this.localPoint(i, e);
      const before = clone(it);
      let moved = false;
      const move = (ev) => {
        const p = this.localPoint(i, ev);
        const [dx, dy] = [p[0] - start[0], p[1] - start[1]];
        if (Math.abs(dx) + Math.abs(dy) > 2) moved = true;
        const d0 = this.toPdf(i, 0, 0), d1 = this.toPdf(i, dx, dy);
        const ddx = d1[0] - d0[0], ddy = d1[1] - d0[1];
        Object.assign(it, shifted(before, ddx, ddy));
        this.drawItems(i);
      };
      const up = () => {
        window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
        if (moved) { this.history.push({ bytes: this.bytes, items: this.items.map((x) => (x.id === it.id ? before : clone(x))), forms: new Map(this.formValues) }); this.future = []; this.changed(); }
      };
      window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
    }

    resizeItem(it, e) {
      const i = it.page;
      const before = clone(it);
      const box0 = this.rectToView(i, it.rect);
      const start = this.localPoint(i, e);
      const ratio = box0.width / Math.max(1, box0.height);
      const move = (ev) => {
        const p = this.localPoint(i, ev);
        let w = Math.max(8, box0.width + p[0] - start[0]);
        let h = Math.max(8, box0.height + p[1] - start[1]);
        if (it.type === 'image' && !ev.shiftKey) h = w / ratio;
        const a = this.toPdf(i, box0.left, box0.top), b = this.toPdf(i, box0.left + w, box0.top + h);
        it.rect = [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])];
        this.drawItems(i);
      };
      const up = () => {
        window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
        this.history.push({ bytes: this.bytes, items: this.items.map((x) => (x.id === it.id ? before : clone(x))), forms: new Map(this.formValues) });
        this.future = [];
        this.changed();
      };
      window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
    }

    editTextItem(it, focus) {
      const p = this.pages[it.page];
      const el = p.overlay.querySelector(`[data-id="${it.id}"]`);
      if (!el) return;
      el.contentEditable = 'plaintext-only';
      el.classList.add('editing');
      if (focus) {
        el.focus();
        const r = document.createRange(); r.selectNodeContents(el); r.collapse(false);
        getSelection().removeAllRanges(); getSelection().addRange(r);
      }
      const finish = () => {
        el.removeEventListener('blur', finish);
        el.contentEditable = 'false';
        el.classList.remove('editing');
        const text = el.innerText.replace(/\n$/, '');
        if (!text.trim() && it.type === 'text') { this.items = this.items.filter((x) => x !== it); }
        else if (it.type === 'replace' && text === it.original) { this.items = this.items.filter((x) => x !== it); }
        else it.text = text;
        this.drawItems(it.page);
        this.changed();
      };
      el.addEventListener('blur', finish);
      el.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.preventDefault(); el.blur(); } e.stopPropagation(); });
    }

    deleteSelected() {
      const it = this.items.find((x) => x.id === this.selected);
      if (!it) return false;
      this.pushHistory();
      this.items = this.items.filter((x) => x !== it);
      this.selected = null;
      this.drawItems(it.page);
      this.changed();
      return true;
    }

    onKey(e) {
      if (!this.root.isConnected || !this.root.offsetParent || this.readOnly) return;
      if ((e.key === 'Delete' || e.key === 'Backspace') && this.selected && !document.activeElement.isContentEditable && !['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName)) {
        if (this.deleteSelected()) e.preventDefault();
      }
      if (e.key === 'Escape' && this.placing) { this.placing = null; this.setTool('select'); if (this.opts.onToolChange) this.opts.onToolChange('select'); }
    }

    // Place an image (picture or signature) with the next click.
    beginPlace(png, w, h, kind) {
      const max = kind === 'signature' ? 200 : 300;
      const k = Math.min(1, max / w);
      this.placing = { png, w: w * k, h: h * k, kind };
      this.setTool('place');
    }

    // ---------------------------------------------------------------- forms

    drawForms(i) {
      const p = this.pages[i];
      if (!p.forms) return;
      p.forms.replaceChildren();
      for (const f of this.fields.filter((x) => x.page === i)) {
        const box = this.rectToView(i, f.rect);
        let input;
        const cur = this.formValues.has(f.name) ? this.formValues.get(f.name) : f.value;
        if (f.type === 'PDFTextField') {
          input = document.createElement(f.multiline ? 'textarea' : 'input');
          input.value = cur || '';
          input.addEventListener('input', () => this.setField(f, input.value));
        } else if (f.type === 'PDFCheckBox') {
          input = document.createElement('input'); input.type = 'checkbox'; input.checked = !!cur;
          input.addEventListener('change', () => this.setField(f, input.checked));
        } else if (f.type === 'PDFRadioGroup') {
          input = document.createElement('input'); input.type = 'radio'; input.name = `dd-radio-${f.name}`; input.checked = cur && cur === f.option;
          input.addEventListener('change', () => { if (input.checked) this.setField(f, f.option); });
        } else if (f.type === 'PDFDropdown' || f.type === 'PDFOptionList') {
          input = document.createElement('select');
          for (const o of ['', ...(f.options || [])]) { const op = document.createElement('option'); op.value = o; op.textContent = o; input.append(op); }
          input.value = cur || '';
          input.addEventListener('change', () => this.setField(f, input.value));
        } else continue;
        input.className = 'pdfe-field';
        input.title = f.name;
        input.disabled = f.readOnly;
        Object.assign(input.style, { left: `${box.left}px`, top: `${box.top}px`, width: `${box.width}px`, height: `${box.height}px`, fontSize: `${Math.min(14, box.height * 0.7)}px` });
        p.forms.append(input);
      }
    }

    setField(f, v) {
      if (!this.formTimer) this.pushHistory();
      clearTimeout(this.formTimer);
      this.formTimer = setTimeout(() => { this.formTimer = null; }, 800);
      this.formValues.set(f.name, v);
      // Other widgets of the same field (e.g. radio buttons) follow.
      if (f.type === 'PDFRadioGroup') for (const p of this.pages) if (p.forms) this.drawForms(p.index);
      this.changed();
    }

    // ---------------------------------------------------------------- undo

    async undo() {
      const prev = this.history.pop();
      if (!prev) return;
      this.future.push({ bytes: this.bytes, items: clone(this.items), forms: new Map(this.formValues) });
      await this.restore(prev);
    }
    async redo() {
      const next = this.future.pop();
      if (!next) return;
      this.history.push({ bytes: this.bytes, items: clone(this.items), forms: new Map(this.formValues) });
      await this.restore(next);
    }
    async restore(state) {
      const reload = state.bytes !== this.bytes;
      this.items = clone(state.items);
      this.formValues = new Map(state.forms);
      if (reload) await this.load(state.bytes, true);
      else for (const p of this.pages) { this.drawItems(p.index); if (p.forms) this.drawForms(p.index); }
      this.changed();
    }

    // ---------------------------------------------------------------- page operations (applied immediately)

    async pageOp(fn) {
      const bytes = await this.build();
      const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
      await fn(doc);
      const out = await doc.save();
      this.history.push({ bytes: this.bytes, items: clone(this.items), forms: new Map(this.formValues) });
      this.future = [];
      this.items = [];
      this.formValues = new Map();
      await this.load(out, true);
      this.changed();
    }
    rotatePage(i, delta) { return this.pageOp((doc) => { const p = doc.getPage(i); p.setRotation(degrees(((p.getRotation().angle + delta) % 360 + 360) % 360)); }); }
    deletePage(i) {
      if (this.pages.length < 2) { this.opts.error && this.opts.error("A PDF needs at least one page, so the last page can't be deleted."); return Promise.resolve(); }
      return this.pageOp((doc) => doc.removePage(i));
    }
    insertBlank(at) {
      return this.pageOp((doc) => {
        const ref = doc.getPage(Math.max(0, Math.min(at - 1, doc.getPageCount() - 1)));
        const { width, height } = ref.getSize();
        doc.insertPage(at, [width, height]);
      });
    }
    movePage(from, to) {
      return this.pageOp((doc) => {
        const page = doc.getPage(from);
        doc.removePage(from);
        doc.insertPage(Math.min(to, doc.getPageCount()), page);
      });
    }
    insertFrom(bytes, at) {
      return this.pageOp(async (doc) => {
        const other = await PDFDocument.load(bytes, { ignoreEncryption: true });
        const pages = await doc.copyPages(other, other.getPageIndices());
        pages.forEach((p, k) => doc.insertPage(at + k, p));
      });
    }

    // ---------------------------------------------------------------- saving

    async getFont(doc, label, fontName) {
      const key = `${label}|${fontName || ''}`;
      if (this.fontCache.has(key) && this.fontCache.get(key).doc === doc) return this.fontCache.get(key).font;
      let font = null;
      const want = fontName ? parseFontName(fontName) : parseFontName(label);
      const files = FONT_FILES[want.family];
      if (files && this.opts.readFont) {
        const file = files[(want.bold ? 1 : 0) + (want.italic ? 2 : 0)];
        const data = await this.opts.readFont(file);
        if (data) { try { font = await doc.embedFont(data, { subset: true }); } catch { font = null; } }
      }
      if (!font) {
        const serif = /times|georgia|garamond|antiqua|serif|cambria/.test(want.family), mono = /courier|consol|mono/.test(want.family);
        const base = mono ? 'Courier' : serif ? 'TimesRoman' : 'Helvetica';
        const suffix = want.bold && want.italic ? (serif ? 'BoldItalic' : 'BoldOblique') : want.bold ? 'Bold' : want.italic ? (serif ? 'Italic' : 'Oblique') : '';
        font = await doc.embedFont(StandardFonts[base + suffix] || StandardFonts.Helvetica);
      }
      this.fontCache.set(key, { doc, font });
      return font;
    }

    // Write pending items and form values into a copy of the current PDF. Returns the new bytes.
    async build() {
      if (!this.items.length && !this.formValues.size) return this.bytes;
      const doc = await PDFDocument.load(this.bytes, { ignoreEncryption: true, updateMetadata: false });
      doc.registerFontkit(fontkit);
      const pages = doc.getPages();
      for (const it of this.items) {
        const page = pages[it.page];
        if (!page) continue;
        const rot = page.getRotation().angle;
        const [r, g, b] = hexToRgb(it.color || '#000000');
        if (it.type === 'text' || it.type === 'replace') {
          const font = await this.getFont(doc, it.font, it.fontName);
          if (it.type === 'replace') page.drawRectangle({ x: it.rect[0], y: it.rect[1], width: it.rect[2] - it.rect[0], height: it.rect[3] - it.rect[1], color: rgb(1, 1, 1) });
          const lines = it.text.split('\n');
          const vp1 = this.pages[it.page].page.getViewport({ scale: 1 });
          const [bx, by] = vp1.convertToViewportPoint(it.x, it.y);
          lines.forEach((line, k) => {
            const [x, y] = vp1.convertToPdfPoint(bx, by + k * it.size * 1.2);
            page.drawText(safeText(line, font), { x, y, size: it.size, font, color: rgb(r, g, b), rotate: degrees(rot) });
          });
        } else if (it.type === 'whiteout') {
          page.drawRectangle({ x: it.rect[0], y: it.rect[1], width: it.rect[2] - it.rect[0], height: it.rect[3] - it.rect[1], color: rgb(1, 1, 1) });
        } else if (it.type === 'image') {
          const img = await doc.embedPng(it.png);
          const [x1, y1, x2, y2] = it.rect;
          drawRotatedImage(page, img, x1, y1, x2 - x1, y2 - y1, rot);
        } else {
          this.addAnnotation(doc, page, it);
        }
      }
      if (this.formValues.size) {
        const form = doc.getForm();
        for (const [name, v] of this.formValues) {
          try {
            const f = form.getField(name);
            const type = fieldType(f);
            if (type === 'PDFTextField') f.setText(String(v || ''));
            else if (type === 'PDFCheckBox') { if (v) f.check(); else f.uncheck(); }
            else if (type === 'PDFRadioGroup') { if (v) f.select(v); }
            else if (type === 'PDFDropdown' || type === 'PDFOptionList') { if (v) f.select(v); else f.clear(); }
          } catch (err) { console.warn('form field', name, err); }
        }
        try { form.updateFieldAppearances(await doc.embedFont(StandardFonts.Helvetica)); } catch (err) { console.warn('appearances', err); }
      }
      return doc.save();
    }

    // Markup annotations with appearance streams (so every viewer shows them the same way).
    addAnnotation(doc, page, it) {
      const ctx = doc.context;
      const [r, g, b] = hexToRgb(it.color || '#ffeb3b');
      const col = `${fmt(r)} ${fmt(g)} ${fmt(b)}`;
      let rect, sub, extra = {}, ops, gs = null;
      const now = `D:${new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14)}`;
      if (it.type === 'highlight') {
        const xs = it.quads.flatMap((q) => [q[0], q[2]]), ys = it.quads.flatMap((q) => [q[1], q[3]]);
        rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
        sub = 'Highlight';
        extra.QuadPoints = it.quads.flatMap((q) => [q[0], q[3], q[2], q[3], q[0], q[1], q[2], q[1]]);
        extra.CA = 1;
        gs = { BM: 'Multiply' };
        ops = `/GS0 gs ${col} rg ${it.quads.map((q) => `${fmt(q[0])} ${fmt(q[1])} ${fmt(q[2] - q[0])} ${fmt(q[3] - q[1])} re f`).join(' ')}`;
        if (it.text) extra.Contents = PDFHexString.fromText(it.text);
      } else if (it.type === 'ink') {
        const pts = it.paths.flat();
        const w = it.width;
        rect = [Math.min(...pts.map((p) => p[0])) - w, Math.min(...pts.map((p) => p[1])) - w, Math.max(...pts.map((p) => p[0])) + w, Math.max(...pts.map((p) => p[1])) + w];
        sub = 'Ink';
        extra.InkList = it.paths.map((path) => path.flat());
        extra.BS = { W: w, S: 'S' };
        ops = `${col} RG ${fmt(w)} w 1 J 1 j ${it.paths.map((path) => path.map((p, k) => `${fmt(p[0])} ${fmt(p[1])} ${k ? 'l' : 'm'}`).join(' ') + ' S').join(' ')}`;
      } else if (it.type === 'rect' || it.type === 'ellipse') {
        const w = it.width;
        rect = [it.rect[0] - w / 2, it.rect[1] - w / 2, it.rect[2] + w / 2, it.rect[3] + w / 2];
        sub = it.type === 'rect' ? 'Square' : 'Circle';
        extra.BS = { W: w, S: 'S' };
        const [x1, y1, x2, y2] = it.rect;
        if (it.type === 'rect') ops = `${col} RG ${fmt(w)} w ${fmt(x1)} ${fmt(y1)} ${fmt(x2 - x1)} ${fmt(y2 - y1)} re S`;
        else {
          const cx = (x1 + x2) / 2, cy = (y1 + y2) / 2, rx = (x2 - x1) / 2, ry = (y2 - y1) / 2, k = 0.5523;
          ops = `${col} RG ${fmt(w)} w ${fmt(cx + rx)} ${fmt(cy)} m ${fmt(cx + rx)} ${fmt(cy + ry * k)} ${fmt(cx + rx * k)} ${fmt(cy + ry)} ${fmt(cx)} ${fmt(cy + ry)} c ${fmt(cx - rx * k)} ${fmt(cy + ry)} ${fmt(cx - rx)} ${fmt(cy + ry * k)} ${fmt(cx - rx)} ${fmt(cy)} c ${fmt(cx - rx)} ${fmt(cy - ry * k)} ${fmt(cx - rx * k)} ${fmt(cy - ry)} ${fmt(cx)} ${fmt(cy - ry)} c ${fmt(cx + rx * k)} ${fmt(cy - ry)} ${fmt(cx + rx)} ${fmt(cy - ry * k)} ${fmt(cx + rx)} ${fmt(cy)} c S`;
        }
      } else if (it.type === 'line') {
        const w = it.width;
        rect = [Math.min(it.a[0], it.b[0]) - w, Math.min(it.a[1], it.b[1]) - w, Math.max(it.a[0], it.b[0]) + w, Math.max(it.a[1], it.b[1]) + w];
        sub = 'Line';
        extra.L = [it.a[0], it.a[1], it.b[0], it.b[1]];
        extra.BS = { W: w, S: 'S' };
        ops = `${col} RG ${fmt(w)} w 1 J ${fmt(it.a[0])} ${fmt(it.a[1])} m ${fmt(it.b[0])} ${fmt(it.b[1])} l S`;
      } else if (it.type === 'note') {
        rect = [it.x, it.y - 20, it.x + 20, it.y];
        sub = 'Text';
        extra.Contents = PDFHexString.fromText(it.text);
        extra.Name = PDFName.of('Comment');
        extra.Open = false;
        const [x, y] = [rect[0], rect[1]];
        ops = `${col} rg 0.4 0.3 0 RG 0.8 w ${fmt(x + 1)} ${fmt(y + 1)} 18 18 re B 0.4 0.3 0 RG ${fmt(x + 4)} ${fmt(y + 13)} m ${fmt(x + 16)} ${fmt(y + 13)} l ${fmt(x + 4)} ${fmt(y + 9)} m ${fmt(x + 16)} ${fmt(y + 9)} l ${fmt(x + 4)} ${fmt(y + 5)} m ${fmt(x + 12)} ${fmt(y + 5)} l S`;
      } else return;
      const resources = gs ? { ExtGState: { GS0: { Type: 'ExtGState', ...gs } } } : {};
      const ap = ctx.stream(ops, { Type: 'XObject', Subtype: 'Form', BBox: rect, Resources: resources });
      const apRef = ctx.register(ap);
      const annot = ctx.obj({
        Type: 'Annot', Subtype: sub, Rect: rect, C: [r, g, b], F: 4, P: page.ref,
        T: PDFHexString.fromText(this.opts.author || 'Doc Studio'), M: PDFHexString.fromText(now),
        AP: { N: apRef }, ...extra,
      });
      page.node.addAnnot(ctx.register(annot));
    }

    destroy() {
      document.removeEventListener('keydown', this.keyHandler);
      if (this.io) this.io.disconnect();
      if (this.task) this.task.destroy();
      this.host.replaceChildren();
    }
  }

  // ---------------------------------------------------------------- helpers

  function shifted(it, dx, dy) {
    const out = clone(it);
    if (out.rect) out.rect = [out.rect[0] + dx, out.rect[1] + dy, out.rect[2] + dx, out.rect[3] + dy];
    if (out.quads) out.quads = out.quads.map((q) => [q[0] + dx, q[1] + dy, q[2] + dx, q[3] + dy]);
    if (out.paths) out.paths = out.paths.map((p) => p.map(([x, y]) => [x + dx, y + dy]));
    if (out.a) { out.a = [out.a[0] + dx, out.a[1] + dy]; out.b = [out.b[0] + dx, out.b[1] + dy]; }
    if (out.x !== undefined) { out.x += dx; out.y += dy; }
    return out;
  }

  function bounds(points, pad) {
    const xs = points.map((p) => p[0]), ys = points.map((p) => p[1]);
    const left = Math.min(...xs) - pad, top = Math.min(...ys) - pad;
    return { left, top, width: Math.max(...xs) + pad - left, height: Math.max(...ys) + pad - top };
  }

  // Merge selection rectangles on the same line into one quad each.
  function mergeQuads(quads) {
    const out = [];
    for (const q of quads.sort((a, b) => b[1] - a[1] || a[0] - b[0])) {
      const last = out[out.length - 1];
      if (last && Math.abs(last[1] - q[1]) < 2 && Math.abs(last[3] - q[3]) < 2 && q[0] <= last[2] + 2) last[2] = Math.max(last[2], q[2]);
      else out.push([...q]);
    }
    return out;
  }

  // Keep only characters the font can draw (pdf-lib throws on others).
  function safeText(text, font) {
    try { font.encodeText(text); return text; } catch { /* fall through */ }
    return [...text].map((ch) => { try { font.encodeText(ch); return ch; } catch { return '?'; } }).join('');
  }

  function drawRotatedImage(page, img, x, y, w, h, rot) {
    if (!rot) { page.drawImage(img, { x, y, width: w, height: h }); return; }
    // Keep the picture upright on rotated pages: rotate about its box.
    const map = { 90: [x + w, y, h, w], 180: [x + w, y + h, w, h], 270: [x, y + h, h, w] }[rot] || [x, y, w, h];
    page.drawImage(img, { x: map[0], y: map[1], width: map[2], height: map[3], rotate: degrees(rot) });
  }

  function fontLabel(fontName) {
    const f = parseFontName(fontName);
    const hit = FONT_CHOICES.find((c) => c.toLowerCase() === f.family);
    return hit || (f.family ? f.family.replace(/\b\w/g, (m) => m.toUpperCase()) : 'Helvetica');
  }
  function cssFont(label, fontName) {
    const f = fontName ? parseFontName(fontName) : parseFontName(label);
    const fam = f.family === 'helvetica' ? 'Arial' : (label || 'Arial');
    return `${f.italic ? 'italic ' : ''}${f.bold ? 'bold ' : ''}"${fam}", Arial, sans-serif`.replace(/^(italic |bold )+/, (m) => m);
  }
  let measureCtx = null;
  function measure(text, px, label) {
    measureCtx = measureCtx || document.createElement('canvas').getContext('2d');
    measureCtx.font = `${px}px "${label || 'Arial'}", Arial, sans-serif`;
    return Math.max(...text.split('\n').map((l) => measureCtx.measureText(l).width));
  }
  // The darkest colour inside the text's box on the rendered page is (nearly always) the text colour.
  function sampleTextColor(canvas, box) {
    try {
      const k = canvas.width / parseFloat(canvas.style.width);
      const d = canvas.getContext('2d').getImageData(Math.floor(box.left * k), Math.floor(box.top * k), Math.max(1, Math.floor(box.width * k)), Math.max(1, Math.floor(box.height * k))).data;
      let best = null, lum = 999;
      for (let i = 0; i < d.length; i += 4) { const l = d[i] * 0.3 + d[i + 1] * 0.59 + d[i + 2] * 0.11; if (d[i + 3] > 200 && l < lum) { lum = l; best = [d[i], d[i + 1], d[i + 2]]; } }
      return best ? '#' + best.map((v) => v.toString(16).padStart(2, '0')).join('') : '#000000';
    } catch { return '#000000'; }
  }

  async function blank() {
    const doc = await PDFDocument.create();
    doc.addPage([612, 792]);
    doc.setCreator('Doc Studio');
    return doc.save();
  }

  // Signature pad: draw with the mouse, or type a name. Resolves to { png, w, h } or null.
  function signatureDialog() {
    return new Promise((resolve) => {
      const back = document.createElement('div');
      back.className = 'backdrop';
      const box = document.createElement('div');
      box.className = 'dialog sigpad';
      box.innerHTML = '<h3>Signature</h3><div class="seg sig-tabs"><button class="on" data-m="draw">Draw</button><button data-m="type">Type</button></div><canvas width="520" height="170"></canvas><input type="text" placeholder="Type your name" hidden><div class="dialog-actions"><button class="ghost" data-a="clear">Clear</button><button class="ghost" data-a="cancel">Cancel</button><button data-a="ok">Use signature</button></div>';
      back.append(box);
      document.body.append(back);
      const canvas = box.querySelector('canvas');
      const input = box.querySelector('input');
      const ctx = canvas.getContext('2d');
      let mode = 'draw', drawn = false;
      ctx.lineWidth = 2.4; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.strokeStyle = '#1a237e';
      const pos = (e) => { const r = canvas.getBoundingClientRect(); return [(e.clientX - r.left) * canvas.width / r.width, (e.clientY - r.top) * canvas.height / r.height]; };
      canvas.addEventListener('pointerdown', (e) => {
        if (mode !== 'draw') return;
        canvas.setPointerCapture(e.pointerId);
        const [x, y] = pos(e); ctx.beginPath(); ctx.moveTo(x, y);
        const mv = (ev) => { const [a, b] = pos(ev); ctx.lineTo(a, b); ctx.stroke(); drawn = true; };
        const up = () => { canvas.removeEventListener('pointermove', mv); canvas.removeEventListener('pointerup', up); };
        canvas.addEventListener('pointermove', mv); canvas.addEventListener('pointerup', up);
      });
      const typed = () => {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.fillStyle = '#1a237e';
        ctx.font = '64px "Segoe Script", "Brush Script MT", cursive';
        ctx.textBaseline = 'middle';
        ctx.fillText(input.value, 20, canvas.height / 2);
        drawn = !!input.value.trim();
      };
      input.addEventListener('input', typed);
      box.querySelectorAll('.sig-tabs button').forEach((b) => b.addEventListener('click', () => {
        mode = b.dataset.m;
        box.querySelectorAll('.sig-tabs button').forEach((x) => x.classList.toggle('on', x === b));
        input.hidden = mode !== 'type';
        ctx.clearRect(0, 0, canvas.width, canvas.height); drawn = false;
        if (mode === 'type') { input.focus(); typed(); }
      }));
      const done = (v) => { back.remove(); resolve(v); };
      box.querySelector('[data-a="clear"]').onclick = () => { ctx.clearRect(0, 0, canvas.width, canvas.height); drawn = false; input.value = ''; };
      box.querySelector('[data-a="cancel"]').onclick = () => done(null);
      box.querySelector('[data-a="ok"]').onclick = () => {
        if (!drawn) return done(null);
        // Trim to the ink.
        const d = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
        let x1 = canvas.width, y1 = canvas.height, x2 = 0, y2 = 0;
        for (let y = 0; y < canvas.height; y++) for (let x = 0; x < canvas.width; x++) if (d[(y * canvas.width + x) * 4 + 3] > 10) { x1 = Math.min(x1, x); y1 = Math.min(y1, y); x2 = Math.max(x2, x); y2 = Math.max(y2, y); }
        const w = x2 - x1 + 8, h = y2 - y1 + 8;
        const c2 = document.createElement('canvas'); c2.width = w; c2.height = h;
        c2.getContext('2d').drawImage(canvas, x1 - 4, y1 - 4, w, h, 0, 0, w, h);
        c2.toBlob(async (blob) => done({ png: new Uint8Array(await blob.arrayBuffer()), w: w / 2, h: h / 2 }), 'image/png');
      };
    });
  }

  return { Editor, blank, signatureDialog, FONT_CHOICES };
})();
