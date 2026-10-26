'use strict';
// Picture editing (images open with View and Edit), like a phone's gallery editor: crop and rotate (with aspect
// ratios, flip and straighten), adjust (light and colour), filters, draw, text, remove background (BgRemove) and
// resize. Saving can change the file type (PNG, JPEG, WebP, BMP).
//
// The picture is kept as a canvas (base). Crop, rotate, drawing, text, background removal and resize make a new
// base; light, colour and filters are settings applied on top, so they stay adjustable until saving. A base is
// never changed in place, so Undo can keep earlier ones (within a memory budget).

const ImageEditor = (() => {
  // ---------------------------------------------------------------- canvas helpers

  const canvas = (w, h) => { const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w)); c.height = Math.max(1, Math.round(h)); return c; };
  const ctx = (c) => c.getContext('2d', { willReadFrequently: true });
  function copy(src) { const c = canvas(src.width, src.height); ctx(c).drawImage(src, 0, 0); return c; }
  // High-quality downscale: halve repeatedly, then the last step.
  function scaled(src, w, h) {
    let cur = src;
    while (cur.width / 2 > w && cur.height / 2 > h) {
      const half = canvas(cur.width / 2, cur.height / 2);
      const g = ctx(half); g.imageSmoothingQuality = 'high'; g.drawImage(cur, 0, 0, half.width, half.height);
      cur = half;
    }
    const out = canvas(w, h);
    const g = ctx(out); g.imageSmoothingQuality = 'high'; g.drawImage(cur, 0, 0, out.width, out.height);
    return out;
  }
  async function decode(bytes, ext) {
    const blob = new Blob([bytes], { type: FORMATS.mimeOf(ext) });
    try {
      const bmp = await createImageBitmap(blob);
      const c = canvas(bmp.width, bmp.height); ctx(c).drawImage(bmp, 0, 0); bmp.close();
      return c;
    } catch {
      const url = URL.createObjectURL(blob);
      try { const img = new Image(); img.src = url; await img.decode(); const c = canvas(img.naturalWidth || 800, img.naturalHeight || 600); ctx(c).drawImage(img, 0, 0, c.width, c.height); return c; } finally { URL.revokeObjectURL(url); }
    }
  }
  // 24-bit BMP (canvas can't write BMP itself).
  function bmpBytes(c) {
    const w = c.width, h = c.height, d = ctx(c).getImageData(0, 0, w, h).data;
    const row = (w * 3 + 3) & ~3, size = 54 + row * h;
    const buf = new ArrayBuffer(size), v = new DataView(buf), u = new Uint8Array(buf);
    u[0] = 0x42; u[1] = 0x4d; v.setUint32(2, size, true); v.setUint32(10, 54, true);
    v.setUint32(14, 40, true); v.setInt32(18, w, true); v.setInt32(22, h, true); v.setUint16(26, 1, true); v.setUint16(28, 24, true);
    v.setUint32(34, row * h, true); v.setInt32(38, 2835, true); v.setInt32(42, 2835, true);
    for (let y = 0; y < h; y++) {
      let o = 54 + (h - 1 - y) * row;
      for (let x = 0; x < w; x++) {
        const i = (y * w + x) * 4, a = d[i + 3] / 255;
        // No transparency in BMP: transparent parts become white.
        u[o++] = Math.round(d[i + 2] * a + 255 * (1 - a)); u[o++] = Math.round(d[i + 1] * a + 255 * (1 - a)); u[o++] = Math.round(d[i] * a + 255 * (1 - a));
      }
    }
    return buf;
  }
  const WRITES = ['png', 'jpg', 'jpeg', 'webp', 'bmp'];
  // A canvas as file bytes. JPEG and BMP have no transparency, so transparent parts become white.
  async function encode(c, ext, quality) {
    ext = String(ext || 'png').toLowerCase();
    if (ext === 'bmp') return bmpBytes(c);
    const type = { jpg: 'image/jpeg', jpeg: 'image/jpeg', jfif: 'image/jpeg', webp: 'image/webp' }[ext] || 'image/png';
    let src = c;
    if (type === 'image/jpeg') { src = canvas(c.width, c.height); const g = ctx(src); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(c, 0, 0); }
    const blob = await new Promise((r) => src.toBlob(r, type, quality ?? 0.92));
    return blob.arrayBuffer();
  }

  // ---------------------------------------------------------------- light, colour and filters

  const ADJUST = [
    ['brightness', 'Brightness', -100, 100], ['contrast', 'Contrast', -100, 100], ['saturation', 'Saturation', -100, 100],
    ['warmth', 'Warmth', -100, 100], ['tint', 'Tint', -100, 100], ['highlights', 'Highlights', -100, 100],
    ['shadows', 'Shadows', -100, 100], ['sharpness', 'Sharpness', 0, 100], ['vignette', 'Vignette', 0, 100],
  ];
  const noAdjust = () => Object.fromEntries(ADJUST.map(([k]) => [k, 0]));
  const lumaOf = (r, g, b) => 0.299 * r + 0.587 * g + 0.114 * b;
  const sat = (k) => (r, g, b) => { const l = lumaOf(r, g, b); return [l + (r - l) * k, l + (g - l) * k, l + (b - l) * k]; };
  const con = (k) => (r, g, b) => [(r - 128) * k + 128, (g - 128) * k + 128, (b - 128) * k + 128];
  const both = (f1, f2) => (r, g, b) => f2(...f1(r, g, b));
  const sepia = (r, g, b) => [0.393 * r + 0.769 * g + 0.189 * b, 0.349 * r + 0.686 * g + 0.168 * b, 0.272 * r + 0.534 * g + 0.131 * b];
  const FILTERS = {
    none: ['Original', null],
    vivid: ['Vivid', both(sat(1.35), con(1.12))],
    warm: ['Warm', (r, g, b) => [r * 1.08 + 8, g * 1.02, b * 0.88]],
    cool: ['Cool', (r, g, b) => [r * 0.9, g * 1.01, b * 1.08 + 8]],
    mono: ['Mono', (r, g, b) => { const l = lumaOf(r, g, b); return [l, l, l]; }],
    noir: ['Noir', both((r, g, b) => { const l = lumaOf(r, g, b); return [l, l, l]; }, con(1.45))],
    sepia: ['Sepia', sepia],
    fade: ['Fade', both(sat(0.75), (r, g, b) => [28 + r * 0.82, 28 + g * 0.82, 30 + b * 0.8])],
    dramatic: ['Dramatic', both(con(1.32), sat(0.8))],
    vintage: ['Vintage', both((r, g, b) => { const s = sepia(r, g, b); return [r * 0.45 + s[0] * 0.55, g * 0.45 + s[1] * 0.55, b * 0.45 + s[2] * 0.55]; }, (r, g, b) => [20 + r * 0.88, 18 + g * 0.88, 14 + b * 0.85])],
  };
  const smooth = (a, b, x) => { const t = Math.max(0, Math.min(1, (x - a) / (b - a))); return t * t * (3 - 2 * t); };
  const isPlain = (p) => p.filter === 'none' && ADJUST.every(([k]) => !p.adjust[k]);

  // The picture with the settings applied (a new canvas the size of src).
  function processed(src, p) {
    if (isPlain(p)) return src;
    const w = src.width, h = src.height;
    const out = copy(src);
    const g = ctx(out);
    const img = g.getImageData(0, 0, w, h), d = img.data;
    const A = p.adjust, f = FILTERS[p.filter] && FILTERS[p.filter][1], amt = (p.amount ?? 100) / 100;
    const br = A.brightness / 100 * 70;
    const ck = A.contrast >= 0 ? 1 + A.contrast / 100 * 1.2 : 1 + A.contrast / 100 * 0.85;
    const sk = A.saturation >= 0 ? 1 + A.saturation / 100 : 1 + A.saturation / 100;
    const wm = A.warmth / 100 * 32, tn = A.tint / 100 * 26, hl = A.highlights / 100 * 70, sh = A.shadows / 100 * 70;
    const vg = A.vignette / 100, cx = w / 2, cy = h / 2, rmax = Math.hypot(cx, cy);
    for (let y = 0, i = 0; y < h; y++) {
      for (let x = 0; x < w; x++, i += 4) {
        let r = d[i], gg = d[i + 1], b = d[i + 2];
        if (f) { const q = f(r, gg, b); r += (q[0] - r) * amt; gg += (q[1] - gg) * amt; b += (q[2] - b) * amt; }
        if (br) { r += br; gg += br; b += br; }
        if (hl || sh) {
          const L = lumaOf(r, gg, b) / 255;
          const add = hl * smooth(0.45, 1, L) + sh * (1 - smooth(0, 0.55, L));
          r += add; gg += add; b += add;
        }
        if (ck !== 1) { r = (r - 128) * ck + 128; gg = (gg - 128) * ck + 128; b = (b - 128) * ck + 128; }
        if (sk !== 1) { const l = lumaOf(r, gg, b); r = l + (r - l) * sk; gg = l + (gg - l) * sk; b = l + (b - l) * sk; }
        if (wm) { r += wm; b -= wm; }
        if (tn) { gg -= tn; r += tn * 0.4; b += tn * 0.4; }
        if (vg) { const k = 1 - vg * 0.75 * smooth(0.35, 1, Math.hypot(x - cx, y - cy) / rmax); r *= k; gg *= k; b *= k; }
        d[i] = r; d[i + 1] = gg; d[i + 2] = b; // Uint8ClampedArray clamps
      }
    }
    if (A.sharpness > 0) sharpen(img, w, h, A.sharpness / 100 * 1.6);
    g.putImageData(img, 0, 0);
    return out;
  }
  // Unsharp mask with a 3x3 blur.
  function sharpen(img, w, h, k) {
    const s = new Uint8ClampedArray(img.data), d = img.data;
    for (let y = 1; y < h - 1; y++) for (let x = 1; x < w - 1; x++) {
      const i = (y * w + x) * 4;
      for (let c = 0; c < 3; c++) {
        const j = i + c;
        const blur = (s[j - 4 - w * 4] + s[j - w * 4] + s[j + 4 - w * 4] + s[j - 4] + s[j] + s[j + 4] + s[j - 4 + w * 4] + s[j + w * 4] + s[j + 4 + w * 4]) / 9;
        d[j] = s[j] + (s[j] - blur) * k;
      }
    }
  }

  // Straightening rotates the picture and zooms it just enough that no empty corners show.
  function straightened(src, deg) {
    if (!deg) return src;
    const t = Math.abs(deg) * Math.PI / 180, w = src.width, h = src.height;
    const s = Math.max((w * Math.cos(t) + h * Math.sin(t)) / w, (w * Math.sin(t) + h * Math.cos(t)) / h);
    const out = canvas(w, h), g = ctx(out);
    g.imageSmoothingQuality = 'high';
    g.translate(w / 2, h / 2); g.rotate(deg * Math.PI / 180); g.scale(s, s);
    g.drawImage(src, -w / 2, -h / 2);
    return out;
  }

  // ---------------------------------------------------------------- icons

  const ICON = {
    crop: '<path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M18 22V8a2 2 0 0 0-2-2H2"/>',
    adjust: '<circle cx="12" cy="12" r="9"/><path d="M12 3v18" /><path d="M12 3a9 9 0 0 1 0 18" fill="currentColor"/>',
    filters: '<circle cx="9" cy="9" r="6"/><circle cx="15" cy="9" r="6"/><circle cx="12" cy="15" r="6"/>',
    draw: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/>',
    text: '<path d="M4 7V4h16v3"/><path d="M9 20h6"/><path d="M12 4v16"/>',
    background: '<path d="M12 3l1.9 5.8L20 10l-5 3.6L16.8 20 12 16.6 7.2 20 9 13.6 4 10l6.1-1.2Z"/>',
    resize: '<path d="M15 3h6v6"/><path d="M9 21H3v-6"/><path d="M21 3l-7 7"/><path d="M3 21l7-7"/>',
    undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>',
    redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>',
    reset: '<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/>',
    rotl: '<path d="M2.5 2v6h6"/><path d="M2.66 15.57a10 10 0 1 0 .57-8.38"/>',
    rotr: '<path d="M21.5 2v6h-6"/><path d="M21.34 15.57a10 10 0 1 1-.57-8.38"/>',
    fliph: '<path d="M12 3v18"/><path d="M8 7 3 12l5 5Z"/><path d="m16 7 5 5-5 5Z"/>',
    flipv: '<path d="M3 12h18"/><path d="M7 8l5-5 5 5Z"/><path d="m7 16 5 5 5-5Z"/>',
  };
  const icon = (name) => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ICON[name]}</svg>`;
  const iconBtn = (name, title, onclick, cls) => { const b = h('button', { type: 'button', class: `ghost ie-icon ${cls || ''}`, title, 'aria-label': title, onclick }); b.innerHTML = icon(name); return b; };
  function slider(label, min, max, value, oninput, onchange) {
    const out = h('output', { text: String(value) });
    const input = h('input', { type: 'range', min, max, step: 1, value });
    input.addEventListener('input', () => { out.textContent = input.value; oninput(+input.value); });
    if (onchange) input.addEventListener('change', () => onchange(+input.value));
    input.addEventListener('dblclick', () => { input.value = 0; out.textContent = '0'; oninput(0); if (onchange) onchange(0); });
    return { el: h('label', { class: 'ie-slider' }, h('span', {}, h('span', { text: label }), out), input), input, set: (v) => { input.value = v; out.textContent = String(v); } };
  }
  const SWATCHES = ['#000000', '#ffffff', '#e5292f', '#f59e0b', '#facc15', '#22c55e', '#2f6fdb', '#a855f7'];
  function swatches(value, onpick) {
    const wrap = h('div', { class: 'ie-swatches' });
    const custom = h('input', { type: 'color', value, title: 'Other colour' });
    const paint = (v) => { wrap.querySelectorAll('.ie-swatch').forEach((s) => s.classList.toggle('on', s.dataset.c === v)); };
    for (const c of SWATCHES) wrap.append(h('button', { type: 'button', class: 'ie-swatch', 'data-c': c, title: c, style: `background:${c}`, onclick: () => { paint(c); onpick(c); } }));
    custom.addEventListener('input', () => { paint(custom.value); onpick(custom.value); });
    wrap.append(custom);
    paint(value);
    return wrap;
  }

  // ---------------------------------------------------------------- the editor

  const TOOLS = [['crop', 'Crop'], ['adjust', 'Adjust'], ['filters', 'Filters'], ['draw', 'Draw'], ['text', 'Text'], ['background', 'Background'], ['resize', 'Resize']];

  class Editor {
    constructor(host, opts) {
      this.host = host;
      this.opts = opts || {};
      this.past = [];
      this.future = [];
      this.tool = 'crop';
      this.params = { adjust: noAdjust(), filter: 'none', amount: 100 };
      this.stage = h('div', { class: 'ie-stage' });
      this.frame = h('div', { class: 'ie-frame' });
      this.view = canvas(10, 10);
      this.view.className = 'ie-canvas';
      this.overlay = h('div', { class: 'ie-overlay' });
      this.frame.append(this.view, this.overlay);
      this.stage.append(this.frame);
      this.tabs = h('div', { class: 'ie-tabs', role: 'tablist' });
      this.panel = h('div', { class: 'ie-panel' });
      this.side = h('div', { class: 'ie-side' }, this.tabs, this.panel);
      this.root = h('div', { class: 'ie', tabindex: '-1' }, this.stage, this.side);
      host.replaceChildren(this.root);
      for (const [id, label] of TOOLS) {
        const b = h('button', { type: 'button', class: 'ie-tab', role: 'tab', 'data-tool': id, title: label, onclick: () => this.setTool(id) });
        b.innerHTML = `${icon(id)}<span>${label}</span>`;
        this.tabs.append(b);
      }
      // Undo, Redo and Revert go on Doc Studio's toolbar.
      this.undoBtn = iconBtn('undo', 'Undo (Ctrl+Z)', () => this.undo());
      this.redoBtn = iconBtn('redo', 'Redo (Ctrl+Y)', () => this.redo());
      this.resetBtn = iconBtn('reset', 'Revert to the original picture', () => this.revert());
      this.tools = [this.undoBtn, this.redoBtn, this.resetBtn];
      this.root.addEventListener('keydown', (e) => {
        if (e.target.matches('input, textarea, select')) return;
        if (e.ctrlKey && !e.shiftKey && e.key.toLowerCase() === 'z') { e.preventDefault(); e.stopPropagation(); this.undo(); }
        else if (e.ctrlKey && (e.key.toLowerCase() === 'y' || (e.shiftKey && e.key.toLowerCase() === 'z'))) { e.preventDefault(); e.stopPropagation(); this.redo(); }
      });
      new ResizeObserver(() => this.fit()).observe(this.stage);
    }

    async load(bytes, ext) {
      this.base = await decode(bytes, ext);
      this.original = this.base;
      this.past = []; this.future = [];
      this.baseChanged();
      this.setTool('adjust');
    }
    get width() { return this.base.width; }
    get height() { return this.base.height; }

    // ------------------------------------------------ state and history

    snapshot() { return { base: this.base, params: JSON.parse(JSON.stringify(this.params)) }; }
    // Keep earlier pictures within about 600 MB (each pixel is 4 bytes); settings-only steps share the picture.
    trim() {
      const limit = Math.max(4, Math.floor(1.5e8 / (this.base.width * this.base.height)));
      while (this.past.length > limit) this.past.shift();
    }
    push() { this.past.push(this.snapshot()); this.future = []; this.trim(); this.paintHistory(); }
    restore(s) {
      const changed = s.base !== this.base;
      if (changed) this.crop = null; // a frame drawn on the other picture no longer fits
      this.base = s.base;
      this.params = JSON.parse(JSON.stringify(s.params));
      if (changed) this.baseChanged(); else this.render();
      this.setTool(this.tool, true);
      this.changed();
    }
    undo() { if (this.cancelSession()) return; if (!this.past.length) return; this.future.push(this.snapshot()); this.restore(this.past.pop()); this.paintHistory(); }
    redo() { if (!this.future.length) return; this.past.push(this.snapshot()); this.restore(this.future.pop()); this.paintHistory(); }
    async revert() {
      if (this.base === this.original && isPlain(this.params)) return;
      const r = await desk.confirm({ type: 'warning', message: 'Revert to the original picture?', detail: 'Every change made in this editor is undone (Undo can bring them back).', buttons: ['Revert', 'Cancel'], defaultId: 1, cancelId: 1 });
      if (r !== 0) return;
      this.cancelSession();
      this.push();
      this.restore({ base: this.original, params: { adjust: noAdjust(), filter: 'none', amount: 100 } });
    }
    paintHistory() { this.undoBtn.disabled = !this.past.length && !this.session; this.redoBtn.disabled = !this.future.length; }
    // A new picture: commit it (with Undo) and redraw.
    commit(next) { this.push(); this.base = next; this.baseChanged(); this.changed(); }
    setParams(fn, record) { if (record) this.push(); fn(this.params); this.render(); this.changed(); }
    changed() { if (this.opts.onChange) this.opts.onChange(); if (this.opts.onInfo) this.opts.onInfo(`${this.base.width} × ${this.base.height} px`); }

    // ------------------------------------------------ drawing on screen

    baseChanged() {
      // A smaller copy for the live preview (at most 1600 px on the long side).
      const k = Math.min(1, 1600 / Math.max(this.base.width, this.base.height));
      this.small = k < 1 ? scaled(this.base, this.base.width * k, this.base.height * k) : this.base;
      this.render();
      this.paintHistory();
    }
    // Source shown on screen: the background session's cut-out while it is open, otherwise the picture.
    previewSource() { return this.session && this.session.preview ? this.session.preview : this.small; }
    render() {
      if (!this.base) return;
      let src = this.previewSource();
      if (this.tool === 'crop' && this.crop) src = straightened(src, this.crop.angle);
      const out = this.tool === 'crop' ? src : processed(src, this.params);
      if (this.view.width !== out.width || this.view.height !== out.height) { this.view.width = out.width; this.view.height = out.height; }
      const g = ctx(this.view);
      g.clearRect(0, 0, out.width, out.height);
      g.drawImage(out, 0, 0);
      this.fit();
    }
    fit() {
      if (!this.view.width) return;
      const pad = 24, sw = Math.max(50, this.stage.clientWidth - pad * 2), sh = Math.max(50, this.stage.clientHeight - pad * 2);
      const k = Math.min(sw / this.view.width, sh / this.view.height, 4);
      this.frame.style.width = `${Math.round(this.view.width * k)}px`;
      this.frame.style.height = `${Math.round(this.view.height * k)}px`;
      if (this.layoutOverlay) this.layoutOverlay();
    }
    // Pointer position in picture pixels.
    toImage(e) {
      const r = this.frame.getBoundingClientRect();
      return { x: (e.clientX - r.left) / r.width * this.base.width, y: (e.clientY - r.top) / r.height * this.base.height };
    }
    // Final picture: the base with the settings, at full size.
    renderFull() { return processed(this.base, this.params); }
    async export(ext, quality) { return encode(this.renderFull(), ext, quality); }

    // ------------------------------------------------ tools

    setTool(id, keep) {
      if (!keep && this.tool !== id) { this.applyCrop(); this.cancelSession(); }
      this.tool = id;
      this.tabs.querySelectorAll('.ie-tab').forEach((b) => b.classList.toggle('on', b.dataset.tool === id));
      this.overlay.replaceChildren();
      this.overlay.className = 'ie-overlay';
      this.layoutOverlay = null;
      this.frame.onpointerdown = null;
      this.panel.replaceChildren();
      this[`${id}Panel`]();
      this.render();
    }
    section(title, ...kids) { return h('div', { class: 'ie-section' }, title ? h('div', { class: 'ie-section-title', text: title }) : '', ...kids); }

    // Crop & rotate
    cropPanel() {
      if (!this.crop) this.crop = { angle: 0, rect: { x: 0, y: 0, w: this.base.width, h: this.base.height }, ratio: 0 };
      const ratios = [['Free', 0], ['Original', this.base.width / this.base.height], ['Square', 1], ['4:3', 4 / 3], ['3:4', 3 / 4], ['16:9', 16 / 9], ['9:16', 9 / 16]];
      const chips = h('div', { class: 'ie-chips' }, ...ratios.map(([l, r]) => h('button', { type: 'button', class: `ghost ie-chip${this.crop.ratio === r ? ' on' : ''}`, text: l, onclick: (e) => { this.crop.ratio = r; e.currentTarget.parentNode.querySelectorAll('.ie-chip').forEach((c) => c.classList.toggle('on', c === e.currentTarget)); this.fitRatio(); } })));
      const straight = slider('Straighten (°)', -45, 45, this.crop.angle, (v) => { this.crop.angle = v; this.render(); });
      this.panel.append(
        this.section('Rotate and flip', h('div', { class: 'ie-row' },
          iconBtn('rotl', 'Rotate left', () => this.rotate(-1)), iconBtn('rotr', 'Rotate right', () => this.rotate(1)),
          iconBtn('fliph', 'Flip left to right', () => this.flip(true)), iconBtn('flipv', 'Flip upside down', () => this.flip(false)))),
        this.section('Straighten', straight.el),
        this.section('Crop', chips, h('p', { class: 'ie-hint', text: 'Drag the frame or its corners on the picture.' })),
        h('div', { class: 'ie-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Reset', onclick: () => { this.crop = null; this.setTool('crop', true); } }), h('button', { type: 'button', text: 'Apply crop', onclick: () => { this.applyCrop(); this.setTool('crop', true); } })));
      // The frame: dark outside, handles on the corners and sides.
      const box = h('div', { class: 'ie-crop' });
      for (const p of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) box.append(h('span', { class: `ie-handle ${p}`, 'data-h': p }));
      this.overlay.append(box);
      this.overlay.classList.add('cropping');
      this.layoutOverlay = () => {
        if (!this.crop) return; // just applied; the panel is about to be rebuilt
        const r = this.crop.rect, W = this.base.width, H = this.base.height;
        Object.assign(box.style, { left: `${r.x / W * 100}%`, top: `${r.y / H * 100}%`, width: `${r.w / W * 100}%`, height: `${r.h / H * 100}%` });
      };
      this.layoutOverlay();
      box.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        const which = e.target.dataset.h || 'move';
        const start = this.toImage(e), r0 = { ...this.crop.rect }, W = this.base.width, H = this.base.height, min = Math.max(16, Math.min(W, H) * 0.03);
        box.setPointerCapture(e.pointerId);
        const move = (ev) => {
          const p = this.toImage(ev), dx = p.x - start.x, dy = p.y - start.y;
          let { x, y, w, h: hh } = r0;
          if (which === 'move') { x = Math.min(Math.max(0, x + dx), W - w); y = Math.min(Math.max(0, y + dy), H - hh); }
          else {
            if (which.includes('w')) { x = Math.min(Math.max(0, r0.x + dx), r0.x + r0.w - min); w = r0.x + r0.w - x; }
            if (which.includes('e')) w = Math.min(Math.max(min, r0.w + dx), W - r0.x);
            if (which.includes('n')) { y = Math.min(Math.max(0, r0.y + dy), r0.y + r0.h - min); hh = r0.y + r0.h - y; }
            if (which.includes('s')) hh = Math.min(Math.max(min, r0.h + dy), H - r0.y);
            if (this.crop.ratio) {
              // Keep the ratio: follow the side being dragged, then fit inside the picture.
              const rr = this.crop.ratio;
              if (which === 'n' || which === 's') w = hh * rr; else hh = w / rr;
              if (x + w > W) { w = W - x; hh = w / rr; }
              if (y + hh > H) { hh = H - y; w = hh * rr; }
              if (which.includes('w')) x = r0.x + r0.w - w;
              if (which.includes('n')) y = r0.y + r0.h - hh;
            }
          }
          this.crop.rect = { x, y, w, h: hh };
          this.layoutOverlay();
        };
        const up = () => { box.removeEventListener('pointermove', move); box.removeEventListener('pointerup', up); };
        box.addEventListener('pointermove', move);
        box.addEventListener('pointerup', up);
      });
    }
    // The largest frame of the chosen ratio, centred.
    fitRatio() {
      const W = this.base.width, H = this.base.height, r = this.crop.ratio;
      if (!r) { this.crop.rect = { x: 0, y: 0, w: W, h: H }; } else {
        let w = W, hh = W / r;
        if (hh > H) { hh = H; w = H * r; }
        this.crop.rect = { x: (W - w) / 2, y: (H - hh) / 2, w, h: hh };
      }
      if (this.layoutOverlay) this.layoutOverlay();
    }
    applyCrop() {
      const c = this.crop;
      if (!c) return;
      const full = c.rect.x < 0.5 && c.rect.y < 0.5 && Math.abs(c.rect.w - this.base.width) < 1 && Math.abs(c.rect.h - this.base.height) < 1;
      if (!c.angle && full) return;
      const src = straightened(this.base, c.angle);
      const out = canvas(c.rect.w, c.rect.h);
      ctx(out).drawImage(src, c.rect.x, c.rect.y, c.rect.w, c.rect.h, 0, 0, out.width, out.height);
      this.crop = null;
      this.commit(out);
    }
    rotate(dir) {
      this.applyCrop();
      const b = this.base, out = canvas(b.height, b.width), g = ctx(out);
      g.translate(out.width / 2, out.height / 2); g.rotate(dir * Math.PI / 2); g.drawImage(b, -b.width / 2, -b.height / 2);
      this.crop = null;
      this.commit(out);
      this.setTool('crop', true);
    }
    flip(horizontal) {
      this.applyCrop();
      const b = this.base, out = canvas(b.width, b.height), g = ctx(out);
      if (horizontal) { g.translate(b.width, 0); g.scale(-1, 1); } else { g.translate(0, b.height); g.scale(1, -1); }
      g.drawImage(b, 0, 0);
      this.crop = null;
      this.commit(out);
      this.setTool('crop', true);
    }

    // Adjust
    adjustPanel() {
      const sliders = ADJUST.map(([k, label, min, max]) => slider(label, min, max, this.params.adjust[k], (v) => { this.params.adjust[k] = v; this.renderSoon(); }, () => this.changed()));
      // One Undo step per slider drag: remember the settings when a drag starts.
      sliders.forEach((s) => s.input.addEventListener('pointerdown', () => this.push()));
      this.panel.append(this.section('Light and colour', ...sliders.map((s) => s.el), h('p', { class: 'ie-hint', text: 'Double-click a slider to reset it.' })),
        h('div', { class: 'ie-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Reset all', onclick: () => { this.setParams((p) => { p.adjust = noAdjust(); }, true); this.setTool('adjust', true); } })));
    }
    renderSoon() { if (this.raf) return; this.raf = requestAnimationFrame(() => { this.raf = 0; this.render(); }); }

    // Filters: a row of previews, plus how strong.
    filtersPanel() {
      const thumb = scaled(this.small, Math.max(1, Math.round(this.small.width * 96 / Math.max(this.small.width, this.small.height))), Math.max(1, Math.round(this.small.height * 96 / Math.max(this.small.width, this.small.height))));
      const grid = h('div', { class: 'ie-filters' });
      for (const [id, [label]] of Object.entries(FILTERS)) {
        const c = processed(thumb, { adjust: noAdjust(), filter: id, amount: 100 });
        c.className = 'ie-filter-thumb';
        grid.append(h('button', { type: 'button', class: `ghost ie-filter${this.params.filter === id ? ' on' : ''}`, onclick: () => { this.setParams((p) => { p.filter = id; }, true); this.setTool('filters', true); } }, c, h('span', { text: label })));
      }
      const amount = slider('Strength', 0, 100, this.params.amount, (v) => { this.params.amount = v; this.renderSoon(); }, () => this.changed());
      amount.input.addEventListener('pointerdown', () => this.push());
      this.panel.append(this.section('Filters', grid), this.params.filter !== 'none' ? this.section('', amount.el) : '');
    }

    // Draw: pen and highlighter; each stroke is one Undo step.
    drawPanel() {
      this.pen = this.pen || { kind: 'pen', color: '#e5292f', size: 8 };
      const kinds = h('div', { class: 'ie-chips' }, ...[['pen', 'Pen'], ['marker', 'Highlighter']].map(([k, l]) => h('button', { type: 'button', class: `ghost ie-chip${this.pen.kind === k ? ' on' : ''}`, text: l, onclick: () => { this.pen.kind = k; this.setTool('draw', true); } })));
      const size = slider('Size', 1, 60, this.pen.size, (v) => { this.pen.size = v; });
      this.panel.append(this.section('Draw', kinds), this.section('Colour', swatches(this.pen.color, (c) => { this.pen.color = c; })), this.section('', size.el), h('p', { class: 'ie-hint', text: 'Draw on the picture. Undo removes the last stroke.' }));
      this.overlay.classList.add('drawing');
      const live = canvas(10, 10);
      live.className = 'ie-live';
      this.overlay.append(live);
      this.frame.onpointerdown = (e) => {
        if (e.button !== 0) return;
        e.preventDefault();
        const r = this.frame.getBoundingClientRect();
        live.width = Math.round(r.width); live.height = Math.round(r.height);
        const k = r.width / this.base.width;
        const pts = [this.toImage(e)];
        const lw = this.pen.size * Math.max(this.base.width, this.base.height) / 1000 * (this.pen.kind === 'marker' ? 3 : 1);
        const style = (g, scale) => { g.lineCap = 'round'; g.lineJoin = 'round'; g.strokeStyle = this.pen.color; g.lineWidth = Math.max(1, lw * scale); g.globalAlpha = this.pen.kind === 'marker' ? 0.38 : 1; };
        const stroke = (g, scale) => { g.beginPath(); pts.forEach((p, i) => (i ? g.lineTo(p.x * scale, p.y * scale) : g.moveTo(p.x * scale, p.y * scale))); if (pts.length === 1) g.lineTo(pts[0].x * scale + 0.01, pts[0].y * scale); g.stroke(); };
        const g = ctx(live);
        const paint = () => { g.clearRect(0, 0, live.width, live.height); style(g, k); stroke(g, k); };
        paint();
        this.frame.setPointerCapture(e.pointerId);
        const move = (ev) => { pts.push(this.toImage(ev)); paint(); };
        const up = () => {
          this.frame.removeEventListener('pointermove', move); this.frame.removeEventListener('pointerup', up);
          g.clearRect(0, 0, live.width, live.height);
          const out = copy(this.base), go = ctx(out);
          style(go, 1); stroke(go, 1);
          this.commit(out);
        };
        this.frame.addEventListener('pointermove', move);
        this.frame.addEventListener('pointerup', up);
      };
    }

    // Text: type it, click the picture to place it, then Add.
    textPanel() {
      this.txt = this.txt || { text: '', size: 48, color: '#ffffff', bold: true, box: false, x: 0.5, y: 0.5 };
      const t = this.txt;
      const area = h('textarea', { rows: 3, placeholder: 'Type your text' });
      area.value = t.text;
      const size = slider('Size', 8, 300, t.size, (v) => { t.size = v; paintText(); });
      const bold = h('input', { type: 'checkbox' }); bold.checked = t.bold;
      const boxed = h('input', { type: 'checkbox' }); boxed.checked = t.box;
      const preview = h('div', { class: 'ie-text-preview' });
      this.overlay.append(preview);
      this.overlay.classList.add('texting');
      const paintText = () => {
        preview.textContent = t.text || 'Your text';
        preview.classList.toggle('empty', !t.text);
        const k = this.frame.clientWidth / this.base.width;
        Object.assign(preview.style, { left: `${t.x * 100}%`, top: `${t.y * 100}%`, fontSize: `${t.size * Math.max(this.base.width, this.base.height) / 1000 * k}px`, color: t.color, fontWeight: t.bold ? '700' : '400', background: t.box ? (t.color === '#000000' ? 'rgba(255,255,255,.85)' : 'rgba(0,0,0,.6)') : 'transparent', textShadow: t.box ? 'none' : '0 1px 3px rgba(0,0,0,.6)' });
      };
      this.layoutOverlay = paintText;
      area.addEventListener('input', () => { t.text = area.value; paintText(); });
      bold.addEventListener('change', () => { t.bold = bold.checked; paintText(); });
      boxed.addEventListener('change', () => { t.box = boxed.checked; paintText(); });
      this.frame.onpointerdown = (e) => {
        if (e.button !== 0) return;
        e.preventDefault();
        const r = this.frame.getBoundingClientRect();
        const place = (ev) => { t.x = Math.min(1, Math.max(0, (ev.clientX - r.left) / r.width)); t.y = Math.min(1, Math.max(0, (ev.clientY - r.top) / r.height)); paintText(); };
        place(e);
        this.frame.setPointerCapture(e.pointerId);
        const up = () => { this.frame.removeEventListener('pointermove', place); this.frame.removeEventListener('pointerup', up); };
        this.frame.addEventListener('pointermove', place);
        this.frame.addEventListener('pointerup', up);
      };
      this.panel.append(
        this.section('Text', area, h('p', { class: 'ie-hint', text: 'Click or drag on the picture to place it.' })),
        this.section('Colour', swatches(t.color, (c) => { t.color = c; paintText(); })),
        this.section('', size.el, h('label', { class: 'check-row' }, bold, h('span', { text: 'Bold' })), h('label', { class: 'check-row' }, boxed, h('span', { text: 'Background behind the text' }))),
        h('div', { class: 'ie-actions' }, h('button', { type: 'button', text: 'Add to picture', onclick: () => { if (!t.text.trim()) { area.focus(); return; } this.bakeText(); t.text = ''; this.setTool('text', true); } })));
      paintText();
      area.focus();
    }
    bakeText() {
      const t = this.txt, out = copy(this.base), g = ctx(out);
      const px = t.size * Math.max(out.width, out.height) / 1000;
      g.font = `${t.bold ? 700 : 400} ${px}px "Segoe UI", sans-serif`;
      g.textBaseline = 'middle'; g.textAlign = 'center';
      const lines = t.text.split('\n'), lh = px * 1.2, cx = t.x * out.width, cy = t.y * out.height;
      const top = cy - (lines.length - 1) * lh / 2;
      if (t.box) {
        const w = Math.max(...lines.map((l) => g.measureText(l).width)) + px * 0.6, hh = lines.length * lh + px * 0.3;
        g.fillStyle = t.color === '#000000' ? 'rgba(255,255,255,.85)' : 'rgba(0,0,0,.6)';
        g.fillRect(cx - w / 2, cy - hh / 2, w, hh);
      } else { g.shadowColor = 'rgba(0,0,0,.6)'; g.shadowBlur = px * 0.08; g.shadowOffsetY = px * 0.03; }
      g.fillStyle = t.color;
      lines.forEach((l, i) => g.fillText(l, cx, top + i * lh));
      this.commit(out);
    }

    // Background: remove it with the AI cut-out, then touch up with Erase / Restore and choose what goes behind.
    backgroundPanel() {
      const s = this.session;
      if (!s) {
        this.panel.append(this.section('Remove background',
          h('p', { class: 'ie-hint', text: 'Finds the main subject (a person, pet or object) and removes everything behind it. It runs on this PC; nothing is uploaded.' }),
          h('button', { type: 'button', text: 'Remove background', onclick: () => this.removeBackground() })));
        return;
      }
      const fill = h('div', { class: 'ie-chips' }, ...[['transparent', 'Transparent'], ['#ffffff', 'White'], ['color', 'Colour…']].map(([v, l]) => h('button', { type: 'button', class: `ghost ie-chip${(v === 'color' ? !['transparent', '#ffffff'].includes(s.fill) : s.fill === v) ? ' on' : ''}`, text: l, onclick: () => { if (v === 'color') picker.click(); else { s.fill = v; this.composite(); this.setTool('background', true); } } })));
      const picker = h('input', { type: 'color', value: /^#/.test(s.fill) && s.fill !== '#ffffff' ? s.fill : '#2f6fdb', class: 'ie-hidden-color' });
      picker.addEventListener('input', () => { s.fill = picker.value; this.composite(); });
      picker.addEventListener('change', () => this.setTool('background', true));
      const modes = h('div', { class: 'ie-chips' }, ...[['erase', 'Erase more'], ['restore', 'Bring back']].map(([v, l]) => h('button', { type: 'button', class: `ghost ie-chip${s.brush === v ? ' on' : ''}`, text: l, onclick: () => { s.brush = v; this.setTool('background', true); } })));
      const size = slider('Brush size', 4, 200, s.size, (v) => { s.size = v; });
      this.panel.append(
        this.section('Behind the subject', fill, picker),
        this.section('Touch up', modes, size.el, h('p', { class: 'ie-hint', text: 'Paint on the picture to erase more or bring parts back.' })),
        h('div', { class: 'ie-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Cancel', onclick: () => { this.cancelSession(); this.setTool('background', true); } }), h('button', { type: 'button', text: 'Apply', onclick: () => this.applySession() })));
      this.overlay.classList.add('brushing');
      this.frame.onpointerdown = (e) => {
        if (e.button !== 0) return;
        e.preventDefault();
        const mg = ctx(s.mask);
        const r = Math.max(2, s.size * Math.max(this.base.width, this.base.height) / 1000);
        let last = this.toImage(e);
        const dab = (p) => {
          mg.save();
          mg.globalCompositeOperation = s.brush === 'erase' ? 'destination-out' : 'source-over';
          mg.strokeStyle = '#000'; mg.lineCap = 'round'; mg.lineWidth = r * 2;
          mg.beginPath(); mg.moveTo(last.x, last.y); mg.lineTo(p.x + 0.01, p.y); mg.stroke();
          mg.restore();
          last = p;
          this.compositeSoon();
        };
        dab(last);
        this.frame.setPointerCapture(e.pointerId);
        const move = (ev) => dab(this.toImage(ev));
        const up = () => { this.frame.removeEventListener('pointermove', move); this.frame.removeEventListener('pointerup', up); this.composite(); };
        this.frame.addEventListener('pointermove', move);
        this.frame.addEventListener('pointerup', up);
      };
    }
    async removeBackground() {
      const b = h('div', { class: 'dialog busy-dialog' }, h('div', { class: 'spinner' }), h('p', { text: 'Finding the subject…' }));
      const back = h('div', { class: 'backdrop' }, b);
      document.body.append(back);
      try {
        const small = await BgRemove.mask(this.base);
        // Mask at full size, edges cleaned up (soft but not foggy).
        const mask = canvas(this.base.width, this.base.height), mg = ctx(mask);
        mg.imageSmoothingQuality = 'high';
        mg.drawImage(small, 0, 0, mask.width, mask.height);
        const id = mg.getImageData(0, 0, mask.width, mask.height), d = id.data;
        for (let i = 3; i < d.length; i += 4) d[i] = Math.max(0, Math.min(255, (d[i] - 40) * 255 / 170));
        mg.putImageData(id, 0, 0);
        this.session = { orig: this.base, mask, fill: 'transparent', brush: 'erase', size: 40 };
        this.composite();
        this.setTool('background', true);
        this.paintHistory();
      } catch (err) {
        logError(err);
        desk.error("Couldn't remove the background", err.message);
      } finally { back.remove(); }
    }
    // The cut-out: the picture through the mask, with the chosen fill behind.
    cutout(scale) {
      const s = this.session, w = Math.round(s.orig.width * scale), hh = Math.round(s.orig.height * scale);
      const out = canvas(w, hh), g = ctx(out);
      g.drawImage(s.mask, 0, 0, w, hh);
      g.globalCompositeOperation = 'source-in';
      g.drawImage(s.orig, 0, 0, w, hh);
      if (s.fill !== 'transparent') { g.globalCompositeOperation = 'destination-over'; g.fillStyle = s.fill; g.fillRect(0, 0, w, hh); }
      return out;
    }
    composite() { if (!this.session) return; this.session.preview = this.cutout(this.small.width / this.base.width); this.render(); }
    compositeSoon() { if (this.craf) return; this.craf = requestAnimationFrame(() => { this.craf = 0; this.composite(); }); }
    applySession() { const out = this.cutout(1); this.session = null; this.commit(out); this.setTool('background', true); }
    cancelSession() { if (!this.session) return false; this.session = null; this.render(); this.paintHistory(); return true; }

    // Resize
    resizePanel() {
      const W = this.base.width, H = this.base.height;
      const wIn = h('input', { type: 'number', min: '1', value: W }), hIn = h('input', { type: 'number', min: '1', value: H });
      const lock = h('input', { type: 'checkbox' }); lock.checked = true;
      wIn.addEventListener('input', () => { if (lock.checked) hIn.value = Math.max(1, Math.round(+wIn.value * H / W)); });
      hIn.addEventListener('input', () => { if (lock.checked) wIn.value = Math.max(1, Math.round(+hIn.value * W / H)); });
      const preset = (label, k) => h('button', { type: 'button', class: 'ghost ie-chip', text: label, onclick: () => { wIn.value = Math.max(1, Math.round(W * k)); hIn.value = Math.max(1, Math.round(H * k)); } });
      const longSide = (px) => Math.min(1, px / Math.max(W, H));
      this.panel.append(
        this.section('Size', h('p', { class: 'ie-hint', text: `Now ${W} × ${H} pixels.` }),
          h('div', { class: 'ie-size' }, h('label', { class: 'tool-field' }, h('span', { text: 'Width' }), wIn), h('span', { class: 'ie-x', text: '×' }), h('label', { class: 'tool-field' }, h('span', { text: 'Height' }), hIn)),
          h('label', { class: 'check-row' }, lock, h('span', { text: 'Keep the proportions' }))),
        this.section('Quick sizes', h('div', { class: 'ie-chips' }, preset('75%', 0.75), preset('50%', 0.5), preset('25%', 0.25), preset('1920 px', longSide(1920)), preset('1280 px', longSide(1280)), preset('800 px', longSide(800)))),
        h('div', { class: 'ie-actions' }, h('button', { type: 'button', text: 'Resize', onclick: () => {
          const w = Math.round(+wIn.value), hh = Math.round(+hIn.value);
          if (!(w > 0 && hh > 0) || (w === W && hh === H)) return;
          if (w * hh > 100e6) { desk.error('That size is too large', 'Choose at most 100 million pixels (for example 10000 × 10000).'); return; }
          this.commit(w < W ? scaled(this.base, w, hh) : (() => { const c = canvas(w, hh), g = ctx(c); g.imageSmoothingQuality = 'high'; g.drawImage(this.base, 0, 0, w, hh); return c; })());
          this.setTool('resize', true);
        } })));
    }
  }

  return { Editor, encode, decode, WRITES };
})();
