'use strict';
// The AI assistants' worker (ai/bridge.js): runs a job with Doc Studio's own engines and returns the result as
// base64 bytes or text. It runs in a hidden window (index.html?worker=1), so the owner's documents are not
// touched; openDocuments and readOpen run in the owner's window to see what is open there. Files are written
// by the bridge, which applies the owner's rules.

const AIWorker = (() => {
  const isWorker = /[?&]worker=1/.test(location.search);
  const extOf = (p) => (p.includes('.') ? p.slice(p.lastIndexOf('.') + 1).toLowerCase() : '');
  const nameOf = (p) => p.split(/[\\/]/).pop();
  const IMAGE_EXTS = ['png', 'jpg', 'jpeg', 'jfif', 'gif', 'webp', 'bmp', 'avif'];
  const quiet = { set() {}, close() {} }; // stands in for a progress box
  function b64(bytes) {
    const u = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let s = '';
    for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
    return btoa(s);
  }
  const text = (bytes) => (typeof bytes === 'string' ? bytes : new TextDecoder().decode(bytes));
  const clip = (s, max) => (max && s.length > max ? { text: s.slice(0, max), truncated: true, total_chars: s.length } : { text: s, truncated: false });

  // ---------------------------------------------------------------- tabs in the worker

  const opened = [];
  async function load(path) {
    const f = await desk.read(path);
    const tab = makeTab({ path: f.path, name: f.name, size: f.size, data: f.data });
    opened.push(tab);
    await setupTab(tab);
    if (tab.body.querySelector('.failure')) throw new Error(`${f.name} couldn't be opened in Doc Studio.`);
    return tab;
  }
  async function newTab(name, data) {
    const tab = makeTab({ path: null, name, size: data.byteLength, data });
    opened.push(tab);
    await setupTab(tab);
    return tab;
  }
  async function closeAll() { for (const t of opened.splice(0)) { t.dirty = false; if (tabs.includes(t)) await closeTab(t); } }

  // A document's text as Markdown (or plain text when Markdown isn't available for it).
  async function textOf(tab) {
    if (['md', 'text', 'csv', 'html'].includes(tab.kind) && tab.textarea) return tab.textarea.value;
    if (tab.kind === 'image') throw Object.assign(new Error('This is a picture. Use docstudio_ocr to read the text in it.'), { code: 'use_ocr' });
    const list = await Convert.targets(tab);
    const to = list.some((t) => t.ext === 'md') ? 'md' : list.some((t) => t.ext === 'txt') ? 'txt' : null;
    if (!to) throw new Error(`Doc Studio can't read the text of ${FORMATS.label(tab.kind)} files.`);
    return text(await Convert.run(tab, to, { office: false }));
  }

  // Markdown -> safe HTML for documents.
  const mdHtml = (md) => DOMPurify.sanitize(marked.parse(md));
  const page = (title, body) => `<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<title>${title.replace(/</g, '&lt;')}</title>\n<style>body{font:16px/1.6 "Segoe UI",sans-serif;max-width:820px;margin:40px auto;padding:0 20px;color:#1d2129}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:4px 8px}code,pre{background:#f3f4f6}pre{padding:10px;overflow:auto}</style>\n</head>\n<body>\n${body}\n</body>\n</html>\n`;

  // ---------------------------------------------------------------- jobs

  const JOBS = {
    async read({ path, max_chars }) {
      const tab = await load(path);
      return { name: tab.name, kind: FORMATS.label(tab.kind), ...clip(await textOf(tab), max_chars || 200000) };
    },
    async create({ ext, content, title }) {
      const enc = (s) => new TextEncoder().encode(s);
      if (ext === 'md' || ext === 'txt') return { data: b64(enc(content.endsWith('\n') ? content : content + '\n')) };
      if (ext === 'html') return { data: b64(enc(page(title, mdHtml(content)))) };
      if (ext === 'docx') return { data: b64(await desk.docxFromHtml(mdHtml(content), title)) };
      if (ext === 'pdf') { const tab = await newTab(`${title}.md`, enc(content).buffer); return { data: b64(await Convert.run(tab, 'pdf')) }; }
      throw new Error(`Can't create .${ext} files.`);
    },
    async convert({ path, to }) {
      const tab = await load(path);
      const list = await Convert.targets(tab);
      if (!list.some((t) => t.ext === to)) throw Object.assign(new Error(`${FORMATS.label(tab.kind)} files can't be converted to .${to}. They can become: ${list.map((t) => '.' + t.ext).join(', ') || 'nothing'}.`), { code: 'bad_format' });
      const out = await Convert.run(tab, to);
      if (out && out.files) return { files: out.files.map((f) => ({ suffix: f.suffix, data: b64(f.data) })) };
      return { data: b64(typeof out === 'string' ? new TextEncoder().encode(out) : out) };
    },
    async merge({ paths }) {
      const r = await Tools.mergeFiles(paths.map((p) => ({ name: nameOf(p), path: p })));
      return { data: b64(r.data), pages: r.pages };
    },
    async split({ path, ranges, every }) {
      const bytes = (await desk.read(path)).data;
      const { PDFDocument } = DDPdf.PDFLib;
      const n = (await PDFDocument.load(bytes, { ignoreEncryption: true })).getPageCount();
      let groups;
      if (ranges) groups = Tools.parseRanges(String(ranges), n);
      else if (every) { const k = Math.max(1, +every); groups = []; for (let i = 0; i < n; i += k) groups.push(Array.from({ length: Math.min(k, n - i) }, (_, j) => i + j)); }
      else groups = Array.from({ length: n }, (_, i) => [i]);
      const out = await Tools.splitPdf(bytes, groups);
      return { files: out.map((d, i) => ({ data: b64(d), label: groups[i].length === 1 ? `page ${groups[i][0] + 1}` : `pages ${groups[i][0] + 1}-${groups[i][groups[i].length - 1] + 1}` })) };
    },
    async compress({ path, level, quality, max_side, format }) {
      const data = (await desk.read(path)).data, ext = extOf(path);
      const q = Math.max(0.1, Math.min(1, (+quality || 75) / 100));
      let out, outExt = ext;
      if (ext === 'pdf') out = await Tools.compressPdf(data, ['lossless', 'balanced', 'small'].includes(level) ? level : 'lossless', quiet);
      else if (IMAGE_EXTS.includes(ext)) { outExt = ['jpg', 'webp', 'png'].includes(format) ? format : (ext === 'png' ? 'png' : 'jpg'); out = (await Tools.imageAs(data, ext, { jpg: 'image/jpeg', webp: 'image/webp', png: 'image/png' }[outExt], q, +max_side || 0)).data; }
      else if (['docx', 'docm', 'xlsx', 'xlsm', 'pptx', 'pptm', 'odt', 'ods', 'odp', 'epub'].includes(ext)) out = await Tools.compressZip(data, q, +max_side || 1600, quiet);
      else throw Object.assign(new Error('Doc Studio compresses PDFs, pictures and Office files.'), { code: 'bad_format' });
      const before = data.byteLength, after = out.byteLength;
      if (after >= before && outExt === ext) return { notSmaller: true, before, after };
      return { data: b64(out), ext: outExt, before, after };
    },
    async ocr({ path }) {
      const data = (await desk.read(path)).data, ext = extOf(path);
      let images;
      if (ext === 'pdf') images = (await Tools.pdfImages(data, 3, 'image/png', 1)).map((x) => ({ data: x.data, ext: 'png' }));
      else if (IMAGE_EXTS.includes(ext)) images = [{ data: (await Tools.imageAs(data, ext, 'image/png', 1, 4000, 1600)).data, ext: 'png' }];
      else throw Object.assign(new Error('OCR reads pictures and PDFs. For other documents use docstudio_read_document.'), { code: 'bad_format' });
      const texts = await desk.ocr(images);
      return { pages: texts.length, text: texts.map((t) => t.trim()).join('\n\n\f\n\n') };
    },
    // Picture edits, in order, then saved as format.
    async image({ path, steps, format }) {
      const host = h('div', { style: 'position:fixed;left:-10000px;top:0;width:1200px;height:800px' });
      document.body.append(host);
      try {
        const ed = new ImageEditor.Editor(host, {});
        await ed.load((await desk.read(path)).data, extOf(path));
        const full = () => ({ x: 0, y: 0, w: ed.width, h: ed.height });
        for (const s of steps) {
          const op = String(s.op || '').toLowerCase();
          if (op === 'rotate') { const t = ((Math.round(+s.turns || (+s.degrees || 90) / 90) % 4) + 4) % 4; for (let i = 0; i < (t === 3 ? 1 : t); i++) ed.rotate(t === 3 ? -1 : 1); }
          else if (op === 'flip') ed.flip(String(s.direction || 'horizontal').startsWith('h'));
          else if (op === 'straighten') { ed.crop = { angle: Math.max(-45, Math.min(45, +s.degrees || 0)), rect: full(), ratio: 0 }; ed.applyCrop(); }
          else if (op === 'crop') {
            let { x = 0, y = 0, width, height } = s;
            const frac = [x, y, width, height].every((v) => v === undefined || (v >= 0 && v <= 1));
            if (frac) { x *= ed.width; y *= ed.height; width = (width ?? 1) * ed.width; height = (height ?? 1) * ed.height; }
            x = Math.max(0, Math.min(ed.width - 1, x)); y = Math.max(0, Math.min(ed.height - 1, y));
            width = Math.max(1, Math.min(ed.width - x, width || ed.width)); height = Math.max(1, Math.min(ed.height - y, height || ed.height));
            ed.crop = { angle: 0, rect: { x, y, w: width, h: height }, ratio: 0 };
            ed.applyCrop();
          } else if (op === 'resize') {
            let w = +s.width || 0, hh = +s.height || 0;
            if (!w && !hh) throw new Error('resize needs a width or a height.');
            if (!w) w = Math.round(hh * ed.width / ed.height);
            if (!hh) hh = Math.round(w * ed.height / ed.width);
            if (w * hh > 100e6) throw new Error('That size is too large (at most 100 million pixels).');
            const c = document.createElement('canvas'); c.width = w; c.height = hh;
            const g = c.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(ed.base, 0, 0, w, hh);
            ed.commit(c);
          } else if (op === 'remove_background') {
            await ed.removeBackground();
            if (!ed.session) throw new Error("The background couldn't be removed from this picture.");
            const fill = String(s.fill || 'transparent').toLowerCase();
            ed.session.fill = fill === 'white' ? '#ffffff' : fill === 'black' ? '#000000' : /^#[0-9a-f]{6}$/.test(fill) ? fill : 'transparent';
            ed.applySession();
          } else if (op === 'adjust') { for (const k of Object.keys(ed.params.adjust)) if (s[k] !== undefined) ed.params.adjust[k] = Math.max(-100, Math.min(100, +s[k] || 0)); }
          else if (op === 'filter') { ed.params.filter = ['none', 'vivid', 'warm', 'cool', 'mono', 'noir', 'sepia', 'fade', 'dramatic', 'vintage'].includes(s.name) ? s.name : 'none'; ed.params.amount = Math.max(0, Math.min(100, s.amount ?? 100)); }
          else throw Object.assign(new Error(`Unknown picture step "${s.op}". Steps: rotate, flip, straighten, crop, resize, remove_background, adjust, filter.`), { code: 'bad_args' });
        }
        const outExt = ImageEditor.WRITES.includes(format) ? (format === 'jpeg' ? 'jpg' : format) : 'png';
        return { data: b64(await ed.export(outExt)), ext: outExt, width: ed.width, height: ed.height };
      } finally { host.remove(); document.querySelectorAll('.backdrop').forEach((b) => b.remove()); }
    },
    async scan({ make, color, dpi, paper, scanner }) {
      const list = await desk.scanners();
      if (!list.length) throw Object.assign(new Error('No scanner was found. Check that it is on and that Windows can see it.'), { code: 'no_scanner' });
      const dev = (scanner && list.find((d) => d.name.toLowerCase().includes(String(scanner).toLowerCase()))) || list[0];
      const r = await desk.scanPage({ id: dev.id, color: ['color', 'gray', 'bw'].includes(color) ? color : 'color', dpi: +dpi || 300, paper: ['letter', 'a4', 'glass'].includes(paper) ? paper : 'letter' });
      const pic = await ImageEditor.decode(r.data, r.ext);
      if (make === 'png' || make === 'image') return { data: b64(await ImageEditor.encode(pic, 'png')), ext: 'png' };
      if (make === 'text') return { text: (await desk.ocr([{ data: await ImageEditor.encode(pic, 'png'), ext: 'png' }]))[0] || '' };
      const { PDFDocument } = DDPdf.PDFLib;
      const doc = await PDFDocument.create();
      doc.setCreator('Doc Studio'); doc.setProducer('Doc Studio');
      const img = color === 'bw' ? await doc.embedPng(await ImageEditor.encode(pic, 'png')) : await doc.embedJpg(await ImageEditor.encode(pic, 'jpg', 0.85));
      const d = +dpi || 300, w = pic.width / d * 72, hh = pic.height / d * 72;
      doc.addPage([w, hh]).drawImage(img, { x: 0, y: 0, width: w, height: hh });
      return { data: b64(await doc.save()), ext: 'pdf' };
    },
  };

  async function run(op, args) {
    try {
      if (!JOBS[op]) throw new Error(`Unknown job ${op}`);
      return await JOBS[op](args || {});
    } catch (err) {
      logError(err);
      return { error: err.message || String(err), code: err.code };
    } finally { await closeAll(); }
  }

  // ---------------------------------------------------------------- the owner's window

  function openDocuments() {
    return tabs.map((t, i) => ({ index: i, name: t.name, path: t.path || null, kind: FORMATS.label(t.kind), unsaved: !!t.dirty, showing: t === active }));
  }
  // An open document's text, with unsaved changes. which: its index, its name, or "active".
  async function readOpen(which, max) {
    try {
      const t = which === 'active' ? active : typeof which === 'number' ? tabs[which] : tabs.find((x) => x.name.toLowerCase() === String(which).toLowerCase()) || tabs[+which];
      if (!t) return { error: 'No open document matches that. Use docstudio_list_open_documents.', code: 'not_found' };
      return { name: t.name, path: t.path || null, unsaved: !!t.dirty, ...clip(await textOf(t), max || 200000) };
    } catch (err) { return { error: err.message, code: err.code }; }
  }

  return { run, openDocuments, readOpen, isWorker };
})();
