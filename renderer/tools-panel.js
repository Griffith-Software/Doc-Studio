'use strict';
// The sidebar's tools: Convert, OCR (scan to text, with the text recognition built into Windows), Compare,
// Merge (PDFs, pictures and documents into one PDF), Split (a PDF into several) and Compress (PDFs, pictures and
// Office files).

const Tools = (() => {
  const { pdfjs, PDFLib } = DDPdf;
  const { PDFDocument } = PDFLib;
  const IMAGE_EXTS = ['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'tif', 'tiff'];
  const ZIP_EXTS = ['docx', 'docm', 'xlsx', 'xlsm', 'pptx', 'pptm', 'odt', 'ods', 'odp', 'epub'];
  const extOf = (p) => (p.includes('.') ? p.slice(p.lastIndexOf('.') + 1).toLowerCase() : '');
  const nameOf = (p) => p.split(/[\\/]/).pop();
  const docView = () => !!active && Shell.view === 'doc';

  // ---------------------------------------------------------------- small UI helpers

  // opts.beforeClose: async check (false keeps the window open); opts.onClose: runs once it has closed.
  function modal(title, cls, opts) {
    const box = h('div', { class: `dialog ${cls || ''}` }, h('h3', { text: title }));
    const back = h('div', { class: 'backdrop' }, box);
    let closed = false;
    const close = async () => {
      if (closed || (opts && opts.beforeClose && !(await opts.beforeClose()))) return;
      closed = true;
      back.remove();
      if (opts && opts.onClose) opts.onClose();
    };
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
    return { box, close };
  }
  function busy(text) {
    const box = h('div', { class: 'dialog busy-dialog' }, h('div', { class: 'spinner' }), h('p', { text }));
    const back = h('div', { class: 'backdrop' }, box);
    document.body.append(back);
    return { set: (t) => { box.querySelector('p').textContent = t; }, close: () => back.remove() };
  }
  // ---------------------------------------------------------------- tool windows
  // Every tool opens the same kind of window (like Merge): what it does, the file(s) it works on, only the options
  // that apply to them, and one button that stays off until the tool has what it needs.

  // Runs a tool's button action; a failure is shown instead of being lost.
  const act = (label, fn) => async () => { try { await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } };
  function toolWindow(title, intro, goText, cls, opts) {
    const { box, close } = modal(title, `tool-dialog ${cls || ''}`, opts);
    if (intro) box.append(h('p', { class: 'muted tool-intro', text: intro }));
    const body = h('div', { class: 'tool-body' });
    const go = h('button', { type: 'button', text: goText, disabled: true });
    const left = h('span', { class: 'tool-left' });
    box.append(body, h('div', { class: 'dialog-actions tool-actions' }, left, h('span', { class: 'spacer' }), h('button', { type: 'button', class: 'ghost', text: 'Cancel', onclick: close }), go));
    return { box, body, go, left, close };
  }
  // A labelled option (select, input...) that can be shown or hidden.
  function field(label, input, hint) {
    const el = h('label', { class: 'tool-field' }, h('span', { text: label }), input, hint ? h('small', { class: 'muted', text: hint }) : '');
    return { el, show: (on) => { el.hidden = !on; } };
  }
  const select = (options, value) => { const el = h('select', {}, options.map(([v, t]) => h('option', { value: v, text: t }))); el.value = value; return el; };
  const kindBadge = (name) => h('span', { class: `tab-kind k-${FORMATS.kindOf(name)}`, text: (extOf(name) || '?').slice(0, 4).toUpperCase() });
  // The open document as a tool input.
  const openEntry = () => ({ name: active.name, path: active.path, tab: active, size: active.data ? active.data.byteLength : null });
  // Its bytes: the open document's current content (with unsaved PDF edits), or the file read from disk.
  async function bytesOf(entry) {
    if (entry.tab) return (await currentFileOf(entry.tab)).data;
    if (!entry.data) entry.data = (await desk.read(entry.path)).data;
    return entry.data;
  }
  // One input file: a row showing it, with Choose file… and, when the open document fits, Use the open document.
  function fileSlot({ label, fits, pickTitle, filters, onChange, preferOpen = true }) {
    let file = null;
    const row = h('div', { class: 'tool-file' });
    const openFits = () => docView() && fits(active);
    const set = async (entry) => {
      file = entry;
      paint();
      try { await onChange(entry); } catch (err) { file = null; paint(); await onChange(null); desk.error(`Couldn't use ${entry.name}`, err.message); }
    };
    const choose = async () => {
      const [p] = await desk.pick({ title: pickTitle, filters });
      if (!p) return;
      const data = (await desk.read(p)).data;
      await set({ name: nameOf(p), path: p, data, size: data.byteLength });
    };
    function paint() {
      row.replaceChildren(
        h('span', { class: 'tool-file-label', text: label }),
        file ? h('span', { class: 'tool-file-name' }, kindBadge(file.name), h('b', { text: file.name, title: file.path || 'Not saved yet' }), file.size != null ? h('small', { class: 'muted', text: formatBytes(file.size) }) : '')
          : h('span', { class: 'tool-file-name muted', text: 'No file chosen' }),
        h('span', { class: 'tool-file-buttons' },
          openFits() && !(file && file.tab === active) ? h('button', { type: 'button', class: 'ghost', text: 'Use the open document', onclick: act('use the open document', () => set(openEntry())) }) : '',
          h('button', { type: 'button', class: 'ghost', text: file ? 'Change…' : 'Choose file…', onclick: act('choose a file', choose) })));
    }
    paint();
    const ready = preferOpen && openFits() ? set(openEntry()) : Promise.resolve(onChange(null));
    return { row, ready, set, get file() { return file; } };
  }
  const setSlot = (slot, entry) => slot.set(entry);
  // Writes beside a file of the same name: "name (2).ext", "name (3).ext"... rather than over it.
  async function freeName(target) {
    if (!(await desk.exists(target))) return target;
    const m = /^(.*?)(\.[^.\\/]+)?$/.exec(target);
    for (let i = 2; ; i++) { const t = `${m[1]} (${i})${m[2] || ''}`; if (!(await desk.exists(t))) return t; }
  }
  const folderOf = (p) => (p ? p.replace(/[\\/][^\\/]*$/, '') : null);

  async function readPath(p) { return { name: nameOf(p), path: p, data: (await desk.read(p)).data }; }
  // The open document as bytes (with its unsaved edits for PDFs).
  async function currentFile() {
    const t = active;
    let data = t.data;
    if (t.kind === 'pdf' && t.pdfEditor) data = await t.pdfEditor.build();
    return { name: t.name, path: t.path, data, tab: t };
  }
  async function pickOne(title, filters) { const [f] = await desk.pick({ title, filters }); return f ? readPath(f) : null; }
  // Use the open document when it fits, otherwise ask for a file.
  async function source(title, fits, filters) {
    if (docView() && fits(active)) return currentFile();
    return pickOne(title, filters);
  }
  async function saveAs(defaultName, ext, data) {
    const target = await desk.saveDialog(defaultName, ext);
    if (!target) return null;
    await desk.write(target, data, { version: false });
    return target;
  }
  async function done(target, message, detail) {
    const r = await desk.confirm({ type: 'info', message, detail: detail || '', buttons: ['Open it', 'Show in folder', 'Close'], defaultId: 0, cancelId: 2 });
    if (r === 0) openPath(target);
    if (r === 1) desk.reveal(target);
  }
  // Open a new, unsaved tab with the given bytes. startMode 'edit' opens it ready to edit.
  async function openNew(name, data, startMode) {
    const tab = makeTab({ path: null, name, size: data.byteLength, data });
    if (startMode) tab.startMode = startMode;
    await setupTab(tab);
    markDirty(tab);
    refreshChrome();
    return tab;
  }
  const base = (f) => (f.path ? f.path.replace(/\.[^.\\/]+$/, '') : stem(f.name));

  // Pictures of PDF pages (for OCR and for "pages as images" compression).
  async function pdfImages(bytes, scale, type, quality, onPage) {
    const task = pdfjs.getDocument({ data: new Uint8Array(bytes).slice(), isEvalSupported: false });
    const doc = await task.promise;
    const out = [];
    try {
      for (let i = 1; i <= doc.numPages; i++) {
        if (onPage) onPage(i, doc.numPages);
        const page = await doc.getPage(i);
        const vp1 = page.getViewport({ scale: 1 });
        const s = Math.min(scale, 4000 / Math.max(vp1.width, vp1.height));
        const vp = page.getViewport({ scale: s });
        const c = document.createElement('canvas');
        c.width = Math.round(vp.width); c.height = Math.round(vp.height);
        const g = c.getContext('2d');
        g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
        await page.render({ canvas: c, canvasContext: g, viewport: vp, annotationMode: pdfjs.AnnotationMode.ENABLE_FORMS }).promise;
        const blob = await new Promise((r) => c.toBlob(r, type, quality));
        out.push({ data: await blob.arrayBuffer(), width: c.width, height: c.height, ptW: vp1.width, ptH: vp1.height });
        page.cleanup();
      }
    } finally { task.destroy(); }
    return out;
  }
  // Any picture as PNG or JPEG bytes (with its size), optionally scaled.
  async function imageAs(bytes, ext, type, quality, maxSide, minSide) {
    const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp', svg: 'image/svg+xml' }[ext] || 'application/octet-stream';
    const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      let w = img.naturalWidth || 1000, hh = img.naturalHeight || 1000;
      let s = 1;
      if (maxSide && Math.max(w, hh) > maxSide) s = maxSide / Math.max(w, hh);
      if (minSide && Math.max(w, hh) < minSide) s = minSide / Math.max(w, hh);
      w = Math.max(1, Math.round(w * s)); hh = Math.max(1, Math.round(hh * s));
      const c = document.createElement('canvas');
      c.width = w; c.height = hh;
      const g = c.getContext('2d');
      if (type === 'image/jpeg') { g.fillStyle = '#fff'; g.fillRect(0, 0, w, hh); }
      g.drawImage(img, 0, 0, w, hh);
      const blob = await new Promise((r) => c.toBlob(r, type, quality));
      return { data: await blob.arrayBuffer(), width: w, height: hh };
    } finally { URL.revokeObjectURL(url); }
  }

  // ---------------------------------------------------------------- Convert

  // The formats a file can become, worked out from its type (the open document's own list when it is open).
  async function targetsOf(entry) {
    if (entry.tab) return Convert.targets(entry.tab);
    let kind = FORMATS.kindOf(entry.name);
    kind = { doc: 'docx', ppt: 'pptx', odp: 'pptx' }[kind] || kind;
    return Convert.targets({ kind, ext: extOf(entry.name) });
  }
  // opts.empty: start with no files (File > Convert files…); otherwise the open document is added.
  async function convert(opts) {
    const items = [];
    let chosen = null, seq = 0;
    const w = toolWindow('Convert', 'Turn documents into other formats. The originals stay as they are; the converted copies are saved separately.', 'Convert', 'convert-tool');
    const listEl = h('div', { class: 'merge-list' });
    const grid = h('div', { class: 'convert-grid' });
    const formatNote = h('p', { class: 'muted tool-note' });
    const where = select([['beside', 'Next to each original'], ['folder', 'In a folder I choose…']], 'beside');
    const whereField = field('Save the converted files', where);
    const paintGo = () => { w.go.disabled = !items.length || !chosen; w.go.textContent = chosen ? `Convert to .${chosen}` : 'Convert'; };
    async function paint() {
      listEl.replaceChildren(...(items.length ? items.map((it, i) => h('div', { class: 'merge-row' },
        kindBadge(it.name), h('span', { class: 'merge-name', text: it.name, title: it.path || '' }),
        h('span'), h('span'),
        h('button', { class: 'ghost', text: '✕', title: 'Remove', onclick: () => { items.splice(i, 1); paint(); } })))
        : [h('p', { class: 'muted', text: 'Add the files to convert. Several files can be converted to the same format at once.' })]));
      whereField.show(items.length > 1);
      // Formats every chosen file can become.
      const mine = ++seq;
      let common = null;
      for (const it of items) {
        const list = (await targetsOf(it)).map((t) => [t.ext, t.label]);
        common = common ? common.filter(([x]) => list.some(([y]) => y === x)) : list;
      }
      if (mine !== seq) return;
      common = common || [];
      if (!common.some(([x]) => x === chosen)) chosen = null;
      grid.replaceChildren(...common.map(([ext, label]) => h('button', { type: 'button', class: `ghost convert-choice${ext === chosen ? ' on' : ''}`, onclick: () => { chosen = ext; paint(); } }, h('b', { text: `.${ext}` }), h('span', { text: label }))));
      formatNote.textContent = !items.length ? 'The formats appear once a file is added.' : common.length ? (items.length > 1 ? 'Formats all of these files can become.' : '') : 'These files have no format in common. Convert them separately.';
      paintGo();
    }
    const add = h('button', { type: 'button', class: 'ghost', text: 'Add files…', onclick: act('add files', async () => { for (const p of await desk.pick({ multi: true, title: 'Add files to convert' })) if (!items.some((x) => x.path === p)) items.push({ name: nameOf(p), path: p }); await paint(); }) });
    const addOpen = h('button', { type: 'button', class: 'ghost', text: 'Add the open document', disabled: !docView(), onclick: () => { if (!items.some((x) => x.tab === active)) items.push(openEntry()); paint(); } });
    w.left.append(add, addOpen);
    w.body.append(listEl, h('div', { class: 'tool-section', text: 'Convert to' }), grid, formatNote, whereField.el);
    w.go.onclick = act('convert', async () => {
      const ext = chosen;
      const list = items.slice();
      let folder = null;
      if (list.length > 1 && where.value === 'folder') { [folder] = await desk.pick({ folder: true, title: 'Choose where to save the converted files' }); if (!folder) return; }
      w.close();
      if (list.length === 1) return convertOne(list[0], ext);
      return convertMany(list, ext, folder);
    });
    if (!opts || !opts.empty) { if (docView()) items.push(openEntry()); }
    await paint();
  }
  // Opens a file as a tab (or uses the open one), runs fn on it, then closes it again if it was opened here.
  async function withTab(entry, fn) {
    if (entry.tab) return fn(entry.tab);
    const was = tabs.find((t) => t.path && entry.path && t.path.toLowerCase() === entry.path.toLowerCase());
    if (was) return fn(was);
    const before = tabs.length;
    await openPath(entry.path);
    const tab = active;
    if (!tab || tabs.length === before) throw new Error(`${entry.name} couldn't be opened.`);
    try { return await fn(tab); } finally { tab.dirty = false; await closeTab(tab); }
  }
  // One file: Save As for the converted copy, then Open it / Show in folder.
  async function convertOne(entry, ext) { await withTab(entry, (tab) => Convert.convertTab(tab, ext)); }
  // Several files: each saved next to its original (or in the chosen folder), then a summary.
  async function convertMany(list, ext, folder) {
    const b = busy('Converting…');
    let ok = 0, last = null;
    const failed = [];
    for (const [i, entry] of list.entries()) {
      b.set(`Converting ${entry.name} (${i + 1} of ${list.length})…`);
      try {
        await withTab(entry, async (tab) => {
          if (!(await Convert.targets(tab)).some((t) => t.ext === ext)) throw new Error(`${FORMATS.label(tab.kind)} files can't be converted to .${ext}`);
          const dir = folder || folderOf(entry.path || tab.path);
          if (!dir) throw new Error('it has not been saved yet, so choose a folder for it');
          const out = await Convert.run(tab, ext);
          const stemPath = `${dir}\\${stem(entry.name)}`;
          if (out && out.files) for (const x of out.files) await desk.write(last = await freeName(`${stemPath}${x.suffix}.${ext}`), x.data, { version: false });
          else await desk.write(last = await freeName(`${stemPath}.${ext}`), out, { version: false });
        });
        ok++;
      } catch (err) { failed.push(`${entry.name}: ${err.message}`); }
    }
    b.close();
    const r = await desk.confirm({ type: failed.length ? 'warning' : 'info', message: `Converted ${ok} of ${list.length} files to .${ext}`, detail: failed.length ? failed.join('\n') : folder ? `Saved in ${folder}` : 'Each converted file was saved next to its original.', buttons: last ? ['Show in folder', 'Close'] : ['Close'], defaultId: 0, cancelId: last ? 1 : 0 });
    if (last && r === 0) desk.reveal(last);
  }

  // ---------------------------------------------------------------- Scan

  // Scan one or more pages from a scanner connected to this PC (through Windows' own scanner support), then make a
  // picture to edit, a PDF, or text (OCR). The settings are remembered.
  async function scan() {
    const pages = [];
    let scanning = false;
    const saved = (() => { try { return JSON.parse(localStorage.getItem('docstudio.scan') || '{}'); } catch { return {}; } })();
    const w = toolWindow('Scan', 'Scan pages from a scanner connected to this PC, then make a picture to edit, a PDF, or text you can edit.', 'Scan a page', 'scan-tool');
    const dev = select([['', 'Looking for scanners…']], '');
    const again = h('button', { type: 'button', class: 'ghost', text: 'Look again' });
    const color = select([['color', 'Colour'], ['gray', 'Greyscale'], ['bw', 'Black and white (text only)']], saved.color || 'color');
    const dpi = select([['150', '150 dpi (quick)'], ['200', '200 dpi'], ['300', '300 dpi (recommended)'], ['600', '600 dpi (photos, slow)']], saved.dpi || '300');
    const paper = select([['letter', 'Letter (8.5 × 11 in)'], ['a4', 'A4'], ['glass', 'The whole glass']], saved.paper || 'letter');
    const make = select([['image', 'A picture to edit (each page)'], ['pdf', 'A PDF (all pages)'], ['docx', 'A text document (OCR)'], ['md', 'Markdown text (OCR)']], saved.make || 'image');
    const strip = h('div', { class: 'scan-pages' });
    const status = h('p', { class: 'muted tool-note' });
    const more = h('button', { type: 'button', class: 'ghost', text: 'Scan another page', hidden: true });
    w.left.append(more);
    w.body.append(
      h('div', { class: 'tool-file' }, h('span', { class: 'tool-file-label', text: 'Scanner' }), dev, h('span', { class: 'tool-file-buttons' }, again)),
      h('div', { class: 'scan-options' }, field('Colour', color).el, field('Resolution', dpi).el, field('Page size', paper).el),
      strip, field('Make', make).el, status);
    const remember = () => { try { localStorage.setItem('docstudio.scan', JSON.stringify({ id: dev.value, color: color.value, dpi: dpi.value, paper: paper.value, make: make.value })); } catch { /* not kept */ } };
    for (const el of [dev, color, dpi, paper, make]) el.addEventListener('change', () => { remember(); paint(); });
    function paint() {
      strip.replaceChildren(...pages.map((p, i) => h('div', { class: 'scan-page' }, p.thumb, h('span', { text: `Page ${i + 1}` }),
        h('button', { type: 'button', class: 'ghost', text: '✕', title: 'Remove this page', disabled: scanning, onclick: () => { pages.splice(i, 1); paint(); } }))));
      strip.hidden = !pages.length;
      more.hidden = !pages.length;
      more.disabled = scanning || !dev.value;
      const goal = { image: pages.length > 1 ? `Open ${pages.length} pictures` : 'Open the picture', pdf: 'Make the PDF', docx: 'Make the document', md: 'Make the Markdown' }[make.value];
      w.go.textContent = pages.length ? goal : 'Scan a page';
      w.go.disabled = scanning || (!pages.length && !dev.value);
      for (const el of [dev, color, dpi, paper, again]) el.disabled = scanning;
    }
    async function findScanners() {
      dev.replaceChildren(h('option', { value: '', text: 'Looking for scanners…' }));
      status.textContent = '';
      paint();
      let list = [];
      try { list = await desk.scanners(); } catch (err) { status.textContent = err.message; }
      dev.replaceChildren(...(list.length ? list.map((d) => h('option', { value: d.id, text: d.name })) : [h('option', { value: '', text: 'No scanner found' })]));
      if (list.some((d) => d.id === saved.id)) dev.value = saved.id;
      if (!list.length && !status.textContent) status.textContent = 'Turn the scanner on and check that Windows can see it (Settings > Bluetooth & devices > Printers & scanners), then press Look again.';
      paint();
    }
    async function scanPage() {
      scanning = true;
      status.textContent = `Scanning page ${pages.length + 1}… (this can take up to a minute)`;
      paint();
      try {
        const r = await desk.scanPage({ id: dev.value, color: color.value, dpi: dpi.value, paper: paper.value });
        const pic = await ImageEditor.decode(r.data, r.ext);
        const k = 72 / Math.max(pic.width, pic.height);
        const thumb = document.createElement('canvas');
        thumb.width = Math.max(1, Math.round(pic.width * k)); thumb.height = Math.max(1, Math.round(pic.height * k));
        thumb.getContext('2d').drawImage(pic, 0, 0, thumb.width, thumb.height);
        pages.push({ pic, thumb, dpi: +dpi.value, bw: color.value === 'bw' });
        status.textContent = `Page ${pages.length} scanned. Put the next page on the glass and press Scan another page, or finish below.`;
      } catch (err) { status.textContent = ''; desk.error("Couldn't scan", err.message); }
      scanning = false;
      paint();
    }
    again.addEventListener('click', act('look for scanners', findScanners));
    more.addEventListener('click', act('scan', scanPage));
    w.go.onclick = act('scan', async () => {
      if (!pages.length) { await scanPage(); return; }
      w.close();
      const when = new Date(), pad = (n) => String(n).padStart(2, '0');
      const name = `Scan ${when.getFullYear()}-${pad(when.getMonth() + 1)}-${pad(when.getDate())} ${pad(when.getHours())}.${pad(when.getMinutes())}`;
      const b = busy('Preparing…');
      try {
        if (make.value === 'image') {
          for (const [i, p] of pages.entries()) { b.set(`Opening page ${i + 1}…`); await openNew(`${name}${pages.length > 1 ? ` (${i + 1})` : ''}.png`, await ImageEditor.encode(p.pic, 'png'), 'edit'); }
        } else if (make.value === 'pdf') {
          const doc = await PDFDocument.create();
          doc.setCreator('Doc Studio'); doc.setProducer('Doc Studio');
          for (const [i, p] of pages.entries()) {
            b.set(`Adding page ${i + 1} of ${pages.length}…`);
            // Photos and greyscale as JPEG (much smaller); black-and-white text as PNG (sharp).
            const img = p.bw ? await doc.embedPng(await ImageEditor.encode(p.pic, 'png')) : await doc.embedJpg(await ImageEditor.encode(p.pic, 'jpg', 0.85));
            const wPt = p.pic.width / p.dpi * 72, hPt = p.pic.height / p.dpi * 72;
            doc.addPage([wPt, hPt]).drawImage(img, { x: 0, y: 0, width: wPt, height: hPt });
          }
          await openNew(`${name}.pdf`, (await doc.save()).buffer);
        } else {
          const images = [];
          for (const p of pages) images.push({ data: await ImageEditor.encode(p.pic, 'png'), ext: 'png' });
          b.set(images.length > 1 ? `Recognising the text on ${images.length} pages…` : 'Recognising the text…');
          const texts = await desk.ocr(images);
          b.close();
          await textOut(texts, `${name} (text)`, make.value);
          return;
        }
      } finally { b.close(); }
      flash('Scanned. Check it, then save.');
    });
    paint();
    await findScanners();
  }

  // ---------------------------------------------------------------- Phone

  // Send photos and files between a phone and this PC over the home network (phone.js serves the phone's page).
  // The phone scans the QR code; files it sends are listed here to open or save; files added under Send to phone
  // can be downloaded on the phone. Everything stops when this window closes.
  async function phone() {
    const got = []; // { id, name, size, done, opened, saved, row }
    const offered = [];
    let unsubscribe = null;
    const pending = () => got.filter((x) => x.done && !x.opened && !x.saved);
    const w = toolWindow('Phone', 'Send photos and files between your phone and this PC over your home Wi-Fi.', 'Open all', 'phone-tool', {
      beforeClose: async () => {
        const n = pending().length;
        if (!n) return true;
        return (await desk.confirm({ type: 'warning', message: `${n} file${n === 1 ? '' : 's'} from your phone ${n === 1 ? 'has' : 'have'} not been opened or saved.`, detail: 'They are deleted when this window closes.', buttons: ['Close anyway', 'Keep the window open'], defaultId: 1, cancelId: 1 })) === 0;
      },
      onClose: () => { if (unsubscribe) unsubscribe(); desk.phoneStop(); },
    });
    const qr = h('canvas', { class: 'phone-qr', width: 232, height: 232 });
    const addr = h('div', { class: 'phone-addr' });
    const status = h('p', { class: 'phone-status', text: 'Starting…' });
    const steps = h('ol', { class: 'phone-steps' },
      h('li', { text: "Open your phone's camera and point it at the code." }),
      h('li', { text: 'Tap the link that appears. A Doc Studio page opens in the browser.' }),
      h('li', { text: 'Take a photo or choose files. They appear below.' }));
    const recvList = h('div', { class: 'merge-list phone-list' });
    const sendList = h('div', { class: 'merge-list phone-list' });
    const saveAll = h('button', { type: 'button', class: 'ghost', text: 'Save all to a folder…', disabled: true });
    const addOpen = h('button', { type: 'button', class: 'ghost', text: 'Add the open document', disabled: !docView() });
    const addFiles = h('button', { type: 'button', class: 'ghost', text: 'Add files…' });
    w.left.append(saveAll);
    w.body.append(
      h('div', { class: 'phone-top' }, h('div', { class: 'phone-qr-box' }, qr), h('div', { class: 'phone-help' }, steps, addr, status,
        // The link is a small web page on this PC, so Windows may ask whether Doc Studio may use the network.
        h('p', { class: 'muted tool-note phone-firewall', text: 'If Windows asks whether to allow Doc Studio on networks, choose Private networks only (your home network), not Public.' }))),
      h('div', { class: 'tool-section', text: 'Received from your phone' }), recvList,
      h('div', { class: 'tool-section phone-send-head' }, h('span', { text: 'Send to your phone' }), h('span', { class: 'phone-send-buttons' }, addOpen, addFiles)), sendList);
    const kb = (n) => (n ? formatBytes(n) : '');
    function paint() {
      recvList.replaceChildren(...(got.length ? got.map((x) => {
        const state = x.failed ? x.failed : !x.done ? `Receiving… ${x.size ? Math.round((x.got || 0) / x.size * 100) + '%' : ''}` : x.opened ? 'Opened' : x.saved ? 'Saved' : kb(x.size);
        return h('div', { class: 'merge-row' }, kindBadge(x.name), h('span', { class: 'merge-name', text: x.name }), h('small', { class: `muted${x.failed ? ' tool-problem' : ''}`, text: state }),
          h('button', { type: 'button', class: 'ghost', text: 'Open', disabled: !x.done || !!x.failed, onclick: act('open the file', () => openOne(x)) }),
          h('button', { type: 'button', class: 'ghost', text: '✕', title: 'Remove', disabled: !x.done && !x.failed, onclick: () => { got.splice(got.indexOf(x), 1); if (x.done) desk.phoneForget(x.id); paint(); } }));
      }) : [h('p', { class: 'muted', text: 'Nothing yet. Photos and files you send from the phone appear here.' })]));
      sendList.replaceChildren(...(offered.length ? offered.map((x) => h('div', { class: 'merge-row' }, kindBadge(x.name), h('span', { class: 'merge-name', text: x.name }), h('small', { class: 'muted', text: x.taken ? 'Downloaded' : kb(x.size) }), h('span'),
        h('button', { type: 'button', class: 'ghost', text: '✕', title: 'Stop offering it', onclick: () => { offered.splice(offered.indexOf(x), 1); desk.phoneUnoffer(x.id); paint(); } })))
        : [h('p', { class: 'muted', text: 'Add a document or files here and they show up on the phone page to download.' })]));
      const ready = got.filter((x) => x.done && !x.failed);
      w.go.disabled = !ready.some((x) => !x.opened);
      saveAll.disabled = !ready.length;
      addOpen.disabled = !docView();
    }
    async function openOne(x) {
      const data = await desk.phoneRead(x.id);
      await openNew(x.name, data);
      x.opened = true;
      paint();
    }
    // The QR code: black on white with a quiet border, whatever the theme.
    function drawQr(text) {
      const code = qrcode(0, 'M');
      code.addData(text);
      code.make();
      const n = code.getModuleCount(), q = 4, cell = Math.floor(qr.width / (n + q * 2)), off = Math.floor((qr.width - cell * (n + q * 2)) / 2);
      const g = qr.getContext('2d');
      g.fillStyle = '#fff'; g.fillRect(0, 0, qr.width, qr.height);
      g.fillStyle = '#000';
      for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) if (code.isDark(r, c)) g.fillRect(off + (c + q) * cell, off + (r + q) * cell, cell, cell);
    }
    unsubscribe = desk.onPhone((ev, d) => {
      if (ev === 'connected') status.textContent = `Phone connected${/iPhone|iPad/.test(d.agent) ? ' (iPhone)' : /Android/.test(d.agent) ? ' (Android)' : ''}. Keep the page open on the phone.`;
      else if (ev === 'incoming') got.unshift({ id: d.id, name: d.name, size: d.size, got: 0 });
      else if (ev === 'progress') { const x = got.find((y) => y.id === d.id); if (x) x.got = d.got; }
      else if (ev === 'received') { const x = got.find((y) => y.id === d.id); if (x) { x.done = true; x.size = d.size; } else got.unshift({ id: d.id, name: d.name, size: d.size, done: true }); flash(`Received ${d.name} from your phone`); }
      else if (ev === 'failed') { const x = got.find((y) => y.id === d.id); if (x) x.failed = d.message; }
      else if (ev === 'downloaded') { const x = offered.find((y) => y.id === d.id); if (x) x.taken = true; }
      else if (ev === 'stopped') { status.textContent = 'The link stopped after 15 minutes without use. Close this window and open Phone again.'; status.classList.add('tool-problem'); }
      paint();
    });
    saveAll.addEventListener('click', act('save the files', async () => {
      const [folder] = await desk.pick({ folder: true, title: 'Choose where to save the files from your phone' });
      if (!folder) return;
      const list = got.filter((x) => x.done && !x.failed);
      const saved = await desk.phoneSaveTo(list.map((x) => x.id), folder);
      list.forEach((x) => { x.saved = true; });
      paint();
      if (saved.length) desk.reveal(saved[0]);
    }));
    // The open document goes as it is now (with unsaved edits); a saved, unchanged file goes straight from disk.
    addOpen.addEventListener('click', act('add the document', async () => {
      const t = active;
      let item;
      if (t.path && !t.dirty) item = { path: t.path, name: t.name };
      else if (FORMATS.isEditable(t.kind) && hasEditor(t)) item = { name: `${stem(t.name)}.${nativeExt(t)}`, data: await nativeBytes(t, true) };
      else if (t.data) item = { name: t.name, data: t.data };
      else throw new Error('This document has nothing to send yet.');
      offered.push(...(await desk.phoneOffer([item])));
      paint();
    }));
    addFiles.addEventListener('click', act('add files', async () => {
      const paths = await desk.pick({ multi: true, title: 'Choose files to send to your phone' });
      if (!paths.length) return;
      offered.push(...(await desk.phoneOffer(paths.map((p) => ({ path: p })))));
      paint();
    }));
    w.go.onclick = act('open the files', async () => {
      for (const x of got.filter((y) => y.done && !y.failed && !y.opened)) await openOne(x);
      w.close();
    });
    paint();
    let info;
    try { info = await desk.phoneStart(); } catch (err) { status.textContent = `The phone link couldn't start: ${err.message}`; status.classList.add('tool-problem'); return; }
    if (!info.addresses.length) { status.textContent = "This PC isn't connected to a network. Connect it to the same Wi-Fi as your phone, then open Phone again."; status.classList.add('tool-problem'); return; }
    const urlOf = (a) => `http://${a}:${info.port}/?k=${info.token}`;
    // Can't scan? The full link (with its code) can be copied and sent to the phone some other way.
    const show = (a) => { drawQr(urlOf(a)); qr.dataset.url = urlOf(a); addr.replaceChildren(h('span', { class: 'muted', text: "Can't scan the code? " }), h('button', { type: 'button', class: 'link', text: 'Copy the link', onclick: async () => { await navigator.clipboard.writeText(urlOf(a)); flash('Link copied. Send it to your phone (for example in a message to yourself) and open it there.'); } })); };
    if (info.addresses.length > 1) {
      const pick = select(info.addresses.map((x) => [x.address, `${x.address} (${x.name})`]), info.addresses[0].address);
      pick.addEventListener('change', () => show(pick.value));
      addr.before(field('Network', pick, 'If the phone cannot open the page, try the other network.').el);
    }
    show(info.addresses[0].address);
    status.textContent = 'Waiting for your phone… (it must be on the same Wi-Fi as this PC)';
  }

  // ---------------------------------------------------------------- OCR

  async function ocr() {
    const w = toolWindow('OCR (Scan to Text)', 'Reads the text in a scan, photo or PDF, using the text recognition built into Windows.', 'Recognise text');
    const out = select([['docx', 'A new document (.docx)'], ['md', 'A new Markdown file'], ['txt', 'A new text file'], ['clip', 'The clipboard']], 'docx');
    const src = fileSlot({ label: 'Read from', fits: (t) => t.kind === 'pdf' || t.kind === 'image', pickTitle: 'Choose a scan, photo or PDF', filters: [{ name: 'Pictures and PDFs', extensions: ['pdf', ...IMAGE_EXTS] }], onChange: (f) => { w.go.disabled = !f; } });
    w.body.append(src.row, field('Put the text in', out).el);
    w.go.onclick = act('recognise the text', async () => { const f = src.file; w.close(); await recognise({ name: f.name, data: await bytesOf(f) }, out.value); });
    await src.ready;
  }
  async function recognise(f, where) {
    const ext = extOf(f.name);
    const b = busy('Preparing the pages…');
    let texts;
    try {
      let images;
      if (ext === 'pdf') images = (await pdfImages(f.data, 3, 'image/png', 1, (i, n) => b.set(`Preparing page ${i} of ${n}…`))).map((x) => ({ data: x.data, ext: 'png' }));
      else images = [{ data: (await imageAs(f.data, ext, 'image/png', 1, 4000, 1600)).data, ext: 'png' }];
      b.set(images.length > 1 ? `Recognising the text on ${images.length} pages…` : 'Recognising the text…');
      texts = await desk.ocr(images);
    } catch (err) { b.close(); throw new Error(`${err.message}${/language/i.test(err.message) ? '' : '\nWindows text recognition needs a language with OCR support (Settings > Time & language > Language).'}`); }
    b.close();
    await textOut(texts, `${stem(f.name)} (text)`, where);
  }
  // Recognised text (one string per page) into a new document, a text file or the clipboard.
  async function textOut(texts, title, where) {
    const pages = texts.map((t) => t.split('\n').map((l) => l.trim()));
    if (!pages.some((p) => p.some(Boolean))) { desk.error('No text was found', 'The picture may be too small, blurred or rotated.'); return; }
    if (where === 'clip') { await navigator.clipboard.writeText(pages.map((p) => p.join('\n')).join('\n\n')); flash('Recognised text copied to the clipboard'); return; }
    if (where === 'txt' || where === 'md') { await openNew(`${title}.${where}`, new TextEncoder().encode(pages.map((p) => p.join('\n')).join(where === 'md' ? '\n\n---\n\n' : '\n\n')).buffer); return; }
    const esc = (x) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const html = pages.map((p, i) => `${i ? '<p></p>' : ''}${p.filter(Boolean).map((l) => `<p>${esc(l)}</p>`).join('')}`).join('');
    await openNew(`${title}.docx`, await desk.docxFromHtml(html, title));
    flash('Recognised text opened as a new document - check it, then save');
  }

  // ---------------------------------------------------------------- Compare

  // first: a tab to start from (an editor's Tools > Compare documents…).
  async function compare(first) {
    const fits = (t) => !['image', 'audio', 'video', 'unknown'].includes(t.kind);
    const w = toolWindow('Compare documents', 'Shows what changed between two versions of a document: removed text is struck through, added text is underlined.', 'Compare');
    let a = null, b = null;
    const paintGo = () => { w.go.disabled = !(a && b && a.file && b.file); };
    a = fileSlot({ label: 'Original', fits, pickTitle: 'Choose the original document', onChange: paintGo, preferOpen: !first });
    b = fileSlot({ label: 'Changed version', fits: (t) => fits(t) && !!t.path && !(a && a.file && a.file.tab === t), pickTitle: 'Choose the changed version', onChange: paintGo, preferOpen: false });
    w.body.append(a.row, b.row);
    w.go.onclick = act('compare the documents', async () => {
      const x = a.file, y = b.file;
      w.close();
      let tab = x.tab;
      if (!tab) { await openPath(x.path); tab = active; }
      if (tab) await DocTools.compareDocuments(tab, y.path);
    });
    await a.ready;
    if (first) await setSlot(a, { name: first.name, path: first.path, tab: first });
    await b.ready;
    paintGo();
  }

  // ---------------------------------------------------------------- Merge

  // Any file as PDF bytes: PDFs as they are, pictures as pages, other documents through Doc Studio's converters.
  async function asPdf(entry) {
    const ext = extOf(entry.name);
    if (entry.tab && entry.tab.kind !== 'pdf') return Convert.run(entry.tab, 'pdf');
    if (ext === 'pdf') return entry.data;
    if (IMAGE_EXTS.includes(ext)) {
      const doc = await PDFDocument.create();
      const jpg = ext === 'jpg' || ext === 'jpeg';
      const img = jpg ? await doc.embedJpg(entry.data) : await doc.embedPng((await imageAs(entry.data, ext, 'image/png', 1)).data);
      const page = doc.addPage([img.width * 0.75, img.height * 0.75]);
      page.drawImage(img, { x: 0, y: 0, width: img.width * 0.75, height: img.height * 0.75 });
      return doc.save();
    }
    // Open it, convert it, close it again.
    const before = tabs.length;
    await openPath(entry.path);
    const tab = active;
    if (!tab || tabs.length === before) throw new Error(`${entry.name} couldn't be opened.`);
    try { return await Convert.run(tab, 'pdf'); } finally { tab.dirty = false; await closeTab(tab); }
  }
  async function merge() {
    const items = [];
    const { box, close } = modal('Merge into one PDF', 'tool-dialog merge-dialog');
    const listEl = h('div', { class: 'merge-list' });
    const paint = () => {
      listEl.replaceChildren(...(items.length ? items.map((it, i) => h('div', { class: 'merge-row' },
        h('span', { class: `tab-kind k-${FORMATS.kindOf(it.name)}`, text: (extOf(it.name) || '?').slice(0, 4).toUpperCase() }),
        h('span', { class: 'merge-name', text: it.name, title: it.path || '' }),
        h('button', { class: 'ghost', text: '↑', title: 'Move up', disabled: i === 0, onclick: () => { [items[i - 1], items[i]] = [items[i], items[i - 1]]; paint(); } }),
        h('button', { class: 'ghost', text: '↓', title: 'Move down', disabled: i === items.length - 1, onclick: () => { [items[i + 1], items[i]] = [items[i], items[i + 1]]; paint(); } }),
        h('button', { class: 'ghost', text: '✕', title: 'Remove', onclick: () => { items.splice(i, 1); paint(); } })))
        : [h('p', { class: 'muted', text: 'Add PDFs, pictures, documents, spreadsheets, presentations or any other file Doc Studio opens. They are joined in this order.' })]));
      go.disabled = items.length < 1;
    };
    const add = h('button', { class: 'ghost', text: 'Add files…', onclick: act('add files', async () => { for (const p of await desk.pick({ multi: true, title: 'Add files to merge' })) items.push({ name: nameOf(p), path: p }); paint(); }) });
    const addOpen = h('button', { class: 'ghost', text: 'Add the open document', disabled: !docView(), onclick: () => { items.push({ name: active.name, path: active.path, tab: active }); paint(); } });
    const go = h('button', { text: 'Merge', onclick: act('merge the files', async () => {
      close();
      const b = busy('Merging…');
      try {
        const { data, pages } = await mergeFiles(items, (i) => b.set(`Adding ${items[i].name} (${i + 1} of ${items.length})…`));
        b.close();
        const target = await saveAs('Merged.pdf', 'pdf', data);
        if (target) await done(target, `Merged ${items.length} file(s) into ${nameOf(target)}`, `${pages} page(s)`);
      } catch (err) { b.close(); throw err; }
    }) });
    box.append(listEl, h('div', { class: 'dialog-actions', style: 'justify-content:space-between' }, h('span', {}, add, ' ', addOpen), h('span', {}, h('button', { class: 'ghost', text: 'Cancel', onclick: close }), ' ', go)));
    if (docView()) items.push({ name: active.name, path: active.path, tab: active });
    paint();
  }
  // items: [{ name, path?, data?, tab? }] -> one PDF
  async function mergeFiles(items, onItem) {
    const out = await PDFDocument.create();
    for (const [i, it] of items.entries()) {
      if (onItem) onItem(i);
      if (!it.data && !it.tab && it.path && (extOf(it.name) === 'pdf' || IMAGE_EXTS.includes(extOf(it.name)))) it.data = (await desk.read(it.path)).data;
      if (it.tab && it.tab.kind === 'pdf') it.data = (await currentFileOf(it.tab)).data;
      const bytes = await asPdf(it);
      const src = await PDFDocument.load(bytes, { ignoreEncryption: true });
      for (const p of await out.copyPages(src, src.getPageIndices())) out.addPage(p);
    }
    return { data: await out.save(), pages: out.getPageCount() };
  }
  // groups of page indexes -> one PDF each
  async function splitPdf(bytes, groups) {
    const src = await PDFDocument.load(bytes, { ignoreEncryption: true });
    const out = [];
    for (const g of groups) { const d = await PDFDocument.create(); for (const p of await d.copyPages(src, g)) d.addPage(p); out.push(await d.save()); }
    return out;
  }
  async function currentFileOf(tab) { return { name: tab.name, data: tab.pdfEditor ? await tab.pdfEditor.build() : tab.data }; }

  // ---------------------------------------------------------------- Split

  // "1-3, 5, 8-10" -> [[0,1,2],[4],[7,8,9]]
  function parseRanges(text, n) {
    const out = [];
    for (const part of text.split(/[,;]+/).map((x) => x.trim()).filter(Boolean)) {
      const m = /^(\d+)\s*(?:-\s*(\d+|end))?$/i.exec(part);
      if (!m) throw new Error(`"${part}" isn't a page or a range like 2-5.`);
      const a = +m[1], b = m[2] ? (/end/i.test(m[2]) ? n : +m[2]) : a;
      if (a < 1 || b > n || b < a) throw new Error(`${part} is outside pages 1-${n}.`);
      out.push(Array.from({ length: b - a + 1 }, (_, i) => a - 1 + i));
    }
    return out;
  }
  async function split() {
    let src = null, n = 0, folder = null;
    const w = toolWindow('Split a PDF', 'Saves the pages of a PDF as several smaller PDFs. The original stays as it is.', 'Split');
    const how = select([['each', 'Every page into its own file'], ['every', 'Every few pages'], ['ranges', 'By page ranges'], ['extract', 'Extract some pages into one file']], 'each');
    const size = h('input', { type: 'number', min: '1', value: '2' });
    const ranges = h('input', { type: 'text', placeholder: '1-3, 5, 8-end' });
    const sizeF = field('Pages in each file', size), rangesF = field('Pages', ranges, 'Pages and ranges, separated by commas, for example 1-3, 5, 8-end.');
    const info = h('p', { class: 'muted tool-note' });
    const folderName = h('span', { class: 'tool-folder-name' });
    const folderRow = h('div', { class: 'tool-file' }, h('span', { class: 'tool-file-label', text: 'Save in' }), folderName,
      h('span', { class: 'tool-file-buttons' }, h('button', { type: 'button', class: 'ghost', text: 'Change…', onclick: act('choose a folder', async () => { const [d] = await desk.pick({ folder: true, title: 'Choose where to save the PDFs' }); if (d) { folder = d; paint(); } }) })));
    // The page groups the current settings make (an error message when the ranges can't be read).
    const groups = () => {
      if (!n) return [];
      if (how.value === 'each') return Array.from({ length: n }, (_, i) => [i]);
      if (how.value === 'every') { const k = Math.max(1, +size.value || 1); const g = []; for (let i = 0; i < n; i += k) g.push(Array.from({ length: Math.min(k, n - i) }, (_, j) => i + j)); return g; }
      const r = parseRanges(ranges.value, n);
      return how.value === 'ranges' ? r : (r.length ? [r.flat()] : []);
    };
    function paint() {
      sizeF.show(how.value === 'every');
      rangesF.show(how.value === 'ranges' || how.value === 'extract');
      folderName.textContent = folder || 'Choose a folder';
      folderName.classList.toggle('muted', !folder);
      let g = [], problem = '';
      try { g = groups(); } catch (err) { problem = err.message; }
      info.textContent = !n ? '' : problem || (g.length ? `${n} pages → ${g.length} PDF file${g.length === 1 ? '' : 's'}` : 'Enter the pages to keep.');
      info.classList.toggle('tool-problem', !!problem);
      w.go.disabled = !src || !folder || !!problem || !g.length;
    }
    const slot = fileSlot({ label: 'PDF', fits: (t) => t.kind === 'pdf', pickTitle: 'Choose a PDF to split', filters: [{ name: 'PDF', extensions: ['pdf'] }], onChange: async (f) => {
      src = null; n = 0;
      if (f) {
        src = await PDFDocument.load(await bytesOf(f), { ignoreEncryption: true });
        n = src.getPageCount();
        if (!ranges.value) ranges.value = n > 1 ? `1-${Math.ceil(n / 2)}, ${Math.ceil(n / 2) + 1}-end` : '1';
        folder = folder || folderOf(f.path || (f.tab && f.tab.path));
      }
      paint();
    } });
    for (const el of [how, size, ranges]) el.addEventListener('input', paint);
    w.body.append(slot.row, field('Split', how).el, sizeF.el, rangesF.el, folderRow, info);
    w.go.onclick = act('split the PDF', async () => {
      const g = groups(), name = slot.file.name;
      w.close();
      const b = busy('Splitting…');
      const pad = String(g.length).length;
      let first = null;
      try {
        for (const [i, pages] of g.entries()) {
          b.set(`Saving file ${i + 1} of ${g.length}…`);
          const out = await PDFDocument.create();
          for (const p of await out.copyPages(src, pages)) out.addPage(p);
          const label = pages.length === 1 ? `page ${pages[0] + 1}` : `pages ${pages[0] + 1}-${pages[pages.length - 1] + 1}`;
          const target = await freeName(g.length === 1 ? `${folder}\\${stem(name)} (${label}).pdf` : `${folder}\\${stem(name)} - ${String(i + 1).padStart(pad, '0')} (${label}).pdf`);
          await desk.write(target, await out.save(), { version: false });
          first = first || target;
        }
      } finally { b.close(); }
      if (g.length === 1) await done(first, `Saved ${g[0].length} page(s) as ${nameOf(first)}`);
      else { flash(`Split into ${g.length} files`); desk.reveal(first); }
    });
    await slot.ready;
    paint();
  }

  // ---------------------------------------------------------------- Compress

  async function compressPdf(data, level, b) {
    if (level === 'lossless') {
      const doc = await PDFDocument.load(data, { ignoreEncryption: true, updateMetadata: false });
      return doc.save({ useObjectStreams: true });
    }
    const scale = level === 'small' ? 96 / 72 : 150 / 72;
    const quality = level === 'small' ? 0.55 : 0.72;
    const pages = await pdfImages(data, scale, 'image/jpeg', quality, (i, n) => b.set(`Compressing page ${i} of ${n}…`));
    const doc = await PDFDocument.create();
    for (const p of pages) {
      const img = await doc.embedJpg(p.data);
      const page = doc.addPage([p.ptW, p.ptH]);
      page.drawImage(img, { x: 0, y: 0, width: p.ptW, height: p.ptH });
    }
    return doc.save({ useObjectStreams: true });
  }
  // Office files: pictures inside are made smaller and the package is packed tighter.
  async function compressZip(data, quality, maxSide, b) {
    const zip = await JSZip.loadAsync(data);
    const media = Object.keys(zip.files).filter((n) => /\.(jpe?g|png)$/i.test(n) && !zip.files[n].dir);
    for (const [i, n] of media.entries()) {
      b.set(`Compressing picture ${i + 1} of ${media.length}…`);
      const bytes = await zip.file(n).async('arraybuffer');
      if (bytes.byteLength < 60 * 1024) continue;
      const ext = extOf(n);
      try {
        const out = await imageAs(bytes, ext, ext === 'png' ? 'image/png' : 'image/jpeg', quality, maxSide);
        if (out.data.byteLength < bytes.byteLength) zip.file(n, out.data);
      } catch { /* leave this picture as it is */ }
    }
    const mimetype = zip.file('mimetype');
    const opts = { type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } };
    if (mimetype) opts.mimeType = await mimetype.async('string');
    return zip.generateAsync(opts);
  }
  async function compress() {
    let fields = [], kind = '';
    const w = toolWindow('Compress', 'Makes a PDF, picture or Office file smaller. The original stays as it is; the smaller copy is saved separately.', 'Compress');
    const opts = h('div', { class: 'tool-options' });
    const fits = (t) => ['pdf', 'image'].includes(t.kind) || ZIP_EXTS.includes(extOf(t.name));
    const slot = fileSlot({ label: 'File', fits, pickTitle: 'Choose a file to make smaller', filters: [{ name: 'PDFs, pictures and Office files', extensions: ['pdf', ...IMAGE_EXTS, ...ZIP_EXTS] }], onChange: (f) => {
      const ext = f ? extOf(f.name) : '';
      kind = ext === 'pdf' ? 'pdf' : IMAGE_EXTS.includes(ext) ? 'image' : ZIP_EXTS.includes(ext) ? 'zip' : '';
      if (kind === 'pdf') fields = [['level', 'How much smaller', select([['lossless', 'Tidy up only (no quality loss; text stays selectable)'], ['balanced', 'Smaller (pages become 150 dpi pictures)'], ['small', 'Smallest (pages become 96 dpi pictures)']], 'lossless')]];
      else if (kind === 'image') fields = [['quality', 'Quality (10-100)', h('input', { type: 'number', min: '10', max: '100', value: '75' })], ['max', 'Longest side in pixels (blank keeps the size)', h('input', { type: 'number', min: '1', value: '2000' })], ['format', 'Save as', select([['jpg', 'JPEG'], ['webp', 'WebP'], ['png', 'PNG (lossless)']], ext === 'png' ? 'png' : 'jpg')]];
      else if (kind === 'zip') fields = [['quality', 'Picture quality (10-100)', h('input', { type: 'number', min: '10', max: '100', value: '75' })], ['max', 'Largest picture side in pixels', h('input', { type: 'number', min: '1', value: '1600' })]];
      else fields = [];
      opts.replaceChildren(...fields.map(([, label, input]) => field(label, input).el));
      w.go.disabled = !kind;
    } });
    w.body.append(slot.row, opts);
    w.go.onclick = act('compress the file', async () => {
      const f = slot.file, v = Object.fromEntries(fields.map(([k, , input]) => [k, input.value]));
      w.close();
      const data = await bytesOf(f), ext = extOf(f.name);
      const b = busy('Compressing…');
      let out, outExt = ext;
      try {
        if (kind === 'pdf') out = await compressPdf(data, v.level, b);
        else if (kind === 'image') { outExt = v.format; out = (await imageAs(data, ext, { jpg: 'image/jpeg', webp: 'image/webp', png: 'image/png' }[v.format], Math.max(0.1, Math.min(1, (+v.quality || 75) / 100)), +v.max || 0)).data; }
        else out = await compressZip(data, Math.max(0.1, Math.min(1, (+v.quality || 75) / 100)), +v.max || 1600, b);
      } finally { b.close(); }
      const before = data.byteLength, after = out.byteLength;
      if (after >= before && outExt === ext) { await desk.confirm({ type: 'info', buttons: ['OK'], message: 'This file is already as small as these settings can make it.', detail: `${formatBytes(before)} → ${formatBytes(after)}` }); return; }
      const target = await saveAs(`${base(f)} (compressed).${outExt}`, outExt, out);
      if (target) await done(target, `Saved ${nameOf(target)}`, `${formatBytes(before)} → ${formatBytes(after)} (${Math.round((1 - after / before) * 100)}% smaller)`);
    });
    await slot.ready;
  }

  return { scan, phone, convert, ocr, compare, merge, split, compress, parseRanges, compressPdf, compressZip, pdfImages, mergeFiles, splitPdf, imageAs };
})();
