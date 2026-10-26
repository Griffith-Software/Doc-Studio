'use strict';
// More Google Docs features for the Word editor, and shared pieces used by Sheets and Slides:
//   Suggesting mode and the review panel (accept/reject), equations, picture options (resize handles, size,
//   alt text, crop), dropdown chips, section breaks, list styles and restart numbering, paragraph borders and
//   shading, update style to match, the ruler (indents), pageless view, citations, dictionary, and two
//   inserters shared by all editors: a drawing canvas and a chart maker (both give a PNG).

const DocsExtra = (() => {
  const form = (...a) => DocTools.form(...a);
  const run = (label, fn) => async () => { try { await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } };

  // ---------------------------------------------------------------- Suggesting mode

  function setSuggesting(tab, on) {
    const ed = tab.editor;
    if (!ed) return;
    ed.suggesting = on;
    tab.body.classList.toggle('suggesting', on);
    if (tab.suggestSel) tab.suggestSel.value = on ? 'suggest' : 'edit';
    flash(on ? 'Suggesting: your edits are tracked as suggestions' : 'Editing: changes go straight into the document');
  }
  async function reviewPanel(tab) {
    const ed = tab.editor;
    const list = await ed.trackedChanges();
    const box = h('div', { class: 'dialog versions-dialog' }, h('h3', { text: 'Review suggested edits' }));
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    const names = { ins: 'Added', del: 'Deleted', moveFrom: 'Moved from', moveTo: 'Moved to', rPrChange: 'Formatting', pPrChange: 'Paragraph formatting' };
    if (!list.length) box.append(h('p', { class: 'muted', text: 'There are no suggested edits. Switch the toolbar from Editing to Suggesting to make some.' }));
    else box.append(h('div', { class: 'dialog-actions', style: 'justify-content:flex-start' },
      h('button', { text: 'Accept all', onclick: run('accept all', async () => { close(); await ed.reviewChanges(null, true); flash('All suggestions accepted'); }) }),
      h('button', { class: 'ghost', text: 'Reject all', onclick: run('reject all', async () => { close(); await ed.reviewChanges(null, false); flash('All suggestions rejected'); }) })));
    const rows = h('div', { class: 'version-list' });
    for (const c of list) {
      rows.append(h('div', { class: 'change-row' },
        h('span', { class: `change-kind k-${c.type}`, text: names[c.type] || c.type }),
        h('span', { class: 'change-text', text: c.text || '(formatting)' }),
        h('span', { class: 'muted', text: [c.author, c.date ? new Date(c.date).toLocaleString() : ''].filter(Boolean).join(' · ') }),
        h('button', { class: 'ghost', text: '✓', title: 'Accept', onclick: run('accept', async () => { close(); await ed.reviewChanges(c.id, true); await reviewPanel(tab); }) }),
        h('button', { class: 'ghost', text: '✕', title: 'Reject', onclick: run('reject', async () => { close(); await ed.reviewChanges(c.id, false); await reviewPanel(tab); }) })));
    }
    box.append(rows, h('div', { class: 'dialog-actions' }, h('button', { class: 'ghost', text: 'Close', onclick: close })));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
  }

  // ---------------------------------------------------------------- equations

  async function equation(ed) {
    const v = await form('Insert equation', [{ name: 'src', label: 'Equation: a/b, x^2, x_i, \\sqrt{x}, \\frac{a}{b}, \\sum_{i=1}^{n} x_i, \\int_0^1, \\alpha, \\pi, \\pm, \\le …', value: 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}' }], 'Insert');
    if (!v || !v.src.trim()) return;
    DocTools.restore();
    await ed.insertEquation(v.src.trim());
  }

  // ---------------------------------------------------------------- pictures: select, drag to resize, options

  function wirePictures(tab) {
    const ed = tab.editor;
    if (!ed || ed._picWired === ed.stage) return;
    ed._picWired = ed.stage;
    ed.stage.addEventListener('mousedown', (e) => {
      const img = e.target.closest && e.target.closest('img');
      document.querySelectorAll('.dd-img-sel').forEach((x) => x.classList.remove('dd-img-sel'));
      document.querySelectorAll('.dd-img-handle').forEach((x) => x.remove());
      if (!img || !ed.stage.contains(img) || !img.closest('[data-dd-r]')) { ed.selectedImage = null; return; }
      ed.selectedImage = img;
      img.classList.add('dd-img-sel');
      const handle = h('b', { class: 'dd-img-handle', title: 'Drag to resize' });
      const place = () => { const r = img.getBoundingClientRect(); handle.style.left = `${r.right - 5}px`; handle.style.top = `${r.bottom - 5}px`; };
      place();
      document.body.append(handle);
      handle.addEventListener('mousedown', (ev) => {
        ev.preventDefault(); ev.stopPropagation();
        const info = ed.imageInfo(img);
        if (!info) return;
        const sx = ev.clientX, w0 = img.getBoundingClientRect().width, h0 = img.getBoundingClientRect().height;
        const move = (mv) => { const w = Math.max(16, w0 + mv.clientX - sx); img.style.width = `${w}px`; img.style.height = `${w * h0 / w0}px`; place(); };
        const up = () => {
          window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up);
          handle.remove();
          const k = img.getBoundingClientRect().width / w0;
          if (Math.abs(k - 1) < 0.01) return;
          run('resize the picture', () => ed.setImage({ id: info.id, cx: info.cx * k, cy: info.cy * k }))();
        };
        window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
      });
    });
    ed.stage.addEventListener('scroll', () => document.querySelectorAll('.dd-img-handle').forEach((x) => x.remove()), true);
  }
  async function imageOptions(ed) {
    const img = ed.selectedImage;
    if (!img || !img.isConnected) throw new Error('Click a picture first.');
    const info = ed.imageInfo(img);
    if (!info) throw new Error('This picture can\'t be changed here.');
    const inch = (emu) => (emu / 914400).toFixed(2);
    const v = await form('Picture options', [
      { name: 'w', label: 'Width (inches)', type: 'number', step: '0.01', value: inch(info.cx) },
      { name: 'h', label: 'Height (inches)', type: 'number', step: '0.01', value: inch(info.cy) },
      { name: 'lock', label: 'Keep proportions', type: 'checkbox', value: true },
      { name: 'alt', label: 'Alt text (describes the picture for screen readers)', type: 'textarea', value: info.alt },
      { name: 'l', label: 'Crop left (%)', type: 'number', step: '1', min: '0', value: info.crop.l }, { name: 'r', label: 'Crop right (%)', type: 'number', step: '1', min: '0', value: info.crop.r },
      { name: 't', label: 'Crop top (%)', type: 'number', step: '1', min: '0', value: info.crop.t }, { name: 'b', label: 'Crop bottom (%)', type: 'number', step: '1', min: '0', value: info.crop.b },
    ], 'Apply');
    if (!v) return;
    let cx = parseFloat(v.w) * 914400, cy = parseFloat(v.h) * 914400;
    if (v.lock) { const kx = cx / info.cx, ky = cy / info.cy; const k = Math.abs(kx - 1) > Math.abs(ky - 1) ? kx : ky; cx = info.cx * k; cy = info.cy * k; }
    const c = (x) => Math.max(0, Math.min(90, parseFloat(x) || 0));
    await ed.setImage({ id: info.id, cx, cy, alt: v.alt, crop: { l: c(v.l), t: c(v.t), r: c(v.r), b: c(v.b) } });
  }

  // ---------------------------------------------------------------- dropdown chips

  function wireDropdowns(tab) {
    const ed = tab.editor;
    if (!ed || ed._ddWired === ed.stage) return;
    ed._ddWired = ed.stage;
    ed.stage.addEventListener('click', (e) => {
      const span = e.target.closest && e.target.closest('[data-dd-r]');
      if (!span || tab.mode !== 'edit') return;
      const items = ed.dropdownAt(span);
      if (!items) return;
      const r = span.getBoundingClientRect();
      SheetView.showMenu(r.left, r.bottom + 2, items.map((x) => [x, '', () => {
        // Replace the chip's text; the content control around it is kept when saved.
        const t = [...span.childNodes].find((n) => n.nodeType === 3) || span;
        t.textContent = x;
        span.closest('[contenteditable="true"]').dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertReplacementText' }));
      }]));
    });
    ed.stage.querySelectorAll('[data-dd-r]').forEach((s) => { if (ed.dropdownAt(s)) s.classList.add('dd-chip'); });
  }
  async function dropdownChip(ed) {
    const v = await form('Dropdown', [{ name: 'items', label: 'Options (one per line)', type: 'textarea', value: 'Not started\nIn progress\nDone' }], 'Insert');
    if (!v) return;
    const items = v.items.split('\n').map((x) => x.trim()).filter(Boolean);
    if (!items.length) return;
    DocTools.restore();
    await ed.insertDropdown(items);
  }

  // ---------------------------------------------------------------- list styles, borders & shading

  const LISTS = [['bullet', '• Bullets'], ['arrow', '➢ Arrows'], ['check', '✓ Check marks'], ['square', '▪ Squares'], ['dash', '– Dashes'], ['star', '★ Stars'], ['number', '1. 2. 3.'], ['paren', '1) 2) 3)'], ['letter', 'a. b. c.'], ['LETTER', 'A. B. C.'], ['roman', 'i. ii. iii.'], ['ROMAN', 'I. II. III.']];
  function listStyle(ed, kind) {
    if (!DocTools.restore()) throw new Error('Click in the paragraphs to make into a list.');
    return ed.toggleList(kind);
  }
  async function bordersShading(ed) {
    if (!DocTools.restore()) throw new Error('Click in a paragraph first.');
    const ps = ed.selectedParagraphs();
    const v = await form('Borders and shading', [
      { name: 'border', label: 'Border', type: 'select', value: 'box', options: [['box', 'Box around the paragraph'], ['none', 'No border']] },
      { name: 'color', label: 'Border colour', type: 'color', value: '#000000' },
      { name: 'width', label: 'Border width', type: 'select', value: '6', options: [['4', '½ pt'], ['6', '¾ pt'], ['12', '1½ pt'], ['18', '2¼ pt'], ['24', '3 pt']] },
      { name: 'shade', label: 'Background colour', type: 'color', value: '#ffffff' },
    ], 'Apply');
    if (!v) return;
    for (const p of ps) {
      const cur = p.dataset.ddBorder ? JSON.parse(p.dataset.ddBorder) : {};
      cur.border = v.border === 'none' ? null : { color: v.color, size: +v.width };
      cur.shade = v.shade.toLowerCase() === '#ffffff' ? null : v.shade;
      p.dataset.ddBorder = JSON.stringify(cur);
      p.style.border = cur.border ? `${(+v.width / 8) * 1.33}px solid ${v.color}` : '';
      p.style.padding = cur.border ? '4px' : '';
      p.style.background = cur.shade || '';
    }
    if (ed.opts.onChange) ed.opts.onChange();
  }

  // ---------------------------------------------------------------- ruler (left, first-line and right indents)

  function toggleRuler(tab) {
    if (tab.rulerEl) { tab.rulerEl.remove(); tab.rulerEl = null; clearInterval(tab.rulerTimer); return; }
    const ed = tab.editor;
    if (!ed) return;
    const ruler = h('div', { class: 'dd-ruler' });
    const track = h('div', { class: 'dd-ruler-track' });
    const left = h('b', { class: 'dd-ind left', title: 'Left indent' });
    const first = h('b', { class: 'dd-ind first', title: 'First-line indent' });
    const right = h('b', { class: 'dd-ind right', title: 'Right indent' });
    track.append(left, first, right);
    ruler.append(track);
    tab.editBox.prepend(ruler);
    tab.rulerEl = ruler;
    const TW = 15; // twips per CSS pixel (1440 twips per 96 px)
    const state = () => {
      const p = ed.caretParagraph();
      const art = p && p.closest('article');
      if (!p || !art) return null;
      const ar = art.getBoundingClientRect(), pr = p.getBoundingClientRect();
      const cs = getComputedStyle(p);
      return { p, ar, left: pr.left - ar.left, right: ar.right - pr.right, first: parseFloat(cs.textIndent) || 0 };
    };
    const paint = () => {
      const s = state();
      ruler.classList.toggle('off', !s);
      if (!s) return;
      const rr = ruler.getBoundingClientRect();
      track.style.left = `${s.ar.left - rr.left}px`;
      track.style.width = `${s.ar.width}px`;
      // ticks every quarter inch
      if (track.dataset.w !== String(Math.round(s.ar.width))) {
        track.dataset.w = String(Math.round(s.ar.width));
        track.querySelectorAll('.tick').forEach((t) => t.remove());
        for (let x = 0, i = 0; x <= s.ar.width; x += 24, i++) track.append(h('i', { class: `tick${i % 4 === 0 ? ' major' : ''}`, style: `left:${x}px`, text: i % 4 === 0 && i ? String(i / 4) : '' }));
      }
      left.style.left = `${s.left}px`;
      first.style.left = `${s.left + s.first}px`;
      right.style.left = `${s.ar.width - s.right}px`;
    };
    const drag = (which) => (e) => {
      e.preventDefault();
      const s = state();
      if (!s) return;
      const ps = ed.selectedParagraphs();
      const x0 = e.clientX;
      const base = { left: s.left, right: s.right, first: s.first };
      const move = (ev) => {
        const dx = Math.round((ev.clientX - x0) / 6) * 6; // snap to 1/16 inch
        const v = { ...base };
        if (which === 'left') v.left = Math.max(0, base.left + dx);
        if (which === 'first') v.first = base.first + dx;
        if (which === 'right') v.right = Math.max(0, base.right - dx);
        for (const p of ps) {
          p.dataset.ddInd = JSON.stringify({ left: v.left * TW, right: v.right * TW, firstLine: v.first * TW });
          p.style.marginLeft = `${v.left}px`; p.style.marginRight = `${v.right}px`; p.style.textIndent = `${v.first}px`;
        }
        paint();
      };
      const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); if (ed.opts.onChange) ed.opts.onChange(); ed.drawPageMarkers(); };
      window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
    };
    left.addEventListener('mousedown', drag('left'));
    first.addEventListener('mousedown', drag('first'));
    right.addEventListener('mousedown', drag('right'));
    document.addEventListener('selectionchange', paint);
    tab.rulerTimer = setInterval(() => { if (!ruler.isConnected) { clearInterval(tab.rulerTimer); document.removeEventListener('selectionchange', paint); } else paint(); }, 1000);
    paint();
  }

  function togglePageless(tab) { tab.body.classList.toggle('pageless'); if (tab.editor) tab.editor.drawPageMarkers(); flash(tab.body.classList.contains('pageless') ? 'Pageless: text fills the window' : 'Pages'); }

  // ---------------------------------------------------------------- citations (sources kept with the document path)

  const srcKey = (tab) => `docstudio.sources:${tab.path || tab.name}`;
  const loadSources = (tab) => { try { return JSON.parse(localStorage.getItem(srcKey(tab)) || '[]'); } catch { return []; } };
  const saveSources = (tab, list) => { try { localStorage.setItem(srcKey(tab), JSON.stringify(list)); } catch { /* not kept */ } };
  function cite(src, style) {
    const authors = src.author.split(/;|\band\b/).map((x) => x.trim()).filter(Boolean);
    const last = (a) => (a.includes(',') ? a.split(',')[0].trim() : a.split(/\s+/).pop());
    const names = authors.length > 2 ? `${last(authors[0])} et al.` : authors.map(last).join(style === 'APA' ? ' & ' : ' and ');
    if (style === 'MLA') return `(${names}${src.page ? ` ${src.page}` : ''})`;
    if (style === 'Chicago') return `(${names} ${src.year}${src.page ? `, ${src.page}` : ''})`;
    return `(${names}, ${src.year})`;
  }
  function reference(src, style) {
    const t = src.title, y = src.year, a = src.author, pub = src.publisher, url = src.url;
    if (style === 'MLA') return `${a}. ${t}. ${pub ? `${pub}, ` : ''}${y}.${url ? ` ${url}.` : ''}`;
    if (style === 'Chicago') return `${a}. ${y}. ${t}. ${pub ? `${pub}.` : ''}${url ? ` ${url}.` : ''}`;
    return `${a} (${y}). ${t}.${pub ? ` ${pub}.` : ''}${url ? ` ${url}` : ''}`;
  }
  async function citations(tab) {
    const ed = tab.editor;
    const list = loadSources(tab);
    const box = h('div', { class: 'dialog versions-dialog' }, h('h3', { text: 'Citations' }));
    const styleSel = h('select', null, ['APA', 'MLA', 'Chicago'].map((x) => h('option', { value: x, text: x })));
    styleSel.value = localStorage.getItem('docstudio.citeStyle') || 'APA';
    styleSel.addEventListener('change', () => { try { localStorage.setItem('docstudio.citeStyle', styleSel.value); } catch { /* fine */ } });
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    box.append(h('label', null, h('span', { text: 'Style' }), styleSel));
    const rows = h('div', { class: 'version-list' });
    const fill = () => {
      rows.replaceChildren(...(list.length ? list.map((s, i) => h('div', { class: 'version-row' }, h('span', { text: `${s.author} (${s.year}) ${s.title}` }), h('span'),
        h('button', { class: 'ghost', text: 'Cite', onclick: () => { close(); DocTools.insertText(` ${cite(s, styleSel.value)}`); } }),
        h('button', { class: 'ghost', text: 'Remove', onclick: () => { list.splice(i, 1); saveSources(tab, list); fill(); } }))) : [h('p', { class: 'muted', text: 'No sources yet. Add one, then click Cite to put the citation at the cursor.' })]));
    };
    fill();
    box.append(rows, h('div', { class: 'dialog-actions' },
      h('button', { class: 'ghost', text: 'Add source…', onclick: async () => {
        const v = await form('Add source', [{ name: 'author', label: 'Author(s) (Last, First; separate several with ;)', value: '' }, { name: 'title', label: 'Title', value: '' }, { name: 'year', label: 'Year', value: String(new Date().getFullYear()) }, { name: 'publisher', label: 'Publisher or journal', value: '' }, { name: 'url', label: 'Web address (optional)', value: '' }, { name: 'page', label: 'Page (optional)', value: '' }], 'Add');
        if (v && v.author && v.title) { list.push(v); saveSources(tab, list); fill(); }
      } }),
      h('button', { class: 'ghost', text: 'Insert bibliography', onclick: run('insert the bibliography', async () => {
        close();
        if (!list.length) return;
        const sorted = [...list].sort((a, b) => a.author.localeCompare(b.author));
        const heading = styleSel.value === 'MLA' ? 'Works Cited' : styleSel.value === 'Chicago' ? 'Bibliography' : 'References';
        DocTools.restore();
        document.execCommand('insertParagraph');
        DocTools.insertText(heading);
        for (const s of sorted) { document.execCommand('insertParagraph'); document.execCommand('insertText', false, reference(s, styleSel.value)); }
        void ed;
      }) }),
      h('button', { text: 'Close', onclick: close })));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
  }

  // ---------------------------------------------------------------- dictionary (online lookup of the selected word)

  async function dictionary() {
    DocTools.restore();
    let word = getSelection().toString().trim();
    if (active && active.textarea && document.activeElement === active.textarea) word = active.textarea.value.slice(active.textarea.selectionStart, active.textarea.selectionEnd).trim();
    // The lookup is online, so the word is always shown first with where it goes (also for a selected word).
    word = word.split(/\s+/)[0].replace(/[^\p{L}'-]/gu, '');
    const v = await form('Dictionary', [{ name: 'w', label: 'Word (sent to dictionaryapi.dev to look it up)', value: word }], 'Look up');
    if (!v) return;
    word = v.w.trim().split(/\s+/)[0].replace(/[^\p{L}'-]/gu, '');
    if (!word) return;
    const r = await desk.dictionary(word);
    if (!r || !r.length) { desk.confirm({ type: 'info', buttons: ['OK'], message: `No definition found for "${word}"` }); return; }
    const lines = [];
    for (const entry of r.slice(0, 2)) {
      if (entry.phonetic) lines.push(entry.phonetic);
      for (const m of entry.meanings || []) {
        lines.push('', m.partOfSpeech);
        (m.definitions || []).slice(0, 3).forEach((d, i) => lines.push(`  ${i + 1}. ${d.definition}${d.example ? `\n     "${d.example}"` : ''}`));
        if (m.synonyms && m.synonyms.length) lines.push(`  Synonyms: ${m.synonyms.slice(0, 6).join(', ')}`);
      }
    }
    desk.confirm({ type: 'info', buttons: ['OK'], message: word, detail: lines.join('\n') + '\n\n(Definitions from dictionaryapi.dev)' });
  }

  // ---------------------------------------------------------------- drawing canvas (a picture for any editor)

  function drawing(title) {
    return new Promise((resolve) => {
      const W = 800, H = 480;
      const canvas = h('canvas', { width: W, height: H, class: 'draw-canvas' });
      const g = canvas.getContext('2d');
      g.fillStyle = '#ffffff'; g.fillRect(0, 0, W, H);
      const history = [];
      const snap = () => { history.push(g.getImageData(0, 0, W, H)); if (history.length > 40) history.shift(); };
      let tool = 'pen', color = '#1a73e8', width = 3, fill = false;
      const tools = h('div', { class: 'draw-tools' });
      const toolBtn = (id, label) => { const b = h('button', { class: `ghost${id === tool ? ' on' : ''}`, text: label, onclick: () => { tool = id; tools.querySelectorAll('button[data-t]').forEach((x) => x.classList.toggle('on', x.dataset.t === id)); } }); b.dataset.t = id; return b; };
      const col = h('input', { type: 'color', value: color }); col.addEventListener('input', () => { color = col.value; });
      const wid = h('select', null, [1, 2, 3, 5, 8, 12].map((x) => h('option', { value: x, text: `${x}px` }))); wid.value = '3'; wid.addEventListener('change', () => { width = +wid.value; });
      const fl = h('input', { type: 'checkbox' }); fl.addEventListener('change', () => { fill = fl.checked; });
      tools.append(toolBtn('pen', '✎ Pen'), toolBtn('line', '╱ Line'), toolBtn('arrow', '→ Arrow'), toolBtn('rect', '▭ Box'), toolBtn('ellipse', '◯ Oval'), toolBtn('text', 'T Text'), toolBtn('erase', '⌫ Eraser'), col, wid, h('label', { class: 'check-row' }, fl, h('span', { text: 'Filled' })),
        h('button', { class: 'ghost', text: 'Undo', onclick: () => { const d = history.pop(); if (d) g.putImageData(d, 0, 0); } }),
        h('button', { class: 'ghost', text: 'Clear', onclick: () => { snap(); g.fillStyle = '#fff'; g.fillRect(0, 0, W, H); } }));
      const box = h('div', { class: 'dialog draw-dialog' }, h('h3', { text: title || 'Drawing' }), tools, canvas);
      const back = h('div', { class: 'backdrop' }, box);
      const done = (v) => { back.remove(); resolve(v); };
      const pos = (e) => { const r = canvas.getBoundingClientRect(); return [(e.clientX - r.left) * W / r.width, (e.clientY - r.top) * H / r.height]; };
      canvas.addEventListener('mousedown', async (e) => {
        const [x0, y0] = pos(e);
        snap();
        if (tool === 'text') {
          const v = await form('Text', [{ name: 't', label: 'Text', value: '' }, { name: 's', label: 'Size (px)', type: 'number', value: '24' }], 'Add');
          if (v && v.t) { g.fillStyle = color; g.font = `${+v.s || 24}px "Segoe UI", sans-serif`; g.textBaseline = 'top'; g.fillText(v.t, x0, y0); }
          return;
        }
        const base = g.getImageData(0, 0, W, H);
        g.lineCap = 'round'; g.lineJoin = 'round';
        let last = [x0, y0];
        const shape = (x1, y1) => {
          g.putImageData(base, 0, 0);
          g.strokeStyle = color; g.fillStyle = color; g.lineWidth = width;
          g.beginPath();
          if (tool === 'rect') { g.rect(x0, y0, x1 - x0, y1 - y0); }
          else if (tool === 'ellipse') { g.ellipse((x0 + x1) / 2, (y0 + y1) / 2, Math.abs(x1 - x0) / 2, Math.abs(y1 - y0) / 2, 0, 0, Math.PI * 2); }
          else { g.moveTo(x0, y0); g.lineTo(x1, y1); }
          if (fill && (tool === 'rect' || tool === 'ellipse')) g.fill(); else g.stroke();
          if (tool === 'arrow') { const a = Math.atan2(y1 - y0, x1 - x0), L = 10 + width * 2; g.beginPath(); g.moveTo(x1, y1); g.lineTo(x1 - L * Math.cos(a - 0.45), y1 - L * Math.sin(a - 0.45)); g.lineTo(x1 - L * Math.cos(a + 0.45), y1 - L * Math.sin(a + 0.45)); g.closePath(); g.fill(); }
        };
        const move = (ev) => {
          const [x, y] = pos(ev);
          if (tool === 'pen' || tool === 'erase') { g.strokeStyle = tool === 'erase' ? '#ffffff' : color; g.lineWidth = tool === 'erase' ? width * 4 : width; g.beginPath(); g.moveTo(last[0], last[1]); g.lineTo(x, y); g.stroke(); last = [x, y]; } else shape(x, y);
        };
        const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
        window.addEventListener('mousemove', move); window.addEventListener('mouseup', up);
      });
      box.append(h('div', { class: 'dialog-actions' }, h('button', { class: 'ghost', text: 'Cancel', onclick: () => done(null) }), h('button', { text: 'Insert', onclick: async () => {
        // Trim the white margin around the drawing.
        const d = g.getImageData(0, 0, W, H).data;
        let x1 = W, y1 = H, x2 = 0, y2 = 0;
        for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { const i = (y * W + x) * 4; if (d[i] < 250 || d[i + 1] < 250 || d[i + 2] < 250) { if (x < x1) x1 = x; if (y < y1) y1 = y; if (x > x2) x2 = x; if (y > y2) y2 = y; } }
        if (x2 <= x1) { done(null); return; }
        const pad = 8;
        x1 = Math.max(0, x1 - pad); y1 = Math.max(0, y1 - pad); x2 = Math.min(W - 1, x2 + pad); y2 = Math.min(H - 1, y2 + pad);
        const out = h('canvas', { width: x2 - x1 + 1, height: y2 - y1 + 1 });
        out.getContext('2d').drawImage(canvas, x1, y1, out.width, out.height, 0, 0, out.width, out.height);
        const bytes = await (await new Promise((r) => out.toBlob(r, 'image/png'))).arrayBuffer();
        done({ bytes, ext: 'png', width: out.width, height: out.height });
      } })));
      back.addEventListener('keydown', (e) => { if (e.key === 'Escape') done(null); });
      document.body.append(back);
    });
  }

  // ---------------------------------------------------------------- chart maker (a picture for documents)

  async function chartImage() {
    const v = await form('Insert chart', [
      { name: 'type', label: 'Chart type', type: 'select', value: 'bar', options: [['bar', 'Column'], ['hbar', 'Bar'], ['line', 'Line'], ['pie', 'Pie'], ['doughnut', 'Doughnut']] },
      { name: 'title', label: 'Title', value: '' },
      { name: 'data', label: 'Data: first row = series names, first column = labels (comma or tab separated)', type: 'textarea', value: 'Quarter, Sales, Costs\nQ1, 120, 80\nQ2, 140, 90\nQ3, 160, 95\nQ4, 180, 100' },
    ], 'Insert');
    if (!v) return null;
    const rows = v.data.split('\n').map((l) => l.split(/\t|,/).map((x) => x.trim())).filter((r) => r.some(Boolean));
    if (rows.length < 2) throw new Error('Enter a header row and at least one row of numbers.');
    const labels = rows.slice(1).map((r) => r[0]);
    const palette = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E'];
    const pie = v.type === 'pie' || v.type === 'doughnut';
    const datasets = rows[0].slice(1).map((name, i) => ({ label: name, data: rows.slice(1).map((r) => +r[i + 1] || 0), backgroundColor: pie ? labels.map((_, k) => palette[k % palette.length]) : palette[i % palette.length], borderColor: pie ? '#fff' : palette[i % palette.length], borderWidth: v.type === 'line' ? 3 : 1, fill: false, tension: 0.2 }));
    const canvas = h('canvas', { width: 900, height: 520 });
    const holder = h('div', { style: 'position:fixed;left:-10000px;top:0;width:900px;height:520px' }, canvas);
    document.body.append(holder);
    try {
      const chart = new Chart(canvas, { type: v.type === 'hbar' ? 'bar' : v.type, data: { labels, datasets: pie ? datasets.slice(0, 1) : datasets }, options: { indexAxis: v.type === 'hbar' ? 'y' : 'x', animation: false, responsive: false, devicePixelRatio: 2, plugins: { title: { display: !!v.title, text: v.title, font: { size: 20 } }, legend: { position: 'bottom' } } }, plugins: [{ id: 'white', beforeDraw: (c) => { const x = c.ctx; x.save(); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.restore(); } }] });
      chart.update();
      const bytes = await (await new Promise((r) => canvas.toBlob(r, 'image/png'))).arrayBuffer();
      chart.destroy();
      return { bytes, ext: 'png', width: 900, height: 520, rows, type: v.type, title: v.title };
    } finally { holder.remove(); }
  }

  // ---------------------------------------------------------------- menus added to the Word editor

  function wordMenus(tab) {
    const ed = () => tab.editor;
    const it = (label, fn, o) => ({ label, run: run(label.toLowerCase(), fn), ...(o || {}) });
    const sub = (label, items) => ({ label, sub: items });
    const SEP = { sep: true };
    const en = { enabled: () => tab.mode === 'edit' && !!tab.editor };
    const pic = async (p) => { if (!p) return; DocTools.restore(); await ed().insertImage(p.bytes, p.ext, p.width, p.height); };
    return {
      Insert: [
        it('Equation…', () => equation(ed()), en),
        it('Drawing…', async () => pic(await drawing('Drawing')), en),
        it('Chart…', async () => pic(await chartImage()), en),
        it('Dropdown…', () => dropdownChip(ed()), en),
        sub('Section break', [it('Next page', () => { DocTools.restore(); return ed().insertSectionBreak('nextPage'); }, en), it('Continuous', () => { DocTools.restore(); return ed().insertSectionBreak('continuous'); }, en)]),
      ],
      Format: [
        sub('Bullets & numbering', [...LISTS.map(([k, t]) => it(t, () => listStyle(ed(), k), en)), SEP, it('Restart numbering', () => { DocTools.restore(); return ed().restartNumbering(); }, en)]),
        it('Borders & shading…', () => bordersShading(ed()), en),
        it('Picture options…', () => imageOptions(ed()), { enabled: () => !!(tab.editor && tab.editor.selectedImage && tab.editor.selectedImage.isConnected) }),
        it('Update paragraph style to match', () => { DocTools.restore(); const p = ed().caretParagraph(); if (!p) throw new Error('Click in a paragraph first.'); return ed().updateStyleToMatch(p); }, en),
      ],
      View: [
        it('Show ruler', () => toggleRuler(tab), { ...en, checked: () => !!tab.rulerEl }),
        it('Pageless', () => togglePageless(tab), { checked: () => tab.body.classList.contains('pageless') }),
        it('Suggesting mode', () => setSuggesting(tab, !(tab.editor && tab.editor.suggesting)), { ...en, checked: () => !!(tab.editor && tab.editor.suggesting) }),
      ],
      Tools: [
        it('Review suggested edits…', () => reviewPanel(tab), { enabled: () => !!tab.editor, key: 'Ctrl+Alt+O' }),
        it('Citations…', () => citations(tab), en),
        it('Dictionary…', () => dictionary(), { key: 'Ctrl+Shift+Y' }),
      ],
    };
  }

  return { setSuggesting, reviewPanel, equation, wirePictures, wireDropdowns, imageOptions, toggleRuler, togglePageless, citations, dictionary, drawing, chartImage, wordMenus, LISTS, listStyle };
})();
