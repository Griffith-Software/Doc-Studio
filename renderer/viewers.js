'use strict';
// Read-only renderers. Each takes the file bytes and a container and fills it.
// They return an optional status string for the footer.

const Viewers = {
  paginate: (root) => paginate(root),
  // ---------- PDF: Chromium's own viewer (search, zoom, print, page thumbnails)
  async pdf(tab, data, box) {
    const url = URL.createObjectURL(new Blob([data], { type: 'application/pdf' }));
    tab.urls.push(url);
    box.classList.add('fill');
    box.append(h('iframe', { class: 'pdf-frame', src: url, title: tab.name }));
    return 'Use the PDF toolbar for search, zoom and printing';
  },

  // ---------- Word, faithful layout (pages, headers, footers, footnotes)
  async docx(tab, data, box) {
    const pages = h('div', { class: 'docx-view' });
    box.append(pages);
    await docx.renderAsync(data, pages, null, {
      className: 'docx', inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: true,
      renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true,
      renderChanges: true, renderComments: true, useBase64URL: true, experimental: true,
    });
    pages.querySelectorAll('a[href]').forEach(linkOut);
    // docx-preview can list a note twice; keep the first copy.
    for (const ol of pages.querySelectorAll('ol')) {
      const seen = new Set();
      for (const li of [...ol.children]) { if (li.tagName !== 'LI') continue; if (seen.has(li.innerHTML)) li.remove(); else seen.add(li.innerHTML); }
    }
    if (pages.isConnected) paginate(pages);
    return `${pages.querySelectorAll('section.docx').length} page(s)`;
  },

  // ---------- spreadsheets Doc Studio can't edit in place (xls, xlsb, ods, numbers)
  async sheetview(tab, data, box) {
    const wb = XLSX.read(data, { type: 'array', cellDates: true, dense: true });
    return showWorkbook(wb, box);
  },

  // ---------- RTF
  async rtf(tab, data, box) {
    const els = await RtfEditor.render(data);
    const page = h('article', { class: 'paper rtf-view' });
    els.forEach((el) => page.append(el));
    page.querySelectorAll('a[href]').forEach(linkOut);
    box.append(page);
    return `${countWords(page.innerText)} words`;
  },

  // ---------- legacy .doc: text only
  async doc(tab, data, box) {
    const r = await desk.docExtract(data);
    const page = h('article', { class: 'paper doc-view' });
    if (r.headers && r.headers.trim()) page.append(h('div', { class: 'doc-header', text: r.headers.trim() }));
    r.body.split(/\n{2,}|\r/).forEach((para) => { if (para.trim()) page.append(h('p', { text: para.trim() })); });
    if (r.footers && r.footers.trim()) page.append(h('div', { class: 'doc-footer', text: r.footers.trim() }));
    box.append(h('div', { class: 'notice' }, 'Older .doc files (97-2003) are shown as text only - formatting and images are not displayed.'), page);
    return `${countWords(r.body)} words`;
  },

  // ---------- OpenDocument text
  async odt(tab, data, box) {
    const zip = await JSZip.loadAsync(data);
    const xml = new DOMParser().parseFromString(await zip.file('content.xml').async('string'), 'application/xml');
    const styles = await odfStyles(zip, xml);
    const page = h('article', { class: 'paper odt-view' });
    const body = xml.getElementsByTagNameNS(ODF.office, 'text')[0];
    if (body) for (const n of body.childNodes) { const el = await odfNode(n, styles, zip, tab); if (el) page.append(el); }
    box.append(page);
    return `${countWords(page.innerText)} words`;
  },

  // ---------- OpenDocument presentation: slide text and pictures
  async odp(tab, data, box) {
    const zip = await JSZip.loadAsync(data);
    const xml = new DOMParser().parseFromString(await zip.file('content.xml').async('string'), 'application/xml');
    const styles = await odfStyles(zip, xml);
    const slides = [...xml.getElementsByTagNameNS(ODF.draw, 'page')];
    const wrap = h('div', { class: 'slides-text' });
    let i = 0;
    for (const s of slides) {
      i++;
      const card = h('section', { class: 'slide-card' }, h('div', { class: 'slide-num', text: `Slide ${i}${s.getAttributeNS(ODF.draw, 'name') ? ' - ' + s.getAttributeNS(ODF.draw, 'name') : ''}` }));
      for (const frame of s.getElementsByTagNameNS(ODF.draw, 'frame')) {
        for (const n of frame.childNodes) {
          if (n.localName === 'text-box') { for (const c of n.childNodes) { const el = await odfNode(c, styles, zip, tab); if (el) card.append(el); } }
          if (n.localName === 'image') { const img = await odfImage(n, zip, tab); if (img) card.append(img); }
        }
      }
      wrap.append(card);
    }
    box.append(h('div', { class: 'notice' }, 'OpenDocument presentations are shown as slide text and pictures, not full slide layouts.'), wrap);
    return `${slides.length} slide(s)`;
  },

  // ---------- EPUB: one chapter at a time
  async epub(tab, data, box) {
    const zip = await JSZip.loadAsync(data);
    const read = (p) => zip.file(p) ? zip.file(p).async('string') : Promise.resolve(null);
    const container = new DOMParser().parseFromString(await read('META-INF/container.xml'), 'application/xml');
    const opfPath = container.querySelector('rootfile').getAttribute('full-path');
    const opfDir = opfPath.includes('/') ? opfPath.slice(0, opfPath.lastIndexOf('/') + 1) : '';
    const opf = new DOMParser().parseFromString(await read(opfPath), 'application/xml');
    const manifest = {};
    opf.querySelectorAll('manifest > item').forEach((it) => {
      manifest[it.getAttribute('id')] = { href: resolvePath(opfDir, decodeURIComponent(it.getAttribute('href'))), type: it.getAttribute('media-type') };
    });
    const spine = [...opf.querySelectorAll('spine > itemref')].map((r) => manifest[r.getAttribute('idref')]).filter(Boolean);
    const title = (opf.getElementsByTagNameNS('*', 'title')[0] || {}).textContent || tab.name;
    const blobCache = {};
    const blobFor = async (p) => {
      if (blobCache[p]) return blobCache[p];
      const f = zip.file(p);
      if (!f) return null;
      const url = URL.createObjectURL(new Blob([await f.async('arraybuffer')], { type: FORMATS.mimeOf(FORMATS.extOf(p)) }));
      tab.urls.push(url);
      return (blobCache[p] = url);
    };
    const rewriteCss = async (css, dir) => {
      const parts = [];
      let last = 0;
      for (const m of css.matchAll(/url\(\s*['"]?([^'")]+)['"]?\s*\)/g)) {
        parts.push(css.slice(last, m.index));
        const u = /^(data:|https?:)/.test(m[1]) ? null : await blobFor(resolvePath(dir, m[1]));
        parts.push(u ? `url("${u}")` : m[0]);
        last = m.index + m[0].length;
      }
      return parts.join('') + css.slice(last);
    };

    const frame = h('iframe', { class: 'epub-frame', sandbox: 'allow-same-origin', title });
    const select = h('select', { class: 'epub-toc', 'aria-label': 'Chapter' }, spine.map((s, i) => h('option', { value: i, text: `${i + 1}. ${s.href.split('/').pop()}` })));
    const prev = h('button', { class: 'ghost', text: '‹ Previous' });
    const next = h('button', { class: 'ghost', text: 'Next ›' });
    let index = 0;
    const show = async (i) => {
      index = Math.max(0, Math.min(spine.length - 1, i));
      select.value = index;
      const item = spine[index];
      const dir = item.href.includes('/') ? item.href.slice(0, item.href.lastIndexOf('/') + 1) : '';
      const doc = new DOMParser().parseFromString(await read(item.href) || '', item.type === 'text/html' ? 'text/html' : 'application/xhtml+xml');
      doc.querySelectorAll('script').forEach((s) => s.remove());
      for (const img of doc.querySelectorAll('img[src]')) { const u = await blobFor(resolvePath(dir, decodeURIComponent(img.getAttribute('src')))); if (u) img.setAttribute('src', u); }
      for (const img of doc.getElementsByTagNameNS('http://www.w3.org/2000/svg', 'image')) {
        const href = img.getAttributeNS('http://www.w3.org/1999/xlink', 'href') || img.getAttribute('href');
        const u = href && await blobFor(resolvePath(dir, decodeURIComponent(href)));
        if (u) { img.setAttributeNS('http://www.w3.org/1999/xlink', 'href', u); img.setAttribute('href', u); }
      }
      let css = '';
      for (const link of doc.querySelectorAll('link[rel~="stylesheet"][href]')) {
        const p = resolvePath(dir, decodeURIComponent(link.getAttribute('href')));
        const text = await read(p);
        if (text) css += await rewriteCss(text, p.includes('/') ? p.slice(0, p.lastIndexOf('/') + 1) : '');
        link.remove();
      }
      for (const st of doc.querySelectorAll('style')) { st.textContent = await rewriteCss(st.textContent, dir); }
      const base = 'body{max-width:44em;margin:2em auto;padding:0 1.5em;font:17px/1.6 Georgia,serif;color:#222;background:#fff}img,svg{max-width:100%;height:auto}';
      const bodyHtml = doc.body ? doc.body.innerHTML : new XMLSerializer().serializeToString(doc.documentElement);
      const headStyles = [...doc.querySelectorAll('style')].map((s) => s.outerHTML).join('');
      frame.srcdoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${base}</style><style>${css}</style>${headStyles}</head><body>${bodyHtml}</body></html>`;
      prev.disabled = index === 0; next.disabled = index === spine.length - 1;
      tab.setStatus && tab.setStatus(`${title} - chapter ${index + 1} of ${spine.length}`);
    };
    frame.addEventListener('load', () => {
      const d = frame.contentDocument;
      if (!d) return;
      d.addEventListener('click', (e) => {
        const a = e.target.closest && e.target.closest('a[href]');
        if (!a) return;
        e.preventDefault();
        const href = a.getAttribute('href');
        if (/^(https?:|mailto:)/i.test(href)) { desk.openExternal(href); return; }
        const dir = spine[index].href.includes('/') ? spine[index].href.slice(0, spine[index].href.lastIndexOf('/') + 1) : '';
        const [file, hash] = href.split('#');
        const target = file ? resolvePath(dir, decodeURIComponent(file)) : spine[index].href;
        const j = spine.findIndex((s) => s.href === target);
        if (j >= 0 && j !== index) show(j).then(() => hash && frame.addEventListener('load', () => frame.contentDocument.getElementById(hash)?.scrollIntoView(), { once: true }));
        else if (hash) d.getElementById(hash)?.scrollIntoView();
      });
    });
    prev.onclick = () => show(index - 1);
    next.onclick = () => show(index + 1);
    select.onchange = () => show(+select.value);
    box.classList.add('fill');
    box.append(h('div', { class: 'epub-bar' }, h('strong', { text: title }), h('span', { class: 'spacer' }), prev, select, next), frame);
    await show(0);
    return `${spine.length} chapter(s)`;
  },

  // ---------- images with zoom
  async image(tab, data, box) {
    const url = URL.createObjectURL(new Blob([data], { type: FORMATS.mimeOf(tab.ext) }));
    tab.urls.push(url);
    const img = h('img', { src: url, alt: tab.name, draggable: 'false' });
    const view = h('div', { class: 'image-view fit' }, img);
    let zoom = 1;
    const label = h('span', { class: 'zoom-label', text: 'Fit' });
    const set = (z) => {
      if (z === 'fit') { view.classList.add('fit'); img.style.width = ''; label.textContent = 'Fit'; return; }
      view.classList.remove('fit');
      zoom = Math.max(0.05, Math.min(16, z));
      img.style.width = `${img.naturalWidth * zoom}px`;
      label.textContent = `${Math.round(zoom * 100)}%`;
    };
    const current = () => (view.classList.contains('fit') ? img.clientWidth / img.naturalWidth : zoom);
    tab.tools = [
      h('button', { class: 'ghost', title: 'Zoom out', text: '−', onclick: () => set(current() / 1.25) }),
      label,
      h('button', { class: 'ghost', title: 'Zoom in', text: '+', onclick: () => set(current() * 1.25) }),
      h('button', { class: 'ghost', text: 'Fit', onclick: () => set('fit') }),
      h('button', { class: 'ghost', text: '100%', onclick: () => set(1) }),
    ];
    view.addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); e.stopPropagation(); set(current() * (e.deltaY < 0 ? 1.1 : 1 / 1.1)); } }, { passive: false });
    box.classList.add('fill');
    box.append(view);
    await img.decode().catch(() => { throw new Error('This image could not be decoded.'); });
    return `${img.naturalWidth} × ${img.naturalHeight} px`;
  },

  async audio(tab, data, box) { return media('audio', tab, data, box); },
  async video(tab, data, box) { return media('video', tab, data, box); },
};

// docx-preview only starts a new page at explicit breaks; move content that overflows a page onto a new page
// (whole paragraphs/tables), repeating the header and footer, so View shows real pages.
function paginate(root) {
  for (let guard = 0; guard < 2000; guard++) {
    let split = false;
    for (const sec of root.querySelectorAll('section.docx')) {
      const art = sec.querySelector(':scope > article');
      const pageH = parseFloat(sec.style.minHeight);
      if (!art || !pageH || sec.offsetHeight <= pageH + 2) continue;
      const footer = sec.querySelector(':scope > footer');
      const limit = pageH - parseFloat(getComputedStyle(sec).paddingBottom) - (footer ? footer.offsetHeight + 8 : 0);
      const blocks = [...art.children];
      let cut = blocks.findIndex((b) => b.offsetTop + b.offsetHeight > limit);
      if (cut === 0) cut = 1;
      if (cut < 0 || cut >= blocks.length) continue;
      const next = sec.cloneNode(false);
      const header = sec.querySelector(':scope > header');
      if (header) next.append(header.cloneNode(true));
      const nart = art.cloneNode(false);
      next.append(nart);
      for (const b of blocks.slice(cut)) nart.append(b);
      if (footer) next.append(footer.cloneNode(true));
      sec.after(next);
      split = true;
      break;
    }
    if (!split) break;
  }
}

function media(tag, tab, data, box) {
  const url = URL.createObjectURL(new Blob([data], { type: FORMATS.mimeOf(tab.ext) }));
  tab.urls.push(url);
  box.classList.add('fill');
  box.append(h('div', { class: 'media-view' }, h(tag, { src: url, controls: true })));
  return '';
}

function linkOut(a) {
  a.addEventListener('click', (e) => {
    e.preventDefault();
    const href = a.getAttribute('href');
    if (/^(https?:|mailto:)/i.test(href)) desk.openExternal(href);
    else if (href.startsWith('#')) a.ownerDocument.getElementById(href.slice(1))?.scrollIntoView();
  });
}

function resolvePath(dir, rel) {
  const out = [];
  for (const part of (dir + rel).split('/')) {
    if (part === '..') out.pop();
    else if (part && part !== '.') out.push(part);
  }
  return out.join('/');
}

// ---------- workbook grid (shared by spreadsheets and CSV table view)

const MAX_ROWS = 5000;
function showWorkbook(wb, box) {
  const grid = h('div', { class: 'sheet-grid' });
  const tabs = h('div', { class: 'sheet-tabs', role: 'tablist' });
  const render = (name) => {
    tabs.querySelectorAll('button').forEach((b) => b.classList.toggle('active', b.dataset.name === name));
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: true });
    const cols = rows.reduce((m, r) => Math.max(m, r.length), 0);
    const table = h('table');
    const head = h('tr', null, h('th', { class: 'corner' }));
    for (let c = 0; c < Math.max(cols, 1); c++) head.append(h('th', { text: XLSX.utils.encode_col(c) }));
    table.append(h('thead', null, head));
    const tbody = h('tbody');
    rows.slice(0, MAX_ROWS).forEach((r, i) => {
      const tr = h('tr', null, h('th', { text: i + 1 }));
      for (let c = 0; c < cols; c++) {
        const v = r[c] == null ? '' : String(r[c]);
        tr.append(h('td', { class: /^-?[\d,.]+%?$/.test(v) ? 'num' : null, text: v }));
      }
      tbody.append(tr);
    });
    table.append(tbody);
    grid.replaceChildren(table);
    if (rows.length > MAX_ROWS) grid.append(h('div', { class: 'notice', text: `Showing the first ${MAX_ROWS.toLocaleString()} of ${rows.length.toLocaleString()} rows.` }));
    grid.scrollTo(0, 0);
  };
  wb.SheetNames.forEach((n) => tabs.append(h('button', { 'data-name': n, text: n, onclick: () => render(n) })));
  box.classList.add('fill');
  box.append(h('div', { class: 'sheet-view' }, grid, wb.SheetNames.length > 1 ? tabs : null));
  render(wb.SheetNames[0]);
  return `${wb.SheetNames.length} sheet(s)`;
}

// ---------- OpenDocument helpers

const ODF = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
};

async function odfStyles(zip, contentXml) {
  const map = {};
  const collect = (doc) => {
    for (const st of doc.getElementsByTagNameNS(ODF.style, 'style')) {
      const tp = st.getElementsByTagNameNS(ODF.style, 'text-properties')[0];
      const pp = st.getElementsByTagNameNS(ODF.style, 'paragraph-properties')[0];
      const css = [];
      if (tp) {
        if (tp.getAttributeNS(ODF.fo, 'font-weight') === 'bold') css.push('font-weight:bold');
        if (tp.getAttributeNS(ODF.fo, 'font-style') === 'italic') css.push('font-style:italic');
        const u = tp.getAttributeNS(ODF.style, 'text-underline-style');
        if (u && u !== 'none') css.push('text-decoration:underline');
        if ((tp.getAttributeNS(ODF.style, 'text-line-through-style') || 'none') !== 'none') css.push('text-decoration:line-through');
        const color = tp.getAttributeNS(ODF.fo, 'color');
        if (color && color !== '#000000') css.push(`color:${color}`);
        const size = tp.getAttributeNS(ODF.fo, 'font-size');
        if (size && /pt$/.test(size)) css.push(`font-size:${size}`);
      }
      if (pp) {
        const align = pp.getAttributeNS(ODF.fo, 'text-align');
        if (align) css.push(`text-align:${align === 'end' ? 'right' : align === 'start' ? 'left' : align}`);
      }
      map[st.getAttributeNS(ODF.style, 'name')] = css.join(';');
    }
  };
  if (zip.file('styles.xml')) collect(new DOMParser().parseFromString(await zip.file('styles.xml').async('string'), 'application/xml'));
  collect(contentXml);
  return map;
}

async function odfImage(n, zip, tab) {
  const href = n.getAttributeNS(ODF.xlink, 'href');
  const f = href && zip.file(href);
  if (!f) return null;
  const url = URL.createObjectURL(new Blob([await f.async('arraybuffer')], { type: FORMATS.mimeOf(FORMATS.extOf(href)) }));
  tab.urls.push(url);
  return h('img', { src: url, alt: '' });
}

async function odfNode(n, styles, zip, tab) {
  if (n.nodeType === 3) return document.createTextNode(n.nodeValue);
  if (n.nodeType !== 1) return null;
  const kids = async (el) => { for (const c of n.childNodes) { const x = await odfNode(c, styles, zip, tab); if (x) el.append(x); } return el; };
  const style = styles[n.getAttributeNS(ODF.text, 'style-name')] || '';
  switch (n.localName) {
    case 'h': return kids(h(`h${Math.min(6, +(n.getAttributeNS(ODF.text, 'outline-level') || 1))}`, { style }));
    case 'p': return kids(h('p', { style }));
    case 'span': return kids(h('span', { style }));
    case 'a': {
      const href = n.getAttributeNS(ODF.xlink, 'href') || '';
      const a = await kids(h('a', { href }));
      linkOut(a);
      return a;
    }
    case 's': return document.createTextNode(' '.repeat(+(n.getAttributeNS(ODF.text, 'c') || 1)));
    case 'tab': return document.createTextNode('\t');
    case 'line-break': return h('br');
    case 'list': return kids(h('ul'));
    case 'list-item': return kids(h('li'));
    case 'table': return kids(h('table', { class: 'odf-table' }));
    case 'table-header-rows': case 'table-rows': return kids(document.createDocumentFragment());
    case 'table-row': return kids(h('tr'));
    case 'table-cell': return kids(h('td', { colspan: n.getAttributeNS(ODF.table, 'number-columns-spanned') || null }));
    case 'frame': return kids(document.createDocumentFragment());
    case 'text-box': return kids(h('div', { class: 'odf-box' }));
    case 'image': return odfImage(n, zip, tab);
    case 'section': case 'soft-page-break': case 'bookmark': case 'bookmark-start': case 'bookmark-end':
      return n.localName === 'section' ? kids(document.createDocumentFragment()) : null;
    case 'sequence-decls': case 'tracked-changes': case 'variable-decls': case 'user-field-decls': case 'annotation':
      return null;
    default: return kids(document.createDocumentFragment());
  }
}
