'use strict';
// File > Convert to… / Download as: turns the open document (with any unsaved edits) into another format.
//
// When Word, Excel, PowerPoint or LibreOffice is installed, conversions between office formats go through it for
// the closest result. Every conversion also works without them:
//   - PDF: the document is drawn by Doc Studio (pages, slides, sheets) and printed by Chromium in a hidden window.
//   - Word: from HTML (Markdown, web pages, OpenDocument/RTF pages, PDF text) with the docx library.
//   - Markdown: Word through mammoth (semantic HTML) then Turndown; other formats through their HTML.
//   - OpenDocument text and RTF: Doc Studio's own writers. EPUB: a small EPUB 3 package.
//   - Spreadsheets: SheetJS writers. Slides: PDF, PNG pictures and a text outline. PDF: text, Markdown, Word, PNG.

const Convert = (() => {
  const LABELS = {
    pdf: 'PDF document', docx: 'Document', odt: 'OpenDocument text', rtf: 'Rich Text', doc: 'Document (97-2003)',
    md: 'Markdown', html: 'Web page (HTML)', txt: 'Plain text', epub: 'EPUB book',
    xlsx: 'Spreadsheet', xls: 'Spreadsheet (97-2003)', xlsb: 'Binary spreadsheet', ods: 'OpenDocument spreadsheet',
    csv: 'CSV (current sheet)', tsv: 'TSV (current sheet)', json: 'JSON (current sheet)',
    pptx: 'Presentation', ppt: 'Presentation (97-2003)', odp: 'OpenDocument presentation', png: 'PNG pictures', jpg: 'JPEG picture', webp: 'WebP picture', bmp: 'BMP picture',
  };

  // What the open document can become.
  async function targets(tab) {
    const k = tab.kind;
    let list = [];
    if (k === 'docx') list = ['pdf', 'odt', 'rtf', 'doc', 'md', 'html', 'txt', 'epub'];
    else if (k === 'odt' || k === 'rtf') list = ['pdf', 'docx', k === 'odt' ? 'rtf' : 'odt', 'doc', 'md', 'html', 'txt', 'epub'];
    else if (k === 'md') list = ['pdf', 'docx', 'odt', 'rtf', 'html', 'txt', 'epub'];
    else if (k === 'html') list = ['pdf', 'docx', 'odt', 'rtf', 'md', 'txt', 'epub'];
    else if (k === 'text') list = ['pdf', 'docx', 'odt', 'rtf', 'md', 'html'];
    else if (k === 'csv') list = ['xlsx', 'ods', 'pdf', 'md', 'html', 'json'];
    else if (k === 'sheet') list = ['pdf', 'ods', 'xls', 'xlsb', 'csv', 'tsv', 'html', 'md', 'json'];
    else if (k === 'sheetview') list = ['xlsx', 'ods', 'csv', 'html', 'pdf'];
    else if (k === 'pptx') list = ['pdf', 'png', 'odp', 'ppt', 'md', 'txt'];
    else if (k === 'pdf') list = ['docx', 'md', 'txt', 'html', 'png'];
    else if (k === 'image') list = ['pdf', 'png', 'jpg', 'webp', 'bmp'];
    else if (k === 'epub') list = ['md', 'html', 'txt'];
    list = list.filter((x) => x !== tab.ext);
    const out = [];
    for (const ext of list) {
      // Office-only targets appear only when a program here can write them.
      if (['doc', 'ppt', 'odp'].includes(ext) && !(await route(officeNative(tab), ext))) continue;
      out.push({ ext, label: LABELS[ext] || ext.toUpperCase() });
    }
    return out;
  }
  // Which office program can convert (cached: the answer only changes if Office is installed or removed).
  const routes = new Map();
  async function route(from, to) {
    const key = `${from}>${to}`;
    if (!routes.has(key)) routes.set(key, await desk.convertRoute(from, to));
    return routes.get(key);
  }
  const officeNative = (tab) => ({ docx: 'docx', odt: 'odt', rtf: 'rtf', sheet: 'xlsx', pptx: 'pptx' }[tab.kind] || tab.ext);

  // ---------------------------------------------------------------- source content

  async function nativeBytes(tab) {
    switch (tab.kind) {
      case 'docx': return currentDocxBytes(tab);
      case 'sheet': if (tab.sheetView) tab.sheetView.endEdit(true); return tab.workbook.save();
      case 'pptx': return tab.pptEditor ? tab.pptEditor.save() : tab.data;
      case 'odt': return tab.odtEditor ? tab.odtEditor.save() : tab.data;
      case 'rtf': return tab.rtfEditor ? tab.rtfEditor.save() : tab.data;
      case 'pdf': return tab.pdfEditor && tab.dirty ? tab.pdfEditor.build() : tab.data;
      default: return tab.textarea ? new TextEncoder().encode(tab.textarea.value).buffer : tab.data;
    }
  }
  const textOf = (tab) => (tab.textarea ? tab.textarea.value : new TextDecoder().decode(tab.data || new ArrayBuffer(0)));
  const escapeHtml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  // Semantic HTML of a text document (headings, lists, tables, pictures as data: URLs).
  async function htmlOf(tab) {
    switch (tab.kind) {
      case 'md': return sanitize(marked.parse(textOf(tab), { gfm: true }));
      case 'html': return sanitize(textOf(tab), { WHOLE_DOCUMENT: false });
      case 'text': return textOf(tab).split(/\n{2,}/).map((p) => `<p>${escapeHtml(p).replace(/\n/g, '<br>')}</p>`).join('\n');
      case 'docx': {
        // Word's Title/Subtitle styles become headings too.
        const r = await mammoth.convertToHtml({ arrayBuffer: await currentDocxBytes(tab) }, { styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='Subtitle'] => h2:fresh", "p[style-name='Quote'] => blockquote > p:fresh", "p[style-name='Intense Quote'] => blockquote > p:fresh"] });
        return r.value;
      }
      case 'odt': case 'rtf': {
        const ed = tab.richEditor;
        const clone = ed.article.cloneNode(true);
        clone.querySelectorAll('[contenteditable]').forEach((e) => e.removeAttribute('contenteditable'));
        await inlineImages(clone);
        return clone.innerHTML;
      }
      case 'pdf': return (await pdfParagraphs(tab)).map((pg) => pg.map((p) => (p.heading ? `<h2>${escapeHtml(p.text)}</h2>` : `<p>${escapeHtml(p.text)}</p>`)).join('\n')).join('\n<hr>\n');
      case 'epub': { const box = tab.searchRoot && tab.searchRoot(); return box ? box.innerHTML : ''; }
      default: return '';
    }
  }

  // blob: pictures don't survive leaving this window; turn them into data: URLs.
  async function inlineImages(root) {
    for (const img of root.querySelectorAll('img')) {
      const src = img.getAttribute('src') || '';
      if (!src.startsWith('blob:')) continue;
      try {
        const b = await (await fetch(src)).blob();
        img.setAttribute('src', await new Promise((r) => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b); }));
      } catch { /* leave it */ }
    }
  }

  // ---------------------------------------------------------------- converting

  // opts.office === false skips Word/Excel/PowerPoint/LibreOffice (Doc Studio's own route only).
  async function run(tab, ext, opts) {
    const k = tab.kind;
    const native = officeNative(tab);
    // Office/LibreOffice first for office-to-office and office-to-PDF.
    if (!(opts && opts.office === false) && ['docx', 'odt', 'rtf', 'sheet', 'pptx'].includes(k) && ['pdf', 'docx', 'odt', 'rtf', 'doc', 'xls', 'xlsb', 'ods', 'ppt', 'odp', 'txt'].includes(ext) && await route(native, ext)) {
      return (await desk.convert(await nativeBytes(tab), native, ext)).data;
    }
    if (['doc', 'ppt', 'odp'].includes(ext)) throw new Error(`.${ext} files can only be written with Microsoft Office or LibreOffice installed.`);
    if (ext === 'pdf') return toPdf(tab);
    if (k === 'sheet' || k === 'sheetview' || k === 'csv') return fromSheet(tab, ext);
    if (k === 'pptx') return fromSlides(tab, ext);
    if (k === 'image') return fromImage(tab, ext);
    if (k === 'pdf' && ext === 'png') return pdfPages(tab);
    if (k === 'pdf' && ext === 'txt') return enc((await pdfParagraphs(tab)).map((pg) => pg.map((p) => p.text).join('\n\n')).join('\n\n\f'));
    const html = await htmlOf(tab);
    switch (ext) {
      case 'docx': return desk.docxFromHtml(html, stem(tab.name));
      case 'odt': return OdtEditor.fromHtml(html);
      case 'rtf': return rtfFromHtml(html);
      case 'md': return enc(toMarkdown(html));
      case 'html': return enc(standaloneHtml(stem(tab.name), html));
      case 'txt': return enc(htmlText(html));
      case 'epub': return epub(stem(tab.name), html);
      default: throw new Error(`Doc Studio can't convert ${FORMATS.label(k)} files to .${ext}.`);
    }
  }
  const enc = (s) => new TextEncoder().encode(s).buffer;

  function toMarkdown(html) {
    const td = new TurndownService({ headingStyle: 'atx', codeBlockStyle: 'fenced', bulletListMarker: '-', emDelimiter: '*' });
    td.use(turndownPluginGfm.gfm);
    // Pictures embedded as data: URLs would make the Markdown unreadable; keep their alt text as a placeholder.
    // GitHub tables need a header row: tables without one use their first row.
    const d = document.createElement('div');
    d.innerHTML = html;
    for (const t of d.querySelectorAll('table')) {
      if (t.querySelector('th')) continue;
      const first = t.querySelector('tr');
      if (!first) continue;
      for (const td0 of [...first.children]) { const th = document.createElement('th'); th.innerHTML = td0.innerHTML; td0.replaceWith(th); }
      // Block content inside cells would break the table; keep the text only.
      t.querySelectorAll('td p, th p').forEach((p) => { p.after(' '); p.replaceWith(...p.childNodes); });
    }
    html = d.innerHTML;
    td.addRule('dataImages', { filter: (n) => n.nodeName === 'IMG' && /^data:/.test(n.getAttribute('src') || ''), replacement: (_c, n) => `![${n.getAttribute('alt') || 'picture'}]()` });
    return td.turndown(html).replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }
  function htmlText(html) {
    const d = document.createElement('div');
    d.innerHTML = html;
    d.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,tr,div,br').forEach((e) => e.append('\n'));
    return d.textContent.replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }
  function standaloneHtml(title, body) {
    return `<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<title>${escapeHtml(title)}</title>\n<style>body{font:16px/1.6 "Segoe UI",system-ui,sans-serif;max-width:820px;margin:40px auto;padding:0 20px;color:#1f2328}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:4px 8px}img{max-width:100%}pre{background:#f4f4f4;padding:10px;overflow:auto}code{background:#f4f4f4;padding:1px 4px}</style>\n</head>\n<body>\n${body}\n</body>\n</html>\n`;
  }

  // RTF from HTML: drawn in a hidden box so computed styles are available to the RTF writer.
  async function rtfFromHtml(html) {
    const box = document.createElement('article');
    box.className = 'paper convert-scratch';
    box.style.cssText = 'position:fixed;left:-10000px;top:0;width:700px;color:#000;background:#fff;font:11pt Calibri,sans-serif';
    box.innerHTML = html;
    document.body.append(box);
    try { return await RtfEditor.writeRtf(box, '\\paperw12240\\paperh15840\\margl1440\\margr1440\\margt1440\\margb1440'); } finally { box.remove(); }
  }

  // ---------------------------------------------------------------- PDF

  const PAGE_CSS = (size, margin) => `<style>@page{size:${size};margin:${margin}}</style>`;
  async function toPdf(tab) {
    const k = tab.kind;
    if (k === 'pdf') return nativeBytes(tab);
    if (k === 'image') return fromImage(tab, 'pdf');
    if (k === 'docx') {
      // The paginated Word layout, one sheet of paper per page.
      const box = document.createElement('div');
      box.style.cssText = 'position:fixed;left:-20000px;top:0;';
      document.body.append(box);
      try {
        await Viewers.docx({ urls: [] }, await currentDocxBytes(tab), box);
        const view = box.querySelector('.docx-view');
        if (view) Viewers.paginate(view);
        const sec = box.querySelector('section.docx');
        const w = sec ? sec.style.width : '8.5in', hgt = sec ? sec.style.minHeight : '11in';
        const css = `${PAGE_CSS(`${w} ${hgt}`, '0')}<style>.docx-wrapper{background:none!important;padding:0!important;display:block!important}.docx-wrapper>section.docx{margin:0!important;box-shadow:none!important;page-break-after:always;break-after:page;overflow:hidden}</style>`;
        return desk.pdfFromHtml(css + box.innerHTML, { margins: { marginType: 'none' } });
      } finally { box.remove(); }
    }
    if (k === 'pptx') {
      const deck = tab.pptEditor ? tab.pptEditor.deck : await new PptxEditor.Deck().load(tab.data);
      const W = 960, H = W * deck.cy / deck.cx;
      const parts = [];
      for (const [i, s] of deck.slides.entries()) {
        parts.push(`<div class="pp-page">${await slideHtml(deck, s, i, W)}</div>`);
      }
      const css = `${PAGE_CSS(`${W}px ${H}px`, '0')}<style>.pp-page{width:${W}px;height:${H}px;overflow:hidden;page-break-after:always;break-after:page}.pp-page:last-child{page-break-after:auto}</style>`;
      return desk.pdfFromHtml(css + parts.join(''), { margins: { marginType: 'none' } });
    }
    if (k === 'sheet' || k === 'sheetview' || k === 'csv') {
      const wb = await workbookOf(tab);
      const css = `${PAGE_CSS('Letter landscape', '0.5in')}<style>table{border-collapse:collapse;font:9pt Calibri,sans-serif}td,th{border:1px solid #bbb;padding:2px 5px;white-space:nowrap}h2{font:600 12pt "Segoe UI",sans-serif;margin:0 0 8px}.sheet{page-break-after:always}</style>`;
      const body = wb.SheetNames.map((n) => `<div class="sheet"><h2>${escapeHtml(n)}</h2>${XLSX.utils.sheet_to_html(wb.Sheets[n], { header: '', footer: '' }).replace(/^[\s\S]*?<table/, '<table').replace(/<\/table>[\s\S]*$/, '</table>')}</div>`).join('');
      return desk.pdfFromHtml(css + body, {});
    }
    // Flowing text documents.
    let html = await htmlOf(tab);
    let page = 'Letter';
    let margin = '1in';
    if (k === 'odt' && tab.odtEditor) {
      const pl = tab.odtEditor.pageLayout();
      page = `${pl.width}px ${pl.height}px`;
      margin = `${pl.top}px ${pl.right}px ${pl.bottom}px ${pl.left}px`;
      const clone = tab.odtEditor.article.cloneNode(true);
      clone.removeAttribute('contenteditable');
      await inlineImages(clone);
      html = `<div class="odt-page" style="width:auto;min-height:0;padding:0;box-shadow:none;margin:0">${clone.outerHTML}</div>`;
    } else if (k === 'text') {
      html = `<pre style="white-space:pre-wrap;font:10pt Consolas,monospace;margin:0">${escapeHtml(textOf(tab))}</pre>`;
    } else if (k === 'rtf') {
      html = `<article class="paper rtf-view" style="box-shadow:none;padding:0;max-width:none">${html}</article>`;
    } else {
      html = `<article class="markdown-body" style="max-width:none;padding:0">${html}</article>`;
    }
    return desk.pdfFromHtml(PAGE_CSS(page, margin) + html, {});
  }

  // One slide as self-contained HTML: pictures inlined, chart canvases swapped for pictures of themselves.
  async function slideHtml(deck, s, i, W) {
    const el = await new PptxEditor.Renderer(deck, { slideNumber: i + 1 }).slide(s, W, false);
    const holder = document.createElement('div');
    holder.style.cssText = 'position:fixed;left:-20000px;top:0';
    holder.append(el);
    document.body.append(holder);
    try {
      await new Promise((r) => setTimeout(r, 60)); // charts draw once they are on the page
      for (const c of el.querySelectorAll('canvas')) {
        try { const img = document.createElement('img'); img.src = c.toDataURL('image/png'); img.style.cssText = c.style.cssText || 'width:100%;height:100%'; c.replaceWith(img); } catch { /* leave it */ }
      }
      await inlineImages(el);
      return el.outerHTML;
    } finally { holder.remove(); }
  }

  // ---------------------------------------------------------------- spreadsheets

  async function workbookOf(tab) {
    if (tab.kind === 'csv') return XLSX.read(textOf(tab), { type: 'string', raw: false });
    const bytes = tab.kind === 'sheet' ? await nativeBytes(tab) : tab.data;
    return XLSX.read(bytes, { type: 'array', cellStyles: true, cellNF: true, cellFormula: true, cellDates: true });
  }
  async function fromSheet(tab, ext) {
    const wb = await workbookOf(tab);
    const name = tab.sheetView ? tab.sheetView.sheet.name : wb.SheetNames[0];
    const ws = wb.Sheets[name] || wb.Sheets[wb.SheetNames[0]];
    switch (ext) {
      case 'csv': return enc('﻿' + XLSX.utils.sheet_to_csv(ws));
      case 'tsv': return enc('﻿' + XLSX.utils.sheet_to_csv(ws, { FS: '\t' }));
      case 'json': return enc(JSON.stringify(XLSX.utils.sheet_to_json(ws, { defval: null }), null, 2) + '\n');
      case 'html': return enc(standaloneHtml(stem(tab.name), wb.SheetNames.map((n) => `<h2>${escapeHtml(n)}</h2>` + XLSX.utils.sheet_to_html(wb.Sheets[n]).replace(/^[\s\S]*?<table/, '<table').replace(/<\/table>[\s\S]*$/, '</table>')).join('\n')));
      case 'md': {
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
        if (!rows.length) return enc('');
        const w = Math.max(...rows.map((r) => r.length));
        const cell = (v) => String(v ?? '').replace(/\|/g, '\\|').replace(/\n/g, ' ');
        const line = (r) => `| ${Array.from({ length: w }, (_, i) => cell(r[i])).join(' | ')} |`;
        return enc([line(rows[0]), `| ${Array(w).fill('---').join(' | ')} |`, ...rows.slice(1).map(line)].join('\n') + '\n');
      }
      case 'xlsx': case 'ods': case 'xls': case 'xlsb': case 'fods':
        if (ext !== 'xlsx') keepModernAsValues(wb);
        return fixOds(XLSX.write(wb, { bookType: ext === 'xls' ? 'biff8' : ext, type: 'array', cellStyles: true }), ext);
      default: throw new Error(`Spreadsheets can't be converted to .${ext}.`);
    }
  }

  // Functions newer than the target format (XLOOKUP, FILTER, LET...; stored with an _xlfn. prefix) can't be
  // written as formulas there, so those cells keep their calculated value instead.
  function keepModernAsValues(wb) {
    for (const n of wb.SheetNames) {
      const ws = wb.Sheets[n];
      for (const k of Object.keys(ws)) { const c = ws[k]; if (k[0] !== '!' && c && c.f && /_xl(fn|ws|pm)\.|#REF!/i.test(c.f)) { delete c.f; delete c.F; } }
    }
    return wb;
  }

  // SheetJS writes references to other sheets in OpenDocument formulas as Sheet![.A1:.B2]; the format wants
  // [Sheet.A1:.B2] ('Sheet name'![.A1] likewise). Rewrite them in content.xml (or the flat .fods file).
  async function fixOds(bytes, ext) {
    if (ext !== 'ods' && ext !== 'fods') return bytes;
    const fix = (xml) => xml.replace(/table:formula="[^"]*"/g, (f) => f.replace(/((?:&apos;|')(?:[^'&]|&(?!apos;))+(?:&apos;|')|[A-Za-z_][\w.]*)!\[\.([^\]]+)\]/g, '[$1.$2]'));
    if (ext === 'fods') { const t = typeof bytes === 'string' ? bytes : new TextDecoder().decode(bytes); return new TextEncoder().encode(fix(t)).buffer; }
    const zip = await JSZip.loadAsync(bytes);
    zip.file('content.xml', fix(await zip.file('content.xml').async('string')));
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' });
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }

  // ---------------------------------------------------------------- presentations

  async function fromSlides(tab, ext) {
    const deck = tab.pptEditor ? tab.pptEditor.deck : await new PptxEditor.Deck().load(tab.data);
    if (ext === 'png') {
      const W = 1920, H = Math.round(W * deck.cy / deck.cx);
      const pages = [];
      for (const [i, s] of deck.slides.entries()) pages.push(await slideHtml(deck, s, i, W));
      return { files: (await desk.pngFromHtml(pages, W, H)).map((data, i) => ({ suffix: `-${i + 1}`, data })) };
    }
    // Text outline: titles as headings, other text as bullets, plus speaker notes.
    const { NS } = PptxEditor;
    const lines = [];
    for (const [i, s] of deck.slides.entries()) {
      let title = '';
      const body = [];
      for (const sp of s.doc.getElementsByTagNameNS(NS.p, 'sp')) {
        const ph = deck.phOf(sp);
        const paras = [...sp.getElementsByTagNameNS(NS.a, 'p')].map((p) => ({ text: [...p.getElementsByTagNameNS(NS.a, 't')].map((t) => t.textContent).join(''), lvl: +((p.getElementsByTagNameNS(NS.a, 'pPr')[0] || { getAttribute: () => 0 }).getAttribute('lvl') || 0) })).filter((p) => p.text.trim());
        if (ph && /title/i.test(ph.type) && !title) title = paras.map((p) => p.text).join(' ');
        else body.push(...paras);
      }
      const notes = tab.pptEditor ? await tab.pptEditor.getNotes(i) : '';
      if (ext === 'md') {
        lines.push(`## ${title || `Slide ${i + 1}`}`, '', ...body.map((p) => `${'  '.repeat(p.lvl)}- ${p.text}`));
        if (notes) lines.push('', `> ${notes.replace(/\n/g, '\n> ')}`);
        lines.push('');
      } else {
        lines.push(`Slide ${i + 1}: ${title}`, ...body.map((p) => `${'  '.repeat(p.lvl + 1)}${p.text}`));
        if (notes) lines.push(`  Notes: ${notes}`);
        lines.push('');
      }
    }
    if (ext === 'md' || ext === 'txt') return enc(lines.join('\n'));
    throw new Error(`Presentations can't be converted to .${ext} on this PC.`);
  }

  // ---------------------------------------------------------------- PDF sources

  async function pdfDoc(tab) {
    const bytes = new Uint8Array(await nativeBytes(tab)).slice();
    const task = DDPdf.pdfjs.getDocument({ data: bytes, isEvalSupported: false });
    const doc = await task.promise;
    doc.close = () => task.destroy();
    return doc;
  }
  // Lines grouped into paragraphs by vertical gaps; big text becomes a heading.
  async function pdfParagraphs(tab) {
    const doc = await pdfDoc(tab);
    const pages = [];
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const tc = await page.getTextContent();
      const lines = [];
      for (const it of tc.items) {
        if (!it.str) continue;
        const y = it.transform[5], size = Math.abs(it.transform[3]) || 10;
        const last = lines[lines.length - 1];
        if (last && Math.abs(last.y - y) < size * 0.4) { last.text += (last.text.endsWith(' ') || it.str.startsWith(' ') ? '' : ' ') + it.str; last.size = Math.max(last.size, size); }
        else lines.push({ y, size, text: it.str });
        if (it.hasEOL && lines.length) lines[lines.length - 1].eol = true;
      }
      const sizes = lines.map((l) => l.size).sort((a, b) => a - b);
      const body = sizes[Math.floor(sizes.length / 2)] || 10;
      const paras = [];
      let cur = null;
      lines.forEach((l, i) => {
        const gap = i ? lines[i - 1].y - l.y : 0;
        const heading = l.size > body * 1.25;
        if (!cur || heading || cur.heading || gap > Math.max(l.size, lines[i - 1].size) * 1.6) { cur = { text: l.text.trim(), heading }; paras.push(cur); }
        else cur.text += (cur.text.endsWith('-') ? '' : ' ') + l.text.trim();
      });
      pages.push(paras.filter((p) => p.text));
    }
    await doc.close();
    return pages;
  }
  async function pdfPages(tab) {
    const doc = await pdfDoc(tab);
    const files = [];
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const vp = page.getViewport({ scale: 2 });
      const c = document.createElement('canvas');
      c.width = vp.width; c.height = vp.height;
      await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
      files.push({ suffix: `-${n}`, data: await (await new Promise((r) => c.toBlob(r, 'image/png'))).arrayBuffer() });
    }
    await doc.close();
    return { files };
  }

  // ---------------------------------------------------------------- pictures

  // A picture (with the image editor's changes when it is being edited) as another picture type or a PDF page.
  async function fromImage(tab, ext) {
    {
      const c = tab.imageEditor ? tab.imageEditor.renderFull() : await ImageEditor.decode(tab.data, tab.ext);
      const bytes = await ImageEditor.encode(c, ext === 'pdf' ? 'png' : ext);
      if (ext !== 'pdf') return bytes;
      const { PDFDocument } = DDPdf.PDFLib;
      const pdf = await PDFDocument.create();
      pdf.setCreator('Doc Studio'); pdf.setProducer('Doc Studio');
      const png = await pdf.embedPng(bytes);
      // Fit on a Letter page (portrait or landscape to match the picture), with half-inch margins.
      const land = c.width > c.height;
      const W = land ? 792 : 612, H = land ? 612 : 792;
      const s = Math.min((W - 72) / c.width, (H - 72) / c.height, 1);
      const p = pdf.addPage([W, H]);
      p.drawImage(png, { x: (W - c.width * s) / 2, y: (H - c.height * s) / 2, width: c.width * s, height: c.height * s });
      return (await pdf.save()).buffer;
    }
  }

  // ---------------------------------------------------------------- EPUB

  async function epub(title, html) {
    const zip = new JSZip();
    zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });
    zip.file('META-INF/container.xml', '<?xml version="1.0" encoding="UTF-8"?>\n<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>');
    // Chapters at each top-level heading.
    const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
    const images = [];
    doc.querySelectorAll('img').forEach((img, i) => {
      const m = /^data:image\/(png|jpe?g|gif|svg\+xml);base64,(.*)$/.exec(img.getAttribute('src') || '');
      if (!m) { img.remove(); return; }
      const ext = m[1] === 'svg+xml' ? 'svg' : m[1].replace('jpeg', 'jpg');
      const name = `images/img${i + 1}.${ext}`;
      zip.file(`OEBPS/${name}`, m[2], { base64: true });
      images.push({ name, type: `image/${m[1]}` });
      img.setAttribute('src', name);
      if (!img.getAttribute('alt')) img.setAttribute('alt', '');
    });
    const chapters = [];
    let cur = null;
    for (const n of [...doc.body.childNodes]) {
      if (n.nodeType === 1 && /^H[12]$/.test(n.tagName) || !cur) { cur = { title: n.nodeType === 1 && /^H[12]$/.test(n.tagName) ? n.textContent.trim() : title, nodes: [] }; chapters.push(cur); }
      cur.nodes.push(n);
    }
    const ser = new XMLSerializer();
    const xhtml = (t, body) => `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" lang="en"><head><meta charset="utf-8"/><title>${escapeHtml(t)}</title></head><body>${body}</body></html>`;
    chapters.forEach((c, i) => zip.file(`OEBPS/ch${i + 1}.xhtml`, xhtml(c.title, c.nodes.map((n) => (n.nodeType === 1 ? ser.serializeToString(n).replace(/ xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '') : escapeHtml(n.textContent))).join(''))));
    zip.file('OEBPS/nav.xhtml', xhtml(title, `<nav epub:type="toc"><h1>${escapeHtml(title)}</h1><ol>${chapters.map((c, i) => `<li><a href="ch${i + 1}.xhtml">${escapeHtml(c.title || `Chapter ${i + 1}`)}</a></li>`).join('')}</ol></nav>`));
    const id = `urn:uuid:${crypto.randomUUID()}`;
    zip.file('OEBPS/content.opf', `<?xml version="1.0" encoding="UTF-8"?>\n<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="uid"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:identifier id="uid">${id}</dc:identifier><dc:title>${escapeHtml(title)}</dc:title><dc:language>en</dc:language><meta property="dcterms:modified">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</meta></metadata><manifest><item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>${chapters.map((_, i) => `<item id="ch${i + 1}" href="ch${i + 1}.xhtml" media-type="application/xhtml+xml"/>`).join('')}${images.map((im, i) => `<item id="img${i + 1}" href="${im.name}" media-type="${im.type}"/>`).join('')}</manifest><spine>${chapters.map((_, i) => `<itemref idref="ch${i + 1}"/>`).join('')}</spine></package>`);
    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', mimeType: 'application/epub+zip' });
  }

  // ---------------------------------------------------------------- UI

  // Convert the open document: ask where to save, write, then offer to open the result.
  async function convertTab(tab, ext) {
    if (!tab) return;
    const base = tab.path ? tab.path.replace(/\.[^.\\/]+$/, '') : stem(tab.name);
    flash(`Converting to .${ext}…`);
    let out;
    try { out = await run(tab, ext); } catch (err) { logError(err); desk.error(`Couldn't convert ${tab.name} to .${ext}`, err.message); return; }
    if (out && out.files) {
      // Several pictures: name-1.png, name-2.png ... next to the chosen name.
      const target = await desk.saveDialog(`${base}.${ext}`, ext);
      if (!target) return;
      const stemPath = target.replace(/\.[^.\\/]+$/, '');
      for (const f of out.files) await desk.write(`${stemPath}${f.suffix}.${ext}`, f.data, { version: false });
      flash(`Saved ${out.files.length} picture(s)`);
      desk.reveal(`${stemPath}${out.files[0].suffix}.${ext}`);
      return;
    }
    const target = await desk.saveDialog(`${base}.${ext}`, ext);
    if (!target) return;
    await desk.write(target, out, { version: false });
    const r = await desk.confirm({ type: 'info', message: `Converted to ${target.split(/[\\/]/).pop()}`, buttons: ['Open it', 'Show in folder', 'Close'], defaultId: 0, cancelId: 2 });
    if (r === 0) openPath(target);
    if (r === 1) desk.reveal(target);
  }

  // Convert to… and Convert files… open the Convert tool window (tools-panel.js): the open document, or no
  // files yet.
  function dialog() { return Tools.convert(); }
  function batch() { return Tools.convert({ empty: true }); }

  return { targets, run, convertTab, dialog, batch, toMarkdown, keepModernAsValues, fixOds, LABELS };
})();
