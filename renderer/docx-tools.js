'use strict';
// Word editor commands added for Google Docs parity: footnotes, comments, bookmarks, page numbers, headers and
// footers, table of contents, horizontal lines, line and paragraph spacing, page setup (size, orientation,
// margins, columns, page colour) and table rows/columns. Each is a structural edit: applied to the XML, then the
// document is redrawn (and the edit can be undone as one step). Inline items go in at the caret: a private marker
// character is typed there first and replaced by the real XML element when the document is rebuilt.

(() => {
  const E = DocxEditor.Editor.prototype;
  const { W, R } = DocxEditor;
  const REL = (t) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${t}`;
  const CT = (t) => `application/vnd.openxmlformats-officedocument.wordprocessingml.${t}+xml`;
  const NSDECL = `xmlns:w="${W}" xmlns:r="${R}"`;
  const SETTINGS_TAIL = ['hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType', 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'];
  const SETTINGS_EARLY_AFTER = ['writeProtection', 'view', 'zoom', 'removePersonalInformation', 'removeDateAndTime', 'doNotDisplayPageBoundaries'];

  // ---------------------------------------------------------------- helpers

  const fieldRuns = (ctx, d, instr, result, rPr) => {
    const r = (child) => { const x = ctx.wEl(d, 'r'); if (rPr) x.append(rPr.cloneNode(true)); x.append(child); return x; };
    const t = ctx.wEl(d, 'instrText'); t.setAttribute('xml:space', 'preserve'); t.textContent = ` ${instr} `;
    const res = ctx.wEl(d, 't'); res.setAttribute('xml:space', 'preserve'); res.textContent = result;
    return [r(ctx.wEl(d, 'fldChar', { fldCharType: 'begin' })), r(t), r(ctx.wEl(d, 'fldChar', { fldCharType: 'separate' })), r(res), r(ctx.wEl(d, 'fldChar', { fldCharType: 'end' }))];
  };
  async function styleIds(ctx) {
    const d = await ctx.partDoc(`${ctx.docDir}styles.xml`);
    return new Set(d ? [...d.getElementsByTagNameNS(W, 'style')].map((s) => s.getAttributeNS(W, 'styleId')) : []);
  }
  // Word's built-in styles for notes and comments, added when a document doesn't have them yet.
  const BUILTIN = {
    FootnoteReference: '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>',
    FootnoteText: '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>',
    CommentReference: '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>',
  };
  async function ensureStyles(ctx, ids) {
    const d = await ctx.partDoc(`${ctx.docDir}styles.xml`);
    if (!d) return new Set();
    const have = new Set([...d.getElementsByTagNameNS(W, 'style')].map((x) => x.getAttributeNS(W, 'styleId')));
    for (const id of ids) {
      if (have.has(id) || !BUILTIN[id]) continue;
      let xml = BUILTIN[id];
      if (!have.has('Normal')) xml = xml.replace('<w:basedOn w:val="Normal"/>', '');
      if (!have.has('DefaultParagraphFont')) xml = xml.replace('<w:basedOn w:val="DefaultParagraphFont"/>', '');
      d.documentElement.append(imp(d, xml.replace('<w:style ', `<w:style ${NSDECL} `)));
      have.add(id);
    }
    return have;
  }
  async function settingsDoc(ctx) {
    const path = ctx.relTarget(REL('settings')) || `${ctx.docDir}settings.xml`;
    let d = await ctx.partDoc(path);
    if (!d) d = (await ctx.createPart('settings.xml', REL('settings'), CT('settings'), `<w:settings ${NSDECL}/>`)).doc;
    return d;
  }
  function settingsSet(ctx, d, el, early) {
    const root = d.documentElement;
    const old = ctx.kids(root).find((c) => c.localName === el.localName && c.namespaceURI === W);
    if (old) { root.replaceChild(el, old); return; }
    if (early) {
      const after = [...ctx.kids(root)].reverse().find((c) => SETTINGS_EARLY_AFTER.includes(c.localName));
      root.insertBefore(el, after ? after.nextSibling : root.firstChild);
    } else {
      const before = ctx.kids(root).find((c) => SETTINGS_TAIL.includes(c.localName));
      root.insertBefore(el, before || null);
    }
  }
  const imp = (d, xml) => { const n = d.importNode(DocxEditor.parseXml(xml).documentElement, true); for (const a of [...n.attributes]) if (a.name.startsWith('xmlns') && d.documentElement.lookupNamespaceURI(a.localName === 'xmlns' ? null : a.localName) === a.value) n.removeAttributeNode(a); return n; };

  // Type a marker at the caret (or around the selection) and turn it into XML in the structural edit.
  E.atCaret = async function (make, opts) {
    opts = opts || {};
    const sel = getSelection();
    if (!sel.rangeCount) throw new Error('Click in the document where it should go first.');
    const range = sel.getRangeAt(0).cloneRange();
    const host = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
    const area = host && host.closest('[contenteditable="true"]');
    if (!area || !this.stage.contains(area)) throw new Error('Click in the document where it should go first.');
    if (opts.bodyOnly && area.tagName !== 'ARTICLE') throw new Error('This can only be added to the main text, not headers, footers or notes.');
    this.markerSeq = ((this.markerSeq || 0) + 1) % 120;
    const start = String.fromCharCode(0xe100 + this.markerSeq * 2), end = String.fromCharCode(0xe101 + this.markerSeq * 2);
    const ranged = opts.ranged && !range.collapsed;
    return this.structural(async (ctx) => {
      if (opts.before) await opts.before(ctx);
      if (ranged) {
        if (!ctx.atMarker(end, (d, rPr) => make(ctx, d, rPr, 'end'))) throw new Error('The insertion point was lost.');
        if (!ctx.atMarker(start, (d, rPr) => make(ctx, d, rPr, 'start'))) throw new Error('The insertion point was lost.');
      } else if (!ctx.atMarker(start, (d, rPr) => make(ctx, d, rPr, 'point'))) throw new Error('The insertion point was lost.');
    }, () => {
      area.focus();
      const put = (at, ch) => { const r = document.createRange(); r.setStart(at.container, at.offset); r.collapse(true); sel.removeAllRanges(); sel.addRange(r); document.execCommand('insertText', false, ch); };
      if (ranged) {
        const s0 = { container: range.startContainer, offset: range.startOffset };
        put({ container: range.endContainer, offset: range.endOffset }, end);
        put(s0, start);
      } else {
        sel.removeAllRanges();
        const r = range.cloneRange();
        r.collapse(!opts.atEnd);
        sel.addRange(r);
        document.execCommand('insertText', false, start);
      }
    }, { anyPart: !opts.bodyOnly });
  };

  // ---------------------------------------------------------------- footnotes

  E.insertFootnote = function (text) {
    let id;
    return this.atCaret((ctx, d, rPr) => {
      const r = ctx.wEl(d, 'r');
      const pr = ctx.wEl(d, 'rPr');
      pr.append(ctx.wEl(d, 'rStyle', { val: 'FootnoteReference' }));
      r.append(pr, ctx.wEl(d, 'footnoteReference', { id }));
      void rPr;
      return [r];
    }, {
      bodyOnly: true,
      before: async (ctx) => {
        ctx.noteStyles = await ensureStyles(ctx, ['FootnoteReference', 'FootnoteText']);
        let path = ctx.relTarget(REL('footnotes'));
        let notes = path ? await ctx.partDoc(path) : null;
        if (!notes) {
          notes = (await ctx.createPart('footnotes.xml', REL('footnotes'), CT('footnotes'), `<w:footnotes ${NSDECL}><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>`)).doc;
          const st = await settingsDoc(ctx);
          settingsSet(ctx, st, imp(st, `<w:footnotePr ${NSDECL}><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`));
        }
        id = Math.max(0, ...[...notes.getElementsByTagNameNS(W, 'footnote')].map((n) => +n.getAttributeNS(W, 'id'))) + 1;
        const s = ctx.noteStyles;
        const esc = (x) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;');
        notes.documentElement.append(imp(notes, `<w:footnote ${NSDECL} w:id="${id}"><w:p><w:pPr>${s.has('FootnoteText') ? '<w:pStyle w:val="FootnoteText"/>' : '<w:spacing w:after="0"/>'}</w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${esc(text || '')}</w:t></w:r></w:p></w:footnote>`));
      },
    });
  };

  // ---------------------------------------------------------------- comments

  E.insertComment = function (text, author) {
    let id;
    const esc = (x) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    return this.atCaret((ctx, d, rPr, which) => {
      if (which === 'start') return [ctx.wEl(d, 'commentRangeStart', { id })];
      const ref = ctx.wEl(d, 'r');
      const pr = ctx.wEl(d, 'rPr');
      if (ctx.noteStyles.has('CommentReference')) pr.append(ctx.wEl(d, 'rStyle', { val: 'CommentReference' }));
      ref.append(pr, ctx.wEl(d, 'commentReference', { id }));
      return which === 'end' ? [ctx.wEl(d, 'commentRangeEnd', { id }), ref] : [ctx.wEl(d, 'commentRangeStart', { id }), ctx.wEl(d, 'commentRangeEnd', { id }), ref];
    }, {
      ranged: true,
      bodyOnly: true,
      before: async (ctx) => {
        ctx.noteStyles = await ensureStyles(ctx, ['CommentReference']);
        let path = ctx.relTarget(REL('comments'));
        let c = path ? await ctx.partDoc(path) : null;
        if (!c) c = (await ctx.createPart('comments.xml', REL('comments'), CT('comments'), `<w:comments ${NSDECL}/>`)).doc;
        id = Math.max(-1, ...[...c.getElementsByTagNameNS(W, 'comment')].map((n) => +n.getAttributeNS(W, 'id'))) + 1;
        const when = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
        const paras = String(text || '').split('\n').map((line) => `<w:p><w:r><w:t xml:space="preserve">${esc(line)}</w:t></w:r></w:p>`).join('');
        c.documentElement.append(imp(c, `<w:comment ${NSDECL} w:id="${id}" w:author="${esc(author || 'Doc Studio')}" w:date="${when}" w:initials="${esc((author || 'DD').slice(0, 2).toUpperCase())}">${paras}</w:comment>`));
      },
    });
  };
  // Comments already in the document: { id, author, date, text }.
  E.comments = async function () {
    const rels = this.zip.file(this.relsPath) ? DocxEditor.parseXml(await this.zip.file(this.relsPath).async('string')) : null;
    const rel = rels && [...rels.documentElement.children].find((r) => r.getAttribute('Type') === REL('comments'));
    if (!rel) return [];
    const path = this.docPath.replace(/[^/]+$/, '') + rel.getAttribute('Target').replace(/^\.?\//, '');
    const f = this.zip.file(path);
    if (!f) return [];
    const d = DocxEditor.parseXml(await f.async('string'));
    return [...d.getElementsByTagNameNS(W, 'comment')].map((c) => ({ id: c.getAttributeNS(W, 'id'), author: c.getAttributeNS(W, 'author'), date: c.getAttributeNS(W, 'date'), text: [...c.getElementsByTagNameNS(W, 'p')].map((p) => p.textContent).join('\n') }));
  };
  // Delete a comment: its text and its anchors.
  E.deleteComment = function (id) {
    return this.structural(async (ctx) => {
      const path = ctx.relTarget(REL('comments'));
      const c = path && await ctx.partDoc(path);
      if (c) for (const n of [...c.getElementsByTagNameNS(W, 'comment')]) if (n.getAttributeNS(W, 'id') === String(id)) n.remove();
      for (const name of ['commentRangeStart', 'commentRangeEnd', 'commentReference']) {
        for (const n of [...ctx.doc.getElementsByTagNameNS(W, name)]) {
          if (n.getAttributeNS(W, 'id') !== String(id)) continue;
          const run = name === 'commentReference' ? n.parentNode : null;
          if (run && ctx.kids(run).every((k) => ctx.isW(k, 'rPr') || k === n)) run.remove(); else n.remove();
        }
      }
    }, null, { anyPart: true });
  };

  // ---------------------------------------------------------------- bookmarks

  E.insertBookmark = function (name) {
    const clean = String(name || 'Bookmark').replace(/[^\w]/g, '_').replace(/^(\d)/, '_$1').slice(0, 40);
    let id;
    return this.atCaret((ctx, d) => [ctx.wEl(d, 'bookmarkStart', { id, name: clean }), ctx.wEl(d, 'bookmarkEnd', { id })], {
      bodyOnly: true,
      before: (ctx) => { id = 1 + Math.max(0, ...[...ctx.doc.getElementsByTagNameNS(W, 'bookmarkStart')].map((b) => +b.getAttributeNS(W, 'id') || 0)); },
    });
  };
  E.bookmarks = function () { return [...this.doc.getElementsByTagNameNS(W, 'bookmarkStart')].map((b) => b.getAttributeNS(W, 'name')).filter((n) => n && !n.startsWith('_')); };

  // ---------------------------------------------------------------- fields at the caret: page number, page count, date

  E.insertField = function (instr, result) {
    return this.atCaret((ctx, d, rPr) => fieldRuns(ctx, d, instr, result, rPr), {});
  };

  // ---------------------------------------------------------------- headers, footers and page numbers

  // Make sure every section has a default header/footer of `kind`; returns the documents to write into.
  async function ensureHeaderFooter(ctx, kind) {
    const tag = kind === 'header' ? 'hdr' : 'ftr';
    const refName = `${kind}Reference`;
    const targets = [];
    let created = null;
    for (const sp of ctx.sectPrs()) {
      const ref = ctx.kids(sp).find((c) => ctx.isW(c, refName) && (c.getAttributeNS(W, 'type') || 'default') === 'default');
      if (ref) {
        const target = ctx.relById(ref.getAttributeNS(R, 'id'));
        const d = target && await ctx.partDoc(target);
        if (d && !targets.includes(d)) targets.push(d);
        continue;
      }
      if (!created) {
        let n = 1;
        while (ctx.hasPart(`${ctx.docDir}${kind}${n}.xml`)) n++;
        created = await ctx.createPart(`${kind}${n}.xml`, REL(kind), CT(kind), `<w:${tag} ${NSDECL}><w:p><w:pPr><w:pStyle w:val="${kind === 'header' ? 'Header' : 'Footer'}"/></w:pPr></w:p></w:${tag}>`);
        if (!ctx.styleSet.has(kind === 'header' ? 'Header' : 'Footer')) for (const ps of created.doc.getElementsByTagNameNS(W, 'pStyle')) ps.parentNode.removeChild(ps);
        targets.push(created.doc);
      }
      const el = ctx.wEl(ctx.doc, refName, { type: 'default' });
      el.setAttributeNS(R, 'r:id', created.rid);
      // References come first in sectPr, headers before footers.
      const refs = ctx.kids(sp).filter((c) => ctx.isW(c, 'headerReference') || ctx.isW(c, 'footerReference'));
      const after = kind === 'header' ? refs.filter((c) => ctx.isW(c, 'headerReference')).pop() : refs[refs.length - 1];
      sp.insertBefore(el, after ? after.nextSibling : sp.firstChild);
    }
    return targets;
  }
  function hfContext(ctx) {
    ctx.relById = (rid) => {
      const r = [...ctx.bodyRelsDoc().documentElement.children].find((x) => x.getAttribute('Id') === rid);
      return r ? ctx.docDir + r.getAttribute('Target').replace(/^\.?\//, '') : null;
    };
    ctx.hasPart = (path) => ctx.knownPaths().includes(path);
  }

  E.addHeaderFooter = function (kind) {
    return this.structural(async (ctx) => {
      hfContext(ctx);
      ctx.styleSet = await styleIds(ctx);
      await ensureHeaderFooter(ctx, kind);
    }, null, { anyPart: true });
  };

  // Page numbers: { where: 'header'|'footer', align: 'left'|'center'|'right', withTotal: bool }
  E.insertPageNumbers = function (o) {
    return this.structural(async (ctx) => {
      hfContext(ctx);
      ctx.styleSet = await styleIds(ctx);
      const docs = await ensureHeaderFooter(ctx, o.where);
      for (const d of docs) {
        const root = d.documentElement;
        const p = ctx.wEl(d, 'p');
        const pPr = ctx.wEl(d, 'pPr');
        const sty = o.where === 'header' ? 'Header' : 'Footer';
        if (ctx.styleSet.has(sty)) pPr.append(ctx.wEl(d, 'pStyle', { val: sty }));
        pPr.append(ctx.wEl(d, 'jc', { val: o.align === 'left' ? 'left' : o.align === 'center' ? 'center' : 'right' }));
        p.append(pPr);
        if (o.withTotal) {
          const t = (x) => { const r = ctx.wEl(d, 'r'); const tt = ctx.wEl(d, 't'); tt.setAttribute('xml:space', 'preserve'); tt.textContent = x; r.append(tt); return r; };
          p.append(t('Page '), ...fieldRuns(ctx, d, 'PAGE', '1', null), t(' of '), ...fieldRuns(ctx, d, 'NUMPAGES', '1', null));
        } else p.append(...fieldRuns(ctx, d, 'PAGE', '1', null));
        // Replace an empty last paragraph, otherwise add below what is there.
        const last = [...ctx.kids(root)].reverse().find((c) => ctx.isW(c, 'p'));
        if (last && !last.textContent.trim() && !last.getElementsByTagNameNS(W, 'drawing').length) root.replaceChild(p, last); else root.append(p);
      }
    }, null, { anyPart: true });
  };

  // ---------------------------------------------------------------- table of contents

  E.insertTableOfContents = function () {
    return this.structural(async (ctx) => {
      const ids = await styleIds(ctx);
      const levelOf = (p) => {
        const ps = p.getElementsByTagNameNS(W, 'pStyle')[0];
        const id = ps && ps.getAttributeNS(W, 'val');
        const info = id && ctx.styles.byId[id];
        const m = /^heading\s*(\d)$/i.exec((info && info.name) || id || '');
        if (m) return +m[1];
        const ol = p.getElementsByTagNameNS(W, 'outlineLvl')[0];
        return ol ? +ol.getAttributeNS(W, 'val') + 1 : 0;
      };
      const heads = [...ctx.doc.getElementsByTagNameNS(W, 'body')[0].getElementsByTagNameNS(W, 'p')].map((p) => ({ lvl: levelOf(p), text: p.textContent.replace(/[-]/g, '').trim() })).filter((x) => x.lvl >= 1 && x.lvl <= 3 && x.text);
      const d = ctx.doc;
      const para = (lvl) => {
        const p = ctx.wEl(d, 'p');
        const pPr = ctx.wEl(d, 'pPr');
        if (ids.has(`TOC${lvl}`)) pPr.append(ctx.wEl(d, 'pStyle', { val: `TOC${lvl}` }));
        else pPr.append(ctx.wEl(d, 'ind', { left: 360 * (lvl - 1) }));
        p.append(pPr);
        return p;
      };
      const text = (d0, s) => { const r = ctx.wEl(d0, 'r'); const t = ctx.wEl(d0, 't'); t.setAttribute('xml:space', 'preserve'); t.textContent = s; r.append(t); return r; };
      const run = (el) => { const r = ctx.wEl(d, 'r'); r.append(el); return r; };
      const instr = ctx.wEl(d, 'instrText'); instr.setAttribute('xml:space', 'preserve'); instr.textContent = ' TOC \\o "1-3" \\h \\z \\u ';
      const entries = heads.length ? heads : [{ lvl: 1, text: 'No headings yet: use Heading 1-3 styles, then update this table (in Microsoft Word: right-click > Update Field).' }];
      const ps = entries.map((h) => { const p = para(h.lvl); p.append(text(d, h.text)); return p; });
      ps[0].insertBefore(run(ctx.wEl(d, 'fldChar', { fldCharType: 'begin' })), ps[0].children[1]);
      ps[0].insertBefore(run(instr), ps[0].children[2]);
      ps[0].insertBefore(run(ctx.wEl(d, 'fldChar', { fldCharType: 'separate' })), ps[0].children[3]);
      ps[ps.length - 1].append(run(ctx.wEl(d, 'fldChar', { fldCharType: 'end' })));
      let at = ctx.caretNode;
      for (const p of ps) { ctx.insertAfterBlock(at, p, d); at = p; }
      // Word fills in page numbers and links when it next opens the file.
      const st = await settingsDoc(ctx);
      settingsSet(ctx, st, ctx.wEl(st, 'updateFields', { val: 'true' }));
    });
  };

  // ---------------------------------------------------------------- horizontal line

  E.insertHorizontalLine = function () {
    return this.structural((ctx) => {
      const d = ctx.doc;
      const p = ctx.wEl(d, 'p');
      p.append(imp(d, `<w:pPr ${NSDECL}><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr>`));
      ctx.insertAfterBlock(ctx.caretNode, p, d);
    });
  };

  // ---------------------------------------------------------------- line and paragraph spacing

  // spacing: { line (multiple of single), before, after (twips), keepNext }
  E.setSpacing = function (ps, spacing) {
    for (const p of ps) {
      const cur = p.dataset.ddSpacing ? JSON.parse(p.dataset.ddSpacing) : {};
      const next = { ...cur, ...spacing };
      p.dataset.ddSpacing = JSON.stringify(next);
      if (next.line != null) p.style.lineHeight = String(1.15 * next.line);
      if (next.before != null) p.style.marginTop = `${next.before / 15}px`;
      if (next.after != null) p.style.marginBottom = `${next.after / 15}px`;
    }
    this.drawPageMarkers();
    if (this.opts.onChange) this.opts.onChange();
  };
  E.spacingOf = function (p) {
    if (p.dataset.ddSpacing) return JSON.parse(p.dataset.ddSpacing);
    const cs = getComputedStyle(p);
    const lh = parseFloat(cs.lineHeight), fs = parseFloat(cs.fontSize);
    return { line: lh && fs ? Math.round(lh / fs / 1.15 * 100) / 100 : 1, before: Math.round(parseFloat(cs.marginTop) * 15), after: Math.round(parseFloat(cs.marginBottom) * 15) };
  };

  // ---------------------------------------------------------------- page setup

  // { width, height (twips), orientation: 'portrait'|'landscape', margins: { top, right, bottom, left } (twips), color: '#rrggbb'|null, columns }
  E.pageSetupInfo = function () {
    const sp = [...this.doc.getElementsByTagNameNS(W, 'sectPr')].pop();
    const sz = sp && DocxEditor.child(sp, 'pgSz');
    const mar = sp && DocxEditor.child(sp, 'pgMar');
    const cols = sp && DocxEditor.child(sp, 'cols');
    const bg = this.doc.getElementsByTagNameNS(W, 'background')[0];
    const n = (el, a, d) => (el && el.getAttributeNS(W, a) ? +el.getAttributeNS(W, a) : d);
    const w = n(sz, 'w', 12240), h = n(sz, 'h', 15840);
    return {
      width: w, height: h, orientation: (sz && sz.getAttributeNS(W, 'orient')) === 'landscape' || w > h ? 'landscape' : 'portrait',
      margins: { top: n(mar, 'top', 1440), right: n(mar, 'right', 1440), bottom: n(mar, 'bottom', 1440), left: n(mar, 'left', 1440) },
      color: bg && bg.getAttributeNS(W, 'color') && bg.getAttributeNS(W, 'color') !== 'auto' ? `#${bg.getAttributeNS(W, 'color')}` : null,
      columns: n(cols, 'num', 1),
    };
  };
  E.pageSetup = function (o) {
    return this.structural(async (ctx) => {
      for (const sp of ctx.sectPrs()) {
        let sz = ctx.child(sp, 'pgSz');
        if (!sz) { sz = ctx.wEl(ctx.doc, 'pgSz'); sp.insertBefore(sz, ctx.kids(sp).find((c) => !/Reference$|^footnotePr$|^endnotePr$|^type$/.test(c.localName)) || null); }
        let w = o.width, h = o.height;
        if (o.orientation === 'landscape' && w < h) [w, h] = [h, w];
        if (o.orientation === 'portrait' && w > h) [w, h] = [h, w];
        sz.setAttributeNS(W, 'w:w', String(Math.round(w)));
        sz.setAttributeNS(W, 'w:h', String(Math.round(h)));
        if (o.orientation === 'landscape') sz.setAttributeNS(W, 'w:orient', 'landscape'); else sz.removeAttributeNS(W, 'orient');
        let mar = ctx.child(sp, 'pgMar');
        if (!mar) { mar = ctx.wEl(ctx.doc, 'pgMar', { header: 720, footer: 720, gutter: 0 }); sz.after(mar); }
        for (const k of ['top', 'right', 'bottom', 'left']) mar.setAttributeNS(W, `w:${k}`, String(Math.round(o.margins[k])));
        if (o.columns) {
          let cols = ctx.child(sp, 'cols');
          if (!cols) { cols = ctx.wEl(ctx.doc, 'cols'); const after = ctx.kids(sp).find((c) => ['formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'].includes(c.localName)); sp.insertBefore(cols, after || null); }
          cols.setAttributeNS(W, 'w:num', String(o.columns));
          cols.setAttributeNS(W, 'w:space', '720');
          for (const c of ctx.kids(cols)) cols.removeChild(c);
          cols.removeAttributeNS(W, 'equalWidth');
        }
      }
      // Page colour: <w:background> before the body, and Word's "show background" setting.
      const root = ctx.doc.documentElement;
      const old = ctx.child(root, 'background');
      if (old) root.removeChild(old);
      const st = await settingsDoc(ctx);
      if (o.color && o.color.toLowerCase() !== '#ffffff') {
        root.insertBefore(ctx.wEl(ctx.doc, 'background', { color: o.color.replace('#', '').toUpperCase() }), ctx.child(root, 'body'));
        settingsSet(ctx, st, ctx.wEl(st, 'displayBackgroundShape'), true);
      } else {
        const ds = ctx.kids(st.documentElement).find((c) => c.localName === 'displayBackgroundShape');
        if (ds) ds.remove();
      }
    }, null, { anyPart: true });
  };

  // ---------------------------------------------------------------- tracked changes: review, accept, reject

  const CHANGE_TAGS = ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange'];
  // The tracked changes in the document as it is now (suggestions typed this session included).
  E.trackedChanges = async function () {
    const built = await this.build({ keepCaret: true });
    const zip = await JSZip.loadAsync(built.bytes);
    const d = DocxEditor.parseXml(await zip.file(this.docPath).async('string'));
    const out = [];
    for (const name of CHANGE_TAGS) {
      for (const el of d.getElementsByTagNameNS(W, name)) {
        // Run-level ins/del only (paragraph-mark changes inside rPr are accepted with their paragraph).
        if ((name === 'ins' || name === 'del') && el.parentNode && el.parentNode.localName === 'rPr') continue;
        const text = name === 'del' || name === 'moveFrom' ? [...el.getElementsByTagNameNS(W, 'delText'), ...el.getElementsByTagNameNS(W, 't')].map((t) => t.textContent).join('') : name.endsWith('Change') ? (el.parentNode.parentNode ? el.parentNode.parentNode.textContent : '') : el.textContent;
        out.push({ id: el.getAttributeNS(W, 'id'), type: name, author: el.getAttributeNS(W, 'author'), date: el.getAttributeNS(W, 'date'), text: text.slice(0, 200) });
      }
    }
    return out;
  };
  function resolveChange(ctx, el, accept) {
    const name = el.localName;
    const unwrap = () => { while (el.firstChild) el.parentNode.insertBefore(el.firstChild, el); el.remove(); };
    const delToText = () => { for (const dt of [...el.getElementsByTagNameNS(W, 'delText')]) { const t = ctx.wEl(el.ownerDocument, 't'); t.setAttribute('xml:space', 'preserve'); t.textContent = dt.textContent; dt.replaceWith(t); } };
    if (name === 'ins' || name === 'moveTo') { if (el.parentNode.localName === 'rPr') { el.remove(); return; } if (accept) unwrap(); else el.remove(); }
    else if (name === 'del' || name === 'moveFrom') { if (el.parentNode.localName === 'rPr') { el.remove(); return; } if (accept) el.remove(); else { delToText(); unwrap(); } }
    else if (name === 'rPrChange' || name === 'pPrChange') {
      // Reject: put the old properties back.
      if (!accept) { const pr = el.parentNode; const old = ctx.child(el, name === 'rPrChange' ? 'rPr' : 'pPr'); for (const c of [...ctx.kids(pr)]) if (c !== el && !(name === 'pPrChange' && ['sectPr', 'rPr'].includes(c.localName))) c.remove(); if (old) for (const c of [...ctx.kids(old)]) pr.insertBefore(c, el); }
      el.remove();
    }
  }
  // which: a change id, or null for all; accept: true/false.
  E.reviewChanges = function (which, accept) {
    return this.structural((ctx) => {
      for (const name of CHANGE_TAGS) {
        for (const el of [...ctx.doc.getElementsByTagNameNS(W, name)]) {
          if (which != null && el.getAttributeNS(W, 'id') !== String(which)) continue;
          if (!el.parentNode) continue;
          resolveChange(ctx, el, accept);
        }
      }
    }, null, { anyPart: true });
  };


  // ---------------------------------------------------------------- equations (typed as text, written as Word math)
  // Input: a small LaTeX-like language: a/b or \frac{a}{b}, x^2, x_i, \sqrt{x}, \sum_{i=1}^{n}, \int_a^b,
  // Greek letters (\alpha ...), \pm \times \div \le \ge \ne \approx \infty \cdot \to, ( ) [ ].

  const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  const GREEK = { alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω', Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω' };
  const SYMBOLS = { pm: '±', times: '×', div: '÷', le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠', approx: '≈', infty: '∞', cdot: '·', to: '→', rightarrow: '→', leftarrow: '←', partial: '∂', nabla: '∇', in: '∈', forall: '∀', exists: '∃', degree: '°' };
  function parseMath(src) {
    let i = 0;
    const peek = () => src[i];
    const group = () => {
      while (src[i] === ' ') i++;
      if (src[i] === '{') { i++; const n = seq('}'); i++; return n; }
      return atom();
    };
    const atom = () => {
      while (src[i] === ' ') i++;
      const c = src[i];
      if (c === undefined) return { t: 'text', v: '' };
      if (c === '\\') {
        i++;
        let name = '';
        while (/[A-Za-z]/.test(src[i] || '')) name += src[i++];
        if (name === 'frac') { const num = group(); const den = group(); return { t: 'frac', num, den }; }
        if (name === 'sqrt') { let deg = null; if (src[i] === '[') { i++; deg = seq(']'); i++; } return { t: 'rad', deg, e: group() }; }
        if (name === 'sum' || name === 'prod' || name === 'int') return { t: 'nary', chr: { sum: '∑', prod: '∏', int: '∫' }[name] };
        if (GREEK[name]) return { t: 'text', v: GREEK[name] };
        if (SYMBOLS[name]) return { t: 'text', v: SYMBOLS[name] };
        return { t: 'text', v: name };
      }
      if (c === '(' || c === '[') { const close = c === '(' ? ')' : ']'; i++; const e = seq(close); i++; return { t: 'delim', open: c, close, e }; }
      if (/[0-9.]/.test(c)) { let v = ''; while (/[0-9.]/.test(src[i] || '')) v += src[i++]; return { t: 'text', v }; }
      i++;
      return { t: 'text', v: c };
    };
    const seq = (end) => {
      const items = [];
      while (i < src.length && src[i] !== end) {
        if (src[i] === ' ') { i++; continue; }
        let base = atom();
        // sub/superscripts
        let sub = null, sup = null;
        for (;;) {
          if (src[i] === '_') { i++; sub = group(); } else if (src[i] === '^') { i++; sup = group(); } else break;
        }
        if (base.t === 'nary') { base.sub = sub; base.sup = sup; base.e = { t: 'seq', items: [] }; items.push(base); continue; }
        if (sub && sup) base = { t: 'subsup', e: base, sub, sup }; else if (sub) base = { t: 'sub', e: base, sub }; else if (sup) base = { t: 'sup', e: base, sup };
        // a/b -> fraction of the neighbouring items
        if (src[i] === '/') { i++; const den = (() => { const a = atom(); let s2 = null, p2 = null; for (;;) { if (src[i] === '_') { i++; s2 = group(); } else if (src[i] === '^') { i++; p2 = group(); } else break; } return s2 && p2 ? { t: 'subsup', e: a, sub: s2, sup: p2 } : s2 ? { t: 'sub', e: a, sub: s2 } : p2 ? { t: 'sup', e: a, sup: p2 } : a; })(); items.push({ t: 'frac', num: base, den }); continue; }
        items.push(base);
      }
      // An n-ary operator takes everything after it as its argument.
      const k = items.findIndex((x) => x.t === 'nary');
      if (k >= 0 && k < items.length - 1) { items[k].e = { t: 'seq', items: items.slice(k + 1) }; items.length = k + 1; }
      return { t: 'seq', items };
    };
    return seq(undefined);
  }
  function mathXml(n) {
    const esc = (x) => x.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const e = (x) => `<m:e>${mathXml(x)}</m:e>`;
    switch (n.t) {
      case 'seq': return n.items.map(mathXml).join('');
      case 'text': return n.v ? `<m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t xml:space="preserve">${esc(n.v)}</m:t></m:r>` : '';
      case 'frac': return `<m:f><m:num>${mathXml(n.num)}</m:num><m:den>${mathXml(n.den)}</m:den></m:f>`;
      case 'sup': return `<m:sSup>${e(n.e)}<m:sup>${mathXml(n.sup)}</m:sup></m:sSup>`;
      case 'sub': return `<m:sSub>${e(n.e)}<m:sub>${mathXml(n.sub)}</m:sub></m:sSub>`;
      case 'subsup': return `<m:sSubSup>${e(n.e)}<m:sub>${mathXml(n.sub)}</m:sub><m:sup>${mathXml(n.sup)}</m:sup></m:sSubSup>`;
      case 'rad': return `<m:rad><m:radPr>${n.deg ? '' : '<m:degHide m:val="1"/>'}</m:radPr><m:deg>${n.deg ? mathXml(n.deg) : ''}</m:deg>${e(n.e)}</m:rad>`;
      case 'delim': return `<m:d><m:dPr><m:begChr m:val="${n.open}"/><m:endChr m:val="${n.close}"/></m:dPr>${e(n.e)}</m:d>`;
      case 'nary': return `<m:nary><m:naryPr><m:chr m:val="${n.chr}"/>${n.sub ? '' : '<m:subHide m:val="1"/>'}${n.sup ? '' : '<m:supHide m:val="1"/>'}</m:naryPr><m:sub>${n.sub ? mathXml(n.sub) : ''}</m:sub><m:sup>${n.sup ? mathXml(n.sup) : ''}</m:sup>${e(n.e || { t: 'seq', items: [] })}</m:nary>`;
      default: return '';
    }
  }
  DocxEditor.mathXml = (src) => mathXml(parseMath(src));
  E.insertEquation = function (src) {
    const inner = mathXml(parseMath(src));
    return this.atCaret((ctx, d) => [imp(d, `<m:oMath xmlns:m="${M}" xmlns:w="${W}">${inner}</m:oMath>`)], { bodyOnly: true });
  };

  // ---------------------------------------------------------------- pictures: size, alt text, crop

  // The picture (wp:inline / wp:anchor) the user selected, identified by its docPr id.
  E.imageInfo = function (img) {
    const span = img.closest('[data-dd-r]');
    const run = span && this.runAt(+span.dataset.ddR);
    const drawing = run && run.node.getElementsByTagNameNS(W, 'drawing')[0];
    const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
    const holder = drawing && (drawing.getElementsByTagNameNS(WP, 'inline')[0] || drawing.getElementsByTagNameNS(WP, 'anchor')[0]);
    if (!holder) return null;
    const ext = holder.getElementsByTagNameNS(WP, 'extent')[0];
    const docPr = holder.getElementsByTagNameNS(WP, 'docPr')[0];
    const src = holder.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'srcRect')[0];
    const pct = (a) => (src && src.getAttribute(a) ? +src.getAttribute(a) / 1000 : 0);
    return { id: docPr && docPr.getAttribute('id'), cx: +ext.getAttribute('cx'), cy: +ext.getAttribute('cy'), alt: docPr ? docPr.getAttribute('descr') || '' : '', crop: { l: pct('l'), t: pct('t'), r: pct('r'), b: pct('b') }, anchor: holder.localName === 'anchor' };
  };
  // o: { id, cx, cy (EMU), alt, crop: { l, t, r, b } (percent) }
  E.setImage = function (o) {
    const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    return this.structural((ctx) => {
      const docPr = [...ctx.doc.getElementsByTagNameNS(WP, 'docPr')].find((x) => x.getAttribute('id') === String(o.id));
      if (!docPr) throw new Error('The picture was not found.');
      const holder = docPr.parentNode;
      if (o.cx && o.cy) {
        const ext = holder.getElementsByTagNameNS(WP, 'extent')[0];
        ext.setAttribute('cx', Math.round(o.cx)); ext.setAttribute('cy', Math.round(o.cy));
        for (const e of holder.getElementsByTagNameNS(A, 'ext')) if (e.parentNode.localName === 'xfrm') { e.setAttribute('cx', Math.round(o.cx)); e.setAttribute('cy', Math.round(o.cy)); }
      }
      if (o.alt != null) { if (o.alt) docPr.setAttribute('descr', o.alt); else docPr.removeAttribute('descr'); }
      if (o.crop) {
        const blipFill = [...holder.getElementsByTagName('*')].find((x) => x.localName === 'blipFill');
        if (blipFill) {
          for (const s0 of [...blipFill.children].filter((x) => x.localName === 'srcRect')) s0.remove();
          const c = o.crop;
          if (c.l || c.t || c.r || c.b) {
            const s0 = blipFill.ownerDocument.createElementNS(A, 'a:srcRect');
            for (const k of ['l', 't', 'r', 'b']) if (c[k]) s0.setAttribute(k, Math.round(c[k] * 1000));
            const blip = [...blipFill.children].find((x) => x.localName === 'blip');
            blip.after(s0);
          }
        }
      }
    }, null, { anyPart: true });
  };

  // ---------------------------------------------------------------- section breaks

  E.insertSectionBreak = function (type) {
    return this.structural((ctx) => {
      const last = ctx.sectPrs().pop();
      const p = ctx.wEl(ctx.doc, 'p');
      const pPr = ctx.wEl(ctx.doc, 'pPr');
      const sp = last ? last.cloneNode(true) : ctx.wEl(ctx.doc, 'sectPr');
      for (const t of ctx.kids(sp).filter((c) => c.localName === 'type')) t.remove();
      sp.insertBefore(ctx.wEl(ctx.doc, 'type', { val: type === 'continuous' ? 'continuous' : 'nextPage' }), ctx.kids(sp).find((c) => !/Reference$/.test(c.localName)) || null);
      pPr.append(sp);
      p.append(pPr);
      ctx.insertAfterBlock(ctx.caretNode, p, ctx.doc);
    });
  };

  // ---------------------------------------------------------------- "Update style to match" (the caret's formatting)

  E.updateStyleToMatch = function (p) {
    const styleId = this.styleOf(p) || this.styles.defaultPara || 'Normal';
    const sel = getSelection();
    const el = sel.anchorNode && (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement);
    const cs = getComputedStyle(el && p.contains(el) ? el : p);
    const font = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    const sz = Math.round(parseFloat(cs.fontSize) * 0.75 * 2);
    const rgb = /rgba?\((\d+)[,\s]+(\d+)[,\s]+(\d+)/.exec(cs.color);
    const color = rgb ? [rgb[1], rgb[2], rgb[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase() : null;
    const bold = +cs.fontWeight >= 600, italic = cs.fontStyle === 'italic';
    const jc = { center: 'center', right: 'right', end: 'right', justify: 'both' }[getComputedStyle(p).textAlign] || 'left';
    return this.structural(async (ctx) => {
      const d = await ctx.partDoc(`${ctx.docDir}styles.xml`);
      const st = [...d.getElementsByTagNameNS(W, 'style')].find((x) => x.getAttributeNS(W, 'styleId') === styleId);
      if (!st) throw new Error(`The style ${styleId} was not found.`);
      let rPr = ctx.child(st, 'rPr');
      if (!rPr) { rPr = ctx.wEl(d, 'rPr'); st.append(rPr); }
      const set = (name, attrs) => { const old = ctx.child(rPr, name); if (old) old.remove(); if (attrs) ctx.setOrdered(rPr, ctx.wEl(d, name, attrs), ctx.RPR_ORDER); };
      set('rFonts', { ascii: font, hAnsi: font, cs: font });
      set('b', bold ? {} : null); set('bCs', bold ? {} : null);
      set('i', italic ? {} : null); set('iCs', italic ? {} : null);
      if (color) set('color', { val: color });
      set('sz', { val: sz }); set('szCs', { val: sz });
      let pPr = ctx.child(st, 'pPr');
      if (!pPr) { pPr = ctx.wEl(d, 'pPr'); st.insertBefore(pPr, rPr); }
      const jcOld = ctx.child(pPr, 'jc'); if (jcOld) jcOld.remove();
      if (jc !== 'left') ctx.setOrdered(pPr, ctx.wEl(d, 'jc', { val: jc }), ctx.PPR_ORDER);
    }, null, { anyPart: true });
  };

  // ---------------------------------------------------------------- restart numbering at the caret's list item

  E.restartNumbering = function () {
    const caret = this.caretParagraph();
    if (!caret) throw new Error('Click in a numbered list item first.');
    return this.structural(async (ctx) => {
      const p = ctx.caretNode;
      const numIdEl = p && p.getElementsByTagNameNS(W, 'numId')[0];
      if (!numIdEl) throw new Error('Click in a numbered list item first.');
      const oldId = numIdEl.getAttributeNS(W, 'val');
      const path = ctx.relTarget(REL('numbering'));
      const nd = path && await ctx.partDoc(path);
      if (!nd) throw new Error('This list has no numbering definitions.');
      const num = [...nd.getElementsByTagNameNS(W, 'num')].find((n) => n.getAttributeNS(W, 'numId') === oldId);
      const absId = num && ctx.child(num, 'abstractNumId').getAttributeNS(W, 'val');
      const newId = 1 + Math.max(0, ...[...nd.getElementsByTagNameNS(W, 'num')].map((n) => +n.getAttributeNS(W, 'numId')));
      nd.documentElement.append(imp(nd, `<w:num ${NSDECL} w:numId="${newId}"><w:abstractNumId w:val="${absId}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`));
      // This item and the ones after it in the same list continue from 1.
      const paras = [...ctx.doc.getElementsByTagNameNS(W, 'body')[0].getElementsByTagNameNS(W, 'p')];
      for (const q of paras.slice(paras.indexOf(p))) {
        const id = q.getElementsByTagNameNS(W, 'numId')[0];
        if (id && id.getAttributeNS(W, 'val') === oldId) id.setAttributeNS(W, 'w:val', String(newId));
      }
    });
  };

  // ---------------------------------------------------------------- dropdown chips (Word drop-down list content controls)

  E.insertDropdown = function (items, initial) {
    const esc = (x) => String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    const id = Math.floor(Math.random() * 2e9);
    return this.atCaret((ctx, d, rPr) => [imp(d, `<w:sdt ${NSDECL}><w:sdtPr><w:id w:val="${id}"/><w:dropDownList>${items.map((x) => `<w:listItem w:displayText="${esc(x)}" w:value="${esc(x)}"/>`).join('')}</w:dropDownList></w:sdtPr><w:sdtContent><w:r>${rPr ? new XMLSerializer().serializeToString(rPr) : ''}<w:t xml:space="preserve">${esc(initial || items[0])}</w:t></w:r></w:sdtContent></w:sdt>`)], {});
  };
  // Drop-down content controls in the document, keyed by run id, so clicking the text opens the list.
  E.dropdownAt = function (span) {
    const run = span && this.runAt(+span.dataset.ddR);
    const sdt = run && [...run.chain].reverse().find((w) => w.localName === 'sdt');
    const pr = sdt && DocxEditor.child(sdt, 'sdtPr');
    const list = pr && (DocxEditor.child(pr, 'dropDownList') || DocxEditor.child(pr, 'comboBox'));
    return list ? [...list.getElementsByTagNameNS(W, 'listItem')].map((li) => li.getAttributeNS(W, 'displayText') || li.getAttributeNS(W, 'value')) : null;
  };

  // ---------------------------------------------------------------- tables: rows and columns at the caret cell

  E.tableOp = function (op) {
    const caret = this.caretParagraph();
    if (!caret || !caret.closest('td,th')) throw new Error('Click inside a table cell first.');
    return this.structural((ctx) => {
      const p = ctx.caretNode;
      let tc = p; while (tc && !ctx.isW(tc, 'tc')) tc = tc.parentNode;
      if (!tc) throw new Error('Click inside a table cell first.');
      const tr = tc.parentNode, tbl = tr.parentNode;
      const cells = (row) => ctx.kids(row).filter((c) => ctx.isW(c, 'tc'));
      const gridIndex = (row, cell) => { let g = 0; for (const c of cells(row)) { if (c === cell) return g; const span = c.getElementsByTagNameNS(W, 'gridSpan')[0]; g += span ? +span.getAttributeNS(W, 'val') : 1; } return g; };
      const cellAtGrid = (row, gi) => { let g = 0; for (const c of cells(row)) { const span = c.getElementsByTagNameNS(W, 'gridSpan')[0]; const n = span ? +span.getAttributeNS(W, 'val') : 1; if (gi < g + n) return c; g += n; } return null; };
      const emptyCopy = (cell) => {
        const n = cell.cloneNode(false);
        const pr = ctx.child(cell, 'tcPr');
        if (pr) { const c = pr.cloneNode(true); for (const x of [...c.getElementsByTagNameNS(W, 'vMerge')]) x.remove(); n.append(c); }
        const firstP = ctx.child(cell, 'p');
        const np = ctx.wEl(ctx.doc, 'p');
        if (firstP && ctx.child(firstP, 'pPr')) np.append(ctx.child(firstP, 'pPr').cloneNode(true));
        n.append(np);
        return n;
      };
      const gi = gridIndex(tr, tc);
      const grid = ctx.child(tbl, 'tblGrid');
      const rows = ctx.kids(tbl).filter((c) => ctx.isW(c, 'tr'));
      if (op === 'rowAbove' || op === 'rowBelow') {
        const nr = tr.cloneNode(false);
        const trPr = ctx.child(tr, 'trPr');
        if (trPr) nr.append(trPr.cloneNode(true));
        for (const c of cells(tr)) nr.append(emptyCopy(c));
        tbl.insertBefore(nr, op === 'rowAbove' ? tr : tr.nextSibling);
      } else if (op === 'colLeft' || op === 'colRight') {
        for (const row of rows) {
          const c = cellAtGrid(row, gi) || cells(row).pop();
          if (!c) continue;
          const nc = emptyCopy(c);
          const span = nc.getElementsByTagNameNS(W, 'gridSpan')[0]; if (span) span.remove();
          row.insertBefore(nc, op === 'colLeft' ? c : c.nextSibling);
        }
        if (grid) { const cols = ctx.kids(grid); const ref = cols[gi] || cols[cols.length - 1]; if (ref) grid.insertBefore(ref.cloneNode(true), op === 'colLeft' ? ref : ref.nextSibling); }
      } else if (op === 'deleteRow') {
        tr.remove();
        if (!ctx.kids(tbl).some((c) => ctx.isW(c, 'tr'))) tbl.remove();
      } else if (op === 'deleteColumn') {
        for (const row of rows) { const c = cellAtGrid(row, gi); if (c) c.remove(); }
        if (grid && ctx.kids(grid)[gi]) ctx.kids(grid)[gi].remove();
        if (!rows.some((row) => cells(row).length)) tbl.remove();
      } else if (op === 'deleteTable') {
        tbl.remove();
      }
      // The body must keep at least one paragraph after the table's removal.
      const body = ctx.doc.getElementsByTagNameNS(W, 'body')[0];
      if (!ctx.kids(body).some((c) => ctx.isW(c, 'p') || ctx.isW(c, 'tbl'))) body.insertBefore(ctx.wEl(ctx.doc, 'p'), ctx.child(body, 'sectPr'));
    });
  };
})();
