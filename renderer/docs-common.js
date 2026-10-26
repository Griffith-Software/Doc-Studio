'use strict';
// Google Docs-style tools shared by the text editors (Word, OpenDocument, RTF; some also Markdown):
// special characters, date, change case, checklists, paint format, document outline, non-printing characters,
// automatic substitutions, compare documents, the table grid picker, and the Word dialogs (spacing, page setup,
// page numbers, footnotes, comments, bookmarks). Menu items for each editor are built by docMenus(tab).

const DocTools = (() => {
  // ---------------------------------------------------------------- where the caret is

  function editableAt() {
    const sel = getSelection();
    if (!sel.rangeCount) return null;
    const n = sel.anchorNode;
    const el = n && (n.nodeType === 1 ? n : n.parentElement);
    return el && el.closest('[contenteditable="true"]');
  }
  // Keep the last selection inside the document, so dialogs (which take focus) can put it back.
  let saved = null;
  document.addEventListener('selectionchange', () => {
    const ed = editableAt();
    if (ed && active && active.body.contains(ed)) saved = { range: getSelection().getRangeAt(0).cloneRange(), el: ed };
    const ta = active && active.textarea;
    if (ta && document.activeElement === ta) saved = { ta, start: ta.selectionStart, end: ta.selectionEnd };
  });
  function restore() {
    if (!saved) return false;
    if (saved.ta) { if (!saved.ta.isConnected) return false; saved.ta.focus(); saved.ta.setSelectionRange(saved.start, saved.end); return true; }
    if (!saved.el.isConnected) return false;
    saved.el.focus({ preventScroll: true });
    const s = getSelection(); s.removeAllRanges(); s.addRange(saved.range);
    return true;
  }
  function insertText(text) {
    if (!restore()) { flash('Click in the document where it should go first.'); return false; }
    document.execCommand('insertText', false, text);
    return true;
  }

  // ---------------------------------------------------------------- special characters

  const CHARS = {
    'Popular': '©®™°±×÷≠≈≤≥∞µ€£¥¢§¶•…–—‘’“”«»†‡′″½¼¾¹²³√∑∏∫∆∂π',
    'Arrows': '←↑→↓↔↕⇐⇑⇒⇓⇔↖↗↘↙↩↪⇄⇆➔➜➤⟵⟶⟷',
    'Math': '±×÷=≠≈≡≤≥<>∞∑∏∫∮√∛∜∂∆∇∈∉∋∩∪⊂⊃⊆⊇∧∨¬∀∃∅∝∠⊥∥°′″%‰',
    'Greek': 'αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ',
    'Currency': '$€£¥¢₹₽₩₪₫₱₴₦₺₿฿₡₲',
    'Punctuation': '…–—‘’‚“”„«»‹›¡¿·•‣◦§¶†‡※‽⁂',
    'Latin': 'ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝßàáâãäåæçèéêëìíîïñòóôõöøùúûüýÿŒœŠšŽžŁłĆćŃńŚśŹźŻż',
    'Shapes': '■□▪▫▲△▶▷▼▽◀◁◆◇○●◎★☆♠♣♥♦✓✔✗✘☐☑☒',
    'Emoji': '😀😃😄😁😆😅😂🙂😉😊😍😘😎🤔😐😴😢😭😡👍👎👏🙌🙏💪👋✌️👀🎉🎂🎁🔥⭐✨💡📌📎📅📈📉✅❌⚠️❓❗💬📝📞📧🏠🚗✈️🌞🌧️❄️☕🍕',
  };
  function specialCharacters() {
    const box = h('div', { class: 'dialog chars-dialog' }, h('h3', { text: 'Insert special characters' }));
    const search = h('input', { type: 'search', placeholder: 'Search by name, or type a character', class: 'chars-search' });
    const tabsRow = h('div', { class: 'chars-tabs' });
    const grid = h('div', { class: 'chars-grid' });
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    const show = (set) => {
      grid.replaceChildren(...[...set].filter((c) => c.trim() && c !== '️').map((c) => h('button', { class: 'ghost char', text: c, title: `U+${c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`, onclick: () => { insertText(c); } })));
    };
    for (const name of Object.keys(CHARS)) tabsRow.append(h('button', { class: 'ghost', text: name, onclick: (e) => { tabsRow.querySelectorAll('.on').forEach((b) => b.classList.remove('on')); e.target.classList.add('on'); show(CHARS[name]); } }));
    search.addEventListener('input', () => {
      const q = search.value.trim();
      if (!q) { show(CHARS.Popular); return; }
      if ([...q].length === 1 && !/[a-z0-9]/i.test(q)) { show(q); return; }
      // Code point (U+00E9 or 00e9) or a name match against the category names.
      const cp = /^(?:u\+)?([0-9a-f]{4,6})$/i.exec(q);
      if (cp) { try { show(String.fromCodePoint(parseInt(cp[1], 16))); return; } catch { /* invalid */ } }
      const hits = Object.entries(CHARS).filter(([n]) => n.toLowerCase().includes(q.toLowerCase())).map(([, v]) => v).join('');
      show(hits);
    });
    box.append(search, tabsRow, grid, h('div', { class: 'dialog-actions' }, h('button', { class: 'ghost', text: 'Close', onclick: close })));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
    tabsRow.firstChild.classList.add('on');
    show(CHARS.Popular);
    search.focus();
  }

  // ---------------------------------------------------------------- text changes that keep formatting

  // Transform the text of the selection node by node, so bold/colours etc. stay where they are.
  function transformSelection(fn) {
    if (!restore()) return;
    const sel = getSelection();
    const r = sel.getRangeAt(0);
    if (r.collapsed) { flash('Select some text first.'); return; }
    const root = editableAt();
    const nodes = [];
    const walker = document.createTreeWalker(r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) if (r.intersectsNode(n)) nodes.push(n);
    const sc = r.startContainer, so = r.startOffset, ec = r.endContainer, eo = r.endOffset;
    let i = 0;
    for (const n of nodes) {
      const a = n === sc ? so : 0, b = n === ec ? eo : n.data.length;
      n.data = n.data.slice(0, a) + fn(n.data.slice(a, b), i++) + n.data.slice(b);
    }
    try { const nr = document.createRange(); nr.setStart(sc, Math.min(so, sc.length ?? so)); nr.setEnd(ec, Math.min(eo, ec.length ?? eo)); sel.removeAllRanges(); sel.addRange(nr); } catch { /* fine */ }
    root.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertReplacementText' }));
  }
  const toTitle = (s) => s.toLowerCase().replace(/(^|[\s\-("'])(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
  function changeCase(kind) {
    if (active && active.textarea && document.activeElement !== active.textarea && saved && saved.ta) {
      const ta = active.textarea; restore();
      const t = ta.value.slice(ta.selectionStart, ta.selectionEnd);
      if (!t) { flash('Select some text first.'); return; }
      document.execCommand('insertText', false, kind === 'upper' ? t.toUpperCase() : kind === 'lower' ? t.toLowerCase() : toTitle(t));
      return;
    }
    transformSelection((t) => (kind === 'upper' ? t.toUpperCase() : kind === 'lower' ? t.toLowerCase() : toTitle(t)));
  }

  // ---------------------------------------------------------------- checklists (☐ / ☑ at the start of paragraphs)

  const BOX = '☐', TICK = '☑';
  function paragraphsOfSelection() {
    const ed = active && (active.editor || active.richEditor);
    if (!ed) return [];
    if (!restore()) return [];
    return ed.selectedParagraphs();
  }
  function toggleChecklist() {
    const ps = paragraphsOfSelection();
    if (!ps.length) { flash('Click in the document first.'); return; }
    const all = ps.every((p) => /^[☐☑]\s/.test(p.textContent));
    const sel = getSelection();
    for (const p of ps) {
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      let first = walker.nextNode();
      while (first && !first.data.length) first = walker.nextNode();
      if (all) { if (first) first.data = first.data.replace(/^[☐☑]\s/, ''); continue; }
      if (/^[☐☑]\s/.test(p.textContent)) continue;
      if (first) first.data = `${BOX} ${first.data}`;
      else { const r = document.createRange(); r.selectNodeContents(p); r.collapse(true); sel.removeAllRanges(); sel.addRange(r); document.execCommand('insertText', false, `${BOX} `); }
    }
    const root = ps[0].closest('[contenteditable="true"]');
    if (root) root.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertReplacementText' }));
  }
  // Clicking a box ticks it.
  document.addEventListener('click', (e) => {
    const root = e.target.closest && e.target.closest('[contenteditable="true"]');
    if (!root || !document.caretRangeFromPoint) return;
    const r = document.caretRangeFromPoint(e.clientX, e.clientY);
    if (!r || r.startContainer.nodeType !== 3) return;
    const n = r.startContainer;
    for (const off of [r.startOffset, r.startOffset - 1]) {
      const ch = n.data[off];
      if (ch !== BOX && ch !== TICK) continue;
      // Only a box at the start of its paragraph is a checklist box.
      const p = n.parentElement.closest('p,div,li,h1,h2,h3,h4,h5,h6');
      if (!p || !p.textContent.startsWith(ch)) continue;
      n.data = n.data.slice(0, off) + (ch === BOX ? TICK : BOX) + n.data.slice(off + 1);
      root.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertReplacementText' }));
      e.preventDefault();
      return;
    }
  });

  // ---------------------------------------------------------------- paint format

  let painter = null;
  function paintFormat(btn) {
    if (painter) { stopPaint(); return; }
    if (!restore()) { flash('Click in the text whose formatting you want to copy.'); return; }
    const sel = getSelection();
    const n = sel.anchorNode;
    const el = n && (n.nodeType === 1 ? n : n.parentElement);
    if (!el) return;
    const cs = getComputedStyle(el);
    painter = {
      bold: +cs.fontWeight >= 600 || cs.fontWeight === 'bold', italic: cs.fontStyle === 'italic',
      underline: /underline/.test(cs.textDecorationLine), strike: /line-through/.test(cs.textDecorationLine),
      color: cs.color, font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(), size: `${Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2}pt`,
      bg: (() => { for (let x = el; x && x.isContentEditable && x.getAttribute('contenteditable') !== 'true'; x = x.parentElement) { const b = getComputedStyle(x).backgroundColor; if (b && b !== 'rgba(0, 0, 0, 0)' && b !== 'transparent') return b; } return null; })(),
      btn,
    };
    if (btn) btn.classList.add('on');
    document.body.classList.add('painting');
    flash('Select the text to apply the formatting to (Esc cancels)');
  }
  function stopPaint() { if (painter && painter.btn) painter.btn.classList.remove('on'); painter = null; document.body.classList.remove('painting'); }
  document.addEventListener('mouseup', () => {
    if (!painter) return;
    setTimeout(() => {
      const sel = getSelection();
      if (!sel.rangeCount || sel.getRangeAt(0).collapsed || !editableAt()) return;
      const p = painter;
      stopPaint();
      document.execCommand('styleWithCSS', false, true);
      const state = (cmd) => { try { return document.queryCommandState(cmd); } catch { return false; } };
      if (state('bold') !== p.bold) document.execCommand('bold');
      if (state('italic') !== p.italic) document.execCommand('italic');
      if (state('underline') !== p.underline) document.execCommand('underline');
      if (state('strikeThrough') !== p.strike) document.execCommand('strikeThrough');
      document.execCommand('foreColor', false, p.color);
      document.execCommand('fontName', false, p.font);
      document.execCommand('fontSize', false, '7');
      const root = editableAt();
      root.querySelectorAll('[style*="xxx-large"], font[size="7"]').forEach((x) => { x.style.fontSize = p.size; x.removeAttribute('size'); });
      if (p.bg) document.execCommand('hiliteColor', false, p.bg);
      root.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'formatFontName' }));
    }, 0);
  });
  document.addEventListener('keydown', (e) => { if (painter && e.key === 'Escape') stopPaint(); });

  // ---------------------------------------------------------------- font size − / +

  const SIZES = [6, 7, 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
  function stepFontSize(dir) {
    if (!restore()) return;
    const sel = getSelection();
    const n = sel.anchorNode;
    const el = n && (n.nodeType === 1 ? n : n.parentElement);
    const cur = Math.round(parseFloat(getComputedStyle(el).fontSize) * 0.75 * 2) / 2;
    const next = dir > 0 ? SIZES.find((s) => s > cur) || cur + 2 : [...SIZES].reverse().find((s) => s < cur) || Math.max(1, cur - 1);
    if (sel.getRangeAt(0).collapsed) { flash(`Select text to change its size (now ${cur} pt)`); return; }
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontSize', false, '7');
    const root = editableAt();
    root.querySelectorAll('[style*="xxx-large"], font[size="7"]').forEach((x) => {
      if (x.tagName === 'FONT') { const sp = document.createElement('span'); sp.style.fontSize = `${next}pt`; while (x.firstChild) sp.append(x.firstChild); x.replaceWith(sp); } else x.style.fontSize = `${next}pt`;
    });
    root.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'formatFontSize' }));
    flash(`${next} pt`);
  }

  // ---------------------------------------------------------------- document outline

  function headingsOf(tab) {
    let root = null;
    if (tab.kind === 'docx') root = tab.mode === 'edit' && tab.editor ? tab.editor.stage : tab.viewBox;
    else if (tab.richEditor) root = tab.richEditor.article;
    else if (tab.previewPane) root = tab.previewPane;
    if (!root) return [];
    const out = [];
    for (const el of root.querySelectorAll('h1,h2,h3,h4,h5,h6,p')) {
      let lvl = 0;
      if (/^H\d$/.test(el.tagName)) lvl = +el.tagName[1];
      else { const m = /docx_(?:heading|berschrift|titre)(\d)/i.exec(el.className); if (m) lvl = +m[1]; else if (/docx_title\b/i.test(el.className)) lvl = 1; }
      const text = el.textContent.replace(/[-]/g, '').trim();
      if (lvl && text) out.push({ el, lvl, text });
    }
    return out;
  }
  function toggleOutline(tab) {
    if (!tab) return;
    if (tab.outlineEl) { tab.outlineEl.remove(); tab.outlineEl = null; tab.body.classList.remove('with-outline'); return; }
    const panel = h('aside', { class: 'outline-panel' }, h('div', { class: 'outline-head' }, h('b', { text: 'Outline' }), h('button', { class: 'ghost', text: '×', title: 'Close', onclick: () => toggleOutline(tab) })));
    const list = h('div', { class: 'outline-list' });
    panel.append(list);
    const fill = () => {
      const hs = headingsOf(tab);
      list.replaceChildren(...(hs.length ? hs.map((x) => h('div', { class: `outline-item l${Math.min(4, x.lvl)}`, text: x.text, title: x.text, onclick: () => x.el.scrollIntoView({ block: 'start', behavior: 'smooth' }) })) : [h('p', { class: 'muted', text: 'Headings you add to the document appear here.' })]));
    };
    fill();
    tab.outlineFill = fill;
    tab.body.append(panel);
    tab.body.classList.add('with-outline');
    tab.outlineEl = panel;
    clearInterval(tab.outlineTimer);
    tab.outlineTimer = setInterval(() => { if (!tab.outlineEl || !tab.outlineEl.isConnected) { clearInterval(tab.outlineTimer); return; } fill(); }, 2500);
  }

  // ---------------------------------------------------------------- non-printing characters

  function toggleMarks(tab) { tab.body.classList.toggle('show-marks'); flash(tab.body.classList.contains('show-marks') ? 'Showing paragraph marks' : 'Paragraph marks hidden'); }

  // ---------------------------------------------------------------- automatic substitutions (Tools > Preferences)

  const PREF_KEY = 'docstudio.prefs';
  const prefs = (() => { try { return { smartQuotes: true, substitutions: true, ...JSON.parse(localStorage.getItem(PREF_KEY) || '{}') }; } catch { return { smartQuotes: true, substitutions: true }; } })();
  const SUBS = [['(c)', '©'], ['(r)', '®'], ['(tm)', '™'], ['-->', '→'], ['<--', '←'], ['<->', '↔'], ['==>', '⇒'], ['--', '—'], ['...', '…'], ['1/2', '½'], ['1/4', '¼'], ['3/4', '¾'], ['!=', '≠'], ['<=', '≤'], ['>=', '≥']];
  document.addEventListener('beforeinput', (e) => {
    if (e.inputType !== 'insertText' || !e.data) return;
    const root = e.target.closest && e.target.closest('[contenteditable="true"]');
    if (!root || !active || !active.body.contains(root)) return;
    const sel = getSelection();
    if (!sel.rangeCount || !sel.getRangeAt(0).collapsed) return;
    const r = sel.getRangeAt(0);
    if (r.startContainer.nodeType !== 3) return;
    const node = r.startContainer, off = r.startOffset;
    const before = node.data.slice(0, off);
    // Smart quotes.
    if (prefs.smartQuotes && (e.data === '"' || e.data === "'")) {
      const open = !before || /[\s([{—–-]$/.test(before);
      e.preventDefault();
      document.execCommand('insertText', false, e.data === '"' ? (open ? '“' : '”') : (open ? '‘' : '’'));
      return;
    }
    // Substitutions when a word ends.
    if (prefs.substitutions && /[\s.,;:!?)]/.test(e.data)) {
      for (const [from, to] of SUBS) {
        if (!before.endsWith(from)) continue;
        if (from === '...' && e.data === '.') continue;
        e.preventDefault();
        const rr = document.createRange();
        rr.setStart(node, off - from.length); rr.setEnd(node, off);
        sel.removeAllRanges(); sel.addRange(rr);
        document.execCommand('insertText', false, to + e.data);
        return;
      }
    }
  });
  function preferences() {
    const box = h('div', { class: 'dialog' }, h('h3', { text: 'Preferences' }));
    const row = (key, label) => { const cb = h('input', { type: 'checkbox' }); cb.checked = !!prefs[key]; cb.addEventListener('change', () => { prefs[key] = cb.checked; try { localStorage.setItem(PREF_KEY, JSON.stringify(prefs)); } catch { /* not saved */ } }); return h('label', { class: 'check-row' }, cb, h('span', { text: label })); };
    box.append(row('smartQuotes', 'Use smart quotes (“ ” ‘ ’)'), row('substitutions', 'Automatic substitutions: (c) → ©, -- → —, ... → …, --> → →, 1/2 → ½ and more'));
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    box.append(h('div', { class: 'dialog-actions' }, h('button', { text: 'Done', onclick: close })));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
  }

  // ---------------------------------------------------------------- compare documents

  async function textParagraphs(tab) {
    const html = await Convert.run(tab, 'txt', { office: false }).then((b) => new TextDecoder().decode(b)).catch(() => null);
    if (html != null) return html.split(/\n+/).map((x) => x.trim()).filter(Boolean);
    return documentText(tab).split(/\n+/).map((x) => x.trim()).filter(Boolean);
  }
  function diff(a, b) {
    // Longest common subsequence on items (paragraphs or words).
    const n = a.length, m = b.length;
    const L = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = a[i] === b[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const out = [];
    let i = 0, j = 0;
    while (i < n && j < m) {
      if (a[i] === b[j]) { out.push({ t: '=', v: a[i] }); i++; j++; } else if (L[i + 1][j] >= L[i][j + 1]) out.push({ t: '-', v: a[i++] }); else out.push({ t: '+', v: b[j++] });
    }
    while (i < n) out.push({ t: '-', v: a[i++] });
    while (j < m) out.push({ t: '+', v: b[j++] });
    return out;
  }
  // otherPath: the changed version. Without it, the Compare tool window opens with this document as the original.
  async function compareDocuments(tab, otherPath) {
    if (!tab) return;
    if (!otherPath) return Tools.compare(tab);
    const f = otherPath;
    if (!f) return;
    flash('Comparing…');
    const before = tabs.length;
    const cur = active;
    await openPath(f);
    const other = active;
    if (!other || other === cur) { activate(cur); return; }
    let A, B;
    try { A = await textParagraphs(cur); B = await textParagraphs(other); } finally { if (tabs.length > before) await closeTab(other); activate(cur); }
    if (A.length * B.length > 4e6) { desk.error('These documents are too long to compare here.'); return; }
    const ops = diff(A, B);
    // Pair removed + added paragraphs and show word-level changes inside them.
    const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const parts = [];
    let adds = 0, dels = 0;
    for (let k = 0; k < ops.length; k++) {
      const o = ops[k];
      if (o.t === '=') { parts.push(`<p>${esc(o.v)}</p>`); continue; }
      if (o.t === '-' && ops[k + 1] && ops[k + 1].t === '+') {
        const words = diff(o.v.split(/(\s+)/), ops[k + 1].v.split(/(\s+)/));
        parts.push(`<p class="changed">${words.map((w) => (w.t === '=' ? esc(w.v) : w.t === '-' ? `<del>${esc(w.v)}</del>` : `<ins>${esc(w.v)}</ins>`)).join('')}</p>`);
        adds++; dels++; k++;
        continue;
      }
      if (o.t === '-') { dels++; parts.push(`<p class="removed"><del>${esc(o.v)}</del></p>`); } else { adds++; parts.push(`<p class="added"><ins>${esc(o.v)}</ins></p>`); }
    }
    const summary = `<div class="notice">Comparing <b>${esc(cur.name)}</b> with <b>${esc(f.split(/[\\/]/).pop())}</b>: ${dels} paragraph(s) removed or changed, ${adds} added or changed. <del>Struck</del> text is only in ${esc(cur.name)}; <ins>underlined</ins> text is only in the other file.</div>`;
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Comparison</title><style>body{font:15px/1.6 "Segoe UI",sans-serif;max-width:860px;margin:30px auto;padding:0 20px;color:#1f2328}del{background:#ffd7d5;color:#82071e}ins{background:#ccffd8;color:#055d20;text-decoration:none;border-bottom:1px solid #2da44e}.notice{background:#f3f4f6;padding:10px 14px;border-radius:6px;margin-bottom:18px}</style></head><body>${summary}${parts.join('\n')}</body></html>`;
    untitledSeq++;
    const t = makeTab({ path: null, name: `Comparison ${untitledSeq}.html`, size: 0, data: new TextEncoder().encode(html).buffer });
    await setupTab(t);
    refreshChrome();
  }

  // ---------------------------------------------------------------- table grid picker

  function tableGrid(anchor, onPick) {
    const MAX = 10;
    const panel = h('div', { class: 'menu-panel grid-picker' });
    const label = h('div', { class: 'grid-label', text: 'Insert table' });
    const grid = h('div', { class: 'grid-cells' });
    for (let r = 1; r <= MAX; r++) for (let c = 1; c <= MAX; c++) {
      const cell = h('div', { class: 'grid-cell', 'data-r': r, 'data-c': c });
      cell.addEventListener('mouseenter', () => { grid.querySelectorAll('.grid-cell').forEach((x) => x.classList.toggle('on', +x.dataset.r <= r && +x.dataset.c <= c)); label.textContent = `${c} × ${r}`; });
      cell.addEventListener('mousedown', (e) => { e.preventDefault(); panel.remove(); onPick(r, c); });
      grid.append(cell);
    }
    panel.append(grid, label);
    const rr = anchor.getBoundingClientRect();
    panel.style.left = `${rr.left}px`;
    panel.style.top = `${rr.bottom + 4}px`;
    document.body.append(panel);
    const off = (e) => { if (!panel.contains(e.target)) { panel.remove(); document.removeEventListener('mousedown', off, true); } };
    setTimeout(() => document.addEventListener('mousedown', off, true), 0);
  }

  // ---------------------------------------------------------------- Word dialogs

  const IN = 1440; // twips per inch
  function form(title, fields, okText) {
    return new Promise((resolve) => {
      const box = h('form', { class: 'dialog form-dialog' }, h('h3', { text: title }));
      const els = {};
      for (const f of fields) {
        let input;
        if (f.type === 'select') input = h('select', { name: f.name }, f.options.map(([v, t]) => h('option', { value: v, text: t })));
        else if (f.type === 'checkbox') input = h('input', { type: 'checkbox', name: f.name });
        else if (f.type === 'textarea') input = h('textarea', { name: f.name, rows: 4 });
        else input = h('input', { type: f.type || 'text', name: f.name, step: f.step || null, min: f.min ?? null });
        if (f.type === 'checkbox') input.checked = !!f.value; else input.value = f.value ?? '';
        els[f.name] = input;
        box.append(f.type === 'checkbox' ? h('label', { class: 'check-row' }, input, h('span', { text: f.label })) : h('label', null, h('span', { text: f.label }), input));
      }
      const back = h('div', { class: 'backdrop' }, box);
      const done = (v) => { back.remove(); resolve(v); };
      box.append(h('div', { class: 'dialog-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Cancel', onclick: () => done(null) }), h('button', { type: 'submit', text: okText || 'OK' })));
      box.addEventListener('submit', (e) => { e.preventDefault(); done(Object.fromEntries(Object.entries(els).map(([k, el]) => [k, el.type === 'checkbox' ? el.checked : el.value]))); });
      back.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.stopPropagation(); done(null); } });
      back.addEventListener('mousedown', (e) => { if (e.target === back) done(null); });
      document.body.append(back);
      const first = box.querySelector('input,select,textarea');
      if (first) first.focus();
    });
  }
  const PAPER = { letter: ['Letter (8.5" × 11")', 12240, 15840], legal: ['Legal (8.5" × 14")', 12240, 20160], tabloid: ['Tabloid (11" × 17")', 15840, 24480], statement: ['Statement (5.5" × 8.5")', 7920, 12240], executive: ['Executive (7.25" × 10.5")', 10440, 15120], a3: ['A3 (29.7 × 42 cm)', 16838, 23811], a4: ['A4 (21 × 29.7 cm)', 11906, 16838], a5: ['A5 (14.8 × 21 cm)', 8391, 11906], b5: ['B5 (17.6 × 25 cm)', 9979, 14175] };
  async function pageSetup(ed) {
    const cur = ed.pageSetupInfo();
    const portrait = [Math.min(cur.width, cur.height), Math.max(cur.width, cur.height)];
    const match = Object.entries(PAPER).find(([, [, w, hh]]) => Math.abs(w - portrait[0]) < 60 && Math.abs(hh - portrait[1]) < 60);
    const inch = (t) => (t / IN).toFixed(2);
    const v = await form('Page setup', [
      { name: 'orientation', label: 'Orientation', type: 'select', value: cur.orientation, options: [['portrait', 'Portrait'], ['landscape', 'Landscape']] },
      { name: 'paper', label: 'Paper size', type: 'select', value: match ? match[0] : 'custom', options: [...Object.entries(PAPER).map(([k, [t]]) => [k, t]), ['custom', `Current (${inch(portrait[0])}" × ${inch(portrait[1])}")`]] },
      { name: 'top', label: 'Top margin (inches)', type: 'number', step: '0.05', min: '0', value: inch(cur.margins.top) },
      { name: 'bottom', label: 'Bottom margin (inches)', type: 'number', step: '0.05', min: '0', value: inch(cur.margins.bottom) },
      { name: 'left', label: 'Left margin (inches)', type: 'number', step: '0.05', min: '0', value: inch(cur.margins.left) },
      { name: 'right', label: 'Right margin (inches)', type: 'number', step: '0.05', min: '0', value: inch(cur.margins.right) },
      { name: 'columns', label: 'Columns', type: 'select', value: String(cur.columns || 1), options: [['1', 'One'], ['2', 'Two'], ['3', 'Three']] },
      { name: 'color', label: 'Page colour', type: 'color', value: cur.color || '#ffffff' },
    ], 'Apply');
    if (!v) return;
    const size = v.paper === 'custom' ? portrait : [PAPER[v.paper][1], PAPER[v.paper][2]];
    const tw = (x) => Math.max(0, Math.round(parseFloat(x) * IN) || 0);
    await ed.pageSetup({ width: size[0], height: size[1], orientation: v.orientation, margins: { top: tw(v.top), right: tw(v.right), bottom: tw(v.bottom), left: tw(v.left) }, color: v.color, columns: +v.columns });
  }
  async function customSpacing(ed) {
    if (!restore()) { flash('Click in a paragraph first.'); return; }
    const ps = ed.selectedParagraphs();
    if (!ps.length) return;
    const cur = ed.spacingOf(ps[0]);
    const v = await form('Custom spacing', [
      { name: 'line', label: 'Line spacing (1 = single)', type: 'number', step: '0.05', min: '0.5', value: cur.line },
      { name: 'before', label: 'Space before (pt)', type: 'number', step: '1', min: '0', value: Math.round((cur.before || 0) / 20) },
      { name: 'after', label: 'Space after (pt)', type: 'number', step: '1', min: '0', value: Math.round((cur.after || 0) / 20) },
      { name: 'keepNext', label: 'Keep with next paragraph', type: 'checkbox', value: !!cur.keepNext },
    ], 'Apply');
    if (!v) return;
    ed.setSpacing(ps, { line: Math.max(0.5, parseFloat(v.line) || 1), before: Math.max(0, parseFloat(v.before) || 0) * 20, after: Math.max(0, parseFloat(v.after) || 0) * 20, keepNext: v.keepNext });
  }
  function setLineSpacing(ed, line) {
    if (!restore()) { flash('Click in a paragraph first.'); return; }
    ed.setSpacing(ed.selectedParagraphs(), { line });
  }
  function addSpace(ed, which, on) {
    if (!restore()) { flash('Click in a paragraph first.'); return; }
    ed.setSpacing(ed.selectedParagraphs(), { [which]: on ? 240 : 0 });
  }
  const run = (label, fn) => async () => { try { await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } };
  function tableOp(ed, op) {
    restore();
    const p = ed.caretParagraph();
    if (!p || !p.closest('td,th')) { flash('Click inside a table cell first.'); return; }
    return ed.tableOp(op);
  }
  async function footnote(ed) {
    const v = await form('Insert footnote', [{ name: 'text', label: 'Footnote text', type: 'textarea', value: '' }], 'Insert');
    if (!v) return;
    restore();
    await ed.insertFootnote(v.text);
  }
  async function comment(ed) {
    const v = await form('Add comment', [{ name: 'text', label: 'Comment', type: 'textarea', value: '' }, { name: 'author', label: 'Your name (shown with the comment)', value: localStorage.getItem('docstudio.author') || '' }], 'Comment');
    if (!v || !v.text.trim()) return;
    try { localStorage.setItem('docstudio.author', v.author || ''); } catch { /* not kept */ }
    restore();
    await ed.insertComment(v.text, v.author || 'Doc Studio');
  }
  async function commentsPanel(ed) {
    const list = await ed.comments();
    const box = h('div', { class: 'dialog versions-dialog' }, h('h3', { text: 'Comments' }));
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    if (!list.length) box.append(h('p', { class: 'muted', text: 'This document has no comments. Select text and use Insert > Comment to add one.' }));
    for (const c of list) {
      box.append(h('div', { class: 'comment-row' },
        h('div', null, h('b', { text: c.author || 'Comment' }), h('span', { class: 'muted', text: c.date ? `  ${new Date(c.date).toLocaleString()}` : '' })),
        h('div', { class: 'comment-text', text: c.text }),
        h('button', { class: 'ghost', text: 'Resolve (delete)', onclick: run('delete the comment', async () => { close(); await ed.deleteComment(c.id); flash('Comment removed'); }) })));
    }
    box.append(h('div', { class: 'dialog-actions' }, h('button', { class: 'ghost', text: 'Close', onclick: close })));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
  }
  async function bookmark(ed) {
    const v = await form('Insert bookmark', [{ name: 'name', label: 'Bookmark name', value: 'Bookmark1' }], 'Insert');
    if (!v) return;
    restore();
    await ed.insertBookmark(v.name);
    flash('Bookmark added - link to it with Insert > Link and #name');
  }
  async function pageNumbers(ed) {
    const v = await form('Page numbers', [
      { name: 'where', label: 'Position', type: 'select', value: 'footer', options: [['footer', 'Bottom of page (footer)'], ['header', 'Top of page (header)']] },
      { name: 'align', label: 'Alignment', type: 'select', value: 'right', options: [['right', 'Right'], ['center', 'Center'], ['left', 'Left']] },
      { name: 'withTotal', label: 'Show "Page X of Y"', type: 'checkbox', value: false },
    ], 'Insert');
    if (!v) return;
    await ed.insertPageNumbers(v);
  }

  // ---------------------------------------------------------------- menus for the text editors

  function docMenus(tab) {
    const it = (label, fn, o) => ({ label, run: run(label.toLowerCase(), fn), ...(o || {}) });
    const sub = (label, items) => ({ label, sub: items });
    const SEP = { sep: true };
    const k = tab.kind;
    const editing = () => tab.mode === 'edit' || tab.mode === 'split';
    const en = { enabled: editing };
    const m = { Insert: [], Format: [], View: [], Tools: [] };
    if (k === 'docx') {
      const ed = () => tab.editor;
      m.Insert.push(
        it('Table…', () => new Promise((resolve) => tableGrid($('#menu-names'), async (r, c) => { restore(); await ed().insertTable(r, c); resolve(); })), en),
        it('Horizontal line', () => { restore(); return ed().insertHorizontalLine(); }, en),
        it('Special characters…', () => specialCharacters(), en),
        it('Date', () => insertText(new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })), en),
        it('Checklist', () => toggleChecklist(), { ...en, key: 'Ctrl+Shift+9' }),
        SEP,
        it('Footnote…', () => footnote(ed()), { ...en, key: 'Ctrl+Alt+F' }),
        it('Comment…', () => comment(ed()), { ...en, key: 'Ctrl+Alt+M' }),
        it('Bookmark…', () => bookmark(ed()), en),
        SEP,
        sub('Headers & footers', [it('Header', () => ed().addHeaderFooter('header'), en), it('Footer', () => ed().addHeaderFooter('footer'), en)]),
        it('Page numbers…', () => pageNumbers(ed()), en),
        it('Table of contents', () => { restore(); return ed().insertTableOfContents(); }, en),
        sub('Break', [it('Page break', () => { restore(); return ed().insertPageBreak(); }, { ...en, key: 'Ctrl+Enter' })]),
      );
      m.Format.push(
        sub('Text', [it('UPPERCASE', () => changeCase('upper'), en), it('lowercase', () => changeCase('lower'), en), it('Title Case', () => changeCase('title'), en), SEP,
          it('Increase font size', () => stepFontSize(1), { ...en, key: 'Ctrl+Shift+.' }), it('Decrease font size', () => stepFontSize(-1), { ...en, key: 'Ctrl+Shift+,' })]),
        sub('Line & paragraph spacing', [
          ...[[1, 'Single'], [1.15, '1.15'], [1.5, '1.5'], [2, 'Double']].map(([v, t]) => it(t, () => setLineSpacing(ed(), v), en)), SEP,
          it('Add space before paragraph', () => addSpace(ed(), 'before', true), en), it('Remove space before paragraph', () => addSpace(ed(), 'before', false), en),
          it('Add space after paragraph', () => addSpace(ed(), 'after', true), en), it('Remove space after paragraph', () => addSpace(ed(), 'after', false), en), SEP,
          it('Custom spacing…', () => customSpacing(ed()), en)]),
        sub('Table', [
          it('Insert row above', () => tableOp(ed(), 'rowAbove'), en), it('Insert row below', () => tableOp(ed(), 'rowBelow'), en),
          it('Insert column left', () => tableOp(ed(), 'colLeft'), en), it('Insert column right', () => tableOp(ed(), 'colRight'), en), SEP,
          it('Delete row', () => tableOp(ed(), 'deleteRow'), en), it('Delete column', () => tableOp(ed(), 'deleteColumn'), en), it('Delete table', () => tableOp(ed(), 'deleteTable'), en)]),
        it('Page setup…', () => pageSetup(ed()), en),
      );
      m.View.push(it('Show outline', () => toggleOutline(tab), { checked: () => !!tab.outlineEl }), it('Show paragraph marks', () => toggleMarks(tab), { checked: () => tab.body.classList.contains('show-marks') }));
      m.Tools.push(it('Comments…', () => commentsPanel(ed()), { enabled: () => !!tab.editor }), it('Compare documents…', () => compareDocuments(tab)), it('Preferences…', () => preferences()));
      if (typeof DocsExtra !== 'undefined') for (const [k2, list] of Object.entries(DocsExtra.wordMenus(tab))) m[k2] = [...(m[k2] || []), SEP, ...list];
    } else if (k === 'odt' || k === 'rtf') {
      m.Insert.push(it('Special characters…', () => specialCharacters(), en), it('Date', () => insertText(new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })), en), it('Checklist', () => toggleChecklist(), en));
      m.Format.push(sub('Text', [it('UPPERCASE', () => changeCase('upper'), en), it('lowercase', () => changeCase('lower'), en), it('Title Case', () => changeCase('title'), en), SEP,
        it('Increase font size', () => stepFontSize(1), en), it('Decrease font size', () => stepFontSize(-1), en)]));
      m.View.push(it('Show outline', () => toggleOutline(tab), { checked: () => !!tab.outlineEl }), it('Show paragraph marks', () => toggleMarks(tab), { checked: () => tab.body.classList.contains('show-marks') }));
      m.Tools.push(it('Compare documents…', () => compareDocuments(tab)), it('Preferences…', () => preferences()));
    } else if (k === 'md' || k === 'text' || k === 'html') {
      m.Insert.push(it('Special characters…', () => specialCharacters(), en), it('Date', () => insertText(new Date().toISOString().slice(0, 10)), en));
      m.Format.push(sub('Text', [it('UPPERCASE', () => changeCase('upper'), en), it('lowercase', () => changeCase('lower'), en), it('Title Case', () => changeCase('title'), en)]));
      if (k === 'md') m.View.push(it('Show outline', () => toggleOutline(tab), { checked: () => !!tab.outlineEl }));
      m.Tools.push(it('Compare documents…', () => compareDocuments(tab)));
    }
    return m;
  }

  // Keyboard shortcuts from Google Docs that the editors don't already have.
  document.addEventListener('keydown', (e) => {
    if (!active || !['docx', 'odt', 'rtf'].includes(active.kind) || active.mode !== 'edit' || !editableAt()) return;
    const k = e.key.toLowerCase();
    if (e.ctrlKey && e.shiftKey && (k === '.' || k === '>')) { e.preventDefault(); stepFontSize(1); }
    else if (e.ctrlKey && e.shiftKey && (k === ',' || k === '<')) { e.preventDefault(); stepFontSize(-1); }
    else if (e.ctrlKey && e.shiftKey && (e.code === 'Digit9')) { e.preventDefault(); toggleChecklist(); }
    else if (e.ctrlKey && !e.shiftKey && !e.altKey && k === 'enter' && active.kind === 'docx') { e.preventDefault(); run('insert a page break', () => active.editor.insertPageBreak())(); }
    else if (e.ctrlKey && e.altKey && k === 'f' && active.kind === 'docx') { e.preventDefault(); footnote(active.editor); }
    else if (e.ctrlKey && e.altKey && k === 'm' && active.kind === 'docx') { e.preventDefault(); comment(active.editor); }
    else if (e.ctrlKey && !e.shiftKey && (k === '\\' || e.code === 'Space') && !e.altKey) { e.preventDefault(); document.execCommand('removeFormat'); }
    else if (e.ctrlKey && !e.shiftKey && !e.altKey && k === '.') { e.preventDefault(); document.execCommand('superscript'); }
    else if (e.ctrlKey && !e.shiftKey && !e.altKey && k === ',') { e.preventDefault(); document.execCommand('subscript'); }
    else if (e.altKey && e.shiftKey && e.code === 'Digit5') { e.preventDefault(); document.execCommand('strikeThrough'); }
  });

  return { docMenus, specialCharacters, changeCase, toggleChecklist, paintFormat, stepFontSize, toggleOutline, toggleMarks, preferences, compareDocuments, tableGrid, restore, insertText, setLineSpacing, customSpacing, pageSetup, footnote, comment, commentsPanel, bookmark, pageNumbers, form, diff };
})();
