'use strict';
// Doc Studio renderer: tabs, editors (Word, Markdown, text/HTML/CSV), find, print and save.

const $ = (sel) => document.querySelector(sel);
// Logs a failed command, except guidance such as "Select the data first" (the dialog already says it; it is not a bug).
const HINT = /^(Select|Click|Double-click|Right-click|Copy|Create|Choose|Put|Type|Open|Add|Pick|Draw|Insert|There is|There are|No |Nothing)\b/;
function logError(err) { if (!(err && HINT.test(err.message || ''))) console.error(err); }
const tabs = [];
let active = null;
let tabSeq = 0;
let untitledSeq = 0;

const MODE_LABELS = { view: 'View', edit: 'Edit', split: 'Split', preview: 'Preview', table: 'Table' };

// ======================================================================
// Opening files
// ======================================================================

async function openPath(file) {
  const existing = tabs.find((t) => t.path && t.path.toLowerCase() === file.toLowerCase());
  if (existing) { activate(existing); return; }
  let f;
  try { f = await desk.read(file); } catch (err) { desk.error(`Couldn't open ${file.split(/[\\/]/).pop()}`, err.message); return; }
  const tab = makeTab({ path: f.path, name: f.name, size: f.size, data: f.data });
  try {
    await setupTab(tab);
  } catch (err) {
    console.error(err);
    tab.body.replaceChildren(failure(tab, err));
    tab.modes = [];
  }
  refreshChrome();
}

// Only one Open dialog at a time (a second click while it is showing does nothing).
let openDialogShowing = false;
async function openDialog() {
  if (openDialogShowing) return;
  openDialogShowing = true;
  let files = [];
  try { files = await desk.openDialog(); } finally { openDialogShowing = false; }
  for (const f of files) await openPath(f);
}

// Starting content for new text-based documents.
const STARTERS = {
  html: `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Untitled</title>
</head>
<body>
  <h1>Untitled</h1>
  <p></p>
</body>
</html>
`,
  csv: 'Column 1,Column 2,Column 3\n',
  json: '{\n  \n}\n',
};

// New files in formats Doc Studio edits through another editor (saved by converting on save).
const NEW_VIA = { ods: 'sheet', xls: 'sheet', odp: 'pptx', doc: 'docx' };

function newDocument(ext) {
  untitledSeq++;
  const tab = makeTab({ path: null, name: `Untitled ${untitledSeq}.${ext}`, size: 0, data: null });
  tab.isNew = true;
  if (NEW_VIA[ext]) tab.kind = NEW_VIA[ext];
  if (STARTERS[ext]) tab.data = new TextEncoder().encode(STARTERS[ext]).buffer;
  setupTab(tab).then(refreshChrome);
}

function makeTab({ path, name, size, data }) {
  const ext = FORMATS.extOf(name);
  const tab = {
    id: ++tabSeq, path, name, ext, size, data,
    kind: FORMATS.kindOf(name),
    dirty: false, mode: null, modes: [], tools: [], urls: [], status: '', info: '',
    el: h('section', { class: 'pane' }),
  };
  tab.body = h('div', { class: 'pane-body' });
  tab.el.append(tab.body);
  tab.setStatus = (s) => { tab.status = s; if (tab === active) refreshStatus(); };
  $('#stage').append(tab.el);
  tabs.push(tab);
  activate(tab);
  return tab;
}

async function setupTab(tab) {
  tab.body.replaceChildren(h('div', { class: 'loading', text: 'Opening…' }));
  let kind = tab.kind;
  if (kind === 'unknown' && tab.data && !TextCodec.looksBinary(tab.data)) kind = tab.kind = 'text';
  switch (kind) {
    case 'docx': return setupDocx(tab);
    case 'md': return setupTextual(tab, ['edit', 'split', 'preview'], tab.isNew ? 'edit' : 'preview');
    case 'html': return setupTextual(tab, ['preview', 'edit'], tab.isNew ? 'edit' : 'preview');
    case 'csv': return setupTextual(tab, ['table', 'edit'], tab.isNew ? 'edit' : 'table');
    case 'sheet': return setupSheet(tab);
    case 'pdf': return setupPdf(tab);
    case 'pptx': return setupPptx(tab);
    case 'ppt': return editConverted(tab, 'view');
    case 'odt': return setupRich(tab, 'odt');
    case 'rtf': return !tab.isNew && await desk.convertRoute('rtf', 'docx') ? setupConvertible(tab) : setupRich(tab, 'rtf');
    case 'doc': case 'odp': case 'sheetview': return setupConvertible(tab);
    case 'text': return setupTextual(tab, ['edit'], 'edit');
    case 'image': return setupImage(tab);
    case 'unknown':
      tab.body.replaceChildren(failure(tab, null));
      return;
    default: {
      const box = h('div', { class: 'view-box' });
      tab.body.replaceChildren(box);
      tab.info = (await Viewers[kind](tab, tab.data, box)) || '';
      tab.searchRoot = () => box;
      tab.printable = () => (['pdf', 'epub', 'audio', 'video'].includes(kind) ? null : box);
    }
  }
}

function failure(tab, err) {
  const box = h('div', { class: 'failure' },
    h('h2', { text: err ? `Doc Studio couldn't display this ${FORMATS.label(tab.kind).toLowerCase()}` : "Doc Studio can't display this kind of file" }),
    h('p', { text: err ? (err.message || String(err)) : `${tab.name} doesn't look like a document or text file.` }));
  if (tab.path) box.append(h('button', { class: 'ghost', text: 'Show in folder', onclick: () => desk.reveal(tab.path) }));
  return box;
}

// ======================================================================
// Word documents: View = paginated layout, Edit = the same rendering made editable (DocxEditor patches the
// original XML on save, so the document's own formatting is kept).
// ======================================================================

async function setupDocx(tab) {
  if (!tab.data) tab.data = await desk.docxBlank();
  tab.modes = ['view', 'edit'];
  tab.viewBox = h('div', { class: 'view-box' });
  tab.editHost = h('div');
  tab.editBox = h('div', { class: 'view-box docx-edit-box' }, tab.editHost);
  tab.body.replaceChildren(tab.viewBox, tab.editBox);
  tab.searchRoot = () => (tab.mode === 'edit' ? tab.editBox : tab.viewBox);
  tab.printable = () => tab.viewBox.querySelector('.docx-view');
  tab.beforePrint = async () => { if (!tab.viewRendered || tab.editedSinceView) await renderDocxView(tab); };
  tab.onMode = async (mode) => {
    tab.viewBox.hidden = mode !== 'view';
    tab.editBox.hidden = mode !== 'edit';
    if (mode === 'edit' && !tab.editor) await buildDocxEditor(tab);
    if (mode === 'view' && (!tab.viewRendered || tab.editedSinceView)) await renderDocxView(tab);
    if (mode === 'edit') { const a = tab.editHost.querySelector('article'); if (a && !tab.editHost.contains(document.activeElement)) a.focus(); }
    tab.tools = mode === 'edit' ? tab.editTools : [];
    updateDocxInfo(tab);
  };
  await setMode(tab, tab.startMode || (tab.isNew ? 'edit' : 'view'));
}

async function currentDocxBytes(tab) {
  return tab.editor && tab.everEdited ? (await tab.editor.build()).bytes : tab.data;
}

async function renderDocxView(tab) {
  const scroll = tab.viewBox.scrollTop;
  const box = h('div');
  const data = await currentDocxBytes(tab);
  const pages = await Viewers.docx(tab, data, box);
  tab.viewBox.replaceChildren(...(tab.editor && tab.everEdited && tab.dirty ? [h('div', { class: 'notice', text: 'Showing your edits - not saved yet.' })] : []), ...box.childNodes);
  // Page splitting needs real measurements, so it runs once the pages are on screen.
  const wasHidden = tab.viewBox.hidden;
  tab.viewBox.hidden = false;
  const view = tab.viewBox.querySelector('.docx-view');
  if (view) { Viewers.paginate(view); tab.viewInfo = `${view.querySelectorAll('section.docx').length} page(s)`; }
  tab.viewBox.hidden = wasHidden;
  tab.viewBox.scrollTop = scroll;
  if (!tab.viewInfo || !view) tab.viewInfo = pages;
  tab.viewRendered = true;
  tab.editedSinceView = false;
  if (typeof Shell !== 'undefined') Shell.pagesChanged(tab);
}

async function buildDocxEditor(tab) {
  tab.editHost.replaceChildren(h('div', { class: 'loading', text: 'Preparing the editor…' }));
  const editor = new DocxEditor.Editor(tab.editHost, {
    onChange: () => { markDirty(tab); tab.everEdited = true; tab.editedSinceView = true; updateDocxInfo(tab); },
    // After every (re)draw: picture handles and dropdown chips.
    onLoad: () => { if (typeof DocsExtra !== 'undefined') { DocsExtra.wirePictures(tab); DocsExtra.wireDropdowns(tab); } },
  });
  await editor.load(tab.data);
  tab.editor = editor;
  tab.editTools = docxToolbar(tab, editor);
  document.execCommand('styleWithCSS', false, true);
  tab.editHost.addEventListener('click', (e) => {
    const a = e.target.closest('a[href]');
    if (a && e.ctrlKey) { e.preventDefault(); desk.openExternal(a.getAttribute('href')); }
  });
}

function docxToolbar(tab, editor) {
  // Remember the selection inside the document so toolbar controls (which take focus) can act on it.
  let saved = null;
  const inDoc = (n) => {
    if (!n || !editor.stage || !editor.stage.contains(n)) return null;
    return (n.nodeType === 1 ? n : n.parentElement).closest('[contenteditable="true"]');
  };
  document.addEventListener('selectionchange', () => {
    const s = getSelection();
    if (s.rangeCount && inDoc(s.anchorNode)) { saved = s.getRangeAt(0).cloneRange(); if (tab === active && tab.mode === 'edit') refreshState(); }
  });
  const restore = () => {
    if (!saved || !editor.stage.contains(saved.startContainer)) return false;
    const art = inDoc(saved.startContainer);
    if (art) art.focus({ preventScroll: true });
    const s = getSelection(); s.removeAllRanges(); s.addRange(saved);
    return true;
  };
  const changed = () => { editor.opts.onChange(); refreshState(); };
  const cmd = (name, value) => () => { if (!restore()) return; document.execCommand(name, false, value); changed(); };
  const btn = (label, title, fn, state) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title, 'data-state': state || '' });
    b.innerHTML = label;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', fn);
    return b;
  };

  // Paragraph styles come from the document's own styles.xml: its quick styles plus any style the document uses.
  const used = new Set([...editor.stage.querySelectorAll('article p')].map((p) => editor.styleOf(p)));
  const styleList = editor.styles.paragraph.filter((st) => (st.quick && !st.hidden) || used.has(st.id))
    .sort((a, b) => (a.id === editor.styles.defaultPara ? -1 : b.id === editor.styles.defaultPara ? 1 : 0));
  const prettyName = (n) => n.replace(/^heading (\d)$/i, 'Heading $1').replace(/^title$/i, 'Title').replace(/^normal$/i, 'Normal').replace(/^subtitle$/i, 'Subtitle');
  const styleSel = h('select', { class: 'tool-select style-select', title: 'Paragraph style', 'aria-label': 'Paragraph style' },
    styleList.map((st) => h('option', { value: st.id, text: prettyName(st.name) })));
  styleSel.addEventListener('change', () => { restore(); editor.setParagraphStyle(editor.selectedParagraphs(), styleSel.value); });

  const fonts = ['Calibri', 'Cambria', 'Aptos', 'Arial', 'Georgia', 'Segoe UI', 'Times New Roman', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Garamond', 'Courier New', 'Consolas'];
  const fontSel = h('select', { class: 'tool-select font-select', title: 'Font', 'aria-label': 'Font' }, h('option', { value: '', text: 'Font' }), fonts.map((f) => h('option', { value: f, text: f })));
  fontSel.addEventListener('change', () => { if (fontSel.value && restore()) { document.execCommand('fontName', false, fontSel.value); changed(); } });
  const sizes = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
  const sizeSel = h('select', { class: 'tool-select size-select', title: 'Font size', 'aria-label': 'Font size' }, h('option', { value: '', text: 'Size' }), sizes.map((z) => h('option', { value: z, text: z })));
  sizeSel.addEventListener('change', () => {
    if (!sizeSel.value || !restore()) return;
    // execCommand only knows sizes 1-7: apply 7, then turn those into the chosen point size.
    document.execCommand('fontSize', false, '7');
    editor.stage.querySelectorAll('article [style*="xxx-large"], article font[size="7"]').forEach((el) => {
      if (el.tagName === 'FONT') { const sp = document.createElement('span'); sp.style.fontSize = `${sizeSel.value}pt`; while (el.firstChild) sp.append(el.firstChild); el.replaceWith(sp); }
      else el.style.fontSize = `${sizeSel.value}pt`;
    });
    changed();
  });

  const color = h('input', { type: 'color', class: 'tool-color', title: 'Text colour', value: '#c0392b' });
  color.addEventListener('change', () => { if (restore()) { document.execCommand('foreColor', false, color.value); changed(); } });
  const hilite = h('input', { type: 'color', class: 'tool-color', title: 'Highlight', value: '#fff176' });
  hilite.addEventListener('change', () => { if (restore()) { document.execCommand('hiliteColor', false, hilite.value); changed(); } });

  const structural = (fn) => async () => {
    restore();
    try { await fn(); } catch (err) { logError(err); desk.error("Couldn't change the document", err.message); }
  };

  const imageInput = h('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif,image/webp,image/bmp', hidden: true });
  imageInput.addEventListener('change', structural(async () => {
    const file = imageInput.files[0];
    imageInput.value = '';
    if (!file) return;
    const { bytes, ext, width, height } = await imageForWord(file);
    restore();
    await editor.insertImage(bytes, ext, width, height);
  }));

  const paint = btn(ICONS.paint, 'Paint format: copy formatting, then select text to apply it', () => DocTools.paintFormat(paint));
  const spacingBtn = btn(ICONS.spacing, 'Line & paragraph spacing', () => {
    const r = spacingBtn.getBoundingClientRect();
    const sp = editor.selectedParagraphs()[0];
    const cur = sp ? editor.spacingOf(sp).line : 1;
    SheetView.showMenu(r.left, r.bottom + 2, [
      ...[[1, 'Single'], [1.15, '1.15'], [1.5, '1.5'], [2, 'Double']].map(([v, t]) => [`${Math.abs(cur - v) < 0.03 ? '✓ ' : ''}${t}`, '', () => { restore(); editor.setSpacing(editor.selectedParagraphs(), { line: v }); }]),
      ['Add space before paragraph', '', () => { restore(); editor.setSpacing(editor.selectedParagraphs(), { before: 240 }); }],
      ['Add space after paragraph', '', () => { restore(); editor.setSpacing(editor.selectedParagraphs(), { after: 240 }); }],
      ['Remove space after paragraph', '', () => { restore(); editor.setSpacing(editor.selectedParagraphs(), { after: 0 }); }],
      ['Custom spacing…', '', () => DocTools.customSpacing(editor)],
    ]);
  });
  const tableBtn = btn(ICONS.table, 'Insert table', () => DocTools.tableGrid(tableBtn, structural(async (r, c) => { restore(); await editor.insertTable(r, c); })));
  const tools = [
    paint, sep(),
    styleSel, fontSel,
    btn('−', 'Decrease font size (Ctrl+Shift+,)', () => DocTools.stepFontSize(-1)), sizeSel, btn('+', 'Increase font size (Ctrl+Shift+.)', () => DocTools.stepFontSize(1)), sep(),
    btn('<b>B</b>', 'Bold (Ctrl+B)', cmd('bold'), 'bold'),
    btn('<i>I</i>', 'Italic (Ctrl+I)', cmd('italic'), 'italic'),
    btn('<u>U</u>', 'Underline (Ctrl+U)', cmd('underline'), 'underline'),
    btn('<s>S</s>', 'Strikethrough', cmd('strikeThrough'), 'strikeThrough'),
    btn('x<sup>2</sup>', 'Superscript', cmd('superscript'), 'superscript'),
    btn('x<sub>2</sub>', 'Subscript', cmd('subscript'), 'subscript'),
    color, hilite, sep(),
    spacingBtn,
    btn(ICONS.task, 'Checklist (Ctrl+Shift+9)', () => DocTools.toggleChecklist()),
    btn(ICONS.ul, 'Bulleted list', structural(() => editor.toggleList('bullet'))),
    btn(ICONS.ol, 'Numbered list', structural(() => editor.toggleList('number'))),
    btn(ICONS.outdent, 'Decrease indent', () => { restore(); editor.indent(editor.selectedParagraphs(), -1); }),
    btn(ICONS.indent, 'Increase indent', () => { restore(); editor.indent(editor.selectedParagraphs(), 1); }), sep(),
    btn(ICONS.left, 'Align left', cmd('justifyLeft'), 'justifyLeft'),
    btn(ICONS.center, 'Center', cmd('justifyCenter'), 'justifyCenter'),
    btn(ICONS.right, 'Align right', cmd('justifyRight'), 'justifyRight'),
    btn(ICONS.justify, 'Justify', cmd('justifyFull'), 'justifyFull'), sep(),
    btn(ICONS.link, 'Link (Ctrl+K)', () => { restore(); insertLink(); }),
    btn(ICONS.image, 'Insert picture', () => imageInput.click()),
    tableBtn,
    btn(ICONS.comment, 'Add comment (Ctrl+Alt+M)', () => DocTools.comment(editor)),
    btn(ICONS.pagebreak, 'Page break', structural(() => editor.insertPageBreak())),
    btn(ICONS.clear, 'Clear formatting', cmd('removeFormat')),
    imageInput,
  ];
  // Editing / Suggesting, as in Google Docs.
  const suggestSel = h('select', { class: 'tool-select suggest-select', title: 'Editing or suggesting', 'aria-label': 'Editing or suggesting' }, h('option', { value: 'edit', text: '✎ Editing' }), h('option', { value: 'suggest', text: '✍ Suggesting' }));
  suggestSel.addEventListener('change', () => DocsExtra.setSuggesting(tab, suggestSel.value === 'suggest'));
  tab.suggestSel = suggestSel;
  tools.push(sep(), suggestSel);

  function refreshState() {
    tools.forEach((t) => {
      const c = t.dataset && t.dataset.state;
      if (c) { try { t.classList.toggle('on', document.queryCommandState(c)); } catch { /* not a state command */ } }
    });
    const p = editor.caretParagraph();
    if (p) {
      const st = editor.styleOf(p);
      if (st && ![...styleSel.options].some((o) => o.value === st) && editor.styles.byId[st]) styleSel.append(h('option', { value: st, text: prettyName(editor.styles.byId[st].name) }));
      styleSel.value = st;
    }
    const sel = getSelection();
    const el = sel.anchorNode && (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement);
    if (el) {
      const cs = getComputedStyle(el);
      const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      fontSel.value = fonts.includes(fam) ? fam : '';
      fontSel.options[0].textContent = fonts.includes(fam) ? 'Font' : (fam || 'Font');
      const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
      sizeSel.value = sizes.includes(pt) ? String(pt) : '';
      sizeSel.options[0].textContent = sizes.includes(pt) ? 'Size' : String(pt || 'Size');
    }
  }
  tab.editHost.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key.toLowerCase() === 'k') { e.preventDefault(); insertLink(); }
  });
  return tools;
}

// Word takes PNG/JPEG/GIF/BMP pictures; anything else is converted to PNG first.
async function imageForWord(file) {
  const buf = await file.arrayBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: file.type }));
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    let ext = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' }[file.type];
    let bytes = buf;
    if (!ext) {
      const c = document.createElement('canvas');
      c.width = img.naturalWidth; c.height = img.naturalHeight;
      c.getContext('2d').drawImage(img, 0, 0);
      bytes = await (await new Promise((r) => c.toBlob(r, 'image/png'))).arrayBuffer();
      ext = 'png';
    }
    return { bytes, ext, width: img.naturalWidth, height: img.naturalHeight };
  } finally { URL.revokeObjectURL(url); }
}

async function insertLink() {
  const sel = getSelection();
  const range = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
  const host = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement);
  const editable = host && host.closest('[contenteditable="true"]');
  if (!editable) return;
  const existing = host.closest('a[href]');
  const v = await ask('Link', [
    { name: 'url', label: 'Address (leave empty to remove the link)', value: existing ? existing.getAttribute('href') : 'https://' },
    ...(range.collapsed && !existing ? [{ name: 'text', label: 'Text to show', value: '' }] : []),
  ]);
  editable.focus();
  sel.removeAllRanges(); sel.addRange(range);
  if (!v) return;
  const url = v.url.trim();
  if (!url || url === 'https://') {
    if (existing) { const r = document.createRange(); r.selectNodeContents(existing); sel.removeAllRanges(); sel.addRange(r); }
    document.execCommand('unlink');
    editable.dispatchEvent(new InputEvent('input'));
    return;
  }
  if (!/^(https?:|mailto:)/i.test(url)) { desk.error('Links must start with https://, http:// or mailto:'); return; }
  if (range.collapsed && !existing) {
    const text = v.text || url;
    document.execCommand('insertText', false, text);
    const r = sel.getRangeAt(0);
    r.setStart(r.startContainer, Math.max(0, r.startOffset - text.length));
    sel.removeAllRanges(); sel.addRange(r);
  }
  document.execCommand('createLink', false, url);
  editable.dispatchEvent(new InputEvent('input'));
}

function updateDocxInfo(tab) {
  if (tab.mode === 'edit' && tab.editor) tab.info = `${countWords(tab.editor.articles.map((a) => a.innerText).join('\n')).toLocaleString()} words`;
  else tab.info = tab.viewInfo || '';
  if (tab === active) refreshStatus();
}

// ======================================================================
// OpenDocument text and Rich Text: Doc Studio's own editors (ODT is patched in place; RTF is rewritten from the page).
// View shows the same page read-only.
// ======================================================================

async function setupRich(tab, which) {
  const M = which === 'odt' ? OdtEditor : RtfEditor;
  if (!tab.data) tab.data = await M.blank();
  tab.modes = ['view', 'edit'];
  const box = h('div', { class: `view-box rich-box ${which}-box` });
  tab.body.replaceChildren(box);
  const ed = new M.Editor(box, { onChange: () => { markDirty(tab); updateRichInfo(tab); } });
  await ed.load(tab.data);
  tab[which === 'odt' ? 'odtEditor' : 'rtfEditor'] = ed;
  tab.richEditor = ed;
  tab.richTools = richToolbar(tab, ed);
  tab.searchRoot = () => box;
  tab.printable = () => box;
  tab.onMode = (mode) => {
    ed.article.contentEditable = mode === 'edit' ? 'true' : 'false';
    tab.tools = mode === 'edit' ? tab.richTools : [];
    if (mode === 'edit' && !box.contains(document.activeElement)) ed.article.focus({ preventScroll: true });
    updateRichInfo(tab);
  };
  await setMode(tab, tab.startMode || (tab.isNew ? 'edit' : 'view'));
}

// Pictures: View (zoom) and Edit (image-editor.js). SVG and icons are view-only.
const IMAGE_EDITABLE = new Set(['png', 'jpg', 'jpeg', 'jfif', 'webp', 'bmp', 'gif', 'avif', 'apng']);
async function setupImage(tab) {
  const viewBox = h('div', { class: 'view-box' });
  const editBox = h('div', { class: 'view-box fill image-edit-box' });
  editBox.style.display = 'none';
  tab.body.replaceChildren(viewBox, editBox);
  tab.searchRoot = () => viewBox;
  tab.printable = () => viewBox;
  // View shows the picture as it is now (with unsaved edits).
  const showView = async () => {
    const data = tab.imageEditor && tab.dirty ? await tab.imageEditor.export('png') : tab.data;
    viewBox.replaceChildren();
    tab.info = (await Viewers.image(tab, data, viewBox)) || '';
  };
  if (!IMAGE_EDITABLE.has(tab.ext) || !tab.data) { await showView(); return; }
  tab.modes = ['view', 'edit'];
  tab.onMode = async (mode) => {
    viewBox.style.display = mode === 'view' ? '' : 'none';
    editBox.style.display = mode === 'edit' ? '' : 'none';
    if (mode !== 'edit') { await showView(); return; }
    if (!tab.imageEditor) {
      const ed = new ImageEditor.Editor(editBox, { onChange: () => markDirty(tab), onInfo: (s) => { tab.info = s; if (tab === active) refreshStatus(); } });
      await ed.load(tab.data, tab.ext);
      tab.imageEditor = ed;
    }
    tab.tools = tab.imageEditor.tools;
    tab.info = `${tab.imageEditor.width} × ${tab.imageEditor.height} px`;
    tab.imageEditor.fit();
    tab.imageEditor.root.focus({ preventScroll: true });
  };
  await setMode(tab, tab.startMode || (tab.isNew ? 'edit' : 'view'));
}

function updateRichInfo(tab) {
  const ed = tab.richEditor;
  if (!ed) return;
  tab.info = `${countWords(ed.article.innerText).toLocaleString()} words`;
  if (tab === active) refreshStatus();
}

function richToolbar(tab, ed) {
  let saved = null;
  document.addEventListener('selectionchange', () => {
    const s = getSelection();
    if (s.rangeCount && ed.article && ed.article.contains(s.anchorNode)) saved = s.getRangeAt(0).cloneRange();
  });
  const restore = () => {
    if (!saved || !ed.article.contains(saved.startContainer)) return false;
    ed.article.focus({ preventScroll: true });
    const s = getSelection(); s.removeAllRanges(); s.addRange(saved);
    return true;
  };
  const cmd = (name, value) => () => { if (!restore()) return; document.execCommand('styleWithCSS', false, true); document.execCommand(name, false, value); };
  const btn = (label, title, fn) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title });
    b.innerHTML = label;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', async () => { try { await fn(); } catch (err) { logError(err); desk.error("Couldn't change the document", err.message); } });
    return b;
  };
  const paras = () => { const ps = ed.selectedParagraphs(); return ps; };
  const tools = [];
  if (ed.paragraphStyles) {
    const list = ed.paragraphStyles();
    const sel = h('select', { class: 'tool-select style-select', title: 'Paragraph style' }, h('option', { value: '', text: 'Style' }), list.map((st) => h('option', { value: st.id, text: st.name })));
    sel.addEventListener('change', () => { if (sel.value && restore()) ed.setParagraphStyle(paras(), sel.value); sel.value = ''; });
    tools.push(sel);
  }
  const fonts = ['Calibri', 'Cambria', 'Aptos', 'Arial', 'Liberation Serif', 'Liberation Sans', 'Georgia', 'Segoe UI', 'Times New Roman', 'Verdana', 'Courier New'];
  const fontSel = h('select', { class: 'tool-select font-select', title: 'Font' }, h('option', { value: '', text: 'Font' }), fonts.map((f) => h('option', { value: f, text: f })));
  fontSel.addEventListener('change', () => { if (fontSel.value && restore()) { document.execCommand('styleWithCSS', false, true); document.execCommand('fontName', false, fontSel.value); } fontSel.value = ''; });
  const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
  const sizeSel = h('select', { class: 'tool-select size-select', title: 'Font size' }, h('option', { value: '', text: 'Size' }), sizes.map((z) => h('option', { value: z, text: z })));
  sizeSel.addEventListener('change', () => {
    if (!sizeSel.value || !restore()) return;
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontSize', false, '7');
    ed.article.querySelectorAll('[style*="xxx-large"], font[size="7"]').forEach((el) => {
      if (el.tagName === 'FONT') { const sp = document.createElement('span'); sp.style.fontSize = `${sizeSel.value}pt`; while (el.firstChild) sp.append(el.firstChild); el.replaceWith(sp); } else el.style.fontSize = `${sizeSel.value}pt`;
    });
    sizeSel.value = '';
    ed.article.dispatchEvent(new InputEvent('input'));
  });
  const color = h('input', { type: 'color', class: 'tool-color', title: 'Text colour', value: '#c0392b' });
  color.addEventListener('change', () => cmd('foreColor', color.value)());
  const hilite = h('input', { type: 'color', class: 'tool-color', title: 'Highlight', value: '#fff176' });
  hilite.addEventListener('change', () => cmd('hiliteColor', hilite.value)());
  const imageInput = h('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif,image/bmp,image/webp', hidden: true });
  imageInput.addEventListener('change', async () => {
    const file = imageInput.files[0];
    imageInput.value = '';
    if (!file) return;
    const { bytes, ext, width, height } = await imageForWord(file);
    restore();
    await ed.insertImage(bytes, ext === 'jpeg' ? 'jpg' : ext, width, height);
  });
  const align = (v) => () => { restore(); ed.align(paras(), v); };
  const paint = btn(ICONS.paint, 'Paint format: copy formatting, then select text to apply it', () => DocTools.paintFormat(paint));
  tools.unshift(paint, sep());
  tools.push(fontSel, btn('−', 'Decrease font size (Ctrl+Shift+,)', () => DocTools.stepFontSize(-1)), sizeSel, btn('+', 'Increase font size (Ctrl+Shift+.)', () => DocTools.stepFontSize(1)), sep(),
    btn('<b>B</b>', 'Bold (Ctrl+B)', cmd('bold')),
    btn('<i>I</i>', 'Italic (Ctrl+I)', cmd('italic')),
    btn('<u>U</u>', 'Underline (Ctrl+U)', cmd('underline')),
    btn('<s>S</s>', 'Strikethrough', cmd('strikeThrough')),
    btn('x<sup>2</sup>', 'Superscript', cmd('superscript')),
    btn('x<sub>2</sub>', 'Subscript', cmd('subscript')),
    color, hilite, sep(),
    btn(ICONS.left, 'Align left', align('left')),
    btn(ICONS.center, 'Center', align('center')),
    btn(ICONS.right, 'Align right', align('right')),
    btn(ICONS.justify, 'Justify', align('justify')), sep(),
    btn(ICONS.task, 'Checklist (Ctrl+Shift+9)', () => DocTools.toggleChecklist()),
    btn(ICONS.image, 'Insert picture', () => imageInput.click()),
    imageInput);
  return tools;
}

// ======================================================================
// Older and other-suite formats (.doc, .rtf, .odt, .ppt, .odp, .xls, .xlsb, .ods): View uses Doc Studio's own
// viewer; Edit converts the file to its modern equivalent with the office program installed on this PC (Word,
// Excel, PowerPoint or LibreOffice) and opens the full editor. Saving converts back to the original format.
// ======================================================================

const EDIT_AS = { doc: 'docx', rtf: 'docx', odt: 'docx', odp: 'pptx', ppt: 'pptx', sheetview: 'xlsx' };

async function setupConvertible(tab) {
  // Word and PowerPoint draw these far better than a text-only preview, so convert straight away when they can.
  if (['doc', 'odp'].includes(tab.kind) && await desk.convertRoute(tab.ext, EDIT_AS[tab.kind])) return editConverted(tab, 'view');
  const box = h('div', { class: 'view-box' });
  tab.body.replaceChildren(box);
  tab.viewEl = box;
  tab.modes = ['view', 'edit'];
  tab.mode = 'view';
  tab.searchRoot = () => box;
  tab.printable = () => box;
  tab.onMode = async (mode) => { if (mode === 'edit') await editConverted(tab, 'edit'); };
  tab.info = (await Viewers[tab.kind](tab, tab.data, box)) || '';
}

async function editConverted(tab, startMode) {
  const from = tab.ext;
  const to = EDIT_AS[tab.kind];
  const back = tab.viewEl;
  tab.body.replaceChildren(h('div', { class: 'loading', text: 'Converting for editing…' }));
  let data, via, note = '';
  try {
    const route = await desk.convertRoute(from, to);
    if (route) ({ data, via } = await desk.convert(tab.data, from, to));
    else if (tab.kind === 'sheetview') {
      // No Excel or LibreOffice: SheetJS keeps values, formulas, merges and number formats, not cell styling.
      const wb = XLSX.read(tab.data, { type: 'array', cellStyles: true, cellNF: true, cellFormula: true });
      data = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
      via = 'sheetjs';
      note = 'Microsoft Excel and LibreOffice are not installed, so cell colours and fonts from this file are not kept.';
    } else if (tab.kind === 'doc') {
      const r = await desk.docExtract(tab.data);
      const html = r.body.split(/\n{2,}|\r/).filter((x) => x.trim()).map((x) => `<p>${x.trim().replace(/&/g, '&amp;').replace(/</g, '&lt;')}</p>`).join('');
      data = await desk.docxFromHtml(html || '<p></p>', stem(tab.name));
      via = 'text';
      note = 'Microsoft Word and LibreOffice are not installed, so only the text of this file could be opened. It will be saved as a new .docx file.';
    } else {
      throw new Error(`Editing ${FORMATS.label(tab.kind)} files needs Microsoft Office or LibreOffice installed on this PC.`);
    }
  } catch (err) {
    if (back) tab.body.replaceChildren(back);
    tab.mode = 'view';
    if (!back) { tab.body.replaceChildren(failure(tab, err)); tab.modes = []; return; }
    throw err;
  }
  tab.original = { ext: from, kind: tab.kind, via };
  tab.kind = FORMATS.kindOf(`x.${to}`);
  tab.data = data;
  tab.modes = [];
  tab.onMode = null;
  tab.viewEl = null;
  tab.startMode = startMode;
  tab.body.replaceChildren();
  await setupTab(tab);
  if (note) desk.error('Some formatting is not available', note);
  else if (startMode === 'edit') flash(`Opened for editing with ${{ word: 'Microsoft Word', excel: 'Microsoft Excel', powerpoint: 'Microsoft PowerPoint', libreoffice: 'LibreOffice' }[via] || via}. Saving writes .${from} again.`);
}

// ======================================================================
// PDF: View = Chromium's viewer, Edit = PdfEditor (annotations, text, signatures, forms, pages)
// ======================================================================

async function setupPdf(tab) {
  if (!tab.data) tab.data = (await PdfEditor.blank()).buffer;
  tab.modes = ['view', 'edit'];
  tab.editHost = h('div', { class: 'pdf-edit-host' });
  tab.body.replaceChildren(tab.editHost);
  tab.searchRoot = () => (tab.pdfEditor ? tab.pdfEditor.main : null);
  // Printing draws every page as a picture first.
  tab.printBox = h('div', { class: 'pdf-print' });
  tab.printable = () => (tab.printBox.childElementCount ? tab.printBox : null);
  tab.beforePrint = async () => { if (tab.pdfEditor) await tab.pdfEditor.printImages(tab.printBox); };
  tab.onMode = async (mode) => {
    if (!tab.pdfEditor) await buildPdfEditor(tab);
    tab.pdfEditor.setReadOnly(mode === 'view');
    tab.tools = mode === 'edit' ? tab.pdfTools : [];
    if (tab === active) refreshChrome();
  };
  await setMode(tab, tab.startMode || (tab.isNew ? 'edit' : 'view'));
}

async function buildPdfEditor(tab) {
  tab.editHost.replaceChildren(h('div', { class: 'loading', text: 'Preparing the editor…' }));
  const ed = new PdfEditor.Editor(tab.editHost, {
    onChange: () => { markDirty(tab); tab.viewBytes = null; },
    ask: (title, fields) => ask(title, fields),
    error: (msg) => desk.error(msg),
    readFont: (file) => desk.readFont(file),
    pickPdf: async () => { const files = await desk.openDialog(); const f = files.find((x) => /\.pdf$/i.test(x)); if (!f) return null; return (await desk.read(f)).data; },
    onToolChange: (t) => tab.setPdfTool && tab.setPdfTool(t),
  });
  await ed.load(tab.data);
  if (ed.encrypted) desk.error('This PDF is protected', 'Doc Studio can show it, but changes may not save correctly while it is encrypted.');
  tab.pdfEditor = ed;
  tab.pdfTools = pdfToolbar(tab, ed);
  tab.info = `${ed.pages.length} page(s)${ed.fields.length ? ` · ${ed.fields.length} form field(s)` : ''}`;
  requestAnimationFrame(() => ed.fitWidth());
}

function pdfToolbar(tab, ed) {
  const toolBtns = [];
  const tool = (id, label, title) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title, 'data-tool': id });
    b.innerHTML = label;
    b.addEventListener('click', () => setTool(id));
    toolBtns.push(b);
    return b;
  };
  const setTool = (id) => { ed.setTool(id); toolBtns.forEach((b) => b.classList.toggle('on', b.dataset.tool === id)); };
  tab.setPdfTool = setTool;
  const act = (label, title, fn) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title });
    b.innerHTML = label;
    b.addEventListener('click', async () => { try { await fn(); } catch (err) { logError(err); desk.error("Couldn't do that", err.message); } });
    return b;
  };
  const color = h('input', { type: 'color', class: 'tool-color', title: 'Ink and text colour', value: ed.color });
  color.addEventListener('input', () => { ed.color = color.value; });
  const hi = h('input', { type: 'color', class: 'tool-color', title: 'Highlight colour', value: ed.hiColor });
  hi.addEventListener('input', () => { ed.hiColor = hi.value; });
  const font = h('select', { class: 'tool-select font-select', title: 'Font for new text' }, PdfEditor.FONT_CHOICES.map((f) => h('option', { value: f, text: f })));
  font.addEventListener('change', () => { ed.font = font.value; });
  const size = h('select', { class: 'tool-select size-select', title: 'Text size' }, [8, 9, 10, 11, 12, 14, 16, 18, 24, 32, 48].map((z) => h('option', { value: z, text: z })));
  size.value = '12';
  size.addEventListener('change', () => { ed.fontSize = +size.value; });
  const width = h('select', { class: 'tool-select size-select', title: 'Line width' }, [1, 2, 3, 5, 8].map((z) => h('option', { value: z, text: `${z} pt` })));
  width.value = '2';
  width.addEventListener('change', () => { ed.lineWidth = +width.value; });
  const imageInput = h('input', { type: 'file', accept: 'image/*', hidden: true });
  imageInput.addEventListener('change', async () => {
    const file = imageInput.files[0];
    imageInput.value = '';
    if (!file) return;
    const png = await toPng(file);
    ed.beginPlace(png.bytes, png.width * 0.75, png.height * 0.75, 'image');
    flash('Click on the page where the picture should go');
  });
  const tools = [
    tool('select', ICONS.pointer, 'Select and move (Delete removes the selected item)'),
    tool('text', 'T', 'Add text: click where it should go'),
    tool('edittext', ICONS.edittext, 'Edit existing text: click a line of text'),
    tool('highlight', ICONS.highlighter, 'Highlight: select text'),
    tool('highlightbox', ICONS.hibox, 'Highlight an area: drag a box'),
    tool('draw', ICONS.pen, 'Draw freehand'),
    tool('rect', ICONS.square, 'Rectangle'),
    tool('ellipse', ICONS.circle, 'Ellipse'),
    tool('line', ICONS.rule, 'Line'),
    tool('whiteout', ICONS.eraser, 'White-out: cover part of the page'),
    tool('note', ICONS.note, 'Sticky note: click where it should go'),
    act(ICONS.signature, 'Add a signature', async () => { const s = await PdfEditor.signatureDialog(); if (s) { ed.beginPlace(s.png, s.w, s.h, 'signature'); flash('Click on the page where the signature should go'); } }),
    act(ICONS.image, 'Add a picture', () => imageInput.click()),
    sep(), color, hi, font, size, width, sep(),
    act(ICONS.rotateLeft, 'Rotate page left', () => ed.rotatePage(ed.current || 0, -90)),
    act(ICONS.rotateRight, 'Rotate page right', () => ed.rotatePage(ed.current || 0, 90)),
    act(ICONS.pageAdd, 'Insert a blank page after this one', () => ed.insertBlank((ed.current || 0) + 1)),
    act(ICONS.pagesFrom, 'Insert pages from another PDF after this page', async () => { const b = await ed.opts.pickPdf(); if (b) await ed.insertFrom(b, (ed.current || 0) + 1); }),
    act(ICONS.pageDel, 'Delete this page', async () => {
      const r = await desk.confirm({ type: 'warning', message: `Delete page ${(ed.current || 0) + 1}?`, buttons: ['Delete', 'Cancel'], defaultId: 1, cancelId: 1 });
      if (r === 0) await ed.deletePage(ed.current || 0);
    }), sep(),
    act('−', 'Zoom out', () => ed.setZoom(ed.scale / 1.2)),
    act('+', 'Zoom in', () => ed.setZoom(ed.scale * 1.2)),
    act('Fit', 'Fit width', () => ed.fitWidth()),
    imageInput,
  ];
  setTool('select');
  return tools;
}

// ======================================================================
// PowerPoint: View = every slide drawn by Doc Studio's own renderer, Edit = PptxUI (text, shapes, pictures,
// slides, notes). Saving rewrites only the parts that changed.
// ======================================================================

async function setupPptx(tab) {
  if (!tab.data) tab.data = await PptxUI.blank();
  tab.modes = ['view', 'edit'];
  tab.viewBox = h('div', { class: 'view-box pp-view' });
  tab.editHost = h('div', { class: 'pp-edit-host' });
  tab.body.replaceChildren(tab.viewBox, tab.editHost);
  tab.searchRoot = () => (tab.mode === 'edit' && tab.pptEditor ? tab.pptEditor.stage : tab.viewBox);
  // Printing always prints every slide (drawn fresh, with any edits).
  tab.beforePrint = async () => { if (tab.pptEditor) tab.pptEditor.commitText(); if (!tab.viewRendered || tab.editedSinceView) await renderPptxView(tab); };
  tab.printable = () => tab.viewBox;
  tab.menuExtras = () => PptxUI.menus(tab, () => tab.pptEditor);
  tab.onMode = async (mode) => {
    tab.viewBox.hidden = mode !== 'view';
    tab.editHost.hidden = mode !== 'edit';
    if (mode === 'edit' && !tab.pptEditor) await buildPptxEditor(tab);
    if (mode === 'view') {
      if (tab.pptEditor) tab.pptEditor.commitText();
      if (!tab.viewRendered || tab.editedSinceView) await renderPptxView(tab);
    }
    if (mode === 'edit') tab.pptEditor.renderCurrent();
    tab.tools = mode === 'edit' ? tab.pptTools : [];
    if (tab === active) refreshChrome();
  };
  await setMode(tab, tab.startMode || (tab.isNew ? 'edit' : 'view'));
}

async function renderPptxView(tab) {
  const deck = tab.pptEditor ? tab.pptEditor.deck : await new PptxEditor.Deck().load(tab.data);
  const width = Math.min(1100, Math.max(400, (tab.body.clientWidth || 1000) - 64));
  const out = [];
  if (tab.pptEditor && tab.dirty) out.push(h('div', { class: 'notice', text: 'Showing your edits - not saved yet.' }));
  for (const [i, s] of deck.slides.entries()) {
    try { out.push(await new PptxEditor.Renderer(deck, { slideNumber: i + 1 }).slide(s, width, false)); }
    catch (err) { console.error(err); out.push(h('div', { class: 'notice', text: `Slide ${i + 1} couldn't be drawn: ${err.message}` })); }
  }
  tab.viewBox.replaceChildren(...out);
  tab.info = `${deck.slides.length} slide(s)`;
  tab.viewRendered = true;
  tab.editedSinceView = false;
  if (typeof Shell !== 'undefined') Shell.pagesChanged(tab);
}

async function buildPptxEditor(tab) {
  tab.editHost.replaceChildren(h('div', { class: 'loading', text: 'Preparing the editor…' }));
  const ed = new PptxUI.Editor(tab.editHost, {
    onChange: () => { markDirty(tab); tab.editedSinceView = true; tab.info = `Slide ${ed.current + 1} of ${ed.deck.slides.length}`; if (tab === active) refreshStatus(); },
    error: (msg) => desk.error(msg),
  });
  await ed.load(tab.data);
  tab.pptEditor = ed;
  tab.pptTools = pptxToolbar(tab, ed);
}

function pptxToolbar(tab, ed) {
  const act = (label, title, fn) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title });
    b.innerHTML = label;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', async () => { try { await fn(); } catch (err) { logError(err); desk.error("Couldn't change the presentation", err.message); } });
    return b;
  };
  const fmt = (cmd, value) => () => { if (!ed.format(cmd, value)) flash('Double-click a text box to edit its text first.'); };
  const fonts = ['Calibri', 'Calibri Light', 'Aptos', 'Arial', 'Segoe UI', 'Georgia', 'Times New Roman', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Century Gothic', 'Courier New'];
  const fontSel = h('select', { class: 'tool-select font-select', title: 'Font' }, h('option', { value: '', text: 'Font' }), fonts.map((f) => h('option', { value: f, text: f })));
  fontSel.addEventListener('mousedown', () => { ed.keepSel = getSelection().rangeCount ? getSelection().getRangeAt(0).cloneRange() : null; });
  const restoreSel = () => { if (ed.keepSel && ed.editingEl) { ed.editingEl.focus(); const s = getSelection(); s.removeAllRanges(); s.addRange(ed.keepSel); } };
  fontSel.addEventListener('change', () => { restoreSel(); if (fontSel.value) fmt('fontName', fontSel.value)(); fontSel.value = ''; });
  const sizes = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 54, 60, 72, 96];
  const sizeSel = h('select', { class: 'tool-select size-select', title: 'Font size' }, h('option', { value: '', text: 'Size' }), sizes.map((z) => h('option', { value: z, text: z })));
  sizeSel.addEventListener('mousedown', () => { ed.keepSel = getSelection().rangeCount ? getSelection().getRangeAt(0).cloneRange() : null; });
  sizeSel.addEventListener('change', () => { restoreSel(); if (sizeSel.value) fmt('size', +sizeSel.value)(); sizeSel.value = ''; });
  const color = h('input', { type: 'color', class: 'tool-color', title: 'Text colour', value: '#c0392b' });
  color.addEventListener('mousedown', () => { ed.keepSel = getSelection().rangeCount ? getSelection().getRangeAt(0).cloneRange() : null; });
  color.addEventListener('change', () => { restoreSel(); fmt('foreColor', color.value)(); });
  const imageInput = h('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif,image/webp,image/bmp', hidden: true });
  imageInput.addEventListener('change', async () => {
    const file = imageInput.files[0];
    imageInput.value = '';
    if (!file) return;
    try {
      const { bytes, ext, width, height } = await imageForWord(file);
      await ed.insertPicture(bytes, ext === 'jpeg' ? 'jpeg' : ext, width, height);
    } catch (err) { desk.error("Couldn't insert the picture", err.message); }
  });
  const shapeMenu = act(ICONS.square, 'Insert a shape', () => {
    const r = shapeMenu.getBoundingClientRect();
    SheetView.showMenu(r.left, r.bottom + 2, [
      ['Rectangle', '', () => ed.insertShape('rect')],
      ['Rounded rectangle', '', () => ed.insertShape('roundRect')],
      ['Oval', '', () => ed.insertShape('ellipse')],
      ['Triangle', '', () => ed.insertShape('triangle')],
      ['Right arrow', '', () => ed.insertShape('rightArrow')],
      ['Star', '', () => ed.insertShape('star5')],
    ]);
  });
  const present = h('button', { class: 'tool present-btn', title: 'Present (Ctrl+F5)', text: '▶ Present' });
  present.addEventListener('click', () => PptxUI.present(tab, {}).catch((err) => desk.error("Couldn't start the slideshow", err.message)));
  const shapeMenuBtn = (label, title, items) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title });
    b.innerHTML = label;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => { const r = b.getBoundingClientRect(); SheetView.showMenu(r.left, r.bottom + 2, items.map(([t, fn]) => [t, '', () => Promise.resolve(fn()).catch((err) => desk.error("Couldn't do that", err.message))])); });
    return b;
  };
  return [
    present, sep(),
    act(ICONS.pageAdd, 'New slide (choose a layout)', () => ed.addSlideDialog()),
    act(ICONS.copy || '⧉', 'Duplicate this slide', () => ed.duplicateSlide(ed.current)),
    act(ICONS.pageDel, 'Delete this slide', async () => {
      const r = await desk.confirm({ type: 'warning', message: `Delete slide ${ed.current + 1}?`, buttons: ['Delete', 'Cancel'], defaultId: 1, cancelId: 1 });
      if (r === 0) await ed.deleteSlide(ed.current);
    }), sep(),
    act('T', 'Insert a text box', () => ed.insertTextBox()),
    shapeMenu,
    act(ICONS.image, 'Insert a picture', () => imageInput.click()),
    shapeMenuBtn(ICONS.rule, 'Line or arrow', [['Line', () => ed.insertLine(false)], ['Arrow', () => ed.insertLine(true)]]),
    act(ICONS.table, 'Insert table', () => new Promise((resolve) => DocTools.tableGrid(document.querySelector('#tool-left'), async (r, c) => { await ed.insertTable(r, c); resolve(); }))),
    act(ICONS.trash || '✕', 'Delete the selected item (Delete)', () => ed.deleteSelected()), sep(),
    shapeMenuBtn(ICONS.fill, 'Fill colour', [['Blue', () => ed.setFill('#4472C4')], ['Orange', () => ed.setFill('#ED7D31')], ['Green', () => ed.setFill('#70AD47')], ['Red', () => ed.setFill('#E53935')], ['Yellow', () => ed.setFill('#FFC000')], ['White', () => ed.setFill('#FFFFFF')], ['Black', () => ed.setFill('#000000')], ['No fill', () => ed.setFill(null)]]),
    shapeMenuBtn(ICONS.border, 'Border', [['Black', () => ed.setBorder({ color: '#000000' })], ['Blue', () => ed.setBorder({ color: '#4472C4' })], ['1 pt', () => ed.setBorder({ width: 1 })], ['3 pt', () => ed.setBorder({ width: 3 })], ['6 pt', () => ed.setBorder({ width: 6 })], ['Dashed', () => ed.setBorder({ dash: 'dash' })], ['Solid', () => ed.setBorder({ dash: 'solid' })], ['No border', () => ed.setBorder({ none: true })]]),
    shapeMenuBtn(ICONS.arrange, 'Arrange', [['Bring to front', () => ed.order('front')], ['Send to back', () => ed.order('back')], ['Align left', () => ed.align('left')], ['Align center', () => ed.align('center')], ['Align right', () => ed.align('right')], ['Align top', () => ed.align('top')], ['Align middle', () => ed.align('middle')], ['Align bottom', () => ed.align('bottom')], ['Group (Shift+click several shapes first)', () => ed.group()], ['Ungroup', () => ed.ungroup()]]),
    shapeMenuBtn('Background', 'Slide background', [['White', () => ed.setBackground('#FFFFFF')], ['Light grey', () => ed.setBackground('#F2F2F2')], ['Light blue', () => ed.setBackground('#DEEBF7')], ['Dark', () => ed.setBackground('#1F2937')], ['Custom…', async () => { const v = await DocTools.form('Slide background', [{ name: 'c', label: 'Colour', type: 'color', value: '#ffffff' }, { name: 'all', label: 'Apply to all slides', type: 'checkbox', value: false }], 'Apply'); if (v) await ed.setBackground(v.c, v.all); }], ['Reset', () => ed.setBackground(null)]]),
    act('Layout', 'Apply a layout to this slide', async () => { const ls = await ed.layouts(); const r = document.querySelector('#tool-left').getBoundingClientRect(); SheetView.showMenu(r.left + 300, r.bottom + 2, ls.map((l) => [l.name, '', () => ed.applyLayout(l.path).catch((err) => desk.error("Couldn't apply the layout", err.message))])); }),
    shapeMenuBtn('Theme', 'Theme colours and fonts', ed.themeNames().map((n) => [n, () => ed.applyTheme(n)])),
    shapeMenuBtn('Transition', 'Slide transition', [['None', () => ed.setTransition('none')], ['Fade', () => ed.setTransition('fade', 'med')], ['Push', () => ed.setTransition('push', 'med')], ['Wipe', () => ed.setTransition('wipe', 'med')], ['Cover', () => ed.setTransition('cover', 'med')], ['Split', () => ed.setTransition('split', 'med')], ['Zoom', () => ed.setTransition('zoom', 'med')], ['Fade on every slide', () => ed.setTransition('fade', 'med', true)]]),
    shapeMenuBtn('Animate', 'Animate the selected shape', [['Appear', () => ed.setAnimation('appear')], ['Fade in', () => ed.setAnimation('fade')], ['Fly in from bottom', () => ed.setAnimation('fly')], ['Remove animation', () => ed.setAnimation(null)]]), sep(),
    fontSel, sizeSel,
    act('<b>B</b>', 'Bold (Ctrl+B)', fmt('bold')),
    act('<i>I</i>', 'Italic (Ctrl+I)', fmt('italic')),
    act('<u>U</u>', 'Underline (Ctrl+U)', fmt('underline')),
    act('<s>S</s>', 'Strikethrough', fmt('strikeThrough')),
    color, sep(),
    act(ICONS.left, 'Align left', fmt('align', 'left')),
    act(ICONS.center, 'Center', fmt('align', 'center')),
    act(ICONS.right, 'Align right', fmt('align', 'right')),
    imageInput,
  ];
}

// Any image -> PNG bytes (pdf-lib embeds PNG/JPEG; PNG keeps transparency).
async function toPng(file) {
  const url = URL.createObjectURL(file);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    const c = document.createElement('canvas');
    c.width = img.naturalWidth; c.height = img.naturalHeight;
    c.getContext('2d').drawImage(img, 0, 0);
    const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
    return { bytes: new Uint8Array(await blob.arrayBuffer()), width: img.naturalWidth, height: img.naturalHeight };
  } finally { URL.revokeObjectURL(url); }
}

// ======================================================================
// Excel workbooks: SheetModel (file in/out, styles), SheetCalc (formulas), SheetView (grid UI)
// ======================================================================

async function setupSheet(tab) {
  const data = tab.data || await SheetModel.blankBytes();
  if (!tab.data) tab.data = data;
  const wb = await new SheetModel.Workbook().load(data);
  const host = h('div', { class: 'sheet-host' });
  tab.body.replaceChildren(host);
  const view = new SheetView.View(host, wb, {
    onChange: () => markDirty(tab),
    onSelection: () => { updateSheetInfo(tab); if (tab.refreshSheetTools) tab.refreshSheetTools(); },
    onStatus: (msg) => flash(msg),
    onInitialCalc: (n) => { tab.status = `${n} formula value(s) calculated on open`; },
    ask: (title, fields) => ask(title, fields),
    error: (msg, detail) => desk.error(msg, detail),
    readClipboard: () => desk.readClipboard(),
    onWorkbook: (wb) => { tab.workbook = wb; updateSheetInfo(tab); },
  });
  tab.workbook = wb;
  tab.sheetView = view;
  tab.modes = [];
  tab.menuExtras = () => SheetView.sheetMenus(tab, view);
  tab.tools = sheetToolbar(tab, view);
  tab.printable = () => view.printable();
  tab.onActivate = () => setTimeout(() => { view.render(true); view.grid.focus({ preventScroll: true }); }, 0);
  updateSheetInfo(tab);
  if (tab === active) tab.onActivate();
}

function updateSheetInfo(tab) {
  const v = tab.sheetView;
  if (!v) return;
  const stats = v.selectionStats();
  tab.info = stats || `${tab.workbook.sheets.length} sheet(s)`;
  if (tab === active) refreshStatus();
}

function sheetToolbar(tab, view) {
  const focusGrid = () => view.grid.focus({ preventScroll: true });
  // Errors (including from async actions) are shown, not lost.
  const run = (fn) => () => {
    if (view.editing) view.endEdit(true);
    Promise.resolve().then(fn).catch((err) => { logError(err); desk.error("Couldn't do that", err.message); }).finally(() => { if (!document.querySelector('.backdrop')) focusGrid(); });
  };
  const btn = (label, title, fn, state) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title, 'data-state': state || '' });
    b.innerHTML = label;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', run(fn));
    return b;
  };
  const fonts = ['Calibri', 'Aptos', 'Arial', 'Cambria', 'Georgia', 'Segoe UI', 'Times New Roman', 'Verdana', 'Tahoma', 'Courier New', 'Consolas'];
  const fontSel = h('select', { class: 'tool-select font-select', title: 'Font', 'aria-label': 'Font' }, h('option', { value: '', text: 'Font' }), fonts.map((f) => h('option', { value: f, text: f })));
  fontSel.addEventListener('change', () => { if (fontSel.value) run(() => view.styleSelection({ font: { name: fontSel.value } }))(); });
  const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
  const sizeSel = h('select', { class: 'tool-select size-select', title: 'Font size', 'aria-label': 'Font size' }, h('option', { value: '', text: '' }), sizes.map((z) => h('option', { value: z, text: z })));
  sizeSel.addEventListener('change', () => { if (sizeSel.value) run(() => view.styleSelection({ font: { size: +sizeSel.value } }))(); });
  const color = h('input', { type: 'color', class: 'tool-color', title: 'Font colour', value: '#c00000' });
  color.addEventListener('change', run(() => view.styleSelection({ font: { color: color.value } })));
  const fill = h('input', { type: 'color', class: 'tool-color', title: 'Fill colour', value: '#ffff00' });
  fill.addEventListener('change', run(() => view.styleSelection({ fill: fill.value })));
  const borders = h('select', { class: 'tool-select', title: 'Borders', 'aria-label': 'Borders' },
    [['', 'Borders'], ['all', 'All borders'], ['outside', 'Outside borders'], ['bottom', 'Bottom border'], ['none', 'No border']].map(([v, l]) => h('option', { value: v, text: l })));
  borders.addEventListener('change', () => { const v = borders.value; borders.value = ''; if (v) run(() => view.borders(v))(); });
  const FORMATS_LIST = [
    ['General', 'General'], ['0', 'Number (0)'], ['#,##0.00', 'Number (1,234.00)'], ['\\$#,##0.00', 'Currency'],
    ['_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', 'Accounting'], ['0%', 'Percent'], ['0.00%', 'Percent (0.00%)'],
    ['m/d/yyyy', 'Short date'], ['dddd, mmmm d, yyyy', 'Long date'], ['h:mm:ss AM/PM', 'Time'], ['0.00E+00', 'Scientific'], ['# ?/?', 'Fraction'], ['@', 'Text'],
  ];
  const numSel = h('select', { class: 'tool-select num-select', title: 'Number format', 'aria-label': 'Number format' }, FORMATS_LIST.map(([v, l]) => h('option', { value: v, text: l })), h('option', { value: '__custom', text: 'Custom', hidden: true }));
  numSel.addEventListener('change', () => { if (numSel.value !== '__custom') run(() => view.styleSelection({ numFmt: numSel.value }))(); });
  const wrapBtn = btn(ICONS.wrap, 'Wrap text', () => view.styleSelection({ align: { wrap: !view.activeCss().wrap } }));
  const align = (hz) => () => view.styleSelection({ align: { horizontal: view.activeCss().h === hz ? null : hz } });
  // A compact dropdown that runs an action and resets.
  function menuSelect(label, options, fn) {
    const sel = h('select', { class: 'tool-select menu-select', title: label, 'aria-label': label }, h('option', { value: '', text: label }), options.map(([v, l]) => h('option', { value: v, text: l })));
    sel.addEventListener('change', () => {
      const v = sel.value;
      sel.value = '';
      if (!v) return;
      if (view.editing) view.endEdit(true);
      Promise.resolve(fn(v)).catch((err) => { logError(err); desk.error("Couldn't do that", err.message); }).finally(() => focusGrid());
    });
    return sel;
  }
  const paint = h('button', { class: 'tool', title: 'Paint format: copy this cell\'s format, then select cells to apply it', 'aria-label': 'Paint format' });
  paint.innerHTML = ICONS.paint;
  paint.addEventListener('mousedown', (e) => e.preventDefault());
  paint.addEventListener('click', () => { if (view.editing) view.endEdit(true); view.startPaintFormat(paint); focusGrid(); });
  const stepSize = (d) => () => { const cur = view.activeCss().font_.size || 11; const next = d > 0 ? sizes.find((z) => z > cur) || cur + 2 : [...sizes].reverse().find((z) => z < cur) || Math.max(1, cur - 1); view.styleSelection({ font: { size: next } }); };
  const fnMenu = menuSelect('Σ', [['SUM', 'SUM'], ['AVERAGE', 'AVERAGE'], ['COUNT', 'COUNT'], ['MAX', 'MAX'], ['MIN', 'MIN'], ['more', 'More functions…']], (v) => {
    if (v === 'more') return view.functionPicker();
    if (v === 'SUM') return view.autoSum();
    view.beginEdit(`=${v}(`, 'cell', true);
  });
  fnMenu.title = 'Functions';
  const tools = [
    paint, sep(),
    btn('$', 'Format as currency', () => view.styleSelection({ numFmt: '\\$#,##0.00' })),
    btn('%', 'Format as percent', () => view.styleSelection({ numFmt: '0.00%' })), sep(),
    fontSel, btn('−', 'Decrease font size', stepSize(-1)), sizeSel, btn('+', 'Increase font size', stepSize(1)), sep(),
    btn('<b>B</b>', 'Bold (Ctrl+B)', () => view.toggleFont('bold'), 'bold'),
    btn('<i>I</i>', 'Italic (Ctrl+I)', () => view.toggleFont('italic'), 'italic'),
    btn('<u>U</u>', 'Underline (Ctrl+U)', () => view.toggleFont('underline'), 'underline'),
    btn('<s>S</s>', 'Strikethrough', () => view.toggleFont('strike'), 'strike'),
    color, fill, borders, sep(),
    btn(ICONS.left, 'Align left', align('left'), 'h-left'),
    btn(ICONS.center, 'Center', align('center'), 'h-center'),
    btn(ICONS.right, 'Align right', align('right'), 'h-right'),
    menuSelect('Vertical align', [['top', 'Top'], ['center', 'Middle'], ['bottom', 'Bottom']], (v) => view.styleSelection({ align: { vertical: v } })),
    wrapBtn,
    menuSelect('Rotation', [['0', 'None'], ['45', 'Tilt up'], ['135', 'Tilt down'], ['90', 'Rotate up'], ['180', 'Rotate down'], ['255', 'Stack vertically']], (v) => view.styleSelection({ align: { textRotation: v === '0' ? null : v } })),
    btn(ICONS.merge, 'Merge & center (click again to unmerge)', () => view.toggleMerge()), sep(),
    btn(ICONS.link, 'Insert link (Ctrl+K)', () => view.insertLink()),
    btn(ICONS.comment, 'Insert note (Shift+F2)', () => view.editNote()),
    btn(ICONS.chart, 'Insert chart (column chart from the selection)', () => view.insertChart({ type: 'column' })), sep(),
    numSel,
    btn('.0<span class="arrow">←</span>', 'Increase decimals', () => view.changeDecimals(1)),
    btn('.0<span class="arrow">→</span>', 'Decrease decimals', () => view.changeDecimals(-1)), sep(),
    fnMenu,
    btn(ICONS.recalc, 'Recalculate all (F9)', () => view.recalcAll()), sep(),
    menuSelect('Insert', [['rowAbove', 'Row above'], ['rowBelow', 'Row below'], ['colLeft', 'Column left'], ['colRight', 'Column right']], (v) => ({ rowAbove: () => view.insertRows(false), rowBelow: () => view.insertRows(true), colLeft: () => view.insertCols(false), colRight: () => view.insertCols(true) }[v]())),
    menuSelect('Delete', [['rows', 'Rows'], ['cols', 'Columns']], (v) => (v === 'rows' ? view.deleteRows() : view.deleteCols())),
    btn(ICONS.sortAsc, 'Sort A to Z', () => view.sortSelection(true)),
    btn(ICONS.sortDesc, 'Sort Z to A', () => view.sortSelection(false)),
    btn(ICONS.filter, 'Filter (on/off)', () => view.toggleFilter()),
    menuSelect('Conditional', [['gt', 'Greater than…'], ['lt', 'Less than…'], ['between', 'Between…'], ['eq', 'Equal to…'], ['text', 'Text that contains…'], ['dup', 'Duplicate values'], ['top10', 'Top 10 items'], ['bar', 'Data bars'], ['scale', 'Color scale'], ['icons', 'Icon set (arrows)'], ['clear', 'Clear rules from selection'], ['clearAll', 'Clear rules from sheet']], (v) => view.cfPreset(v)),
    btn(ICONS.freeze, 'Freeze panes at the selected cell (click again to unfreeze)', () => view.toggleFreeze()),
  ];
  wrapBtn.dataset.state = 'wrap';
  tab.refreshSheetTools = () => {
    const css = view.activeCss();
    for (const t of tools) {
      const st = t.dataset && t.dataset.state;
      if (!st) continue;
      let on = false;
      if (st.startsWith('h-')) on = css.h === st.slice(2);
      else if (st === 'wrap') on = css.wrap;
      else on = !!css.font_[st];
      t.classList.toggle('on', on);
    }
    fontSel.value = fonts.includes(css.font_.name) ? css.font_.name : '';
    fontSel.options[0].textContent = fonts.includes(css.font_.name) ? 'Font' : css.font_.name;
    sizeSel.value = sizes.includes(css.font_.size) ? String(css.font_.size) : '';
    sizeSel.options[0].textContent = sizes.includes(css.font_.size) ? '' : String(css.font_.size);
    const known = FORMATS_LIST.find(([v]) => v === css.fmt);
    numSel.value = known ? css.fmt : '__custom';
    numSel.options[numSel.options.length - 1].textContent = known ? 'Custom' : `Custom: ${css.fmt}`;
  };
  tab.refreshSheetTools();
  return tools;
}

// ======================================================================
// Text-based formats: Markdown, HTML, CSV and plain text/code
// ======================================================================

async function setupTextual(tab, modes, initial) {
  const decoded = tab.data ? TextCodec.decode(tab.data) : { text: '', encoding: 'utf-8', bom: false, eol: tab.kind === 'text' ? '\r\n' : '\n' };
  tab.codec = decoded;
  tab.modes = modes.length > 1 ? modes : [];
  const wrap = tab.kind === 'md' || tab.kind === 'text' && ['txt', 'text', 'log', 'md', 'eml', 'nfo', ''].includes(tab.ext);
  const ta = h('textarea', { class: `code${wrap ? '' : ' nowrap'}`, spellcheck: tab.kind === 'md' || ['txt', 'text', ''].includes(tab.ext) ? 'true' : 'false', 'aria-label': tab.name });
  ta.value = decoded.text;
  tab.textarea = ta;
  const backdrop = h('div', { class: `ta-backdrop${wrap ? '' : ' nowrap'}`, 'aria-hidden': 'true' });
  tab.backdrop = backdrop;
  // Hidden copy of the editor text, used to measure where each Markdown block starts (scroll sync).
  const measure = h('div', { class: `ta-backdrop ta-measure${wrap ? '' : ' nowrap'}`, 'aria-hidden': 'true' });
  tab.measure = measure;
  const editPane = h('div', { class: 'edit-pane' }, measure, backdrop, ta);
  const previewPane = h('div', { class: 'preview-pane' });
  tab.body.replaceChildren(h('div', { class: 'split' }, editPane, previewPane));
  tab.splitEl = tab.body.firstChild;

  let timer = null;
  ta.addEventListener('input', () => {
    markDirty(tab);
    if (!$('#findbar').hidden && tab === active) runFind(true);
    clearTimeout(timer);
    timer = setTimeout(() => { if (tab.mode !== 'edit') renderPreview(tab); updateTextInfo(tab); }, 150);
  });
  ta.addEventListener('keydown', (e) => textKeys(e, tab));
  ['keyup', 'click', 'select'].forEach((ev) => ta.addEventListener(ev, () => refreshStatus()));

  // Keep the find highlights in step with the editor, and in Split view keep both panes scrolled together.
  ta.addEventListener('scroll', () => {
    backdrop.scrollTop = ta.scrollTop;
    backdrop.scrollLeft = ta.scrollLeft;
    syncScroll(tab, 'editor');
  });
  previewPane.addEventListener('scroll', () => syncScroll(tab, 'preview'));
  const remap = () => { clearTimeout(tab.remapTimer); tab.remapTimer = setTimeout(() => { buildScrollMap(tab); syncScroll(tab, 'editor', true); }, 120); };
  tab.remap = remap;
  new ResizeObserver(remap).observe(editPane);
  new ResizeObserver(remap).observe(previewPane);
  previewPane.addEventListener('load', remap, true); // pictures change the preview's height

  tab.previewPane = previewPane;
  tab.searchRoot = () => (tab.mode === 'edit' || tab.mode === 'split' ? ta : previewPane);
  tab.printable = () => (tab.mode === 'edit' || tab.kind === 'text' ? ta : previewPane);
  tab.onMode = (mode) => {
    tab.splitEl.dataset.mode = mode;
    if (mode !== 'edit') renderPreview(tab);
    if (mode !== 'preview' && mode !== 'table') ta.focus();
    tab.tools = mode === 'edit' || mode === 'split' ? textTools(tab) : [];
    if (tab.kind === 'html' && mode === 'preview') tab.tools = [h('span', { class: 'hint', text: 'Scripts are disabled in the preview' })];
  };
  updateTextInfo(tab);
  await setMode(tab, initial);
}

function textKeys(e, tab) {
  const ta = e.target;
  if (e.key === 'Tab' && !e.ctrlKey) {
    e.preventDefault();
    const { selectionStart: s, selectionEnd: en, value } = ta;
    if (s !== en && value.slice(s, en).includes('\n')) {
      // Indent / outdent every selected line.
      const lineStart = value.lastIndexOf('\n', s - 1) + 1;
      const block = value.slice(lineStart, en);
      const out = e.shiftKey ? block.replace(/^(\t| {1,4})/gm, '') : block.replace(/^/gm, '\t');
      ta.setSelectionRange(lineStart, en);
      document.execCommand('insertText', false, out);
      ta.setSelectionRange(lineStart, lineStart + out.length);
    } else if (!e.shiftKey) {
      document.execCommand('insertText', false, '\t');
    }
  }
  if (tab.kind === 'md' && e.key === 'Enter' && !e.shiftKey && !e.ctrlKey) {
    // Continue Markdown lists.
    const { selectionStart: s, value } = ta;
    const lineStart = value.lastIndexOf('\n', s - 1) + 1;
    const line = value.slice(lineStart, s);
    const m = line.match(/^(\s*)([-*+]|\d+[.)])(\s+)(\[[ xX]\]\s+)?/);
    if (m) {
      e.preventDefault();
      if (line.trim() === m[0].trim()) { ta.setSelectionRange(lineStart, s); document.execCommand('insertText', false, '\n'); return; }
      let marker = m[2];
      if (/^\d/.test(marker)) marker = `${parseInt(marker, 10) + 1}${marker.slice(-1)}`;
      document.execCommand('insertText', false, `\n${m[1]}${marker}${m[3]}${m[4] ? '[ ] ' : ''}`);
    }
  }
  if (tab.kind === 'md' && e.ctrlKey && !e.shiftKey) {
    const k = e.key.toLowerCase();
    if (k === 'b') { e.preventDefault(); wrapSel(ta, '**', '**', 'bold'); }
    if (k === 'i') { e.preventDefault(); wrapSel(ta, '*', '*', 'italic'); }
    if (k === 'k') { e.preventDefault(); mdLink(ta); }
  }
}

function wrapSel(ta, before, after, placeholder) {
  ta.focus();
  const { selectionStart: s, selectionEnd: e, value } = ta;
  const sel = value.slice(s, e) || placeholder;
  if (value.slice(s - before.length, s) === before && value.slice(e, e + after.length) === after) {
    ta.setSelectionRange(s - before.length, e + after.length);
    document.execCommand('insertText', false, sel);
    ta.setSelectionRange(s - before.length, s - before.length + sel.length);
    return;
  }
  document.execCommand('insertText', false, before + sel + after);
  ta.setSelectionRange(s + before.length, s + before.length + sel.length);
}

function prefixLines(ta, prefix, numbered) {
  ta.focus();
  const { selectionStart: s, selectionEnd: e, value } = ta;
  const lineStart = value.lastIndexOf('\n', s - 1) + 1;
  let lineEnd = value.indexOf('\n', e);
  if (lineEnd < 0) lineEnd = value.length;
  const lines = value.slice(lineStart, lineEnd).split('\n');
  const strip = /^(#{1,6}\s|>\s?|[-*+]\s(\[[ xX]\]\s)?|\d+[.)]\s)/;
  const all = lines.every((l) => l.startsWith(prefix) || (numbered && /^\d+[.)]\s/.test(l)));
  const out = lines.map((l, i) => (all ? l.replace(strip, '') : (numbered ? `${i + 1}. ` : prefix) + l.replace(strip, ''))).join('\n');
  ta.setSelectionRange(lineStart, lineEnd);
  document.execCommand('insertText', false, out);
  ta.setSelectionRange(lineStart, lineStart + out.length);
}

async function mdLink(ta) {
  const { selectionStart: s, selectionEnd: e, value } = ta;
  const text = value.slice(s, e);
  const v = await ask('Link', [{ name: 'url', label: 'Address', value: 'https://' }, { name: 'text', label: 'Text to show', value: text }]);
  ta.focus();
  ta.setSelectionRange(s, e);
  if (v && v.url.trim()) document.execCommand('insertText', false, `[${v.text || v.url}](${v.url.trim()})`);
}

function textTools(tab) {
  const ta = tab.textarea;
  const btn = (label, title, fn) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title });
    b.innerHTML = label;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', fn);
    return b;
  };
  const wrapBtn = h('button', { class: `tool text-tool${ta.classList.contains('nowrap') ? '' : ' on'}`, title: 'Word wrap', text: 'Wrap' });
  wrapBtn.onclick = () => { ta.classList.toggle('nowrap'); tab.backdrop.classList.toggle('nowrap'); tab.measure.classList.toggle('nowrap'); if (tab.remap) tab.remap(); wrapBtn.classList.toggle('on', !ta.classList.contains('nowrap')); };
  const tools = [];
  if (tab.kind === 'md') {
    tools.push(
      btn('<b>B</b>', 'Bold (Ctrl+B)', () => wrapSel(ta, '**', '**', 'bold')),
      btn('<i>I</i>', 'Italic (Ctrl+I)', () => wrapSel(ta, '*', '*', 'italic')),
      btn('<s>S</s>', 'Strikethrough', () => wrapSel(ta, '~~', '~~', 'text')),
      btn('H1', 'Heading 1', () => prefixLines(ta, '# ')),
      btn('H2', 'Heading 2', () => prefixLines(ta, '## ')),
      btn('H3', 'Heading 3', () => prefixLines(ta, '### ')), sep(),
      btn(ICONS.ul, 'Bulleted list', () => prefixLines(ta, '- ')),
      btn(ICONS.ol, 'Numbered list', () => prefixLines(ta, '', true)),
      btn(ICONS.task, 'Task list', () => prefixLines(ta, '- [ ] ')),
      btn(ICONS.quote, 'Quote', () => prefixLines(ta, '> ')), sep(),
      btn(ICONS.code, 'Inline code', () => wrapSel(ta, '`', '`', 'code')),
      btn(ICONS.codeblock, 'Code block', () => wrapSel(ta, '\n```\n', '\n```\n', 'code')),
      btn(ICONS.link, 'Link (Ctrl+K)', () => mdLink(ta)),
      btn(ICONS.table, 'Table', () => { ta.focus(); document.execCommand('insertText', false, '\n| Column 1 | Column 2 | Column 3 |\n| --- | --- | --- |\n|  |  |  |\n'); }),
      btn(ICONS.rule, 'Horizontal line', () => { ta.focus(); document.execCommand('insertText', false, '\n\n---\n\n'); }),
      sep(),
    );
  }
  if (tab.ext === 'json') {
    tools.push(h('button', { class: 'tool text-tool', text: 'Format JSON', onclick: () => {
      try {
        const out = JSON.stringify(JSON.parse(ta.value), null, 2);
        ta.focus(); ta.select(); document.execCommand('insertText', false, out);
      } catch (err) { desk.error('This JSON has an error', err.message); }
    } }));
  }
  tools.push(wrapBtn);
  return tools;
}

// Markdown preview: resolve relative images and links against the file's folder.
let previewBaseDir = null;
DOMPurify.addHook('afterSanitizeAttributes', (node) => {
  if (!previewBaseDir) return;
  if (node.tagName === 'IMG') {
    const src = node.getAttribute('src');
    if (src && !/^(https?:|data:|blob:|file:)/i.test(src)) node.setAttribute('src', fileUrl(localPath(previewBaseDir, src)));
  }
  if (node.tagName === 'A' && node.getAttribute('href')) node.setAttribute('data-href', node.getAttribute('href'));
});

// ---- Split view scroll sync
// Each top-level Markdown block (heading, paragraph, list, table, code...) is an anchor: its first line in the
// editor and its top in the preview. Scrolling either pane moves the other to the same point between anchors,
// so both sides stay aligned even where one side is much taller (pictures, tables, long code lines).

function buildScrollMap(tab) {
  const ta = tab.textarea, pane = tab.previewPane;
  const body = pane && pane.querySelector('.markdown-body');
  tab.scrollMap = null;
  if (tab.kind !== 'md' || tab.mode !== 'split' || !body || !ta.clientHeight) return;
  let tokens;
  try { tokens = marked.lexer(ta.value, { gfm: true }); } catch { return; }
  const starts = [];
  let line = 0;
  for (const t of tokens) {
    if (t.type !== 'space' && t.type !== 'def') starts.push(line);
    line += (t.raw.match(/\n/g) || []).length;
  }
  const blocks = [...body.children];
  // Measure where those lines start in the editor (wrapping included) with a hidden copy of the text.
  const lines = ta.value.split('\n');
  const want = new Set(starts);
  const frag = [];
  const markers = new Map();
  let chunk = '';
  lines.forEach((l, i) => {
    if (want.has(i)) { if (chunk) frag.push(chunk); chunk = ''; const m = h('span'); markers.set(i, m); frag.push(m); }
    chunk += l + (i < lines.length - 1 ? '\n' : '');
  });
  frag.push(chunk + ' ');
  tab.measure.replaceChildren(...frag);
  const edPad = parseFloat(getComputedStyle(ta).paddingTop) || 0;
  const pvTop = pane.getBoundingClientRect().top - pane.scrollTop;
  const edMax = Math.max(0, ta.scrollHeight - ta.clientHeight);
  const pvMax = Math.max(0, pane.scrollHeight - pane.clientHeight);
  const pairs = [[0, 0]];
  // Pair blocks in order; if the counts differ (raw HTML can render as several elements) pair what lines up.
  const n = Math.min(starts.length, blocks.length);
  for (let i = 0; i < n; i++) {
    const m = markers.get(starts[i]);
    if (!m) continue;
    const ed = Math.min(edMax, Math.max(0, m.offsetTop - edPad));
    const pv = Math.min(pvMax, Math.max(0, blocks[i].getBoundingClientRect().top - pvTop - 16));
    const last = pairs[pairs.length - 1];
    if (ed > last[0] && pv >= last[1]) pairs.push([ed, pv]);
  }
  tab.measure.replaceChildren();
  const last = pairs[pairs.length - 1];
  if (edMax > last[0] && pvMax >= last[1]) pairs.push([edMax, pvMax]);
  else pairs[pairs.length - 1] = [Math.max(last[0], edMax), Math.max(last[1], pvMax)];
  pairs.heights = [ta.scrollHeight, pane.scrollHeight, ta.clientWidth, pane.clientWidth];
  tab.scrollMap = pairs;
}

function mapScroll(pairs, y, from, to) {
  for (let i = 1; i < pairs.length; i++) {
    const a = pairs[i - 1], b = pairs[i];
    if (y <= b[from] || i === pairs.length - 1) {
      const span = b[from] - a[from];
      const f = span > 0 ? Math.min(1, Math.max(0, (y - a[from]) / span)) : 0;
      return a[to] + f * (b[to] - a[to]);
    }
  }
  return 0;
}

function syncScroll(tab, source, force) {
  if (tab.mode !== 'split' || !tab.previewPane) return;
  const ta = tab.textarea, pane = tab.previewPane;
  // Ignore the scroll event caused by our own update of the other pane.
  if (!force && tab.syncTarget === source && Date.now() < tab.syncUntil) return;
  const [src, dst] = source === 'editor' ? [ta, pane] : [pane, ta];
  // Anything that changed either pane's size (pictures loading, fonts, typing) makes the map stale: rebuild it.
  const hs = tab.scrollMap && tab.scrollMap.heights;
  if (tab.kind === 'md' && (!tab.scrollMap || !hs || hs[0] !== ta.scrollHeight || hs[1] !== pane.scrollHeight || hs[2] !== ta.clientWidth || hs[3] !== pane.clientWidth)) buildScrollMap(tab);
  let y;
  if (tab.scrollMap) y = mapScroll(tab.scrollMap, src.scrollTop, source === 'editor' ? 0 : 1, source === 'editor' ? 1 : 0);
  else y = src.scrollTop / Math.max(1, src.scrollHeight - src.clientHeight) * (dst.scrollHeight - dst.clientHeight);
  y = Math.round(y);
  if (Math.abs(dst.scrollTop - y) < 1) return;
  tab.syncTarget = source === 'editor' ? 'preview' : 'editor';
  tab.syncUntil = Date.now() + 80;
  dst.scrollTop = y;
}

function renderPreview(tab) {
  const pane = tab.previewPane;
  const text = tab.textarea.value;
  if (tab.kind === 'md') {
    previewBaseDir = tab.path ? dirOf(tab.path) : null;
    const scroll = pane.scrollTop;
    const body = h('article', { class: 'markdown-body' });
    body.innerHTML = sanitize(marked.parse(text, { gfm: true, breaks: false }), { FORBID_TAGS: ['style', 'form', 'button'], ADD_ATTR: ['checked', 'disabled', 'type'] });
    previewBaseDir = null;
    body.querySelectorAll('input:not([type="checkbox"])').forEach((i) => i.remove());
    body.querySelectorAll('a[data-href]').forEach((a) => a.addEventListener('click', (e) => {
      e.preventDefault();
      const href = a.dataset.href;
      if (/^(https?:|mailto:)/i.test(href)) desk.openExternal(href);
      else if (href.startsWith('#')) body.querySelector(`[id="${CSS.escape(decodeURIComponent(href.slice(1)))}"]`)?.scrollIntoView();
      else if (tab.path) openPath(localPath(dirOf(tab.path), href.split('#')[0]));
    }));
    // marked doesn't add heading ids; add GitHub-style ones so #anchors work.
    const seen = {};
    body.querySelectorAll('h1,h2,h3,h4,h5,h6').forEach((hd) => {
      let id = hd.textContent.trim().toLowerCase().replace(/[^\p{L}\p{N}\s-]/gu, '').replace(/\s/g, '-');
      if (seen[id] != null) id += `-${++seen[id]}`; else seen[id] = 0;
      hd.id = id;
    });
    pane.replaceChildren(body);
    pane.scrollTop = scroll;
    if (tab.remap) tab.remap();
  } else if (tab.kind === 'html') {
    const frame = pane.querySelector('iframe') || pane.appendChild(h('iframe', { class: 'html-frame', sandbox: '', title: 'HTML preview' }));
    frame.srcdoc = text;
  } else if (tab.kind === 'csv') {
    const wb = XLSX.read(text, { type: 'string', raw: true, FS: tab.ext === 'tsv' ? '\t' : undefined });
    pane.replaceChildren();
    pane.classList.add('fill');
    showWorkbook(wb, pane);
  }
}

function updateTextInfo(tab) {
  const v = tab.textarea.value;
  const lines = v.length ? v.split('\n').length : 0;
  tab.info = tab.kind === 'md' || ['txt', 'text', ''].includes(tab.ext)
    ? `${countWords(v).toLocaleString()} words · ${lines.toLocaleString()} lines`
    : `${lines.toLocaleString()} lines`;
  if (tab === active) refreshStatus();
}

// ======================================================================
// Tabs, modes and window chrome
// ======================================================================

async function setMode(tab, mode) {
  if (!tab.modes.includes(mode) && tab.modes.length) mode = tab.modes[0];
  tab.mode = mode;
  if (tab.onMode) {
    try { await tab.onMode(mode); } catch (err) { logError(err); desk.error(`Couldn't switch to ${MODE_LABELS[mode]}`, err.message); }
  }
  if (tab === active) refreshChrome();
}

function activate(tab) {
  active = tab;
  tabs.forEach((t) => t.el.classList.toggle('active', t === tab));
  closeFind();
  refreshChrome();
  if (tab && tab.onActivate) tab.onActivate();
}

function refreshChrome() {
  // Tab strip
  const strip = $('#tabs');
  strip.replaceChildren(...tabs.map((t) => {
    const close = h('button', { class: 'tab-close', title: 'Close (Ctrl+W)', 'aria-label': `Close ${t.name}`, text: '×' });
    close.addEventListener('click', (e) => { e.stopPropagation(); closeTab(t); });
    const el = h('div', { class: `tab${t === active ? ' active' : ''}${t.dirty ? ' dirty' : ''}`, role: 'tab', 'aria-selected': t === active ? 'true' : 'false', title: t.path || t.name, tabindex: '0' },
      h('span', { class: `tab-kind k-${t.kind}`, text: (t.ext || '?').slice(0, 4).toUpperCase() }),
      h('span', { class: 'tab-name', text: t.name }),
      h('span', { class: 'tab-dot', 'aria-hidden': 'true', text: '●' }),
      close);
    el.addEventListener('click', () => activate(t));
    el.addEventListener('auxclick', (e) => { if (e.button === 1) closeTab(t); });
    el.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') activate(t); });
    return el;
  }));
  // "+" after the last tab: a new document of any type.
  const plus = h('button', { class: 'tab-new', title: 'New document', 'aria-label': 'New document', text: '+' });
  plus.addEventListener('click', () => {
    const r = plus.getBoundingClientRect();
    SheetView.showMenu(r.left, r.bottom + 2, [
      ['Markdown', 'Ctrl+Shift+N', () => newDocument('md')], null,
      ['Document (.docx)', 'Ctrl+N', () => newDocument('docx')], ['Spreadsheet (.xlsx)', '', () => newDocument('xlsx')], ['Presentation (.pptx)', '', () => newDocument('pptx')], ['PDF', '', () => newDocument('pdf')], null,
      ['OpenDocument text', '', () => newDocument('odt')], ['OpenDocument spreadsheet', '', () => newDocument('ods')], ['OpenDocument presentation', '', () => newDocument('odp')], ['Rich Text', '', () => newDocument('rtf')], null,
      ['Plain text', '', () => newDocument('txt')], ['Web page (HTML)', '', () => newDocument('html')], ['CSV table', '', () => newDocument('csv')], ['JSON', '', () => newDocument('json')],
    ]);
  });
  strip.append(plus);
  const at = strip.querySelector('.tab.active');
  if (at) at.scrollIntoView({ block: 'nearest', inline: 'nearest' });

  // Toolbar (shown while there are editing tools)
  const bar = $('#toolbar');
  const sw = $('#mode-switch');
  if (!active) { bar.hidden = true; sw.hidden = true; if (typeof Toolbar !== 'undefined') Toolbar.layout(null, []); } else {
    const items = [...(typeof commonTools === 'function' ? commonTools(active) : []), ...(active.tools || [])];
    // Two lines: main tools beside the zoom, the rest (with the menus) below - see toolbar-layout.js.
    if (typeof Toolbar !== 'undefined') Toolbar.layout(active.tools && active.tools.length ? active : null, active.tools && active.tools.length ? items : []);
    else $('#tool-left').replaceChildren(...items);
    sw.hidden = active.modes.length < 2;
    sw.replaceChildren(...active.modes.map((m) => h('button', {
      class: m === active.mode ? 'on' : '', text: MODE_LABELS[m], title: `${MODE_LABELS[m]} (Ctrl+E switches)`,
      onclick: () => setMode(active, m),
    })));
    bar.hidden = !(active.tools && active.tools.length);
  }
  if (typeof Shell !== 'undefined') Shell.refresh();
  if (typeof Menus !== 'undefined') Menus.render();
  document.title = active ? `${active.name}${active.dirty ? ' •' : ''} - Doc Studio` : 'Doc Studio';
  refreshStatus();
}

function refreshStatus() {
  const left = $('#status-left');
  const right = $('#status-right');
  if (!active) { left.textContent = 'Ready'; right.textContent = ''; return; }
  const parts = [FORMATS.label(active.kind)];
  if (active.size) parts.push(formatBytes(active.size));
  if (active.info) parts.push(active.info);
  if (active.status) parts.push(active.status);
  if (active.zoom && active.zoom !== 1) parts.push(`${Math.round(active.zoom * 100)}%`);
  left.textContent = parts.join('  ·  ');
  const r = [];
  const ta = active.textarea;
  if (ta && (active.mode === 'edit' || active.mode === 'split')) {
    const before = ta.value.slice(0, ta.selectionStart);
    const line = before.split('\n').length;
    const col = ta.selectionStart - before.lastIndexOf('\n');
    r.push(`Ln ${line}, Col ${col}`);
    if (ta.selectionEnd > ta.selectionStart) r.push(`${ta.selectionEnd - ta.selectionStart} selected`);
    r.push(active.codec.encoding.toUpperCase() + (active.codec.bom ? ' BOM' : ''), active.codec.eol === '\r\n' ? 'CRLF' : 'LF');
  }
  right.textContent = r.join('   ');
}

function markDirty(tab) {
  tab.editSeq = (tab.editSeq || 0) + 1; // tells the recovery copy (recovery.js) there is something new
  if (tab.dirty) return;
  tab.dirty = true;
  refreshChrome();
}

async function closeTab(tab) {
  if (tab.dirty) {
    const r = await desk.confirm({ type: 'warning', message: `Save changes to ${tab.name}?`, detail: "Your changes will be lost if you don't save them.", buttons: ['Save', "Don't save", 'Cancel'], defaultId: 0, cancelId: 2 });
    if (r === 2) return false;
    if (r === 0 && !(await save(tab))) return false;
  }
  const i = tabs.indexOf(tab);
  tabs.splice(i, 1);
  Recovery.forget(tab);
  tab.urls.forEach((u) => URL.revokeObjectURL(u));
  if (tab.sheetView) tab.sheetView.destroy();
  if (tab.pdfEditor) tab.pdfEditor.destroy();
  if (tab.pptEditor) tab.pptEditor.destroy();
  tab.el.remove();
  if (active === tab) active = tabs[Math.min(i, tabs.length - 1)] || null;
  if (active) activate(active); else refreshChrome();
  return true;
}

// ======================================================================
// Saving
// ======================================================================

// The format each editor works in. Saving to another extension converts from it.
function nativeExt(tab) {
  switch (tab.kind) {
    case 'docx': return 'docx';
    case 'sheet': return /^xl[st]m$/.test(tab.ext) ? 'xlsm' : 'xlsx';
    case 'pptx': return 'pptx';
    case 'pdf': return 'pdf';
    case 'odt': return 'odt';
    case 'rtf': return 'rtf';
    case 'image': return ImageEditor.WRITES.includes(tab.ext) ? tab.ext : 'png';
    default: return tab.ext || 'txt';
  }
}
const TEXTUAL = new Set(['md', 'html', 'csv', 'text']);
const SHEETJS_OUT = { xls: 'biff8', xlsb: 'xlsb', ods: 'ods', fods: 'fods' };
const EXPORTS = {
  docx: ['odt', 'rtf', 'doc', 'pdf', 'txt'], odt: ['docx', 'rtf', 'doc', 'pdf'], rtf: ['docx', 'odt', 'doc', 'pdf'],
  png: ['jpg', 'webp', 'bmp'], jpg: ['png', 'webp', 'bmp'], jpeg: ['png', 'webp', 'bmp'], webp: ['png', 'jpg', 'bmp'], bmp: ['png', 'jpg', 'webp'],
  xlsx: ['ods', 'xls', 'xlsb', 'csv', 'pdf'], xlsm: ['xlsx', 'xls', 'xlsb', 'ods', 'pdf'], pptx: ['odp', 'ppt', 'pdf'],
};

async function canExport(tab, native, ext) {
  if (!ext || ext === native || TEXTUAL.has(tab.kind)) return true;
  if (tab.kind === 'image') return ImageEditor.WRITES.includes(ext);
  if (tab.kind === 'sheet' && SHEETJS_OUT[ext]) return true;
  if (tab.kind === 'odt' && ext === 'docx' && !(await desk.convertRoute('odt', 'docx'))) return false;
  return !!(await desk.convertRoute(native, ext));
}

async function exportBytes(tab, bytes, native, ext) {
  if (ext === native || TEXTUAL.has(tab.kind)) return bytes;
  if (tab.kind === 'image') return tab.imageEditor.export(ext);
  if (await desk.convertRoute(native, ext)) return (await desk.convert(bytes, native, ext)).data;
  if (tab.kind === 'sheet' && SHEETJS_OUT[ext]) {
    const wb = Convert.keepModernAsValues(XLSX.read(bytes, { type: 'array', cellStyles: true, cellNF: true, cellFormula: true }));
    return Convert.fixOds(XLSX.write(wb, { bookType: SHEETJS_OUT[ext], type: 'array', cellStyles: true }), ext);
  }
  throw new Error(`Doc Studio can't write .${ext} files on this PC.`);
}

async function exportChoices(tab, native) {
  const out = [];
  for (const e of EXPORTS[native] || []) if (await canExport(tab, native, e)) out.push(e);
  return out;
}

function hasEditor(tab) {
  switch (tab.kind) {
    case 'docx': return !!tab.editor;
    case 'pdf': return !!tab.pdfEditor;
    case 'pptx': return !!tab.pptEditor;
    case 'odt': return !!tab.odtEditor;
    case 'rtf': return !!tab.rtfEditor;
    case 'image': return !!tab.imageEditor;
    default: return true;
  }
}

// The document's current content in its editor's own format. commit: finish a cell edit or PDF text box first
// (Save does; the recovery copy doesn't, so it never interrupts typing).
async function nativeBytes(tab, commit) {
  if (tab.kind === 'docx') return (await tab.editor.build()).bytes;
  if (tab.kind === 'sheet') { if (commit) tab.sheetView.endEdit(true); return tab.workbook.save(); }
  if (tab.kind === 'pptx') return tab.pptEditor.save();
  if (tab.kind === 'odt') return tab.odtEditor.save();
  if (tab.kind === 'rtf') return tab.rtfEditor.save();
  if (tab.kind === 'image') return tab.imageEditor.export(nativeExt(tab));
  if (tab.kind === 'pdf') { if (commit && document.activeElement && document.activeElement.isContentEditable) document.activeElement.blur(); return tab.pdfEditor.build(); }
  return TextCodec.encode(tab.textarea.value, tab.codec);
}

async function save(tab, saveAs) {
  if (!tab) return false;
  const editable = FORMATS.isEditable(tab.kind);
  // View-only files (or a document never opened in Edit) - Save As writes its bytes as they are. A document that
  // was never saved (from the phone, a scan, OCR...) always gets Save As, and then becomes the saved file.
  if (!editable || !hasEditor(tab)) {
    if (!saveAs && tab.path) { flash('Nothing to save - this file has no edits.'); return true; }
    if (!tab.data) return false;
    const target = await desk.saveDialog(tab.path || tab.name, tab.ext || 'bin');
    if (!target) return false;
    await desk.write(target, tab.data);
    if (!tab.path) {
      tab.path = target;
      tab.name = target.split(/[\\/]/).pop();
      tab.ext = FORMATS.extOf(tab.name);
      tab.dirty = false;
      tab.isNew = false;
      Recovery.forget(tab);
      refreshChrome();
      flash(`Saved ${tab.name}`);
    } else flash(`Saved a copy to ${target}`);
    return true;
  }

  const native = nativeExt(tab);
  let target = tab.path;
  let ext = target ? FORMATS.extOf(target) : (tab.ext || native);
  // Formats Doc Studio can't write here (templates, .numbers, .doc without Word...) go through Save As.
  if (target && !saveAs && !(await canExport(tab, native, ext))) saveAs = true;
  if (!target || saveAs) {
    const choices = TEXTUAL.has(tab.kind) ? [] : await exportChoices(tab, native);
    const def = TEXTUAL.has(tab.kind) || choices.includes(ext) ? ext : native;
    const suggested = (target ? target.replace(/\.[^.\\/]+$/, '') : stem(tab.name)) + `.${def}`;
    target = await desk.saveDialog(suggested, def, [native, ...choices].filter((x) => x !== def));
    if (!target) return false;
    ext = FORMATS.extOf(target);
    if (!(await canExport(tab, native, ext))) {
      desk.error(`Doc Studio can't save this document as .${ext} on this PC.`, `It can be saved as: ${[native, ...choices].map((x) => '.' + x).join(', ')}`);
      return false;
    }
  }

  try {
    const bytes = await nativeBytes(tab, true);
    const converted = ext !== native && !TEXTUAL.has(tab.kind);
    if (converted) flash(`Converting to .${ext}…`);
    const out = await exportBytes(tab, bytes, native, ext);
    // First overwrite of an existing file: keep a backup copy of the original in Doc Studio's data folder.
    if (target === tab.path && !tab.backedUp && !tab.isNew) { await desk.backup(target); tab.backedUp = true; }
    await desk.write(target, out);
    const renamed = target !== tab.path;
    tab.path = target;
    tab.name = target.split(/[\\/]/).pop();
    tab.size = out.byteLength;
    tab.dirty = false;
    tab.isNew = false;
    Recovery.forget(tab);
    if (['docx', 'sheet', 'pptx', 'odt', 'rtf'].includes(tab.kind)) tab.data = bytes;
    if (tab.kind === 'image') tab.data = out;
    if (tab.kind === 'pdf') { tab.data = bytes; tab.viewBytes = null; const ed = tab.pdfEditor; ed.items = []; ed.formValues = new Map(); await ed.load(bytes, true); }
    if (renamed) {
      const kind = FORMATS.kindOf(tab.name);
      // A text file saved with a new extension (e.g. .txt -> .md) switches to that handler on next open.
      if (!converted && kind !== tab.kind && FORMATS.isEditable(kind)) flash('Saved. Reopen the file to use the editor for its new type.');
      tab.ext = FORMATS.extOf(tab.name);
    }
    refreshChrome();
    if (converted) flash(`Saved ${tab.name}`);
    else if (!renamed || FORMATS.kindOf(tab.name) === tab.kind) flash(`Saved ${tab.name}`);
    return true;
  } catch (err) {
    desk.error(`Couldn't save ${tab.name}`, err.message);
    return false;
  }
}

// ======================================================================
// Find (CSS Custom Highlights for rendered content, selection for text editors)
// ======================================================================

const find = { hits: [], index: -1 };

function openFind(withReplace) {
  if (!active) return;
  if (active.kind === 'pdf' && active.mode !== 'edit') { flash('Click inside the PDF and press Ctrl+F to use its search.'); return; }
  $('#findbar').hidden = false;
  if (withReplace) $('#replace-row').hidden = false;
  const input = $('#find-input');
  const ta = active.textarea;
  if (ta && document.activeElement === ta && ta.selectionEnd > ta.selectionStart) input.value = ta.value.slice(ta.selectionStart, ta.selectionEnd);
  else if (getSelection().toString() && !getSelection().toString().includes('\n')) input.value = getSelection().toString();
  input.focus();
  input.select();
  runFind();
}

function closeFind() {
  $('#findbar').hidden = true;
  $('#replace-row').hidden = true;
  CSS.highlights.clear();
  tabs.forEach((t) => { if (t.backdrop) { t.backdrop.replaceChildren(); t.backdrop.parentElement.classList.remove('finding'); } });
  find.hits = []; find.index = -1;
  $('#find-count').textContent = '';
}

// The search as a global RegExp (plain text is escaped), or null for an empty or invalid pattern.
function findPattern() {
  const q = $('#find-input').value;
  if (!q) return null;
  const regex = $('#find-regex').getAttribute('aria-pressed') === 'true';
  const flags = ($('#find-case').getAttribute('aria-pressed') === 'true' ? 'g' : 'gi') + (regex ? 'm' : '');
  try { return new RegExp(regex ? q : q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags + 'u'); } catch { return null; }
}
function matchesIn(text, re) {
  const out = [];
  re.lastIndex = 0;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    if (!m[0].length) { re.lastIndex++; continue; }
    out.push({ start: m.index, end: m.index + m[0].length, groups: m });
  }
  return out;
}

function runFind(keepIndex) {
  CSS.highlights.clear();
  const re = findPattern();
  $('#find-input').classList.toggle('bad', !!$('#find-input').value && !re);
  if (active && active.sheetView) {
    find.hits = re ? active.sheetView.findAll($('#find-input').value, re).map((x) => ({ cell: x })) : [];
    if (!keepIndex) find.index = find.hits.length ? 0 : -1;
    find.index = Math.min(find.index, find.hits.length - 1);
    showHit();
    return;
  }
  if (active && active.backdrop) { active.backdrop.replaceChildren(); active.backdrop.parentElement.classList.remove('finding'); }
  find.hits = [];
  const root = active && active.searchRoot && active.searchRoot();
  if (!re || !root) { $('#find-count').textContent = re || !$('#find-input').value ? '' : 'Invalid pattern'; return; }
  if (root.tagName === 'TEXTAREA') {
    for (const m of matchesIn(root.value, re)) find.hits.push({ ta: root, start: m.start, end: m.end, m: m.groups });
  } else {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: (n) => (n.parentElement && n.parentElement.closest('style,script,[hidden]') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
    });
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      for (const m of matchesIn(n.nodeValue, re)) {
        const r = new Range(); r.setStart(n, m.start); r.setEnd(n, m.end);
        find.hits.push({ range: r, m: m.groups });
      }
    }
    if (find.hits.length) CSS.highlights.set('find-hit', new Highlight(...find.hits.map((x) => x.range)));
  }
  if (!keepIndex) find.index = find.hits.length ? 0 : -1;
  find.index = Math.min(find.index, find.hits.length - 1);
  showHit();
}

function stepFind(dir) {
  if (!find.hits.length) { runFind(); return; }
  find.index = (find.index + dir + find.hits.length) % find.hits.length;
  showHit();
}

function showHit() {
  const count = $('#find-count');
  if (!find.hits.length) { count.textContent = $('#find-input').value ? 'No results' : ''; return; }
  count.textContent = `${find.index + 1} of ${find.hits.length}`;
  const hit = find.hits[find.index];
  if (hit.cell) { active.sheetView.select(hit.cell.r, hit.cell.c); return; }
  if (hit.ta) {
    // Select in the textarea so the match is visible and scrolled to, then give focus back to the find box.
    const input = document.activeElement === $('#replace-input') ? $('#replace-input') : $('#find-input');
    hit.ta.focus();
    hit.ta.setSelectionRange(hit.start, hit.end);
    input.focus();
    paintTextHits(hit.ta);
    refreshStatus();
  } else {
    CSS.highlights.set('find-current', new Highlight(hit.range));
    const el = hit.range.startContainer.parentElement;
    if (el) el.scrollIntoView({ block: 'center', inline: 'nearest' });
  }
}

// ---- replace: text editors (undoable), editable pages and spreadsheet cells

function replacementFor(hit) {
  const rep = $('#replace-input').value;
  if ($('#find-regex').getAttribute('aria-pressed') !== 'true' || !hit.m) return rep;
  const m = hit.m;
  return rep.replace(/\$(\d+|&)/g, (x, g) => (g === '&' ? m[0] : m[+g] ?? ''));
}
function hitEditable(hit) {
  if (hit.ta) return !hit.ta.readOnly && hit.ta.closest('[data-mode="edit"],[data-mode="split"]') !== null;
  if (hit.cell) return true;
  const el = hit.range.startContainer.parentElement;
  return !!(el && el.isContentEditable);
}
function replaceHit(hit) {
  const text = replacementFor(hit);
  if (hit.cell) {
    const sv = active.sheetView;
    const c = sv.sheet.get(hit.cell.r, hit.cell.c);
    const src = c ? (c.formula ? `=${c.formula}` : String(c.value ?? '')) : '';
    const re = findPattern();
    sv.setFromText(hit.cell.r, hit.cell.c, src.replace(re, $('#find-regex').getAttribute('aria-pressed') === 'true' ? $('#replace-input').value : () => $('#replace-input').value));
    return;
  }
  if (hit.ta) {
    hit.ta.focus();
    hit.ta.setSelectionRange(hit.start, hit.end);
    document.execCommand('insertText', false, text);
    return;
  }
  const host = hit.range.startContainer.parentElement.closest('[contenteditable="true"]');
  host.focus();
  const sel = getSelection();
  sel.removeAllRanges();
  sel.addRange(hit.range);
  if (text) document.execCommand('insertText', false, text); else document.execCommand('delete');
}
function replaceOne() {
  if (!find.hits.length) { runFind(); return; }
  const hit = find.hits[find.index];
  if (!hitEditable(hit)) { flash('Switch to Edit to replace text.'); return; }
  replaceHit(hit);
  runFind(true);
  $('#replace-input').focus();
}
function replaceAll() {
  runFind();
  const hits = find.hits.filter(hitEditable);
  if (!hits.length) { flash(find.hits.length ? 'Switch to Edit to replace text.' : 'No matches to replace.'); return; }
  const n = hits.length;
  const ta = hits[0].ta;
  if (ta) {
    // One undo step for the whole text.
    const re = findPattern();
    const rep = $('#replace-input').value;
    const value = ta.value.replace(re, $('#find-regex').getAttribute('aria-pressed') === 'true' ? rep : () => rep);
    ta.focus();
    ta.select();
    document.execCommand('insertText', false, value);
  } else {
    // Last to first, so earlier matches stay where they are.
    for (const hit of [...hits].reverse()) replaceHit(hit);
  }
  runFind();
  flash(`Replaced ${n} match${n === 1 ? '' : 'es'}`);
}

// Textareas can't style ranges, so matches are painted on a mirror layer behind the (transparent) textarea.
function paintTextHits(ta) {
  const backdrop = active.backdrop;
  if (!backdrop) return;
  const v = ta.value;
  const out = [];
  let last = 0;
  find.hits.forEach((hit, i) => {
    out.push(document.createTextNode(v.slice(last, hit.start)));
    out.push(h('mark', { class: i === find.index ? 'current' : null, text: v.slice(hit.start, hit.end) }));
    last = hit.end;
  });
  out.push(document.createTextNode(v.slice(last) + ' '));
  backdrop.replaceChildren(...out);
  backdrop.parentElement.classList.add('finding');
  backdrop.scrollTop = ta.scrollTop;
  backdrop.scrollLeft = ta.scrollLeft;
}

// ======================================================================
// Print and PDF export: copy the printable content into #print-root
// ======================================================================

function preparePrint() {
  if (!active || !active.printable) return false;
  const src = active.printable();
  if (!src) {
    flash(active.kind === 'pdf' ? 'Use the print button in the PDF toolbar.' : 'This file type can\'t be printed from Doc Studio.');
    return false;
  }
  const root = $('#print-root') || document.body.appendChild(h('div', { id: 'print-root' }));
  if (src.tagName === 'TEXTAREA') root.replaceChildren(h('pre', { class: 'print-text', text: src.value }));
  else {
    const clone = src.cloneNode(true);
    clone.removeAttribute('contenteditable');
    root.replaceChildren(clone);
  }
  root.className = `print-${active.kind}`;
  document.body.classList.add('has-print');
  return true;
}
function endPrint() { document.body.classList.remove('has-print'); const r = $('#print-root'); if (r) r.replaceChildren(); }

async function doPrint() {
  if (active && active.beforePrint) await active.beforePrint();
  if (!preparePrint()) return;
  flash('Opening the print dialog…');
  const r = await desk.print();
  endPrint();
  if (r && (r.reason === 'no-response' || r.reason === 'no-printers')) {
    const c = await desk.confirm({ type: 'warning', message: r.reason === 'no-printers' ? 'No printer is set up on this PC.' : "Windows' printing service isn't answering.", detail: r.reason === 'no-printers' ? 'Add a printer in Windows Settings, or save the document as a PDF instead.' : 'A printer driver may be stuck. You can save the document as a PDF and print that later, or restart the PC and try again.', buttons: ['Export as PDF…', 'Cancel'], defaultId: 0, cancelId: 1 });
    if (c === 0) await doExportPdf();
  }
}
async function doExportPdf() {
  if (active && active.beforePrint) await active.beforePrint();
  if (!preparePrint()) return;
  const out = await desk.exportPdf((active.path ? active.path.replace(/\.[^.\\/]+$/, '') : stem(active.name)) + '.pdf');
  endPrint();
  if (out) flash(`Exported ${out.split(/[\\/]/).pop()}`);
}

// ======================================================================
// Small UI helpers
// ======================================================================

function sep() { return h('span', { class: 'tool-sep', 'aria-hidden': 'true' }); }
function stem(name) { return name.replace(/\.[^.]+$/, ''); }
function dirOf(p) { return p.slice(0, Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/'))); }
function localPath(dir, rel) {
  const clean = decodeURIComponent(rel.split('?')[0]);
  if (/^[a-z]:[\\/]/i.test(clean)) return clean;
  return resolvePath(dir.replace(/\\/g, '/') + '/', clean.replace(/\\/g, '/')).replace(/\//g, '\\');
}
function fileUrl(p) { return 'file:///' + p.replace(/\\/g, '/').split('/').map((s, i) => (i === 0 ? s : encodeURIComponent(s))).join('/'); }

let flashTimer = null;
function flash(msg) {
  const s = $('#status-left');
  s.textContent = msg;
  s.classList.add('flash');
  clearTimeout(flashTimer);
  flashTimer = setTimeout(() => { s.classList.remove('flash'); refreshStatus(); }, 3500);
}

// Promise-based in-app form (Electron has no window.prompt).
function ask(title, fields) {
  return new Promise((resolve) => {
    const inputs = fields.map((f) => h('input', { name: f.name, type: f.type || 'text', value: f.value || '', min: f.type === 'number' ? '1' : null }));
    const form = h('form', { class: 'dialog' }, h('h3', { text: title }),
      ...fields.map((f, i) => h('label', null, h('span', { text: f.label }), inputs[i])),
      h('div', { class: 'dialog-actions' }, h('button', { type: 'button', class: 'ghost', text: 'Cancel', onclick: () => done(null) }), h('button', { type: 'submit', text: 'OK' })));
    const backdrop = h('div', { class: 'backdrop' }, form);
    const done = (v) => { backdrop.remove(); resolve(v); };
    form.addEventListener('submit', (e) => { e.preventDefault(); done(Object.fromEntries(inputs.map((i) => [i.name, i.value]))); });
    backdrop.addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.stopPropagation(); done(null); } });
    backdrop.addEventListener('mousedown', (e) => { if (e.target === backdrop) done(null); });
    document.body.append(backdrop);
    inputs[0].focus();
    inputs[0].select();
  });
}

function showHelp() {
  desk.confirm({
    type: 'info', buttons: ['OK'], message: 'Supported file types',
    detail: [
      'View and edit',
      '  Documents: .docx (View shows the pages; Edit keeps every bit of the original formatting)',
      '  Markdown: .md, .markdown (edit, split or preview)',
      '  HTML: .html, .htm (preview or edit source)',
      '  CSV / TSV: table view or edit as text',
      '  Text and code: .txt, .json, .xml, .yaml, .log, .ini, .js, .py, .cs and more',
      '',
      'View',
      '  PDF, spreadsheets (.xlsx, .xlsm, .xls, .xlsb), OpenDocument (.odt, .ods, .odp)',
      '  Presentations (.pptx), Rich Text (.rtf), 97-2003 documents (.doc, text only)',
      '  EPUB books, images (PNG, JPEG, GIF, WebP, SVG, BMP, ICO, AVIF), audio and video',
      '',
      'Shortcuts: Ctrl+O open, Ctrl+S save, Ctrl+E switch View/Edit, Ctrl+F find, Ctrl+P print, Ctrl+Shift+E export PDF, Ctrl+W close tab.',
    ].join('\n'),
  });
}

const svg = (d) => `<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${d}</svg>`;
const ICONS = {
  ul: svg('<line x1="9" y1="6" x2="20" y2="6"/><line x1="9" y1="12" x2="20" y2="12"/><line x1="9" y1="18" x2="20" y2="18"/><circle cx="4.5" cy="6" r="1"/><circle cx="4.5" cy="12" r="1"/><circle cx="4.5" cy="18" r="1"/>'),
  ol: svg('<line x1="10" y1="6" x2="20" y2="6"/><line x1="10" y1="12" x2="20" y2="12"/><line x1="10" y1="18" x2="20" y2="18"/><path d="M4 4h1v4M4 8h2M4 11.5c.5-.7 2-.7 2 .3 0 .8-2 1.7-2 2.7h2"/>'),
  task: svg('<rect x="3" y="5" width="6" height="6" rx="1"/><path d="M4.5 8l1 1 2-2"/><line x1="12" y1="8" x2="21" y2="8"/><rect x="3" y="14" width="6" height="6" rx="1"/><line x1="12" y1="17" x2="21" y2="17"/>'),
  indent: svg('<line x1="11" y1="6" x2="21" y2="6"/><line x1="11" y1="12" x2="21" y2="12"/><line x1="11" y1="18" x2="21" y2="18"/><path d="M3 8l4 4-4 4"/>'),
  outdent: svg('<line x1="11" y1="6" x2="21" y2="6"/><line x1="11" y1="12" x2="21" y2="12"/><line x1="11" y1="18" x2="21" y2="18"/><path d="M7 8l-4 4 4 4"/>'),
  left: svg('<line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="15" y2="12"/><line x1="3" y1="18" x2="18" y2="18"/>'),
  center: svg('<line x1="3" y1="6" x2="21" y2="6"/><line x1="6" y1="12" x2="18" y2="12"/><line x1="4" y1="18" x2="20" y2="18"/>'),
  right: svg('<line x1="3" y1="6" x2="21" y2="6"/><line x1="9" y1="12" x2="21" y2="12"/><line x1="6" y1="18" x2="21" y2="18"/>'),
  justify: svg('<line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/>'),
  link: svg('<path d="M10 13a5 5 0 0 0 7.5.5l3-3a5 5 0 0 0-7-7l-1.7 1.7"/><path d="M14 11a5 5 0 0 0-7.5-.5l-3 3a5 5 0 0 0 7 7l1.7-1.7"/>'),
  image: svg('<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/>'),
  table: svg('<rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="12" y1="3" x2="12" y2="21"/>'),
  rule: svg('<line x1="3" y1="12" x2="21" y2="12"/>'),
  sortAsc: svg('<path d="M4 6h7M4 12h5M4 18h3"/><path d="M17 4v16M14 17l3 3 3-3"/>'),
  sortDesc: svg('<path d="M4 6h3M4 12h5M4 18h7"/><path d="M17 4v16M14 17l3 3 3-3"/>'),
  filter: svg('<path d="M3 4h18l-7 8v6l-4 2v-8z"/>'),
  freeze: svg('<rect x="3" y="3" width="18" height="18" rx="1"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="3" x2="9" y2="21"/>'),
  wrap: svg('<line x1="3" y1="6" x2="21" y2="6"/><path d="M3 12h15a3 3 0 0 1 0 6h-4"/><path d="M16 16l-2 2 2 2"/><line x1="3" y1="18" x2="10" y2="18"/>'),
  merge: svg('<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M8 12h8M10 10l-2 2 2 2M14 10l2 2-2 2"/>'),
  recalc: svg('<path d="M21 12a9 9 0 1 1-3-6.7L21 8"/><path d="M21 3v5h-5"/>'),
  pointer: svg('<path d="M5 3l14 8-6 2-2 6z"/>'),
  edittext: svg('<path d="M4 7V5h16v2M9 19h6M12 5v14"/><path d="M17 14l3 3-4 4h-3v-3z"/>'),
  highlighter: svg('<path d="M9 11l-5 5v4h4l5-5"/><path d="M15 5l4 4-6 6-4-4z"/>'),
  hibox: svg('<rect x="3" y="7" width="18" height="10" rx="1" fill="currentColor" fill-opacity=".25"/>'),
  pen: svg('<path d="M4 20c4-1 6-6 9-9s5-5 7-7"/><path d="M15 4l5 5"/>'),
  square: svg('<rect x="4" y="5" width="16" height="14" rx="1"/>'),
  circle: svg('<ellipse cx="12" cy="12" rx="9" ry="7"/>'),
  eraser: svg('<path d="M16 3l5 5-11 11H5l-2-2z"/><path d="M9 21h12"/>'),
  note: svg('<path d="M4 4h16v12l-4 4H4z"/><path d="M16 20v-4h4"/>'),
  signature: svg('<path d="M3 17c3-6 5-9 6-8s-2 7 0 7 3-5 5-5 1 4 3 4 2-2 4-2"/><path d="M3 21h18"/>'),
  rotateLeft: svg('<path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/>'),
  rotateRight: svg('<path d="M21 12a9 9 0 1 1-3-6.7L21 8"/><path d="M21 3v5h-5"/>'),
  pageAdd: svg('<path d="M14 3H6v18h12V7z"/><path d="M12 10v6M9 13h6"/>'),
  pagesFrom: svg('<path d="M14 3H8v14h10V7z"/><path d="M5 7v14h10"/>'),
  pageDel: svg('<path d="M14 3H6v18h12V7z"/><path d="M9 12l6 4M15 12l-6 4"/>'),
  copy: svg('<rect x="8" y="8" width="12" height="12" rx="1"/><path d="M16 8V5a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1v10a1 1 0 0 0 1 1h3"/>'),
  trash: svg('<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>'),
  pagebreak: svg('<path d="M5 3v5h14V3M5 21v-5h14v5"/><line x1="3" y1="12" x2="6" y2="12"/><line x1="9" y1="12" x2="15" y2="12"/><line x1="18" y1="12" x2="21" y2="12"/>'),
  clear: svg('<path d="M4 7V4h16v3M9 20h6M12 4v16"/><line x1="3" y1="21" x2="21" y2="3"/>'),
  quote: svg('<path d="M7 7H4v6h3l-2 4M17 7h-3v6h3l-2 4"/>'),
  code: svg('<path d="M8 6l-6 6 6 6M16 6l6 6-6 6"/>'),
  paint: svg('<rect x="4" y="3" width="14" height="6" rx="1"/><path d="M18 6h2v5h-8v3"/><rect x="10" y="14" width="4" height="7" rx="1"/>'),
  spacing: svg('<line x1="11" y1="6" x2="21" y2="6"/><line x1="11" y1="12" x2="21" y2="12"/><line x1="11" y1="18" x2="21" y2="18"/><path d="M5 4v16M3 6l2-2 2 2M3 18l2 2 2-2"/>'),
  comment: svg('<path d="M4 4h16v12H8l-4 4z"/><path d="M12 7v6M9 10h6"/>'),
  chart: svg('<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>'),
  fill: svg('<path d="M5 12l7-7 7 7-7 7z"/><path d="M19 15c0 1.5 1 2.5 1 3.5a1 1 0 0 1-2 0c0-1 1-2 1-3.5z"/>'),
  border: svg('<rect x="4" y="4" width="16" height="16" rx="1" stroke-dasharray="3 2"/>'),
  arrange: svg('<rect x="3" y="3" width="11" height="11" rx="1"/><rect x="10" y="10" width="11" height="11" rx="1" fill="currentColor" fill-opacity=".25"/>'),
  codeblock: svg('<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M10 9l-3 3 3 3M14 9l3 3-3 3"/>'),
};

// ======================================================================
// Wiring
// ======================================================================

desk.onOpenPath((p) => openPath(p));
desk.onCommand((cmd, arg) => command(cmd, arg));
function command(cmd, arg) {
  switch (cmd) {
    case 'replace': return openFind(true);
    case 'menuSearch': return Menus.focusSearch();
    case 'shortcuts': return showShortcuts();
    case 'wordCount': return showWordCount(active);
    case 'compact': return Menus.toggleCompact();
    case 'zoom': return zoomTab(active, arg);
    case 'backups': return desk.openBackups();
    case 'about': return showAbout();
    case 'new': return newDocument(arg);
    case 'open': return openDialog();
    case 'save': return save(active, false);
    case 'saveAs': return save(active, true);
    case 'exportPdf': return doExportPdf();
    case 'print': return doPrint();
    case 'reveal': return active && active.path && desk.reveal(active.path);
    case 'closeTab': return active && closeTab(active);
    case 'find': return openFind();
    case 'undo': case 'redo': {
      const sv = active && active.sheetView;
      if (sv && !sv.editing && !['INPUT', 'TEXTAREA'].includes((document.activeElement || {}).tagName)) return sv[cmd]();
      // Word: the editor's undo also reaches back past lists/tables/pictures (which re-render the page).
      if (active && active.kind === 'docx' && active.mode === 'edit' && active.editor) return active.editor[cmd]();
      if (active && active.kind === 'pptx' && active.mode === 'edit' && active.pptEditor && !(document.activeElement && ['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName))) return active.pptEditor[cmd]();
      if (active && active.kind === 'pdf' && active.mode === 'edit' && active.pdfEditor && !(document.activeElement && (document.activeElement.isContentEditable || ['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)))) return active.pdfEditor[cmd]();
      return document.execCommand(cmd);
    }
    case 'help': return showHelp();
    case 'toggleMode': {
      if (!active || active.modes.length < 2) return;
      return setMode(active, active.modes[(active.modes.indexOf(active.mode) + 1) % active.modes.length]);
    }
    case 'nextTab': {
      if (tabs.length < 2) return;
      return activate(tabs[(tabs.indexOf(active) + arg + tabs.length) % tabs.length]);
    }
  }
}

desk.onCloseRequest(async () => {
  const dirty = tabs.filter((t) => t.dirty);
  if (dirty.length) {
    const r = await desk.confirm({
      type: 'warning',
      message: dirty.length === 1 ? `Save changes to ${dirty[0].name}?` : `You have unsaved changes in ${dirty.length} documents.`,
      detail: dirty.length > 1 ? dirty.map((t) => `• ${t.name}`).join('\n') : "Your changes will be lost if you don't save them.",
      buttons: [dirty.length > 1 ? 'Save all' : 'Save', "Don't save", 'Cancel'], defaultId: 0, cancelId: 2,
    });
    if (r === 2) return;
    if (r === 0) {
      for (const t of dirty) { activate(t); if (!(await save(t))) return; }
    }
  }
  await Recovery.forgetAll();
  desk.closeConfirmed();
});


$('#find-input').addEventListener('input', () => runFind());
$('#find-input').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); stepFind(e.shiftKey ? -1 : 1); }
  if (e.key === 'Escape') { closeFind(); const ta = active && active.textarea; if (ta && !ta.closest('[data-mode="preview"],[data-mode="table"]')) ta.focus(); }
});
$('#find-next').addEventListener('click', () => stepFind(1));
for (const id of ['#find-case', '#find-regex']) $(id).addEventListener('click', () => { const b = $(id); b.setAttribute('aria-pressed', b.getAttribute('aria-pressed') === 'true' ? 'false' : 'true'); runFind(); $('#find-input').focus(); });
$('#find-toggle-replace').addEventListener('click', () => { const r = $('#replace-row'); r.hidden = !r.hidden; if (!r.hidden) $('#replace-input').focus(); });
$('#replace-one').addEventListener('click', replaceOne);
$('#replace-all').addEventListener('click', replaceAll);
$('#replace-input').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); if (e.ctrlKey) replaceAll(); else replaceOne(); }
  if (e.key === 'Escape') closeFind();
});
$('#find-prev').addEventListener('click', () => stepFind(-1));
$('#find-close').addEventListener('click', closeFind);
document.addEventListener('keydown', (e) => {
  if (e.key === 'F3') { e.preventDefault(); if ($('#findbar').hidden) openFind(); else stepFind(e.shiftKey ? -1 : 1); }
});

// Drag and drop files anywhere in the window.
let dragDepth = 0;
const hasFiles = (e) => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
document.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; e.preventDefault(); dragDepth++; $('#drop-hint').hidden = false; });
document.addEventListener('dragleave', (e) => { if (!hasFiles(e)) return; dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) $('#drop-hint').hidden = true; });
document.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
document.addEventListener('drop', async (e) => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth = 0;
  $('#drop-hint').hidden = true;
  for (const f of e.dataTransfer.files) {
    const p = desk.pathForFile(f);
    if (p) await openPath(p);
  }
});

refreshChrome();
