'use strict';
// Shell features shared by every document: zoom, rename, make a copy, version history, details, word count,
// keyboard shortcuts, spelling, and the toolbar's common buttons (undo, redo, print, zoom).

// ---------------------------------------------------------------- zoom (the document, not the app's own controls)

const ZOOMS = [0.5, 0.67, 0.75, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];
function canZoom(tab) { return !!tab && ['docx', 'odt', 'rtf', 'md', 'html', 'text', 'csv', 'pptx', 'pdf', 'epub', 'doc', 'odp', 'image', 'sheetview', 'sheet'].includes(tab.kind) && !(tab.kind === 'pdf' && !tab.pdfEditor); }
function zoomTab(tab, dir, exact) {
  if (!canZoom(tab)) { flash('Zoom isn\'t available for this view yet.'); return; }
  const cur = tab.zoom || 1;
  let z = exact || (dir === 0 ? 1 : dir > 0 ? ZOOMS.find((x) => x > cur + 0.001) || cur : [...ZOOMS].reverse().find((x) => x < cur - 0.001) || cur);
  if (tab.kind === 'sheet' && tab.sheetView) { z = Math.max(0.5, Math.min(2, z)); tab.zoom = z; tab.sheetView.setZoom(z); refreshChrome(); return; }
  if (tab.kind === 'pdf' && tab.pdfEditor) { tab.pdfEditor.setZoom(tab.pdfEditor.scale * z / cur); tab.zoom = z; refreshChrome(); return; }
  tab.zoom = z;
  applyZoom(tab);
  refreshChrome();
}
function applyZoom(tab) {
  const z = tab.zoom || 1;
  const set = (el) => { if (el) el.style.zoom = z === 1 ? '' : String(z); };
  if (tab.textarea) {
    const px = `${(14 * z).toFixed(2)}px`;
    for (const el of [tab.textarea, tab.backdrop, tab.measure]) if (el) el.style.fontSize = z === 1 ? '' : px;
    set(tab.previewPane && tab.previewPane.firstElementChild);
    if (tab.remap) tab.remap();
    return;
  }
  if (tab.kind === 'docx') { set(tab.viewBox && tab.viewBox.querySelector('.docx-view')); set(tab.editHost); return; }
  if (tab.richEditor) { set(tab.richEditor.host.firstElementChild); return; }
  if (tab.kind === 'pptx') { set(tab.viewBox); if (tab.pptEditor) { tab.pptEditor.zoom = z; tab.pptEditor.renderCurrent(); } return; }
  set(tab.body.querySelector('.view-box') && tab.body.querySelector('.view-box').firstElementChild);
}

// ---------------------------------------------------------------- toolbar: undo, redo, print, zoom

function commonTools(tab) {
  if (!tab) return [];
  const btn = (icon, title, fn, enabled) => {
    const b = h('button', { class: 'tool', title, 'aria-label': title });
    b.innerHTML = icon;
    b.disabled = enabled === false;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', fn);
    return b;
  };
  const editable = FORMATS.isEditable(tab.kind) && (tab.kind === 'sheet' || ['edit', 'split'].includes(tab.mode));
  // Search, print and zoom live in the app bar and the page bar.
  return [
    btn(SHELL_ICONS.undo, 'Undo (Ctrl+Z)', () => command('undo'), editable),
    btn(SHELL_ICONS.redo, 'Redo (Ctrl+Y)', () => command('redo'), editable),
    sep(),
  ];
}

const SHELL_ICONS = {
  undo: svg('<path d="M9 14L4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>'),
  redo: svg('<path d="M15 14l5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>'),
  print: svg('<path d="M6 9V3h12v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>'),
  search: svg('<circle cx="11" cy="11" r="7"/><path d="M21 21l-5-5"/>'),
};

// ---------------------------------------------------------------- rename, copy, versions, details

async function renameTab(tab) {
  if (!tab) return;
  const ext = tab.name.includes('.') ? tab.name.slice(tab.name.lastIndexOf('.')) : '';
  const v = await ask('Rename', [{ name: 'name', label: 'Name', value: stem(tab.name) }]);
  if (!v || !v.name.trim()) return;
  const newName = v.name.trim().endsWith(ext) ? v.name.trim() : v.name.trim() + ext;
  if (newName === tab.name) return;
  try {
    if (tab.path) tab.path = await desk.rename(tab.path, newName);
    tab.name = newName;
    refreshChrome();
    flash(`Renamed to ${newName}`);
  } catch (err) { desk.error("Couldn't rename the file", err.message); }
}

// The document as it is now, unsaved edits included, in the format its editor works in.
async function tabBytes(tab) {
  if (!(hasEditor(tab) && FORMATS.isEditable(tab.kind))) return tab.data;
  if (tab.kind === 'docx') return currentDocxBytes(tab);
  if (tab.kind === 'sheet') { tab.sheetView.endEdit(true); return tab.workbook.save(); }
  if (tab.kind === 'pptx') return tab.pptEditor.save();
  if (tab.kind === 'odt') return tab.odtEditor.save();
  if (tab.kind === 'rtf') return tab.rtfEditor.save();
  if (tab.kind === 'pdf') return tab.pdfEditor ? tab.pdfEditor.build() : tab.data;
  if (tab.textarea) return TextCodec.encode(tab.textarea.value, tab.codec);
  return tab.data;
}

// A copy of the document as it is now (with unsaved edits), saved under a new name; the open file is unchanged.
async function makeCopy(tab) {
  if (!tab) return;
  let bytes;
  try { bytes = await tabBytes(tab); } catch (err) { desk.error("Couldn't make a copy", err.message); return; }
  const ext = tab.kind === 'sheet' ? nativeExt(tab) : (tab.ext || 'bin');
  const dir = tab.path ? tab.path.slice(0, Math.max(tab.path.lastIndexOf('\\'), tab.path.lastIndexOf('/')) + 1) : '';
  const target = await desk.saveDialog(`${dir}Copy of ${stem(tab.name)}.${ext}`, ext);
  if (!target) return;
  await desk.write(target, bytes);
  const r = await desk.confirm({ type: 'info', message: `Saved a copy as ${target.split(/[\\/]/).pop()}`, buttons: ['Open the copy', 'Close'], defaultId: 1, cancelId: 1 });
  if (r === 0) openPath(target);
}

async function showVersions(tab) {
  if (!tab || !tab.path) return;
  const list = await desk.versions(tab.path);
  const box = h('div', { class: 'dialog versions-dialog' }, h('h3', { text: `Version history - ${tab.name}` }));
  const back = h('div', { class: 'backdrop' }, box);
  const close = () => back.remove();
  if (!list.length) box.append(h('p', { text: 'No saved versions yet. Doc Studio keeps a version each time the file is saved (the newest 40).' }));
  else {
    const rows = h('div', { class: 'version-list' });
    for (const v of list) {
      const when = new Date(v.time);
      const label = isNaN(when) ? v.id : when.toLocaleString();
      rows.append(h('div', { class: 'version-row' },
        h('span', { class: 'version-time', text: label }), h('span', { class: 'muted', text: formatBytes(v.size) }),
        h('button', { class: 'ghost', text: 'Open', onclick: async () => { close(); await openVersion(tab, v, label); } }),
        h('button', { class: 'ghost', text: 'Restore', onclick: async () => { close(); await restoreVersion(tab, v, label); } })));
    }
    box.append(h('p', { class: 'muted', text: 'Open shows a version in a new tab; Restore saves it over the current file (the current file is kept as a version too).' }), rows);
  }
  box.append(h('div', { class: 'dialog-actions' }, h('button', { class: 'ghost', text: 'Close', onclick: close })));
  back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
  back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
  document.body.append(back);
  box.querySelector('button').focus();
}
async function openVersion(tab, v, label) {
  const data = await desk.readVersion(tab.path, v.id);
  const t = makeTab({ path: null, name: `${stem(tab.name)} (${label.replace(/[\\/:]/g, '-')}).${tab.ext}`, size: data.byteLength, data });
  try { await setupTab(t); } catch (err) { t.body.replaceChildren(failure(t, err)); }
  refreshChrome();
}
async function restoreVersion(tab, v, label) {
  if (tab.dirty) {
    const r = await desk.confirm({ type: 'warning', message: `Restore the version from ${label}?`, detail: 'Your unsaved changes will be lost.', buttons: ['Restore', 'Cancel'], defaultId: 1, cancelId: 1 });
    if (r !== 0) return;
  }
  const data = await desk.readVersion(tab.path, v.id);
  await desk.write(tab.path, data);
  const path = tab.path;
  tab.dirty = false;
  await closeTab(tab);
  await openPath(path);
  flash(`Restored the version from ${label}`);
}

async function showDetails(tab) {
  if (!tab) return;
  const rows = [['Name', tab.name], ['Type', FORMATS.label(tab.kind)]];
  if (tab.path) {
    rows.push(['Location', tab.path.slice(0, Math.max(tab.path.lastIndexOf('\\'), tab.path.lastIndexOf('/')))]);
    try {
      const st = await desk.stat(tab.path);
      rows.push(['Size', formatBytes(st.size)], ['Created', new Date(st.created).toLocaleString()], ['Modified', new Date(st.modified).toLocaleString()]);
    } catch { /* file moved */ }
  } else rows.push(['Location', 'Not saved yet']);
  const counts = wordCounts(tab);
  if (counts) rows.push(['Words', counts.words.toLocaleString()], ['Characters', counts.chars.toLocaleString()]);
  if (tab.info) rows.push(['Status', tab.info]);
  if (tab.dirty) rows.push(['Changes', 'Unsaved changes']);
  desk.confirm({ type: 'info', buttons: ['OK'], message: 'Document details', detail: rows.map(([k, v]) => `${k}: ${v}`).join('\n') });
}

// ---------------------------------------------------------------- word count

function documentText(tab) {
  if (!tab) return '';
  if (tab.textarea) return tab.textarea.value;
  if (tab.kind === 'docx' && tab.editor && tab.mode === 'edit') return tab.editor.articles.map((a) => a.innerText).join('\n');
  if (tab.richEditor) return tab.richEditor.article.innerText;
  if (tab.kind === 'pptx') { const root = tab.mode === 'edit' && tab.pptEditor ? null : tab.viewBox; return root ? root.innerText : (tab.pptEditor.deck.slides.map((s) => s.doc.documentElement.textContent).join('\n')); }
  const root = tab.searchRoot && tab.searchRoot();
  return root ? (root.value != null ? root.value : root.innerText) : '';
}
function wordCounts(tab, selectionOnly) {
  if (!tab || ['sheet', 'image', 'audio', 'video'].includes(tab.kind)) return null;
  let text = documentText(tab);
  if (selectionOnly) {
    const ta = tab.textarea;
    text = ta && document.activeElement === ta ? ta.value.slice(ta.selectionStart, ta.selectionEnd) : getSelection().toString();
  }
  const words = countWords(text);
  return { words, chars: text.replace(/\n/g, '').length, charsNoSpaces: text.replace(/\s/g, '').length, paragraphs: text.split(/\n+/).filter((x) => x.trim()).length };
}
function showWordCount(tab) {
  const all = wordCounts(tab);
  if (!all) { flash('Word count isn\'t available for this file.'); return; }
  const sel = wordCounts(tab, true);
  const line = (k, a, b) => `${k}: ${a.toLocaleString()}${sel && sel.words ? `   (selection: ${b.toLocaleString()})` : ''}`;
  desk.confirm({ type: 'info', buttons: ['OK'], message: 'Word count', detail: [
    tab.kind === 'docx' && tab.viewInfo ? `Pages: ${String(tab.viewInfo).replace(/\D+$/, '').replace(/^\D+/, '') || tab.viewInfo}` : null,
    line('Words', all.words, sel && sel.words), line('Characters', all.chars, sel && sel.chars),
    line('Characters excluding spaces', all.charsNoSpaces, sel && sel.charsNoSpaces), `Paragraphs: ${all.paragraphs.toLocaleString()}`,
  ].filter(Boolean).join('\n') });
}

// ---------------------------------------------------------------- spelling

let spelling = true;
function spellingOn() { return spelling; }
function toggleSpelling() {
  spelling = !spelling;
  document.querySelectorAll('textarea, [contenteditable="true"]').forEach((el) => { el.spellcheck = spelling; });
  desk.setSpellcheck(spelling);
  flash(spelling ? 'Spelling underlines on - right-click a word for suggestions' : 'Spelling underlines off');
}

// ---------------------------------------------------------------- keyboard shortcuts

function showShortcuts() {
  const groups = [
    ['General', [['Search the menus', 'Alt+/'], ['Open', 'Ctrl+O'], ['Save', 'Ctrl+S'], ['Save as', 'Ctrl+Shift+S'], ['Print', 'Ctrl+P'], ['Export as PDF', 'Ctrl+Shift+E'], ['Close tab', 'Ctrl+W'], ['Next / previous tab', 'Ctrl+Tab / Ctrl+Shift+Tab'], ['Switch View / Edit', 'Ctrl+E'], ['Full screen', 'F11'], ['Compact mode', 'Ctrl+Shift+F'], ['Open a menu', 'Alt + first letter']]],
    ['Editing', [['Undo / redo', 'Ctrl+Z / Ctrl+Y'], ['Cut / copy / paste', 'Ctrl+X / C / V'], ['Paste without formatting', 'Ctrl+Shift+V'], ['Select all', 'Ctrl+A'], ['Find', 'Ctrl+F'], ['Find next / previous', 'F3 / Shift+F3'], ['Find and replace', 'Ctrl+H'], ['Word count', 'Ctrl+Shift+C']]],
    ['Text', [['Bold / italic / underline', 'Ctrl+B / I / U'], ['Link', 'Ctrl+K'], ['Zoom in / out / 100%', 'Ctrl+= / Ctrl+- / Ctrl+0']]],
    ['Spreadsheets', [['Edit cell', 'F2 or double-click'], ['Absolute reference', 'F4'], ['Fill down / right', 'Ctrl+D / Ctrl+R'], ['Jump to edge', 'Ctrl+arrow'], ['Recalculate', 'F9']]],
    ['Slides', [['Next / previous slide', 'Page Down / Page Up'], ['Move selected item', 'Arrows (Shift = 10x)'], ['Delete selected item', 'Delete']]],
  ];
  const box = h('div', { class: 'dialog shortcuts-dialog' }, h('h3', { text: 'Keyboard shortcuts' }),
    h('div', { class: 'shortcut-cols' }, groups.map(([g, list]) => h('div', null, h('h4', { text: g }), h('table', null, list.map(([a, k]) => h('tr', null, h('td', { text: a }), h('td', null, h('kbd', { text: k })))))))));
  const back = h('div', { class: 'backdrop' }, box);
  const close = () => back.remove();
  box.append(h('div', { class: 'dialog-actions' }, h('button', { text: 'Close', onclick: close })));
  back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
  back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
  document.body.append(back);
  box.querySelector('button').focus();
}

async function showAbout() {
  const v = await desk.appVersion();
  desk.confirm({ type: 'info', buttons: ['OK'], message: `Doc Studio ${v}`, title: 'About Doc Studio', detail: 'View, edit, create and convert documents, spreadsheets, presentations and PDFs, keeping each file\'s own formatting.\n\nDoc Studio by Griffith Software' });
}
