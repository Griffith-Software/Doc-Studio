'use strict';
// Auto-save for recovery: unsaved work survives a crash or a power cut. Every few seconds, each document with new
// unsaved edits is copied into Doc Studio's recovery folder (main.js 'recovery:*'); the document's own file is not
// touched. Saving, closing or discarding the document removes its copy. On the next start, copies left behind
// (Doc Studio didn't get to close normally) are offered back.

const Recovery = (() => {
  const SESSION = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const EVERY = 5000;
  // Kinds whose content Doc Studio can rebuild into a file.
  const KINDS = new Set(['docx', 'sheet', 'pptx', 'odt', 'rtf', 'pdf', 'md', 'html', 'csv', 'text', 'image']);
  const keyOf = (tab) => `${SESSION}-${tab.id}`;
  const kept = new Set();
  let busy = false;

  // built: the document is open in its editor (its content is rebuilt); otherwise its bytes are copied as they are
  // (a new, unsaved document that has not been switched to Edit, such as an OCR result).
  async function snapshot(tab, built) {
    const seq = tab.editSeq;
    const bytes = built ? await nativeBytes(tab, false) : tab.data;
    if (!tabs.includes(tab) || !tab.dirty) return;
    await desk.recoverySave(keyOf(tab), { path: tab.path || null, name: tab.name, kind: tab.kind, isNew: !!tab.isNew }, bytes);
    tab.snapSeq = seq;
    kept.add(keyOf(tab));
  }
  async function tick() {
    if (busy) return;
    busy = true;
    try {
      for (const tab of [...tabs]) {
        if (!tab.dirty || tab.snapSeq === tab.editSeq) continue;
        const built = KINDS.has(tab.kind) && hasEditor(tab) && (['docx', 'sheet', 'pptx', 'odt', 'rtf', 'pdf', 'image'].includes(tab.kind) || !!tab.textarea);
        if (!built && !tab.data) continue;
        // Building a slide deck ends text editing, so wait until the text box being typed in is finished; same for
        // a PDF text box.
        if (built && tab.kind === 'pptx' && tab.pptEditor.editingEl) continue;
        if (built && tab.kind === 'pdf' && document.activeElement && document.activeElement.isContentEditable && tab.el.contains(document.activeElement)) continue;
        try { await snapshot(tab, built); } catch (err) { console.warn(`Recovery copy of ${tab.name} failed`, err); }
      }
    } finally { busy = false; }
  }
  // The document was saved, closed or discarded: its copy is no longer needed.
  function forget(tab) {
    const key = keyOf(tab);
    tab.snapSeq = undefined;
    if (!kept.delete(key)) return Promise.resolve();
    return desk.recoveryClear(key).catch(() => {});
  }
  function forgetAll() { const keys = [...kept]; kept.clear(); tabs.forEach((t) => { t.snapSeq = undefined; }); return keys.length ? desk.recoveryClear(keys).catch(() => {}) : Promise.resolve(); }

  async function restore(items) {
    for (const it of items) {
      const tab = makeTab({ path: it.meta.path, name: it.meta.name, size: it.data.byteLength, data: it.data });
      tab.kind = it.meta.kind;
      tab.isNew = !!it.meta.isNew;
      const n = /^Untitled (\d+)\./.exec(tab.name);
      if (n && tab.isNew) untitledSeq = Math.max(untitledSeq, +n[1]);
      try { await setupTab(tab); } catch (err) { console.error(err); tab.body.replaceChildren(failure(tab, err)); tab.modes = []; }
      markDirty(tab);
    }
    refreshChrome();
    await tick();
  }
  async function offer() {
    let items = [];
    try { items = await desk.recoveryTake(); } catch { return; }
    if (!items.length) return;
    const when = (t) => new Date(t).toLocaleString([], { dateStyle: 'medium', timeStyle: 'short' });
    const box = h('div', { class: 'dialog recovery-dialog' },
      h('h3', { text: 'Recover unsaved work' }),
      h('p', { text: `Doc Studio closed before ${items.length === 1 ? 'this document was' : 'these documents were'} saved, for example because the power went out. Restore ${items.length === 1 ? 'it' : 'them'} to keep working, then save.` }),
      h('div', { class: 'recovery-list' }, ...items.map((it) => h('div', { class: 'recovery-row' },
        h('b', { text: it.meta.name }),
        h('small', { text: `${it.meta.path || 'Never saved'} · last copied ${when(it.meta.time)}` })))));
    const back = h('div', { class: 'backdrop' }, box);
    const keys = items.map((x) => x.key);
    const discard = h('button', { type: 'button', class: 'ghost', text: 'Discard' });
    const ok = h('button', { type: 'button', text: items.length === 1 ? 'Restore' : `Restore all ${items.length}` });
    box.append(h('div', { class: 'dialog-actions' }, discard, ok));
    discard.addEventListener('click', async () => {
      const r = await desk.confirm({ type: 'warning', message: `Discard ${items.length === 1 ? items[0].meta.name : `${items.length} unsaved documents`}?`, detail: "The unsaved changes can't be brought back afterwards.", buttons: ['Discard', 'Cancel'], defaultId: 1, cancelId: 1 });
      if (r !== 0) return;
      back.remove();
      await desk.recoveryClear(keys);
    });
    ok.addEventListener('click', async () => {
      back.remove();
      await restore(items);
      // The restored documents now have copies of their own, so the old ones can go.
      await desk.recoveryClear(keys);
      flash(`Restored ${items.length === 1 ? items[0].meta.name : `${items.length} documents`}. Save to keep the changes.`);
    });
    document.body.append(back);
    ok.focus();
  }

  setInterval(tick, EVERY);
  // The AI worker window (index.html?worker=1) never offers recovery; the owner's window does.
  if (!/[?&]worker=1/.test(location.search)) setTimeout(offer, 600);
  return { tick, forget, forgetAll, offer };
})();
