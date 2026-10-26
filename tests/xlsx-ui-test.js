'use strict';
// Spreadsheet UI test with real keyboard and mouse input: `electron tests/xlsx-ui-test.js` (DOCSTUDIO_OUT=<dir>).

const { app, BrowserWindow, Menu } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');

const outDir = process.env.DOCSTUDIO_OUT;
app.setPath('userData', fs.mkdtempSync(path.join(os.tmpdir(), 'docstudio-xlui-')));
fs.mkdirSync(outDir, { recursive: true });
require('../main.js');

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const checks = [];
const check = (name, ok, detail) => checks.push({ name, ok: !!ok, detail });

app.whenReady().then(async () => {
  let win;
  while (!(win = BrowserWindow.getAllWindows()[0])) await wait(100);
  win.setSize(1300, 850);
  win.show(); win.focus();
  await new Promise((r) => (win.webContents.isLoading() ? win.webContents.once('did-finish-load', r) : r()));
  const wc = win.webContents;
  const js = (code) => wc.executeJavaScript(code);
  const key = async (keyCode, modifiers = []) => {
    wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
    if (keyCode.length === 1 && !modifiers.some((m) => m === 'control')) wc.sendInputEvent({ type: 'char', keyCode, modifiers });
    wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
    await wait(40);
  };
  const type = async (text) => { for (const ch of text) await key(ch, /[A-Z!*()=]/.test(ch) && !/[=]/.test(ch) ? ['shift'] : []); };
  const cellXY = (r, c) => js(`(() => { const v = active.sheetView; const g = v.grid.getBoundingClientRect(); return { x: Math.round(g.left + 46 + v.colLeft[${c}] - v.main.scrollLeft + 10), y: Math.round(g.top + 22 + v.rowTop[${r}] - v.main.scrollTop + 8) }; })()`);
  const click = async (r, c, extra = {}) => {
    const p = await cellXY(r, c);
    wc.sendInputEvent({ type: 'mouseDown', x: p.x, y: p.y, button: 'left', clickCount: extra.clickCount || 1, modifiers: extra.modifiers || [] });
    wc.sendInputEvent({ type: 'mouseUp', x: p.x, y: p.y, button: 'left', clickCount: extra.clickCount || 1, modifiers: extra.modifiers || [] });
    await wait(80);
  };
  const drag = async (r1, c1, r2, c2) => {
    const a = await cellXY(r1, c1), b = await cellXY(r2, c2);
    wc.sendInputEvent({ type: 'mouseDown', x: a.x, y: a.y, button: 'left', clickCount: 1 });
    wc.sendInputEvent({ type: 'mouseMove', x: (a.x + b.x) >> 1, y: (a.y + b.y) >> 1, button: 'left' });
    wc.sendInputEvent({ type: 'mouseMove', x: b.x, y: b.y, button: 'left' });
    wc.sendInputEvent({ type: 'mouseUp', x: b.x, y: b.y, button: 'left', clickCount: 1 });
    await wait(80);
  };
  const val = (a) => js(`(() => { const p = SheetModel.parseAddr(${JSON.stringify(a)}); const c = active.sheetView.sheet.get(p.r, p.c); return c ? (c.value && c.value.code ? c.value.code : c.value) : null; })()`);
  const formula = (a) => js(`(() => { const p = SheetModel.parseAddr(${JSON.stringify(a)}); const c = active.sheetView.sheet.get(p.r, p.c); return c ? c.formula : null; })()`);
  const activeAddr = () => js(`SheetModel.addr(active.sheetView.active.r, active.sheetView.active.c)`);

  await js(`newDocument('xlsx'); true`);
  await wait(1500);
  await click(1, 1);
  // Typing replaces the cell; Enter commits and moves down.
  await type('Apples'); await key('Return');
  check('typed text committed', (await val('A1')) === 'Apples', [await val('A1'), await val('A2'), await activeAddr()]);
  check('Enter moved down', (await activeAddr()) === 'A2');
  await type('5'); await key('Tab');
  check('Tab moved right', (await activeAddr()) === 'B2');
  await key('Left'); await key('Down');
  await type('7'); await key('Return');
  await type('2.5'); await key('Return');
  check('numbers typed', (await val('A2')) === 5 && (await val('A3')) === 7 && (await val('A4')) === 2.5);
  // Build a formula by clicking cells.
  await click(5, 1);
  await type('=SUM(');
  await drag(2, 1, 4, 1);
  await type(')');
  await key('Return');
  check('formula built by dragging a range', (await formula('A5')) === 'SUM(A2:A4)' && (await val('A5')) === 14.5, [await formula('A5'), await val('A5')]);
  await click(6, 1);
  await type('=');
  await click(2, 1);
  await type('*');
  await click(3, 1);
  await key('Return');
  check('formula built by clicking cells', (await formula('A6')) === 'A2*A3' && (await val('A6')) === 35, [await formula('A6'), await val('A6')]);
  // F2 edits the existing content.
  await click(1, 1);
  await key('F2'); await type('!'); await key('Return');
  check('F2 appends to existing text', (await val('A1')) === 'Apples!');
  // Escape cancels.
  await click(1, 1); await type('zzz'); await key('Escape');
  check('Escape cancels the edit', (await val('A1')) === 'Apples!');
  // Delete clears.
  await click(4, 1); await key('Delete');
  check('Delete clears and dependents recalc', (await val('A4')) === null && (await val('A5')) === 12);
  // Ctrl+Z through the menu accelerator restores.
  await key('z', ['control']);
  await wait(100);
  check('Ctrl+Z restores', (await val('A4')) === 2.5 && (await val('A5')) === 14.5, [await val('A4'), await val('A5')]);
  // Copy A6 (formula) and paste into B6: references shift.
  await click(6, 1);
  wc.copy(); await wait(150);
  await click(6, 2);
  wc.paste(); await wait(200);
  check('copy/paste shifts formula references', (await formula('B6')) === 'B2*B3', await formula('B6'));
  // Ctrl+B on a range.
  await drag(1, 1, 2, 2);
  await key('b', ['control']);
  check('Ctrl+B bolds the selection', await js(`active.workbook.styles.font(active.sheetView.sheet.get(2,2) ? active.sheetView.sheet.get(2,2).s : 0).bold && active.workbook.styles.font(active.sheetView.sheet.get(1,1).s).bold`));
  // Status bar sum.
  await drag(2, 1, 4, 1);
  check('status shows Sum of selection', (await js(`document.querySelector('#status-left').textContent`)).includes('Sum: 14.5'));
  // Double-click to edit.
  await click(3, 1, { clickCount: 2 });
  check('double-click opens the editor', await js(`!!active.sheetView.editing && !active.sheetView.editor.hidden`));
  await key('Escape');
  // Arrows after typing commit and move (Enter mode).
  await click(8, 1); await type('x'); await key('Right');
  check('arrow commits typed entry and moves', (await val('A8')) === 'x' && (await activeAddr()) === 'B8');
  // Ctrl+Down jumps to the edge of data.
  await click(1, 1); await key('Down', ['control']);
  check('Ctrl+Down jumps to end of block', (await activeAddr()) === 'A6', await activeAddr());
  fs.writeFileSync(path.join(outDir, 'ui.png'), (await wc.capturePage()).toPNG());

  const summary = { passed: checks.filter((x) => x.ok).length, failed: checks.filter((x) => !x.ok) };
  console.log(JSON.stringify(summary, null, 2));
  app.exit(0);
}).catch((e) => { console.log('ERROR', e); app.exit(1); });
