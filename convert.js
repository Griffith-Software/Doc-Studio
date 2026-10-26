'use strict';
// Format conversion through the office programs installed on this PC (main process only).
//
// Microsoft Office (Word, Excel, PowerPoint) is driven over COM by a small PowerShell script; LibreOffice is run
// headless. Files are opened read-only with macros disabled, converted into a private temp folder and read back.

const { execFile } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

// Office SaveAs format codes.
const WORD = { docx: 16, doc: 0, rtf: 6, odt: 23, pdf: 17, txt: 7, html: 10 };
const EXCEL = { xlsx: 51, xlsm: 52, xls: 56, xlsb: 50, ods: 60, csv: 62, pdf: -1 };
const POWERPOINT = { pptx: 24, ppt: 1, odp: 35, pdf: 32 };
const APP_FOR = {
  docx: 'word', docm: 'word', dotx: 'word', doc: 'word', dot: 'word', rtf: 'word', odt: 'word', ott: 'word', txt: 'word', html: 'word', htm: 'word',
  xlsx: 'excel', xlsm: 'excel', xls: 'excel', xlsb: 'excel', ods: 'excel', fods: 'excel', csv: 'excel',
  pptx: 'powerpoint', pptm: 'powerpoint', ppsx: 'powerpoint', ppt: 'powerpoint', pps: 'powerpoint', pot: 'powerpoint', odp: 'powerpoint',
};

const SCRIPT = String.raw`param([string]$App, [string]$In, [string]$Out, [int]$Format)
$ErrorActionPreference = 'Stop'
switch ($App) {
  'word' {
    $w = New-Object -ComObject Word.Application
    try {
      $w.Visible = $false; $w.DisplayAlerts = 0; $w.AutomationSecurity = 3
      $d = $w.Documents.Open($In, $false, $true, $false, '', '', $true, '', '', 0, 0, $false, $false)
      $d.SaveAs2($Out, $Format)
      $d.Close(0)
    } finally { $w.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w) }
  }
  'excel' {
    $x = New-Object -ComObject Excel.Application
    try {
      $x.Visible = $false; $x.DisplayAlerts = $false; $x.AutomationSecurity = 3; $x.AskToUpdateLinks = $false
      $b = $x.Workbooks.Open($In, 0, $true)
      if ($Format -eq -1) { $b.ExportAsFixedFormat(0, $Out) } else { $b.SaveAs($Out, $Format) }
      $b.Close($false)
    } finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
  }
  'powerpoint' {
    $p = New-Object -ComObject PowerPoint.Application
    try {
      $p.AutomationSecurity = 3
      $s = $p.Presentations.Open($In, -1, 0, 0)
      $s.SaveAs($Out, $Format)
      $s.Close()
    } finally { $p.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($p) }
  }
}
'ok'
`;

let toolsCache = null;

function run(file, args, timeout) {
  return new Promise((resolve, reject) => {
    execFile(file, args, { windowsHide: true, timeout: timeout || 180000, maxBuffer: 4 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) {
        const lines = (stderr || '').split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
        reject(new Error(lines.find((l) => /exception|error|cannot|failed/i.test(l) && !/^\+|CategoryInfo|FullyQualified/.test(l)) || lines[0] || err.message));
      } else resolve(stdout);
    });
  });
}

async function hasProgId(progId) {
  try { await run('reg.exe', ['query', `HKCR\\${progId}\\CLSID`], 15000); return true; } catch { return false; }
}

function libreOfficePath() {
  const dirs = [process.env.ProgramFiles, process.env['ProgramFiles(x86)'], process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, 'Programs')].filter(Boolean);
  for (const d of dirs) {
    const p = path.join(d, 'LibreOffice', 'program', 'soffice.exe');
    if (fs.existsSync(p)) return p;
  }
  return null;
}

async function tools() {
  if (toolsCache) return toolsCache;
  if (process.platform !== 'win32') return (toolsCache = { word: false, excel: false, powerpoint: false, libreoffice: !!libreOfficePath() });
  const [word, excel, powerpoint] = await Promise.all(['Word.Application', 'Excel.Application', 'PowerPoint.Application'].map(hasProgId));
  toolsCache = { word, excel, powerpoint, libreoffice: !!libreOfficePath() };
  return toolsCache;
}

// Which program (if any) can turn `from` into `to`.
async function route(from, to) {
  const t = await tools();
  const app = APP_FOR[from];
  const table = { word: WORD, excel: EXCEL, powerpoint: POWERPOINT }[app];
  if (app && t[app] && table && table[to] !== undefined && (APP_FOR[to] === app || to === 'pdf')) return app;
  if (t.libreoffice) return 'libreoffice';
  return null;
}

// One conversion at a time (Office refuses a second automation start while one is closing), with one retry.
let queue = Promise.resolve();
function convert(data, from, to) {
  const job = queue.then(() => convertOnce(data, from, to).catch(async (err) => {
    await new Promise((r) => setTimeout(r, 2500));
    try { return await convertOnce(data, from, to); } catch { throw err; }
  }));
  queue = job.catch(() => {});
  return job;
}

async function convertOnce(data, from, to) {
  from = from.toLowerCase(); to = to.toLowerCase();
  const via = await route(from, to);
  if (!via) throw new Error(`No installed program can convert .${from} to .${to}.`);
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'docstudio-conv-'));
  try {
    const src = path.join(dir, `source.${from}`);
    await fs.promises.writeFile(src, Buffer.from(data));
    let out = path.join(dir, `result.${to}`);
    if (via === 'libreoffice') {
      const outDir = path.join(dir, 'out');
      await fs.promises.mkdir(outDir);
      await run(libreOfficePath(), ['--headless', '--norestore', '--nolockcheck', '--convert-to', to, '--outdir', outDir, src]);
      out = path.join(outDir, `source.${to}`);
    } else {
      const table = { word: WORD, excel: EXCEL, powerpoint: POWERPOINT }[via];
      const script = path.join(dir, 'convert.ps1');
      await fs.promises.writeFile(script, '﻿' + SCRIPT, 'utf8');
      await run('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, '-App', via, '-In', src, '-Out', out, '-Format', String(table[to])]);
    }
    const buf = await fs.promises.readFile(out);
    return { data: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), via };
  } finally {
    fs.promises.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

module.exports = { tools, route, convert };
