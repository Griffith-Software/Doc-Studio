'use strict';
// Per-type file icons in File Explorer (blue DOCX, green XLSX, red PDF...). Windows 11 records "open with Doc
// Studio" as the whole app, which has one icon, so Explorer is pointed at a small icon handler instead
// (iconhandler/IconHandler.cs -> resources\iconhandler\DocStudioIcons.dll) that answers with
// resources\filetypes\<ext>.ico for each file.
//
// File Explorer keeps the handler loaded while it runs, so its file can't be replaced during an update. Doc Studio
// therefore copies it to %LOCALAPPDATA%\Doc Studio Icons\<version>\ and registers it there for the current Windows
// user (HKCU; no admin). Earlier versions' folders are removed when they are no longer in use. The uninstaller
// removes the registration (build/installer.nsh).

const { app } = require('electron');
const { execFile } = require('child_process');
const fs = require('fs');
const path = require('path');

const CLSID = '{6C1E2B7A-4F3D-4D8E-9B21-5A7C0D3E9F41}';
const APP_KEY = 'HKCU\\Software\\Classes\\Applications\\Doc Studio.exe';
const CLSID_KEY = `HKCU\\Software\\Classes\\CLSID\\${CLSID}`;

const reg = (args) => new Promise((resolve) => execFile('reg.exe', args, { windowsHide: true }, (err, stdout) => resolve(err ? null : stdout)));
const regSet = (key, name, value) => reg(['add', key, ...(name === null ? ['/ve'] : ['/v', name]), '/t', 'REG_SZ', '/d', value, '/f']);
async function regGet(key, name) {
  const out = await reg(['query', key, ...(name === null ? ['/ve'] : ['/v', name])]);
  const m = out && /REG_SZ\s+(.*)$/m.exec(out);
  return m ? m[1].trim() : null;
}

async function register() {
  if (process.platform !== 'win32' || !app.isPackaged) return;
  const source = path.join(process.resourcesPath, 'iconhandler', 'DocStudioIcons.dll');
  const icons = path.join(process.resourcesPath, 'filetypes');
  if (!fs.existsSync(source) || !fs.existsSync(icons)) return;
  const base = path.join(process.env.LOCALAPPDATA || path.join(app.getPath('home'), 'AppData', 'Local'), 'Doc Studio Icons');
  const dir = path.join(base, app.getVersion());
  const dll = path.join(dir, 'DocStudioIcons.dll');
  fs.mkdirSync(dir, { recursive: true });
  if (!fs.existsSync(dll)) fs.copyFileSync(source, dll);
  const pointer = path.join(dir, 'filetypes.txt');
  if (!fs.existsSync(pointer) || fs.readFileSync(pointer, 'utf8').trim() !== icons) fs.writeFileSync(pointer, icons);
  fixAutoFiles(icons);
  const codeBase = `file:///${dll.replace(/\\/g, '/')}`;
  // Already registered for this copy? (The installer resets the app's icon, so this is checked every start.)
  if ((await regGet(`${CLSID_KEY}\\InprocServer32`, 'CodeBase')) === codeBase && (await regGet(`${APP_KEY}\\DefaultIcon`, null)) === '%1' && (await regGet(`${APP_KEY}\\shellex\\IconHandler`, null)) === CLSID) { cleanOld(base, dir); return; }
  const inproc = `${CLSID_KEY}\\InprocServer32`;
  await regSet(CLSID_KEY, null, 'Doc Studio icon handler');
  await regSet(inproc, null, 'mscoree.dll');
  await regSet(inproc, 'ThreadingModel', 'Both');
  await regSet(inproc, 'Class', 'DocStudio.IconHandler');
  await regSet(inproc, 'Assembly', 'DocStudioIcons, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null');
  await regSet(inproc, 'RuntimeVersion', 'v4.0.30319');
  await regSet(inproc, 'CodeBase', codeBase);
  await regSet('HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Shell Extensions\\Approved', CLSID, 'Doc Studio icon handler');
  await regSet(`${APP_KEY}\\shellex\\IconHandler`, null, CLSID);
  await regSet(`${APP_KEY}\\DefaultIcon`, null, '%1');
  // Tell File Explorer the icons changed (it picks the new ones up for windows opened from now on).
  execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', 'Add-Type -Namespace DS -Name N -MemberDefinition \'[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, uint f, System.IntPtr a, System.IntPtr b);\'; [DS.N]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)'], { windowsHide: true }, () => {});
  cleanOld(base, dir);
}
// When someone picks Doc Studio by browsing for it, Windows makes its own entry for the type (<ext>_auto_file)
// with no icon, so those files show the plain program icon. Give each such entry that opens Doc Studio its type's
// icon (only the picture changes, not which app opens the files). One PowerShell run, in the background.
function fixAutoFiles(icons) {
  const exts = fs.readdirSync(icons).filter((f) => f.endsWith('.ico')).map((f) => f.slice(0, -4));
  const script = `$d='${icons.replace(/'/g, "''")}'; foreach ($x in @(${exts.map((x) => `'${x}'`).join(',')})) { $k = "HKCU:\\Software\\Classes\\$($x)_auto_file"; $c = (Get-ItemProperty "$k\\shell\\open\\command" -EA 0).'(default)'; if ($c -and $c -match 'Doc Studio') { $i = "\`"$d\\$x.ico\`""; if ((Get-ItemProperty "$k\\DefaultIcon" -EA 0).'(default)' -ne $i) { New-Item "$k\\DefaultIcon" -Force | Out-Null; Set-ItemProperty "$k\\DefaultIcon" -Name '(default)' -Value $i } } }`;
  execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { windowsHide: true }, () => {});
}
// Remove earlier versions' copies (one still loaded in File Explorer can't be removed until Explorer restarts).
function cleanOld(base, keep) {
  for (const name of fs.readdirSync(base)) {
    const d = path.join(base, name);
    if (d !== keep) fs.promises.rm(d, { recursive: true, force: true }).catch(() => {});
  }
}

module.exports = { register: () => register().catch((err) => console.error('shell icons', err)) };
