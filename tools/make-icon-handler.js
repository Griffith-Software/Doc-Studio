#!/usr/bin/env node
'use strict';
// Builds build/iconhandler/DocStudioIcons.dll (the File Explorer icon handler, iconhandler/IconHandler.cs) with
// the C# compiler that comes with Windows (.NET Framework 4), so nothing has to be installed. Shipped in
// resources\iconhandler; Doc Studio copies it to a versioned folder and registers it (ai/../main.js iconHandler).
//   node tools/make-icon-handler.js
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const csc = path.join(process.env.WINDIR || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');
const out = path.join(root, 'build', 'iconhandler');
fs.mkdirSync(out, { recursive: true });
if (!fs.existsSync(csc)) {
  // Not Windows (or no .NET Framework): keep a previously built handler if there is one.
  if (fs.existsSync(path.join(out, 'DocStudioIcons.dll'))) { console.log('icon handler: csc not found, keeping the built DLL'); process.exit(0); }
  console.error('icon handler: the .NET Framework C# compiler was not found'); process.exit(1);
}
execFileSync(csc, ['/nologo', '/target:library', '/optimize+', '/platform:anycpu', `/out:${path.join(out, 'DocStudioIcons.dll')}`, path.join(root, 'iconhandler', 'IconHandler.cs')], { stdio: 'inherit' });
console.log('icon handler -> build/iconhandler/DocStudioIcons.dll');
