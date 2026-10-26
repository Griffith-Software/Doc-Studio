'use strict';
// Writes renderer/vendor/THIRD-PARTY-NOTICES.txt for the packages bundled into renderer/vendor/formula.js.
// Run by `npm run bundle` (esbuild writes the metafile this reads).
const fs = require('fs');
const path = require('path');

const metaPath = process.argv[2] || 'build/formula-meta.json';
const which = process.argv[3] || 'formula';
const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
// Package name -> its folder (nested copies such as pdf-lib/node_modules/tslib are read from where they are).
const pkgs = new Map();
for (const input of Object.keys(meta.inputs)) {
  const parts = input.split(/[\\/]/);
  const i = parts.lastIndexOf('node_modules');
  if (i < 0) continue;
  const n = parts[i + 1].startsWith('@') ? 2 : 1;
  const name = parts.slice(i + 1, i + 1 + n).join('/');
  if (!pkgs.has(name)) pkgs.set(name, parts.slice(0, i + 1 + n).join('/'));
}
let out = `Doc Studio includes the following open-source packages in renderer/vendor/${which}.js.\n\n`;
for (const [name, dir] of [...pkgs].sort((a, b) => a[0].localeCompare(b[0]))) {
  const pj = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8'));
  const lic = fs.readdirSync(dir).find((f) => /^licen[cs]e/i.test(f));
  out += `${'='.repeat(70)}\n${name} ${pj.version} - ${pj.license || 'see below'}\n${'='.repeat(70)}\n`;
  out += lic ? `${fs.readFileSync(path.join(dir, lic), 'utf8').trim()}\n\n` : `(${pj.license || 'licence'} - no licence file shipped in the package)\n\n`;
}
fs.writeFileSync(path.join('renderer', 'vendor', which === 'formula' ? 'THIRD-PARTY-NOTICES.txt' : `THIRD-PARTY-NOTICES-${which}.txt`), out);
fs.unlinkSync(metaPath);
console.log(`notices: ${[...pkgs.keys()].join(', ')}`);
