'use strict';
// Checks that the packaged app.asar contains every script the renderer loads and the main-process converters.
// Run after `npm run dist`: node tests/check-package.js
const fs = require('fs');
const path = require('path');
const asar = require('@electron/asar');

const root = path.join(__dirname, '..');
const archive = path.join(root, 'dist', 'win-unpacked', 'resources', 'app.asar');
const files = new Set(asar.listPackage(archive).map((f) => f.replace(/\\/g, '/')));
const html = fs.readFileSync(path.join(root, 'renderer', 'index.html'), 'utf8');
const needed = [...html.matchAll(/src="\.\.\/(node_modules\/[^"]+)"/g)].map((m) => '/' + m[1]);
needed.push('/node_modules/word-extractor/package.json', '/node_modules/docx/dist/html.cjs', '/node_modules/docx/dist/index.cjs', '/convert.js', '/renderer/templates/blank.pptx');
// Our own renderer scripts (including the bundled formula engine).
needed.push(...[...html.matchAll(/src="(?!\.\.)([^"]+\.js)"/g)].map((m) => '/renderer/' + m[1]));
let missing = 0;
for (const f of needed) {
  const ok = files.has(f);
  if (!ok) missing++;
  console.log(`${ok ? 'OK     ' : 'MISSING'} ${f}`);
}
// Build-time-only packages must not ship.
for (const dev of ['/node_modules/fast-formula-parser/', '/node_modules/@formulajs/', '/node_modules/esbuild/', '/src/', '/tests/', '/packaging/', '/release/']) {
  const leaked = [...files].some((f) => f.startsWith(dev));
  if (leaked) missing++;
  console.log(`${leaked ? 'SHIPPED' : 'OK     '} (not packaged) ${dev}`);
}
process.exit(missing ? 1 : 0);
