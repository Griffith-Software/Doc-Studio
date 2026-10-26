'use strict';
// Runs before every build: every script must parse, or the build stops (a broken script would leave part of the
// window missing). node tests/check-syntax.js
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const files = ['main.js', 'preload.js', 'convert.js', 'phone.js', 'ai/bridge.js', 'shell-icons.js', ...fs.readdirSync(path.join(root, 'renderer')).filter((f) => f.endsWith('.js')).map((f) => `renderer/${f}`)];
let bad = 0;
for (const f of files) {
  try { execFileSync(process.execPath, ['--check', path.join(root, f)], { stdio: 'pipe' }); } catch (err) {
    bad++;
    console.error(`SYNTAX ERROR in ${f}\n${String(err.stderr || err.message).split('\n').slice(0, 5).join('\n')}`);
  }
}
if (bad) { console.error(`${bad} file(s) with syntax errors - not building.`); process.exit(1); }
console.log(`Syntax OK (${files.length} scripts)`);
