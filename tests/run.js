#!/usr/bin/env node
'use strict';
// `npm test`: the syntax check, then the suites that need no sample files (each in the real app, with its own
// temporary profile). The other suites take sample documents; see README.md (Tests).
//   npm test
const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const root = path.join(__dirname, '..');
const electron = require('electron'); // the path to electron.exe when required from Node
const suites = ['image-test', 'toolbar-test', 'xlsx-ui-test'];
let failed = 0;

const syntax = spawnSync(process.execPath, [path.join(__dirname, 'check-syntax.js')], { cwd: root, stdio: 'inherit' });
if (syntax.status !== 0) failed++;

for (const name of suites) {
  const out = fs.mkdtempSync(path.join(os.tmpdir(), `docstudio-${name}-`));
  const r = spawnSync(electron, [path.join(__dirname, `${name}.js`)], { cwd: root, env: { ...process.env, DOCSTUDIO_OUT: out }, encoding: 'utf8', timeout: 10 * 60 * 1000 });
  const ok = r.status === 0;
  if (!ok) failed++;
  // Each suite prints its checks; show the failing ones and the summary line.
  const lines = `${r.stdout || ''}`.split(/\r?\n/).filter((l) => /^FAIL|passed|failed/i.test(l));
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${lines.length ? `  (${lines[lines.length - 1].trim()})` : ''}`);
  if (!ok) lines.filter((l) => /^FAIL/.test(l)).slice(0, 10).forEach((l) => console.log(`      ${l}`));
  fs.rmSync(out, { recursive: true, force: true });
}

console.log(failed ? `\n${failed} failed` : '\nAll passed');
process.exit(failed ? 1 : 0);
