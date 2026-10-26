'use strict';
// electron-builder afterSign hook (package.json build.afterSign): fills in the program's OriginalFilename, which
// electron-builder leaves empty when it writes the version resource. Runs after electron-builder has edited the
// resources and before the installer is made, with the rcedit that electron-builder itself downloaded.
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

function rcedit() {
  const cache = path.join(process.env.LOCALAPPDATA || '', 'electron-builder', 'Cache', 'winCodeSign');
  for (const d of fs.existsSync(cache) ? fs.readdirSync(cache) : []) {
    const exe = path.join(cache, d, 'rcedit-x64.exe');
    if (fs.existsSync(exe)) return exe;
  }
  throw new Error(`rcedit-x64.exe was not found in ${cache}`);
}

exports.default = async function afterSign(context) {
  if (context.electronPlatformName !== 'win32') return;
  const name = `${context.packager.appInfo.productFilename}.exe`;
  execFileSync(rcedit(), [path.join(context.appOutDir, name), '--set-version-string', 'OriginalFilename', name]);
};
