#!/usr/bin/env node
'use strict';
// Builds build/ai: Doc Studio's AI assistant packages, shipped inside the installer (resources\ai) so Settings >
// AI assistants can set them up. One MCP server (integrations/mcp/mcp-server.js) and one guide (GUIDE.md) feed:
//   .claude-plugin/marketplace.json + claude-code/   Claude Code plugin (a local marketplace)
//   doc-studio.mcpb                                  Claude Desktop extension (runs with Claude's own Node)
//   gemini/                                          Gemini CLI extension
//   mcp-server.js + codex/README.md                  OpenAI Codex (and any other MCP client)
//   node tools/make-ai-packages.js
// No dependencies: the .mcpb zip is written here (deflate from zlib, CRC-32 below).

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const root = path.join(__dirname, '..');
const src = path.join(root, 'integrations');
const out = path.join(root, 'build', 'ai');
const pkg = require('../package.json');
const { TOOLS } = require('../integrations/mcp/mcp-server');
const server = fs.readFileSync(path.join(src, 'mcp', 'mcp-server.js'));
const guide = fs.readFileSync(path.join(src, 'GUIDE.md'), 'utf8');
const version = pkg.version;

fs.rmSync(out, { recursive: true, force: true });
const put = (rel, data) => { const f = path.join(out, rel); fs.mkdirSync(path.dirname(f), { recursive: true }); fs.writeFileSync(f, data); };
const json = (o) => JSON.stringify(o, null, 2) + '\n';
const withVersion = (file) => json({ ...JSON.parse(fs.readFileSync(file, 'utf8')), version });

// Claude Code: a marketplace with one plugin.
put('.claude-plugin/marketplace.json', json({
  name: 'doc-studio',
  owner: { name: 'Griffith Software' },
  description: 'The Doc Studio plugin, installed from Doc Studio\'s own folder.',
  plugins: [{ name: 'doc-studio', source: './claude-code', description: pkg.description, version }],
}));
put('claude-code/.claude-plugin/plugin.json', withVersion(path.join(src, 'claude-code', '.claude-plugin', 'plugin.json')));
put('claude-code/.mcp.json', fs.readFileSync(path.join(src, 'claude-code', '.mcp.json')));
put('claude-code/server/mcp-server.js', server);
put('claude-code/skills/doc-studio/SKILL.md', `---\nname: doc-studio\ndescription: Use Doc Studio (the owner's document app) to read, create, convert, merge, split, compress, OCR, scan and edit pictures. Use when the owner asks to make or change a Word/PDF/Markdown document, convert a file, combine or split PDFs, make a file smaller, read text from a scan or photo, scan a page, or edit a picture (crop, resize, remove background).\n---\n\n${guide}`);

// Gemini CLI extension.
put('gemini/gemini-extension.json', withVersion(path.join(src, 'gemini', 'gemini-extension.json')));
put('gemini/GEMINI.md', guide);
put('gemini/server/mcp-server.js', server);

// Codex and other MCP clients: the server on its own (run with Doc Studio.exe and ELECTRON_RUN_AS_NODE=1).
put('mcp-server.js', server);
put('codex/README.md', `# Doc Studio for OpenAI Codex\n\nIn Doc Studio, open Settings > AI assistants > Codex and press Copy the setup. It copies a block like this,\nwith the paths of this PC, to add to %USERPROFILE%\\.codex\\config.toml:\n\n\`\`\`toml\n[mcp_servers.doc_studio]\ncommand = 'C:\\Users\\<you>\\AppData\\Local\\Programs\\Doc Studio\\Doc Studio.exe'\nargs = ['C:\\Users\\<you>\\AppData\\Local\\Programs\\Doc Studio\\resources\\ai\\mcp-server.js']\nenv = { ELECTRON_RUN_AS_NODE = '1' }\n\`\`\`\n\n${guide}`);

// Claude Desktop extension (.mcpb).
const manifest = {
  manifest_version: '0.2',
  name: 'doc-studio',
  display_name: 'Doc Studio',
  version,
  description: 'Read, create, convert, merge, split, compress, OCR, scan and edit pictures with Doc Studio.',
  long_description: 'Doc Studio is a document app for Windows. With this extension Claude can read any document as Markdown, create Word, PDF, HTML and Markdown documents, convert between formats, merge, split and compress PDFs, read text from scans and photos, scan a page, and edit pictures (crop, rotate, resize, remove background, filters). It writes new files only: to change one of your files it copies it first, it may change files an AI made (asking you for files from an earlier session), and it never deletes or uploads anything. Doc Studio must be installed.',
  author: { name: 'Griffith Software' },
  icon: 'icon.png',
  server: { type: 'node', entry_point: 'server/mcp-server.js', mcp_config: { command: 'node', args: ['${__dirname}/server/mcp-server.js'] } },
  tools: TOOLS.map((t) => ({ name: t.name, description: t.description.split('. ')[0] + '.' })),
  keywords: ['documents', 'pdf', 'word', 'markdown', 'convert', 'ocr', 'scan', 'images'],
  compatibility: { platforms: ['win32'], runtimes: { node: '>=18.0.0' } },
};
put('doc-studio.mcpb', zip([
  ['manifest.json', Buffer.from(json(manifest))],
  ['server/mcp-server.js', server],
  ['icon.png', fs.readFileSync(path.join(root, 'build', 'icon.png'))],
]));

put('README.md', `Doc Studio ${version} - AI assistant packages (see Settings > AI assistants in Doc Studio).\n`);
console.log(`AI packages ${version}: ${TOOLS.length} tools -> build/ai`);

// ---------------------------------------------------------------- a small zip writer

function crc32(buf) {
  let c = ~0;
  for (const b of buf) { c ^= b; for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1)); }
  return ~c >>> 0;
}
function zip(entries) {
  const locals = [], centrals = [];
  let offset = 0;
  for (const [name, data] of entries) {
    const n = Buffer.from(name), comp = zlib.deflateRawSync(data), crc = crc32(data);
    const local = Buffer.alloc(30); local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6); local.writeUInt16LE(8, 8);
    local.writeUInt32LE(crc, 14); local.writeUInt32LE(comp.length, 18); local.writeUInt32LE(data.length, 22); local.writeUInt16LE(n.length, 26);
    const central = Buffer.alloc(46); central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(0x0800, 8); central.writeUInt16LE(8, 10);
    central.writeUInt32LE(crc, 16); central.writeUInt32LE(comp.length, 20); central.writeUInt32LE(data.length, 24); central.writeUInt16LE(n.length, 28); central.writeUInt32LE(offset, 42);
    locals.push(local, n, comp); centrals.push(central, n);
    offset += 30 + n.length + comp.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10); end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}

// Every JSON file in the packages must parse (a stray backslash once broke the Claude Code plugin's .mcp.json).
for (const f of ['.claude-plugin/marketplace.json', 'claude-code/.claude-plugin/plugin.json', 'claude-code/.mcp.json', 'gemini/gemini-extension.json']) JSON.parse(fs.readFileSync(path.join(out, f), 'utf8'));
