#!/usr/bin/env node
'use strict';
// Doc Studio MCP server: lets AI assistants (Claude Code, Claude Desktop, Gemini CLI, OpenAI Codex) use Doc Studio.
// stdio JSON-RPC, one message per line. No dependencies: it runs with Node 18+ or with Doc Studio's own runtime
// (Doc Studio.exe with ELECTRON_RUN_AS_NODE=1).
//
// It talks to the running Doc Studio through the app's local AI bridge (<APPDATA>\Doc Studio\ai-bridge.json:
// a port on 127.0.0.1 and a secret key), starting Doc Studio when it isn't running. Doc Studio applies the owner's
// rules: the AI writes only new files, may copy a file and work on the copy, may change only files an AI assistant
// created (an earlier session's file needs the owner's OK), and never deletes or sends anything off the PC.

const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const crypto = require('crypto');
const { spawn } = require('child_process');

const VERSION = '1.0.1';
const PROTOCOLS = ['2025-06-18', '2025-03-26', '2024-11-05'];
const SESSION = crypto.randomBytes(8).toString('hex');
const APPDATA = process.env.APPDATA || path.join(os.homedir(), 'AppData', 'Roaming');
const LOCALAPPDATA = process.env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local');
const BRIDGE = path.join(process.env.DOC_STUDIO_DATA || path.join(APPDATA, 'Doc Studio'), 'ai-bridge.json');
// The program to start when Doc Studio is closed: DOC_STUDIO_EXE, else where Doc Studio last ran from
// (install.json, written by Doc Studio itself, so any install folder works), else the standard per-user place.
const installed = (() => { try { return JSON.parse(fs.readFileSync(path.join(path.dirname(BRIDGE), 'install.json'), 'utf8')).exe; } catch { return null; } })();
const EXE = process.env.DOC_STUDIO_EXE || (installed && fs.existsSync(installed) ? installed : path.join(LOCALAPPDATA, 'Programs', 'Doc Studio', 'Doc Studio.exe'));

const INSTRUCTIONS = [
  'Doc Studio is a document app on this PC (Word-compatible documents, spreadsheets, presentations, PDF, Markdown, pictures).',
  'Use these tools to read, create, convert, merge, split, compress, OCR, scan and edit pictures.',
  'Rules: you can only create NEW files; give outputs as full paths, or leave output out and Doc Studio picks a free name next to the source.',
  'To change an existing file the owner did not create with an AI, copy it first (docstudio_copy_document) and change the copy.',
  'You may change files an AI assistant created; a file from an earlier session asks the owner first.',
  'Doc Studio never deletes anything and never sends files off this PC itself. Use docstudio_show to open a result in Doc Studio for the owner to check.',
].join(' ');

const obj = (properties, required) => ({ type: 'object', properties, required: required || [], additionalProperties: false });
const str = (description) => ({ type: 'string', description });
const outPath = str('Optional full path for the new file (it must not exist, unless an AI assistant created it). Leave out to get a free name next to the source.');
const ann = (title, read, idem) => ({ title, readOnlyHint: !!read, destructiveHint: false, idempotentHint: !!idem, openWorldHint: false });

const TOOLS = [
  { name: 'docstudio_status', description: 'Check Doc Studio: version, the documents open in it, and the Documents folder.', inputSchema: obj({}), annotations: ann('Doc Studio status', true, true) },
  { name: 'docstudio_list_open_documents', description: 'List the documents open in Doc Studio (index, name, path, type, whether it has unsaved changes, which one is showing).', inputSchema: obj({}), annotations: ann('List open documents', true, true) },
  { name: 'docstudio_read_document', description: 'Read a document as Markdown text (Word, OpenDocument, RTF, PDF, spreadsheets, presentations, EPUB, HTML, Markdown, text). Give a file path, or open_document to read one that is open in Doc Studio (including its unsaved changes). For pictures and scans use docstudio_ocr.',
    inputSchema: obj({ path: str('Full path of the file to read.'), open_document: { type: ['integer', 'string'], description: 'Instead of path: an open document\'s index or name from docstudio_list_open_documents, or "active" for the one showing.' }, max_chars: { type: 'integer', description: 'Most characters to return (default 200000).' } }),
    annotations: ann('Read a document', true, true) },
  { name: 'docstudio_create_document', description: 'Create a new document from Markdown: .md, .txt, .html, .docx (Word) or .pdf. For other formats create one of these and then use docstudio_convert.',
    inputSchema: obj({ content: str('The document in Markdown (headings, lists, tables, bold, links...).'), output: str('Full path of the new file; its extension sets the format. Leave out to save in the Documents folder named after the title.'), format: { type: 'string', enum: ['md', 'txt', 'html', 'docx', 'pdf'], description: 'Format when no output is given (default md).' }, title: str('Document title (also the file name when no output is given).') }, ['content']),
    annotations: ann('Create a document') },
  { name: 'docstudio_write_document', description: 'Rewrite a document an AI assistant created (.md, .txt, .html, .docx or .pdf) with new Markdown content. Files from an earlier session ask the owner first. To change any other file, copy it with docstudio_copy_document and write the copy.',
    inputSchema: obj({ path: str('Full path of the AI-created file.'), content: str('The whole new document in Markdown.'), title: str('Optional title.') }, ['path', 'content']),
    annotations: ann('Rewrite an AI-created document') },
  { name: 'docstudio_find_replace', description: 'Find and replace text in a text file an AI assistant created (.md, .txt, .html, .csv, .json, .xml).',
    inputSchema: obj({ path: str('Full path of the AI-created text file.'), find: str('Text to find.'), replace: str('Replacement text.'), all: { type: 'boolean', description: 'Replace every match (default true).' }, match_case: { type: 'boolean', description: 'Match upper/lower case exactly (default false).' } }, ['path', 'find', 'replace']),
    annotations: ann('Find and replace') },
  { name: 'docstudio_copy_document', description: 'Copy any file to a new file, so the copy can be changed (the original is never touched).',
    inputSchema: obj({ source: str('Full path of the file to copy.'), output: str('Optional full path of the copy. Leave out for "name (copy).ext" next to the original.') }, ['source']),
    annotations: ann('Copy a document') },
  { name: 'docstudio_convert', description: 'Convert a file to another format and save it as a new file (for example docx to pdf, pdf to docx or md, md to docx, xlsx to csv, pptx to pdf or png, png to jpg).',
    inputSchema: obj({ source: str('Full path of the file to convert.'), to: str('Target format, for example "pdf", "docx", "md", "html", "txt", "odt", "rtf", "epub", "xlsx", "csv", "png", "jpg", "webp", "bmp".'), output: outPath }, ['source', 'to']),
    annotations: ann('Convert a file') },
  { name: 'docstudio_merge_pdfs', description: 'Join PDFs, pictures and documents (in order) into one new PDF.',
    inputSchema: obj({ sources: { type: 'array', items: { type: 'string' }, description: 'Full paths, in order.' }, output: outPath }, ['sources']),
    annotations: ann('Merge into one PDF') },
  { name: 'docstudio_split_pdf', description: 'Split a PDF into new PDFs: every page, every N pages, or by page ranges.',
    inputSchema: obj({ source: str('Full path of the PDF.'), ranges: str('Optional page ranges, one file each, for example "1-3, 5, 8-end".'), every: { type: 'integer', description: 'Optional: pages per file.' }, output_folder: str('Optional folder for the new files (default: next to the PDF).') }, ['source']),
    annotations: ann('Split a PDF') },
  { name: 'docstudio_compress', description: 'Make a PDF, picture or Office file smaller, saved as a new file.',
    inputSchema: obj({ source: str('Full path of the file.'), level: { type: 'string', enum: ['lossless', 'balanced', 'small'], description: 'PDFs: lossless (tidy up only), balanced (150 dpi pages) or small (96 dpi pages).' }, quality: { type: 'integer', description: 'Pictures and Office files: picture quality 10-100 (default 75).' }, max_side: { type: 'integer', description: 'Pictures and Office files: largest picture side in pixels.' }, format: { type: 'string', enum: ['jpg', 'webp', 'png'], description: 'Pictures: output type.' }, output: outPath }, ['source']),
    annotations: ann('Compress a file') },
  { name: 'docstudio_ocr', description: 'Read the text in a scan, photo or PDF with the text recognition built into Windows. Returns the text (pages separated by a form feed).',
    inputSchema: obj({ source: str('Full path of the picture or PDF.') }, ['source']), annotations: ann('OCR (scan to text)', true, true) },
  { name: 'docstudio_edit_image', description: 'Edit a picture and save the result as a new file. Steps run in order: rotate {turns: 1|2|3 or degrees: 90|180|270}, flip {direction: "horizontal"|"vertical"}, straighten {degrees: -45..45}, crop {x, y, width, height in pixels, or 0-1 fractions}, resize {width and/or height in pixels}, remove_background {fill: "transparent"|"white"|"black"|"#rrggbb"} (AI cut-out, runs on this PC), adjust {brightness, contrast, saturation, warmth, tint, highlights, shadows: -100..100; sharpness, vignette: 0..100}, filter {name: vivid|warm|cool|mono|noir|sepia|fade|dramatic|vintage, amount: 0-100}.',
    inputSchema: obj({ source: str('Full path of the picture.'), steps: { type: 'array', items: { type: 'object' }, description: 'Edit steps, each with "op" and its settings.' }, format: { type: 'string', enum: ['png', 'jpg', 'webp', 'bmp'], description: 'Output type (default: the output extension, else the source type). Use png to keep a transparent background.' }, output: outPath }, ['source', 'steps']),
    annotations: ann('Edit a picture') },
  { name: 'docstudio_scan', description: 'Scan one page from the scanner connected to this PC (the owner must put the page on the glass first) and save it as a PDF or PNG, or return its text.',
    inputSchema: obj({ make: { type: 'string', enum: ['pdf', 'png', 'text'], description: 'What to make (default pdf).' }, color: { type: 'string', enum: ['color', 'gray', 'bw'] }, dpi: { type: 'integer', enum: [150, 200, 300, 600] }, paper: { type: 'string', enum: ['letter', 'a4', 'glass'] }, scanner: str('Optional part of the scanner name.'), output: str('Optional full path (default: Documents\\Scan <date>.pdf).') }),
    annotations: ann('Scan a page') },
  { name: 'docstudio_show', description: 'Open a file in Doc Studio so the owner can see it (for example a document you just created).',
    inputSchema: obj({ path: str('Full path of the file.') }, ['path']), annotations: ann('Show in Doc Studio', false, true) },
];

// ---------------------------------------------------------------- the Doc Studio bridge

const readBridge = () => { try { return JSON.parse(fs.readFileSync(BRIDGE, 'utf8')); } catch { return null; } };
function post(bridge, body, timeout) {
  return new Promise((resolve, reject) => {
    const data = Buffer.from(JSON.stringify(body));
    const req = http.request({ host: '127.0.0.1', port: bridge.port, path: '/call', method: 'POST', headers: { 'Content-Type': 'application/json', 'Content-Length': data.length, Authorization: `Bearer ${bridge.token}` } }, (res) => {
      const parts = [];
      res.on('data', (c) => parts.push(c));
      res.on('end', () => { try { resolve(JSON.parse(Buffer.concat(parts).toString('utf8'))); } catch (err) { reject(err); } });
    });
    req.on('error', reject);
    req.setTimeout(timeout || 15 * 60 * 1000, () => req.destroy(new Error('Doc Studio took too long to answer.')));
    req.end(data);
  });
}
// The running Doc Studio's bridge, starting Doc Studio when needed.
async function bridge() {
  let b = readBridge();
  if (b && (await post(b, { tool: 'status', session: SESSION }, 5000).then(() => true, () => false))) return b;
  // AI assistants are off until the owner turns them on in Doc Studio; then there is nothing to connect to.
  const settings = (() => { try { return JSON.parse(fs.readFileSync(path.join(path.dirname(BRIDGE), 'settings.json'), 'utf8')); } catch { return {}; } })();
  if (settings.aiEnabled !== true) throw new Error('AI assistants are turned off in Doc Studio. Ask the owner to turn them on in Doc Studio (Settings > AI assistants), then try again.');
  if (!fs.existsSync(EXE)) throw new Error(`Doc Studio is not installed at ${EXE}. Install it, or set DOC_STUDIO_EXE to Doc Studio.exe.`);
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE; // this server may itself run on Doc Studio's runtime
  spawn(EXE, [], { detached: true, stdio: 'ignore', env, windowsHide: false }).unref();
  for (let i = 0; i < 60; i++) {
    await new Promise((r) => setTimeout(r, 500));
    b = readBridge();
    if (b && (await post(b, { tool: 'status', session: SESSION }, 5000).then(() => true, () => false))) return b;
  }
  throw new Error('Doc Studio did not start. Open it, then try again.');
}

async function callTool(name, args) {
  const tool = TOOLS.find((t) => t.name === name);
  if (!tool) return { isError: true, content: [{ type: 'text', text: `Unknown tool ${name}` }] };
  try {
    const r = await post(await bridge(), { tool: name.replace(/^docstudio_/, ''), args: args || {}, session: SESSION });
    if (!r.ok) return { isError: true, content: [{ type: 'text', text: `${r.error}${r.code ? ` (${r.code})` : ''}` }] };
    const res = r.result;
    // Text results (reading, OCR) come back as plain text; the rest as JSON.
    if (res && typeof res.text === 'string') {
      const head = { ...res }; delete head.text;
      return { content: [{ type: 'text', text: res.text }, ...(Object.keys(head).length ? [{ type: 'text', text: JSON.stringify(head) }] : [])], structuredContent: res };
    }
    return { content: [{ type: 'text', text: JSON.stringify(res, null, 2) }], structuredContent: res };
  } catch (err) {
    return { isError: true, content: [{ type: 'text', text: err.message }] };
  }
}

// ---------------------------------------------------------------- JSON-RPC over stdio

function reply(id, result, error) {
  process.stdout.write(JSON.stringify(error ? { jsonrpc: '2.0', id, error } : { jsonrpc: '2.0', id, result }) + '\n');
}
async function onMessage(m) {
  if (!m || m.jsonrpc !== '2.0') return;
  const { id, method, params } = m;
  if (id === undefined) return; // notifications (initialized, cancelled...)
  try {
    if (method === 'initialize') {
      const want = params && params.protocolVersion;
      reply(id, { protocolVersion: PROTOCOLS.includes(want) ? want : PROTOCOLS[0], capabilities: { tools: { listChanged: false } }, serverInfo: { name: 'doc-studio', title: 'Doc Studio', version: VERSION }, instructions: INSTRUCTIONS });
    } else if (method === 'ping') reply(id, {});
    else if (method === 'tools/list') reply(id, { tools: TOOLS });
    else if (method === 'tools/call') reply(id, await callTool(params && params.name, params && params.arguments));
    else reply(id, null, { code: -32601, message: `Method not found: ${method}` });
  } catch (err) { reply(id, null, { code: -32603, message: err.message }); }
}
function main() {
  let buf = '';
  process.stdin.setEncoding('utf8');
  process.stdin.on('data', (chunk) => {
    buf += chunk;
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).trim();
      buf = buf.slice(i + 1);
      if (!line) continue;
      let m;
      try { m = JSON.parse(line); } catch { reply(null, null, { code: -32700, message: 'Parse error' }); continue; }
      onMessage(m);
    }
  });
  process.stdin.on('end', () => process.exit(0));
}

module.exports = { TOOLS, INSTRUCTIONS, VERSION, EXE };
if (require.main === module) main();
