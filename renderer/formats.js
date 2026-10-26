'use strict';
// Which handler opens which file extension, and helpers shared by viewers and editors.

const FORMATS = (() => {
  const kinds = {
    docx: ['docx', 'docm', 'dotx'],
    md: ['md', 'markdown', 'mdown', 'mkd', 'mdx'],
    html: ['html', 'htm', 'xhtml'],
    csv: ['csv', 'tsv'],
    text: ['txt', 'text', 'log', 'json', 'jsonc', 'xml', 'yaml', 'yml', 'toml', 'ini', 'cfg', 'conf', 'env', 'properties',
      'js', 'mjs', 'cjs', 'ts', 'tsx', 'jsx', 'css', 'scss', 'less', 'py', 'rb', 'php', 'java', 'kt', 'cs', 'c', 'h', 'cpp', 'hpp',
      'go', 'rs', 'swift', 'sh', 'bash', 'ps1', 'psm1', 'bat', 'cmd', 'sql', 'r', 'lua', 'pl', 'vb', 'vbs', 'tex', 'bib', 'srt',
      'vtt', 'eml', 'nfo', 'diff', 'patch', 'gitignore', 'editorconfig', 'reg', 'svgz', 'rst', 'adoc', 'org'],
    pdf: ['pdf'],
    sheet: ['xlsx', 'xlsm', 'xltx', 'xltm'],
    sheetview: ['xlsb', 'xls', 'ods', 'fods', 'numbers'],
    pptx: ['pptx', 'ppsx', 'pptm'],
    ppt: ['ppt', 'pps', 'pot'],
    rtf: ['rtf'],
    odt: ['odt', 'ott'],
    odp: ['odp'],
    doc: ['doc', 'dot'],
    epub: ['epub'],
    image: ['png', 'jpg', 'jpeg', 'jfif', 'gif', 'webp', 'bmp', 'svg', 'ico', 'avif', 'apng'],
    audio: ['mp3', 'wav', 'ogg', 'oga', 'm4a', 'aac', 'flac', 'opus'],
    video: ['mp4', 'webm', 'ogv', 'mov', 'm4v'],
  };
  const byExt = {};
  for (const [kind, exts] of Object.entries(kinds)) exts.forEach((e) => { byExt[e] = kind; });

  const labels = {
    docx: 'Document', md: 'Markdown', html: 'HTML', csv: 'Delimited text', text: 'Text', pdf: 'PDF',
    sheet: 'Spreadsheet', sheetview: 'Spreadsheet', pptx: 'Presentation', rtf: 'Rich Text', odt: 'OpenDocument text', odp: 'OpenDocument presentation',
    doc: 'Document (97-2003)', ppt: 'Presentation (97-2003)', epub: 'EPUB book', image: 'Image', audio: 'Audio', video: 'Video', unknown: 'File',
  };
  const editable = new Set(['docx', 'sheet', 'pdf', 'pptx', 'odt', 'rtf', 'md', 'html', 'csv', 'text', 'image']);
  const mime = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jfif: 'image/jpeg', gif: 'image/gif', webp: 'image/webp',
    bmp: 'image/bmp', svg: 'image/svg+xml', ico: 'image/x-icon', avif: 'image/avif', apng: 'image/apng',
    mp3: 'audio/mpeg', wav: 'audio/wav', ogg: 'audio/ogg', oga: 'audio/ogg', m4a: 'audio/mp4', aac: 'audio/aac', flac: 'audio/flac', opus: 'audio/ogg',
    mp4: 'video/mp4', webm: 'video/webm', ogv: 'video/ogg', mov: 'video/quicktime', m4v: 'video/mp4',
    pdf: 'application/pdf',
  };

  function extOf(name) {
    const base = name.split(/[\\/]/).pop();
    const i = base.lastIndexOf('.');
    return i > 0 ? base.slice(i + 1).toLowerCase() : (base.startsWith('.') ? base.slice(1).toLowerCase() : '');
  }

  return {
    extOf,
    kindOf: (name) => byExt[extOf(name)] || 'unknown',
    label: (kind) => labels[kind] || 'File',
    isEditable: (kind) => editable.has(kind),
    mimeOf: (ext) => mime[ext] || 'application/octet-stream',
  };
})();

// ---------- text decoding that remembers how to write the file back

const TextCodec = {
  // Returns { text, encoding, bom, eol } so saving preserves the original conventions.
  decode(buf) {
    const bytes = new Uint8Array(buf);
    let encoding = 'utf-8';
    let bom = false;
    let start = 0;
    if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) { bom = true; start = 3; }
    else if (bytes[0] === 0xff && bytes[1] === 0xfe) { encoding = 'utf-16le'; bom = true; start = 2; }
    else if (bytes[0] === 0xfe && bytes[1] === 0xff) { encoding = 'utf-16be'; bom = true; start = 2; }
    let text;
    try {
      text = new TextDecoder(encoding, { fatal: true }).decode(bytes.subarray(start));
    } catch {
      encoding = 'windows-1252';
      text = new TextDecoder(encoding).decode(bytes.subarray(start));
    }
    const crlf = (text.match(/\r\n/g) || []).length;
    const lf = (text.match(/(^|[^\r])\n/g) || []).length;
    const eol = crlf > lf ? '\r\n' : '\n';
    return { text: text.replace(/\r\n/g, '\n'), encoding, bom, eol };
  },

  encode(text, meta) {
    const m = meta || { encoding: 'utf-8', bom: false, eol: '\r\n' };
    const out = m.eol === '\r\n' ? text.replace(/\r?\n/g, '\r\n') : text;
    if (m.encoding === 'utf-16le' || m.encoding === 'utf-16be') {
      const le = m.encoding === 'utf-16le';
      const arr = new Uint8Array(2 + out.length * 2);
      arr[0] = le ? 0xff : 0xfe; arr[1] = le ? 0xfe : 0xff;
      for (let i = 0; i < out.length; i++) {
        const c = out.charCodeAt(i);
        arr[2 + i * 2] = le ? c & 0xff : c >> 8;
        arr[3 + i * 2] = le ? c >> 8 : c & 0xff;
      }
      return arr.buffer;
    }
    if (m.encoding === 'windows-1252') {
      // Characters outside Windows-1252 can't be written in that encoding; upgrade to UTF-8 with BOM instead.
      const ok = [...out].every((ch) => ch.charCodeAt(0) < 0x80 || CP1252.has(ch));
      if (ok) return Uint8Array.from([...out].map((ch) => (ch.charCodeAt(0) < 0x80 ? ch.charCodeAt(0) : CP1252.get(ch)))).buffer;
      m.encoding = 'utf-8'; m.bom = true;
    }
    const body = new TextEncoder().encode(out);
    if (!m.bom) return body.buffer;
    const arr = new Uint8Array(body.length + 3);
    arr.set([0xef, 0xbb, 0xbf]); arr.set(body, 3);
    return arr.buffer;
  },

  looksBinary(buf) {
    const bytes = new Uint8Array(buf, 0, Math.min(buf.byteLength, 8192));
    if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff)) return false;
    let odd = 0;
    for (const b of bytes) { if (b === 0) return true; if (b < 9 || (b > 13 && b < 32)) odd++; }
    return odd > bytes.length * 0.1;
  },
};

// Byte -> char map for Windows-1252 high range, inverted for encoding.
const CP1252 = (() => {
  const dec = new TextDecoder('windows-1252');
  const map = new Map();
  for (let b = 0x80; b <= 0xff; b++) map.set(dec.decode(Uint8Array.of(b)), b);
  return map;
})();

// ---------- small DOM helpers

function h(tag, attrs, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'text') el.textContent = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat()) if (c != null && c !== false) el.append(c.nodeType ? c : String(c));
  return el;
}

function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1048576) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1048576).toFixed(1)} MB`;
}

function countWords(text) {
  const m = text.trim().match(/\S+/g);
  return m ? m.length : 0;
}

function sanitize(html, opts) {
  return DOMPurify.sanitize(html, { ADD_ATTR: ['target'], FORBID_TAGS: ['style', 'form', 'input', 'button'], ...opts });
}
