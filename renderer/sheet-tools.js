'use strict';
// Google Sheets-style features for the workbook editor (added to SheetView.View):
//   Insert: chart, picture, checkbox, dropdown, note, link, function, pivot table
//   Format: vertical align, text rotation, custom number formats, alternating colours, convert to table, paint format
//   Data:   sort range (advanced), remove duplicates, trim whitespace, split text to columns, randomize range,
//           column stats, named ranges, protect sheet, data validation
//   View:   gridlines, show formulas, zoom
//   Sheet tabs: duplicate, delete, hide/unhide, colour, move
// Changes that add package parts (charts, pictures, notes, tables) edit the saved package and reload it, so they
// undo as one step like the other structural edits.

(() => {
  const P = SheetView.View.prototype;
  const { key, addr, colName, parseAddr, parseRange, rangeRef, quoteSheet } = SheetModel;
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const REL = (t) => `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${t}`;
  const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const parseXml = (t) => new DOMParser().parseFromString(t, 'application/xml');
  const serialize = (d) => { const s = new XMLSerializer().serializeToString(d); return s.startsWith('<?xml') ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s; };
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const kidsOf = (el, name) => [...el.childNodes].filter((c) => c.nodeType === 1 && (!name || c.localName === name));

  // An element in the worksheet's namespace (keeping the document's prefix, if it uses one).
  function wsEl(doc, name, attrs) {
    const p = doc.documentElement.prefix;
    const e = doc.createElementNS(NS, p ? `${p}:${name}` : name);
    for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v);
    return e;
  }
  function insertOrdered(root, el) {
    const old = kidsOf(root, el.localName)[0];
    if (old) { root.replaceChild(el, old); return el; }
    const rank = WS_ORDER.indexOf(el.localName);
    const after = kidsOf(root).find((c) => WS_ORDER.indexOf(c.localName) > rank);
    root.insertBefore(el, after || null);
    return el;
  }
  const rangeOf = (view) => view.effectiveRange();
  const textOf = (view, r, c) => SheetView.display(view.sheet.get(r, c), view.wb.styles, view.wb.date1904).text;
  const valueOf = (view, r, c) => { const x = view.sheet.get(r, c); return x ? x.value : null; };

  // ======================================================================
  // Package edits: change the saved package, then reload (one undo step)
  // ======================================================================

  P.packageOp = async function (label, fn) {
    if (this.editing) this.endEdit(true);
    const before = await this.wb.save();
    const zip = await JSZip.loadAsync(before);
    const docs = new Map();
    const pkg = {
      zip,
      doc: async (path) => { if (!docs.has(path)) docs.set(path, zip.file(path) ? parseXml(await zip.file(path).async('string')) : null); return docs.get(path); },
      put: (path, docOrText) => { if (typeof docOrText === 'string') { docs.delete(path); zip.file(path, docOrText); } else docs.set(path, docOrText); },
      exists: (path) => !!zip.file(path) || (docs.has(path) && docs.get(path) !== null),
      relsPath: (part) => part.replace(/([^/]+)$/, '_rels/$1.rels'),
      rels: async (part) => {
        const p = pkg.relsPath(part);
        let d = await pkg.doc(p);
        if (!d) { d = parseXml(`<Relationships xmlns="${PKG}"/>`); docs.set(p, d); }
        return d;
      },
      addRel: async (part, type, target) => {
        const d = await pkg.rels(part);
        const used = new Set([...d.documentElement.children].map((x) => x.getAttribute('Id')));
        let n = 1; while (used.has(`rId${n}`)) n++;
        const el = d.createElementNS(PKG, 'Relationship');
        el.setAttribute('Id', `rId${n}`); el.setAttribute('Type', type); el.setAttribute('Target', target);
        d.documentElement.append(el);
        return `rId${n}`;
      },
      relTarget: async (part, type) => {
        const d = await pkg.rels(part);
        const r = [...d.documentElement.children].find((x) => x.getAttribute('Type') === type);
        return r ? resolve(part, r.getAttribute('Target')) : null;
      },
      override: async (path, type) => {
        const ct = await pkg.doc('[Content_Types].xml');
        const root = ct.documentElement;
        if ([...root.children].some((x) => x.getAttribute('PartName') === `/${path}`)) return;
        const o = ct.createElementNS(CT_NS, 'Override');
        o.setAttribute('PartName', `/${path}`); o.setAttribute('ContentType', type);
        root.append(o);
      },
      default: async (ext, type) => {
        const ct = await pkg.doc('[Content_Types].xml');
        const root = ct.documentElement;
        if ([...root.children].some((x) => x.localName === 'Default' && (x.getAttribute('Extension') || '').toLowerCase() === ext)) return;
        const o = ct.createElementNS(CT_NS, 'Default');
        o.setAttribute('Extension', ext); o.setAttribute('ContentType', type);
        root.insertBefore(o, root.firstChild);
      },
      free: (pattern) => { let n = 1; while (pkg.exists(pattern.replace('#', n))) n++; return pattern.replace('#', n); },
    };
    await fn(pkg);
    for (const [p, d] of docs) if (d) zip.file(p, serialize(d));
    const bytes = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    await this.reloadFrom(bytes, this.sheetIndex);
    const after = await this.wb.save();
    this.pushUndo({ snapshot: { before, after }, sheetIndex: this.sheetIndex, label });
    this.changed();
  };
  function resolve(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const p of target.split('/')) { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); }
    return parts.join('/');
  }
  function relative(fromPart, toPath) {
    const a = fromPart.split('/'); a.pop();
    const b = toPath.split('/');
    while (a.length && b.length && a[0] === b[0]) { a.shift(); b.shift(); }
    return '../'.repeat(a.length) + b.join('/');
  }

  // The sheet's drawing part (created when missing), for charts and pictures.
  async function drawingOf(pkg, sheetPath) {
    let path = await pkg.relTarget(sheetPath, REL('drawing'));
    if (path) return { path, doc: await pkg.doc(path) };
    path = pkg.free('xl/drawings/drawing#.xml');
    const doc = parseXml(`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`);
    pkg.put(path, doc);
    const rid = await pkg.addRel(sheetPath, REL('drawing'), relative(sheetPath, path));
    await pkg.override(path, 'application/vnd.openxmlformats-officedocument.drawing+xml');
    const sheet = await pkg.doc(sheetPath);
    const d = wsEl(sheet, 'drawing');
    d.setAttributeNS(RNS, 'r:id', rid);
    insertOrdered(sheet.documentElement, d);
    return { path, doc };
  }
  const anchorXml = (c1, r1, c2, r2, inner) => `<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" editAs="oneCell"><xdr:from><xdr:col>${c1 - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r1 - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${c2 - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r2 - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>${inner}<xdr:clientData/></xdr:twoCellAnchor>`;
  function nextShapeId(doc) { return 1 + Math.max(1, ...[...doc.getElementsByTagName('*')].filter((e) => e.localName === 'cNvPr').map((e) => +e.getAttribute('id') || 0)); }

  // ======================================================================
  // Insert > Chart
  // ======================================================================

  // spec: { type: 'column'|'bar'|'line'|'area'|'pie'|'scatter', title }
  P.insertChart = async function (spec) {
    const g = rangeOf(this);
    const s = this.sheet;
    const r2 = Math.min(g.r2, Math.max(s.maxRow, g.r1)), c2 = Math.min(g.c2, Math.max(s.maxCol, g.c1));
    if (r2 - g.r1 < 1 && c2 - g.c1 < 1) throw new Error('Select the data for the chart first (with a header row).');
    const isText = (r, c) => { const v = valueOf(this, r, c); return v !== null && v !== '' && typeof v === 'string'; };
    const headerRow = [...Array(c2 - g.c1 + 1)].some((_, i) => isText(g.r1, g.c1 + i));
    const catCol = [...Array(r2 - g.r1)].some((_, i) => isText(g.r1 + 1 + i, g.c1)) || (spec.type !== 'scatter' && c2 > g.c1 && headerRow && !isText(g.r1, g.c1));
    const d1 = headerRow ? g.r1 + 1 : g.r1;
    const firstVal = catCol ? g.c1 + 1 : g.c1;
    if (firstVal > c2) throw new Error('The selection needs at least one column of numbers.');
    const sh = quoteSheet(s.name);
    const ref = (r1, c1, rr2, cc2) => `${sh}!$${colName(c1)}$${r1}:$${colName(cc2)}$${rr2}`;
    const cell = (r, c) => `${sh}!$${colName(c)}$${r}`;
    const strCache = (vals) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
    const numCache = (vals) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numCache>`;
    const rows = [];
    for (let r = d1; r <= r2; r++) rows.push(r);
    const cats = catCol ? rows.map((r) => textOf(this, r, g.c1)) : rows.map((_, i) => String(i + 1));
    const series = [];
    for (let c = firstVal; c <= c2; c++) {
      const name = headerRow ? textOf(this, g.r1, c) : `Series ${c - firstVal + 1}`;
      const vals = rows.map((r) => { const v = valueOf(this, r, c); return typeof v === 'number' ? v : null; });
      series.push({ c, name, vals });
    }
    const type = spec.type || 'column';
    const ser = series.map((x, i) => {
      const tx = headerRow ? `<c:tx><c:strRef><c:f>${esc(cell(g.r1, x.c))}</c:f>${strCache([x.name])}</c:strRef></c:tx>` : '';
      const cat = catCol ? `<c:strRef><c:f>${esc(ref(d1, g.c1, r2, g.c1))}</c:f>${strCache(cats)}</c:strRef>` : `<c:numRef><c:f>${esc(ref(d1, x.c, r2, x.c))}</c:f>${numCache(rows.map((_, k) => k + 1))}</c:numRef>`;
      const val = `<c:numRef><c:f>${esc(ref(d1, x.c, r2, x.c))}</c:f>${numCache(x.vals)}</c:numRef>`;
      if (type === 'scatter') {
        const xs = catCol ? `<c:xVal><c:numRef><c:f>${esc(ref(d1, g.c1, r2, g.c1))}</c:f>${numCache(rows.map((r) => valueOf(this, r, g.c1)))}</c:numRef></c:xVal>` : '';
        return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>${xs}<c:yVal>${val}</c:yVal><c:smooth val="0"/></c:ser>`;
      }
      const marker = type === 'line' ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : '';
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${type === 'bar' || type === 'column' ? '<c:invertIfNegative val="0"/>' : ''}${marker}<c:cat>${cat}</c:cat><c:val>${val}</c:val>${type === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    }).join('');
    const axes = '<c:axId val="500000001"/><c:axId val="500000002"/>';
    let plot;
    if (type === 'pie') plot = `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
    else if (type === 'line') plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/>${axes}</c:lineChart>`;
    else if (type === 'area') plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}${axes}</c:areaChart>`;
    else if (type === 'scatter') plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${ser}${axes}</c:scatterChart>`;
    else plot = `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="150"/>${axes}</c:barChart>`;
    const catAx = type === 'scatter' ? `<c:valAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
      : `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${type === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${type === 'bar' ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="${type === 'area' ? 'midCat' : 'between'}"/></c:valAx>`;
    const title = spec.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="en-US" sz="1400" b="0"/><a:t>${esc(spec.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${RNS}"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${type === 'pie' ? '' : catAx + valAx}</c:plotArea><c:legend><c:legendPos val="${type === 'pie' ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
    const sheetPath = s.path;
    const at = { c: c2 + 2, r: g.r1 };
    await this.packageOp('Insert chart', async (pkg) => {
      const chartPath = pkg.free('xl/charts/chart#.xml');
      pkg.put(chartPath, xml);
      await pkg.override(chartPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      const dr = await drawingOf(pkg, sheetPath);
      const rid = await pkg.addRel(dr.path, REL('chart'), relative(dr.path, chartPath));
      const id = nextShapeId(dr.doc);
      const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${RNS}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
      dr.doc.documentElement.append(dr.doc.importNode(parseXml(anchorXml(at.c, at.r, at.c + 8, at.r + 16, frame).replace('editAs="oneCell"', '')).documentElement, true));
    });
  };

  // ======================================================================
  // Insert > Image (over cells)
  // ======================================================================

  P.insertPicture = async function (bytes, ext, w, h) {
    const sheetPath = this.sheet.path;
    const at = { ...this.active };
    // Size in cells: about 64 px per column and 20 px per row, picture at most 480 px wide.
    const sc = Math.min(1, 480 / w, 360 / h);
    const cols = Math.max(1, Math.round((w * sc) / 64)), rows = Math.max(1, Math.round((h * sc) / 20));
    await this.packageOp('Insert picture', async (pkg) => {
      const media = pkg.free(`xl/media/image#.${ext}`);
      pkg.zip.file(media, bytes);
      await pkg.default(ext, ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : `image/${ext}`);
      const dr = await drawingOf(pkg, sheetPath);
      const rid = await pkg.addRel(dr.path, REL('image'), relative(dr.path, media));
      const id = nextShapeId(dr.doc);
      const pic = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Picture ${id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${RNS}" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(w * sc * 9525)}" cy="${Math.round(h * sc * 9525)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`;
      dr.doc.documentElement.append(dr.doc.importNode(parseXml(anchorXml(at.c, at.r, at.c + cols, at.r + rows, pic)).documentElement, true));
    });
  };

  // ======================================================================
  // Notes (Excel comments) - shown with a red corner, edited from Insert > Note
  // ======================================================================

  P.notes = function () {
    const s = this.sheet;
    if (s._notes) return s._notes;
    s._notes = new Map();
    return s._notes;
  };
  P.loadNotes = async function () {
    const s = this.sheet;
    const map = new Map();
    if (!s.isNew && s.path) {
      const rels = await this.wb.sheetRels(s);
      const rel = rels.list.find((r) => r.type.endsWith('/comments'));
      if (rel && rel.path) {
        const d = await this.wb.part(rel.path);
        if (d) {
          const authors = [...d.getElementsByTagNameNS(NS, 'author')].map((a) => a.textContent);
          for (const c of d.getElementsByTagNameNS(NS, 'comment')) {
            const p = parseAddr(c.getAttribute('ref'));
            if (!p) continue;
            const text = [...c.getElementsByTagNameNS(NS, 't')].map((t) => t.textContent).join('');
            map.set(key(p.r, p.c), { text, author: authors[+c.getAttribute('authorId')] || '' });
          }
        }
      }
    }
    s._notes = map;
    this.render(true);
    return map;
  };
  // Write the complete set of notes for the sheet (comments part + VML shapes Excel needs to show them).
  P.saveNotes = async function (map, label) {
    const s = this.sheet;
    const sheetPath = s.path;
    const author = (localStorage.getItem('docstudio.author') || 'Doc Studio').trim() || 'Doc Studio';
    await this.packageOp(label || 'Note', async (pkg) => {
      const sheet = await pkg.doc(sheetPath);
      let cpath = await pkg.relTarget(sheetPath, REL('comments'));
      if (!cpath && map.size) {
        cpath = pkg.free('xl/comments#.xml');
        await pkg.addRel(sheetPath, REL('comments'), relative(sheetPath, cpath));
        await pkg.override(cpath, 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
      }
      if (!cpath) return;
      const authors = [...new Set([...map.values()].map((n) => n.author || author))];
      const list = [...map.entries()].sort((a, b) => { const x = SheetModel.unkey(a[0]), y = SheetModel.unkey(b[0]); return x.r - y.r || x.c - y.c; });
      const cellRef = (k) => { const p = SheetModel.unkey(k); return addr(p.r, p.c); };
      pkg.put(cpath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<comments xmlns="${NS}"><authors>${authors.map((a) => `<author>${esc(a)}</author>`).join('')}</authors><commentList>${list.map(([k, n]) => `<comment ref="${cellRef(k)}" authorId="${authors.indexOf(n.author || author)}"><text><r><t xml:space="preserve">${esc(n.text)}</t></r></text></comment>`).join('')}</commentList></comments>`);
      // VML: one hidden note shape per comment (Excel shows it on hover).
      let vpath = await pkg.relTarget(sheetPath, REL('vmlDrawing'));
      if (!vpath) {
        vpath = pkg.free('xl/drawings/vmlDrawing#.vml');
        const rid = await pkg.addRel(sheetPath, REL('vmlDrawing'), relative(sheetPath, vpath));
        await pkg.default('vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing');
        const ld = wsEl(sheet, 'legacyDrawing');
        ld.setAttributeNS(RNS, 'r:id', rid);
        insertOrdered(sheet.documentElement, ld);
      }
      const shapes = list.map(([k], i) => {
        const p = SheetModel.unkey(k);
        return `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${p.c}, 15, ${Math.max(0, p.r - 2)}, 2, ${p.c + 2}, 15, ${p.r + 2}, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${p.r - 1}</x:Row><x:Column>${p.c - 1}</x:Column></x:ClientData></v:shape>`;
      }).join('');
      pkg.put(vpath, `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes}</xml>`);
    });
    await this.loadNotes();
  };
  P.editNote = async function () {
    const notes = this.sheet._notes || await this.loadNotes();
    const k = key(this.active.r, this.active.c);
    const cur = notes.get(k);
    const v = await DocTools.form(`Note for ${addr(this.active.r, this.active.c)}`, [{ name: 'text', label: 'Note (leave empty to delete it)', type: 'textarea', value: cur ? cur.text : '' }], 'Save');
    if (!v) return;
    const next = new Map(notes);
    if (v.text.trim()) next.set(k, { text: v.text, author: cur ? cur.author : (localStorage.getItem('docstudio.author') || 'Doc Studio') });
    else if (cur) next.delete(k); else return;
    await this.saveNotes(next, v.text.trim() ? 'Note' : 'Delete note');
  };

  // ======================================================================
  // Data validation: dropdowns and checkboxes
  // ======================================================================

  P.validations = function () {
    const s = this.sheet;
    if (s._dv && s._dvDoc === s.doc) return s._dv;
    const out = [];
    for (const dv of s.doc.getElementsByTagNameNS(NS, 'dataValidation')) {
      const ranges = (dv.getAttribute('sqref') || '').split(/\s+/).map(parseRange).filter(Boolean);
      const f1 = dv.getElementsByTagNameNS(NS, 'formula1')[0];
      const formula = f1 ? f1.textContent : '';
      let values = null;
      if (dv.getAttribute('type') === 'list' && /^".*"$/.test(formula)) values = formula.slice(1, -1).split(',');
      const checkbox = !!values && values.length === 2 && values.map((x) => x.trim().toUpperCase()).sort().join(',') === 'FALSE,TRUE';
      out.push({ el: dv, ranges, type: dv.getAttribute('type'), formula, values, checkbox });
    }
    s._dv = out; s._dvDoc = s.doc;
    return out;
  };
  P.dvAt = function (r, c) { return this.validations().find((v) => v.ranges.some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)) || null; };
  // spec: { type: 'list', values } | { checkbox: true } | { type: 'whole'|'decimal', op, a, b } | null (remove)
  P.setValidation = function (spec) {
    const g = rangeOf(this);
    const s = this.sheet;
    return this.snapshotOp(spec ? 'Data validation' : 'Remove data validation', () => {
      const doc = s.doc, root = doc.documentElement;
      let dvs = kidsOf(root, 'dataValidations')[0];
      // Drop rules that cover only the selection; shrink nothing else (keeps it simple and Excel-valid).
      if (dvs) for (const dv of kidsOf(dvs, 'dataValidation')) {
        const ranges = (dv.getAttribute('sqref') || '').split(/\s+/).map(parseRange).filter(Boolean);
        if (ranges.every((x) => x.r1 >= g.r1 && x.r2 <= g.r2 && x.c1 >= g.c1 && x.c2 <= g.c2)) dv.remove();
      }
      if (spec) {
        if (!dvs) dvs = insertOrdered(root, wsEl(doc, 'dataValidations'));
        const attrs = { allowBlank: '1', showInputMessage: '1', showErrorMessage: '1', sqref: rangeRef(g) };
        let f1 = null, f2 = null;
        if (spec.checkbox || spec.type === 'list') {
          attrs.type = 'list';
          f1 = spec.checkbox ? '"TRUE,FALSE"' : `"${spec.values.map((v) => String(v).replace(/[",]/g, '')).join(',')}"`;
          if (spec.checkbox) attrs.showDropDown = '1'; // no in-cell arrow in Excel for checkboxes
        } else {
          attrs.type = spec.type; attrs.operator = spec.op || 'between';
          f1 = String(spec.a); if (spec.b != null && spec.b !== '') f2 = String(spec.b);
        }
        const dv = wsEl(doc, 'dataValidation', attrs);
        const a = wsEl(doc, 'formula1'); a.textContent = f1; dv.append(a);
        if (f2) { const b = wsEl(doc, 'formula2'); b.textContent = f2; dv.append(b); }
        dvs.append(dv);
        if (spec.checkbox) {
          const r2 = Math.min(g.r2, g.r1 + 5000), c2 = Math.min(g.c2, g.c1 + 200);
          for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) { const v = valueOf(this, r, c); if (typeof v !== 'boolean') s.setCell(r, c, { value: false, formula: null }); }
        }
      }
      if (dvs) { if (!kidsOf(dvs, 'dataValidation').length) dvs.remove(); else dvs.setAttribute('count', kidsOf(dvs, 'dataValidation').length); }
      s._dv = null;
      s.structureChanged = true;
    });
  };

  // ======================================================================
  // Drawing extras into cells
  // ======================================================================

  P.decorateCell = function (r, c, cell, disp, d) {
    let out = disp;
    if (this.showFormulas && cell && cell.formula && !cell.arrayMaster) out = { text: `=${cell.formula}`, kind: 'text' };
    const dv = this.validations().length ? this.dvAt(r, c) : null;
    if (dv && dv.checkbox) {
      const v = cell ? cell.value : null;
      out = { text: v === true || (typeof v === 'string' && v.toUpperCase() === 'TRUE') ? '☑' : '☐', kind: 'bool' };
      d.classList.add('sc-check');
    } else if (dv && dv.values && r === this.active.r && c === this.active.c) {
      const b = document.createElement('b');
      b.className = 'sg-dd-btn';
      b.title = 'Choose from the list';
      d.append(b);
    }
    const notes = this.sheet._notes;
    if (notes && notes.size) {
      const n = notes.get(key(r, c));
      if (n) { const m = document.createElement('i'); m.className = 'sg-note-mark'; d.append(m); d.title = `${n.author ? `${n.author}:\n` : ''}${n.text}`; }
    }
    if (cell && cell.formula && /^HYPERLINK\(/i.test(cell.formula)) { d.classList.add('sc-link'); d.title = 'Ctrl+click to open the link'; }
    return out;
  };

  // Clicks: checkbox toggles, the dropdown arrow, Ctrl+click on links.
  function wireClicks(view) {
    if (view._toolsWired) return;
    view._toolsWired = true;
    view.grid.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      const t = e.target;
      if (t.classList && t.classList.contains('sg-dd-btn')) {
        e.preventDefault(); e.stopPropagation();
        const dv = view.dvAt(view.active.r, view.active.c);
        if (!dv || !dv.values) return;
        const r = t.getBoundingClientRect();
        SheetView.showMenu(r.left, r.bottom + 2, dv.values.map((v) => [v, '', () => { view.setFromText(view.active.r, view.active.c, v); view.grid.focus(); }]));
        return;
      }
      const h = view.hit(e);
      if (h.inColHdr || h.inRowHdr) return;
      const cell = view.sheet.get(h.r, h.c);
      if (e.ctrlKey && cell && cell.formula && /^HYPERLINK\(/i.test(cell.formula)) {
        const m = /^HYPERLINK\(\s*"([^"]+)"/i.exec(cell.formula);
        if (m && /^(https?:|mailto:)/i.test(m[1])) { e.preventDefault(); e.stopPropagation(); desk.openExternal(m[1]); }
        return;
      }
      const dv = view.validations().length ? view.dvAt(h.r, h.c) : null;
      if (dv && dv.checkbox && !e.shiftKey && !e.ctrlKey) setTimeout(() => toggleCheckbox(view, h.r, h.c), 0);
    }, true);
    view.grid.addEventListener('keydown', (e) => {
      if (view.editing || e.key !== ' ' || e.ctrlKey || e.altKey) return;
      const dv = view.dvAt(view.active.r, view.active.c);
      if (dv && dv.checkbox) { e.preventDefault(); e.stopPropagation(); toggleCheckbox(view, view.active.r, view.active.c); }
    }, true);
  }
  function toggleCheckbox(view, r, c) {
    const v = valueOf(view, r, c);
    const on = v === true || (typeof v === 'string' && v.toUpperCase() === 'TRUE');
    view.apply([{ r, c, patch: { value: !on, formula: null } }]);
  }

  // Load notes and wire clicks whenever a sheet is shown.
  const showSheet = P.showSheet;
  P.showSheet = function (i) {
    const out = showSheet.call(this, i);
    wireClicks(this);
    if (this.sheet && !this.sheet._notes) this.loadNotes().catch(() => {});
    if (this.zoom && this.zoom !== 1) this.grid.style.zoom = String(this.zoom);
    return out;
  };

  // ======================================================================
  // Protection: a protected sheet can't be edited until unprotected
  // ======================================================================

  P.isProtected = function () { return !!(this.sheet && this.sheet.doc) && kidsOf(this.sheet.doc.documentElement, 'sheetProtection').some((p) => p.getAttribute('sheet') === '1' || p.getAttribute('sheet') === 'true'); };
  P.toggleProtect = function () {
    const s = this.sheet;
    const on = this.isProtected();
    return this.snapshotOp(on ? 'Unprotect sheet' : 'Protect sheet', () => {
      const root = s.doc.documentElement;
      const old = kidsOf(root, 'sheetProtection')[0];
      if (old) old.remove();
      if (!on) insertOrdered(root, wsEl(s.doc, 'sheetProtection', { sheet: '1', objects: '1', scenarios: '1' }));
      s.structureChanged = true;
    });
  };
  const apply = P.apply;
  P.apply = function (items, sheet) {
    if ((!sheet || sheet === this.sheet) && this.isProtected()) {
      this.opts.error && this.opts.error('This sheet is protected.', 'Use Data > Protect sheet to unprotect it first.');
      return;
    }
    return apply.call(this, items, sheet);
  };

  // ======================================================================
  // Data tools
  // ======================================================================

  // The selection, or the data block around a single selected cell; trimmed to the used area.
  function dataRange(view) {
    let g = rangeOf(view);
    if (g.r1 === g.r2 && g.c1 === g.c2 && view.currentRegion) g = view.currentRegion();
    const s = view.sheet;
    return { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, Math.max(s.maxRow, g.r1)), c2: Math.min(g.c2, Math.max(s.maxCol, g.c1)) };
  }
  const snap = (cell) => (cell ? { value: cell.value, formula: cell.formula, s: cell.s } : { value: null, formula: null, s: 0 });
  // Move whole rows of a range into a new order (values, formulas, formats).
  function rewriteRows(view, g, order, label) {
    const items = [];
    const rows = order.map((r) => { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(snap(view.sheet.get(r, c))); return row; });
    for (let i = 0; i < g.r2 - g.r1 + 1; i++) {
      for (let c = g.c1; c <= g.c2; c++) {
        const src = rows[i] ? rows[i][c - g.c1] : { value: null, formula: null, s: view.sheet.get(g.r1 + i, c) ? view.sheet.get(g.r1 + i, c).s : 0 };
        const patch = { value: src.value, formula: src.formula && rows[i] ? SheetModel.shiftFormula(src.formula, g.r1 + i - order[i], 0) : null, s: src.s };
        items.push({ r: g.r1 + i, c, patch });
      }
    }
    view.apply(items);
    void label;
  }
  P.removeDuplicates = async function () {
    const g = dataRange(this);
    const v = await DocTools.form(`Remove duplicates in ${rangeRef(g)}`, [{ name: 'header', label: 'Data has a header row', type: 'checkbox', value: true }], 'Remove duplicates');
    if (!v) return;
    const start = v.header ? g.r1 + 1 : g.r1;
    const seen = new Set();
    const keep = [];
    let removed = 0;
    for (let r = start; r <= g.r2; r++) {
      const k = Array.from({ length: g.c2 - g.c1 + 1 }, (_, i) => textOf(this, r, g.c1 + i).trim()).join('\u0001');
      if (seen.has(k)) { removed++; continue; }
      seen.add(k); keep.push(r);
    }
    if (!removed) { this.opts.onStatus && this.opts.onStatus('No duplicate rows found'); return; }
    rewriteRows(this, { ...g, r1: start }, keep, 'Remove duplicates');
    this.opts.onStatus && this.opts.onStatus(`Removed ${removed} duplicate row(s); ${keep.length} unique row(s) remain`);
  };
  P.trimWhitespace = function () {
    const g = dataRange(this);
    const items = [];
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const cell = this.sheet.get(r, c);
      if (!cell || cell.formula || typeof cell.value !== 'string') continue;
      const t = cell.value.replace(/[\s\u00a0]+/g, ' ').trim();
      if (t !== cell.value) items.push({ r, c, patch: { value: t } });
    }
    if (items.length) this.apply(items);
    this.opts.onStatus && this.opts.onStatus(`Trimmed ${items.length} cell(s)`);
  };
  P.splitToColumns = async function () {
    const g = rangeOf(this);
    const v = await DocTools.form('Split text to columns', [{ name: 'sep', label: 'Separator', type: 'select', value: 'auto', options: [['auto', 'Detect automatically'], [',', 'Comma'], [';', 'Semicolon'], ['.', 'Period'], [' ', 'Space'], ['\t', 'Tab'], ['custom', 'Custom…']] }, { name: 'custom', label: 'Custom separator', value: '' }], 'Split');
    if (!v) return;
    const r2 = Math.min(g.r2, Math.max(this.sheet.maxRow, g.r1));
    const texts = [];
    for (let r = g.r1; r <= r2; r++) texts.push(textOf(this, r, g.c1));
    let sep = v.sep === 'custom' ? v.custom : v.sep;
    if (sep === 'auto') sep = [',', ';', '\t', '|', ' '].find((x) => texts.some((t) => t.includes(x))) || ',';
    if (!sep) return;
    const items = [];
    texts.forEach((t, i) => {
      if (!t.includes(sep)) return;
      t.split(sep).forEach((part, j) => items.push({ r: g.r1 + i, c: g.c1 + j, patch: SheetView.parseInput ? SheetView.parseInput(part.trim(), true, this.wb.date1904) : { value: part.trim() } }));
    });
    if (items.length) this.apply(items);
  };
  P.randomizeRange = function () {
    const g = dataRange(this);
    const order = [];
    for (let r = g.r1; r <= g.r2; r++) order.push(r);
    for (let i = order.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [order[i], order[j]] = [order[j], order[i]]; }
    rewriteRows(this, g, order, 'Randomize range');
  };
  P.columnStats = function () {
    const g = dataRange(this);
    const c = this.active.c;
    const vals = [];
    let blanks = 0;
    for (let r = g.r1; r <= g.r2; r++) {
      const v = valueOf(this, r, c);
      if (v === null || v === '') blanks++; else vals.push(v);
    }
    const nums = vals.filter((v) => typeof v === 'number');
    const freq = new Map();
    for (const v of vals) { const k = String(v); freq.set(k, (freq.get(k) || 0) + 1); }
    const top = [...freq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 5);
    const sum = nums.reduce((a, b) => a + b, 0);
    const sorted = [...nums].sort((a, b) => a - b);
    const median = sorted.length ? (sorted.length % 2 ? sorted[(sorted.length - 1) / 2] : (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2) : null;
    const fmt = (n) => (n == null ? '—' : Number.isInteger(n) ? n.toLocaleString() : n.toLocaleString(undefined, { maximumFractionDigits: 4 }));
    desk.confirm({ type: 'info', buttons: ['OK'], message: `Column stats - ${colName(c)} (${rangeRef({ r1: g.r1, c1: c, r2: g.r2, c2: c })})`, detail: [
      `Count: ${vals.length}`, `Blank: ${blanks}`, `Unique: ${freq.size}`,
      ...(nums.length ? [`Sum: ${fmt(sum)}`, `Average: ${fmt(sum / nums.length)}`, `Median: ${fmt(median)}`, `Min: ${fmt(sorted[0])}`, `Max: ${fmt(sorted[sorted.length - 1])}`] : []),
      '', 'Most frequent:', ...top.map(([k, n]) => `  ${k}  (${n})`),
    ].join('\n') });
  };
  P.sortRangeDialog = async function () {
    const g = dataRange(this);
    const cols = [];
    for (let c = g.c1; c <= g.c2; c++) cols.push([String(c), `Column ${colName(c)}${textOf(this, g.r1, c) ? ` (${textOf(this, g.r1, c)})` : ''}`]);
    const v = await DocTools.form(`Sort range ${rangeRef(g)}`, [
      { name: 'header', label: 'Data has a header row', type: 'checkbox', value: true },
      { name: 'col', label: 'Sort by', type: 'select', value: String(this.active.c >= g.c1 && this.active.c <= g.c2 ? this.active.c : g.c1), options: cols },
      { name: 'order', label: 'Order', type: 'select', value: 'asc', options: [['asc', 'A → Z'], ['desc', 'Z → A']] },
    ], 'Sort');
    if (!v) return;
    return this.sortRange({ ...g, r1: v.header ? g.r1 + 1 : g.r1 }, +v.col, v.order === 'asc');
  };

  // ---------- named ranges

  P.namedRanges = function () { return this.wb.definedNames.filter((d) => !d.name.startsWith('_xlnm.')); };
  P.addNamedRange = async function () {
    const g = rangeOf(this);
    const v = await DocTools.form('Named range', [{ name: 'name', label: 'Name (letters, numbers, _; no spaces)', value: '' }, { name: 'ref', label: 'Range', value: `${quoteSheet(this.sheet.name)}!$${colName(g.c1)}$${g.r1}:$${colName(g.c2)}$${g.r2}` }], 'Done');
    if (!v || !v.name) return;
    const name = v.name.trim();
    if (!/^[A-Za-z_\\][\w.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name)) { this.opts.error && this.opts.error('That name is not allowed.', 'Use letters, numbers and _; it must not look like a cell address.'); return; }
    if (this.wb.definedNames.some((d) => d.name.toLowerCase() === name.toLowerCase() && d.local == null)) { this.opts.error && this.opts.error(`The name ${name} is already used.`); return; }
    return this.snapshotOp('Named range', () => {
      const doc = this.wb.wbDoc, root = doc.documentElement;
      const p = root.prefix ? `${root.prefix}:` : '';
      let dn = kidsOf(root, 'definedNames')[0];
      if (!dn) {
        dn = doc.createElementNS(NS, `${p}definedNames`);
        const after = kidsOf(root).find((c) => ['calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'].includes(c.localName));
        root.insertBefore(dn, after || null);
      }
      const el = doc.createElementNS(NS, `${p}definedName`);
      el.setAttribute('name', name);
      el.textContent = v.ref;
      dn.append(el);
      this.wb.definedNames.push({ name, local: null, ref: v.ref, el });
      this.wb.wbChanged = true;
    });
  };
  P.deleteNamedRange = function (d) {
    return this.snapshotOp('Delete named range', () => {
      d.el.parentNode.removeChild(d.el);
      this.wb.definedNames = this.wb.definedNames.filter((x) => x !== d);
      this.wb.wbChanged = true;
    });
  };
  P.namedRangesDialog = function () {
    const list = this.namedRanges();
    const box = h('div', { class: 'dialog versions-dialog' }, h('h3', { text: 'Named ranges' }));
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    if (!list.length) box.append(h('p', { class: 'muted', text: 'No named ranges yet.' }));
    for (const d of list) box.append(h('div', { class: 'version-row' }, h('b', { text: d.name }), h('span', { class: 'muted', text: d.ref }), h('button', { class: 'ghost', text: 'Go to', onclick: () => { close(); this.goTo(d.name); } }), h('button', { class: 'ghost', text: 'Delete', onclick: () => { close(); this.deleteNamedRange(d); } })));
    box.append(h('div', { class: 'dialog-actions' }, h('button', { class: 'ghost', text: 'Add a range…', onclick: () => { close(); this.addNamedRange(); } }), h('button', { text: 'Close', onclick: close })));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
  };

  // ---------- pivot table (a summary sheet with live SUMIFS/COUNTIFS/... formulas)

  P.pivotTable = async function () {
    const g = dataRange(this);
    if (g.r2 <= g.r1) throw new Error('Select a block of data with a header row first.');
    const heads = [];
    for (let c = g.c1; c <= g.c2; c++) heads.push([String(c), textOf(this, g.r1, c) || `Column ${colName(c)}`]);
    const v = await DocTools.form(`Pivot table from ${rangeRef(g)}`, [
      { name: 'rows', label: 'Rows', type: 'select', value: heads[0][0], options: heads },
      { name: 'cols', label: 'Columns (optional)', type: 'select', value: '', options: [['', 'None'], ...heads] },
      { name: 'values', label: 'Values', type: 'select', value: heads[heads.length - 1][0], options: heads },
      { name: 'fn', label: 'Summarize by', type: 'select', value: 'SUM', options: [['SUM', 'SUM'], ['COUNT', 'COUNT'], ['AVERAGE', 'AVERAGE'], ['MIN', 'MIN'], ['MAX', 'MAX']] },
    ], 'Create');
    if (!v) return;
    const src = this.sheet;
    const sh = quoteSheet(src.name);
    const col = (c) => `${sh}!$${colName(+c)}$${g.r1 + 1}:$${colName(+c)}$${g.r2}`;
    const uniq = (c) => { const set = new Map(); for (let r = g.r1 + 1; r <= g.r2; r++) { const t = textOf(this, r, +c); if (t !== '' && !set.has(t.toLowerCase())) set.set(t.toLowerCase(), { t, v: valueOf(this, r, +c) }); } return [...set.values()].sort((a, b) => (typeof a.v === 'number' && typeof b.v === 'number' ? a.v - b.v : String(a.t).localeCompare(String(b.t)))); };
    const rowsU = uniq(v.rows);
    const colsU = v.cols ? uniq(v.cols) : null;
    if (rowsU.length > 2000) throw new Error('Too many different row values for a pivot table.');
    const fn = { SUM: 'SUMIFS', COUNT: 'COUNTIFS', AVERAGE: 'AVERAGEIFS', MIN: '_xlfn.MINIFS', MAX: '_xlfn.MAXIFS' }[v.fn];
    const crit = (range, cellRef) => `${range},${cellRef}`;
    const formula = (rowCell, colCell) => {
      const parts = [];
      if (v.fn !== 'COUNT') parts.push(col(v.values));
      parts.push(crit(col(v.rows), rowCell));
      if (colCell) parts.push(crit(col(v.cols), colCell));
      return `${fn}(${parts.join(',')})`;
    };
    const lit = (x) => (typeof x.v === 'number' ? { value: x.v } : { value: x.t });
    await this.snapshotOp('Pivot table', () => {
      let n = 1;
      while (this.wb.sheet(`Pivot Table ${n}`)) n++;
      const ps = this.wb.addSheet(`Pivot Table ${n}`);
      const set = (r, c, patch) => ps.setCell(r, c, patch);
      const valName = heads.find((x) => x[0] === v.values)[1];
      set(1, 1, { value: heads.find((x) => x[0] === v.rows)[1] });
      if (!colsU) {
        set(1, 2, { value: `${v.fn[0]}${v.fn.slice(1).toLowerCase()} of ${valName}` });
        rowsU.forEach((x, i) => { set(2 + i, 1, lit(x)); set(2 + i, 2, { formula: formula(`$A${2 + i}`) }); });
        const last = 1 + rowsU.length;
        set(last + 1, 1, { value: 'Grand total' });
        set(last + 1, 2, { formula: v.fn === 'COUNT' || v.fn === 'SUM' ? `SUM(B2:B${last})` : `${v.fn === 'AVERAGE' ? 'AVERAGE' : v.fn}(${col(v.values)})` });
      } else {
        colsU.forEach((x, j) => set(1, 2 + j, lit(x)));
        const tc = 2 + colsU.length;
        set(1, tc, { value: 'Grand total' });
        rowsU.forEach((x, i) => {
          set(2 + i, 1, lit(x));
          colsU.forEach((_, j) => set(2 + i, 2 + j, { formula: formula(`$A${2 + i}`, `${colName(2 + j)}$1`) }));
          set(2 + i, tc, { formula: formula(`$A${2 + i}`) });
        });
        const last = 1 + rowsU.length;
        set(last + 1, 1, { value: 'Grand total' });
        for (let j = 0; j <= colsU.length; j++) set(last + 1, 2 + j, { formula: v.fn === 'COUNT' || v.fn === 'SUM' ? `SUM(${colName(2 + j)}2:${colName(2 + j)}${last})` : `${v.fn === 'AVERAGE' ? 'AVERAGE' : v.fn}(${colName(2 + j)}2:${colName(2 + j)}${last})` });
      }
      for (const cell of ps.cells.values()) if (cell.formula) cell.formulaChanged = true;
      this.wb.formulasChanged = true;
    });
    this.showSheet(this.wb.sheets.length - 1);
  };

  // ======================================================================
  // Format
  // ======================================================================

  P.customNumberFormat = async function () {
    const v = await DocTools.form('Custom number format', [{ name: 'fmt', label: 'Format code (for example #,##0.00 or 0.0% or mmm d, yyyy or "$"#,##0;[Red]-"$"#,##0)', value: this.activeCss().fmt || 'General' }], 'Apply');
    if (!v || !v.fmt.trim()) return;
    this.styleSelection({ numFmt: v.fmt.trim() });
  };
  P.alternatingColors = async function () {
    const v = await DocTools.form('Alternating colours', [{ name: 'color', label: 'Band colour', type: 'color', value: '#e8f0fe' }, { name: 'header', label: 'Highlight the header row', type: 'checkbox', value: true }], 'Apply');
    if (!v) return;
    const g = dataRange(this);
    const sel = this.sel;
    this.sel = { ...g };
    try {
      await this.addCfRule({ type: 'expression', formulas: [`MOD(ROW(),2)=${g.r1 % 2 === 0 ? 1 : 0}`], dxf: { fill: v.color } });
      if (v.header) { this.sel = { ...g, r2: g.r1 }; this.styleSelection({ fill: '#4472c4', font: { bold: true, color: '#ffffff' } }); }
    } finally { this.sel = sel; this.paintSelection(); }
  };
  P.convertToTable = async function () {
    const g = dataRange(this);
    if (g.r2 <= g.r1) throw new Error('Select a block of data with a header row first.');
    if (this.wb.tables.some((t) => t.sheet === this.sheet && !(t.g.r2 < g.r1 || t.g.r1 > g.r2 || t.g.c2 < g.c1 || t.g.c1 > g.c2))) throw new Error('This range overlaps a table already.');
    // Header cells must be unique text.
    const names = [];
    const fixes = [];
    for (let c = g.c1; c <= g.c2; c++) {
      let n = textOf(this, g.r1, c).trim() || `Column${c - g.c1 + 1}`;
      let k = 2;
      while (names.some((x) => x.toLowerCase() === n.toLowerCase())) n = `${n}${k++}`;
      names.push(n);
      const cur = valueOf(this, g.r1, c);
      if (cur !== n) fixes.push({ r: g.r1, c, patch: { value: n, formula: null } });
    }
    if (fixes.length) this.apply(fixes);
    const sheetPath = this.sheet.path;
    const hadFilter = this.sheet.autoFilter && parseRange(this.sheet.autoFilter.ref);
    const ids = this.wb.tables.map((t) => +t.doc.documentElement.getAttribute('id') || 0);
    const id = Math.max(0, ...ids) + 1;
    let tn = id;
    while (this.wb.tables.some((t) => t.name.toLowerCase() === `table${tn}`)) tn++;
    const ref = rangeRef(g);
    await this.packageOp('Convert to table', async (pkg) => {
      const path = pkg.free('xl/tables/table#.xml');
      pkg.put(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<table xmlns="${NS}" id="${id}" name="Table${tn}" displayName="Table${tn}" ref="${ref}" totalsRowShown="0"><autoFilter ref="${ref}"/><tableColumns count="${names.length}">${names.map((n, i) => `<tableColumn id="${i + 1}" name="${esc(n)}"/>`).join('')}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`);
      await pkg.override(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml');
      const rid = await pkg.addRel(sheetPath, REL('table'), relative(sheetPath, path));
      const sheet = await pkg.doc(sheetPath);
      const root = sheet.documentElement;
      // A sheet filter over the same cells would clash with the table's own.
      if (hadFilter && !(hadFilter.r2 < g.r1 || hadFilter.r1 > g.r2 || hadFilter.c2 < g.c1 || hadFilter.c1 > g.c2)) { const af = kidsOf(root, 'autoFilter')[0]; if (af) af.remove(); }
      let parts = kidsOf(root, 'tableParts')[0];
      if (!parts) parts = insertOrdered(root, wsEl(sheet, 'tableParts'));
      const tp = wsEl(sheet, 'tablePart');
      tp.setAttributeNS(RNS, 'r:id', rid);
      parts.append(tp);
      parts.setAttribute('count', kidsOf(parts, 'tablePart').length);
    });
  };

  // Paint format: copy the active cell's format, apply it to the next selection.
  P.startPaintFormat = function (btn) {
    if (this.painter) { this.painter.btn && this.painter.btn.classList.remove('on'); this.painter = null; return; }
    const cell = this.sheet.get(this.active.r, this.active.c);
    this.painter = { s: cell ? cell.s : this.sheet.defaultStyleAt(this.active.r, this.active.c), btn };
    if (btn) btn.classList.add('on');
    this.opts.onStatus && this.opts.onStatus('Select the cells to apply the format to (Esc cancels)');
    const done = () => {
      window.removeEventListener('mouseup', done, true);
      setTimeout(() => {
        if (!this.painter) return;
        const p = this.painter;
        this.painter = null;
        if (p.btn) p.btn.classList.remove('on');
        const g = this.effectiveRange();
        const items = [];
        for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 300); c++) items.push({ r, c, patch: { s: p.s } });
        this.apply(items);
      }, 0);
    };
    setTimeout(() => window.addEventListener('mouseup', done, true), 0);
  };

  // ======================================================================
  // View
  // ======================================================================

  P.toggleGridlines = function () {
    const s = this.sheet;
    return this.snapshotOp(s.showGrid ? 'Hide gridlines' : 'Show gridlines', () => {
      s.showGrid = !s.showGrid;
      const doc = s.doc, root = doc.documentElement;
      let views = kidsOf(root, 'sheetViews')[0];
      if (!views) { views = wsEl(doc, 'sheetViews'); insertOrdered(root, views); }
      let v = kidsOf(views, 'sheetView')[0];
      if (!v) { v = wsEl(doc, 'sheetView', { workbookViewId: '0' }); views.append(v); }
      if (s.showGrid) v.removeAttribute('showGridLines'); else v.setAttribute('showGridLines', '0');
      s.structureChanged = true;
    });
  };
  P.toggleShowFormulas = function () { this.showFormulas = !this.showFormulas; this.render(true); };
  P.setZoom = function (z) {
    this.zoom = Math.max(0.5, Math.min(2, z));
    this.grid.style.zoom = this.zoom === 1 ? '' : String(this.zoom);
    this.layoutDirty = true;
    this.render(true);
  };

  // ======================================================================
  // Sheet tabs: duplicate, delete, hide/unhide, colour, move
  // ======================================================================

  P.duplicateSheet = function (i) {
    const src = this.wb.sheets[i];
    let n = 1, name = `Copy of ${src.name}`;
    while (this.wb.sheet(name)) name = `Copy of ${src.name} ${++n}`;
    const wbPath = this.wb.wbPath;
    return this.packageOp('Duplicate sheet', async (pkg) => {
      const xml = await pkg.zip.file(src.path).async('string');
      const doc = parseXml(xml);
      const root = doc.documentElement;
      // Parts that belong to one sheet only (tables, drawings, notes) aren't copied.
      for (const name2 of ['tableParts', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls']) for (const e of kidsOf(root, name2)) e.remove();
      for (const v of doc.getElementsByTagNameNS(NS, 'sheetView')) v.removeAttribute('tabSelected');
      const path = pkg.free(wbPath.replace(/[^/]+$/, 'worksheets/sheet#.xml'));
      pkg.put(path, doc);
      // Hyperlinks and printer settings keep their relationships.
      const srcRels = await pkg.doc(pkg.relsPath(src.path));
      if (srcRels) {
        const keep = [...srcRels.documentElement.children].filter((r) => /\/hyperlink$|\/printerSettings$/.test(r.getAttribute('Type')));
        if (keep.length) { const d = parseXml(`<Relationships xmlns="${PKG}"/>`); for (const r of keep) d.documentElement.append(d.importNode(r, true)); pkg.put(pkg.relsPath(path), d); }
      }
      await pkg.override(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
      const rid = await pkg.addRel(wbPath, REL('worksheet'), relative(wbPath, path));
      const wb = await pkg.doc(wbPath);
      const sheets = wb.getElementsByTagNameNS(NS, 'sheet');
      const ids = [...sheets].map((x) => +x.getAttribute('sheetId'));
      const el = wb.importNode(sheets[i], false);
      el.setAttribute('name', name);
      el.setAttribute('sheetId', String(Math.max(0, ...ids) + 1));
      el.removeAttribute('state');
      el.setAttributeNS(RNS, 'r:id', rid);
      sheets[i].after(el);
      shiftLocalNames(wb, (x) => (x > i ? x + 1 : x));
    }).then(() => this.showSheet(i + 1));
  };
  function shiftLocalNames(wbDoc, map) {
    for (const d of [...wbDoc.getElementsByTagNameNS(NS, 'definedName')]) {
      if (!d.hasAttribute('localSheetId')) continue;
      const n = map(+d.getAttribute('localSheetId'));
      if (n == null) d.remove(); else d.setAttribute('localSheetId', String(n));
    }
  }
  P.deleteSheet = async function (i, confirmed) {
    const visible = this.wb.sheets.filter((s) => s.state === 'visible');
    if (visible.length <= 1) { this.opts.error && this.opts.error('A workbook needs at least one visible sheet.'); return; }
    if (!confirmed) {
      const r = await desk.confirm({ type: 'warning', message: `Delete the sheet "${this.wb.sheets[i].name}"?`, detail: 'Formulas on other sheets that use it will show #REF!.', buttons: ['Delete', 'Cancel'], defaultId: 1, cancelId: 1 });
      if (r !== 0) return;
    }
    const victim = this.wb.sheets[i];
    const re = new RegExp(`(?:'${victim.name.replace(/'/g, "''").replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}'|${/^[A-Za-z_][\w.]*$/.test(victim.name) ? victim.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') : '(?!)'})!\\$?[A-Z]{1,3}\\$?\\d+(?::\\$?[A-Z]{1,3}\\$?\\d+)?`, 'gi');
    // Formulas elsewhere that point at the sheet become #REF!.
    for (const s of this.wb.sheets) {
      if (s === victim) continue;
      for (const [k, cell] of s.cells) {
        if (!cell.formula || !re.test(cell.formula)) continue;
        re.lastIndex = 0;
        if (cell.shared != null) s.unshare(cell.shared);
        cell.formula = cell.formula.replace(re, '#REF!');
        cell.formulaChanged = true;
        s.changed.add(k);
      }
    }
    this.wb.formulasChanged = true;
    const wbPath = this.wb.wbPath;
    await this.packageOp('Delete sheet', async (pkg) => {
      const wb = await pkg.doc(wbPath);
      const el = wb.getElementsByTagNameNS(NS, 'sheet')[i];
      const rid = el.getAttributeNS(RNS, 'id') || el.getAttribute('r:id');
      el.remove();
      const rels = await pkg.rels(wbPath);
      for (const rr of [...rels.documentElement.children]) if (rr.getAttribute('Id') === rid) rr.remove();
      pkg.zip.remove(victim.path);
      if (pkg.zip.file(pkg.relsPath(victim.path))) pkg.zip.remove(pkg.relsPath(victim.path));
      const ct = await pkg.doc('[Content_Types].xml');
      for (const o of [...ct.documentElement.children]) if (o.getAttribute('PartName') === `/${victim.path}`) o.remove();
      shiftLocalNames(wb, (x) => (x === i ? null : x > i ? x - 1 : x));
      const bv = wb.getElementsByTagNameNS(NS, 'workbookView')[0];
      if (bv) { bv.removeAttribute('activeTab'); bv.removeAttribute('firstSheet'); }
    });
    this.showSheet(Math.max(0, Math.min(i, this.wb.sheets.length - 1)));
  };
  P.setSheetState = function (i, state) {
    const s = this.wb.sheets[i];
    if (state === 'hidden' && this.wb.sheets.filter((x) => x.state === 'visible').length <= 1) { this.opts.error && this.opts.error('A workbook needs at least one visible sheet.'); return; }
    return this.snapshotOp(state === 'hidden' ? 'Hide sheet' : 'Unhide sheet', () => {
      s.state = state;
      if (state === 'visible') s.el.removeAttribute('state'); else s.el.setAttribute('state', 'hidden');
      const bv = this.wb.wbDoc.getElementsByTagNameNS(NS, 'workbookView')[0];
      if (bv) { bv.removeAttribute('activeTab'); bv.removeAttribute('firstSheet'); }
      this.wb.wbChanged = true;
    }).then(() => { if (state === 'hidden' && i === this.sheetIndex) this.showSheet(this.wb.sheets.findIndex((x) => x.state === 'visible')); else this.paintTabs(); });
  };
  P.setTabColor = function (i, color) {
    const s = this.wb.sheets[i];
    return this.snapshotOp('Sheet colour', () => {
      const doc = s.doc, root = doc.documentElement;
      let pr = kidsOf(root, 'sheetPr')[0];
      if (!pr) { pr = wsEl(doc, 'sheetPr'); root.insertBefore(pr, root.firstChild); }
      const old = kidsOf(pr, 'tabColor')[0];
      if (old) old.remove();
      if (color) pr.insertBefore(wsEl(doc, 'tabColor', { rgb: `FF${color.replace('#', '').toUpperCase()}` }), pr.firstChild);
      if (!pr.attributes.length && !kidsOf(pr).length) pr.remove();
      s.structureChanged = true;
      s._tabColor = color;
    }).then(() => this.paintTabs());
  };
  P.tabColorOf = function (s) {
    if (s._tabColor !== undefined) return s._tabColor;
    const pr = s.doc && kidsOf(s.doc.documentElement, 'sheetPr')[0];
    const tc = pr && kidsOf(pr, 'tabColor')[0];
    const rgb = tc && tc.getAttribute('rgb');
    s._tabColor = rgb ? `#${rgb.slice(-6)}` : (tc && tc.getAttribute('theme') != null && this.wb.theme[+tc.getAttribute('theme')] ? `#${this.wb.theme[+tc.getAttribute('theme')]}` : null);
    return s._tabColor;
  };
  P.moveSheet = function (i, dir) {
    const j = i + dir;
    if (j < 0 || j >= this.wb.sheets.length) return;
    const wbPath = this.wb.wbPath;
    return this.packageOp('Move sheet', async (pkg) => {
      const wb = await pkg.doc(wbPath);
      const els = [...wb.getElementsByTagNameNS(NS, 'sheet')];
      if (dir < 0) els[i].parentNode.insertBefore(els[i], els[j]); else els[j].after(els[i]);
      shiftLocalNames(wb, (x) => (x === i ? j : x === j ? i : x));
      const bv = wb.getElementsByTagNameNS(NS, 'workbookView')[0];
      if (bv) { bv.removeAttribute('activeTab'); bv.removeAttribute('firstSheet'); }
    }).then(() => this.showSheet(j));
  };
  P.unhideSheetsMenu = function (x, y) {
    const hidden = this.wb.sheets.map((s, i) => ({ s, i })).filter((o) => o.s.state !== 'visible' && !o.s.unsupported);
    if (!hidden.length) { this.opts.onStatus && this.opts.onStatus('No hidden sheets'); return; }
    SheetView.showMenu(x, y, hidden.map((o) => [o.s.name, '', () => this.setSheetState(o.i, 'visible')]));
  };

  // Tabs: colours, and a right-click menu.
  const paintTabs = P.paintTabs;
  P.paintTabs = function () {
    paintTabs.call(this);
    const visible = this.wb.sheets.map((s, i) => ({ s, i })).filter((o) => o.s.state === 'visible');
    const buttons = [...this.tabsBar.querySelectorAll('.sg-tab:not(.sg-add)')];
    buttons.forEach((b, k) => {
      const o = visible[k];
      if (!o) return;
      const color = this.tabColorOf(o.s);
      if (color) b.style.boxShadow = `inset 0 -3px 0 ${color}`;
      b.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        const run = (fn) => () => Promise.resolve(fn()).catch((err) => this.opts.error && this.opts.error("Couldn't do that", err.message));
        const colours = [['None', null], ['Red', '#e53935'], ['Orange', '#fb8c00'], ['Yellow', '#fdd835'], ['Green', '#43a047'], ['Blue', '#1e88e5'], ['Purple', '#8e24aa'], ['Grey', '#757575']];
        SheetView.showMenu(e.clientX, e.clientY, [
          ['Rename…', '', run(() => this.renameSheet(o.i))],
          ['Duplicate', '', run(() => this.duplicateSheet(o.i))],
          ['Delete', '', run(() => this.deleteSheet(o.i))],
          null,
          ['Move left', '', run(() => this.moveSheet(o.i, -1))],
          ['Move right', '', run(() => this.moveSheet(o.i, 1))],
          null,
          ['Hide sheet', '', run(() => this.setSheetState(o.i, 'hidden'))],
          ['Unhide sheets…', '', () => this.unhideSheetsMenu(e.clientX, e.clientY)],
          null,
          ...colours.map(([n, c]) => [`Colour: ${n}`, '', run(() => this.setTabColor(o.i, c))]),
        ]);
      });
    });
  };

  // ======================================================================
  // Insert > Function: every function the engine knows, searchable
  // ======================================================================

  const POPULAR = ['SUM', 'AVERAGE', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'IF', 'IFS', 'IFERROR', 'SUMIF', 'SUMIFS', 'COUNTIF', 'COUNTIFS', 'AVERAGEIF', 'VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'FILTER', 'SORT', 'UNIQUE', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'TODAY', 'NOW', 'DATE', 'YEAR', 'MONTH', 'DAY', 'EOMONTH', 'NETWORKDAYS', 'TEXT', 'CONCAT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM', 'UPPER', 'LOWER', 'PROPER', 'SUBSTITUTE', 'FIND', 'SEARCH', 'AND', 'OR', 'NOT', 'ABS', 'PMT', 'HYPERLINK'];
  P.functionPicker = function () {
    const all = (this.calc && this.calc.engine && this.calc.engine.names) ? this.calc.engine.names() : POPULAR;
    const box = h('div', { class: 'dialog chars-dialog' }, h('h3', { text: 'Insert function' }));
    const search = h('input', { type: 'search', class: 'chars-search', placeholder: `Search ${all.length} functions` });
    const list = h('div', { class: 'fn-list' });
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    const pick = (name) => {
      close();
      this.select(this.active.r, this.active.c);
      this.beginEdit(`=${name}(`, 'cell', true);
    };
    const fill = () => {
      const q = search.value.trim().toUpperCase();
      const names = q ? all.filter((n) => n.includes(q)).sort((a, b) => (a.startsWith(q) ? 0 : 1) - (b.startsWith(q) ? 0 : 1) || a.localeCompare(b)) : POPULAR.filter((n) => all.includes(n) || n === 'HYPERLINK');
      list.replaceChildren(...names.slice(0, 200).map((n) => h('button', { class: 'ghost fn-item', text: n, onclick: () => pick(n) })));
    };
    search.addEventListener('input', fill);
    search.addEventListener('keydown', (e) => { if (e.key === 'Enter') { const b = list.querySelector('button'); if (b) b.click(); } });
    box.append(search, h('p', { class: 'muted', text: 'Popular functions are listed; type to search them all. Picking one starts the formula in the selected cell.' }), list, h('div', { class: 'dialog-actions' }, h('button', { class: 'ghost', text: 'Close', onclick: close })));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
    fill();
    search.focus();
  };

  P.insertLink = async function () {
    const v = await DocTools.form('Insert link', [{ name: 'url', label: 'Link (https://… or mailto:…)', value: 'https://' }, { name: 'text', label: 'Text to show', value: textOf(this, this.active.r, this.active.c) }], 'Insert');
    if (!v || !/^(https?:|mailto:)/i.test(v.url.trim())) return;
    const q = (x) => x.replace(/"/g, '""');
    this.setFromText(this.active.r, this.active.c, `=HYPERLINK("${q(v.url.trim())}","${q(v.text || v.url)}")`);
    this.styleSelection({ font: { color: '#1155cc', underline: true } });
  };

  // ======================================================================
  // Menus for the workbook editor
  // ======================================================================

  SheetView.sheetMenus = function (tab, view) {
    const run = (label, fn) => async () => { try { if (view.editing) view.endEdit(true); await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } finally { if (!document.querySelector('.backdrop')) view.grid.focus({ preventScroll: true }); } };
    const it = (label, fn, o) => ({ label, run: run(label.toLowerCase(), fn), ...(o || {}) });
    const sub = (label, items) => ({ label, sub: items });
    const SEP = { sep: true };
    const chart = (type) => it(`${type[0].toUpperCase()}${type.slice(1)} chart`, async () => { const v = await DocTools.form('Chart title', [{ name: 't', label: 'Title (optional)', value: '' }], 'Insert chart'); if (v) await view.insertChart({ type, title: v.t }); });
    return {
      Edit: [it('Paste values only', async () => { const t = await desk.readClipboard(); if (t) view.pasteText(t.replace(/\r?\n$/, '')); }), it('Delete values', () => view.clearSelection(), { key: 'Delete' })],
      View: [
        it('Gridlines', () => view.toggleGridlines(), { checked: () => view.sheet.showGrid }),
        it('Formulas', () => view.toggleShowFormulas(), { checked: () => !!view.showFormulas, key: 'Ctrl+`' }),
        sub('Zoom', [50, 75, 90, 100, 125, 150, 200].map((z) => it(`${z}%`, () => view.setZoom(z / 100), { checked: () => Math.round((view.zoom || 1) * 100) === z }))),
        it('Freeze panes at the selected cell', () => view.toggleFreeze()),
      ],
      Insert: [
        sub('Chart', ['column', 'bar', 'line', 'area', 'pie', 'scatter'].map(chart)),
        it('Pivot table…', () => view.pivotTable()),
        it('Image over cells…', () => new Promise((resolve) => {
          const input = h('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif,image/bmp,image/webp' });
          input.addEventListener('change', async () => { const f = input.files[0]; if (f) { const { bytes, ext, width, height } = await imageForWord(f); await view.insertPicture(bytes, ext === 'jpeg' ? 'jpg' : ext, width, height); } resolve(); });
          input.click();
        })),
        SEP,
        it('Function…', () => view.functionPicker()),
        it('Link…', () => view.insertLink(), { key: 'Ctrl+K' }),
        it('Checkbox', () => view.setValidation({ checkbox: true })),
        it('Dropdown…', async () => { const v = await DocTools.form('Dropdown', [{ name: 'values', label: 'Choices (one per line)', type: 'textarea', value: 'Option 1\nOption 2\nOption 3' }], 'Done'); if (v) { const values = v.values.split('\n').map((x) => x.trim()).filter(Boolean); if (values.length) await view.setValidation({ type: 'list', values }); } }),
        it('Note…', () => view.editNote(), { key: 'Shift+F2' }),
        SEP,
        it('New sheet', () => view.addSheet()),
      ],
      Format: [
        sub('Number', [it('Custom number format…', () => view.customNumberFormat())]),
        sub('Vertical align', [['top', 'Top'], ['center', 'Middle'], ['bottom', 'Bottom']].map(([v, t]) => it(t, () => view.styleSelection({ align: { vertical: v } }), { checked: () => view.activeCss().v === v }))),
        sub('Rotation', [[0, 'None'], [45, 'Tilt up'], [135, 'Tilt down'], [90, 'Rotate up'], [180, 'Rotate down'], [255, 'Stack vertically']].map(([v, t]) => it(t, () => view.styleSelection({ align: { textRotation: v ? String(v) : null } }), { checked: () => (view.activeCss().rotation || 0) === v }))),
        it('Alternating colours…', () => view.alternatingColors()),
        it('Convert to table', () => view.convertToTable()),
        it('Clear formatting', () => { const g = view.effectiveRange(); const items = []; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 300); c++) items.push({ r, c, patch: { s: 0 } }); view.apply(items); }, { key: 'Ctrl+\\' }),
      ],
      Data: [
        it('Sort range…', () => view.sortRangeDialog()),
        it('Create a filter', () => view.toggleFilter()),
        SEP,
        sub('Data cleanup', [it('Remove duplicates…', () => view.removeDuplicates()), it('Trim whitespace', () => view.trimWhitespace())]),
        it('Split text to columns…', () => view.splitToColumns()),
        it('Randomize range', () => view.randomizeRange()),
        it('Column stats', () => view.columnStats()),
        SEP,
        it('Named ranges…', () => view.namedRangesDialog()),
        it('Protect sheet', () => view.toggleProtect(), { checked: () => view.isProtected() }),
        sub('Data validation', [it('Dropdown…', async () => { const v = await DocTools.form('Dropdown', [{ name: 'values', label: 'Choices (one per line)', type: 'textarea', value: '' }], 'Done'); if (v) { const values = v.values.split('\n').map((x) => x.trim()).filter(Boolean); if (values.length) await view.setValidation({ type: 'list', values }); } }),
          it('Number between…', async () => { const v = await DocTools.form('Allow numbers between', [{ name: 'a', label: 'Minimum', type: 'number', value: '0' }, { name: 'b', label: 'Maximum', type: 'number', value: '100' }], 'Done'); if (v) await view.setValidation({ type: 'decimal', op: 'between', a: v.a, b: v.b }); }),
          it('Checkbox', () => view.setValidation({ checkbox: true })),
          it('Remove validation', () => view.setValidation(null))]),
      ],
      Tools: [it('Recalculate everything', () => view.recalcAll(), { key: 'F9' })],
    };
  };
})();
