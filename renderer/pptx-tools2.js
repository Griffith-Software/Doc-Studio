'use strict';
// More Google Slides features for the presentation editor (added to PptxUI.Editor):
//   Image options: crop, brightness, contrast, transparency, greyscale, reflection (drawn here and kept for PowerPoint)
//   Insert: video and audio (played in the editor and the slideshow), real charts with their data in an embedded
//           workbook (Edit chart data), diagrams (process, timeline, cycle, hierarchy) as grouped shapes
//   Comments on slides (PowerPoint comments), and a theme builder (colours, fonts, background on the master)

(() => {
  const E = PptxUI.Editor.prototype;
  const R = PptxEditor.Renderer.prototype;
  const { NS, REL, parseXml, serialize, kids, first, aEl, relative } = PptxEditor;
  const SHAPES = new Set(['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp', 'AlternateContent']);
  const imp = (doc, xml) => { const n = doc.importNode(parseXml(xml).documentElement, true); for (const a of [...n.attributes]) if (a.name.startsWith('xmlns')) n.removeAttributeNode(a); return n; };
  const NSD = `xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"`;
  const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
  const RT = {
    video: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
    audio: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
    media: 'http://schemas.microsoft.com/office/2007/relationships/media',
    chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
    package: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
    comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    authors: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors',
  };
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const hex = (c) => c.replace('#', '').toUpperCase();
  const cur = (ed) => ed.deck.slides[ed.current];
  const list = (ed, slide) => kids(ed.tree(slide)).filter((x) => SHAPES.has(x.localName));
  const freePath = (deck, pattern) => { let n = 1; const taken = (p) => !!deck.zip.file(p) || deck.newFiles.has(p) || (deck.docs.has(p) && deck.docs.get(p)); while (taken(pattern.replace('#', n))) n++; return pattern.replace('#', n); };
  const ctype = (ext) => ({ mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', wmv: 'video/x-ms-wmv', avi: 'video/avi', mp3: 'audio/mpeg', wav: 'audio/wav', m4a: 'audio/mp4', aac: 'audio/aac', ogg: 'audio/ogg', wma: 'audio/x-ms-wma' }[ext] || 'application/octet-stream');
  const AUDIO = new Set(['mp3', 'wav', 'm4a', 'aac', 'ogg', 'wma']);
  const selectedOf = (ed, kind) => {
    const s = ed.selectedNodes()[0];
    if (!s || (kind && s.node.localName !== kind)) return null;
    return s;
  };

  // ======================================================================
  // Image options (crop, adjustments, reflection)
  // ======================================================================

  E.imageOptionsOf = function () {
    const s = selectedOf(this, 'pic');
    if (!s) return null;
    const bf = first(s.node, 'blipFill'), blip = first(bf, 'blip');
    const src = first(bf, 'srcRect');
    const lum = first(blip, 'lum'), alpha = first(blip, 'alphaModFix');
    const pct = (el, a) => (el && el.hasAttribute(a) ? +el.getAttribute(a) / 1000 : 0);
    const eff = first(first(s.node, 'spPr'), 'effectLst');
    return {
      crop: { l: pct(src, 'l'), t: pct(src, 't'), r: pct(src, 'r'), b: pct(src, 'b') },
      bright: pct(lum, 'bright'), contrast: pct(lum, 'contrast'),
      transparency: alpha ? 100 - pct(alpha, 'amt') : 0,
      grey: !!first(blip, 'grayscl'),
      reflection: !!(eff && first(eff, 'reflection')),
    };
  };
  // o: { crop: {l,t,r,b} (percent), bright, contrast (-100..100), transparency (0..100), grey, reflection }
  E.setImageOptions = async function (o) {
    const s = selectedOf(this, 'pic');
    if (!s) throw new Error('Select a picture first.');
    this.snapshot();
    const slide = cur(this), doc = slide.doc;
    const bf = first(s.node, 'blipFill'), blip = first(bf, 'blip');
    if (o.crop) {
      for (const x of kids(bf, 'srcRect')) x.remove();
      const c = o.crop;
      if (c.l || c.t || c.r || c.b) {
        const el = aEl(doc, 'srcRect');
        for (const k of ['l', 't', 'r', 'b']) if (+c[k]) el.setAttribute(k, Math.round(Math.max(0, Math.min(90, +c[k])) * 1000));
        blip.after(el);
      }
    }
    for (const n of ['lum', 'alphaModFix', 'grayscl']) for (const x of kids(blip, n)) x.remove();
    const ext = first(blip, 'extLst');
    const add = (el) => blip.insertBefore(el, ext);
    if (+o.transparency) add(aEl(doc, 'alphaModFix', { amt: Math.round((100 - Math.max(0, Math.min(100, +o.transparency))) * 1000) }));
    if (o.grey) add(aEl(doc, 'grayscl'));
    if (+o.bright || +o.contrast) { const l = aEl(doc, 'lum'); if (+o.bright) l.setAttribute('bright', Math.round(+o.bright * 1000)); if (+o.contrast) l.setAttribute('contrast', Math.round(+o.contrast * 1000)); add(l); }
    const spPr = first(s.node, 'spPr');
    let eff = first(spPr, 'effectLst');
    if (eff) for (const x of kids(eff, 'reflection')) x.remove();
    if (o.reflection) {
      if (!eff) { eff = aEl(doc, 'effectLst'); spPr.insertBefore(eff, kids(spPr).find((c) => ['scene3d', 'sp3d', 'extLst'].includes(c.localName)) || null); }
      eff.insertBefore(aEl(doc, 'reflection', { blurRad: '6350', stA: '50000', endA: '300', endPos: '55000', dir: '5400000', sy: '-100000', algn: 'bl', rotWithShape: '0' }), kids(eff).find((c) => c.localName === 'softEdge') || null);
    }
    if (eff && !kids(eff).length) eff.remove();
    this.markSlide(slide);
    this.changed();
    await this.redrawKeep([s.idx]);
  };

  // Drawing: picture adjustments and reflection; video/audio players.
  const shape0 = R.shape;
  R.shape = async function (node, slide, rels, layer, group) {
    const el = await shape0.call(this, node, slide, rels, layer, group);
    if (!el || node.localName !== 'pic') return el;
    const blip = first(first(node, 'blipFill'), 'blip');
    const img = el.querySelector('img');
    if (img && blip) {
      const f = [];
      const lum = first(blip, 'lum');
      if (lum) { const b = +(lum.getAttribute('bright') || 0) / 100000, c = +(lum.getAttribute('contrast') || 0) / 100000; if (b) f.push(`brightness(${(1 + b).toFixed(3)})`); if (c) f.push(`contrast(${(1 + c).toFixed(3)})`); }
      if (first(blip, 'grayscl')) f.push('grayscale(1)');
      if (f.length) img.style.filter = f.join(' ');
      const a = first(blip, 'alphaModFix');
      if (a) img.style.opacity = +(a.getAttribute('amt') || 100000) / 100000;
    }
    const eff = first(first(node, 'spPr'), 'effectLst');
    if (eff && first(eff, 'reflection')) { el.style.webkitBoxReflect = 'below 2px linear-gradient(transparent 50%, rgba(255,255,255,.45))'; el.style.overflow = 'visible'; el.classList.add('pp-reflect'); }
    const nvPr = first(first(node, 'nvPicPr'), 'nvPr');
    const mediaEl = nvPr && (first(nvPr, 'videoFile') || first(nvPr, 'audioFile'));
    if (mediaEl) {
      const p14 = nvPr.getElementsByTagNameNS(P14, 'media')[0];
      const rid = (p14 && (p14.getAttributeNS(NS.r, 'embed') || p14.getAttribute('r:embed'))) || mediaEl.getAttributeNS(NS.r, 'link') || mediaEl.getAttribute('r:link');
      const rel = rels.find((r) => r.id === rid);
      if (rel && rel.path) {
        const bytes = await this.deck.bytesOf(rel.path);
        if (bytes) {
          const ext = rel.path.split('.').pop().toLowerCase();
          const url = URL.createObjectURL(new Blob([bytes], { type: ctype(ext) }));
          const isVideo = mediaEl.localName === 'videoFile';
          const m = document.createElement(isVideo ? 'video' : 'audio');
          m.src = url; m.preload = 'metadata'; m.className = isVideo ? 'pp-video' : 'pp-audio';
          if (img && isVideo) m.poster = img.src;
          if (this.editable) {
            // In the editor the shape still drags and resizes; a small button plays it.
            m.controls = false;
            const btn = document.createElement('button');
            btn.className = 'pp-play'; btn.textContent = '▶'; btn.title = 'Play / pause';
            btn.addEventListener('mousedown', (e) => e.stopPropagation());
            btn.addEventListener('click', (e) => { e.stopPropagation(); if (m.paused) { m.play(); btn.textContent = '❚❚'; } else { m.pause(); btn.textContent = '▶'; } });
            m.addEventListener('ended', () => { btn.textContent = '▶'; });
            el.append(m, btn);
          } else {
            m.controls = true;
            for (const t of ['mousedown', 'click', 'keydown']) m.addEventListener(t, (e) => e.stopPropagation());
            el.append(m);
          }
          if (isVideo && img) img.style.display = 'none';
          el.classList.add('pp-media');
        }
      }
    }
    return el;
  };

  // ======================================================================
  // Insert > Video / Audio
  // ======================================================================

  // A poster picture for the media shape: the video's first frame, or a speaker icon for audio.
  async function posterFor(bytes, ext, audio) {
    const canvas = document.createElement('canvas');
    let w = 640, h = 360;
    if (!audio) {
      try {
        const v = document.createElement('video');
        v.muted = true; v.src = URL.createObjectURL(new Blob([bytes], { type: ctype(ext) }));
        await new Promise((res, rej) => { v.onloadeddata = res; v.onerror = rej; setTimeout(rej, 6000); });
        v.currentTime = Math.min(0.5, (v.duration || 1) / 4);
        await new Promise((res) => { v.onseeked = res; setTimeout(res, 2000); });
        w = v.videoWidth || w; h = v.videoHeight || h;
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(v, 0, 0, w, h);
        URL.revokeObjectURL(v.src);
      } catch { canvas.width = w; canvas.height = h; const g = canvas.getContext('2d'); g.fillStyle = '#202124'; g.fillRect(0, 0, w, h); g.fillStyle = '#fff'; g.beginPath(); g.moveTo(w / 2 - 40, h / 2 - 50); g.lineTo(w / 2 + 50, h / 2); g.lineTo(w / 2 - 40, h / 2 + 50); g.fill(); }
    } else {
      w = h = 256; canvas.width = w; canvas.height = h;
      const g = canvas.getContext('2d');
      g.fillStyle = '#1a73e8'; g.beginPath(); g.arc(128, 128, 120, 0, Math.PI * 2); g.fill();
      g.fillStyle = '#fff'; g.beginPath(); g.moveTo(70, 105); g.lineTo(105, 105); g.lineTo(145, 70); g.lineTo(145, 186); g.lineTo(105, 151); g.lineTo(70, 151); g.closePath(); g.fill();
      g.strokeStyle = '#fff'; g.lineWidth = 10; g.lineCap = 'round';
      g.beginPath(); g.arc(150, 128, 30, -0.8, 0.8); g.stroke(); g.beginPath(); g.arc(150, 128, 55, -0.8, 0.8); g.stroke();
    }
    const png = await (await new Promise((r) => canvas.toBlob(r, 'image/png'))).arrayBuffer();
    return { png, w, h };
  }

  E.insertMedia = async function (bytes, fileName) {
    const ext = (fileName.split('.').pop() || '').toLowerCase();
    const audio = AUDIO.has(ext);
    const deck = this.deck, slide = cur(this);
    const poster = await posterFor(bytes, ext, audio);
    this.snapshot();
    const media = freePath(deck, `ppt/media/media#.${ext}`);
    deck.newFiles.set(media, new Uint8Array(bytes));
    const img = freePath(deck, 'ppt/media/image#.png');
    deck.newFiles.set(img, new Uint8Array(poster.png));
    await this.ensureDefault(ext, ctype(ext));
    await this.ensureDefault('png', 'image/png');
    const rLink = deck.addRel(slide.path, audio ? RT.audio : RT.video, relative(slide.path, media));
    const rMedia = deck.addRel(slide.path, RT.media, relative(slide.path, media));
    const rImg = deck.addRel(slide.path, REL.image, relative(slide.path, img));
    slide.rels = await deck.rels(slide.path);
    let w = audio ? 812800 : poster.w * 9525, h = audio ? 812800 : poster.h * 9525;
    const s = Math.min(1, (deck.cx * 0.6) / w, (deck.cy * 0.6) / h);
    w = Math.round(w * s); h = Math.round(h * s);
    const id = this.nextId(slide);
    const xml = `<p:pic ${NSD}><p:nvPicPr><p:cNvPr id="${id}" name="${esc(fileName)}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:${audio ? 'audioFile' : 'videoFile'} r:link="${rLink}"/><p:extLst><p:ext uri="{DAA4B4D9-6D2D-4E5D-B4F2-0A5F1E7F6A9F}"><p14:media xmlns:p14="${P14}" r:embed="${rMedia}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${rImg}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round((deck.cx - w) / 2)}" y="${Math.round((deck.cy - h) / 2)}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
    const node = slide.doc.importNode(parseXml(xml).documentElement, true);
    this.tree(slide).append(node);
    this.markSlide(slide);
    await this.afterInsert();
  };

  // ======================================================================
  // Insert > Chart (a real chart; its data lives in an embedded workbook)
  // ======================================================================

  const CHART_TYPES = [['column', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['pie', 'Pie'], ['area', 'Area']];
  function chartXml(type, title, rows) {
    const n = rows.length - 1;
    const col = (i) => String.fromCharCode(65 + i);
    const strCache = (vals) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
    const numCache = (vals) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
    const cats = rows.slice(1).map((r) => r[0]);
    const series = rows[0].slice(1).map((name, i) => ({ name, col: col(i + 1), vals: rows.slice(1).map((r) => +r[i + 1] || 0) }));
    const ser = (type === 'pie' ? series.slice(0, 1) : series).map((s, i) => {
      const tx = `<c:tx><c:strRef><c:f>Sheet1!$${s.col}$1</c:f>${strCache([s.name])}</c:strRef></c:tx>`;
      const cat = `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f>${strCache(cats)}</c:strRef></c:cat>`;
      const val = `<c:val><c:numRef><c:f>Sheet1!$${s.col}$2:$${s.col}$${n + 1}</c:f>${numCache(s.vals)}</c:numRef></c:val>`;
      const marker = type === 'line' ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : '';
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${type === 'bar' || type === 'column' ? '<c:invertIfNegative val="0"/>' : ''}${marker}${cat}${val}${type === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    }).join('');
    const axes = '<c:axId val="600000001"/><c:axId val="600000002"/>';
    let plot;
    if (type === 'pie') plot = `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
    else if (type === 'line') plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/>${axes}</c:lineChart>`;
    else if (type === 'area') plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}${axes}</c:areaChart>`;
    else plot = `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${ser}<c:gapWidth val="150"/>${axes}</c:barChart>`;
    const catAx = `<c:catAx><c:axId val="600000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${type === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="600000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const valAx = `<c:valAx><c:axId val="600000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${type === 'bar' ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="600000001"/><c:crosses val="autoZero"/><c:crossBetween val="${type === 'area' ? 'midCat' : 'between'}"/></c:valAx>`;
    const t = title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800" b="0"/></a:pPr><a:r><a:rPr lang="en-US" sz="1800" b="0"/><a:t>${esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:date1904 val="0"/><c:roundedCorners val="0"/><c:chart>${t}<c:plotArea><c:layout/>${plot}${type === 'pie' ? '' : catAx + valAx}</c:plotArea><c:legend><c:legendPos val="${type === 'pie' ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
  }
  function workbookBytes(rows) {
    const aoa = rows.map((r, i) => r.map((x, j) => (i > 0 && j > 0 ? +x || 0 : x)));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Sheet1');
    return new Uint8Array(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  }
  const parseRows = (text) => {
    const rows = text.split('\n').map((l) => l.split(/\t|,/).map((x) => x.trim())).filter((r) => r.some(Boolean));
    if (rows.length < 2 || rows[0].length < 2) throw new Error('Enter a header row (labels, then series names) and at least one row of numbers.');
    const w = rows[0].length;
    return rows.map((r) => Array.from({ length: w }, (_, i) => r[i] ?? ''));
  };
  // spec: { type, title, rows: [[label, series...], [cat, n...]...] }
  E.insertRealChart = async function (spec) {
    const deck = this.deck, slide = cur(this);
    this.snapshot();
    const chartPath = freePath(deck, 'ppt/charts/chart#.xml');
    const xlsxPath = freePath(deck, 'ppt/embeddings/Microsoft_Excel_Worksheet#.xlsx');
    deck.docs.set(chartPath, parseXml(chartXml(spec.type, spec.title, spec.rows)));
    deck.dirty.add(chartPath);
    deck.newFiles.set(xlsxPath, workbookBytes(spec.rows));
    deck.addRel(chartPath, RT.package, relative(chartPath, xlsxPath));
    await this.addOverride(chartPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
    await this.ensureDefault('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    const rid = deck.addRel(slide.path, RT.chart, relative(slide.path, chartPath));
    slide.rels = await deck.rels(slide.path);
    const id = this.nextId(slide);
    const cx = deck.cx, cy = deck.cy;
    const xml = `<p:graphicFrame ${NSD}><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${Math.round(cx * 0.15)}" y="${Math.round(cy * 0.2)}"/><a:ext cx="${Math.round(cx * 0.7)}" cy="${Math.round(cy * 0.65)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="${NS.c}" r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`;
    this.tree(slide).append(imp(slide.doc, xml));
    this.markSlide(slide);
    await this.afterInsert();
  };
  // The chart part behind the selected chart, with its data read back from the chart's caches.
  E.chartAt = async function () {
    const s = selectedOf(this, 'graphicFrame');
    const c = s && s.node.getElementsByTagNameNS(NS.c, 'chart')[0];
    if (!c) return null;
    const slide = cur(this);
    const rid = c.getAttributeNS(NS.r, 'id') || c.getAttribute('r:id');
    const rel = (await this.deck.rels(slide.path)).find((r) => r.id === rid);
    if (!rel || !rel.path) return null;
    const doc = await this.deck.doc(rel.path);
    const sers = [...doc.getElementsByTagNameNS(NS.c, 'ser')];
    const pts = (el) => (el ? [...el.getElementsByTagNameNS(NS.c, 'pt')].map((p) => (first(p, 'v') || {}).textContent || '') : []);
    const plot = [...doc.getElementsByTagNameNS(NS.c, 'plotArea')[0].children].find((x) => /Chart$/.test(x.localName));
    const type = !plot ? 'column' : plot.localName === 'barChart' ? (first(plot, 'barDir') && first(plot, 'barDir').getAttribute('val') === 'bar' ? 'bar' : 'column') : plot.localName.replace('Chart', '').replace('pie3D', 'pie').replace('doughnut', 'pie');
    const cats = pts(sers[0] && first(sers[0], 'cat'));
    const rows = [['', ...sers.map((x) => pts(first(x, 'tx'))[0] || 'Series')]];
    cats.forEach((cat, i) => rows.push([cat, ...sers.map((x) => pts(first(x, 'val'))[i] || '')]));
    const t = doc.getElementsByTagNameNS(NS.c, 'title')[0];
    const title = t ? [...t.getElementsByTagNameNS(NS.a, 't')].map((x) => x.textContent).join('') : '';
    return { path: rel.path, doc, rows, type: CHART_TYPES.some(([k]) => k === type) ? type : 'column', title, idx: s.idx };
  };
  E.editChartData = async function (spec) {
    const info = await this.chartAt();
    if (!info) throw new Error('Select a chart first.');
    const deck = this.deck;
    this.snapshot();
    deck.docs.set(info.path, parseXml(chartXml(spec.type, spec.title, spec.rows)));
    deck.dirty.add(info.path);
    // Keep the chart's own relationships (its workbook) and refresh the workbook.
    const rels = await deck.rels(info.path);
    let pk = rels.find((r) => r.type === RT.package);
    if (!pk) { const x = freePath(deck, 'ppt/embeddings/Microsoft_Excel_Worksheet#.xlsx'); deck.addRel(info.path, RT.package, relative(info.path, x)); await this.ensureDefault('xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); pk = (await deck.rels(info.path)).find((r) => r.type === RT.package); }
    const d = deck.docs.get(info.path);
    const ext = d.getElementsByTagNameNS(NS.c, 'externalData')[0];
    if (ext) ext.setAttributeNS(NS.r, 'r:id', pk.id);
    deck.newFiles.set(pk.path, workbookBytes(spec.rows));
    this.changed();
    await this.redrawKeep([info.idx]);
  };

  // ======================================================================
  // Insert > Diagram (process, timeline, cycle, hierarchy) as one group of shapes
  // ======================================================================

  function spXml(id, name, x, y, w, h, prst, fill, text, o) {
    o = o || {};
    const rot = o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : '';
    const fillXml = fill == null ? '<a:noFill/>' : /^(accent\d|tx\d|bg\d|lt\d|dk\d)$/.test(fill) ? `<a:solidFill><a:schemeClr val="${fill}"/></a:solidFill>` : `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`;
    const ln = o.line ? `<a:ln w="${o.lineW || 19050}"><a:solidFill><a:schemeClr val="${o.line}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>';
    const color = o.textColor || (fill == null ? null : 'FFFFFF');
    const tx = text != null ? `<p:txBody><a:bodyPr wrap="square" lIns="45720" tIns="45720" rIns="45720" bIns="45720" anchor="${o.anchor || 'ctr'}"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${o.align || 'ctr'}"/><a:r><a:rPr lang="en-US" sz="${o.size || 1400}"${o.bold === false ? '' : ' b="1"'} dirty="0">${color ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` : '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'}</a:rPr><a:t>${esc(text)}</a:t></a:r></a:p></p:txBody>` : '';
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)} ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${rot}><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(Math.max(1, w))}" cy="${Math.round(Math.max(1, h))}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${fillXml}${ln}</p:spPr>${tx}</p:sp>`;
  }
  function lineXml(id, x, y, w, h, color) {
    return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Connector ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:schemeClr val="${color || 'tx2'}"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`;
  }
  const accent = (i) => `accent${(i % 6) + 1}`;
  // kind: 'process' | 'timeline' | 'cycle' | 'hierarchy'; lines: the items (hierarchy: indent children with spaces or "-")
  E.insertDiagram = async function (kind, lines) {
    const items = lines.map((l) => ({ text: l.replace(/^[\s\-•*]+/, '').trim(), level: /^(\s{2,}|\t|[-•*])/.test(l) ? 1 : 0 })).filter((x) => x.text);
    if (!items.length) throw new Error('Type at least one item.');
    const deck = this.deck, slide = cur(this);
    this.snapshot();
    let id = this.nextId(slide);
    const gid = id++;
    const W = deck.cx * 0.84, H = deck.cy * 0.6, X = deck.cx * 0.08, Y = deck.cy * 0.25;
    const parts = [];
    const n = items.length;
    if (kind === 'process') {
      const gap = W * 0.01, w = (W - gap * (n - 1)) / n, h = Math.min(H * 0.45, w * 0.55);
      items.forEach((it, i) => parts.push(spXml(id++, 'Step', X + i * (w + gap), Y + (H - h) / 2, w, h, i === 0 ? 'homePlate' : 'chevron', accent(i), it.text)));
    } else if (kind === 'timeline') {
      const midY = Y + H / 2;
      parts.push(spXml(id++, 'Timeline', X, midY - 25400, W, 50800, 'rect', 'tx2', null));
      const step = W / n, dot = Math.min(step * 0.25, 457200), boxW = step * 0.92, boxH = H * 0.3;
      items.forEach((it, i) => {
        const cxp = X + step * (i + 0.5);
        parts.push(spXml(id++, 'Milestone', cxp - dot / 2, midY - dot / 2, dot, dot, 'ellipse', accent(i), null, { line: 'bg1', lineW: 25400 }));
        const up = i % 2 === 0;
        parts.push(spXml(id++, 'Label', cxp - boxW / 2, up ? midY - dot - boxH : midY + dot, boxW, boxH, 'rect', null, it.text, { anchor: up ? 'b' : 't', size: 1400, textColor: null }));
      });
    } else if (kind === 'cycle') {
      const r = Math.min(W, H) * 0.36, c = { x: X + W / 2, y: Y + H / 2 }, node = Math.min(r * 0.9, (2 * Math.PI * r / n) * 0.6);
      items.forEach((it, i) => {
        const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
        parts.push(spXml(id++, 'Stage', c.x + r * Math.cos(a) - node / 2, c.y + r * Math.sin(a) - node / 2, node, node, 'ellipse', accent(i), it.text, { size: 1200 }));
        const b = a + Math.PI / n;
        const aw = node * 0.4, ah = node * 0.25;
        parts.push(spXml(id++, 'Arrow', c.x + r * Math.cos(b) - aw / 2, c.y + r * Math.sin(b) - ah / 2, aw, ah, 'rightArrow', 'tx2', null, { rot: ((b * 180) / Math.PI + 90 + 360) % 360 }));
      });
    } else {
      // Hierarchy: the first top-level item is the root; other top-level items are its children; indented items
      // hang under the child above them.
      const root = items[0];
      const kidsL = [];
      for (const it of items.slice(1)) { if (it.level === 0 || !kidsL.length) kidsL.push({ text: it.text, subs: [] }); else kidsL[kidsL.length - 1].subs.push(it.text); }
      const m = Math.max(1, kidsL.length), maxSubs = Math.max(0, ...kidsL.map((k) => k.subs.length));
      const rows = 2 + maxSubs;
      const boxH = Math.min(H / (rows + (rows - 1) * 0.5), 640080), vgap = boxH * 0.5;
      const colW = W / m, boxW = Math.min(colW * 0.86, 2286000);
      const rootW = Math.min(W * 0.3, 2743200);
      parts.push(spXml(id++, 'Root', X + W / 2 - rootW / 2, Y, rootW, boxH, 'roundRect', 'accent1', root.text));
      if (kidsL.length) {
        const y1 = Y + boxH, ym = y1 + vgap / 2, y2 = y1 + vgap;
        parts.push(lineXml(id++, X + W / 2, y1, 0, vgap / 2));
        const xs = kidsL.map((_, i) => X + colW * (i + 0.5));
        if (kidsL.length > 1) parts.push(lineXml(id++, xs[0], ym, xs[xs.length - 1] - xs[0], 0));
        kidsL.forEach((k, i) => {
          parts.push(lineXml(id++, xs[i], ym, 0, vgap / 2));
          parts.push(spXml(id++, 'Item', xs[i] - boxW / 2, y2, boxW, boxH, 'roundRect', 'accent2', k.text, { size: 1300 }));
          k.subs.forEach((t, j) => {
            const sy = y2 + (boxH + vgap) * (j + 1);
            parts.push(lineXml(id++, xs[i] - boxW / 2 + boxW * 0.1, sy - vgap, 0, vgap + boxH / 2));
            parts.push(lineXml(id++, xs[i] - boxW / 2 + boxW * 0.1, sy + boxH / 2, boxW * 0.1, 0));
            parts.push(spXml(id++, 'Sub-item', xs[i] - boxW / 2 + boxW * 0.2, sy, boxW * 0.8, boxH, 'roundRect', 'lt2', t, { size: 1100, bold: false, textColor: '262626' }));
          });
        });
      }
    }
    const xml = `<p:grpSp ${NSD}><p:nvGrpSpPr><p:cNvPr id="${gid}" name="${kind[0].toUpperCase()}${kind.slice(1)} diagram ${gid}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="${Math.round(X)}" y="${Math.round(Y)}"/><a:ext cx="${Math.round(W)}" cy="${Math.round(H)}"/><a:chOff x="${Math.round(X)}" y="${Math.round(Y)}"/><a:chExt cx="${Math.round(W)}" cy="${Math.round(H)}"/></a:xfrm></p:grpSpPr>${parts.join('')}</p:grpSp>`;
    this.tree(slide).append(imp(slide.doc, xml));
    this.markSlide(slide);
    await this.afterInsert();
  };

  // ======================================================================
  // Comments (PowerPoint comments: ppt/comments/commentN.xml and ppt/commentAuthors.xml)
  // ======================================================================

  const EMU_PER_POS = 914400 / 576; // comment positions are in 1/576 inch
  const authorName = () => { try { return localStorage.getItem('docstudio.author') || 'Doc Studio'; } catch { return 'Doc Studio'; } };
  async function authorsDoc(ed, create) {
    const deck = ed.deck;
    const rel = (await deck.rels(deck.presPath)).find((r) => r.type === RT.authors);
    if (rel && rel.path) return { path: rel.path, doc: await deck.doc(rel.path) };
    if (!create) return null;
    const path = 'ppt/commentAuthors.xml';
    const doc = parseXml(`<p:cmAuthorLst xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"/>`);
    deck.docs.set(path, doc); deck.dirty.add(path);
    deck.addRel(deck.presPath, RT.authors, relative(deck.presPath, path));
    await ed.addOverride(path, 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml');
    return { path, doc };
  }
  async function commentsDoc(ed, slide, create) {
    const deck = ed.deck;
    const rel = (await deck.rels(slide.path)).find((r) => r.type === RT.comments);
    if (rel && rel.path) return { path: rel.path, doc: await deck.doc(rel.path) };
    if (!create) return null;
    const path = freePath(deck, 'ppt/comments/comment#.xml');
    const doc = parseXml(`<p:cmLst xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}"/>`);
    deck.docs.set(path, doc); deck.dirty.add(path);
    deck.addRel(slide.path, RT.comments, relative(slide.path, path));
    slide.rels = await deck.rels(slide.path);
    await ed.addOverride(path, 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml');
    return { path, doc };
  }
  E.slideComments = async function (i) {
    const slide = this.deck.slides[i ?? this.current];
    const c = slide && await commentsDoc(this, slide, false);
    if (!c || !c.doc) return [];
    const a = await authorsDoc(this, false);
    const names = new Map(a && a.doc ? kids(a.doc.documentElement, 'cmAuthor').map((x) => [x.getAttribute('id'), x.getAttribute('name')]) : []);
    return kids(c.doc.documentElement, 'cm').map((cm) => {
      const pos = first(cm, 'pos');
      return { idx: cm.getAttribute('idx'), authorId: cm.getAttribute('authorId'), author: names.get(cm.getAttribute('authorId')) || 'Unknown', date: cm.getAttribute('dt'), text: (first(cm, 'text') || {}).textContent || '', x: +(pos && pos.getAttribute('x')) || 0, y: +(pos && pos.getAttribute('y')) || 0 };
    });
  };
  // at: { x, y } in EMU (defaults to the selected shape's corner, or the top-left of the slide)
  E.addComment = async function (text, at) {
    if (!text || !text.trim()) return;
    const slide = cur(this);
    this.snapshot();
    const a = await authorsDoc(this, true);
    const name = authorName();
    let au = kids(a.doc.documentElement, 'cmAuthor').find((x) => x.getAttribute('name') === name);
    if (!au) {
      const ids = kids(a.doc.documentElement, 'cmAuthor').map((x) => +x.getAttribute('id'));
      au = a.doc.createElementNS(NS.p, 'p:cmAuthor');
      const initials = name.split(/\s+/).map((w) => w[0]).join('').slice(0, 3).toUpperCase();
      au.setAttribute('id', ids.length ? Math.max(...ids) + 1 : 0); au.setAttribute('name', name); au.setAttribute('initials', initials); au.setAttribute('lastIdx', '0'); au.setAttribute('clrIdx', String(ids.length % 8));
      a.doc.documentElement.append(au);
    }
    const idx = +au.getAttribute('lastIdx') + 1;
    au.setAttribute('lastIdx', idx);
    a.doc && this.deck.dirty.add(a.path);
    const c = await commentsDoc(this, slide, true);
    if (!at) { const s = this.selectedNodes()[0]; const b = s && this.boxOf(s.el); at = b ? { x: b.x + b.w, y: b.y } : { x: 0, y: 0 }; }
    const cm = c.doc.createElementNS(NS.p, 'p:cm');
    cm.setAttribute('authorId', au.getAttribute('id'));
    cm.setAttribute('dt', new Date().toISOString().replace('Z', ''));
    cm.setAttribute('idx', idx);
    const pos = c.doc.createElementNS(NS.p, 'p:pos');
    pos.setAttribute('x', Math.max(0, Math.round(at.x / EMU_PER_POS))); pos.setAttribute('y', Math.max(0, Math.round(at.y / EMU_PER_POS)));
    const t = c.doc.createElementNS(NS.p, 'p:text');
    t.textContent = text.trim();
    cm.append(pos, t);
    c.doc.documentElement.append(cm);
    this.deck.dirty.add(c.path);
    this.changed();
    await this.renderCurrent();
  };
  E.deleteComment = async function (authorId, idx) {
    const slide = cur(this);
    const c = await commentsDoc(this, slide, false);
    if (!c) return;
    this.snapshot();
    for (const cm of kids(c.doc.documentElement, 'cm')) if (cm.getAttribute('authorId') === String(authorId) && cm.getAttribute('idx') === String(idx)) cm.remove();
    this.deck.dirty.add(c.path);
    this.changed();
    await this.renderCurrent();
  };
  // Comment pins on the slide in the editor.
  const renderCurrent0 = E.renderCurrent;
  E.renderCurrent = async function () {
    await renderCurrent0.call(this);
    if (!this.slideEl || this.hideComments) return;
    const list2 = await this.slideComments().catch(() => []);
    for (const c of list2) {
      const pin = h('div', { class: 'pp-comment-pin', title: `${c.author}: ${c.text}`, text: '💬' });
      pin.style.left = `${c.x * EMU_PER_POS * this.k}px`; pin.style.top = `${c.y * EMU_PER_POS * this.k}px`;
      pin.addEventListener('mousedown', (e) => e.stopPropagation());
      pin.addEventListener('click', (e) => { e.stopPropagation(); PptxUI.commentsPanel(this); });
      this.slideEl.append(pin);
    }
  };
  PptxUI.commentsPanel = async function (ed) {
    const run = (label, fn) => async () => { try { await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } };
    const comments = await ed.slideComments();
    const box = h('div', { class: 'dialog versions-dialog' }, h('h3', { text: `Comments on slide ${ed.current + 1}` }));
    const back = h('div', { class: 'backdrop' }, box);
    const close = () => back.remove();
    const rows = h('div', { class: 'version-list' });
    if (!comments.length) rows.append(h('p', { class: 'muted', text: 'No comments on this slide yet.' }));
    for (const c of comments) {
      rows.append(h('div', { class: 'comment-row' },
        h('div', {}, h('b', { text: c.author }), h('span', { class: 'muted', text: c.date ? ` · ${new Date(c.date).toLocaleString()}` : '' })),
        h('div', { class: 'comment-text', text: c.text }),
        h('button', { class: 'ghost', text: 'Resolve', title: 'Delete this comment', onclick: run('delete the comment', async () => { close(); await ed.deleteComment(c.authorId, c.idx); await PptxUI.commentsPanel(ed); }) })));
    }
    const input = h('textarea', { rows: 3, placeholder: 'Add a comment (pinned next to the selected shape)' });
    box.append(rows, input, h('div', { class: 'dialog-actions' },
      h('button', { text: 'Comment', onclick: run('add the comment', async () => { const t = input.value; close(); await ed.addComment(t); await PptxUI.commentsPanel(ed); }) }),
      h('button', { class: 'ghost', text: 'Close', onclick: close })));
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
    back.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    document.body.append(back);
    input.focus();
  };

  // ======================================================================
  // Theme builder: custom colours and fonts, and the master's background and text sizes
  // ======================================================================

  const SLOTS = ['dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
  E.themeNow = async function () {
    const slide = cur(this);
    const doc = slide.themePath && await this.deck.doc(slide.themePath);
    if (!doc) return null;
    const scheme = doc.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
    const colors = {};
    for (const n of SLOTS) { const el = first(scheme, n); const c = el && (first(el, 'srgbClr') || first(el, 'sysClr')); colors[n] = c ? `#${(c.getAttribute('val') === 'windowText' || c.getAttribute('val') === 'window' ? c.getAttribute('lastClr') : c.getAttribute('val')) || '000000'}` : '#000000'; }
    const fs = doc.getElementsByTagNameNS(NS.a, 'fontScheme')[0];
    const face = (n) => { const l = first(first(fs, n), 'latin'); return l ? l.getAttribute('typeface') : ''; };
    return { colors, heading: face('majorFont'), body: face('minorFont'), name: scheme.getAttribute('name') || '' };
  };
  // t: { colors: {slot: '#rrggbb'}, heading, body, background ('#rrggbb' or null), titleSize, bodySize (pt) }
  E.buildTheme = async function (t) {
    this.snapshot();
    const deck = this.deck;
    const themes = [...new Set(deck.slides.map((s) => s.themePath).filter(Boolean))];
    for (const p of themes) {
      const doc = await deck.doc(p);
      const scheme = doc.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
      for (const [slot, c] of Object.entries(t.colors || {})) {
        const el = first(scheme, slot); if (!el || !/^#[0-9a-f]{6}$/i.test(c)) continue;
        while (el.firstChild) el.firstChild.remove();
        el.append(aEl(doc, 'srgbClr', { val: hex(c) }));
      }
      scheme.setAttribute('name', t.name || 'Doc Studio Custom');
      const fs = doc.getElementsByTagNameNS(NS.a, 'fontScheme')[0];
      if (fs) {
        const l1 = first(first(fs, 'majorFont'), 'latin'), l2 = first(first(fs, 'minorFont'), 'latin');
        if (l1 && t.heading) l1.setAttribute('typeface', t.heading);
        if (l2 && t.body) l2.setAttribute('typeface', t.body);
      }
      deck.dirty.add(p);
    }
    const masters = [...new Set(deck.slides.map((s) => s.masterPath).filter(Boolean))];
    for (const p of masters) {
      const doc = await deck.doc(p);
      const cSld = doc.getElementsByTagNameNS(NS.p, 'cSld')[0];
      if (t.background !== undefined) {
        for (const b of kids(cSld, 'bg')) b.remove();
        const bg = t.background ? `<p:bg ${NSD}><p:bgPr><a:solidFill><a:srgbClr val="${hex(t.background)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : `<p:bg ${NSD}><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>`;
        cSld.insertBefore(imp(doc, bg), cSld.firstChild);
      }
      const tx = doc.getElementsByTagNameNS(NS.p, 'txStyles')[0];
      const size = (styleName, pt) => {
        if (!tx || !pt) return;
        const st = first(tx, styleName);
        const l1 = st && first(st, 'lvl1pPr');
        const d = l1 && first(l1, 'defRPr');
        if (d) d.setAttribute('sz', Math.round(pt * 100));
      };
      size('titleStyle', +t.titleSize);
      size('bodyStyle', +t.bodySize);
      deck.dirty.add(p);
    }
    this.changed();
    await deck.loadSlides();
    await this.renderAll();
  };
  async function themeBuilder(ed) {
    const now = (await ed.themeNow()) || { colors: {}, heading: 'Calibri Light', body: 'Calibri' };
    const names = { dk2: 'Dark 2 (text)', lt2: 'Light 2', accent1: 'Accent 1', accent2: 'Accent 2', accent3: 'Accent 3', accent4: 'Accent 4', accent5: 'Accent 5', accent6: 'Accent 6' };
    const fonts = ['Calibri', 'Calibri Light', 'Arial', 'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold', 'Georgia', 'Times New Roman', 'Garamond', 'Trebuchet MS', 'Verdana', 'Century Gothic', 'Tahoma', 'Cambria', 'Consolas'];
    const fontOpts = (curF) => [...new Set([curF, ...fonts].filter(Boolean))].map((f) => [f, f]);
    const v = await DocTools.form('Theme builder', [
      { name: 'name', label: 'Theme name', value: now.name && !/^Office/.test(now.name) ? now.name : 'My theme' },
      ...SLOTS.map((s) => ({ name: s, label: names[s], type: 'color', value: (now.colors[s] || '#4472c4').toLowerCase() })),
      { name: 'heading', label: 'Heading font', type: 'select', value: now.heading, options: fontOpts(now.heading) },
      { name: 'body', label: 'Body font', type: 'select', value: now.body, options: fontOpts(now.body) },
      { name: 'useBg', label: 'Set a background colour on the master', type: 'checkbox', value: false },
      { name: 'bg', label: 'Master background', type: 'color', value: '#ffffff' },
      { name: 'titleSize', label: 'Title text size (pt, blank = keep)', type: 'number', value: '' },
      { name: 'bodySize', label: 'Body text size (pt, blank = keep)', type: 'number', value: '' },
    ], 'Apply theme');
    if (!v) return;
    await ed.buildTheme({ name: v.name, colors: Object.fromEntries(SLOTS.map((s) => [s, v[s]])), heading: v.heading, body: v.body, background: v.useBg ? v.bg : undefined, titleSize: v.titleSize, bodySize: v.bodySize });
  }

  // ======================================================================
  // Menus
  // ======================================================================

  const pickFile = (accept) => new Promise((resolve) => {
    const input = h('input', { type: 'file', accept });
    input.addEventListener('change', async () => { const f = input.files[0]; resolve(f ? { name: f.name, bytes: await f.arrayBuffer() } : null); });
    input.click();
  });
  async function chartDialog(title, init) {
    const v = await DocTools.form(title, [
      { name: 'type', label: 'Chart type', type: 'select', value: init.type, options: CHART_TYPES },
      { name: 'title', label: 'Title', value: init.title || '' },
      { name: 'data', label: 'Data: first row = series names, first column = labels (comma or tab separated)', type: 'textarea', value: init.rows.map((r) => r.join(', ')).join('\n') },
    ], title.startsWith('Edit') ? 'Update' : 'Insert');
    return v && { type: v.type, title: v.title.trim(), rows: parseRows(v.data) };
  }
  const menus0 = PptxUI.menus;
  PptxUI.menus = function (tab, ed) {
    const m = menus0(tab, ed);
    const run = (label, fn) => async () => { try { await fn(); } catch (err) { logError(err); desk.error(`Couldn't ${label}`, err.message); } };
    const it = (label, fn, o) => ({ label, run: run(label.toLowerCase(), fn), ...(o || {}) });
    const sub = (label, items) => ({ label, sub: items });
    const SEP = { sep: true };
    const editing = { enabled: () => tab.mode === 'edit' && !!tab.pptEditor };
    const add = (name, items) => { m[name] = [...(m[name] || []), SEP, ...items]; };
    const imgOptions = async (cropOnly) => {
      const o = ed().imageOptionsOf();
      if (!o) throw new Error('Select a picture first.');
      const crop = [{ name: 'l', label: 'Crop left (%)', type: 'number', value: String(o.crop.l) }, { name: 't', label: 'Crop top (%)', type: 'number', value: String(o.crop.t) }, { name: 'r', label: 'Crop right (%)', type: 'number', value: String(o.crop.r) }, { name: 'b', label: 'Crop bottom (%)', type: 'number', value: String(o.crop.b) }];
      const adj = [{ name: 'bright', label: 'Brightness (-100 to 100)', type: 'number', value: String(o.bright) }, { name: 'contrast', label: 'Contrast (-100 to 100)', type: 'number', value: String(o.contrast) }, { name: 'transparency', label: 'Transparency (0 to 100)', type: 'number', value: String(o.transparency) }, { name: 'grey', label: 'Greyscale', type: 'checkbox', value: o.grey }, { name: 'reflection', label: 'Reflection', type: 'checkbox', value: o.reflection }];
      const v = await DocTools.form(cropOnly ? 'Crop image' : 'Image options', cropOnly ? crop : [...crop, ...adj], 'Apply');
      if (!v) return;
      await ed().setImageOptions(cropOnly ? { ...o, crop: { l: +v.l || 0, t: +v.t || 0, r: +v.r || 0, b: +v.b || 0 } } : { crop: { l: +v.l || 0, t: +v.t || 0, r: +v.r || 0, b: +v.b || 0 }, bright: +v.bright || 0, contrast: +v.contrast || 0, transparency: +v.transparency || 0, grey: !!v.grey, reflection: !!v.reflection });
    };
    const diagram = (kind, label, sample) => it(label, async () => {
      const v = await DocTools.form(`${label} diagram`, [{ name: 'items', label: kind === 'hierarchy' ? 'Items, one per line (first = top; indent with "-" to put an item under the one above)' : 'Items, one per line', type: 'textarea', value: sample }], 'Insert');
      if (v) await ed().insertDiagram(kind, v.items.split('\n'));
    }, editing);
    add('Insert', [
      it('Video…', async () => { const f = await pickFile('video/mp4,video/webm,video/quicktime,.mp4,.m4v,.mov,.webm,.wmv,.avi'); if (f) await ed().insertMedia(f.bytes, f.name); }, editing),
      it('Audio…', async () => { const f = await pickFile('audio/*,.mp3,.wav,.m4a,.aac,.ogg,.wma'); if (f) await ed().insertMedia(f.bytes, f.name); }, editing),
      it('Chart…', async () => { const s = await chartDialog('Insert chart', { type: 'column', title: '', rows: [['Quarter', 'Sales', 'Costs'], ['Q1', 120, 80], ['Q2', 140, 90], ['Q3', 160, 95], ['Q4', 180, 100]] }); if (s) await ed().insertRealChart(s); }, editing),
      sub('Diagram', [
        diagram('process', 'Process', 'Plan\nBuild\nTest\nLaunch'),
        diagram('timeline', 'Timeline', 'Jan: Kick-off\nMar: Beta\nJun: Launch\nSep: Review'),
        diagram('cycle', 'Cycle', 'Plan\nDo\nCheck\nAct'),
        diagram('hierarchy', 'Hierarchy', 'Director\nDesign\n- UX\n- Visual\nEngineering\n- Web\n- Apps\nSales'),
      ]),
      it('Drawing…', async () => { const p = await DocsExtra.drawing('Drawing'); if (p) await ed().insertPicture(p.bytes, 'png', p.width, p.height); }, editing),
      it('Comment…', () => PptxUI.commentsPanel(ed()), { ...editing, key: 'Ctrl+Alt+M' }),
    ]);
    add('Format', [
      it('Crop image…', () => imgOptions(true), editing),
      it('Image options…', () => imgOptions(false), editing),
      it('Reset image', async () => { await ed().setImageOptions({ crop: { l: 0, t: 0, r: 0, b: 0 }, bright: 0, contrast: 0, transparency: 0, grey: false, reflection: false }); }, editing),
      it('Edit chart data…', async () => { const info = await ed().chartAt(); if (!info) throw new Error('Select a chart first.'); const s = await chartDialog('Edit chart data', info); if (s) await ed().editChartData(s); }, editing),
    ]);
    add('Slide', [it('Theme builder…', () => themeBuilder(ed()), editing)]);
    m.Tools = [...(m.Tools || []), it('Comments on this slide…', () => PptxUI.commentsPanel(ed()), editing)];
    return m;
  };

  document.addEventListener('keydown', (e) => {
    if (!active || active.kind !== 'pptx' || !active.pptEditor || active.mode !== 'edit' || document.querySelector('.ppe-show')) return;
    if (e.ctrlKey && e.altKey && e.key.toLowerCase() === 'm') { e.preventDefault(); PptxUI.commentsPanel(active.pptEditor); }
  }, true);
})();
