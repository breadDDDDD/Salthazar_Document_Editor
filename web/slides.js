/* DocStudio - PowerPoint-like editor for .pptx presentations.
 * Plain ES2020, no dependencies. Registers the "slides" editor with window.App.
 * Model: see board decisions "PPTX MODEL JSON" and "PPTX MODEL ADDITIONS" (internal/pptx).
 * Positions are EMU (914400/inch); the slide is laid out in CSS px at 96 dpi (9525 EMU/px)
 * and scaled with a CSS transform, so font sizes can be given in pt directly. */
(function () {
  'use strict';

  const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  const EMU_PX = 9525;
  const EMU_PT = 12700;
  const EMU_IN = 914400;
  const FONTS = ['Calibri', 'Calibri Light', 'Arial', 'Arial Black', 'Aptos', 'Cambria', 'Candara', 'Century Gothic',
    'Consolas', 'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Gill Sans MT', 'Impact',
    'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
  const ZOOMS = [25, 33, 50, 66, 75, 100, 125, 150, 200, 300, 400];
  const LINE_WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
  const DEFAULT_INSET = [91440, 45720, 91440, 45720];
  const DEFAULT_THEME = {
    majorFont: 'Calibri Light', minorFont: 'Calibri',
    colors: { dk1: '#000000', lt1: '#FFFFFF', dk2: '#44546A', lt2: '#E7E6E6', accent1: '#4472C4', accent2: '#ED7D31',
      accent3: '#A5A5A5', accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47', hlink: '#0563C1', folHlink: '#954F72' }
  };
  const STD_COLORS = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
  const SHAPES = [
    ['rect', 'Rectangle'], ['roundRect', 'Rounded rectangle'], ['ellipse', 'Oval'], ['triangle', 'Isosceles triangle'],
    ['rtTriangle', 'Right triangle'], ['diamond', 'Diamond'], ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'],
    ['octagon', 'Octagon'], ['parallelogram', 'Parallelogram'], ['trapezoid', 'Trapezoid'], ['star5', 'Star: 5 points'],
    ['plus', 'Cross'], ['rightArrow', 'Arrow: right'], ['leftArrow', 'Arrow: left'], ['upArrow', 'Arrow: up'],
    ['downArrow', 'Arrow: down'], ['chevron', 'Chevron'], ['homePlate', 'Pentagon arrow'], ['line', 'Line'],
    ['straightConnector1', 'Line arrow']
  ];
  const LINE_GEOMS = new Set(['line', 'straightConnector1', 'bentConnector2', 'bentConnector3', 'curvedConnector3']);
  const ALIGN_CSS = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', justLow: 'justify' };
  const ANCHOR_CSS = { t: 'flex-start', ctr: 'center', b: 'flex-end', just: 'flex-start', dist: 'flex-start' };
  const PH_HINT = { title: 'Click to add title', ctrTitle: 'Click to add title', subTitle: 'Click to add subtitle',
    body: 'Click to add text', obj: 'Click to add text', dt: '', ftr: '', sldNum: '' };
  const TITLE_PH = new Set(['title', 'ctrTitle']);
  let editorCounter = 0;
  let idSeq = 0;
  let clipboard = null; // {shapes:[...], text:string} shared by all slide editors

  /* ---------- small helpers ---------- */
  function h(tag, props, ...children) {
    const el = document.createElement(tag);
    if (props) {
      for (const [k, v] of Object.entries(props)) {
        if (v == null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'text') el.textContent = v;
        else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
        else if (k === 'hidden') el.hidden = true;
        else el.setAttribute(k, v === true ? '' : String(v));
      }
    }
    for (const c of children.flat(Infinity)) {
      if (c == null || c === false) continue;
      el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    }
    return el;
  }
  const SVGNS = 'http://www.w3.org/2000/svg';
  function svgEl(tag, attrs) {
    const el = document.createElementNS(SVGNS, tag);
    for (const [k, v] of Object.entries(attrs || {})) if (v != null) el.setAttribute(k, String(v));
    return el;
  }
  // Deep copy of plain JSON data. Strings (e.g. big data: URLs) are shared, not copied.
  function clone(v) {
    if (Array.isArray(v)) return v.map(clone);
    if (v && typeof v === 'object') { const o = {}; for (const k in v) o[k] = clone(v[k]); return o; }
    return v;
  }
  function readJSON(s) { if (!s) return null; try { return JSON.parse(s); } catch (_) { return null; } }
  const px = (emu) => (+emu || 0) / EMU_PX;
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  function extOf(name) { const m = /\.[^./\\]+$/.exec(name || ''); return m ? m[0].toLowerCase() : ''; }
  function baseName(p) { return (p || '').split(/[\\/]/).pop(); }
  function stripExt(n) { return String(n || '').replace(/\.[^.]+$/, ''); }
  function newId(prefix) { return (prefix || 'n') + (++idSeq) + Date.now().toString(36).slice(-4); }
  function fontStack(f) {
    if (!f) return '';
    const q = (x) => '"' + String(x).replace(/["\\]/g, '') + '"';
    return q(f) + ', "Segoe UI", system-ui, sans-serif';
  }
  function firstFamily(ff) { return String(ff || '').split(',')[0].trim().replace(/^["']|["']$/g, ''); }
  let colorCtx = null;
  function cssToHex(c) {
    if (!c) return null;
    c = String(c).trim();
    if (/^#[0-9a-f]{6}$/i.test(c)) return c.toUpperCase();
    if (/^#[0-9a-f]{3}$/i.test(c)) return ('#' + c[1] + c[1] + c[2] + c[2] + c[3] + c[3]).toUpperCase();
    if (c === 'transparent') return null;
    let m = /rgba?\(([^)]+)\)/i.exec(c);
    if (!m) {
      try {
        colorCtx = colorCtx || document.createElement('canvas').getContext('2d');
        colorCtx.fillStyle = '#000000'; colorCtx.fillStyle = c;
        const r = colorCtx.fillStyle;
        if (/^#/.test(r)) return r.toUpperCase();
        m = /rgba?\(([^)]+)\)/i.exec(r);
      } catch (_) { return null; }
      if (!m) return null;
    }
    const parts = m[1].split(/[\s,/]+/).filter(Boolean).map(parseFloat);
    if (parts.length > 3 && parts[3] === 0) return null;
    return '#' + parts.slice(0, 3).map(v => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function mix(hex, other, t) {
    const a = parseInt((hex || '#000000').slice(1), 16), b = parseInt((other || '#FFFFFF').slice(1), 16);
    const ch = (x, s) => (x >> s) & 255;
    const r = (s) => Math.round(ch(a, s) * (1 - t) + ch(b, s) * t);
    return '#' + [16, 8, 0].map(s => r(s).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function bytesToDataUrl(bytes, mime) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result);
      r.onerror = () => reject(r.error);
      r.readAsDataURL(new Blob([bytes], { type: mime }));
    });
  }
  function imageMime(name) {
    const e = extOf(name);
    return { '.png': 'image/png', '.gif': 'image/gif', '.webp': 'image/webp', '.bmp': 'image/bmp', '.svg': 'image/svg+xml' }[e] || 'image/jpeg';
  }
  function loadImageSize(src) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => resolve({ w: img.naturalWidth || 400, h: img.naturalHeight || 300 });
      img.onerror = () => resolve({ w: 400, h: 300 });
      img.src = src;
    });
  }

  /* ---------- icons ---------- */
  const LOCAL_ICONS = {
    alignObjLeft: 'M4 3v18M7 6h11v4H7zM7 14h7v4H7z',
    alignObjCenter: 'M12 3v18M6 6h12v4H6zM8.5 14h7v4h-7z',
    alignObjRight: 'M20 3v18M6 6h11v4H6zM10 14h7v4h-7z',
    alignObjTop: 'M3 4h18M6 7h4v11H6zM14 7h4v7h-4z',
    alignObjMiddle: 'M3 12h18M6 6h4v12H6zM14 8.5h4v7h-4z',
    alignObjBottom: 'M3 20h18M6 6h4v11H6zM14 10h4v7h-4z',
    distH: 'M3.5 4v16M20.5 4v16M9 7h6v10H9z',
    distV: 'M4 3.5h16M4 20.5h16M7 9h10v6H7z',
    background: 'M3.5 4.5h17v15h-17zM3.5 15l5-4.5 4 3.5 3-2.5 5 4',
    slideSize: 'M3.5 6h17v12h-17zM7 3v2M17 3v2M7 19v2M17 19v2',
    playCurrent: 'M3 4h18v13H3zM10 8v5l4.5-2.5zM12 17v4M8.5 21h7',
    size: 'M4 9V4h5M20 15v5h-5M4 4l7 7M20 20l-7-7',
    shapes: 'M3.5 11h9v9.5h-9z',
    clear: 'M6.5 6.5l11 11M17.5 6.5l-11 11'
  };
  function icon(name, size) {
    size = size || 18;
    try {
      if (window.App && typeof window.App.icon === 'function') {
        const r = window.App.icon(name, { size });
        if (r) return r;
      }
    } catch (_) { /* fall back */ }
    const svg = svgEl('svg', { viewBox: '0 0 24 24', width: size, height: size, fill: 'none', stroke: 'currentColor',
      'stroke-width': '1.5', 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true', focusable: 'false', class: 'icon icon-' + name });
    svg.appendChild(svgEl('path', { d: LOCAL_ICONS[name] || LOCAL_ICONS.shapes }));
    return svg;
  }
  function geomIcon(geom) {
    const svg = svgEl('svg', { viewBox: '0 0 24 24', width: 22, height: 22, fill: 'none', stroke: 'currentColor',
      'stroke-width': '1.5', 'stroke-linejoin': 'round', 'aria-hidden': 'true', focusable: 'false' });
    const g = svgEl('g', { transform: 'translate(3 5)' });
    if (geom === 'ellipse') g.appendChild(svgEl('ellipse', { cx: 9, cy: 7, rx: 8.5, ry: 6.5 }));
    else g.appendChild(svgEl('path', { d: geomPath(geom, 18, 14) }));
    if (geom === 'straightConnector1') g.appendChild(svgEl('path', { d: 'M18 14l-4.5-.6M18 14l-1.6-4.2' }));
    svg.appendChild(g);
    return svg;
  }

  /* ---------- geometry (shape-local px) ---------- */
  function pts(list) { return 'M' + list.map(p => p[0].toFixed(2) + ' ' + p[1].toFixed(2)).join(' L') + ' Z'; }
  function geomPath(geom, w, hh) {
    const m = Math.min(w, hh);
    switch (geom) {
      case 'roundRect': case 'flowChartAlternateProcess': {
        const r = m * 0.1667;
        return `M${r} 0 H${w - r} A${r} ${r} 0 0 1 ${w} ${r} V${hh - r} A${r} ${r} 0 0 1 ${w - r} ${hh} H${r} A${r} ${r} 0 0 1 0 ${hh - r} V${r} A${r} ${r} 0 0 1 ${r} 0 Z`;
      }
      case 'flowChartTerminator': {
        const r = Math.min(hh / 2, w / 2);
        return `M${r} 0 H${w - r} A${r} ${hh / 2} 0 0 1 ${w - r} ${hh} H${r} A${r} ${hh / 2} 0 0 1 ${r} 0 Z`;
      }
      case 'ellipse': case 'flowChartConnector':
        return `M0 ${hh / 2} A${w / 2} ${hh / 2} 0 1 1 ${w} ${hh / 2} A${w / 2} ${hh / 2} 0 1 1 0 ${hh / 2} Z`;
      case 'triangle': case 'flowChartExtract': return pts([[w / 2, 0], [w, hh], [0, hh]]);
      case 'rtTriangle': return pts([[0, 0], [w, hh], [0, hh]]);
      case 'diamond': case 'flowChartDecision': return pts([[w / 2, 0], [w, hh / 2], [w / 2, hh], [0, hh / 2]]);
      case 'pentagon': return pts([[w / 2, 0], [w, hh * 0.382], [w * 0.809, hh], [w * 0.191, hh], [0, hh * 0.382]]);
      case 'hexagon': { const o = m * 0.25; return pts([[o, 0], [w - o, 0], [w, hh / 2], [w - o, hh], [o, hh], [0, hh / 2]]); }
      case 'octagon': { const o = m * 0.2929; return pts([[o, 0], [w - o, 0], [w, o], [w, hh - o], [w - o, hh], [o, hh], [0, hh - o], [0, o]]); }
      case 'parallelogram': case 'flowChartInputOutput': { const o = m * 0.25; return pts([[o, 0], [w, 0], [w - o, hh], [0, hh]]); }
      case 'trapezoid': { const o = m * 0.25; return pts([[o, 0], [w - o, 0], [w, hh], [0, hh]]); }
      case 'chevron': { const o = m * 0.5; return pts([[0, 0], [w - o, 0], [w, hh / 2], [w - o, hh], [0, hh], [o, hh / 2]]); }
      case 'homePlate': { const o = m * 0.5; return pts([[0, 0], [w - o, 0], [w, hh / 2], [w - o, hh], [0, hh]]); }
      case 'plus': case 'flowChartOr': { const o = m * 0.25; return pts([[o, 0], [w - o, 0], [w - o, o], [w, o], [w, hh - o], [w - o, hh - o], [w - o, hh], [o, hh], [o, hh - o], [0, hh - o], [0, o], [o, o]]); }
      case 'star5': {
        const list = [];
        for (let i = 0; i < 10; i++) {
          const a = -Math.PI / 2 + i * Math.PI / 5, r = i % 2 ? 0.382 : 1;
          list.push([w / 2 + Math.cos(a) * r * w / 2, hh * 0.525 + Math.sin(a) * r * hh * 0.525]);
        }
        return pts(list);
      }
      case 'rightArrow': { const hl = Math.min(m * 0.5, w); return pts([[0, hh * 0.25], [w - hl, hh * 0.25], [w - hl, 0], [w, hh / 2], [w - hl, hh], [w - hl, hh * 0.75], [0, hh * 0.75]]); }
      case 'leftArrow': { const hl = Math.min(m * 0.5, w); return pts([[w, hh * 0.25], [hl, hh * 0.25], [hl, 0], [0, hh / 2], [hl, hh], [hl, hh * 0.75], [w, hh * 0.75]]); }
      case 'upArrow': { const hl = Math.min(m * 0.5, hh); return pts([[w * 0.25, hh], [w * 0.25, hl], [0, hl], [w / 2, 0], [w, hl], [w * 0.75, hl], [w * 0.75, hh]]); }
      case 'downArrow': { const hl = Math.min(m * 0.5, hh); return pts([[w * 0.25, 0], [w * 0.75, 0], [w * 0.75, hh - hl], [w, hh - hl], [w / 2, hh], [0, hh - hl], [w * 0.25, hh - hl]]); }
      case 'leftRightArrow': { const hl = Math.min(m * 0.5, w / 2); return pts([[0, hh / 2], [hl, 0], [hl, hh * 0.25], [w - hl, hh * 0.25], [w - hl, 0], [w, hh / 2], [w - hl, hh], [w - hl, hh * 0.75], [hl, hh * 0.75], [hl, hh]]); }
      case 'flowChartDocument': return `M0 0 H${w} V${hh * 0.83} C${w * 0.75} ${hh * 0.65} ${w * 0.5} ${hh * 1.1} 0 ${hh * 0.9} Z`;
      case 'line': case 'straightConnector1': case 'bentConnector2': case 'bentConnector3': case 'curvedConnector3':
        return `M0 0 L${w} ${hh}`;
      default: return `M0 0 H${w} V${hh} H0 Z`;
    }
  }
  function dashArray(dash, wpx) {
    const d = Math.max(wpx, 1);
    switch (dash) {
      case 'dash': case 'sysDash': return `${4 * d} ${3 * d}`;
      case 'dot': case 'sysDot': return `${d} ${d}`;
      case 'lgDash': return `${8 * d} ${3 * d}`;
      case 'dashDot': case 'sysDashDot': return `${4 * d} ${3 * d} ${d} ${3 * d}`;
      case 'lgDashDot': return `${8 * d} ${3 * d} ${d} ${3 * d}`;
      default: return null;
    }
  }

  /* ---------- rendering (shared by canvas, thumbnails, slide show and print) ---------- */
  function themeOf(model) { return (model && model.theme) || DEFAULT_THEME; }
  function themeColor(model, key, dflt) { const c = themeOf(model).colors || {}; return c[key] || DEFAULT_THEME.colors[key] || dflt; }
  function mediaUrl(model, src) {
    if (!src) return '';
    if (typeof src === 'string' && src.startsWith('media:')) return (model.media || {})[src.slice(6)] || '';
    return src;
  }
  function cssUrl(u) { return 'url("' + String(u).replace(/["\\\n]/g, m => '\\' + m) + '")'; }
  function slideSizePx(model) { return { W: px(model.size.cx), H: px(model.size.cy) }; }

  // ctx: {model, index, hint, edit}
  function renderSlide(model, slide, ctx) {
    ctx = Object.assign({ model, index: 0 }, ctx || {});
    const { W, H } = slideSizePx(model);
    const el = h('div', { class: 'sl-slide' });
    el.style.width = W + 'px';
    el.style.height = H + 'px';
    el.style.backgroundColor = themeColor(model, 'lt1', '#FFFFFF');
    const bg = slide.bg;
    if (bg && bg.color) el.style.backgroundColor = bg.color;
    if (bg && bg.image) {
      const u = mediaUrl(model, bg.image);
      if (u) { el.style.backgroundImage = cssUrl(u); el.style.backgroundSize = '100% 100%'; }
    }
    el.style.fontFamily = fontStack(themeOf(model).minorFont || 'Calibri');
    for (const s of slide.decor || []) {
      const d = renderShape(model, s, Object.assign({}, ctx, { decor: true, hint: false }));
      if (d) el.appendChild(d);
    }
    for (const s of slide.shapes || []) {
      const d = renderShape(model, s, ctx);
      if (d) el.appendChild(d);
    }
    return el;
  }

  function renderShape(model, s, ctx) {
    const kind = s.kind || 'shape';
    const w = Math.max(px(s.w), 0), hh = Math.max(px(s.h), 0);
    const el = h('div', { class: 'sl-shape sl-k-' + kind + (ctx.decor ? ' sl-decor' : '') });
    if (!ctx.decor && s.id != null) el.dataset.id = s.id;
    el.style.left = px(s.x) + 'px';
    el.style.top = px(s.y) + 'px';
    el.style.width = w + 'px';
    el.style.height = hh + 'px';
    if (s.rot) el.style.transform = 'rotate(' + s.rot + 'deg)';
    const flip = (s.flipH ? ' scaleX(-1)' : '') + (s.flipV ? ' scaleY(-1)' : '');

    if (kind === 'opaque') {
      if (s.children && s.children.length) {
        // group: children carry absolute slide coordinates, display only
        const inner = h('div', { class: 'sl-group' });
        inner.style.left = -px(s.x) + 'px';
        inner.style.top = -px(s.y) + 'px';
        for (const c of s.children) { const ce = renderShape(model, c, Object.assign({}, ctx, { decor: true, hint: false })); if (ce) inner.appendChild(ce); }
        el.appendChild(inner);
      } else if (s.preview && mediaUrl(model, s.preview)) {
        el.appendChild(h('img', { class: 'sl-opaque-img', src: mediaUrl(model, s.preview), alt: '', draggable: 'false' }));
      } else {
        el.appendChild(h('div', { class: 'sl-opaque-box' }, h('span', null, s.name || 'Embedded object')));
      }
      return el;
    }
    if (kind === 'table') { el.appendChild(renderTable(model, s, ctx)); return el; }

    // geometry (fill + outline)
    const line = s.line && s.line.color ? s.line : null;
    const fillImg = s.fillImage ? mediaUrl(model, s.fillImage) : '';
    const isLine = LINE_GEOMS.has(s.geom);
    if (kind !== 'image' && (s.fill || line || fillImg)) {
      const svg = svgEl('svg', { class: 'sl-geom', width: Math.max(w, 1), height: Math.max(hh, 1), viewBox: `0 0 ${Math.max(w, 1)} ${Math.max(hh, 1)}`, preserveAspectRatio: 'none', overflow: 'visible' });
      if (flip) svg.style.transform = flip.trim();
      let d;
      if (s.geom === 'custom' && s.path) {
        const g = svgEl('g', { transform: `scale(${1 / EMU_PX})` });
        svg.appendChild(g);
        d = svgEl('path', { d: s.path, 'vector-effect': 'non-scaling-stroke' });
        g.appendChild(d);
      } else {
        d = svgEl('path', { d: geomPath(s.geom || 'rect', w, hh) });
        svg.appendChild(d);
      }
      if (fillImg && !isLine) {
        const pid = 'slp' + newId('f');
        const pat = svgEl('pattern', { id: pid, patternUnits: 'objectBoundingBox', width: 1, height: 1 });
        pat.appendChild(svgEl('image', { href: fillImg, x: 0, y: 0, width: Math.max(w, 1), height: Math.max(hh, 1), preserveAspectRatio: 'none' }));
        const defs = svgEl('defs'); defs.appendChild(pat); svg.insertBefore(defs, svg.firstChild);
        d.setAttribute('fill', `url(#${pid})`);
      } else d.setAttribute('fill', s.fill && !isLine ? s.fill : 'none');
      if (line) {
        const lw = Math.max(px(line.width || EMU_PT), 0.5);
        d.setAttribute('stroke', line.color);
        d.setAttribute('stroke-width', lw);
        const da = dashArray(line.dash, lw);
        if (da) d.setAttribute('stroke-dasharray', da);
        if (s.geom === 'straightConnector1' || line.tail || line.head) {
          const mid = 'sla' + newId('m');
          const mk = svgEl('marker', { id: mid, viewBox: '0 0 10 10', refX: 8, refY: 5, markerWidth: 5, markerHeight: 5, orient: 'auto-start-reverse' });
          mk.appendChild(svgEl('path', { d: 'M0 0 L10 5 L0 10 z', fill: line.color }));
          const defs = svgEl('defs'); defs.appendChild(mk); svg.insertBefore(defs, svg.firstChild);
          if (s.geom === 'straightConnector1' || (line.tail && line.tail !== 'none')) d.setAttribute('marker-end', `url(#${mid})`);
          if (line.head && line.head !== 'none') d.setAttribute('marker-start', `url(#${mid})`);
        }
      }
      el.appendChild(svg);
    }

    if (kind === 'image') {
      const src = mediaUrl(model, s.src);
      const wrap = h('div', { class: 'sl-img-wrap' });
      if (flip) wrap.style.transform = flip.trim();
      if (s.geom === 'ellipse') wrap.style.clipPath = 'ellipse(50% 50% at 50% 50%)';
      else if (s.geom === 'roundRect') wrap.style.borderRadius = (Math.min(w, hh) * 0.1667) + 'px';
      if (src) {
        const c = s.crop || {};
        const l = +c.l || 0, t = +c.t || 0, r = +c.r || 0, b = +c.b || 0;
        const iw = w / Math.max(1 - l - r, 0.01), ih = hh / Math.max(1 - t - b, 0.01);
        const img = h('img', { src, alt: s.descr || s.name || '', draggable: 'false' });
        img.style.width = iw + 'px'; img.style.height = ih + 'px';
        img.style.left = (-l * iw) + 'px'; img.style.top = (-t * ih) + 'px';
        wrap.appendChild(img);
      } else wrap.appendChild(h('div', { class: 'sl-opaque-box' }, h('span', null, 'Picture')));
      if (line) {
        const lw = Math.max(px(line.width || EMU_PT), 0.5);
        wrap.style.outline = lw + 'px ' + (line.dash && line.dash !== 'solid' ? 'dashed' : 'solid') + ' ' + line.color;
        wrap.style.outlineOffset = (-lw / 2) + 'px';
      }
      el.appendChild(wrap);
    }

    if (kind !== 'image' && (s.paras && s.paras.length || (ctx.hint && s.ph && PH_HINT[s.ph] !== ''))) {
      el.appendChild(renderTextBox(model, s, s.paras || [], ctx));
    }
    return el;
  }

  function parasEmpty(paras) {
    return !paras || !paras.length || paras.every(p => !(p.runs || []).some(r => r.text && r.text !== '\n'));
  }

  function renderTextBox(model, s, paras, ctx) {
    const inset = Array.isArray(s.inset) && s.inset.length === 4 ? s.inset : DEFAULT_INSET;
    const box = h('div', { class: 'sl-text' });
    box.style.padding = inset.map(v => px(v) + 'px').slice(0, 4).join(' ').replace(/^(\S+) (\S+) (\S+) (\S+)$/, '$2 $3 $4 $1');
    box.style.justifyContent = ANCHOR_CSS[s.anchor] || 'flex-start';
    box.style.color = themeColor(model, 'dk1', '#000000');
    if (s.wrap === 'none') box.style.whiteSpace = 'pre';
    box.style.setProperty('--fs', String((s.fontScale || 100) / 100));
    const body = h('div', { class: 'sl-tbody' });
    fillParas(body, paras, s, ctx);
    if (ctx.hint && s.ph && PH_HINT[s.ph] && parasEmpty(paras)) {
      const hint = h('div', { class: 'sl-hint' });
      hint.style.padding = box.style.padding;
      hint.style.justifyContent = box.style.justifyContent;
      const p0 = paras[0] || { align: s.ph === 'ctrTitle' || s.ph === 'subTitle' ? 'ctr' : 'l' };
      const hp = h('div', { class: 'sl-p' }, PH_HINT[s.ph]);
      styleParaDiv(hp, Object.assign({}, p0, { bullet: null }), { lnRed: s.lnSpcReduction, first: (p0.end || null) });
      hint.appendChild(hp);
      box.appendChild(hint);
    }
    box.appendChild(body);
    return box;
  }

  function runMeta(r) { const o = Object.assign({}, r); delete o.text; return o; }
  function paraMeta(p) { const o = Object.assign({}, p); delete o.runs; return o; }
  function firstRunOf(p) { return (p.runs || []).find(r => r.text && r.text !== '\n') || (p.runs || [])[0] || null; }

  function fillParas(body, paras, s, ctx) {
    body.textContent = '';
    for (const p of paras) {
      const div = h('div', { class: 'sl-p' });
      div.dataset.pp = JSON.stringify(paraMeta(p));
      const fr = firstRunOf(p);
      div.dataset.dr = JSON.stringify(runMeta(fr || p.end || {}));
      const runs = p.runs || [];
      for (const r of runs) {
        if (r.text === '\n') {
          const br = document.createElement('br');
          br.dataset.rr = JSON.stringify(runMeta(r));
          div.appendChild(br);
          continue;
        }
        if (!r.text) continue;
        const span = h('span', { class: 'sl-r' + (r.link ? ' sl-link' : '') });
        span.dataset.rr = JSON.stringify(runMeta(r));
        applyRunStyle(span.style, r);
        span.textContent = r.field === 'slidenum' && ctx && ctx.index != null ? String(ctx.index + 1) : r.text;
        div.appendChild(span);
      }
      if (!div.lastChild || div.lastChild.nodeName === 'BR') div.appendChild(h('br', { class: 'sl-tail' }));
      body.appendChild(div);
    }
    refreshParas(body, s);
  }

  function applyRunStyle(st, r) {
    if (r.font) st.fontFamily = fontStack(r.font);
    const size = r.size ? (r.baseline ? r.size * 0.66 : r.size) : null;
    if (size) st.fontSize = 'calc(var(--fs, 1) * ' + size + 'pt)';
    if (r.b) st.fontWeight = 'bold';
    if (r.i) st.fontStyle = 'italic';
    const deco = [];
    if (r.u || r.link) deco.push('underline');
    if (r.strike) deco.push('line-through');
    if (deco.length) st.textDecoration = deco.join(' ');
    if (r.color) st.color = r.color;
    if (r.highlight) st.backgroundColor = r.highlight;
    if (r.baseline) st.verticalAlign = r.baseline > 0 ? 'super' : 'sub';
    if (r.caps === 'all' || r.caps === true) st.textTransform = 'uppercase';
    else if (r.caps === 'small') st.fontVariant = 'small-caps';
    if (r.spc) st.letterSpacing = (r.spc / 100) + 'pt';
  }

  function autoNumber(scheme, n) {
    const roman = (x) => {
      const t = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
      let out = ''; for (const [v, s] of t) while (x >= v) { out += s; x -= v; } return out;
    };
    const alpha = (x) => { let s = ''; while (x > 0) { x--; s = String.fromCharCode(97 + (x % 26)) + s; x = Math.floor(x / 26); } return s; };
    scheme = String(scheme || 'arabicPeriod');
    let core;
    if (/^romanUc/.test(scheme)) core = roman(n).toUpperCase();
    else if (/^romanLc/.test(scheme)) core = roman(n);
    else if (/^alphaUc/.test(scheme)) core = alpha(n).toUpperCase();
    else if (/^alphaLc/.test(scheme)) core = alpha(n);
    else core = String(n);
    if (/ParenBoth$/.test(scheme)) return '(' + core + ')';
    if (/ParenR$/.test(scheme)) return core + ')';
    if (/Period$/.test(scheme)) return core + '.';
    if (/Minus$/.test(scheme)) return core + ' -';
    return core;
  }

  // Applies paragraph props (from data-pp) to every paragraph div; recomputes auto numbers.
  function refreshParas(body, s) {
    const counters = [];
    const lnRed = s ? s.lnSpcReduction : 0;
    for (const div of body.children) {
      if (!div.classList.contains('sl-p')) continue;
      const pp = readJSON(div.dataset.pp) || {};
      const lvl = pp.level || 0;
      let num = 0;
      if (pp.bullet && pp.bullet.auto) {
        const c = counters[lvl];
        num = c && c.scheme === pp.bullet.auto ? c.n + 1 : (pp.bullet.startAt || 1);
        counters[lvl] = { scheme: pp.bullet.auto, n: num };
        counters.length = lvl + 1;
      } else if (!parasEmptyDiv(div)) counters.length = lvl;
      const span = div.querySelector('[data-rr]');
      const first = (span && readJSON(span.dataset.rr)) || readJSON(div.dataset.dr) || pp.end || null;
      styleParaDiv(div, pp, { lnRed, first, num });
      if (parasEmptyDiv(div)) div.dataset.empty = ''; else delete div.dataset.empty;
    }
  }
  function parasEmptyDiv(div) { return !div.textContent.replace(/\u200b/g, ''); }

  function styleParaDiv(div, pp, c) {
    const st = div.style;
    st.textAlign = ALIGN_CSS[pp.align] || 'left';
    st.marginTop = pp.spaceBefore ? pp.spaceBefore + 'pt' : '';
    st.marginBottom = pp.spaceAfter ? pp.spaceAfter + 'pt' : '';
    if (pp.lineSpacingPts) st.lineHeight = 'calc(var(--fs, 1) * ' + pp.lineSpacingPts + 'pt)';
    else st.lineHeight = String(Math.max((pp.lineSpacing || 100) - (c.lnRed || 0), 20) / 100 * 1.2);
    st.paddingLeft = px(pp.marL || 0) + 'px';
    st.textIndent = px(pp.indent || 0) + 'px';
    const first = c.first || {};
    const size = first.size || (pp.end && pp.end.size) || 18;
    st.fontSize = 'calc(var(--fs, 1) * ' + size + 'pt)';
    if (first.font) st.fontFamily = fontStack(first.font); else st.fontFamily = '';
    if (first.color) st.color = first.color; else st.color = '';
    const b = pp.bullet;
    if (b && (b.char || b.auto)) {
      let txt = b.char || autoNumber(b.auto, c.num || 1);
      if (!((pp.indent || 0) < 0)) txt += '\u00a0';
      div.dataset.bullet = txt;
      st.setProperty('--bw', (pp.indent || 0) < 0 ? px(-pp.indent) + 'px' : '0px');
      st.setProperty('--bf', b.font ? fontStack(b.font) : 'inherit');
      st.setProperty('--bc', b.color || first.color || 'inherit');
      st.setProperty('--bs', 'calc(var(--fs, 1) * ' + (size * (b.size || 100) / 100) + 'pt)');
    } else {
      delete div.dataset.bullet;
      ['--bw', '--bf', '--bc', '--bs'].forEach(k => st.removeProperty(k));
    }
  }

  function renderTable(model, s, ctx) {
    const t = h('table', { class: 'sl-table' });
    const colW = s.colW || [];
    const totalW = colW.reduce((a, b) => a + (+b || 0), 0) || s.w;
    t.style.width = px(totalW) + 'px';
    const cg = h('colgroup');
    colW.forEach(cw => { const col = h('col'); col.style.width = px(cw) + 'px'; cg.appendChild(col); });
    t.appendChild(cg);
    const tb = h('tbody');
    (s.rows || []).forEach((row, ri) => {
      const tr = h('tr');
      if (s.rowH && s.rowH[ri]) tr.style.height = px(s.rowH[ri]) + 'px';
      row.forEach((c, ci) => {
        if (!c || c.hMerge || c.vMerge) return;
        const td = h('td', { 'data-r': ri, 'data-c': ci });
        if (c.span > 1) td.colSpan = c.span;
        if (c.rowSpan > 1) td.rowSpan = c.rowSpan;
        if (c.fill) td.style.backgroundColor = c.fill;
        const bd = c.borders || {};
        for (const [side, key] of [['Left', 'l'], ['Right', 'r'], ['Top', 't'], ['Bottom', 'b']]) {
          const b = bd[key];
          td.style['border' + side] = b && b.color ? Math.max(px(b.width || EMU_PT), 0.5) + 'px ' + (b.dash && b.dash !== 'solid' ? 'dashed' : 'solid') + ' ' + b.color : 'none';
        }
        const inset = Array.isArray(c.inset) && c.inset.length === 4 ? c.inset : DEFAULT_INSET;
        td.style.padding = [inset[1], inset[2], inset[3], inset[0]].map(v => px(v) + 'px').join(' ');
        td.style.verticalAlign = c.anchor === 'ctr' ? 'middle' : c.anchor === 'b' ? 'bottom' : 'top';
        const body = h('div', { class: 'sl-tbody' });
        fillParas(body, c.paras && c.paras.length ? c.paras : [{ runs: [], end: { size: 18 } }], null, ctx);
        td.appendChild(body);
        tr.appendChild(td);
      });
      tb.appendChild(tr);
    });
    t.appendChild(tb);
    return t;
  }

  /* ---------- reading edited text back into paragraphs ---------- */
  const BLOCK_TAGS = /^(DIV|P|LI|H[1-6]|BLOCKQUOTE|PRE|UL|OL|TABLE|TR|TD)$/;
  function hasBlockChild(el) { for (const c of el.children) if (BLOCK_TAGS.test(c.tagName)) return true; return false; }

  function applyElFormat(props, e, isRun) {
    const tag = e.tagName;
    if (tag === 'B' || tag === 'STRONG') props.b = true;
    else if (tag === 'I' || tag === 'EM') props.i = true;
    else if (tag === 'U') props.u = true;
    else if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') props.strike = true;
    else if (tag === 'SUP') props.baseline = 30;
    else if (tag === 'SUB') props.baseline = -25;
    else if (tag === 'A' && e.getAttribute('href')) props.link = e.getAttribute('href');
    else if (tag === 'FONT') {
      if (e.getAttribute('color')) props.color = cssToHex(e.getAttribute('color')) || props.color;
      if (e.getAttribute('face')) props.font = firstFamily(e.getAttribute('face'));
    }
    const st = e.style;
    if (!st) return;
    if (st.fontWeight) props.b = st.fontWeight === 'bold' || st.fontWeight === 'bolder' || +st.fontWeight >= 600;
    if (st.fontStyle) props.i = st.fontStyle === 'italic' || st.fontStyle === 'oblique';
    const td = st.textDecorationLine || st.textDecoration;
    if (td) {
      if (td.includes('none')) { props.u = false; props.strike = false; }
      if (td.includes('underline') && !(isRun && props.link)) props.u = true;
      if (td.includes('line-through')) props.strike = true;
    }
    if (st.color) props.color = cssToHex(st.color) || props.color;
    if (st.backgroundColor) props.highlight = cssToHex(st.backgroundColor);
    if (st.fontFamily && !(isRun && props.font && firstFamily(st.fontFamily) === props.font)) props.font = firstFamily(st.fontFamily);
    const fsz = st.fontSize;
    if (fsz && !/calc|var/.test(fsz)) {
      const m = /^([\d.]+)(pt|px)$/.exec(fsz);
      if (m) props.size = Math.round((m[2] === 'px' ? +m[1] * 0.75 : +m[1]) * 2) / 2;
    }
    if (st.verticalAlign && !isRun) {
      if (st.verticalAlign === 'super') props.baseline = 30;
      else if (st.verticalAlign === 'sub') props.baseline = -25;
      else if (st.verticalAlign === 'baseline') delete props.baseline;
    }
  }

  function isTrailingBr(br, block) {
    if (br.classList.contains('sl-tail')) return true;
    if (br.dataset.rr) return false;
    try {
      const r = document.createRange();
      r.setStartAfter(br);
      r.setEnd(block, block.childNodes.length);
      if (r.toString().replace(/\u200b/g, '') !== '') return false;
      return !r.cloneContents().querySelector('br:not(.sl-tail)');
    } catch (_) { return true; }
  }

  // Parses a contenteditable text body back into model paragraphs.
  function parseEditable(root, fallback) {
    const out = [];
    let cur = null;
    const start = (el) => {
      const prev = cur;
      const pp = (el && el.dataset && readJSON(el.dataset.pp)) || (prev ? paraMeta(prev) : clone(fallback.pp || {}));
      const dr = (el && el.dataset && readJSON(el.dataset.dr)) || (prev ? prev._dr : clone(fallback.dr || {}));
      cur = Object.assign(clone(pp), { runs: [] });
      Object.defineProperty(cur, '_dr', { value: dr, enumerable: false, writable: true });
      out.push(cur);
    };
    const pushRun = (props, text) => {
      const last = cur.runs[cur.runs.length - 1];
      if (last && text !== '\n' && last.text !== '\n' && JSON.stringify(runMeta(last)) === JSON.stringify(props)) last.text += text;
      else cur.runs.push(Object.assign({}, props, { text }));
    };
    const propsFor = (node, block) => {
      const chain = [];
      for (let e = node.nodeType === 1 ? node : node.parentElement; e && e !== block && e !== root; e = e.parentElement) chain.unshift(e);
      let base = null;
      for (let i = chain.length - 1; i >= 0; i--) if (chain[i].dataset && chain[i].dataset.rr) { base = readJSON(chain[i].dataset.rr); break; }
      if (!base && node.nodeType === 1 && node.dataset && node.dataset.rr) base = readJSON(node.dataset.rr);
      const props = clone(base || cur._dr || {});
      for (const e of chain) if (e !== node) applyElFormat(props, e, !!(e.dataset && e.dataset.rr));
      delete props.text;
      for (const k of Object.keys(props)) if (props[k] === undefined) delete props[k];
      return props;
    };
    const walk = (node, block) => {
      for (const n of Array.from(node.childNodes)) {
        if (n.nodeType === 3) {
          const t = n.data.replace(/\u200b/g, '');
          if (!t) continue;
          if (!cur) start(null);
          pushRun(propsFor(n, block), t.replace(/\u00a0/g, ' '));
        } else if (n.nodeType === 1) {
          if (n.tagName === 'BR') {
            if (!cur) start(null);
            if (!isTrailingBr(n, block)) pushRun(propsFor(n, block), '\n');
          } else if (BLOCK_TAGS.test(n.tagName)) {
            if (hasBlockChild(n)) walk(n, n);
            else { start(n); walk(n, n); cur = null; }
          } else if (n.tagName !== 'STYLE' && n.tagName !== 'SCRIPT') walk(n, block);
        }
      }
    };
    walk(root, root);
    if (!out.length) start(null);
    for (const p of out) {
      if (!p.runs.some(r => r.text)) { p.end = p.end || clone(p._dr || {}); p.runs = []; }
      // drop a lone trailing line-break run
      if (p.runs.length && p.runs[p.runs.length - 1].text === '\n' && p.runs.length === 1) p.runs = [];
    }
    return out.map(p => { const o = {}; for (const k of Object.keys(p)) o[k] = p[k]; return o; });
  }

  function shapeText(s) {
    const t = [];
    const paras = (p) => p.map(x => (x.runs || []).map(r => r.text).join('')).join('\n');
    if (s.paras) t.push(paras(s.paras));
    if (s.rows) for (const row of s.rows) t.push(row.map(c => c && c.paras ? paras(c.paras) : '').join('\t'));
    return t.join('\n').trim();
  }
  function slideTitle(sl) {
    const ts = (sl.shapes || []).find(s => TITLE_PH.has(s.ph) && shapeText(s));
    if (ts) return shapeText(ts).split('\n')[0];
    const any = (sl.shapes || []).find(s => s.paras && shapeText(s));
    return any ? shapeText(any).split('\n')[0].slice(0, 80) : 'Untitled';
  }

  /* ---------- blank model and slide templates ---------- */
  function emptyPara(model, o) {
    o = o || {};
    const end = { size: o.size || 18, font: o.font || themeOf(model).minorFont || 'Calibri', color: o.color || themeColor(model, 'dk1', '#000000') };
    if (o.b) end.b = true;
    const p = { align: o.align || 'l', level: 0, bullet: o.bullet || null, runs: [], end };
    if (o.marL != null) p.marL = o.marL;
    if (o.indent != null) p.indent = o.indent;
    return p;
  }
  function scaleBox(model, x, y, w, hh) {
    const sx = model.size.cx / 12192000, sy = model.size.cy / 6858000;
    return { x: Math.round(x * sx), y: Math.round(y * sy), w: Math.round(w * sx), h: Math.round(hh * sy) };
  }
  function defaultPlaceholders(model, layoutName) {
    const n = String(layoutName || '').toLowerCase();
    const major = themeOf(model).majorFont || 'Calibri Light';
    const mk = (ph, name, box, para, extra) => Object.assign({
      id: newId('n'), name, kind: 'text', ph, rot: 0, geom: 'rect', fill: null, line: null,
      anchor: 'ctr', inset: DEFAULT_INSET.slice(), autofit: 'norm', paras: [para]
    }, scaleBox(model, ...box), extra || {});
    if (n.includes('blank')) return [];
    if (n.includes('title slide') || !n) {
      return [
        mk('ctrTitle', 'Title 1', [1524000, 1122363, 9144000, 2387600], emptyPara(model, { align: 'ctr', size: 60, font: major }), { anchor: 'b' }),
        mk('subTitle', 'Subtitle 2', [1524000, 3602038, 9144000, 1655762], emptyPara(model, { align: 'ctr', size: 24 }), { anchor: 't', phIdx: '1' })
      ];
    }
    if (n.includes('section')) {
      return [
        mk('title', 'Title 1', [831850, 1709738, 10515600, 2852737], emptyPara(model, { size: 60, font: major }), { anchor: 'b' }),
        mk('body', 'Text Placeholder 2', [831850, 4589463, 10515600, 1500187], emptyPara(model, { size: 24, color: themeColor(model, 'dk2', '#44546A') }), { anchor: 't', phIdx: '1' })
      ];
    }
    const title = mk('title', 'Title 1', [838200, 365125, 10515600, 1325563], emptyPara(model, { size: 44, font: major }), { anchor: 'ctr' });
    if (n.includes('title only')) return [title];
    return [title, mk('body', 'Content Placeholder 2', [838200, 1825625, 10515600, 4351338],
      emptyPara(model, { size: 28, bullet: { char: '•', font: 'Arial' }, marL: 228600, indent: -228600 }), { anchor: 't', phIdx: '1' })];
  }
  function blankModel() {
    const model = { size: { cx: 12192000, cy: 6858000 }, theme: clone(DEFAULT_THEME), layouts: [], media: {}, slides: [] };
    model.slides.push({ id: newId('s'), src: null, layout: '', hidden: false, bg: null, notes: '', decor: [], shapes: defaultPlaceholders(model, 'Title Slide') });
    return model;
  }
  function normalizeModel(m) {
    m = m && typeof m === 'object' ? m : {};
    if (!m.size || !m.size.cx || !m.size.cy) m.size = { cx: 12192000, cy: 6858000 };
    m.theme = m.theme || clone(DEFAULT_THEME);
    m.theme.colors = Object.assign({}, DEFAULT_THEME.colors, m.theme.colors || {});
    m.layouts = Array.isArray(m.layouts) ? m.layouts : [];
    m.media = m.media || {};
    m.slides = Array.isArray(m.slides) ? m.slides : [];
    for (const s of m.slides) {
      if (s.id == null) s.id = newId('s');
      s.shapes = Array.isArray(s.shapes) ? s.shapes : [];
      s.decor = Array.isArray(s.decor) ? s.decor : [];
      const seen = new Set();
      for (const sh of s.shapes) {
        if (sh.id == null || seen.has(String(sh.id))) sh.id = newId('n');
        seen.add(String(sh.id));
      }
    }
    return m;
  }
  function allParas(s) {
    const out = [];
    if (s.paras) out.push(...s.paras);
    if (s.rows) for (const row of s.rows) for (const c of row) if (c && c.paras && !c.hMerge && !c.vMerge) out.push(...c.paras);
    return out;
  }
  function forEachRun(shapes, fn) {
    for (const s of shapes) for (const p of allParas(s)) {
      for (const r of p.runs || []) if (r.text !== '\n') fn(r, p, s);
      if (p.end) fn(p.end, p, s);
    }
  }
  function stepSize(size, dir) {
    size = size || 18;
    if (dir > 0) return SIZES.find(z => z > size) || Math.min(size + 8, 400);
    return [...SIZES].reverse().find(z => z < size) || Math.max(size - 1, 1);
  }
  function resolveMedia(v, media) {
    if (typeof v === 'string') return v.startsWith('media:') && media && media[v.slice(6)] ? media[v.slice(6)] : v;
    if (Array.isArray(v)) return v.map(x => resolveMedia(x, media));
    if (v && typeof v === 'object') { const o = {}; for (const k in v) o[k] = resolveMedia(v[k], media); return o; }
    return v;
  }
  function kindLabel(s) {
    if (s.ph) return TITLE_PH.has(s.ph) ? 'title placeholder' : 'text placeholder';
    return { text: 'text box', shape: 'shape', image: 'picture', table: 'table', opaque: 'embedded object' }[s.kind] || 'object';
  }

  /* ---------- the editor ---------- */
  function createSlidesEditor(container, doc) {
    const uid = 'sle' + (++editorCounter);
    const listeners = [];
    const on = (t, ty, fn, o) => { t.addEventListener(ty, fn, o); listeners.push([t, ty, fn, o]); };

    let docPath = doc.path || '';
    let docName = doc.name || (docPath ? baseName(docPath) : 'Presentation1.pptx');
    let dirty = false, loading = true, destroyed = false;
    let model = normalizeModel(blankModel());
    let original = null;
    let cur = 0;
    let selected = [];          // ids (strings) of selected shapes on the current slide
    let editing = null;         // active text edit
    let lastCell = null;        // [row, col] of the last edited table cell
    let zoomMode = 'fit', zoomPct = 100, scale = 0.5;
    let view = 'normal', notesOn = true, snapOn = true, tabOut = false;
    let undoStack = [], redoStack = [];
    let thumbTimer = 0, stateTimer = 0, drag = null, menu = null, showExit = null;
    const origBg = new Map(), origGeom = new Map();
    const controls = {};
    const colorBtns = [];
    const needSel = [];
    const last = { color: '#C00000', highlight: '#FFFF00', fill: null, line: null };
    let editor;

    const slide = () => model.slides[cur];
    const W = () => px(model.size.cx);
    const H = () => px(model.size.cy);
    const findShape = (id) => slide() ? (slide().shapes || []).find(s => String(s.id) === String(id)) : null;
    const selShapes = () => selected.map(findShape).filter(Boolean);
    const textShapes = () => selShapes().filter(s => s.paras || s.rows);

    /* ---- DOM skeleton ---- */
    const root = h('div', { class: 'editor-root slides-editor', id: uid });
    const live = h('div', { class: 'visually-hidden', 'aria-live': 'polite', 'aria-atomic': 'true' });
    function announce(t) { live.textContent = ''; setTimeout(() => { if (!destroyed) live.textContent = t; }, 40); }

    const holder = h('div', { class: 'sl-holder' });
    const overlay = h('div', { class: 'sl-overlay', 'aria-hidden': 'true' });
    let slideEl = h('div', { class: 'sl-slide' });
    holder.append(slideEl, overlay);
    const stageHelp = h('div', { id: uid + '-stagehelp', class: 'visually-hidden' },
      'Tab selects the next object, Escape clears the selection, Enter edits text, arrow keys move the selection (Shift for bigger steps, Alt+Left or Right to rotate), Delete removes it. Page Down and Page Up change slide. F6 moves between ribbon, slides, slide and notes.');
    const stage = h('div', { class: 'sl-stage', tabindex: '0', role: 'region', 'aria-roledescription': 'slide editor', 'aria-label': 'Slide', 'aria-describedby': uid + '-stagehelp' },
      h('div', { class: 'sl-stage-inner' }, holder), stageHelp);
    const notes = h('textarea', { id: uid + '-notes', class: 'sl-notes', placeholder: 'Click to add notes', spellcheck: 'true', 'aria-label': 'Speaker notes' });
    const notesWrap = h('div', { class: 'sl-notes-wrap' }, notes);
    const center = h('div', { class: 'sl-center' }, stage, notesWrap);
    const thumbs = h('div', { class: 'panel-side sl-thumbs', role: 'listbox', 'aria-label': 'Slides', 'aria-orientation': 'vertical' });
    const sorter = h('div', { class: 'sl-sorter', role: 'listbox', 'aria-label': 'Slide sorter', hidden: true });
    const main = h('div', { class: 'sl-main' }, thumbs, center, sorter);

    /* ---- ribbon helpers ---- */
    function btn(o) {
      const title = (o.title || o.label) + (o.keys ? ' (' + o.keys + ')' : '');
      const b = h('button', {
        type: 'button', class: 'tool-btn' + (o.big ? ' big' : '') + (o.cls ? ' ' + o.cls : ''), title,
        'aria-label': o.iconOnly ? o.label : null, 'aria-keyshortcuts': o.aria || null,
        'aria-pressed': o.pressable ? 'false' : null, 'aria-haspopup': o.menu ? 'menu' : null, 'aria-expanded': o.menu ? 'false' : null
      },
      o.icon ? h('span', { class: 'ico', 'aria-hidden': 'true' }, icon(o.icon, o.big ? 22 : 18)) : null,
      o.iconOnly ? null : h('span', { class: 'lbl' }, o.label),
      o.menu && !o.noCaret ? h('span', { class: 'sl-caret', 'aria-hidden': 'true' }, icon('chevronDown', 12)) : null);
      b.addEventListener('mousedown', e => e.preventDefault());
      b.addEventListener('click', () => {
        Promise.resolve().then(() => o.run(b)).catch(err => App.toast((o.errPrefix || '') + (err && err.message || String(err)), 'error'));
      });
      if (o.id) controls[o.id] = b;
      if (o.needsSel) needSel.push(b);
      return b;
    }
    function group(label, ...items) {
      return h('div', { class: 'tool-group', role: 'group', 'aria-label': label },
        h('div', { class: 'tool-group-items' }, items),
        h('div', { class: 'tool-group-label', 'aria-hidden': 'true' }, label));
    }
    function row(...items) { return h('div', { class: 'sl-row' }, items); }
    function setPressed(b, v) { if (b) b.setAttribute('aria-pressed', String(!!v)); }
    function colorButton(o) {
      const bar = h('span', { class: 'sl-colorbar', 'aria-hidden': 'true' });
      const mainB = btn({ icon: o.icon, label: o.label, iconOnly: true, keys: o.keys, needsSel: o.needsSel, run: () => o.apply(o.get()) });
      mainB.querySelector('.ico').appendChild(bar);
      const arrow = btn({ icon: 'chevronDown', label: o.label + ' - more colors', iconOnly: true, cls: 'sl-split-arrow', menu: true, noCaret: true, needsSel: o.needsSel,
        run: (b) => openColorMenu(b, o) });
      o.refresh = () => { const c = o.get(); bar.style.background = c || 'transparent'; bar.classList.toggle('is-none', !c); };
      o.refresh();
      colorBtns.push(o);
      return h('span', { class: 'sl-split' }, mainB, arrow);
    }
    function pickCustomColor(initial, cb) {
      const inp = h('input', { type: 'color', class: 'sl-hidden-color', value: initial && /^#[0-9a-f]{6}$/i.test(initial) ? initial : '#000000', 'aria-label': 'Custom color', tabindex: '-1' });
      document.body.appendChild(inp);
      inp.addEventListener('change', () => { const v = inp.value.toUpperCase(); inp.remove(); cb(v); });
      setTimeout(() => { if (document.activeElement !== inp) inp.remove(); }, 120000);
      inp.click();
    }
    function openColorMenu(anchor, o) {
      const keys = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
      const names = { lt1: 'Background 1', dk1: 'Text 1', lt2: 'Background 2', dk2: 'Text 2' };
      const pick = (c) => { o.set(c); o.refresh(); };
      const items = [{ heading: 'Theme colors' }];
      keys.forEach(k => { const c = themeColor(model, k, '#000000'); items.push({ swatch: c, label: (names[k] || 'Accent ' + k.slice(6)) + ' ' + c, run: () => pick(c) }); });
      keys.forEach(k => { const c = mix(themeColor(model, k, '#000000'), '#FFFFFF', 0.6); items.push({ swatch: c, label: (names[k] || 'Accent ' + k.slice(6)) + ', lighter 60% ' + c, run: () => pick(c) }); });
      keys.forEach(k => { const c = mix(themeColor(model, k, '#000000'), '#000000', 0.25); items.push({ swatch: c, label: (names[k] || 'Accent ' + k.slice(6)) + ', darker 25% ' + c, run: () => pick(c) }); });
      items.push({ heading: 'Standard colors' });
      STD_COLORS.forEach(c => items.push({ swatch: c, label: 'Standard ' + c, run: () => pick(c) }));
      items.push({ sep: true });
      if (o.allowNone) items.push({ label: o.noneLabel || 'No color', icon: 'close', run: () => pick(null) });
      items.push({ label: 'More colors…', icon: 'fill', run: () => pickCustomColor(o.get(), pick) });
      openMenu(anchor, items, { label: o.label + ' colors' });
    }

    /* ---- popup menus ---- */
    function openMenu(anchor, items, opts) {
      closeMenu(false);
      opts = opts || {};
      const m = h('div', { class: 'sl-menu', role: 'menu', 'aria-label': opts.label || 'Menu' });
      m.style.setProperty('--accent', getComputedStyle(root).getPropertyValue('--accent'));
      let grid = null;
      const buttons = [];
      for (const it of items) {
        if (!it) continue;
        if (it.sep) { grid = null; m.appendChild(h('div', { class: 'sl-menu-sep', role: 'separator' })); continue; }
        if (it.heading) { grid = null; m.appendChild(h('div', { class: 'sl-menu-heading', role: 'presentation' }, it.heading)); continue; }
        const isCell = !!(it.swatch || it.gridIcon);
        const b = h('button', {
          type: 'button', role: it.checked != null ? 'menuitemradio' : 'menuitem', class: isCell ? 'sl-menu-cell' : 'sl-menu-item',
          title: it.title || it.label, 'aria-label': isCell ? it.label : null,
          'aria-checked': it.checked != null ? String(!!it.checked) : null, disabled: it.disabled || null, tabindex: '-1'
        });
        if (it.swatch) { const sw = h('span', { class: 'sl-swatch', 'aria-hidden': 'true' }); sw.style.background = it.swatch; b.appendChild(sw); }
        else if (it.gridIcon) b.appendChild(it.gridIcon);
        else {
          b.appendChild(h('span', { class: 'sl-menu-ico', 'aria-hidden': 'true' }, it.icon ? icon(it.icon, 16) : (it.checked ? icon('check', 16) : null)));
          b.appendChild(h('span', { class: 'sl-menu-lbl' }, it.label));
          if (it.keys) b.appendChild(h('span', { class: 'sl-menu-keys', 'aria-hidden': 'true' }, it.keys));
        }
        b.addEventListener('mousedown', ev => ev.preventDefault());
        b.addEventListener('click', () => {
          closeMenu(true);
          Promise.resolve().then(() => it.run()).catch(err => App.toast(err && err.message || String(err), 'error'));
        });
        if (isCell) {
          if (!grid) { grid = h('div', { class: 'sl-menu-grid' + (it.gridIcon ? ' icons' : '') }); m.appendChild(grid); }
          grid.appendChild(b);
        } else { grid = null; m.appendChild(b); }
        if (!it.disabled) buttons.push(b);
      }
      document.body.appendChild(m);
      let x, y;
      if (anchor && anchor.nodeType === 1) {
        const r = anchor.getBoundingClientRect();
        x = r.left; y = r.bottom + 2;
        anchor.setAttribute('aria-expanded', 'true');
      } else { x = (anchor && anchor.x) || 100; y = (anchor && anchor.y) || 100; }
      const mr = m.getBoundingClientRect();
      x = Math.max(4, Math.min(x, innerWidth - mr.width - 4));
      if (y + mr.height > innerHeight - 4) y = Math.max(4, innerHeight - mr.height - 4);
      m.style.left = x + 'px'; m.style.top = y + 'px';
      const returnFocus = opts.returnFocus || (anchor && anchor.nodeType === 1 ? anchor : document.activeElement);
      m.addEventListener('keydown', (ev) => {
        const i = buttons.indexOf(document.activeElement);
        const n = buttons.length;
        if (!n) return;
        const cols = (() => {
          const a = document.activeElement;
          if (!a || !a.classList.contains('sl-menu-cell')) return 1;
          const g = a.parentElement; const first = g.firstElementChild;
          let c = 0; for (const x2 of g.children) { if (x2.offsetTop === first.offsetTop) c++; else break; }
          return Math.max(c, 1);
        })();
        let j = null;
        if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); closeMenu(true); return; }
        if (ev.key === 'Tab') { ev.preventDefault(); closeMenu(true); return; }
        if (ev.key === 'ArrowRight') j = i + 1;
        else if (ev.key === 'ArrowLeft') j = i - 1;
        else if (ev.key === 'ArrowDown') j = i + cols;
        else if (ev.key === 'ArrowUp') j = i - cols;
        else if (ev.key === 'Home') j = 0;
        else if (ev.key === 'End') j = n - 1;
        if (j == null) return;
        ev.preventDefault();
        if (j >= n) j = cols > 1 ? Math.min(n - 1, i + 1) : 0;
        if (j < 0) j = cols > 1 ? Math.max(0, i - 1) : n - 1;
        buttons[j].focus();
      });
      const onDown = (ev) => { if (!m.contains(ev.target)) closeMenu(false); };
      setTimeout(() => document.addEventListener('pointerdown', onDown, true), 0);
      menu = { m, anchor, returnFocus, onDown };
      const first = buttons.find(b => b.getAttribute('aria-checked') === 'true') || buttons[0];
      if (first) first.focus({ preventScroll: true });
    }
    function closeMenu(refocus) {
      if (!menu) return;
      const mm = menu; menu = null;
      document.removeEventListener('pointerdown', mm.onDown, true);
      mm.m.remove();
      if (mm.anchor && mm.anchor.nodeType === 1) mm.anchor.setAttribute('aria-expanded', 'false');
      if (refocus && mm.returnFocus && document.contains(mm.returnFocus)) { try { mm.returnFocus.focus({ preventScroll: true }); } catch (_) { /* ignore */ } }
    }

    /* ---- undo / change tracking ---- */
    function snapshot() { return { slides: clone(model.slides), size: clone(model.size), layouts: clone(model.layouts), cur, selected: selected.slice() }; }
    function restore(sn) {
      model.slides = sn.slides; model.size = sn.size; model.layouts = sn.layouts;
      cur = clamp(sn.cur, 0, Math.max(model.slides.length - 1, 0));
      selected = sn.selected.slice();
    }
    function pushUndo() { undoStack.push(snapshot()); if (undoStack.length > 100) undoStack.shift(); redoStack.length = 0; }
    function undo() {
      if (editing) commitEdit();
      if (!undoStack.length) { announce('Nothing to undo'); return; }
      redoStack.push(snapshot());
      restore(undoStack.pop());
      markDirty(); refresh(true); announce('Undone');
    }
    function redo() {
      if (editing) commitEdit();
      if (!redoStack.length) { announce('Nothing to redo'); return; }
      undoStack.push(snapshot());
      restore(redoStack.pop());
      markDirty(); refresh(true); announce('Redone');
    }
    function markDirty() {
      if (!dirty) { dirty = true; App.setDirty(editor, true); }
      updateStatus();
    }
    function setClean() { dirty = false; App.setDirty(editor, false); updateStatus(); }
    function change(fn, opts) {
      opts = opts || {};
      if (editing) commitEdit();
      pushUndo();
      fn();
      markDirty();
      refresh(opts.structure);
    }
    function refresh(structure) {
      if (!model.slides.length) model.slides.push(makeSlide(model.layouts[0] ? model.layouts[0].id : ''));
      cur = clamp(cur, 0, model.slides.length - 1);
      selected = selected.filter(id => findShape(id));
      renderCanvas();
      if (structure) renderThumbs(); else scheduleThumb(cur);
      if (view === 'sorter') renderSorter();
      updateStatus(); updateRibbonState();
    }

    /* ---- canvas ---- */
    function shapeEl(id) { return slideEl.querySelector('.sl-shape[data-id="' + CSS.escape(String(id)) + '"]'); }
    function renderCanvas() {
      const s = slide();
      if (!s) return;
      slideEl = renderSlide(model, s, { index: cur, hint: true, edit: true });
      slideEl.classList.add('sl-slide-edit');
      holder.replaceChildren(slideEl, overlay);
      applyScale();
      if (document.activeElement !== notes) notes.value = s.notes || '';
      stage.setAttribute('aria-label', 'Slide ' + (cur + 1) + ' of ' + model.slides.length + ': ' + slideTitle(s));
    }
    function updateShapeEls(shapes) {
      for (const s of shapes) {
        const old = shapeEl(s.id);
        const n = renderShape(model, s, { model, index: cur, hint: true, edit: true });
        if (old) old.replaceWith(n); else slideEl.appendChild(n);
      }
    }
    function computeFit() {
      const r = stage.getBoundingClientRect();
      if (!r.width || !r.height) return scale || 0.5;
      return clamp(Math.min((r.width - 48) / W(), (r.height - 48) / H()), 0.05, 4);
    }
    function applyScale() {
      scale = zoomMode === 'fit' ? computeFit() : zoomPct / 100;
      holder.style.width = W() * scale + 'px';
      holder.style.height = H() * scale + 'px';
      slideEl.style.transform = 'scale(' + scale + ')';
      const pct = Math.round(scale * 100);
      zoomRange.value = String(clamp(pct, 10, 400));
      zoomLbl.textContent = pct + '%';
      zoomRange.setAttribute('aria-valuetext', pct + '%' + (zoomMode === 'fit' ? ' (fit)' : ''));
      renderOverlay();
    }
    function setZoom(z) {
      if (z === 'fit') zoomMode = 'fit';
      else { zoomMode = 'pct'; zoomPct = clamp(Math.round(z), 10, 400); }
      applyScale();
    }
    function stepZoom(d) {
      const p = Math.round(scale * 100);
      const n = d > 0 ? ZOOMS.find(z => z > p) : [...ZOOMS].reverse().find(z => z < p);
      if (n) setZoom(n);
    }

    const HANDLES = [['nw', 0, 0], ['n', 50, 0], ['ne', 100, 0], ['e', 100, 50], ['se', 100, 100], ['s', 50, 100], ['sw', 0, 100], ['w', 0, 50]];
    function renderOverlay() {
      overlay.textContent = '';
      const shapes = selShapes();
      const single = shapes.length === 1;
      for (const s of shapes) {
        const box = h('div', { class: 'sl-selbox' + (single ? '' : ' multi') + (editing && editing.id === String(s.id) ? ' is-editing' : '') });
        box.style.left = px(s.x) * scale + 'px';
        box.style.top = px(s.y) * scale + 'px';
        box.style.width = Math.max(px(s.w) * scale, 1) + 'px';
        box.style.height = Math.max(px(s.h) * scale, 1) + 'px';
        if (s.rot) box.style.transform = 'rotate(' + s.rot + 'deg)';
        if (single) {
          const isLine = LINE_GEOMS.has(s.geom);
          for (const [name, lx, ly] of HANDLES) {
            if (isLine && !(name === 'nw' || name === 'se')) continue;
            const hd = h('div', { class: 'sl-handle', 'data-h': name });
            hd.style.left = lx + '%'; hd.style.top = ly + '%';
            box.appendChild(hd);
          }
          if (!isLine) {
            box.appendChild(h('div', { class: 'sl-rotline' }));
            const rh = h('div', { class: 'sl-handle sl-rot', 'data-h': 'rot', title: 'Rotate (Shift snaps to 15 degrees)' });
            box.appendChild(rh);
          }
        }
        overlay.appendChild(box);
      }
      if (drag && drag.guides) {
        for (const g of drag.guides) {
          const ln = h('div', { class: 'sl-guide ' + (g.v ? 'v' : 'h') });
          if (g.v) ln.style.left = px(g.at) * scale + 'px'; else ln.style.top = px(g.at) * scale + 'px';
          overlay.appendChild(ln);
        }
      }
      if (drag && drag.marquee) {
        const r = drag.marquee;
        const mq = h('div', { class: 'sl-marquee' });
        mq.style.left = px(r.x) * scale + 'px'; mq.style.top = px(r.y) * scale + 'px';
        mq.style.width = px(r.w) * scale + 'px'; mq.style.height = px(r.h) * scale + 'px';
        overlay.appendChild(mq);
      }
    }
    function bboxOf(shapes) {
      let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
      for (const s of shapes) { x1 = Math.min(x1, s.x); y1 = Math.min(y1, s.y); x2 = Math.max(x2, s.x + s.w); y2 = Math.max(y2, s.y + s.h); }
      return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
    }
    function snapTargets(excl) {
      const xs = [0, model.size.cx / 2, model.size.cx], ys = [0, model.size.cy / 2, model.size.cy];
      for (const s of slide().shapes) {
        if (excl.has(String(s.id))) continue;
        xs.push(s.x, s.x + s.w / 2, s.x + s.w);
        ys.push(s.y, s.y + s.h / 2, s.y + s.h);
      }
      return { xs, ys };
    }
    function snapMove(bb, dx, dy, t) {
      const thr = 6 / scale * EMU_PX;
      const best = (vals, cands) => {
        let b = null;
        for (const v of vals) for (const c of cands) { const d = c - v; if (Math.abs(d) <= thr && (!b || Math.abs(d) < Math.abs(b.d))) b = { d, at: c }; }
        return b;
      };
      const bx = best([bb.x + dx, bb.x + bb.w / 2 + dx, bb.x + bb.w + dx], t.xs);
      const by = best([bb.y + dy, bb.y + bb.h / 2 + dy, bb.y + bb.h + dy], t.ys);
      const guides = [];
      if (bx) { dx += bx.d; guides.push({ v: true, at: bx.at }); }
      if (by) { dy += by.d; guides.push({ v: false, at: by.at }); }
      return { dx, dy, guides };
    }

    /* ---- pointer interaction ---- */
    function toEmu(e) {
      const r = holder.getBoundingClientRect();
      return { x: (e.clientX - r.left) / scale * EMU_PX, y: (e.clientY - r.top) / scale * EMU_PX };
    }
    function beginDrag(e, d) {
      drag = d;
      d.pointerId = e.pointerId;
      d.start = toEmu(e);
      d.sx = e.clientX; d.sy = e.clientY;
      d.moved = false;
      try { stage.setPointerCapture(e.pointerId); } catch (_) { /* ignore */ }
    }
    function endDrag(e, cancel) {
      if (!drag || e.pointerId !== drag.pointerId) return;
      const d = drag; drag = null;
      try { stage.releasePointerCapture(e.pointerId); } catch (_) { /* ignore */ }
      if (d.end) d.end(d.moved && !cancel, e);
      renderOverlay();
    }
    on(stage, 'pointermove', (e) => {
      if (!drag || e.pointerId !== drag.pointerId) return;
      if (!drag.moved && Math.hypot(e.clientX - drag.sx, e.clientY - drag.sy) < 3) return;
      if (!drag.moved) { drag.moved = true; if (drag.mutates) pushUndo(); }
      drag.move(toEmu(e), e);
    });
    on(stage, 'pointerup', (e) => endDrag(e, false));
    on(stage, 'pointercancel', (e) => endDrag(e, true));
    on(stage, 'lostpointercapture', (e) => { if (drag && e.pointerId === drag.pointerId) endDrag(e, false); });

    on(stage, 'pointerdown', (e) => {
      if (e.button !== 0 || loading) return;
      closeMenu(false);
      const t = e.target;
      if (editing && editing.tb.contains(t)) return; // caret placement inside the text being edited
      const hEl = t.closest && t.closest('.sl-handle');
      if (editing) commitEdit();
      if (hEl) { e.preventDefault(); stage.focus({ preventScroll: true }); startHandleDrag(e, hEl.dataset.h); return; }
      const shEl = t.closest && t.closest('.sl-shape[data-id]');
      e.preventDefault();
      stage.focus({ preventScroll: true });
      tabOut = false;
      if (shEl && slideEl.contains(shEl)) {
        const id = shEl.dataset.id;
        const was = selected.includes(id);
        if (e.shiftKey || e.ctrlKey || e.metaKey) {
          selected = was ? selected.filter(x => x !== id) : selected.concat(id);
          selectionChanged();
          if (was) return;
        } else if (!was) { selected = [id]; selectionChanged(); }
        startMoveDrag(e, id, was && selected.length === 1);
      } else {
        if (!e.shiftKey) { selected = []; selectionChanged(); }
        startMarquee(e);
      }
    });
    on(stage, 'dblclick', (e) => {
      if (editing && editing.tb.contains(e.target)) return;
      const shEl = e.target.closest && e.target.closest('.sl-shape[data-id]');
      if (!shEl) return;
      const s = findShape(shEl.dataset.id);
      if (s && canEditText(s)) startEdit(s.id, { point: { x: e.clientX, y: e.clientY }, target: e.target });
    });
    on(stage, 'wheel', (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      stepZoom(e.deltaY < 0 ? 1 : -1);
    }, { passive: false });

    function startMoveDrag(e, id, clickEdits) {
      const shapes = selShapes();
      const starts = shapes.map(s => ({ s, x: s.x, y: s.y }));
      const bb = bboxOf(shapes);
      const targets = snapTargets(new Set(selected));
      beginDrag(e, {
        mutates: true,
        move(p, ev) {
          let dx = p.x - drag.start.x, dy = p.y - drag.start.y;
          if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
          drag.guides = [];
          if (snapOn && !ev.altKey) { const sn = snapMove(bb, dx, dy, targets); dx = sn.dx; dy = sn.dy; drag.guides = sn.guides; }
          for (const st of starts) { st.s.x = Math.round(st.x + dx); st.s.y = Math.round(st.y + dy); }
          updateShapeEls(shapes);
          renderOverlay();
        },
        end(moved, ev) {
          if (moved) { markDirty(); scheduleThumb(cur); updateStatus(); return; }
          if (clickEdits) {
            const s = findShape(id);
            if (s && canEditText(s)) startEdit(s.id, { point: { x: ev.clientX, y: ev.clientY }, target: ev.target });
          }
        }
      });
    }
    function startMarquee(e) {
      const base = e.shiftKey ? selected.slice() : [];
      beginDrag(e, {
        mutates: false,
        move(p) {
          const r = { x: Math.min(p.x, drag.start.x), y: Math.min(p.y, drag.start.y), w: Math.abs(p.x - drag.start.x), h: Math.abs(p.y - drag.start.y) };
          drag.marquee = r;
          const inside = slide().shapes.filter(s => s.x >= r.x && s.y >= r.y && s.x + s.w <= r.x + r.w && s.y + s.h <= r.y + r.h).map(s => String(s.id));
          selected = Array.from(new Set(base.concat(inside)));
          renderOverlay();
        },
        end(moved) { if (moved) selectionChanged(); }
      });
    }
    function startHandleDrag(e, which) {
      const s = selShapes()[0];
      if (!s) return;
      const o = { x: s.x, y: s.y, w: s.w, h: s.h, rot: s.rot || 0, colW: s.colW ? s.colW.slice() : null, rowH: s.rowH ? s.rowH.slice() : null };
      const cx = o.x + o.w / 2, cy = o.y + o.h / 2;
      const done = (moved) => { if (moved) { markDirty(); scheduleThumb(cur); updateStatus(); } };
      if (which === 'rot') {
        beginDrag(e, {
          mutates: true,
          move(p, ev) {
            let a = Math.atan2(p.y - cy, p.x - cx) * 180 / Math.PI + 90;
            if (ev.shiftKey) a = Math.round(a / 15) * 15;
            else { const r = Math.round(a / 90) * 90; if (Math.abs(a - r) < 3) a = r; }
            a = ((a % 360) + 360) % 360;
            s.rot = Math.round(a * 10) / 10;
            if (!s.rot) delete s.rot;
            updateShapeEls([s]); renderOverlay();
            announce(Math.round(s.rot || 0) + ' degrees');
          },
          end: done
        });
        return;
      }
      const dirs = { nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0], se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0] };
      const [dx, dy] = dirs[which] || [1, 1];
      const th = o.rot * Math.PI / 180, cos = Math.cos(th), sin = Math.sin(th);
      const ax = -dx * o.w / 2, ay = -dy * o.h / 2;
      const isLine = LINE_GEOMS.has(s.geom);
      const min = isLine ? 0 : EMU_PT * 4;
      beginDrag(e, {
        mutates: true,
        move(p, ev) {
          const vx = p.x - cx, vy = p.y - cy;
          const qx = vx * cos + vy * sin, qy = -vx * sin + vy * cos;
          let nw = dx ? (qx - ax) * dx : o.w, nh = dy ? (qy - ay) * dy : o.h;
          nw = Math.max(nw, min); nh = Math.max(nh, min);
          const lock = (s.kind === 'image') !== !!ev.shiftKey;
          if (lock && dx && dy && o.w && o.h) { const k = Math.max(nw / o.w, nh / o.h); nw = o.w * k; nh = o.h * k; }
          const lcx = dx ? ax + dx * nw / 2 : 0, lcy = dy ? ay + dy * nh / 2 : 0;
          const ncx = cx + lcx * cos - lcy * sin, ncy = cy + lcx * sin + lcy * cos;
          s.w = Math.round(nw); s.h = Math.round(nh);
          s.x = Math.round(ncx - nw / 2); s.y = Math.round(ncy - nh / 2);
          if (o.colW && o.w) s.colW = o.colW.map(v => Math.round(v * nw / o.w));
          if (o.rowH && o.h) s.rowH = o.rowH.map(v => Math.round(v * nh / o.h));
          updateShapeEls([s]); renderOverlay();
        },
        end: done
      });
    }

    /* ---- selection ---- */
    function describeShape(s) {
      const t = shapeText(s);
      return (s.name || kindLabel(s)) + ', ' + kindLabel(s) + (t ? ': ' + t.slice(0, 100) : '');
    }
    function selectionChanged(speak) {
      renderOverlay(); updateStatus(); updateRibbonState();
      if (speak) {
        const ss = selShapes();
        if (ss.length === 1) announce(describeShape(ss[0]));
        else if (ss.length) announce(ss.length + ' objects selected');
        else announce('Nothing selected');
      }
    }
    function selectAll() { if (editing) return; selected = slide().shapes.map(s => String(s.id)); selectionChanged(true); }

    /* ---- text editing ---- */
    function canEditText(s) { return !!s && s.kind !== 'image' && s.kind !== 'opaque' && !LINE_GEOMS.has(s.geom); }
    function caretToEnd(tb) {
      const r = document.createRange();
      const lastP = tb.lastElementChild || tb;
      const tail = lastP.lastChild;
      if (tail && tail.nodeName === 'BR') r.setStartBefore(tail);
      else if (tail && tail.nodeType === 3) r.setStart(tail, tail.length);
      else if (tail) { r.selectNodeContents(tail); r.collapse(false); }
      else r.setStart(lastP, 0);
      r.collapse(true);
      return r;
    }
    function startEdit(id, opts) {
      opts = opts || {};
      if (editing) commitEdit();
      const s = findShape(id);
      if (!canEditText(s)) return;
      selected = [String(s.id)];
      let cell = null, paras;
      if (s.kind === 'table') {
        const td = opts.target && opts.target.closest ? opts.target.closest('td[data-r]') : null;
        cell = td ? [+td.dataset.r, +td.dataset.c] : (opts.cell || lastCell || [0, 0]);
        const c = s.rows && s.rows[cell[0]] && s.rows[cell[0]][cell[1]];
        if (!c || c.hMerge || c.vMerge) return;
        if (!c.paras || !c.paras.length) c.paras = [emptyPara(model, {})];
        paras = c.paras;
        lastCell = cell;
      } else {
        if (!s.paras || !s.paras.length) {
          s.paras = [emptyPara(model, { align: s.kind === 'shape' || s.fill ? 'ctr' : 'l', color: s.fill ? themeColor(model, 'lt1', '#FFFFFF') : themeColor(model, 'dk1', '#000000') })];
          if (!s.anchor) s.anchor = s.kind === 'shape' ? 'ctr' : 't';
        }
        paras = s.paras;
      }
      updateShapeEls([s]);
      const el = shapeEl(s.id);
      if (!el) return;
      const tb = cell ? el.querySelector('td[data-r="' + cell[0] + '"][data-c="' + cell[1] + '"] .sl-tbody') : el.querySelector('.sl-text > .sl-tbody');
      if (!tb) return;
      const p0 = paras[0] || {};
      pushUndo();
      editing = {
        id: String(s.id), cell, tb, el, before: JSON.stringify(paras), undoLen: undoStack.length, range: null, touched: false,
        fallback: { pp: paraMeta(p0), dr: runMeta(firstRunOf(p0) || p0.end || {}) }
      };
      el.classList.add('sl-editing');
      tb.contentEditable = 'true';
      tb.spellcheck = true;
      tb.setAttribute('role', 'textbox');
      tb.setAttribute('aria-multiline', 'true');
      tb.setAttribute('aria-label', (s.name || kindLabel(s)) + ' text' + (cell ? ', row ' + (cell[0] + 1) + ' column ' + (cell[1] + 1) : ''));
      tb.focus({ preventScroll: true });
      try { document.execCommand('styleWithCSS', false, true); } catch (_) { /* ignore */ }
      let range = null;
      if (opts.point && document.caretRangeFromPoint) {
        range = document.caretRangeFromPoint(opts.point.x, opts.point.y);
        if (range && !tb.contains(range.startContainer)) range = null;
      }
      if (!range) {
        if (opts.selectAll) { range = document.createRange(); range.selectNodeContents(tb); }
        else range = caretToEnd(tb);
      }
      const sel = window.getSelection();
      sel.removeAllRanges(); sel.addRange(range);
      editing.range = range.cloneRange();
      renderOverlay(); updateStatus(); updateRibbonState();
      announce('Editing text. Press Escape to finish.');
    }
    function normalizeEditable(tb) {
      // Wrap stray top-level inline content (left after select-all + type) in a paragraph div.
      const stray = Array.from(tb.childNodes).filter(n => !(n.nodeType === 1 && BLOCK_TAGS.test(n.tagName)) && !(n.nodeType === 3 && !n.data.trim() && tb.children.length));
      if (!stray.length) return;
      const sel = window.getSelection();
      const an = sel.rangeCount ? sel.anchorNode : null, ao = sel.anchorOffset;
      const div = h('div', { class: 'sl-p' });
      div.dataset.pp = JSON.stringify(editing.fallback.pp || {});
      div.dataset.dr = JSON.stringify(editing.fallback.dr || {});
      tb.insertBefore(div, stray[0]);
      for (const n of stray) div.appendChild(n);
      if (an && div.contains(an)) { try { sel.collapse(an, ao); } catch (_) { /* ignore */ } }
    }
    function afterEditInput() {
      if (!editing) return;
      normalizeEditable(editing.tb);
      const s = findShape(editing.id);
      refreshParas(editing.tb, s);
      if (s && !editing.cell && s.autofit === 'shape') {
        const inset = s.inset || DEFAULT_INSET;
        const need = Math.round((editing.tb.offsetHeight + px(inset[1]) + px(inset[3])) * EMU_PX);
        if (need > s.h + EMU_PT) { s.h = need; editing.el.style.height = px(need) + 'px'; editing.touched = true; renderOverlay(); }
      }
      markDirty();
      scheduleRibbonState();
    }
    function commitEdit() {
      if (!editing) return;
      const e = editing;
      editing = null;
      const s = findShape(e.id);
      let paras = null;
      try { paras = parseEditable(e.tb, e.fallback); } catch (err) { console.error(err); }
      e.tb.contentEditable = 'false';
      e.tb.removeAttribute('role'); e.tb.removeAttribute('aria-multiline'); e.tb.removeAttribute('aria-label');
      e.el.classList.remove('sl-editing');
      const sel = window.getSelection();
      if (sel.rangeCount && e.tb.contains(sel.anchorNode)) sel.removeAllRanges();
      if (!s || !paras) { renderCanvas(); return; }
      const changed = JSON.stringify(paras) !== e.before;
      if (changed) { if (e.cell) s.rows[e.cell[0]][e.cell[1]].paras = paras; else s.paras = paras; }
      if (!changed && !e.touched && undoStack.length === e.undoLen) undoStack.pop();
      if (s.kind === 'text' && !s.ph && !s.fill && !(s.line && s.line.color) && parasEmpty(s.paras) && !s.rows) {
        slide().shapes = slide().shapes.filter(x => x !== s);
        selected = selected.filter(x => x !== String(s.id));
        renderCanvas();
        markDirty();
      } else {
        updateShapeEls([s]);
        if (changed) fitAfterEdit(s);
        if (changed || e.touched) markDirty();
      }
      renderOverlay(); scheduleThumb(cur); updateRibbonState(); updateStatus();
    }
    function fitAfterEdit(s) {
      const el = shapeEl(s.id);
      if (!el) return;
      if (s.kind === 'table') {
        const t = el.querySelector('table');
        if (!t) return;
        s.rowH = s.rowH || [];
        let total = 0;
        for (let i = 0; i < t.rows.length; i++) {
          const hp = t.rows[i].offsetHeight;
          if (hp) s.rowH[i] = Math.max(s.rowH[i] || 0, Math.round(hp * EMU_PX));
          total += s.rowH[i] || 0;
        }
        if (total > s.h) { s.h = total; updateShapeEls([s]); }
        return;
      }
      const tb = el.querySelector('.sl-text > .sl-tbody');
      const box = el.querySelector('.sl-text');
      if (!tb || !box) return;
      const inset = s.inset || DEFAULT_INSET;
      if (s.autofit === 'shape') {
        const need = Math.round((tb.offsetHeight + px(inset[1]) + px(inset[3])) * EMU_PX);
        if (Math.abs(need - s.h) > EMU_PT) { s.h = need; updateShapeEls([s]); }
      } else if (s.autofit === 'norm') {
        const avail = px(s.h) - px(inset[1]) - px(inset[3]);
        let f = 100;
        box.style.setProperty('--fs', '1');
        while (tb.offsetHeight > avail + 1 && f > 25) { f -= 7.5; box.style.setProperty('--fs', String(f / 100)); }
        if (f >= 100) { delete s.fontScale; delete s.lnSpcReduction; }
        else { s.fontScale = f; s.lnSpcReduction = f < 85 ? 20 : 10; }
        updateShapeEls([s]);
      }
    }
    function restoreEditRange() {
      if (!editing) return false;
      editing.tb.focus({ preventScroll: true });
      if (editing.range) { const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(editing.range); }
      return true;
    }
    function nextCell(dir) {
      const s = findShape(editing.id);
      const [r0, c0] = editing.cell;
      commitEdit();
      const cells = [];
      (s.rows || []).forEach((row, ri) => row.forEach((c, ci) => { if (c && !c.hMerge && !c.vMerge) cells.push([ri, ci]); }));
      let i = cells.findIndex(([r, c]) => r === r0 && c === c0) + dir;
      if (i >= cells.length) { tableOp('rowBelow', [r0, c0]); i = cells.length; return startEdit(s.id, { cell: [s.rows.length - 1, 0] }); }
      if (i < 0) i = 0;
      startEdit(s.id, { cell: cells[i], selectAll: true });
    }
    on(document, 'selectionchange', () => {
      if (!editing) return;
      const sel = window.getSelection();
      if (sel.rangeCount && editing.tb.contains(sel.anchorNode)) { editing.range = sel.getRangeAt(0).cloneRange(); scheduleRibbonState(); }
    });
    on(stage, 'input', (e) => { if (editing && editing.tb.contains(e.target)) afterEditInput(); });
    on(stage, 'focusout', (e) => {
      if (!editing || !editing.tb.contains(e.target)) return;
      setTimeout(() => {
        if (!editing || destroyed) return;
        const a = document.activeElement;
        if (!a || a === document.body || editing.tb.contains(a) || ribbon.contains(a)) return;
        if (a.closest && (a.closest('.sl-menu') || a.closest('dialog'))) return;
        if (a.classList && a.classList.contains('sl-hidden-color')) return;
        commitEdit();
      }, 0);
    });

    /* ---- formatting ---- */
    function exec(cmd, val) { try { document.execCommand(cmd, false, val == null ? null : val); } catch (_) { /* ignore */ } }
    function applySizeInEdit(pt) {
      exec('styleWithCSS', false);
      exec('fontSize', '7');
      editing.tb.querySelectorAll('font[size="7"]').forEach(f => {
        const sp = document.createElement('span');
        sp.style.fontSize = pt + 'pt';
        while (f.firstChild) sp.appendChild(f.firstChild);
        f.replaceWith(sp);
        sp.querySelectorAll('[data-rr]').forEach(x => { const rr = readJSON(x.dataset.rr) || {}; rr.size = pt; x.dataset.rr = JSON.stringify(rr); x.style.fontSize = ''; });
      });
      exec('styleWithCSS', true);
    }
    function needTextTarget() {
      if (editing || textShapes().length) return true;
      announce('Select a text box, shape or table first');
      App.toast('Select a text box, shape or table first.');
      return false;
    }
    // Run-level formatting: applies to the text selection while editing, else to all text of the selected shapes.
    function setRunProp(prop, value) {
      if (!needTextTarget()) return;
      if (editing) {
        restoreEditRange();
        exec('styleWithCSS', true);
        const rp = currentRunProps() || {};
        switch (prop) {
          case 'b': exec('bold'); break;
          case 'i': exec('italic'); break;
          case 'u': exec('underline'); break;
          case 'strike': exec('strikeThrough'); break;
          case 'super': exec('superscript'); break;
          case 'sub': exec('subscript'); break;
          case 'color': exec('foreColor', value); break;
          case 'highlight': exec('hiliteColor', value || 'transparent'); break;
          case 'font': exec('fontName', value); break;
          case 'size': applySizeInEdit(value); break;
          case 'grow': applySizeInEdit(stepSize(rp.size, value)); break;
          case 'clear': exec('removeFormat'); break;
          case 'link': if (value) exec('createLink', value); else exec('unlink'); break;
        }
        afterEditInput();
        editing && (editing.range = window.getSelection().rangeCount ? window.getSelection().getRangeAt(0).cloneRange() : editing.range);
        return;
      }
      const shapes = textShapes();
      const runs = [];
      forEachRun(shapes, r => runs.push(r));
      const all = (k) => runs.length > 0 && runs.every(r => r[k]);
      change(() => {
        const toggle = { b: !all('b'), i: !all('i'), u: !all('u'), strike: !all('strike') };
        const sup = !runs.every(r => r.baseline > 0), sub = !runs.every(r => r.baseline < 0);
        for (const r of runs) {
          switch (prop) {
            case 'b': case 'i': case 'u': case 'strike': r[prop] = toggle[prop]; break;
            case 'super': if (sup) r.baseline = 30; else delete r.baseline; break;
            case 'sub': if (sub) r.baseline = -25; else delete r.baseline; break;
            case 'color': r.color = value; break;
            case 'highlight': if (value) r.highlight = value; else delete r.highlight; break;
            case 'font': r.font = value; break;
            case 'size': r.size = value; break;
            case 'grow': r.size = stepSize(r.size, value); break;
            case 'clear': ['b', 'i', 'u', 'strike', 'highlight', 'baseline', 'caps', 'spc'].forEach(k => delete r[k]); break;
            case 'link': if (value) r.link = value; else delete r.link; break;
          }
        }
        if (prop === 'size' || prop === 'grow') shapes.forEach(s => { delete s.fontScale; delete s.lnSpcReduction; });
      });
    }
    function editParaDivs() {
      const divs = Array.from(editing.tb.querySelectorAll(':scope > .sl-p'));
      const r = editing.range;
      if (!r) return divs.slice(0, 1);
      const hit = divs.filter(d => r.intersectsNode(d));
      return hit.length ? hit : divs.slice(0, 1);
    }
    // Paragraph-level formatting: fn(pp) mutates the paragraph properties.
    function setParaProp(fn) {
      if (!needTextTarget()) return;
      if (editing) {
        restoreEditRange();
        normalizeEditable(editing.tb);
        const divs = editParaDivs();
        const pps = divs.map(d => readJSON(d.dataset.pp) || {});
        fn(pps);
        divs.forEach((d, i) => { d.dataset.pp = JSON.stringify(pps[i]); });
        refreshParas(editing.tb, findShape(editing.id));
        afterEditInput();
        return;
      }
      const shapes = textShapes();
      change(() => { const ps = []; shapes.forEach(s => ps.push(...allParas(s))); fn(ps); });
    }
    function setAlign(a) { setParaProp(ps => ps.forEach(p => { p.align = a; })); }
    function toggleBullets(kind) {
      setParaProp(ps => {
        const has = (p) => kind === 'num' ? !!(p.bullet && p.bullet.auto) : !!(p.bullet && p.bullet.char);
        const off = ps.length && ps.every(has);
        for (const p of ps) {
          if (off) { p.bullet = null; continue; }
          p.bullet = kind === 'num' ? { auto: 'arabicPeriod' } : { char: '•', font: 'Arial' };
          if (!((p.indent || 0) < 0)) { p.marL = Math.max(p.marL || 0, 342900); p.indent = -342900; }
        }
      });
    }
    function shiftLevel(ps, d) {
      for (const p of ps) {
        const lvl = clamp((p.level || 0) + d, 0, 8);
        if (lvl === (p.level || 0)) continue;
        p.marL = Math.max(0, (p.marL || 0) + (lvl - (p.level || 0)) * 457200);
        p.level = lvl;
      }
    }
    function setLineSpacing(v) { setParaProp(ps => ps.forEach(p => { p.lineSpacing = Math.round(v * 100); delete p.lineSpacingPts; })); }
    function setAnchor(a) {
      const shapes = editing ? [findShape(editing.id)].filter(Boolean) : textShapes();
      if (!shapes.length) { needTextTarget(); return; }
      if (editing) {
        const s = shapes[0];
        if (editing.cell) {
          s.rows[editing.cell[0]][editing.cell[1]].anchor = a;
          const td = editing.tb.parentElement; td.style.verticalAlign = a === 'ctr' ? 'middle' : a === 'b' ? 'bottom' : 'top';
        } else {
          s.anchor = a;
          const box = editing.el.querySelector('.sl-text'); if (box) box.style.justifyContent = ANCHOR_CSS[a];
        }
        editing.touched = true; markDirty(); restoreEditRange(); updateRibbonState();
        return;
      }
      change(() => shapes.forEach(s => {
        if (s.kind === 'table') allCells(s).forEach(c => { c.anchor = a; }); else s.anchor = a;
      }));
    }
    function allCells(s) { const out = []; (s.rows || []).forEach(r => r.forEach(c => { if (c && !c.hMerge && !c.vMerge) out.push(c); })); return out; }
    async function insertLink() {
      if (!needTextTarget()) return;
      const rp = currentRunProps() || {};
      const url = await App.prompt('Link address (leave empty to remove the link)', rp.link || 'https://');
      if (url == null) { restoreEditRange(); return; }
      const u = url.trim();
      if (u && /^\s*(javascript|vbscript|data):/i.test(u)) { App.toast('That kind of link is not allowed.', 'error'); return; }
      if (editing) {
        restoreEditRange();
        if (u && window.getSelection().isCollapsed) { exec('insertText', u); const sel = window.getSelection(); const r = editing.range; if (sel.rangeCount && r) { try { const r2 = document.createRange(); r2.setStart(r.startContainer, r.startOffset); r2.setEnd(sel.anchorNode, sel.anchorOffset); sel.removeAllRanges(); sel.addRange(r2); } catch (_) { /* ignore */ } } }
      }
      setRunProp('link', u && u !== 'https://' ? u : null);
    }

    /* ---- ribbon state ---- */
    function hasDeco(n, d) {
      for (let e = n; e && e !== editing.tb; e = e.parentElement) { const t = getComputedStyle(e).textDecorationLine; if (t && t.includes(d)) return true; }
      return false;
    }
    function currentRunProps() {
      if (editing) {
        const sel = window.getSelection();
        let n = sel.rangeCount ? sel.getRangeAt(0).startContainer : null;
        if (n && n.nodeType === 3) n = n.parentElement;
        if (n && n.nodeType === 1 && n.classList.contains('sl-tbody')) n = n.querySelector('.sl-r') || n.firstElementChild || n;
        if (!n || !editing.tb.contains(n)) n = editing.tb.querySelector('.sl-r') || editing.tb.firstElementChild || editing.tb;
        const cs = getComputedStyle(n);
        const fsv = parseFloat(getComputedStyle(editing.tb).getPropertyValue('--fs')) || 1;
        const pdiv = n.closest ? n.closest('.sl-p') : null;
        let link = null;
        for (let e = n; e && e !== editing.tb; e = e.parentElement) { if (e.tagName === 'A') { link = e.getAttribute('href'); break; } const rr = e.dataset && readJSON(e.dataset.rr); if (rr && rr.link) { link = rr.link; break; } }
        return {
          b: +cs.fontWeight >= 600 || cs.fontWeight === 'bold', i: cs.fontStyle === 'italic', u: hasDeco(n, 'underline'), strike: hasDeco(n, 'line-through'),
          font: firstFamily(cs.fontFamily), size: Math.round(parseFloat(cs.fontSize) * 0.75 / fsv * 2) / 2, color: cssToHex(cs.color), link,
          para: pdiv ? readJSON(pdiv.dataset.pp) || {} : {}
        };
      }
      const shapes = textShapes();
      if (!shapes.length) return null;
      const runs = [];
      forEachRun(shapes, r => runs.push(r));
      const p = allParas(shapes[0])[0] || {};
      const r0 = runs[0] || {};
      const all = (k) => runs.length > 0 && runs.every(r => r[k]);
      return { b: all('b'), i: all('i'), u: all('u'), strike: all('strike'), font: r0.font, size: r0.size, color: r0.color, link: r0.link, para: p };
    }
    function scheduleRibbonState() {
      if (stateTimer) return;
      stateTimer = requestAnimationFrame(() => { stateTimer = 0; if (!destroyed) updateRibbonState(); });
    }
    function updateRibbonState() {
      const rp = currentRunProps();
      ['b', 'i', 'u', 'strike'].forEach(k => setPressed(controls['fmt_' + k], rp && rp[k]));
      const al = rp && rp.para ? rp.para.align || 'l' : null;
      [['l', 'alignL'], ['ctr', 'alignC'], ['r', 'alignR'], ['just', 'alignJ']].forEach(([a, id]) => setPressed(controls[id], al === a));
      const b = rp && rp.para && rp.para.bullet;
      setPressed(controls.bullets, !!(b && b.char));
      setPressed(controls.numbering, !!(b && b.auto));
      if (rp && rp.font) {
        if (![...fontSel.options].some(o => o.value === rp.font)) fontSel.appendChild(h('option', { value: rp.font, text: rp.font }));
        fontSel.value = rp.font;
      } else fontSel.value = '';
      if (document.activeElement !== sizeInput) sizeInput.value = rp && rp.size ? String(rp.size) : '';
      const hasSel = !!editing || selected.length > 0;
      needSel.forEach(b2 => { b2.disabled = !hasSel; });
      const sl = slide();
      setPressed(controls.hideSlide, sl && sl.hidden);
      setPressed(controls.viewNormal, view === 'normal');
      setPressed(controls.viewSorter, view === 'sorter');
      setPressed(controls.notesToggle, notesOn);
      setPressed(controls.snapToggle, snapOn);
      setPressed(stNotesBtn, notesOn);
      colorBtns.forEach(o => o.refresh());
    }

    /* ---- shape operations ---- */
    function requireSel() {
      if (selected.length) return true;
      announce('Select an object first'); App.toast('Select an object on the slide first.');
      return false;
    }
    function deleteShapes() {
      if (!requireSel()) return;
      const set = new Set(selected);
      const n = set.size;
      change(() => { slide().shapes = slide().shapes.filter(s => !set.has(String(s.id))); selected = []; });
      announce(n + (n === 1 ? ' object deleted' : ' objects deleted'));
    }
    function cloneShapeNew(s, dx, dy) {
      const c = clone(s);
      c.id = newId('n');
      c.x = (c.x || 0) + (dx || 0); c.y = (c.y || 0) + (dy || 0);
      if (c.children) c.children.forEach(ch => { ch.x += dx || 0; ch.y += dy || 0; });
      return c;
    }
    function duplicateShapes() {
      if (!requireSel()) return;
      const src = selShapes();
      change(() => {
        const copies = src.map(s => cloneShapeNew(s, 182880, 182880));
        slide().shapes.push(...copies);
        selected = copies.map(c => String(c.id));
      });
    }
    function copyShapes(dt) {
      if (!requireSel()) return false;
      const shapes = selShapes().map(s => resolveMedia(clone(s), model.media));
      const text = shapes.map(shapeText).filter(Boolean).join('\n') || 'DocStudio objects';
      clipboard = { shapes, text, slideId: slide().id };
      try { if (dt) dt.setData('text/plain', text); else if (navigator.clipboard) navigator.clipboard.writeText(text).catch(() => {}); } catch (_) { /* ignore */ }
      announce(shapes.length + (shapes.length === 1 ? ' object copied' : ' objects copied'));
      return true;
    }
    function pasteShapes() {
      if (!clipboard || !clipboard.shapes.length) return false;
      const sl = slide();
      const off = clipboard.slideId === sl.id ? 182880 : 0;
      change(() => {
        const copies = clipboard.shapes.map(s => cloneShapeNew(s, off, off));
        sl.shapes.push(...copies);
        selected = copies.map(c => String(c.id));
      });
      if (off) clipboard.shapes = clipboard.shapes.map(s => cloneShapeNew(s, off, off));
      return true;
    }
    async function pasteFromRibbon() {
      if (editing) {
        restoreEditRange();
        try { const t = await navigator.clipboard.readText(); restoreEditRange(); exec('insertText', t); afterEditInput(); }
        catch (_) { App.toast('Press Ctrl+V to paste text.'); }
        return;
      }
      if (pasteShapes()) return;
      try { const t = await navigator.clipboard.readText(); if (t) insertTextBox(t); else App.toast('The clipboard is empty.'); }
      catch (_) { App.toast('Press Ctrl+V to paste.'); }
    }
    function reorder(mode) {
      if (!requireSel()) return;
      const set = new Set(selected);
      const isSel = (s) => set.has(String(s.id));
      change(() => {
        let arr = slide().shapes.slice();
        if (mode === 'front') arr = arr.filter(s => !isSel(s)).concat(arr.filter(isSel));
        else if (mode === 'back') arr = arr.filter(isSel).concat(arr.filter(s => !isSel(s)));
        else if (mode === 'forward') { for (let i = arr.length - 2; i >= 0; i--) if (isSel(arr[i]) && !isSel(arr[i + 1])) [arr[i], arr[i + 1]] = [arr[i + 1], arr[i]]; }
        else if (mode === 'backward') { for (let i = 1; i < arr.length; i++) if (isSel(arr[i]) && !isSel(arr[i - 1])) [arr[i], arr[i - 1]] = [arr[i - 1], arr[i]]; }
        slide().shapes = arr;
      });
    }
    function alignShapes(how) {
      if (!requireSel()) return;
      const ss = selShapes();
      const ref = ss.length === 1 ? { x: 0, y: 0, w: model.size.cx, h: model.size.cy } : bboxOf(ss);
      change(() => ss.forEach(s => {
        if (how === 'left') s.x = ref.x;
        else if (how === 'center') s.x = Math.round(ref.x + (ref.w - s.w) / 2);
        else if (how === 'right') s.x = ref.x + ref.w - s.w;
        else if (how === 'top') s.y = ref.y;
        else if (how === 'middle') s.y = Math.round(ref.y + (ref.h - s.h) / 2);
        else if (how === 'bottom') s.y = ref.y + ref.h - s.h;
      }));
    }
    function distribute(axis) {
      const ss = selShapes();
      if (ss.length < 3) { App.toast('Select three or more objects to distribute them.'); return; }
      const k = axis === 'h' ? 'x' : 'y', d = axis === 'h' ? 'w' : 'h';
      change(() => {
        const sorted = ss.slice().sort((a, b) => a[k] - b[k]);
        const first = sorted[0];
        const end = Math.max(...sorted.map(s => s[k] + s[d]));
        const total = sorted.reduce((a, s) => a + s[d], 0);
        const gap = (end - first[k] - total) / (sorted.length - 1);
        let pos = first[k];
        for (const s of sorted) { s[k] = Math.round(pos); pos += s[d] + gap; }
      });
    }
    function rotateBy(deg) {
      if (!requireSel()) return;
      change(() => selShapes().forEach(s => { const r = (((s.rot || 0) + deg) % 360 + 360) % 360; if (r) s.rot = r; else delete s.rot; }));
    }
    function flip(axis) {
      if (!requireSel()) return;
      change(() => selShapes().forEach(s => { const k = axis === 'h' ? 'flipH' : 'flipV'; s[k] = !s[k]; }));
    }
    function nudge(dx, dy) {
      const ss = selShapes();
      if (!ss.length) return;
      change(() => ss.forEach(s => { s.x = Math.round(s.x + dx); s.y = Math.round(s.y + dy); if (s.children) s.children.forEach(c => { c.x += dx; c.y += dy; }); }));
    }
    async function sizeDialog() {
      if (!requireSel()) return;
      const ss = selShapes();
      const s = ss[0];
      const inch = (v) => String(Math.round(v / EMU_IN * 100) / 100);
      const field = (id, label, value, step, min) => {
        const inp = h('input', { type: 'number', id: uid + '-sz-' + id, value, step: step || '0.01', min: min == null ? null : String(min), class: 'sl-num' });
        return [h('label', { for: inp.id }, label), inp];
      };
      const f = {
        w: field('w', 'Width (in)', inch(s.w), '0.01', 0), h: field('h', 'Height (in)', inch(s.h), '0.01', 0),
        x: field('x', 'Left (in)', inch(s.x)), y: field('y', 'Top (in)', inch(s.y)), r: field('r', 'Rotation (degrees)', String(s.rot || 0), '1')
      };
      const lock = h('input', { type: 'checkbox', id: uid + '-sz-lock' });
      if (s.kind === 'image') lock.checked = true;
      const body = h('div', { class: 'sl-size-form' }, f.w, f.h, f.x, f.y, f.r, h('span'), h('label', { class: 'sl-check' }, lock, ' Lock aspect ratio'));
      const ratio = s.h ? s.w / s.h : 1;
      f.w[1].addEventListener('input', () => { if (lock.checked && ratio) f.h[1].value = (parseFloat(f.w[1].value) / ratio).toFixed(2); });
      f.h[1].addEventListener('input', () => { if (lock.checked && ratio) f.w[1].value = (parseFloat(f.h[1].value) * ratio).toFixed(2); });
      const res = await App.modal({ title: 'Size and position', body, buttons: [{ label: 'Cancel', value: null }, { label: 'OK', value: 'ok', primary: true }] });
      if (res !== 'ok') { stage.focus(); return; }
      const num = (inp, dflt) => { const v = parseFloat(inp[1].value); return isFinite(v) ? v : dflt; };
      change(() => {
        const nw = Math.max(0, Math.round(num(f.w, s.w / EMU_IN) * EMU_IN)), nh = Math.max(0, Math.round(num(f.h, s.h / EMU_IN) * EMU_IN));
        if (s.colW && s.w) s.colW = s.colW.map(v => Math.round(v * nw / s.w));
        if (s.rowH && s.h) s.rowH = s.rowH.map(v => Math.round(v * nh / s.h));
        s.w = nw; s.h = nh;
        s.x = Math.round(num(f.x, s.x / EMU_IN) * EMU_IN); s.y = Math.round(num(f.y, s.y / EMU_IN) * EMU_IN);
        const r = ((num(f.r, 0) % 360) + 360) % 360; if (r) s.rot = r; else delete s.rot;
      });
      stage.focus();
    }
    function setFill(c) {
      if (!requireSel()) return;
      last.fill = c;
      change(() => selShapes().forEach(s => {
        if (s.kind === 'table') allCells(s).forEach(cell => { cell.fill = c; });
        else if (s.kind !== 'image' && !LINE_GEOMS.has(s.geom)) { s.fill = c; delete s.fillImage; }
      }));
    }
    function setLine(patch) {
      if (!requireSel()) return;
      change(() => selShapes().forEach(s => {
        if (s.kind === 'table' || s.kind === 'opaque') return;
        if (patch === null) { s.line = null; return; }
        s.line = Object.assign({ color: (s.line && s.line.color) || last.line || themeColor(model, 'accent1', '#4472C4'), width: (s.line && s.line.width) || 12700, dash: (s.line && s.line.dash) || 'solid' }, s.line || {}, patch);
        if (!s.line.color) s.line.color = themeColor(model, 'accent1', '#4472C4');
      }));
      if (patch && patch.color) last.line = patch.color;
    }
    function setGeom(geom) {
      if (!requireSel()) return;
      change(() => selShapes().forEach(s => { if (s.kind === 'shape' || s.kind === 'text') s.geom = geom; }));
    }

    /* ---- insert ---- */
    function centerBox(w, hh) { return { x: Math.round((model.size.cx - w) / 2), y: Math.round((model.size.cy - hh) / 2), w: Math.round(w), h: Math.round(hh) }; }
    function insertShapeObj(s, edit) {
      change(() => { slide().shapes.push(s); selected = [String(s.id)]; });
      if (edit) startEdit(s.id); else stage.focus({ preventScroll: true });
      announce('Inserted ' + (s.name || kindLabel(s)));
    }
    function insertTextBox(text) {
      const minor = themeOf(model).minorFont || 'Calibri', color = themeColor(model, 'dk1', '#000000');
      const lines = text ? String(text).replace(/\r\n?/g, '\n').split('\n') : null;
      const paras = lines ? lines.map(l => ({ align: 'l', level: 0, bullet: null, runs: l ? [{ text: l, size: 18, font: minor, color }] : [], end: { size: 18, font: minor, color } })) : [emptyPara(model, {})];
      const hh = Math.round(18 * 1.2 * EMU_PT * paras.length + 91440);
      const s = Object.assign({ id: newId('n'), name: 'TextBox ' + idSeq, kind: 'text', ph: null, rot: 0, geom: 'rect', fill: null, line: null,
        anchor: 't', inset: DEFAULT_INSET.slice(), autofit: 'shape', wrap: 'square', paras }, centerBox(model.size.cx * 0.4, hh));
      insertShapeObj(s, !text);
    }
    function insertGeom(geom, label) {
      const isLine = LINE_GEOMS.has(geom);
      const acc = themeColor(model, 'accent1', '#4472C4');
      let w = EMU_IN * 1.5, hh = EMU_IN * 1.5;
      if (isLine) { w = EMU_IN * 2; hh = 0; }
      else if (geom === 'rightArrow' || geom === 'leftArrow' || geom === 'chevron' || geom === 'homePlate') { w = EMU_IN * 2; hh = EMU_IN; }
      else if (geom === 'upArrow' || geom === 'downArrow') { w = EMU_IN; hh = EMU_IN * 2; }
      const s = Object.assign({ id: newId('n'), name: (label || geom) + ' ' + idSeq, kind: 'shape', ph: null, rot: 0, geom,
        fill: isLine ? null : (last.fill || acc), line: { color: isLine ? acc : mix(acc, '#000000', 0.25), width: isLine ? 19050 : 12700, dash: 'solid' },
        anchor: 'ctr', inset: DEFAULT_INSET.slice(), autofit: 'none' }, centerBox(w, hh));
      if (!isLine) s.paras = [emptyPara(model, { align: 'ctr', color: themeColor(model, 'lt1', '#FFFFFF') })];
      insertShapeObj(s, false);
    }
    async function insertImageBytes(bytes, mime, name) {
      const src = await bytesToDataUrl(bytes, mime);
      const sz = await loadImageSize(src);
      let w = sz.w * EMU_PX, hh = sz.h * EMU_PX;
      const k = Math.min(1, model.size.cx * 0.8 / w, model.size.cy * 0.8 / hh);
      w *= k; hh *= k;
      insertShapeObj(Object.assign({ id: newId('n'), name: 'Picture ' + idSeq, descr: name || '', kind: 'image', ph: null, rot: 0, geom: 'rect',
        fill: null, line: null, src, crop: null }, centerBox(w, hh)), false);
    }
    async function insertPicture() {
      let p = await App.pickFile({ mode: 'open', extensions: ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg'] });
      if (!p) return;
      if (Array.isArray(p)) p = p[0];
      const bytes = await App.readFile(p);
      await insertImageBytes(bytes, imageMime(p), baseName(p));
    }
    function makeCell(ri, ref) {
      const acc = themeColor(model, 'accent1', '#4472C4');
      const border = { color: '#FFFFFF', width: 12700, dash: 'solid' };
      if (ref) {
        const c = clone(ref);
        delete c.hMerge; delete c.vMerge; c.span = 1; c.rowSpan = 1;
        c.paras = (c.paras && c.paras.length ? c.paras.slice(0, 1) : [emptyPara(model, {})]).map(p => Object.assign(paraMeta(p), { runs: [], end: runMeta(firstRunOf(p) || p.end || {}) }));
        return c;
      }
      return { paras: [emptyPara(model, { color: ri === 0 ? themeColor(model, 'lt1', '#FFFFFF') : themeColor(model, 'dk1', '#000000'), b: ri === 0 })],
        fill: ri === 0 ? acc : mix(acc, '#FFFFFF', ri % 2 ? 0.8 : 0.9), span: 1, rowSpan: 1, borders: { l: border, r: border, t: border, b: border } };
    }
    async function insertTable() {
      const v = await App.prompt('Table size: columns x rows (for example 4 x 3)', '3 x 3');
      if (!v) return;
      const m = /(\d+)\s*[x×*,; ]\s*(\d+)/i.exec(v.trim());
      if (!m) { App.toast('Enter the size as columns x rows, for example 4 x 3.', 'error'); return; }
      const cols = clamp(+m[1], 1, 30), rows = clamp(+m[2], 1, 60);
      const w = model.size.cx * 0.8, cw = Math.round(w / cols), rh = 370840;
      const grid = Array.from({ length: rows }, (_, ri) => Array.from({ length: cols }, () => makeCell(ri)));
      insertShapeObj(Object.assign({ id: newId('n'), name: 'Table ' + idSeq, kind: 'table', ph: null, rot: 0, geom: 'rect', fill: null, line: null,
        rows: grid, colW: Array(cols).fill(cw), rowH: Array(rows).fill(rh) }, centerBox(cw * cols, rh * rows)), false);
    }
    function tableTarget() {
      const s = editing ? findShape(editing.id) : selShapes()[0];
      if (!s || s.kind !== 'table') { App.toast('Select a table first.'); return null; }
      return { s, cell: editing && editing.cell ? editing.cell : (lastCell || [0, 0]) };
    }
    function tableOp(op, cellArg) {
      const t = cellArg ? { s: findShape(editing ? editing.id : selected[0]) || selShapes()[0], cell: cellArg } : tableTarget();
      if (!t || !t.s) return;
      const s = t.s;
      const [r, c] = [clamp(t.cell[0], 0, s.rows.length - 1), clamp(t.cell[1], 0, (s.rows[0] || []).length - 1)];
      change(() => {
        const nCols = (s.rows[0] || []).length;
        if (op === 'rowAbove' || op === 'rowBelow') {
          const at = op === 'rowAbove' ? r : r + 1;
          const ref = s.rows[r];
          s.rows.splice(at, 0, ref.map(cell => makeCell(at, cell)));
          s.rowH = s.rowH || [];
          s.rowH.splice(at, 0, s.rowH[r] || 370840);
          s.h += s.rowH[at];
          lastCell = [at, c];
        } else if (op === 'colLeft' || op === 'colRight') {
          const at = op === 'colLeft' ? c : c + 1;
          s.rows.forEach(rw => rw.splice(at, 0, makeCell(0, rw[c])));
          s.colW = s.colW || Array(nCols).fill(Math.round(s.w / nCols));
          const nw = s.colW[c] || Math.round(s.w / nCols);
          s.colW.splice(at, 0, nw);
          s.w += nw;
          lastCell = [r, at];
        } else if (op === 'delRow') {
          if (s.rows.length <= 1) { slide().shapes = slide().shapes.filter(x => x !== s); selected = []; return; }
          s.rows.splice(r, 1);
          if (s.rowH) { s.h -= s.rowH[r] || 0; s.rowH.splice(r, 1); }
          lastCell = [Math.min(r, s.rows.length - 1), c];
        } else if (op === 'delCol') {
          if (nCols <= 1) { slide().shapes = slide().shapes.filter(x => x !== s); selected = []; return; }
          s.rows.forEach(rw => rw.splice(c, 1));
          if (s.colW) { s.w -= s.colW[c] || 0; s.colW.splice(c, 1); }
          lastCell = [r, Math.min(c, nCols - 2)];
        }
        // drop merge flags that no longer have an origin cell
        s.rows.forEach(rw => rw.forEach(cell => { if (cell) { delete cell.hMerge; delete cell.vMerge; cell.span = 1; cell.rowSpan = 1; } }));
      });
    }

    /* ---- slides ---- */
    function layoutById(id) { return (model.layouts || []).find(l => l.id === id) || null; }
    function emptied(s) {
      delete s.fontScale; delete s.lnSpcReduction;
      if (s.paras && s.paras.length) {
        const p0 = s.paras[0];
        s.paras = [Object.assign(paraMeta(p0), { runs: [], end: runMeta(firstRunOf(p0) || p0.end || {}) })];
      }
      return s;
    }
    function layoutTemplate(layoutId) {
      const lay = layoutById(layoutId);
      if (lay && (lay.placeholders || lay.decor)) {
        return { placeholders: (lay.placeholders || []).map(clone), decor: (lay.decor || []).map(clone), bg: lay.bg ? clone(lay.bg) : null };
      }
      const same = model.slides.find(s => s.layout === layoutId && s.shapes.some(x => x.ph));
      if (same) {
        return { placeholders: same.shapes.filter(s => s.ph && !['dt', 'ftr', 'sldNum'].includes(s.ph)).map(s => emptied(clone(s))), decor: clone(same.decor || []), bg: same.bg ? clone(same.bg) : null };
      }
      return { placeholders: defaultPlaceholders(model, lay ? lay.name : 'Title and Content'), decor: [], bg: null };
    }
    function prepPlaceholders(list) {
      return list.map(s => {
        s = clone(s);
        s.id = newId('n');
        delete s.xml;
        if (!s.paras || !s.paras.length) s.paras = [emptyPara(model, { size: TITLE_PH.has(s.ph) ? 44 : 18, font: TITLE_PH.has(s.ph) ? themeOf(model).majorFont : null })];
        else s.paras = s.paras.slice(0, 1).map(p => Object.assign(paraMeta(p), { runs: [], end: p.end || runMeta(firstRunOf(p) || {}) }));
        return s;
      });
    }
    function makeSlide(layoutId) {
      const t = layoutTemplate(layoutId);
      return { id: newId('s'), src: null, layout: layoutId || '', hidden: false, bg: t.bg, notes: '', decor: t.decor, shapes: prepPlaceholders(t.placeholders) };
    }
    function defaultNextLayout() {
      const base = slide();
      const lay = base ? layoutById(base.layout) : null;
      if (lay && /title slide/i.test(lay.name || '')) {
        const tc = (model.layouts || []).find(l => /title and content/i.test(l.name || ''));
        if (tc) return tc.id;
      }
      if (base) return base.layout;
      return (model.layouts.find(l => /title and content/i.test(l.name || '')) || model.layouts[0] || {}).id || '';
    }
    function newSlide(layoutId, at) {
      const lid = layoutId != null ? layoutId : defaultNextLayout();
      const s = makeSlide(lid);
      const pos = at != null ? at : cur + 1;
      change(() => { model.slides.splice(pos, 0, s); cur = pos; selected = []; }, { structure: true });
      announce('New slide ' + (pos + 1) + ' of ' + model.slides.length);
      stage.focus({ preventScroll: true });
    }
    function duplicateSlide(i) {
      i = i == null ? cur : i;
      const c = clone(model.slides[i]);
      c.id = newId('s'); c.src = null;
      c.shapes.forEach(s => { s.id = newId('n'); });
      change(() => { model.slides.splice(i + 1, 0, c); cur = i + 1; selected = []; }, { structure: true });
      announce('Slide duplicated');
    }
    function deleteSlide(i) {
      i = i == null ? cur : i;
      if (model.slides.length <= 1) { App.toast('A presentation needs at least one slide.'); return; }
      change(() => { model.slides.splice(i, 1); if (cur >= i && cur > 0) cur--; selected = []; }, { structure: true });
      announce('Slide deleted. Slide ' + (cur + 1) + ' of ' + model.slides.length);
    }
    function moveSlide(from, to) {
      to = clamp(to, 0, model.slides.length - 1);
      if (from === to) return;
      change(() => { const [s] = model.slides.splice(from, 1); model.slides.splice(to, 0, s); cur = to; selected = []; }, { structure: true });
      announce('Moved to position ' + (to + 1));
    }
    function toggleHidden(i) {
      i = i == null ? cur : i;
      change(() => { model.slides[i].hidden = !model.slides[i].hidden; }, { structure: true });
      announce(model.slides[i].hidden ? 'Slide hidden in slide show' : 'Slide shown in slide show');
    }
    const phGroup = (ph) => TITLE_PH.has(ph) ? 'title' : ph === 'subTitle' || ph === 'body' || ph === 'obj' || !ph ? 'body' : ph;
    function applyLayout(layoutId) {
      const sl = slide();
      const t = layoutTemplate(layoutId);
      change(() => {
        const used = new Set();
        for (const tp of prepPlaceholders(t.placeholders)) {
          const m = sl.shapes.find(s => s.ph && !used.has(s) && phGroup(s.ph) === phGroup(tp.ph));
          if (m) {
            used.add(m);
            Object.assign(m, { x: tp.x, y: tp.y, w: tp.w, h: tp.h, ph: tp.ph });
            if (tp.phIdx != null) m.phIdx = tp.phIdx; else delete m.phIdx;
          } else { sl.shapes.push(tp); used.add(tp); }
        }
        sl.layout = layoutId;
        sl.decor = t.decor;
      }, { structure: true });
    }
    function resetSlide() {
      const sl = slide();
      const t = layoutTemplate(sl.layout);
      change(() => {
        for (const s of sl.shapes) {
          if (!s.ph) continue;
          const tp = t.placeholders.find(p => (s.phIdx != null && p.phIdx === s.phIdx) || phGroup(p.ph) === phGroup(s.ph));
          const g = tp || origGeom.get(sl.id + '/' + s.id);
          if (g) { s.x = g.x; s.y = g.y; s.w = g.w; s.h = g.h; delete s.rot; }
        }
      });
      announce('Slide reset');
    }
    function setSlideSize(cx, cy) {
      if (model.size.cx === cx && model.size.cy === cy) return;
      change(() => {
        const k = Math.min(cx / model.size.cx, cy / model.size.cy);
        const ox = (cx - model.size.cx * k) / 2, oy = (cy - model.size.cy * k) / 2;
        const sc = (s) => {
          s.x = Math.round(s.x * k + ox); s.y = Math.round(s.y * k + oy); s.w = Math.round(s.w * k); s.h = Math.round(s.h * k);
          if (s.colW) s.colW = s.colW.map(v => Math.round(v * k));
          if (s.rowH) s.rowH = s.rowH.map(v => Math.round(v * k));
          forEachRun([s], r => { if (r.size) r.size = Math.max(1, Math.round(r.size * k * 2) / 2); });
          if (s.children) s.children.forEach(sc);
        };
        for (const sl of model.slides) { sl.shapes.forEach(sc); (sl.decor || []).forEach(sc); }
        for (const l of model.layouts) { (l.placeholders || []).forEach(sc); (l.decor || []).forEach(sc); }
        model.size = { cx, cy };
      }, { structure: true });
      setZoom('fit');
    }
    function setBackground(c, all) {
      change(() => { (all ? model.slides : [slide()]).forEach(s => { s.bg = c ? { color: c } : null; }); }, { structure: !!all });
    }
    function bgToAll() {
      const bg = slide().bg ? clone(slide().bg) : null;
      change(() => model.slides.forEach(s => { s.bg = bg ? clone(bg) : null; }), { structure: true });
      announce('Background applied to all slides');
    }
    function resetBackground() {
      change(() => { const s = slide(); s.bg = origBg.has(s.id) ? clone(origBg.get(s.id)) : null; });
    }

    /* ---- ribbon ---- */
    const tabsEl = h('div', { class: 'ribbon-tabs', role: 'tablist', 'aria-label': 'Presentation ribbon' });
    const ribbon = h('div', { class: 'ribbon sl-ribbon' }, tabsEl);
    const rtabs = [];
    function addTab(id, label, content) {
      const tabId = uid + '-tab-' + id, panelId = uid + '-panel-' + id;
      const tab = h('button', { type: 'button', role: 'tab', id: tabId, 'aria-controls': panelId, 'aria-selected': 'false', tabindex: '-1', class: 'ribbon-tab' }, label);
      const toolbar = h('div', { role: 'toolbar', 'aria-label': label + ' tools', class: 'sl-toolbar' }, content);
      const panel = h('div', { role: 'tabpanel', id: panelId, 'aria-labelledby': tabId, class: 'ribbon-panel', hidden: true }, toolbar);
      tab.addEventListener('mousedown', e => e.preventDefault());
      tab.addEventListener('click', () => selectTab(id, true));
      toolbar.addEventListener('keydown', toolbarArrows);
      tabsEl.appendChild(tab);
      ribbon.appendChild(panel);
      rtabs.push({ id, tab, panel });
    }
    function selectTab(id, focusTab) {
      for (const t of rtabs) {
        const onT = t.id === id;
        t.tab.setAttribute('aria-selected', String(onT));
        t.tab.tabIndex = onT ? 0 : -1;
        t.panel.hidden = !onT;
        if (onT && focusTab) t.tab.focus();
      }
    }
    const activeTab = () => rtabs.find(t => t.tab.getAttribute('aria-selected') === 'true') || rtabs[0];
    tabsEl.addEventListener('keydown', e => {
      let i = rtabs.findIndex(t => t.tab === document.activeElement);
      if (i < 0) return;
      if (e.key === 'ArrowRight') i = (i + 1) % rtabs.length;
      else if (e.key === 'ArrowLeft') i = (i - 1 + rtabs.length) % rtabs.length;
      else if (e.key === 'Home') i = 0;
      else if (e.key === 'End') i = rtabs.length - 1;
      else if (e.key === 'ArrowDown') { e.preventDefault(); const f = focusables(rtabs[i].panel)[0]; if (f) f.focus(); return; }
      else return;
      e.preventDefault();
      selectTab(rtabs[i].id, true);
    });
    function focusables(el) { return Array.from(el.querySelectorAll('button,select,input')).filter(x => !x.disabled && x.offsetParent !== null); }
    function toolbarArrows(e) {
      if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
      if (e.target.tagName === 'SELECT' || (e.target.tagName === 'INPUT' && e.target.type !== 'checkbox')) return;
      const items = focusables(e.currentTarget);
      const i = items.indexOf(document.activeElement);
      if (i < 0) return;
      e.preventDefault();
      items[(i + (e.key === 'ArrowRight' ? 1 : items.length - 1)) % items.length].focus();
    }

    const fontSel = h('select', { class: 'tool-select sl-font', 'aria-label': 'Font', title: 'Font' });
    function fillFonts() {
      fontSel.textContent = '';
      fontSel.appendChild(h('option', { value: '', text: '' }));
      const th = themeOf(model);
      const seen = new Set();
      if (th.majorFont) { fontSel.appendChild(h('option', { value: th.majorFont, text: th.majorFont + ' (Headings)' })); seen.add(th.majorFont); }
      if (th.minorFont && !seen.has(th.minorFont)) { fontSel.appendChild(h('option', { value: th.minorFont, text: th.minorFont + ' (Body)' })); seen.add(th.minorFont); }
      FONTS.filter(f => !seen.has(f)).forEach(f => fontSel.appendChild(h('option', { value: f, text: f })));
    }
    fillFonts();
    fontSel.addEventListener('change', () => { if (fontSel.value) setRunProp('font', fontSel.value); });
    const sizeList = h('datalist', { id: uid + '-sizes' }, SIZES.map(s => h('option', { value: String(s) })));
    const sizeInput = h('input', { type: 'text', inputmode: 'decimal', class: 'tool-input sl-size', list: uid + '-sizes', 'aria-label': 'Font size', title: 'Font size (Ctrl+] / Ctrl+[)' });
    sizeInput.addEventListener('change', () => {
      const v = parseFloat(sizeInput.value);
      if (isFinite(v) && v >= 1 && v <= 400) setRunProp('size', Math.round(v * 2) / 2);
    });
    sizeInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); sizeInput.dispatchEvent(new Event('change')); restoreEditRange() || stage.focus(); } });

    const fontColor = { label: 'Font color', icon: 'fontColor', get: () => last.color, set: (c) => { last.color = c || last.color; setRunProp('color', c); }, apply: (c) => setRunProp('color', c) };
    const hiColor = { label: 'Text highlight color', icon: 'highlight', allowNone: true, noneLabel: 'No highlight', get: () => last.highlight, set: (c) => { if (c) last.highlight = c; setRunProp('highlight', c); }, apply: (c) => setRunProp('highlight', c) };
    const fillColor = { label: 'Shape fill', icon: 'fill', allowNone: true, noneLabel: 'No fill', needsSel: true, get: () => last.fill || themeColor(model, 'accent1', '#4472C4'), set: (c) => setFill(c), apply: (c) => setFill(c) };
    const lineColor = { label: 'Shape outline', icon: 'outline', allowNone: true, noneLabel: 'No outline', needsSel: true, get: () => last.line || themeColor(model, 'accent1', '#4472C4'), set: (c) => setLine(c ? { color: c } : null), apply: (c) => setLine({ color: c }) };
    const bgColor = { label: 'Background color', icon: 'background', allowNone: true, noneLabel: 'Original background', get: () => (slide() && slide().bg && slide().bg.color) || themeColor(model, 'lt1', '#FFFFFF'), set: (c) => { if (c) setBackground(c, false); else resetBackground(); }, apply: (c) => setBackground(c, false) };

    function shapeMenu(anchor, run) {
      openMenu(anchor, SHAPES.map(([g, label]) => ({ gridIcon: geomIcon(g), label, run: () => run(g, label) })), { label: 'Shapes' });
    }
    function layoutMenu(anchor, run, currentId) {
      const ls = model.layouts || [];
      const items = ls.length ? ls.map(l => ({ label: l.name || l.id, checked: currentId !== undefined ? l.id === currentId : undefined, run: () => run(l.id) }))
        : ['Title Slide', 'Title and Content', 'Section Header', 'Title Only', 'Blank'].map(n => ({ label: n, run: () => run(n) }));
      openMenu(anchor, items, { label: 'Layouts' });
    }
    function lineWeightMenu(anchor) {
      const curW = (selShapes()[0] && selShapes()[0].line && selShapes()[0].line.width) || 0;
      const items = [{ heading: 'Weight' }];
      LINE_WIDTHS.forEach(w => items.push({ label: w + ' pt', checked: Math.abs(curW - w * EMU_PT) < 100, run: () => setLine({ width: Math.round(w * EMU_PT) }) }));
      items.push({ heading: 'Dashes' });
      [['solid', 'Solid'], ['dash', 'Dash'], ['dot', 'Round dot'], ['lgDash', 'Long dash'], ['dashDot', 'Dash dot']].forEach(([d, l]) => items.push({ label: l, run: () => setLine({ dash: d }) }));
      items.push({ sep: true }, { label: 'No outline', icon: 'close', run: () => setLine(null) });
      openMenu(anchor, items, { label: 'Outline weight and dashes' });
    }
    function lineSpacingMenu(anchor) {
      const p = (currentRunProps() || {}).para || {};
      const curLs = p.lineSpacing || 100;
      openMenu(anchor, [1, 1.15, 1.5, 2, 2.5, 3].map(v => ({ label: String(v.toFixed(v % 1 ? 2 : 1)).replace(/0$/, ''), checked: Math.abs(curLs - v * 100) < 1, run: () => setLineSpacing(v) })), { label: 'Line spacing' });
    }
    function alignTextMenu(anchor) {
      openMenu(anchor, [
        { label: 'Top', icon: 'valignTop', run: () => setAnchor('t') },
        { label: 'Middle', icon: 'valignMiddle', run: () => setAnchor('ctr') },
        { label: 'Bottom', icon: 'valignBottom', run: () => setAnchor('b') }
      ], { label: 'Align text vertically' });
    }
    function arrangeItems() {
      return [
        { heading: 'Order objects' },
        { label: 'Bring to front', icon: 'bringFront', keys: 'Ctrl+Shift+]', run: () => reorder('front') },
        { label: 'Bring forward', icon: 'bringForward', keys: 'Ctrl+]', run: () => reorder('forward') },
        { label: 'Send backward', icon: 'sendBackward', keys: 'Ctrl+[', run: () => reorder('backward') },
        { label: 'Send to back', icon: 'sendBack', keys: 'Ctrl+Shift+[', run: () => reorder('back') },
        { heading: 'Position objects' },
        { label: 'Align left', icon: 'alignObjLeft', run: () => alignShapes('left') },
        { label: 'Align center', icon: 'alignObjCenter', run: () => alignShapes('center') },
        { label: 'Align right', icon: 'alignObjRight', run: () => alignShapes('right') },
        { label: 'Align top', icon: 'alignObjTop', run: () => alignShapes('top') },
        { label: 'Align middle', icon: 'alignObjMiddle', run: () => alignShapes('middle') },
        { label: 'Align bottom', icon: 'alignObjBottom', run: () => alignShapes('bottom') },
        { label: 'Distribute horizontally', icon: 'distH', run: () => distribute('h') },
        { label: 'Distribute vertically', icon: 'distV', run: () => distribute('v') },
        { sep: true },
        { label: 'Rotate right 90°', icon: 'rotateRight', run: () => rotateBy(90) },
        { label: 'Rotate left 90°', icon: 'rotateLeft', run: () => rotateBy(-90) },
        { label: 'Flip horizontal', run: () => flip('h') },
        { label: 'Flip vertical', run: () => flip('v') },
        { sep: true },
        { label: 'Size and position…', icon: 'size', run: () => sizeDialog() }
      ];
    }
    function tableItems() {
      return [
        { label: 'Insert row above', run: () => tableOp('rowAbove') },
        { label: 'Insert row below', run: () => tableOp('rowBelow') },
        { label: 'Insert column left', run: () => tableOp('colLeft') },
        { label: 'Insert column right', run: () => tableOp('colRight') },
        { sep: true },
        { label: 'Delete row', icon: 'delete', run: () => tableOp('delRow') },
        { label: 'Delete column', icon: 'delete', run: () => tableOp('delCol') }
      ];
    }

    addTab('file', 'File', [
      group('Save',
        btn({ icon: 'save', label: 'Save', keys: 'Ctrl+S', big: true, errPrefix: 'Save failed: ', run: () => editor.save() }),
        btn({ icon: 'saveAs', label: 'Save As', keys: 'Ctrl+Shift+S', big: true, errPrefix: 'Save failed: ', run: () => editor.saveAs() })),
      group('Export',
        btn({ icon: 'print', label: 'Print / Save as PDF', keys: 'Ctrl+P', big: true, run: () => printSlides() }),
        btn({ icon: 'download', label: 'Download copy', big: true, errPrefix: 'Download failed: ', run: () => downloadCopy() }))
    ]);
    addTab('home', 'Home', [
      group('Clipboard',
        btn({ icon: 'paste', label: 'Paste', keys: 'Ctrl+V', big: true, run: () => pasteFromRibbon() }),
        h('div', { class: 'sl-stack' },
          btn({ icon: 'cut', label: 'Cut', keys: 'Ctrl+X', needsSel: true, run: () => { if (copyShapes(null)) deleteShapes(); } }),
          btn({ icon: 'copy', label: 'Copy', keys: 'Ctrl+C', needsSel: true, run: () => copyShapes(null) }),
          btn({ icon: 'duplicate', label: 'Duplicate', keys: 'Ctrl+D', needsSel: true, run: () => duplicateShapes() }))),
      group('Undo',
        h('div', { class: 'sl-stack' },
          btn({ icon: 'undo', label: 'Undo', keys: 'Ctrl+Z', run: () => undo() }),
          btn({ icon: 'redo', label: 'Redo', keys: 'Ctrl+Y', run: () => redo() }))),
      group('Slides',
        h('span', { class: 'sl-split sl-split-big' },
          btn({ icon: 'slideNew', label: 'New Slide', keys: 'Ctrl+M', big: true, run: () => newSlide() }),
          btn({ icon: 'chevronDown', label: 'New slide with layout', iconOnly: true, cls: 'sl-split-arrow', menu: true, noCaret: true, run: (b) => layoutMenu(b, (id) => newSlide(id)) })),
        h('div', { class: 'sl-stack' },
          btn({ icon: 'layout', label: 'Layout', menu: true, run: (b) => layoutMenu(b, (id) => applyLayout(id), slide().layout) }),
          btn({ icon: 'reset', label: 'Reset', title: 'Reset placeholder positions to the layout', run: () => resetSlide() }),
          btn({ icon: 'delete', label: 'Delete slide', run: () => deleteSlide() }))),
      group('Font',
        row(fontSel, sizeInput, sizeList,
          btn({ icon: 'growFont', label: 'Increase font size', iconOnly: true, keys: 'Ctrl+]', run: () => setRunProp('grow', 1) }),
          btn({ icon: 'shrinkFont', label: 'Decrease font size', iconOnly: true, keys: 'Ctrl+[', run: () => setRunProp('grow', -1) }),
          btn({ icon: 'clearFormat', label: 'Clear formatting', iconOnly: true, keys: 'Ctrl+Space', run: () => setRunProp('clear') })),
        row(
          btn({ id: 'fmt_b', icon: 'bold', label: 'Bold', iconOnly: true, keys: 'Ctrl+B', pressable: true, run: () => setRunProp('b') }),
          btn({ id: 'fmt_i', icon: 'italic', label: 'Italic', iconOnly: true, keys: 'Ctrl+I', pressable: true, run: () => setRunProp('i') }),
          btn({ id: 'fmt_u', icon: 'underline', label: 'Underline', iconOnly: true, keys: 'Ctrl+U', pressable: true, run: () => setRunProp('u') }),
          btn({ id: 'fmt_strike', icon: 'strike', label: 'Strikethrough', iconOnly: true, pressable: true, run: () => setRunProp('strike') }),
          btn({ icon: 'superscript', label: 'Superscript', iconOnly: true, run: () => setRunProp('super') }),
          btn({ icon: 'subscript', label: 'Subscript', iconOnly: true, run: () => setRunProp('sub') }),
          colorButton(fontColor), colorButton(hiColor))),
      group('Paragraph',
        row(
          btn({ id: 'bullets', icon: 'bullets', label: 'Bullets', iconOnly: true, pressable: true, run: () => toggleBullets('bullet') }),
          btn({ id: 'numbering', icon: 'numbering', label: 'Numbering', iconOnly: true, pressable: true, run: () => toggleBullets('num') }),
          btn({ icon: 'outdent', label: 'Decrease list level', iconOnly: true, keys: 'Shift+Tab', run: () => setParaProp(ps => shiftLevel(ps, -1)) }),
          btn({ icon: 'indent', label: 'Increase list level', iconOnly: true, keys: 'Tab', run: () => setParaProp(ps => shiftLevel(ps, 1)) }),
          btn({ icon: 'lineSpacing', label: 'Line spacing', iconOnly: true, menu: true, run: (b) => lineSpacingMenu(b) })),
        row(
          btn({ id: 'alignL', icon: 'alignLeft', label: 'Align left', iconOnly: true, keys: 'Ctrl+L', pressable: true, run: () => setAlign('l') }),
          btn({ id: 'alignC', icon: 'alignCenter', label: 'Center', iconOnly: true, keys: 'Ctrl+E', pressable: true, run: () => setAlign('ctr') }),
          btn({ id: 'alignR', icon: 'alignRight', label: 'Align right', iconOnly: true, keys: 'Ctrl+R', pressable: true, run: () => setAlign('r') }),
          btn({ id: 'alignJ', icon: 'alignJustify', label: 'Justify', iconOnly: true, keys: 'Ctrl+J', pressable: true, run: () => setAlign('just') }),
          btn({ icon: 'valignMiddle', label: 'Align text vertically', iconOnly: true, menu: true, run: (b) => alignTextMenu(b) }))),
      group('Drawing',
        btn({ icon: 'shapes', label: 'Shapes', big: true, menu: true, run: (b) => shapeMenu(b, (g, l) => insertGeom(g, l)) }),
        h('div', { class: 'sl-stack' },
          row(colorButton(fillColor)),
          row(colorButton(lineColor)),
          btn({ icon: 'lineSpacing', label: 'Weight', title: 'Outline weight and dashes', menu: true, needsSel: true, run: (b) => lineWeightMenu(b) }))),
      group('Arrange',
        btn({ icon: 'layers', label: 'Arrange', big: true, menu: true, needsSel: true, run: (b) => openMenu(b, arrangeItems(), { label: 'Arrange' }) }),
        h('div', { class: 'sl-stack' },
          btn({ icon: 'bringFront', label: 'Bring to front', iconOnly: true, keys: 'Ctrl+Shift+]', needsSel: true, run: () => reorder('front') }),
          btn({ icon: 'sendBack', label: 'Send to back', iconOnly: true, keys: 'Ctrl+Shift+[', needsSel: true, run: () => reorder('back') }),
          btn({ icon: 'alignObjCenter', label: 'Align objects', iconOnly: true, menu: true, needsSel: true, run: (b) => openMenu(b, arrangeItems().slice(6, 14), { label: 'Align objects' }) })))
    ]);
    addTab('insert', 'Insert', [
      group('Slides',
        btn({ icon: 'slideNew', label: 'New Slide', keys: 'Ctrl+M', big: true, run: () => newSlide() }),
        btn({ icon: 'layout', label: 'With layout', big: true, menu: true, run: (b) => layoutMenu(b, (id) => newSlide(id)) })),
      group('Text',
        btn({ icon: 'textbox', label: 'Text Box', big: true, run: () => insertTextBox() }),
        btn({ icon: 'link', label: 'Link', big: true, keys: 'Ctrl+K', run: () => insertLink() })),
      group('Illustrations',
        btn({ icon: 'image', label: 'Picture', big: true, errPrefix: 'Could not insert picture: ', run: () => insertPicture() }),
        btn({ icon: 'shapes', label: 'Shapes', big: true, menu: true, run: (b) => shapeMenu(b, (g, l) => insertGeom(g, l)) })),
      group('Tables',
        btn({ icon: 'table', label: 'Table', big: true, run: () => insertTable() }),
        btn({ icon: 'table', label: 'Rows & columns', big: true, menu: true, run: (b) => openMenu(b, tableItems(), { label: 'Table rows and columns' }) }))
    ]);
    addTab('design', 'Design', [
      group('Slide size',
        btn({ icon: 'slideSize', label: 'Widescreen (16:9)', big: true, run: () => setSlideSize(12192000, 6858000) }),
        btn({ icon: 'slideSize', label: 'Standard (4:3)', big: true, run: () => setSlideSize(9144000, 6858000) })),
      group('Background',
        row(colorButton(bgColor)),
        btn({ icon: 'copy', label: 'Apply to all', title: 'Apply this background to all slides', run: () => bgToAll() }),
        btn({ icon: 'reset', label: 'Reset background', run: () => resetBackground() }))
    ]);
    addTab('show', 'Slide Show', [
      group('Start slide show',
        btn({ icon: 'play', label: 'From Beginning', keys: 'F5', big: true, run: () => startShow(0) }),
        btn({ icon: 'playCurrent', label: 'From Current Slide', keys: 'Shift+F5', big: true, run: () => startShow(cur) })),
      group('Set up',
        btn({ id: 'hideSlide', icon: 'hide', label: 'Hide Slide', big: true, pressable: true, run: () => toggleHidden() }))
    ]);
    addTab('view', 'View', [
      group('Presentation views',
        btn({ id: 'viewNormal', icon: 'normal', label: 'Normal', big: true, pressable: true, run: () => setView('normal') }),
        btn({ id: 'viewSorter', icon: 'sorter', label: 'Slide Sorter', big: true, pressable: true, run: () => setView('sorter') })),
      group('Show',
        btn({ id: 'notesToggle', icon: 'notes', label: 'Notes', big: true, pressable: true, run: () => toggleNotes() }),
        btn({ id: 'snapToggle', icon: 'grid', label: 'Smart guides', big: true, pressable: true, title: 'Snap to smart guides while moving (hold Alt to bypass)', run: () => { snapOn = !snapOn; updateRibbonState(); } })),
      group('Zoom',
        btn({ icon: 'fit', label: 'Fit to Window', big: true, run: () => setZoom('fit') }),
        btn({ icon: 'zoomIn', label: 'Zoom in', iconOnly: true, keys: 'Ctrl+=', run: () => stepZoom(1) }),
        btn({ icon: 'zoomOut', label: 'Zoom out', iconOnly: true, keys: 'Ctrl+-', run: () => stepZoom(-1) }),
        btn({ icon: 'find', label: '100%', big: true, title: 'Zoom to 100%', run: () => setZoom(100) }))
    ]);
    selectTab('home', false);

    /* ---- status bar ---- */
    const stSlide = h('span', { class: 'sl-st' }), stSel = h('span', { class: 'sl-st sl-st-sel' }), stSave = h('span', { class: 'sl-st' });
    const zoomRange = h('input', { type: 'range', min: '10', max: '400', step: '1', value: '100', class: 'sl-zoom-range', 'aria-label': 'Zoom', title: 'Zoom (Ctrl+= / Ctrl+-, Ctrl+wheel)' });
    zoomRange.addEventListener('input', () => setZoom(+zoomRange.value));
    const zoomLbl = h('button', { type: 'button', class: 'tool-btn sl-zoom-lbl', title: 'Choose zoom level', 'aria-label': 'Zoom level', 'aria-haspopup': 'menu', 'aria-expanded': 'false' }, '100%');
    zoomLbl.addEventListener('click', () => openMenu(zoomLbl, [{ label: 'Fit slide to window', icon: 'fit', checked: zoomMode === 'fit', run: () => setZoom('fit') }, { sep: true }]
      .concat(ZOOMS.map(z => ({ label: z + '%', checked: zoomMode !== 'fit' && zoomPct === z, run: () => setZoom(z) }))), { label: 'Zoom' }));
    const stNotesBtn = btn({ icon: 'notes', label: 'Notes', pressable: true, title: 'Show or hide speaker notes', run: () => toggleNotes() });
    const statusbar = h('div', { class: 'statusbar sl-statusbar', role: 'group', 'aria-label': 'Presentation status' },
      stSlide, stSel, h('span', { class: 'sl-st-spacer' }), stSave, stNotesBtn,
      btn({ icon: 'normal', label: 'Normal view', iconOnly: true, run: () => setView('normal') }),
      btn({ icon: 'sorter', label: 'Slide sorter', iconOnly: true, run: () => setView('sorter') }),
      btn({ icon: 'play', label: 'Slide show from current slide', iconOnly: true, keys: 'Shift+F5', run: () => startShow(cur) }),
      btn({ icon: 'zoomOut', label: 'Zoom out', iconOnly: true, run: () => stepZoom(-1) }),
      zoomRange,
      btn({ icon: 'zoomIn', label: 'Zoom in', iconOnly: true, run: () => stepZoom(1) }),
      zoomLbl,
      btn({ icon: 'fit', label: 'Fit slide to current window', iconOnly: true, run: () => setZoom('fit') }));
    function updateStatus() {
      const n = model.slides.length;
      stSlide.textContent = n ? 'Slide ' + (cur + 1) + ' of ' + n + (slide() && slide().hidden ? ' (hidden)' : '') : 'No slides';
      const ss = selShapes();
      stSel.textContent = loading ? 'Opening presentation…' : editing ? 'Editing text' : ss.length === 1 ? (ss[0].name || kindLabel(ss[0])) : ss.length ? ss.length + ' objects selected' : '';
      stSave.textContent = dirty ? 'Unsaved changes' : '';
    }
    function toggleNotes() {
      notesOn = !notesOn;
      notesWrap.hidden = !notesOn;
      updateRibbonState();
      if (zoomMode === 'fit') requestAnimationFrame(applyScale);
    }
    function setView(v) {
      if (editing) commitEdit();
      view = v;
      const so = v === 'sorter';
      sorter.hidden = !so; center.hidden = so; thumbs.hidden = so;
      if (so) { renderSorter(); const it = sorter.children[cur]; if (it) { it.focus(); it.scrollIntoView({ block: 'nearest' }); } }
      else { renderThumbs(); renderCanvas(); stage.focus({ preventScroll: true }); }
      updateRibbonState();
    }

    /* ---- thumbnails and slide sorter ---- */
    const THUMB_W = 168, SORTER_W = 220;
    function makeThumb(i, width, inSorter) {
      const s = model.slides[i];
      const sc = width / W();
      const item = h('div', {
        class: 'sl-thumb' + (s.hidden ? ' is-hidden' : ''), role: 'option', id: uid + (inSorter ? '-so-' : '-th-') + i,
        'aria-selected': String(i === cur), tabindex: i === cur ? '0' : '-1', draggable: 'true', 'data-i': String(i),
        'aria-label': 'Slide ' + (i + 1) + (s.hidden ? ', hidden' : '') + ': ' + slideTitle(s)
      });
      const frame = h('div', { class: 'sl-thumb-frame' });
      frame.style.width = width + 'px';
      frame.style.height = H() * sc + 'px';
      const se = renderSlide(model, s, { index: i });
      se.style.transform = 'scale(' + sc + ')';
      frame.appendChild(se);
      item.append(h('span', { class: 'sl-thumb-num', 'aria-hidden': 'true' }, String(i + 1)), frame);
      return item;
    }
    function renderThumbs() {
      const focused = thumbs.contains(document.activeElement);
      thumbs.replaceChildren(...model.slides.map((_, i) => makeThumb(i, THUMB_W, false)));
      const it = thumbs.children[cur];
      if (it) { it.scrollIntoView({ block: 'nearest' }); if (focused) it.focus({ preventScroll: true }); }
    }
    function renderSorter() {
      const focused = sorter.contains(document.activeElement);
      sorter.replaceChildren(...model.slides.map((_, i) => makeThumb(i, SORTER_W, true)));
      const it = sorter.children[cur];
      if (it && focused) it.focus({ preventScroll: true });
    }
    function scheduleThumb(i) {
      clearTimeout(thumbTimer);
      thumbTimer = setTimeout(() => { if (!destroyed) updateThumb(i); }, 180);
    }
    function updateThumb(i) {
      const old = thumbs.children[i];
      if (!old || thumbs.children.length !== model.slides.length) { renderThumbs(); return; }
      const n = makeThumb(i, THUMB_W, false);
      const had = old.contains(document.activeElement);
      old.replaceWith(n);
      if (had) n.focus({ preventScroll: true });
    }
    function markListSel(list) {
      Array.from(list.children).forEach((t, i) => { t.setAttribute('aria-selected', String(i === cur)); t.tabIndex = i === cur ? 0 : -1; });
      const it = list.children[cur];
      if (it) it.scrollIntoView({ block: 'nearest' });
    }
    function goto(i, opts) {
      opts = opts || {};
      if (editing) commitEdit();
      if (!model.slides.length) return;
      i = clamp(i, 0, model.slides.length - 1);
      if (i !== cur) { cur = i; selected = []; }
      renderCanvas();
      markListSel(thumbs);
      if (view === 'sorter') markListSel(sorter);
      updateStatus(); updateRibbonState();
      if (opts.focusList) { const list = view === 'sorter' ? sorter : thumbs; const it = list.children[cur]; if (it) it.focus({ preventScroll: true }); }
      if (opts.speak) announce('Slide ' + (cur + 1) + ' of ' + model.slides.length + ': ' + slideTitle(slide()));
    }
    function slideMenuItems(i) {
      const s = model.slides[i];
      const items = [
        { label: 'New slide', icon: 'slideNew', keys: 'Ctrl+M', run: () => newSlide(undefined, i + 1) },
        { label: 'Duplicate slide', icon: 'duplicate', keys: 'Ctrl+D', run: () => duplicateSlide(i) },
        { label: 'Delete slide', icon: 'delete', keys: 'Delete', run: () => deleteSlide(i) },
        { sep: true },
        { heading: 'Layout' }
      ];
      for (const l of model.layouts || []) items.push({ label: l.name || l.id, checked: s.layout === l.id, run: () => { goto(i); applyLayout(l.id); } });
      items.push({ sep: true },
        { label: 'Reset slide', icon: 'reset', run: () => { goto(i); resetSlide(); } },
        { label: s.hidden ? 'Show slide' : 'Hide slide', icon: s.hidden ? 'show' : 'hide', run: () => toggleHidden(i) },
        { sep: true },
        { label: 'Move up', icon: 'chevronUp', keys: 'Ctrl+Up', disabled: i === 0, run: () => moveSlide(i, i - 1) },
        { label: 'Move down', icon: 'chevronDown', keys: 'Ctrl+Down', disabled: i >= model.slides.length - 1, run: () => moveSlide(i, i + 1) });
      return items;
    }
    function bindSlideList(list, isSorter) {
      let dragFrom = -1;
      const itemOf = (t) => t && t.closest ? t.closest('.sl-thumb') : null;
      const clearDrop = () => list.querySelectorAll('.drop-before,.drop-after,.is-dragging').forEach(x => x.classList.remove('drop-before', 'drop-after', 'is-dragging'));
      const refocus = () => { const it = list.children[cur]; if (it) it.focus({ preventScroll: true }); };
      on(list, 'click', (e) => { const it = itemOf(e.target); if (!it) return; goto(+it.dataset.i); it.focus({ preventScroll: true }); });
      on(list, 'dblclick', (e) => { const it = itemOf(e.target); if (it && isSorter) { goto(+it.dataset.i); setView('normal'); } });
      on(list, 'keydown', (e) => {
        const it = itemOf(e.target);
        if (!it) return;
        const i = +it.dataset.i, n = model.slides.length, ctrl = e.ctrlKey || e.metaKey;
        let cols = 1;
        if (isSorter) { const top = list.children[0] ? list.children[0].offsetTop : 0; cols = 0; for (const c of list.children) { if (c.offsetTop === top) cols++; else break; } cols = Math.max(cols, 1); }
        let j = null;
        switch (e.key) {
          case 'ArrowDown': j = i + cols; break;
          case 'ArrowUp': j = i - cols; break;
          case 'ArrowRight': if (isSorter) j = i + 1; break;
          case 'ArrowLeft': if (isSorter) j = i - 1; break;
          case 'Home': j = 0; break;
          case 'End': j = n - 1; break;
          case 'PageDown': j = i + 5 * cols; break;
          case 'PageUp': j = i - 5 * cols; break;
          case 'Delete': case 'Backspace': e.preventDefault(); deleteSlide(i); refocus(); return;
          case 'Enter': e.preventDefault(); if (isSorter) setView('normal'); else stage.focus(); return;
          case 'ContextMenu': e.preventDefault(); openMenu(it, slideMenuItems(i), { label: 'Slide ' + (i + 1) }); return;
          case 'F10': if (e.shiftKey) { e.preventDefault(); openMenu(it, slideMenuItems(i), { label: 'Slide ' + (i + 1) }); } return;
          case 'd': case 'D': if (ctrl) { e.preventDefault(); e.stopPropagation(); duplicateSlide(i); refocus(); } return;
          case 'h': case 'H': if (ctrl && e.shiftKey) { e.preventDefault(); toggleHidden(i); refocus(); } return;
          default: return;
        }
        if (j == null) return;
        e.preventDefault();
        j = clamp(j, 0, n - 1);
        if (ctrl && e.key.startsWith('Arrow')) { moveSlide(i, j); refocus(); }
        else goto(j, { focusList: true, speak: false });
      });
      on(list, 'contextmenu', (e) => {
        e.preventDefault();
        const it = itemOf(e.target);
        if (it) goto(+it.dataset.i);
        openMenu({ x: e.clientX, y: e.clientY }, slideMenuItems(it ? +it.dataset.i : cur), { label: 'Slide', returnFocus: it || list });
      });
      on(list, 'dragstart', (e) => {
        const it = itemOf(e.target);
        if (!it) return;
        dragFrom = +it.dataset.i;
        e.dataTransfer.effectAllowed = 'move';
        try { e.dataTransfer.setData('application/x-docstudio-slide', String(dragFrom)); } catch (_) { /* ignore */ }
        it.classList.add('is-dragging');
      });
      on(list, 'dragover', (e) => {
        if (dragFrom < 0) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        const it = itemOf(e.target);
        list.querySelectorAll('.drop-before,.drop-after').forEach(x => x.classList.remove('drop-before', 'drop-after'));
        if (!it) return;
        const r = it.getBoundingClientRect();
        const after = isSorter ? e.clientX > r.left + r.width / 2 : e.clientY > r.top + r.height / 2;
        it.classList.add(after ? 'drop-after' : 'drop-before');
      });
      on(list, 'drop', (e) => {
        if (dragFrom < 0) return;
        e.preventDefault();
        const it = list.querySelector('.drop-before,.drop-after');
        let to = model.slides.length - 1;
        if (it) { const i = +it.dataset.i; to = i + (it.classList.contains('drop-after') ? 1 : 0); if (to > dragFrom) to--; }
        const from = dragFrom;
        dragFrom = -1;
        clearDrop();
        moveSlide(from, to);
      });
      on(list, 'dragend', () => { dragFrom = -1; clearDrop(); });
    }
    bindSlideList(thumbs, false);
    bindSlideList(sorter, true);

    /* ---- notes ---- */
    let notesUndo = false;
    on(notes, 'focus', () => { notesUndo = false; });
    on(notes, 'input', () => {
      if (!notesUndo) { pushUndo(); notesUndo = true; }
      slide().notes = notes.value;
      markDirty();
    });

    /* ---- clipboard and drops on the canvas ---- */
    on(stage, 'copy', (e) => { if (editing || !selected.length) return; e.preventDefault(); copyShapes(e.clipboardData); });
    on(stage, 'cut', (e) => { if (editing || !selected.length) return; e.preventDefault(); if (copyShapes(e.clipboardData)) deleteShapes(); });
    on(stage, 'paste', async (e) => {
      const dt = e.clipboardData;
      if (editing && editing.tb.contains(e.target)) {
        e.preventDefault();
        const t = (dt && dt.getData('text/plain')) || '';
        exec('insertText', t.replace(/\r\n?/g, '\n'));
        return;
      }
      if (editing) return;
      e.preventDefault();
      const file = dt && Array.from(dt.files || []).find(f => /^image\//.test(f.type));
      if (file) { try { await insertImageBytes(new Uint8Array(await file.arrayBuffer()), file.type, file.name); } catch (err) { App.toast('Could not paste the picture: ' + err.message, 'error'); } return; }
      const text = dt ? dt.getData('text/plain') : '';
      if (clipboard && (!text || text === clipboard.text)) { pasteShapes(); return; }
      if (text) insertTextBox(text);
    });
    on(stage, 'dragover', (e) => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; } });
    on(stage, 'drop', async (e) => {
      const files = e.dataTransfer ? Array.from(e.dataTransfer.files || []).filter(f => /^image\//.test(f.type) || /\.(png|jpe?g|gif|bmp|webp|svg)$/i.test(f.name)) : [];
      if (!files.length) return;
      e.preventDefault();
      for (const f of files) {
        try { await insertImageBytes(new Uint8Array(await f.arrayBuffer()), f.type || imageMime(f.name), f.name); }
        catch (err) { App.toast('Could not insert ' + f.name + ': ' + err.message, 'error'); }
      }
    });
    on(stage, 'contextmenu', (e) => {
      if (editing && editing.tb.contains(e.target)) return;
      e.preventDefault();
      const shEl = e.target.closest && e.target.closest('.sl-shape[data-id]');
      if (editing) commitEdit();
      if (shEl && !selected.includes(shEl.dataset.id)) { selected = [shEl.dataset.id]; selectionChanged(); }
      if (!shEl && selected.length) { selected = []; selectionChanged(); }
      openCanvasMenu({ x: e.clientX, y: e.clientY });
    });
    function openCanvasMenu(pt) {
      if (!pt) {
        const ss = selShapes();
        const r = holder.getBoundingClientRect();
        if (ss.length) { const bb = bboxOf(ss); pt = { x: r.left + px(bb.x + bb.w / 2) * scale, y: r.top + px(bb.y + bb.h / 2) * scale }; }
        else pt = { x: r.left + 20, y: r.top + 20 };
      }
      const ss = selShapes();
      let items;
      if (ss.length) {
        items = [
          { label: 'Cut', icon: 'cut', keys: 'Ctrl+X', run: () => { if (copyShapes(null)) deleteShapes(); } },
          { label: 'Copy', icon: 'copy', keys: 'Ctrl+C', run: () => copyShapes(null) },
          { label: 'Paste', icon: 'paste', keys: 'Ctrl+V', disabled: !clipboard, run: () => pasteShapes() },
          { label: 'Duplicate', icon: 'duplicate', keys: 'Ctrl+D', run: () => duplicateShapes() },
          { label: 'Delete', icon: 'delete', keys: 'Delete', run: () => deleteShapes() },
          { sep: true }
        ];
        if (ss.length === 1 && canEditText(ss[0])) items.push({ label: 'Edit text', icon: 'textbox', keys: 'Enter', run: () => startEdit(ss[0].id) });
        if (ss.length === 1 && ss[0].kind === 'table') items.push(...tableItems());
        items.push({ sep: true },
          { label: 'Bring to front', icon: 'bringFront', run: () => reorder('front') },
          { label: 'Bring forward', icon: 'bringForward', run: () => reorder('forward') },
          { label: 'Send backward', icon: 'sendBackward', run: () => reorder('backward') },
          { label: 'Send to back', icon: 'sendBack', run: () => reorder('back') },
          { sep: true },
          { label: 'Size and position…', icon: 'size', run: () => sizeDialog() });
      } else {
        items = [
          { label: 'Paste', icon: 'paste', keys: 'Ctrl+V', disabled: !clipboard, run: () => pasteShapes() },
          { sep: true },
          { label: 'New slide', icon: 'slideNew', keys: 'Ctrl+M', run: () => newSlide() },
          { label: 'Text box', icon: 'textbox', run: () => insertTextBox() },
          { label: 'Picture…', icon: 'image', run: () => insertPicture() },
          { sep: true },
          { label: 'Reset slide', icon: 'reset', run: () => resetSlide() },
          { label: slide().hidden ? 'Show slide' : 'Hide slide', icon: 'hide', run: () => toggleHidden() }
        ];
      }
      openMenu(pt, items, { label: 'Context menu', returnFocus: stage });
    }

    /* ---- keyboard ---- */
    on(stage, 'keydown', (e) => {
      const ctrl = e.ctrlKey || e.metaKey;
      const k = e.key;
      const lk = k.length === 1 ? k.toLowerCase() : k;
      if (editing && editing.tb.contains(e.target)) {
        if (k === 'Escape' || k === 'F2') { e.preventDefault(); commitEdit(); stage.focus({ preventScroll: true }); return; }
        if (k === 'Tab' && !ctrl && !e.altKey) {
          e.preventDefault();
          if (editing.cell) nextCell(e.shiftKey ? -1 : 1);
          else setParaProp(ps => shiftLevel(ps, e.shiftKey ? -1 : 1));
          return;
        }
        if (ctrl && !e.altKey) {
          const al = { l: 'l', e: 'ctr', r: 'r', j: 'just' }[lk];
          if (al && !e.shiftKey) { e.preventDefault(); setAlign(al); return; }
          if (lk === ']' || lk === '[' || k === '>' || k === '<') { e.preventDefault(); setRunProp('grow', lk === ']' || k === '>' ? 1 : -1); return; }
          if (lk === 'k') { e.preventDefault(); insertLink(); return; }
          if (k === ' ') { e.preventDefault(); setRunProp('clear'); return; }
        }
        return;
      }
      if (e.target !== stage) return;
      if (loading) return;
      const ss = selShapes();
      if (k.startsWith('Arrow') && !ctrl && ss.length) {
        e.preventDefault();
        if (e.altKey && (k === 'ArrowLeft' || k === 'ArrowRight')) { rotateBy(k === 'ArrowRight' ? 15 : -15); return; }
        const step = e.shiftKey ? EMU_IN / 10 : EMU_IN / 50;
        nudge(k === 'ArrowLeft' ? -step : k === 'ArrowRight' ? step : 0, k === 'ArrowUp' ? -step : k === 'ArrowDown' ? step : 0);
        return;
      }
      if ((k === 'Delete' || k === 'Backspace') && ss.length) { e.preventDefault(); deleteShapes(); return; }
      if (k === 'Tab' && !ctrl && !e.altKey) {
        const list = slide().shapes;
        if (!list.length || (!selected.length && tabOut)) { tabOut = false; return; }
        e.preventDefault();
        let i;
        if (selected.length) { i = list.findIndex(s => String(s.id) === selected[selected.length - 1]); i = (i + (e.shiftKey ? -1 : 1) + list.length) % list.length; }
        else i = e.shiftKey ? list.length - 1 : 0;
        selected = [String(list[i].id)];
        selectionChanged(true);
        return;
      }
      tabOut = false;
      if (k === 'Escape') { if (selected.length) { e.preventDefault(); selected = []; tabOut = true; selectionChanged(true); } return; }
      if ((k === 'Enter' || k === 'F2') && ss.length === 1 && canEditText(ss[0])) { e.preventDefault(); startEdit(ss[0].id, { selectAll: k === 'Enter' && false }); return; }
      if (k === 'PageDown' || (k === 'ArrowDown' && !ss.length && !ctrl) || (k === 'ArrowRight' && !ss.length && !ctrl)) { e.preventDefault(); goto(cur + 1, { speak: true }); return; }
      if (k === 'PageUp' || (k === 'ArrowUp' && !ss.length && !ctrl) || (k === 'ArrowLeft' && !ss.length && !ctrl)) { e.preventDefault(); goto(cur - 1, { speak: true }); return; }
      if (k === 'Home' && !ss.length) { e.preventDefault(); goto(0, { speak: true }); return; }
      if (k === 'End' && !ss.length) { e.preventDefault(); goto(model.slides.length - 1, { speak: true }); return; }
      if (k === 'ContextMenu' || (e.shiftKey && k === 'F10')) { e.preventDefault(); openCanvasMenu(null); return; }
      if (ctrl && !e.altKey) {
        let handled = true;
        switch (lk) {
          case 'a': selectAll(); break;
          case 'd': duplicateShapes(); break;
          case 'b': setRunProp('b'); break;
          case 'i': setRunProp('i'); break;
          case 'u': setRunProp('u'); break;
          case 'l': setAlign('l'); break;
          case 'e': setAlign('ctr'); break;
          case 'r': setAlign('r'); break;
          case 'j': setAlign('just'); break;
          case 'k': insertLink(); break;
          case ']': case '}': if (e.shiftKey) reorder('front'); else if (ss.length) reorder('forward'); break;
          case '[': case '{': if (e.shiftKey) reorder('back'); else if (ss.length) reorder('backward'); break;
          case '>': setRunProp('grow', 1); break;
          case '<': setRunProp('grow', -1); break;
          default: handled = false;
        }
        if (handled) { e.preventDefault(); e.stopPropagation(); }
      }
    });
    function isTextTarget(t) {
      return !!t && (t.isContentEditable || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || (t.tagName === 'INPUT' && !['button', 'checkbox', 'radio', 'range', 'color'].includes(t.type)));
    }
    function focusCycle(dir) {
      const zones = [];
      zones.push(activeTab().tab);
      if (view === 'sorter') zones.push(sorter.children[cur] || sorter);
      else {
        zones.push(thumbs.children[cur] || thumbs);
        zones.push(editing ? editing.tb : stage);
        if (notesOn) zones.push(notes);
      }
      zones.push(statusbar.querySelector('button') || statusbar);
      const a = document.activeElement;
      let i = zones.findIndex(z => z === a || (z.contains && z.contains(a)));
      if (i < 0 && ribbon.contains(a)) i = 0;
      if (i < 0 && thumbs.contains(a)) i = 1;
      const next = zones[(i + dir + zones.length) % zones.length];
      if (editing && next === editing.tb) restoreEditRange(); else next.focus();
    }
    on(root, 'keydown', (e) => {
      const ctrl = e.ctrlKey || e.metaKey;
      const lk = e.key.length === 1 ? e.key.toLowerCase() : e.key;
      if (e.key === 'F6') { e.preventDefault(); focusCycle(e.shiftKey ? -1 : 1); return; }
      if (e.key === 'F5' && !ctrl) { e.preventDefault(); startShow(e.shiftKey ? cur : 0); return; }
      if (!ctrl || e.altKey) return;
      const inText = isTextTarget(e.target);
      if (lk === 'p' && !e.shiftKey) { e.preventDefault(); printSlides(); }
      else if (lk === 'm' && !e.shiftKey) { e.preventDefault(); newSlide(); }
      else if (lk === 'z' && !e.shiftKey && !inText) { e.preventDefault(); undo(); }
      else if ((lk === 'y' || (lk === 'z' && e.shiftKey)) && !inText) { e.preventDefault(); redo(); }
      else if ((e.key === '=' || e.key === '+') && !e.shiftKey) { e.preventDefault(); stepZoom(1); }
      else if (e.key === '-' || e.key === '_') { e.preventDefault(); stepZoom(-1); }
      else if (e.key === '0' && !inText) { e.preventDefault(); setZoom('fit'); }
    });

    /* ---- slide show ---- */
    function startShow(from) {
      if (loading) return;
      if (editing) commitEdit();
      closeMenu(false);
      const order = [];
      model.slides.forEach((s, i) => { if (!s.hidden || i === from) order.push(i); });
      if (!order.length) return;
      let pos = Math.max(0, order.indexOf(from));
      let black = null, done = false, fsEntered = false, wheelAt = 0;
      const prevFocus = document.activeElement;
      const frame = h('div', { class: 'sl-show-frame' });
      const counter = h('div', { class: 'sl-show-counter', 'aria-hidden': 'true' });
      const sr = h('div', { class: 'visually-hidden', 'aria-live': 'polite', 'aria-atomic': 'true' });
      const ov = h('div', { class: 'sl-show', role: 'dialog', 'aria-modal': 'true', tabindex: '-1',
        'aria-label': 'Slide show. Space, arrow keys or click to move; B for black screen; Escape to end.' }, frame, counter, sr);
      document.body.appendChild(ov);
      function draw() {
        frame.textContent = '';
        ov.classList.toggle('is-black', black === 'b');
        ov.classList.toggle('is-white', black === 'w');
        if (black) { frame.hidden = true; counter.textContent = ''; return; }
        frame.hidden = false;
        if (pos >= order.length) {
          frame.style.width = ''; frame.style.height = '';
          frame.appendChild(h('div', { class: 'sl-show-end' }, 'End of slide show. Click or press any key to exit.'));
          counter.textContent = '';
          sr.textContent = 'End of slide show';
          return;
        }
        const i = order[pos];
        const sc = Math.min(window.innerWidth / W(), window.innerHeight / H());
        frame.style.width = W() * sc + 'px';
        frame.style.height = H() * sc + 'px';
        const se = renderSlide(model, model.slides[i], { index: i });
        se.style.transform = 'scale(' + sc + ')';
        frame.appendChild(se);
        counter.textContent = (pos + 1) + ' / ' + order.length;
        sr.textContent = 'Slide ' + (i + 1) + ': ' + slideTitle(model.slides[i]);
      }
      const next = () => { if (black) { black = null; draw(); return; } if (pos >= order.length) { exit(); return; } pos++; draw(); };
      const prev = () => { if (black) { black = null; draw(); return; } if (pos > 0) { pos = Math.min(pos, order.length) - 1; draw(); } };
      function exit() {
        if (done) return;
        done = true;
        showExit = null;
        window.removeEventListener('resize', draw);
        document.removeEventListener('fullscreenchange', onFs);
        if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
        ov.remove();
        if (!destroyed) {
          const lastI = order[Math.min(pos, order.length - 1)];
          if (lastI != null) goto(lastI);
          if (prevFocus && document.contains(prevFocus)) prevFocus.focus({ preventScroll: true }); else stage.focus();
        }
      }
      function onFs() { if (!document.fullscreenElement && fsEntered) exit(); }
      ov.addEventListener('keydown', (e) => {
        const k = e.key;
        e.stopPropagation();
        if (k === 'Escape' || k === '-' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); exit(); return; }
        if (e.ctrlKey || e.metaKey || e.altKey) return;
        e.preventDefault();
        const back = ['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k);
        if (pos >= order.length && !back) { exit(); return; }
        if ([' ', 'ArrowRight', 'ArrowDown', 'PageDown', 'Enter', 'n', 'N'].includes(k)) next();
        else if (back) prev();
        else if (k === 'Home') { black = null; pos = 0; draw(); }
        else if (k === 'End') { black = null; pos = order.length - 1; draw(); }
        else if (k === 'b' || k === 'B' || k === '.') { black = black === 'b' ? null : 'b'; draw(); }
        else if (k === 'w' || k === 'W' || k === ',') { black = black === 'w' ? null : 'w'; draw(); }
      });
      ov.addEventListener('click', () => next());
      ov.addEventListener('contextmenu', (e) => { e.preventDefault(); prev(); });
      ov.addEventListener('wheel', (e) => { const t = Date.now(); if (t - wheelAt < 250) return; wheelAt = t; if (e.deltaY > 0) next(); else if (e.deltaY < 0) prev(); }, { passive: true });
      window.addEventListener('resize', draw);
      document.addEventListener('fullscreenchange', onFs);
      try {
        const p = ov.requestFullscreen ? ov.requestFullscreen() : null;
        if (p && p.then) p.then(() => { fsEntered = true; }).catch(() => {});
      } catch (_) { /* windowed fallback */ }
      showExit = exit;
      draw();
      ov.focus();
    }

    /* ---- print / export PDF ---- */
    async function printSlides() {
      if (editing) commitEdit();
      const oldRoot = document.getElementById('slides-print-root'); if (oldRoot) oldRoot.remove();
      const oldStyle = document.getElementById('slides-print-style'); if (oldStyle) oldStyle.remove();
      const pr = h('div', { id: 'slides-print-root', 'aria-hidden': 'true' });
      model.slides.forEach((s, i) => {
        if (s.hidden) return;
        const page = h('div', { class: 'sl-print-page' });
        page.style.width = W() + 'px';
        page.style.height = H() + 'px';
        page.appendChild(renderSlide(model, s, { index: i }));
        pr.appendChild(page);
      });
      const wIn = (model.size.cx / EMU_IN).toFixed(3), hIn = (model.size.cy / EMU_IN).toFixed(3);
      const style = h('style', { id: 'slides-print-style' }, '@page { size: ' + wIn + 'in ' + hIn + 'in; margin: 0; }');
      document.head.appendChild(style);
      document.body.appendChild(pr);
      document.body.classList.add('printing-slides');
      const cleanup = () => {
        document.body.classList.remove('printing-slides');
        pr.remove(); style.remove();
        window.removeEventListener('afterprint', cleanup);
      };
      window.addEventListener('afterprint', cleanup);
      await Promise.all(Array.from(pr.querySelectorAll('img')).map(im => (im.decode ? im.decode() : Promise.resolve()).catch(() => {})));
      setTimeout(() => { window.print(); setTimeout(cleanup, 1000); }, 60);
    }

    /* ---- open / save ---- */
    async function serialize() {
      if (editing) commitEdit();
      const form = new FormData();
      form.append('model', JSON.stringify(model));
      if (original) form.append('original', new Blob([original], { type: PPTX_MIME }), 'original.pptx');
      const res = await App.api.postForm('/api/pptx/from-json', form);
      return new Uint8Array(await res.arrayBuffer());
    }
    async function save() {
      if (loading) return false;
      if (!docPath || extOf(docPath) !== '.pptx') return saveAs();
      stSave.textContent = 'Saving…';
      try {
        const bytes = await serialize();
        docPath = await App.writeFile(docPath, bytes);
        doc.path = docPath;
        setClean();
        App.toast('Saved ' + docName, 'success');
        return true;
      } catch (err) { updateStatus(); throw err; }
    }
    async function saveAs() {
      if (loading) return false;
      if (editing) commitEdit();
      let p = await App.pickFile({ mode: 'save', extensions: ['.pptx'], suggestedName: stripExt(docName || 'Presentation1') + '.pptx' });
      if (!p) return false;
      if (Array.isArray(p)) p = p[0];
      if (extOf(p) !== '.pptx') p += '.pptx';
      stSave.textContent = 'Saving…';
      try {
        const bytes = await serialize();
        docPath = await App.writeFile(p, bytes);
        doc.path = docPath;
        docName = baseName(docPath);
        doc.name = docName;
        App.setTitle(editor, docName);
        setClean();
        App.toast('Saved as ' + docName, 'success');
        return true;
      } catch (err) { updateStatus(); throw err; }
    }
    async function downloadCopy() {
      const bytes = await serialize();
      App.download(new Blob([bytes], { type: PPTX_MIME }), stripExt(docName || 'Presentation1') + '.pptx');
    }
    async function load() {
      let j = null;
      if (doc.model) j = { model: clone(doc.model) };
      else if (doc.bytes && doc.bytes.length) {
        const res = await App.api.postBytes('/api/pptx/to-json', doc.bytes, PPTX_MIME);
        j = await res.json();
        original = doc.bytes;
      } else {
        try { const res = await App.api.postBytes('/api/pptx/to-json', new Uint8Array(0), PPTX_MIME); j = await res.json(); }
        catch (err) { console.warn('slides: blank template unavailable, using built-in model', err); j = { model: blankModel() }; }
      }
      if (destroyed) return;
      model = normalizeModel(j && (j.model || j));
      if (!model.slides.length) model.slides.push(makeSlide((model.layouts.find(l => /title slide/i.test(l.name || '')) || model.layouts[0] || {}).id || ''));
      for (const s of model.slides) {
        origBg.set(s.id, s.bg ? clone(s.bg) : null);
        for (const sh of s.shapes) if (sh.ph) origGeom.set(s.id + '/' + sh.id, { x: sh.x, y: sh.y, w: sh.w, h: sh.h });
      }
      fillFonts();
      loading = false;
      cur = 0;
      renderThumbs();
      renderCanvas();
      updateStatus(); updateRibbonState();
    }

    /* ---- assemble ---- */
    root.append(ribbon, main, statusbar, live);
    container.appendChild(root);
    stage.querySelector('.sl-stage-inner').appendChild(h('div', { class: 'sl-loading' }, 'Opening presentation…'));
    let ro = null;
    if (window.ResizeObserver) { ro = new ResizeObserver(() => { if (zoomMode === 'fit' && !destroyed) applyScale(); }); ro.observe(stage); }
    updateStatus();

    editor = {
      save, saveAs,
      isDirty: () => dirty,
      focus() {
        if (view === 'sorter') { const it = sorter.children[cur]; if (it) { it.focus(); return; } }
        if (editing) restoreEditRange(); else stage.focus({ preventScroll: true });
      },
      destroy() {
        destroyed = true;
        if (showExit) showExit();
        closeMenu(false);
        listeners.forEach(([t, ty, fn, o]) => t.removeEventListener(ty, fn, o));
        listeners.length = 0;
        clearTimeout(thumbTimer);
        if (stateTimer) cancelAnimationFrame(stateTimer);
        if (ro) ro.disconnect();
        root.remove();
      },
      // for tests / debugging
      _model: () => model
    };
    editor._ready = load().then(() => {
      const ld = stage.querySelector('.sl-loading'); if (ld) ld.remove();
      return editor;
    }, (err) => { editor.destroy(); throw err; });
    if (window.DocStudioSlides) window.DocStudioSlides.last = editor;
    return editor;
  }

  /* ---------- registration ---------- */
  async function createAsync(container, doc) {
    const ed = createSlidesEditor(container, doc);
    await ed._ready;
    return ed;
  }
  function register() {
    window.App.registerEditor('slides', { label: 'Presentation', icon: 'pptx', extensions: ['.pptx'], create: createAsync });
  }
  if (window.App && typeof window.App.registerEditor === 'function') register();
  else window.addEventListener('DOMContentLoaded', () => { if (window.App) register(); else console.error('slides.js: window.App is missing'); });

  // exposed for testing
  window.DocStudioSlides = { renderSlide, parseEditable, geomPath, autoNumber, create: createAsync };
})();
