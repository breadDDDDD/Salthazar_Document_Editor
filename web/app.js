/* DocStudio application shell.
 * Defines window.App (see the FRONTEND CONTRACT): editor registry, document tabs,
 * Home screen, file picker, dialogs, toasts, settings and global shortcuts.
 * Plain ES2020, no dependencies.
 */
(function () {
  "use strict";

  const RECENT_KEY = "docstudio.recent";
  const SETTINGS_KEY = "docstudio.settings";
  const LASTDIR_KEY = "docstudio.lastDir";
  const IMAGE_EXTS = [".png", ".jpg", ".jpeg"];
  const isMac = /Mac/.test(navigator.platform);

  const editors = {};       // kind -> definition
  const docs = [];          // open documents: {id, kind, def, editor, name, path, dirty, tab, wrap, panel}
  let activeId = null;      // null = Home
  let nextId = 1;
  let untitled = 1;

  // ---------- small helpers ----------
  function el(tag, attrs, ...children) {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === "class") e.className = v;
      else if (k === "text") e.textContent = v;
      else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? "" : v);
    }
    for (const c of children.flat()) if (c != null) e.append(c);
    return e;
  }
  const $ = (id) => document.getElementById(id);
  const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (_) { return d; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (_) { /* ignore */ } };
  const extOf = (name) => { const m = /\.[^./\\]+$/.exec(name || ""); return m ? m[0].toLowerCase() : ""; };
  const baseName = (p) => (p || "").split(/[\\/]/).filter(Boolean).pop() || p || "";
  const sepOf = (p) => (p.includes("\\") ? "\\" : "/");
  const joinPath = (dir, name) => (!dir ? name : /[\\/]$/.test(dir) ? dir + name : dir + sepOf(dir) + name);
  function fmtSize(n) {
    if (n < 1024) return n + " B";
    if (n < 1048576) return (n / 1024).toFixed(1) + " KB";
    return (n / 1048576).toFixed(1) + " MB";
  }
  function fmtDate(d) {
    const t = new Date(d);
    return isNaN(t) ? "" : t.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" });
  }
  // ---------- icons ----------
  // One inline SVG icon set (24x24 viewBox, 1.5px stroke, round caps, currentColor).
  // Use App.icon(name, {size, title, className}) -> SVGElement, or App.iconHTML(name, size) -> string.
  // Names: home new open save saveAs print settings help close undo redo cut copy paste
  //   bold italic underline strike superscript subscript fontColor highlight clearFormat fontGrow fontShrink
  //   bulletList numberList indent outdent alignLeft alignCenter alignRight alignJustify
  //   alignTop alignMiddle alignBottom lineSpacing find replace selectAll image table link unlink
  //   pageBreak hr symbol date shape textbox slide duplicate trash play zoomIn zoomOut fit
  //   pen highlighter eraser comment stamp signature rotateLeft rotateRight merge split lock unlock
  //   chevronDown chevronLeft chevronRight chevronUp plus minus grid notes doc pdf pptx folder
  //   upload download arrowUp check warning info bringForward sendBackward hand cursor watermark
  //   compress form layers eye eyeOff desktop drive computer more wordCount
  const P = (d) => '<path d="' + d + '"/>';
  const C = (cx, cy, r) => '<circle cx="' + cx + '" cy="' + cy + '" r="' + r + '"/>';
  const FILE = "M14 3H6.5A1.5 1.5 0 0 0 5 4.5v15A1.5 1.5 0 0 0 6.5 21h11a1.5 1.5 0 0 0 1.5-1.5V8zM14 3v5h5";
  const ICONS = {
    home: P("M3.5 10.5 12 3.5l8.5 7M5.5 9v11h4.5v-6h4v6h4.5V9"),
    new: P(FILE) + P("M12 11.5v6M9 14.5h6"),
    open: P("M3.5 19V5.5h5.5l2 2h8v3.5") + P("M3.5 19l2.8-8h15.2l-2.8 8z"),
    save: P("M5 3.5h11l3 3V20.5H5zM8 3.5v5h7v-5M8 20.5v-6.5h8v6.5"),
    saveAs: P("M12 20.5H5v-17h11l3 3V11M8 3.5v5h7v-5M8 20.5v-6h4") + P("M14 21l.6-2.6 4.9-4.9 2 2-4.9 4.9z"),
    print: P("M7 8.5v-5h10v5M7 17H4.5V9.5a1 1 0 0 1 1-1h13a1 1 0 0 1 1 1V17H17M7 13.5h10v7H7z"),
    settings: P("M4 6.5h8.5M16.5 6.5H20M4 12h3M11 12h9M4 17.5h10.5M18.5 17.5H20") + C(14.5, 6.5, 2) + C(9, 12, 2) + C(16.5, 17.5, 2),
    help: C(12, 12, 9) + P("M9.6 9.4a2.5 2.5 0 1 1 3.4 2.3c-.6.3-1 .8-1 1.5v.6M12 16.8v.2"),
    close: P("M6.5 6.5l11 11M17.5 6.5l-11 11"),
    undo: P("M9 14.5 4 9.5l5-5M4 9.5h10.5a5.5 5.5 0 0 1 0 11H11"),
    redo: P("M15 14.5l5-5-5-5M20 9.5H9.5a5.5 5.5 0 0 0 0 11H13"),
    cut: C(6.5, 17.5, 2.5) + C(17.5, 17.5, 2.5) + P("M8.3 15.7 18 4M15.7 15.7 6 4"),
    copy: P("M9 8.5h10.5v12H9z") + P("M15 8.5V4H4.5v12H9"),
    paste: P("M9 4.5H5.5v16h13v-16H15M9 3h6v3H9z"),
    bold: P("M7.5 4.5h5.5a3.75 3.75 0 0 1 0 7.5H7.5zM7.5 12h6.5a3.75 3.75 0 0 1 0 7.5H7.5z"),
    italic: P("M10.5 4.5h7.5M6 19.5h7.5M14.5 4.5l-5 15"),
    underline: P("M7 4v7a5 5 0 0 0 10 0V4M5.5 20.5h13"),
    strike: P("M4.5 12h15M16 7.2C15.4 5.6 13.9 4.5 12 4.5c-2.4 0-4 1.3-4 3.1 0 1.2.7 2 1.9 2.6M8 16.8c.6 1.6 2.1 2.7 4 2.7 2.4 0 4-1.3 4-3.1 0-.5-.1-1-.4-1.4"),
    superscript: P("M4 8l7.5 11.5M11.5 8 4 19.5M15.5 5c.3-.9 1.1-1.5 2.1-1.5 1.1 0 2 .8 2 1.8 0 1.7-4.1 2.3-4.1 4.2h4.1"),
    subscript: P("M4 4.5 11.5 16M11.5 4.5 4 16M15.5 15.5c.3-.9 1.1-1.5 2.1-1.5 1.1 0 2 .8 2 1.8 0 1.7-4.1 2.3-4.1 4.2h4.1"),
    fontColor: P("M6.5 15.5 12 3.5l5.5 12M8.6 11h6.8") + '<path d="M4 20.5h16" stroke-width="3"/>',
    highlight: P("M8.5 14.5 6 17v2.5h4l2.5-2.5M8.5 14.5l7-10 4.5 4.5-10 7zM14 20.5h6.5"),
    clearFormat: P("M5 5h11M10.5 5 8 19M14 14l6 6M20 14l-6 6"),
    fontGrow: P("M3 19.5 8.5 5.5l5.5 14M5.2 14.5h6.6M18 4v6M15 7h6"),
    fontShrink: P("M3 19.5 8.5 5.5l5.5 14M5.2 14.5h6.6M15 7h6"),
    bulletList: P("M9.5 6h11M9.5 12h11M9.5 18h11") + C(4.5, 6, 1) + C(4.5, 12, 1) + C(4.5, 18, 1),
    numberList: P("M10 6h10.5M10 12h10.5M10 18h10.5M4 4h1.5v4.5M4 8.5h3M4 14.6c.2-.6.8-1.1 1.5-1.1.8 0 1.5.6 1.5 1.3 0 1.3-3 1.8-3 3.7h3"),
    indent: P("M3.5 4.5h17M11 9.5h9.5M11 14.5h9.5M3.5 19.5h17M3.5 9l3.5 3-3.5 3"),
    outdent: P("M3.5 4.5h17M11 9.5h9.5M11 14.5h9.5M3.5 19.5h17M7 9l-3.5 3L7 15"),
    alignLeft: P("M3.5 5h17M3.5 10h11M3.5 15h17M3.5 20h11"),
    alignCenter: P("M3.5 5h17M6.5 10h11M3.5 15h17M6.5 20h11"),
    alignRight: P("M3.5 5h17M9.5 10h11M3.5 15h17M9.5 20h11"),
    alignJustify: P("M3.5 5h17M3.5 10h17M3.5 15h17M3.5 20h17"),
    alignTop: P("M4 3.5h16M8 7.5h3v13H8zM13 7.5h3v8h-3z"),
    alignMiddle: P("M3 12h2M19 12h2M8 5.5h3v13H8zM13 8h3v8h-3z"),
    alignBottom: P("M4 20.5h16M8 3.5h3v13H8zM13 8.5h3v8h-3z"),
    lineSpacing: P("M11 6h9.5M11 12h9.5M11 18h9.5M6 4v16M3.5 6.5 6 4l2.5 2.5M3.5 17.5 6 20l2.5-2.5"),
    find: C(10.5, 10.5, 6.5) + P("M15.5 15.5l5 5"),
    replace: P("M4.5 9a7.5 7.5 0 0 1 13.4-3.2l1.6 2M19.5 3.5v4.3h-4.3M19.5 15a7.5 7.5 0 0 1-13.4 3.2l-1.6-2M4.5 20.5v-4.3h4.3"),
    selectAll: P("M4 7V4h3M10 4h4M17 4h3v3M20 10v4M20 17v3h-3M14 20h-4M7 20H4v-3M4 14v-4M9 9h6v6H9z"),
    image: '<rect x="3.5" y="4.5" width="17" height="15" rx="1.5"/>' + C(9, 10, 1.75) + P("M20.5 15.5 15.5 11 7 19.5"),
    table: '<rect x="3.5" y="4.5" width="17" height="15" rx="1.5"/>' + P("M3.5 9.5h17M3.5 14.5h17M9.5 4.5v15M14.5 4.5v15"),
    link: P("M10 14a4 4 0 0 0 5.7 0l3-3A4 4 0 0 0 13 5.3l-1 1M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"),
    unlink: P("M14.5 15.5l-1.5 1.5a4 4 0 0 1-5.7-5.7l1.5-1.5M9.5 8.5 11 7a4 4 0 0 1 5.7 5.7l-1.5 1.5M4 4l16 16"),
    pageBreak: P("M6 3v5h12V3M6 21v-5h12v5M3 12h2.5M8 12h3M13 12h3M18.5 12H21"),
    hr: P("M3 12h18M6 7.5h12M6 16.5h12"),
    symbol: P("M5 20h4.5v-2.6A7 7 0 1 1 14.5 17.4V20H19"),
    date: '<rect x="3.5" y="5" width="17" height="15.5" rx="1.5"/>' + P("M3.5 10h17M8 3v4M16 3v4M7.5 14h2M11 14h2M14.5 14h2M7.5 17h2M11 17h2"),
    shape: P("M3.5 11h9v9.5h-9z") + C(15.5, 8.5, 5),
    textbox: '<rect x="3.5" y="4.5" width="17" height="15" rx="1"/>' + P("M8 9h8M12 9v7"),
    slide: '<rect x="3" y="4" width="18" height="12.5" rx="1.5"/>' + P("M12 16.5V21M8.5 21h7"),
    duplicate: P("M9 8.5h11.5v12H9zM15 8.5V4H3.5v12H9M14.75 11.5v6M11.75 14.5h6"),
    trash: P("M4 6.5h16M10 11v6M14 11v6M6 6.5l1 14h10l1-14M9 6.5v-3h6v3"),
    play: P("M7.5 4.5v15L19.5 12z"),
    zoomIn: C(10.5, 10.5, 6.5) + P("M15.5 15.5l5 5M8 10.5h5M10.5 8v5"),
    zoomOut: C(10.5, 10.5, 6.5) + P("M15.5 15.5l5 5M8 10.5h5"),
    fit: P("M4 9V4h5M15 4h5v5M20 15v5h-5M9 20H4v-5"),
    pen: P("M4 20l1-4.5L15.5 5l3.5 3.5L8.5 19zM13.5 7l3.5 3.5"),
    highlighter: P("M9 15.5 6.5 18v2.5h4L13 18M9 15.5l6.5-11 4.5 4.5-11 6.5zM15 20.5h5.5"),
    eraser: P("M8 20.5h12.5M4.5 15.5 14 6l5 5-8.5 9.5H8.5zM9.5 11l5 5"),
    comment: P("M4 5h16v11H9.5L4 20z"),
    stamp: P("M10 11V8.5a2 2 0 1 1 4 0V11M7 11h10v3H7zM4 14h16v3.5H4zM6 21h12"),
    signature: P("M3 16.5c3-1 5-9.5 8-9.5s-1 9.5 1 9.5c1.5 0 2-3 3-3s1 2 2 2 2-1 4-1M3 20.5h18"),
    rotateLeft: P("M4 4v5h5M4.6 9A8 8 0 1 1 6 17.6"),
    rotateRight: P("M20 4v5h-5M19.4 9A8 8 0 1 0 18 17.6"),
    merge: P("M6 3.5V8a6 6 0 0 0 6 6M18 3.5V8a6 6 0 0 1-6 6v6.5M9 17.5l3 3 3-3"),
    split: P("M12 3.5V10a6 6 0 0 1-6 6v4.5M12 10a6 6 0 0 0 6 6v4.5M9 6.5l3-3 3 3"),
    lock: '<rect x="5" y="10.5" width="14" height="10" rx="1.5"/>' + P("M8 10.5V7a4 4 0 0 1 8 0v3.5"),
    unlock: '<rect x="5" y="10.5" width="14" height="10" rx="1.5"/>' + P("M8 10.5V7a4 4 0 0 1 7.8-1.2"),
    chevronDown: P("M6 9l6 6 6-6"),
    chevronUp: P("M6 15l6-6 6 6"),
    chevronLeft: P("M15 6l-6 6 6 6"),
    chevronRight: P("M9 6l6 6-6 6"),
    plus: P("M12 5v14M5 12h14"),
    minus: P("M5 12h14"),
    grid: P("M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z"),
    notes: P("M5.5 3.5h13v17h-13zM9 8h6M9 12h6M9 16h3.5"),
    doc: P(FILE) + P("M8.5 12.5h7M8.5 16h7"),
    pdf: P(FILE) + P("M8.5 17.5c2.2-1.5 4-4.6 3.6-6.3-.3-1.2-1.6-.6-1.3.9.4 2.1 2.6 4.3 4.7 4.8"),
    pptx: P(FILE) + P("M12 11v3.5h3.5A3.5 3.5 0 1 1 12 11z"),
    folder: P("M3.5 6V5h5.5l2 2h9.5v12.5h-17z"),
    upload: P("M12 15.5V4M7.5 8.5 12 4l4.5 4.5M4 15v5h16v-5"),
    download: P("M12 4v11.5M7.5 11 12 15.5l4.5-4.5M4 15v5h16v-5"),
    arrowUp: P("M12 19.5v-15M6 10.5l6-6 6 6"),
    check: P("M5 12.5l4.5 4.5L19 7.5"),
    warning: P("M12 3.5 2.5 20h19zM12 10v4.5M12 17.2v.3"),
    info: C(12, 12, 9) + P("M12 11v5.5M12 7.8v.2"),
    bringForward: P("M4 15V4h11") + '<rect x="8" y="8" width="12" height="12" rx="1" fill="currentColor" fill-opacity=".18"/>',
    sendBackward: '<rect x="4" y="4" width="12" height="12" rx="1" fill="currentColor" fill-opacity=".18"/>' + P("M20 9v11H9"),
    hand: P("M8 13V6a1.5 1.5 0 0 1 3 0v5.5M11 11V4.5a1.5 1.5 0 0 1 3 0V11M14 11V6.5a1.5 1.5 0 0 1 3 0V14c0 4-2.5 6.5-6 6.5-2.4 0-3.9-1.1-5.3-3.2l-2-3a1.5 1.5 0 0 1 2.4-1.8L8 15"),
    cursor: P("M5.5 3.5 10.5 20l2.5-7.5L20.5 10z"),
    watermark: P("M12 3.5s6 6.5 6 11a6 6 0 0 1-12 0c0-4.5 6-11 6-11z"),
    compress: P("M4 9h5V4M20 9h-5V4M4 15h5v5M20 15h-5v5"),
    form: '<rect x="3.5" y="4.5" width="17" height="15" rx="1.5"/>' + P("M7 9.5h3.5M13 9.5h4M7 14.5h3.5M13 14.5h4"),
    layers: P("M12 3.5 3 8l9 4.5L21 8zM3 12.5 12 17l9-4.5M3 16.5 12 21l9-4.5"),
    eye: P("M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z") + C(12, 12, 3),
    eyeOff: P("M3.5 3.5l17 17M10.6 5.6c.5-.1.9-.1 1.4-.1 6 0 9.5 6.5 9.5 6.5a16 16 0 0 1-2.7 3.5M6.6 6.8C4 8.5 2.5 12 2.5 12S6 18.5 12 18.5c1.7 0 3.2-.5 4.5-1.2M9.9 9.9a3 3 0 0 0 4.2 4.2"),
    desktop: '<rect x="3" y="4" width="18" height="12.5" rx="1.5"/>' + P("M8.5 20.5h7M12 16.5v4"),
    drive: P("M3.5 13.5h17v6h-17zM3.5 13.5 6.5 5h11l3 8.5M16.5 16.5h.5"),
    computer: P("M4 5h16v10.5H4zM2.5 19.5h19"),
    more: C(6, 12, 1) + C(12, 12, 1) + C(18, 12, 1),
    wordCount: P("M4 7.5h8M4 12h8M4 16.5h5M15 9.5l1.5-1.5v8.5M15 16.5h3"),
    slideNew: '<rect x="3" y="4" width="14" height="10" rx="1.5"/>' + P("M19.5 14v6M16.5 17h6M7 18h5"),
    layout: '<rect x="3.5" y="4.5" width="17" height="15" rx="1.5"/>' + P("M3.5 9h17M10 9v10.5"),
    reset: P("M4 4v5h5M4.6 9A8 8 0 1 1 4 12"),
    fill: P("M5 12.5 11.5 6l6 6-6.5 6.5a1.5 1.5 0 0 1-2 0L5 14.5a1.5 1.5 0 0 1 0-2zM9 4l2.5 2M5 12.5h12.5M19.5 15.5s1.5 2 1.5 3a1.5 1.5 0 0 1-3 0c0-1 1.5-3 1.5-3z"),
    outline: '<rect x="4" y="4" width="16" height="16" rx="1.5" stroke-dasharray="3 2.5"/>',
    bringFront: '<rect x="3.5" y="3.5" width="10" height="10" rx="1"/><rect x="10.5" y="10.5" width="10" height="10" rx="1"/>' + '<rect x="7" y="7" width="10" height="10" rx="1" fill="currentColor" fill-opacity=".25"/>',
    sendBack: '<rect x="7" y="7" width="10" height="10" rx="1" fill="currentColor" fill-opacity=".25"/>' + P("M7 13.5H3.5v-10h10V7M17 10.5h3.5v10h-10V17"),
    sorter: P("M3.5 5h7v5h-7zM13.5 5h7v5h-7zM3.5 14h7v5h-7zM13.5 14h7v5h-7z"),
    normal: '<rect x="3" y="4" width="18" height="16" rx="1.5"/>' + P("M8 4v16M4.5 7.5h2M4.5 10.5h2M4.5 13.5h2"),
  };
  // Aliases so editors can use their own vocabulary.
  Object.assign(ICONS, {
    growFont: ICONS.fontGrow, shrinkFont: ICONS.fontShrink, bullets: ICONS.bulletList, numbering: ICONS.numberList,
    valignTop: ICONS.alignTop, valignMiddle: ICONS.alignMiddle, valignBottom: ICONS.alignBottom,
    shapes: ICONS.shape, delete: ICONS.trash, hide: ICONS.eyeOff, show: ICONS.eye, word: ICONS.doc, presentation: ICONS.pptx,
  });
  const SVGNS = "http://www.w3.org/2000/svg";
  // Returns null for an unknown name so callers can fall back to their own glyph.
  function icon(name, opts) {
    if (!hasIcon(name)) return null;
    opts = typeof opts === "number" ? { size: opts } : opts || {};
    const size = opts.size || 20;
    const svg = document.createElementNS(SVGNS, "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("width", size);
    svg.setAttribute("height", size);
    svg.setAttribute("fill", "none");
    svg.setAttribute("stroke", "currentColor");
    svg.setAttribute("stroke-width", opts.strokeWidth || "1.5");
    svg.setAttribute("stroke-linecap", "round");
    svg.setAttribute("stroke-linejoin", "round");
    svg.setAttribute("class", "icon icon-" + name + (opts.className ? " " + opts.className : ""));
    svg.setAttribute("focusable", "false");
    if (opts.title) { svg.setAttribute("role", "img"); svg.setAttribute("aria-label", opts.title); }
    else svg.setAttribute("aria-hidden", "true");
    svg.innerHTML = ICONS[name];
    return svg;
  }
  function iconHTML(name, size) { const s = icon(name, { size }); return s ? s.outerHTML : ""; }
  const hasIcon = (name) => typeof name === "string" && Object.prototype.hasOwnProperty.call(ICONS, name);
  // File type -> {icon, cls}. cls drives the small per-type tint (ft-word/ft-pdf/ft-slides/ft-image/ft-text).
  function fileType(name) {
    const e = extOf(name);
    if (e === ".pdf") return { icon: "pdf", cls: "ft-pdf" };
    if (IMAGE_EXTS.includes(e)) return { icon: "image", cls: "ft-image" };
    if (e === ".docx" || e === ".doc") return { icon: "doc", cls: "ft-word" };
    if (e === ".pptx" || e === ".ppt") return { icon: "pptx", cls: "ft-slides" };
    return { icon: "doc", cls: "ft-text" };
  }
  function iconFor(name, size) {
    const t = fileType(name);
    return icon(t.icon, { size: size || 20, className: t.cls });
  }
  const KIND_ICONS = { word: ["doc", "ft-word"], pdf: ["pdf", "ft-pdf"], slides: ["pptx", "ft-slides"] };
  function kindIcon(kind, def) {
    const k = KIND_ICONS[kind];
    if (k) return icon(k[0], { size: 16, className: k[1] });
    return icon(def && hasIcon(def.icon) ? def.icon : "doc", { size: 16 });
  }
  // Replace the contents of every [data-icon] element in the static HTML with its SVG.
  function hydrateIcons(root) {
    (root || document).querySelectorAll("[data-icon]").forEach((n) => {
      n.textContent = "";
      n.append(icon(n.dataset.icon, { size: Number(n.dataset.iconSize) || 20 }));
    });
  }

  // ---------- API ----------
  async function check(res) {
    if (!res.ok) {
      let msg = "";
      try { msg = (await res.text()).trim(); } catch (_) { /* ignore */ }
      throw new Error(msg || res.status + " " + res.statusText);
    }
    return res;
  }
  const api = {
    get: (url) => fetch(url).then(check),
    postBytes: (url, bytes, contentType) => fetch(url, {
      method: "POST",
      headers: { "Content-Type": contentType || "application/octet-stream" },
      body: bytes,
    }).then(check),
    postForm: (url, form) => fetch(url, { method: "POST", body: form }).then(check),
  };

  async function readFile(path) {
    const res = await api.get("/api/fs/read?path=" + encodeURIComponent(path));
    return new Uint8Array(await res.arrayBuffer());
  }
  async function writeFile(path, data) {
    const res = await api.postBytes("/api/fs/write?path=" + encodeURIComponent(path), data, "application/octet-stream");
    const j = await res.json();
    const saved = j.path || path;
    if (kindForName(saved)) addRecent(saved);
    return saved;
  }
  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = el("a", { href: url, download: filename || "download" });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  // ---------- toasts ----------
  function toast(message, kind) {
    kind = kind || "info";
    const region = kind === "error" ? $("toasts-assertive") : $("toasts-polite");
    if (!region) { console.log("[" + kind + "]", message); return; }
    const glyph = kind === "error" ? "warning" : kind === "success" ? "check" : "info";
    const t = el("div", { class: "toast " + kind },
      el("span", { class: "toast-icon" }, icon(glyph, 18)),
      el("span", { class: "toast-msg", text: String(message) }),
      el("button", { type: "button", "aria-label": "Dismiss notification", title: "Dismiss", onclick: () => t.remove() }, icon("close", 16)));
    region.append(t);
    setTimeout(() => t.remove(), kind === "error" ? 12000 : 5000);
  }

  // ---------- modal dialogs ----------
  // modal({title, body: Node, buttons:[{label, value, primary, danger}], wide, init(dlg, close)}) -> Promise<value>
  // Escape / close button resolves with `cancelValue` (default null).
  function modal(opts) {
    return new Promise((resolve) => {
      const prevFocus = document.activeElement;
      const titleId = "dlg-title-" + nextId++;
      let result = opts.cancelValue === undefined ? null : opts.cancelValue;
      const dlg = el("dialog", { class: "modal" + (opts.wide ? " wide" : ""), "aria-labelledby": titleId });
      const close = (v) => { result = v; dlg.close(); };
      const foot = el("div", { class: "modal-foot" });
      for (const b of opts.buttons || []) {
        const btn = el("button", {
          type: b.submit ? "submit" : "button",
          class: "btn" + (b.primary ? " primary" : "") + (b.danger ? " danger" : ""),
          text: b.label,
        });
        if (!b.submit) btn.addEventListener("click", () => (b.onclick ? b.onclick(close) : close(b.value)));
        if (b.primary) btn.dataset.primary = "1";
        foot.append(btn);
      }
      const form = el("form", { class: "modal-form", method: "dialog" },
        el("div", { class: "modal-head" },
          el("h2", { id: titleId, text: opts.title }),
          el("button", { type: "button", class: "icon-btn", "aria-label": "Close dialog", title: "Close (Esc)", onclick: () => close(opts.cancelValue === undefined ? null : opts.cancelValue) }, icon("close", 18))),
        el("div", { class: "modal-body" }, opts.body),
        foot);
      form.addEventListener("submit", (e) => { e.preventDefault(); if (opts.onSubmit) opts.onSubmit(close); });
      dlg.append(form);
      dlg.addEventListener("close", () => {
        dlg.remove();
        if (prevFocus && prevFocus.focus && document.contains(prevFocus)) prevFocus.focus();
        resolve(result);
      });
      document.body.append(dlg);
      dlg.showModal();
      if (opts.init) opts.init(dlg, close);
      const f = opts.focus ? dlg.querySelector(opts.focus) : foot.querySelector("[data-primary]") || foot.querySelector("button");
      if (f) f.focus();
    });
  }

  function confirmDlg(message, opts) {
    opts = opts || {};
    return modal({
      title: opts.title || "Please confirm",
      body: el("p", { text: message }),
      cancelValue: false,
      buttons: [
        { label: opts.cancelLabel || "Cancel", value: false },
        { label: opts.okLabel || "OK", value: true, primary: !opts.danger, danger: !!opts.danger },
      ],
    });
  }

  function promptDlg(message, def) {
    const id = "prompt-" + nextId++;
    const input = el("input", { type: "text", id, value: def == null ? "" : String(def), autocomplete: "off" });
    return modal({
      title: "Input",
      body: el("div", {}, el("label", { for: id, text: message, style: "display:block;margin-bottom:.4rem" }), input),
      focus: "input",
      onSubmit: (close) => close(input.value),
      buttons: [
        { label: "Cancel", value: null },
        { label: "OK", primary: true, onclick: (close) => close(input.value) },
      ],
      init: () => input.select(),
    }).then((v) => v);
  }

  // ---------- settings ----------
  const defaultSettings = { theme: "system", scale: 100, reducedMotion: false, dyslexic: false, bigFocus: false };
  let settings = Object.assign({}, defaultSettings, lsGet(SETTINGS_KEY, {}));
  function applySettings() {
    const r = document.documentElement;
    r.dataset.theme = settings.theme;
    r.style.setProperty("--ui-scale", String(Math.min(200, Math.max(100, Number(settings.scale) || 100)) / 100));
    r.dataset.reducedMotion = String(!!settings.reducedMotion);
    r.dataset.dyslexic = String(!!settings.dyslexic);
    r.dataset.bigfocus = String(!!settings.bigFocus);
  }
  applySettings();

  function openSettings() {
    const save = () => { lsSet(SETTINGS_KEY, settings); applySettings(); };
    const row = (id, label, hint, control) => el("div", { class: "setting-row" },
      el("label", { for: id }, label, hint ? el("span", { class: "muted", text: hint }) : null), control);
    const theme = el("select", { id: "set-theme", onchange: () => { settings.theme = theme.value; save(); } },
      [["system", "Match system"], ["light", "Light"], ["dark", "Dark"], ["hc", "High contrast"]].map(([v, t]) =>
        el("option", { value: v, selected: settings.theme === v, text: t })));
    const scaleOut = el("output", { for: "set-scale", text: settings.scale + "%" });
    const scale = el("input", {
      type: "range", id: "set-scale", min: 100, max: 200, step: 10, value: settings.scale,
      oninput: () => { settings.scale = Number(scale.value); scaleOut.textContent = scale.value + "%"; scale.setAttribute("aria-valuetext", scale.value + " percent"); save(); },
    });
    const check = (id, key) => {
      const c = el("input", { type: "checkbox", id, onchange: () => { settings[key] = c.checked; save(); } });
      c.checked = !!settings[key];
      return c;
    };
    const body = el("div", {},
      row("set-theme", "Theme", "High contrast uses black, white and yellow", theme),
      row("set-scale", "Interface size", "Makes all text and buttons larger", el("span", {}, scale, " ", scaleOut)),
      row("set-motion", "Reduce motion", "Turns off animations", check("set-motion", "reducedMotion")),
      row("set-dys", "Dyslexia-friendly font", "Rounder font with extra letter spacing", check("set-dys", "dyslexic")),
      row("set-focus", "Large cursor and focus highlight", "Easier to see where you are", check("set-focus", "bigFocus")));
    return modal({
      title: "Settings", body, focus: "#set-theme",
      buttons: [
        { label: "Reset to defaults", onclick: (close) => { settings = Object.assign({}, defaultSettings); save(); close(null); toast("Settings reset"); } },
        { label: "Done", primary: true, value: null },
      ],
    });
  }

  function openHelp() {
    const mod = isMac ? "Cmd" : "Ctrl";
    const rows = [
      [mod + "+N", "New Word document"], [mod + "+O", "Open a file"], [mod + "+S", "Save"],
      [mod + "+Shift+S", "Save as"], [mod + "+W", "Close the current document"],
      ["Ctrl+Tab / Ctrl+Shift+Tab", "Next / previous document"], ["Alt+Home", "Go to the Home screen"],
      ["F1", "Show this help"], ["Left / Right arrows", "Move between document tabs (when a tab has focus)"],
      ["Delete", "Close the focused document tab"], ["Esc", "Close a dialog"],
    ];
    const body = el("div", {},
      el("table", { class: "kbd-table" },
        el("caption", { class: "visually-hidden", text: "Global keyboard shortcuts" }),
        el("thead", {}, el("tr", {}, el("th", { scope: "col", text: "Keys" }), el("th", { scope: "col", text: "Action" }))),
        el("tbody", {}, rows.map(([k, a]) => el("tr", {}, el("td", {}, el("kbd", { text: k })), el("td", { text: a }))))),
      el("p", { class: "muted", style: "margin-top:1rem", text: "Each editor shows its own shortcuts in button tooltips (hover or focus a button). Use the Settings button for themes, larger text and other accessibility options." }));
    return modal({ title: "Keyboard shortcuts", body, wide: true, buttons: [{ label: "Close", primary: true, value: null }] });
  }

  // ---------- file picker ----------
  function pickFile(opts) {
    opts = Object.assign({ mode: "open", extensions: [], suggestedName: "", multiple: false }, opts || {});
    const exts = (opts.extensions || []).map((e) => e.toLowerCase());
    const saving = opts.mode === "save";
    const multiple = !saving && !!opts.multiple;
    let cur = { path: "", parent: "", entries: [] };
    let items = [];          // visible entries
    let active = -1;         // index of focused entry
    const selected = new Set(); // paths (multiple mode)
    let typeBuf = "", typeTimer = 0;
    let closeFn = null;

    const listId = "fp-list-" + nextId++;
    const pathInput = el("input", { type: "text", "aria-label": "Folder path", title: "Type a folder path and press Enter", autocomplete: "off", spellcheck: "false" });
    const upBtn = el("button", { type: "button", class: "btn", title: "Up one folder (Backspace)", "aria-label": "Up one folder" }, icon("arrowUp", 18), el("span", { text: "Up" }));
    const list = el("ul", { class: "fp-list", id: listId, role: "listbox", tabindex: "0", "aria-label": "Files and folders", "aria-multiselectable": multiple ? "true" : null });
    const status = el("p", { class: "visually-hidden", "aria-live": "polite" });
    const places = el("nav", { class: "fp-places", "aria-label": "Places" }, el("h3", { text: "Places" }));
    const nameInput = el("input", { type: "text", id: "fp-name-" + nextId, value: opts.suggestedName || "", autocomplete: "off", spellcheck: "false" });
    const nameRow = saving ? el("div", { class: "fp-filename" }, el("label", { for: nameInput.id, text: "File name" }), nameInput) : null;
    const filterNote = exts.length ? el("p", { class: "muted", style: "margin:0;font-size:.85rem", text: "Showing: " + exts.join(", ") + (multiple ? " · Space selects several files" : "") }) : null;

    const body = el("div", { class: "fp" }, places,
      el("div", { class: "fp-main" }, el("div", { class: "fp-pathbar" }, upBtn, pathInput), list, filterNote, nameRow, status));

    function matches(e) { return e.isDir || !exts.length || exts.includes(extOf(e.name)); }

    function render() {
      list.textContent = "";
      items.forEach((e, i) => {
        const sel = multiple ? selected.has(e.path) : i === active;
        const li = el("li", {
          class: "fp-item" + (i === active ? " active" : ""), role: "option", id: listId + "-" + i,
          "aria-selected": String(sel),
          onclick: (ev) => { setActive(i); if (multiple && !e.isDir && (ev.ctrlKey || ev.metaKey || ev.shiftKey)) toggle(e); else if (multiple && !e.isDir) { selected.clear(); selected.add(e.path); render(); } else if (saving && !e.isDir) nameInput.value = e.name; },
          ondblclick: () => activate(e),
        },
          el("span", { class: "fp-icon", "aria-hidden": "true" }, e.isDir ? icon("folder", { size: 18, className: "ft-folder" }) : iconFor(e.name, 18)),
          el("span", { class: "fp-name", text: e.name }),
          el("span", { class: "fp-meta", text: e.isDir ? "Folder" : fmtSize(e.size || 0) }),
          el("span", { class: "fp-meta", text: e.modTime ? fmtDate(e.modTime) : "" }));
        list.append(li);
      });
      if (!items.length) list.append(el("li", { class: "fp-empty", role: "presentation", text: "This folder has no matching files." }));
      if (active >= 0) {
        list.setAttribute("aria-activedescendant", listId + "-" + active);
        const a = $(listId + "-" + active);
        if (a) a.scrollIntoView({ block: "nearest" });
      } else list.removeAttribute("aria-activedescendant");
    }
    function setActive(i) {
      active = Math.max(-1, Math.min(items.length - 1, i));
      if (saving && active >= 0 && !items[active].isDir) nameInput.value = items[active].name;
      render();
    }
    function toggle(e) {
      if (selected.has(e.path)) selected.delete(e.path); else selected.add(e.path);
      render();
      status.textContent = selected.size + " file" + (selected.size === 1 ? "" : "s") + " selected";
    }

    async function navigate(path, focusList) {
      try {
        const res = await api.get("/api/fs/list?path=" + encodeURIComponent(path || ""));
        cur = await res.json();
      } catch (err) {
        toast("Cannot open folder: " + err.message, "error");
        return;
      }
      items = (cur.entries || []).filter(matches);
      active = items.length ? 0 : -1;
      selected.clear();
      pathInput.value = cur.path || "";
      upBtn.disabled = !cur.path || cur.parent === cur.path;
      if (cur.path) lsSet(LASTDIR_KEY, cur.path);
      render();
      status.textContent = (cur.path || "Computer") + ", " + items.length + " items";
      if (focusList) list.focus();
    }
    function goUp() { if (!upBtn.disabled) navigate(cur.parent || "", true); }

    async function finish() {
      if (saving) {
        let name = nameInput.value.trim();
        if (!name) { toast("Please type a file name.", "error"); nameInput.focus(); return; }
        if (/[\\/]/.test(name) || /^[a-zA-Z]:/.test(name)) { closeFn(name); return; } // full path typed
        if (!cur.path) { toast("Choose a folder first.", "error"); return; }
        if (exts.length && !exts.includes(extOf(name))) name += exts[0];
        const existing = (cur.entries || []).find((e) => e.name.toLowerCase() === name.toLowerCase());
        if (existing && existing.isDir) { navigate(existing.path, true); return; }
        if (existing && !(await confirmDlg('"' + name + '" already exists. Replace it?', { okLabel: "Replace", danger: true, title: "Replace file?" }))) return;
        closeFn(joinPath(cur.path, name));
        return;
      }
      const a = items[active];
      if (multiple) {
        const files = items.filter((e) => selected.has(e.path)).map((e) => e.path);
        if (files.length) { closeFn(files); return; }
        if (a && !a.isDir) { closeFn([a.path]); return; }
      } else if (a && !a.isDir) { closeFn(a.path); return; }
      if (a && a.isDir) { navigate(a.path, true); return; }
      toast("Select a file first.", "error");
    }
    function activate(e) {
      if (e.isDir) navigate(e.path, true);
      else if (saving) { nameInput.value = e.name; finish(); }
      else if (multiple) { selected.add(e.path); finish(); }
      else closeFn(e.path);
    }

    list.addEventListener("keydown", (ev) => {
      const k = ev.key;
      if (k === "ArrowDown") { setActive(active + 1); }
      else if (k === "ArrowUp") { setActive(Math.max(0, active - 1)); }
      else if (k === "Home") { setActive(0); }
      else if (k === "End") { setActive(items.length - 1); }
      else if (k === "PageDown") { setActive(active + 10); }
      else if (k === "PageUp") { setActive(Math.max(0, active - 10)); }
      else if (k === "Enter") { if (items[active]) activate(items[active]); }
      else if (k === "Backspace") { goUp(); }
      else if (k === " " && multiple) { if (items[active] && !items[active].isDir) toggle(items[active]); else if (items[active]) activate(items[active]); }
      else if (k.length === 1 && !ev.ctrlKey && !ev.metaKey && !ev.altKey) {
        clearTimeout(typeTimer);
        typeBuf += k.toLowerCase();
        typeTimer = setTimeout(() => (typeBuf = ""), 800);
        const i = items.findIndex((e) => e.name.toLowerCase().startsWith(typeBuf));
        if (i >= 0) setActive(i);
      } else return;
      ev.preventDefault();
    });
    list.addEventListener("focus", () => { if (active < 0 && items.length) setActive(0); });
    pathInput.addEventListener("keydown", (ev) => { if (ev.key === "Enter") { ev.preventDefault(); navigate(pathInput.value.trim(), true); } });
    nameInput.addEventListener("keydown", (ev) => { if (ev.key === "Enter") { ev.preventDefault(); finish(); } });
    upBtn.addEventListener("click", goUp);

    api.get("/api/fs/home").then((r) => r.json()).then((h) => {
      const add = (name, path, glyph) => places.append(el("button", { type: "button", title: path || name, onclick: () => navigate(path, true) }, icon(glyph, 18), el("span", { text: name })));
      add("Home", h.path, "home");
      (h.places || []).forEach((p) => add(p.name, p.path, p.name === "Downloads" ? "download" : p.name === "Desktop" ? "desktop" : "folder"));
      if (/^[a-zA-Z]:/.test(h.path || "")) add("This PC", "", "computer");
      const last = lsGet(LASTDIR_KEY, "");
      navigate(last || h.path, false).then(() => { if (!cur.path && last) navigate(h.path, false); });
    }).catch((err) => toast("File browser unavailable: " + err.message, "error"));

    const title = saving ? "Save as" : multiple ? "Choose files" : "Open file";
    return modal({
      title, body, wide: true, focus: saving ? "input.fp-name, .fp-filename input" : ".fp-list",
      init: (dlg, close) => { closeFn = close; if (saving) { const dot = nameInput.value.lastIndexOf("."); nameInput.setSelectionRange(0, dot > 0 ? dot : nameInput.value.length); } },
      buttons: [
        { label: "Cancel", value: null },
        { label: saving ? "Save" : "Open", primary: true, onclick: () => finish() },
      ],
    });
  }

  // ---------- recent files ----------
  function getRecent() { return lsGet(RECENT_KEY, []).filter((r) => r && r.path); }
  function addRecent(path) {
    if (!path) return;
    const list = getRecent().filter((r) => r.path.toLowerCase() !== path.toLowerCase());
    list.unshift({ path, name: baseName(path), time: new Date().toISOString() });
    lsSet(RECENT_KEY, list.slice(0, 15));
    renderRecent();
  }
  function removeRecent(path) {
    lsSet(RECENT_KEY, getRecent().filter((r) => r.path !== path));
    renderRecent();
  }
  function renderRecent() {
    const ul = $("recent-list");
    if (!ul) return;
    const list = getRecent();
    ul.textContent = "";
    $("recent-empty").hidden = list.length > 0;
    for (const r of list) {
      ul.append(el("li", { class: "recent-item" },
        el("button", { type: "button", class: "recent-open", title: "Open " + r.path, onclick: () => openPath(r.path) },
          el("span", { class: "recent-icon", "aria-hidden": "true" }, iconFor(r.name, 20)),
          el("span", { class: "recent-text" },
            el("span", { class: "recent-name", text: r.name }),
            el("span", { class: "recent-path", text: r.path })),
          el("span", { class: "recent-date", text: fmtDate(r.time) })),
        el("button", { type: "button", class: "icon-btn recent-remove", "aria-label": "Remove " + r.name + " from recent files", title: "Remove from list", onclick: () => removeRecent(r.path) }, icon("close", 16))));
    }
  }

  // ---------- editors & documents ----------
  function registerEditor(kind, def) {
    editors[kind] = Object.assign({ label: kind, extensions: [], icon: "doc" }, def);
    if (kind === "slides") { const t = $("tile-new-slides"); if (t) t.hidden = false; }
  }
  function kindForName(name) {
    const e = extOf(name);
    for (const [kind, def] of Object.entries(editors)) if ((def.extensions || []).map((x) => x.toLowerCase()).includes(e)) return kind;
    return null;
  }
  const findDoc = (editor) => docs.find((d) => d.editor === editor);
  const docById = (id) => docs.find((d) => d.id === id);
  const activeDoc = () => docById(activeId);

  function setDirty(editor, dirty) {
    const d = findDoc(editor);
    if (!d) return;
    d.dirty = !!dirty;
    d.dirtyMark.hidden = !d.dirty;
    d.tab.setAttribute("aria-label", d.name + (d.dirty ? ", unsaved" : ""));
  }
  function setTitle(editor, name) {
    const d = findDoc(editor);
    if (!d || !name) return;
    d.name = name;
    d.nameEl.textContent = name;
    d.tab.title = name + " (Ctrl+W or Delete closes)";
    d.tab.setAttribute("aria-label", name + (d.dirty ? ", unsaved" : ""));
    d.closeBtn.setAttribute("aria-label", "Close " + name);
    if (activeId === d.id) document.title = name + " – DocStudio";
  }

  function tabButtons() { return [$("tab-home"), ...docs.map((d) => d.tab)]; }

  function activate(id, focusEditor) {
    activeId = id;
    const home = $("home");
    home.hidden = id !== null;
    $("tab-home").setAttribute("aria-selected", String(id === null));
    $("tab-home").tabIndex = id === null ? 0 : -1;
    for (const d of docs) {
      const on = d.id === id;
      d.panel.hidden = !on;
      d.tab.setAttribute("aria-selected", String(on));
      d.tab.tabIndex = on ? 0 : -1;
      d.wrap.classList.toggle("active", on);
    }
    const d = activeDoc();
    document.title = d ? d.name + " – DocStudio" : "DocStudio";
    if (d) d.tab.scrollIntoView({ block: "nearest", inline: "nearest" });
    if (focusEditor && d && d.editor && d.editor.focus) { try { d.editor.focus(); } catch (_) { /* ignore */ } }
    if (id === null) renderRecent();
  }

  async function openDocument(kind, doc) {
    const def = editors[kind];
    if (!def) { toast("No editor available for this file type yet.", "error"); return null; }
    const id = nextId++;
    const name = doc.name || "Untitled";
    const panel = el("section", { role: "tabpanel", id: "doc-panel-" + id, "aria-labelledby": "doc-tab-" + id, class: "doc-panel", hidden: true });
    const nameEl = el("span", { class: "doc-tab-name", text: name });
    const dirtyMark = el("span", { class: "doc-tab-dirty", "aria-hidden": "true", text: "•", hidden: true });
    const tab = el("button", {
      type: "button", role: "tab", class: "doc-tab", id: "doc-tab-" + id, "aria-controls": panel.id,
      "aria-selected": "false", "aria-label": name, title: name + " (Ctrl+W or Delete closes)", tabindex: "-1",
    }, el("span", { class: "doc-tab-icon", "aria-hidden": "true" }, kindIcon(kind, def)), nameEl, dirtyMark);
    const closeBtn = el("button", { type: "button", class: "doc-tab-close", tabindex: "-1", "aria-label": "Close " + name, title: "Close (Ctrl+W)" }, icon("close", 14));
    const wrap = el("div", { class: "doc-tab-wrap", role: "presentation", "data-kind": kind }, tab, closeBtn);
    const d = { id, kind, def, editor: null, name, path: doc.path || "", dirty: false, tab, wrap, panel, nameEl, dirtyMark, closeBtn };
    tab.addEventListener("click", () => activate(id, true));
    closeBtn.addEventListener("click", () => closeDoc(id));
    tab.addEventListener("auxclick", (e) => { if (e.button === 1) closeDoc(id); });
    $("doc-tabs").append(wrap);
    $("main").append(panel);
    docs.push(d);
    activate(id, false);
    try {
      d.editor = await def.create(panel, { name, path: doc.path || "", bytes: doc.bytes || null });
    } catch (err) {
      console.error(err);
      toast("Could not open " + name + ": " + (err && err.message ? err.message : err), "error");
      removeDoc(d);
      return null;
    }
    if (doc.path) addRecent(doc.path);
    activate(id, true);
    return d.editor;
  }

  function removeDoc(d) {
    const i = docs.indexOf(d);
    if (i < 0) return;
    docs.splice(i, 1);
    d.wrap.remove();
    d.panel.remove();
    const next = docs[i] || docs[i - 1];
    activate(next ? next.id : null, true);
    if (!next) $("tab-home").focus();
  }

  async function closeDoc(id) {
    const d = docById(id);
    if (!d) return;
    if (d.dirty) {
      activate(id, false);
      const choice = await modal({
        title: "Unsaved changes",
        body: el("p", { text: '"' + d.name + '" has unsaved changes. Save them before closing?' }),
        cancelValue: "cancel",
        buttons: [{ label: "Cancel", value: "cancel" }, { label: "Don't save", value: "discard", danger: true }, { label: "Save", value: "save", primary: true }],
      });
      if (choice === "cancel") return;
      if (choice === "save") {
        try { await d.editor.save(); } catch (err) { toast("Save failed: " + err.message, "error"); return; }
        if (d.dirty) return; // save was cancelled
      }
    }
    try { if (d.editor && d.editor.destroy) d.editor.destroy(); } catch (err) { console.error(err); }
    removeDoc(d);
  }

  function newDocument(kind) {
    kind = kind || "word";
    const def = editors[kind];
    if (!def) { toast(kind === "slides" ? "The presentation editor is not available in this build." : "No editor available for " + kind + ".", "error"); return Promise.resolve(null); }
    const ext = def.extensions && def.extensions[0] ? def.extensions[0] : ".docx";
    const base = kind === "slides" ? "Presentation " : kind === "pdf" ? "PDF " : "Document ";
    return openDocument(kind, { name: base + untitled++ + ext, path: "", bytes: null });
  }

  async function imagesToPdf(files) {
    // files: [{name, bytes}]
    const form = new FormData();
    for (const f of files) form.append("file", new Blob([f.bytes], { type: extOf(f.name) === ".png" ? "image/png" : "image/jpeg" }), f.name);
    const res = await api.postForm("/api/pdf/images-to-pdf", form);
    return new Uint8Array(await res.arrayBuffer());
  }

  async function openBytes(name, path, bytes) {
    try {
      if (IMAGE_EXTS.includes(extOf(name))) {
        const pdf = await imagesToPdf([{ name, bytes }]);
        return openDocument("pdf", { name: name.replace(/\.[^.]+$/, "") + ".pdf", path: "", bytes: pdf });
      }
      const kind = kindForName(name);
      if (!kind) { toast('Cannot open "' + name + '": unsupported file type.', "error"); return null; }
      return openDocument(kind, { name, path, bytes });
    } catch (err) {
      toast("Could not open " + name + ": " + err.message, "error");
      return null;
    }
  }

  async function openPath(path) {
    const existing = docs.find((d) => d.path && d.path.toLowerCase() === path.toLowerCase());
    if (existing) { activate(existing.id, true); return existing.editor; }
    let bytes;
    try { bytes = await readFile(path); } catch (err) {
      toast("Could not read " + path + ": " + err.message, "error");
      return null;
    }
    return openBytes(baseName(path), path, bytes);
  }

  function allExtensions() {
    const s = new Set(IMAGE_EXTS);
    for (const def of Object.values(editors)) (def.extensions || []).forEach((e) => s.add(e.toLowerCase()));
    return [...s];
  }

  async function openDialog() {
    const p = await pickFile({ mode: "open", extensions: allExtensions() });
    if (p) await openPath(p);
  }

  async function pdfFromImagesDialog() {
    const paths = await pickFile({ mode: "open", extensions: IMAGE_EXTS, multiple: true });
    if (!paths || !paths.length) return;
    try {
      const files = await Promise.all(paths.map(async (p) => ({ name: baseName(p), bytes: await readFile(p) })));
      const pdf = await imagesToPdf(files);
      await openDocument("pdf", { name: "Images.pdf", path: "", bytes: pdf });
      toast("Created a PDF from " + files.length + " image" + (files.length === 1 ? "" : "s"), "success");
    } catch (err) { toast("Could not create PDF: " + err.message, "error"); }
  }

  async function mergeDialog() {
    const paths = await pickFile({ mode: "open", extensions: [".pdf"], multiple: true });
    if (!paths || !paths.length) return;
    if (paths.length < 2) { toast("Select at least two PDFs to merge (Space or Ctrl+click selects several).", "error"); return; }
    try {
      const form = new FormData();
      for (const p of paths) form.append("file", new Blob([await readFile(p)], { type: "application/pdf" }), baseName(p));
      const res = await api.postForm("/api/pdf/merge", form);
      await openDocument("pdf", { name: "Merged.pdf", path: "", bytes: new Uint8Array(await res.arrayBuffer()) });
      toast("Merged " + paths.length + " PDFs. Save to keep the result.", "success");
    } catch (err) { toast("Merge failed: " + err.message, "error"); }
  }

  // ---------- drag and drop ----------
  function setupDrop() {
    const overlay = $("drop-overlay");
    let depth = 0;
    const hasFiles = (e) => e.dataTransfer && [...e.dataTransfer.types].includes("Files");
    window.addEventListener("dragenter", (e) => { if (!hasFiles(e)) return; depth++; overlay.hidden = false; });
    window.addEventListener("dragleave", (e) => { if (!hasFiles(e)) return; depth = Math.max(0, depth - 1); if (!depth) overlay.hidden = true; });
    window.addEventListener("dragover", (e) => { if (hasFiles(e)) { e.preventDefault(); e.dataTransfer.dropEffect = "copy"; } });
    window.addEventListener("drop", async (e) => {
      if (!hasFiles(e)) return;
      depth = 0;
      overlay.hidden = true;
      // Let editors that handle drops themselves (e.g. inserting images) keep them.
      if (e.defaultPrevented) return;
      e.preventDefault();
      const files = [...e.dataTransfer.files];
      if (!files.length) return;
      const loaded = await Promise.all(files.map(async (f) => ({ name: f.name, bytes: new Uint8Array(await f.arrayBuffer()) })));
      const images = loaded.filter((f) => IMAGE_EXTS.includes(extOf(f.name)));
      if (images.length > 1) {
        try {
          await openDocument("pdf", { name: "Images.pdf", path: "", bytes: await imagesToPdf(images) });
        } catch (err) { toast("Could not create PDF: " + err.message, "error"); }
      }
      for (const f of loaded) {
        if (images.length > 1 && images.includes(f)) continue;
        await openBytes(f.name, "", f.bytes);
      }
    });
  }

  // ---------- keyboard ----------
  function setupKeys() {
    window.addEventListener("keydown", (e) => {
      if (document.querySelector("dialog[open]")) return;
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      const d = activeDoc();
      let handled = true;
      if (mod && !e.shiftKey && !e.altKey && k === "o") openDialog();
      else if (mod && !e.shiftKey && !e.altKey && k === "n") newDocument("word");
      else if (mod && !e.altKey && k === "s") {
        if (d && d.editor) (e.shiftKey ? d.editor.saveAs() : d.editor.save()).catch((err) => toast("Save failed: " + err.message, "error"));
        else toast("Open a document to save it.");
      }
      else if (mod && !e.shiftKey && !e.altKey && (k === "w" || k === "f4")) { if (d) closeDoc(d.id); }
      else if (e.ctrlKey && k === "tab") cycle(e.shiftKey ? -1 : 1);
      else if (e.altKey && !mod && e.key === "Home") { activate(null); $("tab-home").focus(); }
      else if (e.key === "F1") openHelp();
      else handled = false;
      if (handled) { e.preventDefault(); e.stopPropagation(); }
    }, true);

    // Roving focus inside the document tab list.
    $("doc-tabs").addEventListener("keydown", (e) => {
      const tabs = tabButtons();
      const i = tabs.indexOf(document.activeElement);
      if (i < 0) return;
      let j = -1;
      if (e.key === "ArrowRight") j = (i + 1) % tabs.length;
      else if (e.key === "ArrowLeft") j = (i - 1 + tabs.length) % tabs.length;
      else if (e.key === "Home") j = 0;
      else if (e.key === "End") j = tabs.length - 1;
      else if (e.key === "Delete" && i > 0) { e.preventDefault(); closeDoc(docs[i - 1].id); return; }
      else return;
      e.preventDefault();
      tabs[j].click();
      tabs[j].focus();
    });
  }
  function cycle(dir) {
    const ids = [null, ...docs.map((d) => d.id)];
    const i = ids.indexOf(activeId);
    activate(ids[(i + dir + ids.length) % ids.length], true);
  }

  // ---------- startup ----------
  function init() {
    $("tab-home").addEventListener("click", () => activate(null));
    $("btn-new").addEventListener("click", () => newDocument("word"));
    $("btn-open").addEventListener("click", openDialog);
    $("btn-settings").addEventListener("click", openSettings);
    $("btn-help").addEventListener("click", openHelp);
    hydrateIcons();
    $("tile-new-doc").addEventListener("click", () => newDocument("word"));
    const slidesTile = $("tile-new-slides");
    if (slidesTile) {
      slidesTile.hidden = !editors.slides;
      slidesTile.addEventListener("click", () => newDocument("slides"));
    }
    $("tile-open").addEventListener("click", openDialog);
    $("tile-images").addEventListener("click", pdfFromImagesDialog);
    $("tile-merge").addEventListener("click", mergeDialog);
    setupDrop();
    setupKeys();
    renderRecent();
    activate(null);
    window.addEventListener("beforeunload", (e) => {
      if (docs.some((d) => d.dirty)) { e.preventDefault(); e.returnValue = ""; }
    });
    // Heartbeat: the server shuts down as soon as the last window closes.
    const winId = Math.random().toString(36).slice(2) + Date.now().toString(36);
    const ping = () => fetch("/api/ping?id=" + winId, { method: "POST" }).catch(() => {});
    ping();
    setInterval(ping, 10000);
    window.addEventListener("pagehide", () => navigator.sendBeacon("/api/bye?id=" + winId));
    window.addEventListener("pageshow", (e) => { if (e.persisted) ping(); });
    const m = /^#open=(.+)$/.exec(location.hash);
    if (m) {
      let p = m[1];
      try { p = decodeURIComponent(p); } catch (_) { /* keep raw */ }
      if (/^[a-zA-Z]:\//.test(p)) p = p.replace(/\//g, "\\");
      history.replaceState(null, "", location.pathname);
      openPath(p);
    }
  }

  window.App = {
    registerEditor,
    setDirty,
    setTitle,
    api,
    readFile,
    writeFile,
    pickFile,
    toast,
    confirm: confirmDlg,
    prompt: promptDlg,
    download,
    // Extras (not required by the contract, but handy for editors):
    modal,            // modal({title, body, buttons:[{label,value,primary,danger}]}) -> Promise
    openPath,         // open a file from disk in a new tab
    openBytes,        // openBytes(name, path, Uint8Array) opens in-memory content in a new tab
    newDocument,      // newDocument(kind)
    addRecent,
    el,               // el(tag, attrs, ...children) DOM helper
    icon,             // icon(name, {size, title, className}) -> SVGElement (see ICONS list above)
    iconHTML,         // iconHTML(name, size) -> SVG markup string
    hasIcon,          // hasIcon(name) -> bool
    iconNames: Object.keys(ICONS),
  };

  if (document.readyState === "complete") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
