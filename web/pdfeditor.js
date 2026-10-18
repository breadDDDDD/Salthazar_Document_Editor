/* DocStudio PDF editor (Acrobat-like): viewer + annotation tools + page organizer.
 *
 * Rendering uses the vendored pdf.js 3.11 (window.pdfjsLib). All document changes are done
 * server side through /api/pdf/* (see the board API contract); this file only keeps the
 * current PDF bytes plus a list of editable overlay objects that get "flattened" into the
 * PDF with /api/pdf/annotate before any page operation, export or save.
 *
 * Overlay object units: PDF points in the coordinate space of the displayed page at scale 1
 * (origin top-left of the page as shown, including the page's own /Rotate and the view rotation).
 */
(function () {
  "use strict";
  const App = window.App;
  const pdfjsLib = window.pdfjsLib;
  if (!App) { console.error("pdfeditor.js: window.App is missing"); return; }
  if (pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = "vendor/pdfjs/pdf.worker.min.js";

  const PX = 96 / 72; // CSS pixels per PDF point at 100 % zoom
  const ZOOMS = [0.25, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];
  const SIG_KEY = "docstudio.signature";
  const FONT_CSS = {
    Helvetica: "Helvetica, Arial, sans-serif",
    "Times-Roman": "'Times New Roman', Times, serif",
    Courier: "'Courier New', Courier, monospace",
  };
  // distance from line top to first baseline, as a fraction of the font size, with line-height 1.2
  const BASELINE = { Helvetica: 0.947, "Times-Roman": 0.938, Courier: 0.866 };
  const STAMPS = {
    approved: { label: "Approved", text: "APPROVED", color: "#1e7a34" },
    draft: { label: "Draft", text: "DRAFT", color: "#1f5fbf" },
    confidential: { label: "Confidential", text: "CONFIDENTIAL", color: "#b3261e" },
    date: { label: "Today's date", text: null, color: "#37474f" },
  };
  const TOOLS = [
    { id: "select", ico: "⬉", label: "Select", key: "V", hint: "Select: click an item to select it. Drag or use arrow keys to move it, Alt+arrows to resize, Delete to remove." },
    { id: "hand", ico: "✋", label: "Hand", key: "M", hint: "Hand: drag to scroll the pages. Esc returns to Select." },
    { id: "text", ico: "T", label: "Add text", key: "T", edit: true, hint: "Click on the page to add text. Esc to cancel." },
    { id: "highlight", ico: "🖍", label: "Highlight", key: "H", edit: true, hint: "Drag over the page to highlight an area. Esc to cancel." },
    { id: "draw", ico: "✎", label: "Draw", key: "D", edit: true, hint: "Drag on the page to draw freehand. Esc to stop drawing." },
    { id: "whiteout", ico: "▭", label: "Whiteout", key: "W", edit: true, extra: " To change existing text: whiteout it, then Add text on top.", hint: "Drag to cover existing content with white. To replace text, whiteout it and then use Add text. Esc to cancel." },
    { id: "note", ico: "🗨", label: "Note", key: "N", edit: true, hint: "Click on the page to add a sticky note. Esc to cancel." },
    { id: "rect", ico: "□", label: "Rectangle", key: "R", edit: true, shape: true, hint: "Drag to draw a rectangle. Esc to cancel." },
    { id: "ellipse", ico: "◯", label: "Ellipse", key: "E", edit: true, shape: true, hint: "Drag to draw an ellipse. Esc to cancel." },
    { id: "line", ico: "╱", label: "Line", key: "L", edit: true, shape: true, hint: "Drag to draw a line. Esc to cancel." },
    { id: "arrow", ico: "↗", label: "Arrow", key: "A", edit: true, shape: true, hint: "Drag to draw an arrow. Esc to cancel." },
  ];
  const TYPE_NAMES = { text: "Text box", highlight: "Highlight", ink: "Drawing", whiteout: "Whiteout", note: "Sticky note", rect: "Rectangle", ellipse: "Ellipse", line: "Line", arrow: "Arrow", image: "Image", stamp: "Stamp" };

  // ---------- helpers ----------
  function el(tag, attrs, ...kids) {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === "class") e.className = v;
      else if (k === "text") e.textContent = v;
      else if (k.startsWith("on") && typeof v === "function") e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? "" : v);
    }
    for (const c of kids.flat()) if (c != null && c !== false) e.append(c);
    return e;
  }
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const baseName = (p) => (p || "").split(/[\\/]/).filter(Boolean).pop() || p || "";
  const stem = (n) => (n || "document").replace(/\.[^.]*$/, "");
  const isTyping = (t) => !!t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName));
  const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (_) { return d; } };
  const lsSet = (k, v) => { try { if (v == null) localStorage.removeItem(k); else localStorage.setItem(k, JSON.stringify(v)); } catch (_) { /* ignore */ } };
  function hexToRgba(hex, a) {
    const m = /^#?([0-9a-f]{6})$/i.exec(hex || "");
    const n = m ? parseInt(m[1], 16) : 0xffeb3b;
    return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
  }
  function blobToDataURL(blob) {
    return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(blob); });
  }
  const imgCache = new Map();
  function loadImage(src) {
    let img = imgCache.get(src);
    if (!img) { img = new Image(); img.src = src; imgCache.set(src, img); }
    if (img.complete && img.naturalWidth) return Promise.resolve(img);
    return new Promise((res, rej) => {
      img.addEventListener("load", () => res(img), { once: true });
      img.addEventListener("error", () => rej(new Error("Could not load the image.")), { once: true });
    });
  }
  // 0-based sorted indices -> "1-3,5"
  function rangeStr(idx) {
    const out = [];
    for (let i = 0; i < idx.length; i++) {
      let j = i;
      while (j + 1 < idx.length && idx[j + 1] === idx[j] + 1) j++;
      out.push(i === j ? String(idx[i] + 1) : `${idx[i] + 1}-${idx[j] + 1}`);
      i = j;
    }
    return out.join(",");
  }
  function pagesLabel(idx) { return idx.length === 1 ? "page " + (idx[0] + 1) : idx.length + " pages"; }
  function blankPdf() {
    const objs = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << >> /Contents 4 0 R >>",
      "<< /Length 0 >>\nstream\n\nendstream",
    ];
    let s = "%PDF-1.4\n";
    const offs = [];
    objs.forEach((o, i) => { offs.push(s.length); s += `${i + 1} 0 obj\n${o}\nendobj\n`; });
    const xref = s.length;
    s += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map((o) => String(o).padStart(10, "0") + " 00000 n \n").join("");
    s += `trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
    return new TextEncoder().encode(s);
  }
  const measureCtx = document.createElement("canvas").getContext("2d");
  function wrapText(ctx, text, maxW) {
    const lines = [];
    for (const para of String(text || "").replace(/\r\n?/g, "\n").split("\n")) {
      const words = para.split(/(\s+)/);
      let line = "";
      for (const w of words) {
        if (!w) continue;
        const test = line + w;
        if (ctx.measureText(test).width <= maxW || !line) {
          if (ctx.measureText(test).width > maxW && !/^\s+$/.test(w)) {
            // single word longer than the line: break by characters
            let part = line;
            for (const ch of w) {
              if (ctx.measureText(part + ch).width > maxW && part) { lines.push(part); part = ""; }
              part += ch;
            }
            line = part;
          } else line = test;
        } else {
          lines.push(line.replace(/\s+$/, ""));
          line = /^\s+$/.test(w) ? "" : w;
        }
      }
      lines.push(line.replace(/\s+$/, ""));
    }
    return lines;
  }
  function roundRectPath(ctx, x, y, w, h, r) {
    ctx.beginPath();
    if (ctx.roundRect) ctx.roundRect(x, y, w, h, r); else ctx.rect(x, y, w, h);
  }

  // Draw an overlay object into a W x H device-pixel box; k = device pixels per point.
  function drawObj(ctx, o, W, H, k) {
    ctx.save();
    const lw = Math.max(1, (o.width || 2) * k);
    ctx.lineWidth = lw;
    ctx.strokeStyle = o.color || "#000000";
    ctx.fillStyle = o.color || "#000000";
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
    switch (o.type) {
      case "highlight":
        ctx.fillStyle = hexToRgba(o.color || "#ffeb3b", 0.4);
        ctx.fillRect(0, 0, W, H);
        break;
      case "whiteout":
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, W, H);
        break;
      case "rect":
        ctx.strokeRect(lw / 2, lw / 2, Math.max(0, W - lw), Math.max(0, H - lw));
        break;
      case "ellipse":
        ctx.beginPath();
        ctx.ellipse(W / 2, H / 2, Math.max(1, W / 2 - lw / 2), Math.max(1, H / 2 - lw / 2), 0, 0, Math.PI * 2);
        ctx.stroke();
        break;
      case "ink": case "line": case "arrow": {
        const p = (o.pts || []).map(([x, y]) => [x * W, y * H]);
        if (!p.length) break;
        ctx.beginPath();
        p.forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)));
        if (p.length === 1) ctx.lineTo(p[0][0] + 0.1, p[0][1]);
        ctx.stroke();
        if (o.type === "arrow" && p.length >= 2) {
          const [x1, y1] = p[0], [x2, y2] = p[p.length - 1];
          const a = Math.atan2(y2 - y1, x2 - x1), L = (10 + (o.width || 2) * 3) * k;
          ctx.beginPath();
          ctx.moveTo(x2, y2);
          ctx.lineTo(x2 - L * Math.cos(a - 0.45), y2 - L * Math.sin(a - 0.45));
          ctx.lineTo(x2 - L * Math.cos(a + 0.45), y2 - L * Math.sin(a + 0.45));
          ctx.closePath();
          ctx.fill();
        }
        break;
      }
      case "image": {
        const img = imgCache.get(o.src);
        if (img && img.complete && img.naturalWidth) ctx.drawImage(img, 0, 0, W, H);
        break;
      }
      case "stamp": {
        const b = 3 * k;
        ctx.lineWidth = b;
        roundRectPath(ctx, b / 2, b / 2, W - b, H - b, 6 * k);
        ctx.stroke();
        let fs = H * 0.5;
        ctx.font = `bold ${fs}px Helvetica, Arial, sans-serif`;
        const tw = ctx.measureText(o.text).width;
        if (tw > W * 0.86) { fs *= (W * 0.86) / tw; ctx.font = `bold ${fs}px Helvetica, Arial, sans-serif`; }
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(o.text, W / 2, H / 2 + fs * 0.04);
        break;
      }
      case "note": {
        ctx.fillStyle = "#fff59d";
        ctx.fillRect(0, 0, W, H);
        ctx.lineWidth = Math.max(1, k);
        ctx.strokeStyle = "#b59f00";
        ctx.strokeRect(ctx.lineWidth / 2, ctx.lineWidth / 2, W - ctx.lineWidth, H - ctx.lineWidth);
        const fs = (o.fontSize || 10) * k, pad = 4 * k;
        ctx.font = `${fs}px Helvetica, Arial, sans-serif`;
        ctx.fillStyle = "#000000";
        ctx.beginPath(); ctx.rect(0, 0, W, H); ctx.clip();
        wrapText(ctx, o.text, W - pad * 2).forEach((ln, i) => ctx.fillText(ln, pad, pad + fs * BASELINE.Helvetica + i * fs * 1.2));
        break;
      }
      case "text": {
        const fs = (o.fontSize || 12) * k;
        ctx.font = `${fs}px ${FONT_CSS[o.font] || FONT_CSS.Helvetica}`;
        const base = BASELINE[o.font] || BASELINE.Helvetica;
        wrapText(ctx, o.text, W).forEach((ln, i) => ctx.fillText(ln, 0, fs * base + i * fs * 1.2));
        break;
      }
    }
    ctx.restore();
  }

  function trimCanvas(c) {
    const ctx = c.getContext("2d");
    const { data, width, height } = ctx.getImageData(0, 0, c.width, c.height);
    let x0 = width, y0 = height, x1 = -1, y1 = -1;
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
      if (data[(y * width + x) * 4 + 3] > 8) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    }
    if (x1 < 0) return null;
    const pad = 4, out = document.createElement("canvas");
    out.width = x1 - x0 + 1 + pad * 2; out.height = y1 - y0 + 1 + pad * 2;
    out.getContext("2d").drawImage(c, x0, y0, x1 - x0 + 1, y1 - y0 + 1, pad, pad, x1 - x0 + 1, y1 - y0 + 1);
    return out.toDataURL("image/png");
  }

  let uidSeq = 0;

  // ---------- dialogs ----------
  function passwordDialog(title, message, confirmTwice) {
    if (!App.modal) return App.prompt(message, "");
    const n = ++uidSeq;
    const a = el("input", { type: "password", id: "pdfe-pw1-" + n, autocomplete: confirmTwice ? "new-password" : "current-password" });
    const b = confirmTwice ? el("input", { type: "password", id: "pdfe-pw2-" + n, autocomplete: "new-password" }) : null;
    const err = el("p", { class: "pdfe-error", role: "alert" });
    const body = el("div", { class: "pdfe-form" },
      el("p", { text: message }),
      el("div", { class: "pdfe-form-row" }, el("label", { for: a.id, text: "Password" }), a),
      b && el("div", { class: "pdfe-form-row" }, el("label", { for: b.id, text: "Confirm password" }), b),
      err);
    return App.modal({
      title, body, cancelValue: null, focus: "input",
      buttons: [{ label: "Cancel", value: null }, { label: "OK", primary: true, submit: true }],
      onSubmit: (close) => {
        if (!a.value) { err.textContent = "Please enter a password."; a.focus(); return; }
        if (b && a.value !== b.value) { err.textContent = "The passwords do not match."; b.focus(); return; }
        close(a.value);
      },
    });
  }

  function signatureDialog() {
    if (!App.modal) { App.toast("Signature dialog is not available.", "error"); return Promise.resolve(null); }
    const n = ++uidSeq;
    const saved = lsGet(SIG_KEY, null);
    let mode = saved ? "saved" : "draw";
    const pad = el("canvas", { width: 900, height: 300, class: "pdfe-sigpad", role: "img", "aria-label": "Signature drawing area. Draw with a mouse, pen or finger. Keyboard users can use the Type tab." });
    const pctx = pad.getContext("2d");
    const ink = el("input", { type: "color", value: "#1a237e", id: "pdfe-sig-ink-" + n, class: "tool-input" });
    let drawing = false, last = null;
    const pos = (e) => { const r = pad.getBoundingClientRect(); return [(e.clientX - r.left) * pad.width / r.width, (e.clientY - r.top) * pad.height / r.height]; };
    pad.addEventListener("pointerdown", (e) => { drawing = true; last = pos(e); pad.setPointerCapture(e.pointerId); e.preventDefault(); });
    pad.addEventListener("pointermove", (e) => {
      if (!drawing) return;
      const p = pos(e);
      pctx.strokeStyle = ink.value; pctx.lineWidth = 5; pctx.lineCap = "round"; pctx.lineJoin = "round";
      pctx.beginPath(); pctx.moveTo(last[0], last[1]); pctx.lineTo(p[0], p[1]); pctx.stroke();
      last = p;
    });
    const stop = () => { drawing = false; };
    pad.addEventListener("pointerup", stop);
    pad.addEventListener("pointercancel", stop);

    const nameIn = el("input", { type: "text", id: "pdfe-sig-name-" + n, autocomplete: "name", placeholder: "Your name" });
    const fontSel = el("select", { id: "pdfe-sig-font-" + n, class: "tool-select" },
      ["Segoe Script", "Brush Script MT", "Lucida Handwriting", "Ink Free", "Comic Sans MS"].map((f) => el("option", { value: f, text: f })));
    const preview = el("canvas", { width: 900, height: 220, class: "pdfe-sigpad", role: "img", "aria-label": "Preview of the typed signature" });
    const typed = () => {
      const c = document.createElement("canvas"); c.width = 1400; c.height = 260;
      const ctx = c.getContext("2d");
      ctx.font = `110px "${fontSel.value}", cursive`; ctx.fillStyle = ink.value; ctx.textBaseline = "alphabetic";
      ctx.fillText(nameIn.value, 20, 180);
      return c;
    };
    const drawPreview = () => {
      const ctx = preview.getContext("2d");
      ctx.clearRect(0, 0, preview.width, preview.height);
      if (!nameIn.value.trim()) return;
      const c = typed();
      const s = Math.min(1, preview.width / c.width);
      ctx.drawImage(c, 0, 0, c.width * s, c.height * s);
    };
    nameIn.addEventListener("input", drawPreview);
    fontSel.addEventListener("change", drawPreview);
    ink.addEventListener("input", drawPreview);
    const remember = el("input", { type: "checkbox", id: "pdfe-sig-rem-" + n, checked: true });

    const panels = {
      draw: el("div", { role: "tabpanel", id: `pdfe-sig-p-draw-${n}`, "aria-labelledby": `pdfe-sig-t-draw-${n}`, class: "pdfe-sig-panel" },
        pad, el("div", { class: "pdfe-row" },
          el("button", { type: "button", class: "btn", text: "Clear", onclick: () => pctx.clearRect(0, 0, pad.width, pad.height) }))),
      type: el("div", { role: "tabpanel", id: `pdfe-sig-p-type-${n}`, "aria-labelledby": `pdfe-sig-t-type-${n}`, class: "pdfe-sig-panel" },
        el("div", { class: "pdfe-form-row" }, el("label", { for: nameIn.id, text: "Type your name" }), nameIn),
        el("div", { class: "pdfe-form-row" }, el("label", { for: fontSel.id, text: "Handwriting style" }), fontSel),
        preview),
      saved: saved ? el("div", { role: "tabpanel", id: `pdfe-sig-p-saved-${n}`, "aria-labelledby": `pdfe-sig-t-saved-${n}`, class: "pdfe-sig-panel" },
        el("img", { src: saved, alt: "Your saved signature", class: "pdfe-sig-saved" }),
        el("div", { class: "pdfe-row" }, el("button", {
          type: "button", class: "btn danger", text: "Forget saved signature",
          onclick: (e) => { lsSet(SIG_KEY, null); e.target.disabled = true; App.toast("Saved signature removed from this computer.", "success"); },
        }))) : null,
    };
    const order = saved ? ["saved", "draw", "type"] : ["draw", "type"];
    const labels = { saved: "Saved", draw: "Draw", type: "Type" };
    const tabs = {};
    const tablist = el("div", { role: "tablist", class: "pdfe-sig-tabs", "aria-label": "Signature method" });
    const show = (m, focus) => {
      mode = m;
      for (const k of order) {
        tabs[k].setAttribute("aria-selected", String(k === m));
        tabs[k].tabIndex = k === m ? 0 : -1;
        panels[k].hidden = k !== m;
      }
      if (focus) tabs[m].focus();
    };
    for (const k of order) {
      tabs[k] = el("button", { type: "button", role: "tab", class: "ribbon-tab", id: `pdfe-sig-t-${k}-${n}`, "aria-controls": panels[k].id, text: labels[k], onclick: () => show(k) });
      tablist.append(tabs[k]);
    }
    tablist.addEventListener("keydown", (e) => {
      const i = order.indexOf(mode);
      let j = null;
      if (e.key === "ArrowRight") j = (i + 1) % order.length;
      else if (e.key === "ArrowLeft") j = (i - 1 + order.length) % order.length;
      else if (e.key === "Home") j = 0; else if (e.key === "End") j = order.length - 1;
      if (j != null) { e.preventDefault(); show(order[j], true); }
    });
    const body = el("div", { class: "pdfe-sig" }, tablist, order.map((k) => panels[k]),
      el("div", { class: "pdfe-row" },
        el("label", { for: ink.id, text: "Ink color" }), ink,
        el("span", { class: "pdfe-spacer" }),
        remember, el("label", { for: remember.id, text: "Remember this signature on this computer" })));
    show(mode);
    const produce = (close) => {
      let url = null;
      if (mode === "saved") url = saved;
      else if (mode === "draw") url = trimCanvas(pad);
      else if (nameIn.value.trim()) url = trimCanvas(typed());
      if (!url) { App.toast(mode === "draw" ? "Draw your signature first, or use the Type tab." : "Type your name first.", "error"); return; }
      if (mode !== "saved" && remember.checked) lsSet(SIG_KEY, url);
      close(url);
    };
    return App.modal({
      title: "Add signature", wide: true, body, cancelValue: null,
      buttons: [{ label: "Cancel", value: null }, { label: "Insert signature", primary: true, onclick: produce }],
      onSubmit: produce,
    });
  }

  // ---------- the editor ----------
  class PdfEditor {
    constructor(container, doc) {
      this.c = container;
      this.doc = doc;
      this.uid = ++uidSeq;
      this.bytes = null;
      this.pdf = null;
      this.pages = [];
      this.objs = [];
      this.nextObj = 1;
      this.zoom = 1;
      this.zoomMode = "width";
      this.viewRot = 0;
      this.tool = "select";
      this.selId = null;
      this.sel = new Set();
      this.anchor = 0;
      this.cur = 0;
      this.thumbFocus = 0;
      this.undoStack = [];
      this.redoStack = [];
      this.dirty = false;
      this.password = null;
      this.savePassword = null;
      this.props = { text: "#000000", highlight: "#ffeb3b", stroke: "#d32f2f", width: 2, font: "Helvetica", fontSize: 12 };
      this.find = { q: "", matches: [], idx: -1, scrollTo: -1 };
      this.textCache = [];
      this.textIdx = [];
      this.visible = new Set();
      this.busy = false;
      this.build();
    }

    get cs() { return this.zoom * PX; }
    rot(P) { return ((P.page.rotate || 0) + this.viewRot) % 360; }
    id(s) { return `pdfe${this.uid}-${s}`; }

    // ----- UI construction -----
    btn(ico, label, title, fn, o = {}) {
      const b = el("button", {
        type: "button", class: "tool-btn" + (o.big ? " big" : ""), title,
        "aria-label": o.iconOnly ? label : null, "aria-pressed": o.pressed != null ? String(o.pressed) : null,
      }, el("span", { class: "ico", "aria-hidden": "true", text: ico }), o.iconOnly ? null : el("span", { class: "lbl", text: label }));
      b.addEventListener("click", (e) => fn(e));
      return b;
    }
    group(label, ...items) {
      return el("div", { class: "tool-group", role: "group", "aria-label": label },
        el("div", { class: "tool-group-items" }, ...items),
        el("div", { class: "tool-group-label", "aria-hidden": "true", text: label }));
    }
    field(label, input) { return el("label", { class: "pdfe-field" }, el("span", { text: label }), input); }
    act(fn) { return () => this.run(fn); }

    build() {
      const root = this.root = el("div", { class: "editor-root pdfe", "data-tool": "select" });
      // ribbon
      const tabsDef = [["home", "Home"], ["edit", "Edit"], ["organize", "Organize"], ["protect", "Protect"], ["forms", "Forms"], ["info", "Info"]];
      this.tabBtns = {};
      this.ribbonPanels = {};
      const tablist = el("div", { class: "ribbon-tabs", role: "tablist", "aria-label": "PDF tools" });
      for (const [k, label] of tabsDef) {
        const b = el("button", { type: "button", role: "tab", class: "ribbon-tab", id: this.id("tab-" + k), "aria-controls": this.id("panel-" + k), "aria-selected": "false", tabindex: "-1", text: label });
        b.addEventListener("click", () => this.showTab(k));
        tablist.append(b);
        this.tabBtns[k] = b;
        this.ribbonPanels[k] = el("div", { class: "ribbon-panel", role: "tabpanel", id: this.id("panel-" + k), "aria-labelledby": b.id, hidden: true });
      }
      tablist.addEventListener("keydown", (e) => {
        const keys = tabsDef.map((t) => t[0]);
        const i = keys.indexOf(this.ribbonTab);
        let j = null;
        if (e.key === "ArrowRight") j = (i + 1) % keys.length;
        else if (e.key === "ArrowLeft") j = (i - 1 + keys.length) % keys.length;
        else if (e.key === "Home") j = 0;
        else if (e.key === "End") j = keys.length - 1;
        if (j != null) { e.preventDefault(); this.showTab(keys[j]); this.tabBtns[keys[j]].focus(); }
      });
      this.buildHome();
      this.buildEdit();
      this.buildOrganize();
      this.buildProtect();
      this.buildForms();
      this.buildInfo();
      const ribbon = el("div", { class: "ribbon" }, tablist, ...Object.values(this.ribbonPanels));

      // thumbnails
      this.thumbList = el("ol", { class: "pdfe-thumbs", "aria-label": "Pages" });
      this.thumbsPanel = el("nav", { class: "panel-side pdfe-thumbs-panel", "aria-label": "Page thumbnails" },
        el("div", { class: "pdfe-thumbs-head" },
          el("h2", { class: "pdfe-h", text: "Pages" }),
          this.btn("▲", "Move selected pages up", "Move selected pages up (Alt+Up)", this.act(() => this.moveSel(-1)), { iconOnly: true }),
          this.btn("▼", "Move selected pages down", "Move selected pages down (Alt+Down)", this.act(() => this.moveSel(1)), { iconOnly: true })),
        this.thumbList);
      // pages
      this.pagesWrap = el("div", { class: "pdfe-pages" });
      this.scroller = el("div", { class: "editor-body pdfe-scroll", tabindex: "0", role: "region", "aria-label": "Document pages. Page Up and Page Down change page." }, this.pagesWrap);
      // side panel (forms / info)
      this.sideTitle = el("h2", { class: "pdfe-h", id: this.id("side-title"), tabindex: "-1" });
      this.sideBody = el("div", { class: "pdfe-side-body" });
      this.side = el("aside", { class: "panel-side right pdfe-side", "aria-labelledby": this.sideTitle.id, hidden: true },
        el("div", { class: "pdfe-side-head" }, this.sideTitle,
          el("button", { type: "button", class: "icon-btn", "aria-label": "Close panel", title: "Close panel (Esc)", text: "✕", onclick: () => this.closeSide() })),
        this.sideBody);
      this.side.addEventListener("keydown", (e) => { if (e.key === "Escape" && !e.defaultPrevented) { e.stopPropagation(); this.closeSide(); } });
      const main = el("div", { class: "editor-main" }, this.thumbsPanel, this.scroller, this.side);

      // status bar
      this.stPage = el("span", { text: "Page – of –" });
      this.stZoom = el("span", { text: "100%" });
      this.stProt = el("span", { hidden: true, text: "🔒 Saved with password" });
      this.stHint = el("span", { class: "pdfe-hint", "aria-live": "polite" });
      const status = el("div", { class: "statusbar" }, this.stPage, this.stZoom, this.stProt, this.stHint);

      root.append(ribbon, main, status);
      this.c.append(root);
      this.showTab("home");

      root.addEventListener("keydown", (e) => this.onKey(e));
      this.scroller.addEventListener("scroll", () => this.onScroll(), { passive: true });
      this.scroller.addEventListener("wheel", (e) => {
        if (!e.ctrlKey) return;
        e.preventDefault();
        this.zoomStep(e.deltaY < 0 ? 1 : -1);
      }, { passive: false });
      this.scroller.addEventListener("pointerdown", (e) => this.onScrollerDown(e));
      this.ro = new ResizeObserver(() => {
        if (!this.zoomMode || !this.pages.length) return;
        cancelAnimationFrame(this.roFrame);
        this.roFrame = requestAnimationFrame(() => this.setZoom(null, this.zoomMode));
      });
      this.ro.observe(this.scroller);
    }

    buildHome() {
      const p = this.ribbonPanels.home;
      this.undoBtn = this.btn("↶", "Undo", "Undo (Ctrl+Z)", this.act(() => this.undo()));
      this.redoBtn = this.btn("↷", "Redo", "Redo (Ctrl+Y)", this.act(() => this.redo()));
      this.thumbToggle = this.btn("▤", "Thumbnails", "Show or hide page thumbnails", () => {
        const show = this.thumbsPanel.hidden;
        this.thumbsPanel.hidden = !show;
        this.thumbToggle.setAttribute("aria-pressed", String(show));
      }, { pressed: true });
      this.zoomSel = el("select", { class: "tool-select", "aria-label": "Zoom", title: "Zoom (Ctrl+Plus / Ctrl+Minus, Ctrl+mouse wheel)" },
        el("option", { value: "width", text: "Fit width" }), el("option", { value: "page", text: "Fit page" }),
        ZOOMS.map((z) => el("option", { value: String(z), text: Math.round(z * 100) + "%" })),
        this.zoomCustom = el("option", { value: "custom", text: "", hidden: true }));
      this.zoomSel.addEventListener("change", () => {
        const v = this.zoomSel.value;
        if (v === "width" || v === "page") this.setZoom(null, v);
        else if (v !== "custom") this.setZoom(parseFloat(v));
      });
      this.pageInput = el("input", { type: "number", class: "tool-input pdfe-pagenum", min: "1", value: "1", "aria-label": "Page number", title: "Type a page number and press Enter" });
      this.pageInput.addEventListener("change", () => this.goto(parseInt(this.pageInput.value, 10) - 1));
      this.pageInput.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); this.goto(parseInt(this.pageInput.value, 10) - 1); } });
      this.pageOf = el("span", { class: "pdfe-of", text: "of 1" });
      this.prevBtn = this.btn("◀", "Previous page", "Previous page (Page Up)", () => this.goto(this.cur - 1), { iconOnly: true });
      this.nextBtn = this.btn("▶", "Next page", "Next page (Page Down)", () => this.goto(this.cur + 1), { iconOnly: true });
      this.findInput = el("input", { type: "search", class: "tool-input pdfe-find", "aria-label": "Find text", placeholder: "Find text", title: "Find text (Ctrl+F). Enter: next match, Shift+Enter: previous" });
      this.findInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") { e.preventDefault(); this.run(() => this.runFind(e.shiftKey ? -1 : 1)); }
        else if (e.key === "Escape" && this.findInput.value) { e.preventDefault(); e.stopPropagation(); this.findInput.value = ""; this.clearFind(); }
      });
      this.findInput.addEventListener("input", () => { if (!this.findInput.value) this.clearFind(); });
      this.findCount = el("span", { class: "pdfe-findcount", "aria-live": "polite" });
      p.append(
        this.group("File",
          this.btn("💾", "Save", "Save (Ctrl+S)", () => this.save().catch((err) => App.toast("Save failed: " + err.message, "error")), { big: true }),
          this.btn("📄", "Save as", "Save as (Ctrl+Shift+S)", () => this.saveAs().catch((err) => App.toast("Save failed: " + err.message, "error"))),
          this.btn("🖨", "Print", "Print (Ctrl+P)", this.act(() => this.print()))),
        this.group("History", this.undoBtn, this.redoBtn),
        this.group("View", this.thumbToggle,
          this.btn("−", "Zoom out", "Zoom out (Ctrl+Minus)", () => this.zoomStep(-1), { iconOnly: true }),
          this.zoomSel,
          this.btn("+", "Zoom in", "Zoom in (Ctrl+Plus)", () => this.zoomStep(1), { iconOnly: true }),
          this.btn("↔", "Fit width", "Fit page width", () => this.setZoom(null, "width")),
          this.btn("⤢", "Fit page", "Fit whole page", () => this.setZoom(null, "page")),
          this.btn("⟳", "Rotate view", "Rotate the view 90° (display only, the file is not changed; editing tools are disabled while rotated)", this.act(() => this.rotateView()))),
        this.group("Page", this.prevBtn, this.pageInput, this.pageOf, this.nextBtn),
        this.group("Find", this.findInput,
          this.btn("⌃", "Previous match", "Previous match (Shift+Enter)", this.act(() => this.runFind(-1)), { iconOnly: true }),
          this.btn("⌄", "Next match", "Next match (Enter)", this.act(() => this.runFind(1)), { iconOnly: true }),
          this.findCount));
    }

    buildEdit() {
      const p = this.ribbonPanels.edit;
      this.toolBtns = {};
      this.editOnly = [];
      const mk = (t) => {
        const b = this.btn(t.ico, t.label, `${t.label} (${t.key})${t.extra || ""}`, () => this.setTool(t.id), { pressed: t.id === "select" });
        this.toolBtns[t.id] = b;
        if (t.edit) this.editOnly.push(b);
        return b;
      };
      const ins = (ico, label, title, fn) => { const b = this.btn(ico, label, title, this.act(fn)); this.editOnly.push(b); return b; };
      this.stampSel = el("select", { class: "tool-select", "aria-label": "Stamp type", title: "Stamp type" },
        Object.entries(STAMPS).map(([k, s]) => el("option", { value: k, text: s.label })));
      this.editOnly.push(this.stampSel);
      this.colorIn = el("input", { type: "color", class: "tool-input", value: this.props.stroke, title: "Color of the selected item or the next one you add" });
      this.colorIn.addEventListener("change", () => this.onProp("color", this.colorIn.value));
      this.widthSel = el("select", { class: "tool-select", title: "Line width for drawing and shapes" },
        [1, 2, 3, 5, 8, 12].map((w) => el("option", { value: String(w), text: w + " pt", selected: w === 2 })));
      this.widthSel.addEventListener("change", () => this.onProp("width", parseFloat(this.widthSel.value)));
      this.fontSel = el("select", { class: "tool-select", title: "Font for text boxes" },
        Object.keys(FONT_CSS).map((f) => el("option", { value: f, text: f === "Times-Roman" ? "Times" : f })));
      this.fontSel.addEventListener("change", () => this.onProp("font", this.fontSel.value));
      this.sizeIn = el("input", { type: "number", class: "tool-input pdfe-size", min: "6", max: "96", step: "1", value: "12", title: "Font size in points for text boxes and notes" });
      this.sizeIn.addEventListener("change", () => this.onProp("fontSize", clamp(parseFloat(this.sizeIn.value) || 12, 6, 96)));
      this.delBtn = this.btn("🗑", "Delete item", "Delete the selected item (Delete)", () => { const o = this.selObj(); if (o) this.deleteObj(o); else this.hint("Select an item first."); });
      p.append(
        this.group("Tools", ...TOOLS.filter((t) => !t.shape).map(mk)),
        this.group("Shapes", ...TOOLS.filter((t) => t.shape).map(mk)),
        this.group("Insert",
          ins("🖼", "Image", "Insert a PNG or JPG image (I)", () => this.insertImage()),
          ins("✍", "Sign", "Add your signature: draw or type it (S)", () => this.insertSignature()),
          this.stampSel,
          ins("🔖", "Stamp", "Place the chosen stamp (P)", () => this.insertStamp())),
        this.group("Properties",
          this.field("Color", this.colorIn), this.field("Width", this.widthSel),
          this.field("Font", this.fontSel), this.field("Size", this.sizeIn)),
        this.group("Item", this.delBtn,
          this.btn("✔", "Apply to PDF", "Merge all added items into the PDF pages now (also done automatically on save)", this.act(() => this.flatten(true)))));
    }

    buildOrganize() {
      const p = this.ribbonPanels.organize;
      p.append(
        this.group("Pages",
          this.btn("⟲", "Rotate left", "Rotate selected pages 90° counter-clockwise", this.act(() => this.rotatePages(270))),
          this.btn("⟳", "Rotate right", "Rotate selected pages 90° clockwise", this.act(() => this.rotatePages(90))),
          this.btn("🗑", "Delete pages", "Delete selected pages (Delete in the thumbnail list)", this.act(() => this.deletePages())),
          this.btn("⬆", "Move up", "Move selected pages up (Alt+Up)", this.act(() => this.moveSel(-1))),
          this.btn("⬇", "Move down", "Move selected pages down (Alt+Down)", this.act(() => this.moveSel(1)))),
        this.group("Selection",
          this.btn("☑", "Select all", "Select all pages (Ctrl+A in the thumbnail list)", () => this.selectAllPages()),
          this.btn("☐", "Clear", "Clear page selection", () => { this.sel.clear(); this.updateThumbSel(); this.hint("Page selection cleared."); })),
        this.group("Insert",
          this.btn("📑", "Insert PDF", "Insert pages from other PDF files after the current page", this.act(() => this.insertPdfFiles())),
          this.btn("🖼", "Insert images", "Insert PNG or JPG images as new pages after the current page", this.act(() => this.insertImagePages()))),
        this.group("Export",
          this.btn("📤", "Extract", "Copy the selected pages into a new PDF", this.act(() => this.extractPages())),
          this.btn("✂", "Split", "Split into one PDF per page (saved as a ZIP file)", this.act(() => this.splitPages()))));
    }

    buildProtect() {
      const p = this.ribbonPanels.protect;
      this.wmText = el("input", { type: "text", class: "tool-input", value: "CONFIDENTIAL", title: "Watermark text" });
      this.wmOpacity = el("input", { type: "number", class: "tool-input pdfe-size", min: "0.05", max: "1", step: "0.05", value: "0.3", title: "Watermark opacity from 0.05 (faint) to 1 (solid)" });
      p.append(
        this.group("Password",
          this.btn("🔒", "Set password", "Require a password to open this PDF (applied when you save)", this.act(() => this.setPassword())),
          this.btn("🔓", "Remove password", "Save this PDF without a password", this.act(() => this.removePassword()))),
        this.group("Watermark",
          this.field("Text", this.wmText), this.field("Opacity", this.wmOpacity),
          this.btn("💧", "Apply watermark", "Add the watermark text to every page", this.act(() => this.watermark()))),
        this.group("Optimize",
          this.btn("🗜", "Compress", "Optimize the file to make it smaller", this.act(() => this.compress()))));
    }

    buildForms() {
      this.ribbonPanels.forms.append(this.group("Form",
        this.btn("📋", "Fill form fields", "Show the fillable form fields of this PDF in a side panel", this.act(() => this.showForms()), { big: true })));
    }

    buildInfo() {
      this.ribbonPanels.info.append(this.group("Document",
        this.btn("ℹ", "Document properties", "Show page count, title, author and page sizes", this.act(() => this.showInfo()), { big: true })));
    }

    showTab(k) {
      this.ribbonTab = k;
      for (const [key, b] of Object.entries(this.tabBtns)) {
        const on = key === k;
        b.setAttribute("aria-selected", String(on));
        b.tabIndex = on ? 0 : -1;
        this.ribbonPanels[key].hidden = !on;
      }
    }

    hint(msg) { this.stHint.textContent = msg || ""; }

    async run(fn) {
      if (this.busy) { App.toast("Please wait for the current operation to finish.", "info"); return; }
      this.busy = true;
      this.root.setAttribute("aria-busy", "true");
      try { await fn(); } catch (err) {
        console.error(err);
        App.toast(err && err.message ? err.message : String(err), "error");
      } finally {
        this.busy = false;
        this.root.removeAttribute("aria-busy");
      }
    }

    // ----- loading & rendering -----
    async open() {
      if (!pdfjsLib) throw new Error("The PDF viewer library (vendor/pdfjs/pdf.min.js) did not load.");
      const bytes = this.doc.bytes && this.doc.bytes.length ? this.doc.bytes : blankPdf();
      await this.load(bytes);
      this.hint(TOOLS[0].hint);
    }

    async load(bytes, opts = {}) {
      const keep = opts.keepScroll ? { top: this.scroller.scrollTop, left: this.scroller.scrollLeft } : null;
      let usedPw = false, cancelled = false;
      const task = pdfjsLib.getDocument({ data: bytes.slice(), password: this.password || undefined, isEvalSupported: false });
      task.onPassword = (update, reason) => {
        const wrong = reason === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD;
        passwordDialog("Password required", wrong ? "That password is not correct. Try again:" : `"${this.doc.name}" is protected. Enter its password to open it:`, false)
          .then((pw) => {
            if (pw == null) { cancelled = true; task.destroy(); } else { usedPw = true; this.password = pw; update(pw); }
          });
      };
      let pdf;
      try { pdf = await task.promise; } catch (err) {
        if (cancelled) throw new Error("The document is password protected and was not unlocked.");
        throw new Error("Could not read the PDF: " + (err && err.message ? err.message : err));
      }
      if (usedPw) {
        // keep an unencrypted working copy; the password is re-applied on save
        pdf.destroy();
        const dec = await this.post("/api/pdf/decrypt", { password: this.password }, [{ blob: new Blob([bytes], { type: "application/pdf" }), name: "document.pdf" }]);
        this.savePassword = this.password;
        this.password = null;
        App.toast("Unlocked. The document will be saved with the same password; use Protect > Remove password to drop it.", "info");
        return this.load(dec, opts);
      }
      if (this.pdf) { try { this.pdf.destroy(); } catch (_) { /* ignore */ } }
      this.pdf = pdf;
      this.bytes = bytes;
      this.textCache = [];
      this.textIdx = [];
      const pages = [];
      for (let i = 1; i <= pdf.numPages; i++) pages.push({ index: i - 1, page: await pdf.getPage(i) });
      this.pages = pages;
      const n = pages.length;
      this.sel = new Set([...this.sel].filter((i) => i < n));
      this.cur = clamp(this.cur, 0, n - 1);
      this.thumbFocus = clamp(this.thumbFocus, 0, n - 1);
      this.objs = this.objs.filter((o) => o.page < n);
      this.buildPages();
      this.buildThumbs();
      this.applySizes();
      if (this.zoomMode) this.setZoom(null, this.zoomMode, true);
      if (keep) { this.scroller.scrollTop = keep.top; this.scroller.scrollLeft = keep.left; } else this.scroller.scrollTop = 0;
      this.onScroll(true);
      this.updateProtUi();
      this.updateUndoUi();
      if (this.find.q) this.runFind(0, true).catch(() => {});
    }

    buildPages() {
      if (this.io) this.io.disconnect();
      this.pagesWrap.textContent = "";
      this.visible.clear();
      this.io = new IntersectionObserver((entries) => {
        for (const en of entries) {
          const i = +en.target.dataset.index;
          if (en.isIntersecting) { this.visible.add(i); const P = this.pages[i]; if (P && P.el === en.target) this.renderPage(P); } else this.visible.delete(i);
        }
      }, { root: this.scroller, rootMargin: "800px 0px" });
      for (const P of this.pages) {
        P.el = el("div", { class: "pdfe-page", "data-index": String(P.index), role: "group", "aria-label": `Page ${P.index + 1}` });
        P.canvas = el("canvas", { class: "pdfe-canvas", "aria-hidden": "true" });
        P.textEl = el("div", { class: "textLayer" });
        P.layer = el("div", { class: "pdfe-layer" });
        P.layer.addEventListener("pointerdown", (e) => this.onLayerDown(e, P));
        P.el.append(P.canvas, P.textEl, P.layer);
        this.pagesWrap.append(P.el);
        this.io.observe(P.el);
      }
    }

    applySizes() {
      const cs = this.cs;
      for (const P of this.pages) {
        P.vp = P.page.getViewport({ scale: cs, rotation: this.rot(P) });
        P.el.style.width = Math.floor(P.vp.width) + "px";
        P.el.style.height = Math.floor(P.vp.height) + "px";
        this.renderLayer(P.index);
      }
      for (const i of this.visible) if (this.pages[i]) this.renderPage(this.pages[i]);
      this.updateZoomUi();
    }

    async renderPage(P) {
      const key = this.cs.toFixed(4) + ":" + this.rot(P);
      if (P.renderedKey === key || P.rendering === key) return;
      P.rendering = key;
      if (P.task) { try { P.task.cancel(); } catch (_) { /* ignore */ } }
      const vp = P.vp, dpr = window.devicePixelRatio || 1;
      const canvas = el("canvas", { class: "pdfe-canvas", "aria-hidden": "true" });
      canvas.width = Math.max(1, Math.floor(vp.width * dpr));
      canvas.height = Math.max(1, Math.floor(vp.height * dpr));
      const task = P.page.render({ canvasContext: canvas.getContext("2d"), viewport: vp, transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : null });
      P.task = task;
      try { await task.promise; } catch (_) { if (P.rendering === key) P.rendering = null; return; }
      if (P.rendering !== key) return;
      P.canvas.replaceWith(canvas);
      P.canvas = canvas;
      P.renderedKey = key;
      P.rendering = null;
      P.task = null;
      try {
        const tc = await this.getText(P.index);
        if (P.renderedKey !== key) return;
        const tl = el("div", { class: "textLayer" });
        tl.style.setProperty("--scale-factor", String(vp.scale));
        const textDivs = [];
        await pdfjsLib.renderTextLayer({ textContentSource: tc, container: tl, viewport: vp, textDivs }).promise;
        if (P.renderedKey !== key) return;
        P.textEl.replaceWith(tl);
        P.textEl = tl;
        P.textDivs = textDivs;
        this.applyHits(P.index);
      } catch (err) { console.warn("text layer", err); }
    }

    getText(i) {
      if (!this.textCache[i]) this.textCache[i] = this.pages[i].page.getTextContent();
      return this.textCache[i];
    }

    // ----- thumbnails -----
    buildThumbs() {
      if (this.thumbIo) this.thumbIo.disconnect();
      this.thumbList.textContent = "";
      this.thumbIo = new IntersectionObserver((entries) => {
        for (const en of entries) if (en.isIntersecting) { const P = this.pages[+en.target.dataset.index]; if (P && P.thumb === en.target) this.renderThumb(P); }
      }, { root: this.thumbsPanel, rootMargin: "400px 0px" });
      for (const P of this.pages) {
        const i = P.index;
        const vp1 = P.page.getViewport({ scale: 1, rotation: this.rot(P) });
        const holder = el("span", { class: "pdfe-thumb-img" });
        holder.style.aspectRatio = `${vp1.width} / ${vp1.height}`;
        const b = el("button", {
          type: "button", class: "pdfe-thumb", "data-index": String(i), "aria-label": `Page ${i + 1}`, "aria-pressed": "false", draggable: "true", tabindex: "-1",
          title: `Page ${i + 1}. Click to go there. Ctrl+click or Ctrl+Space adds to the selection, Shift+click selects a range. Drag or Alt+Up/Down to move.`,
        }, holder, el("span", { class: "pdfe-thumb-num", "aria-hidden": "true", text: String(i + 1) }));
        b.addEventListener("click", (e) => this.onThumbClick(e, i));
        b.addEventListener("keydown", (e) => this.onThumbKey(e, i));
        b.addEventListener("dragstart", (e) => {
          this.dragPages = this.sel.has(i) ? [...this.sel].sort((x, y) => x - y) : [i];
          e.dataTransfer.effectAllowed = "move";
          e.dataTransfer.setData("application/x-docstudio-pages", this.dragPages.join(","));
        });
        b.addEventListener("dragover", (e) => {
          if (!this.dragPages) return;
          e.preventDefault();
          e.stopPropagation();
          const r = b.getBoundingClientRect();
          const after = e.clientY > r.top + r.height / 2;
          b.classList.toggle("drop-after", after);
          b.classList.toggle("drop-before", !after);
        });
        b.addEventListener("dragleave", () => b.classList.remove("drop-before", "drop-after"));
        b.addEventListener("drop", (e) => {
          if (!this.dragPages) return;
          e.preventDefault();
          e.stopPropagation();
          const after = b.classList.contains("drop-after");
          b.classList.remove("drop-before", "drop-after");
          const moving = this.dragPages;
          this.dragPages = null;
          this.run(() => this.movePages(moving, after ? i + 1 : i));
        });
        b.addEventListener("dragend", () => { this.dragPages = null; this.thumbList.querySelectorAll(".drop-before,.drop-after").forEach((x) => x.classList.remove("drop-before", "drop-after")); });
        P.thumb = b;
        this.thumbList.append(el("li", {}, b));
        this.thumbIo.observe(b);
      }
      this.updateThumbSel();
    }

    async renderThumb(P) {
      if (P.thumbDone) return;
      P.thumbDone = true;
      const dpr = window.devicePixelRatio || 1;
      const vp1 = P.page.getViewport({ scale: 1, rotation: this.rot(P) });
      const vp = P.page.getViewport({ scale: (110 / vp1.width) * dpr, rotation: this.rot(P) });
      const c = el("canvas", { "aria-hidden": "true" });
      c.width = Math.max(1, Math.floor(vp.width));
      c.height = Math.max(1, Math.floor(vp.height));
      try {
        await P.page.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise;
        P.thumb.firstChild.replaceChildren(c);
      } catch (_) { P.thumbDone = false; }
    }

    updateThumbSel() {
      for (const P of this.pages) {
        if (!P.thumb) continue;
        P.thumb.setAttribute("aria-pressed", String(this.sel.has(P.index)));
        P.thumb.tabIndex = P.index === this.thumbFocus ? 0 : -1;
        if (P.index === this.cur) P.thumb.setAttribute("aria-current", "page"); else P.thumb.removeAttribute("aria-current");
      }
    }

    onThumbClick(e, i) {
      if (e.shiftKey) {
        this.sel = new Set();
        for (let k = Math.min(this.anchor, i); k <= Math.max(this.anchor, i); k++) this.sel.add(k);
      } else if (e.ctrlKey || e.metaKey) {
        if (this.sel.has(i)) this.sel.delete(i); else this.sel.add(i);
        this.anchor = i;
      } else {
        this.sel = new Set([i]);
        this.anchor = i;
      }
      this.thumbFocus = i;
      this.goto(i);
      this.updateThumbSel();
      this.hint(this.sel.size ? `${this.sel.size} page${this.sel.size > 1 ? "s" : ""} selected.` : "No pages selected.");
    }

    onThumbKey(e, i) {
      const n = this.pages.length;
      const focusThumb = (j) => { this.thumbFocus = j; this.updateThumbSel(); this.pages[j].thumb.focus(); this.pages[j].thumb.scrollIntoView({ block: "nearest" }); };
      let handled = true;
      const ctrl = e.ctrlKey || e.metaKey;
      if (e.altKey && (e.key === "ArrowUp" || e.key === "ArrowDown")) {
        this.run(async () => { if (!this.sel.has(i)) this.sel = new Set([i]); await this.moveSel(e.key === "ArrowUp" ? -1 : 1); const f = Math.min(...this.sel); if (this.pages[f]) focusThumb(f); });
      } else if (["ArrowDown", "ArrowRight", "ArrowUp", "ArrowLeft", "Home", "End"].includes(e.key) && !e.altKey) {
        const j = e.key === "Home" ? 0 : e.key === "End" ? n - 1 : clamp(i + (e.key === "ArrowDown" || e.key === "ArrowRight" ? 1 : -1), 0, n - 1);
        if (e.shiftKey) {
          this.sel = new Set();
          for (let k = Math.min(this.anchor, j); k <= Math.max(this.anchor, j); k++) this.sel.add(k);
        }
        focusThumb(j);
        if (!e.shiftKey && !ctrl) { this.goto(j); }
      } else if (e.key === " " && ctrl) {
        if (this.sel.has(i)) this.sel.delete(i); else this.sel.add(i);
        this.anchor = i;
        this.updateThumbSel();
        this.hint(`Page ${i + 1} ${this.sel.has(i) ? "added to" : "removed from"} the selection. ${this.sel.size} selected.`);
      } else if (e.key === "Delete") {
        if (!this.sel.size) this.sel = new Set([i]);
        this.run(() => this.deletePages());
      } else if (ctrl && e.key.toLowerCase() === "a") {
        this.selectAllPages();
      } else handled = false;
      if (handled) { e.preventDefault(); e.stopPropagation(); }
    }

    selectAllPages() {
      this.sel = new Set(this.pages.map((P) => P.index));
      this.updateThumbSel();
      this.hint(`All ${this.pages.length} pages selected.`);
    }

    // ----- navigation & zoom -----
    onScroll(force) {
      if (this.scrollFrame && !force) return;
      const update = () => {
        this.scrollFrame = 0;
        const mid = this.scroller.scrollTop + this.scroller.clientHeight * 0.35;
        let c = 0;
        for (const P of this.pages) { if (P.el.offsetTop <= mid) c = P.index; else break; }
        if (c !== this.cur || force) { this.cur = c; this.updatePageUi(); }
      };
      if (force === true) update(); else this.scrollFrame = requestAnimationFrame(update);
    }

    updatePageUi() {
      const n = this.pages.length;
      this.pageInput.max = String(n);
      this.pageInput.value = String(this.cur + 1);
      this.pageOf.textContent = "of " + n;
      this.stPage.textContent = `Page ${this.cur + 1} of ${n}`;
      this.prevBtn.disabled = this.cur <= 0;
      this.nextBtn.disabled = this.cur >= n - 1;
      this.updateThumbSel();
    }

    goto(i) {
      if (!this.pages.length || isNaN(i)) return;
      i = clamp(i, 0, this.pages.length - 1);
      this.cur = i;
      this.scroller.scrollTop = this.pages[i].el.offsetTop - 12;
      this.updatePageUi();
      if (this.pages[i].thumb) this.pages[i].thumb.scrollIntoView({ block: "nearest" });
    }

    fitZoom(mode) {
      const W = Math.max(100, this.scroller.clientWidth - 48), H = Math.max(100, this.scroller.clientHeight - 32);
      if (mode === "width") {
        let maxW = 1;
        for (const P of this.pages) maxW = Math.max(maxW, P.page.getViewport({ scale: 1, rotation: this.rot(P) }).width);
        return W / (maxW * PX);
      }
      const P = this.pages[this.cur] || this.pages[0];
      const vp = P.page.getViewport({ scale: 1, rotation: this.rot(P) });
      return Math.min(W / (vp.width * PX), H / (vp.height * PX));
    }

    setZoom(z, mode, noAnchor) {
      if (!this.pages.length) return;
      const P = this.pages[this.cur];
      const frac = P && P.el.offsetHeight ? (this.scroller.scrollTop - P.el.offsetTop) / P.el.offsetHeight : 0;
      this.zoomMode = mode || null;
      const nz = clamp(mode ? this.fitZoom(mode) : z, 0.1, 4);
      if (Math.abs(nz - this.zoom) < 0.0005 && this.pages[0].vp) { this.updateZoomUi(); return; }
      this.zoom = nz;
      this.applySizes();
      if (P && !noAnchor) this.scroller.scrollTop = P.el.offsetTop + frac * P.el.offsetHeight;
    }

    zoomStep(dir) {
      const z = this.zoom;
      const next = dir > 0 ? ZOOMS.find((v) => v > z + 0.001) : [...ZOOMS].reverse().find((v) => v < z - 0.001);
      if (next) this.setZoom(next);
    }

    updateZoomUi() {
      const pct = Math.round(this.zoom * 100) + "%";
      this.stZoom.textContent = "Zoom " + pct;
      if (this.zoomMode) { this.zoomSel.value = this.zoomMode; this.zoomCustom.hidden = true; return; }
      const exact = ZOOMS.find((v) => Math.abs(v - this.zoom) < 0.001);
      if (exact) { this.zoomSel.value = String(exact); this.zoomCustom.hidden = true; } else {
        this.zoomCustom.textContent = pct;
        this.zoomCustom.hidden = false;
        this.zoomSel.value = "custom";
      }
    }

    async rotateView() {
      if (this.objs.length) {
        await this.flatten(true);
        App.toast("Your added items were applied to the pages before rotating the view.", "info");
      }
      this.viewRot = (this.viewRot + 90) % 360;
      if (this.viewRot && TOOLS.find((t) => t.id === this.tool).edit) this.setTool("select");
      for (const b of this.editOnly) b.disabled = this.viewRot !== 0;
      for (const P of this.pages) P.thumbDone = false;
      this.buildThumbs();
      if (this.zoomMode) { this.zoom = this.fitZoom(this.zoomMode); }
      this.applySizes();
      this.goto(this.cur);
      this.hint(this.viewRot ? `View rotated ${this.viewRot}°. Editing tools are disabled until the view is back at 0°.` : "View rotation reset. Editing tools are available again.");
    }

    // ----- tools -----
    setTool(id) {
      const t = TOOLS.find((x) => x.id === id);
      if (!t) return;
      if (t.edit && this.viewRot) { App.toast("Editing tools are not available while the view is rotated. Use Rotate view until it is back at 0°.", "error"); return; }
      this.tool = id;
      for (const [k, b] of Object.entries(this.toolBtns)) b.setAttribute("aria-pressed", String(k === id));
      this.root.dataset.tool = id;
      this.root.classList.toggle("pdfe-drawing", id !== "select" && id !== "hand");
      this.root.classList.toggle("pdfe-hand", id === "hand");
      if (id !== "select") this.selectObj(null);
      this.syncProps();
      this.hint(t.hint);
    }

    colorKey(type) { return type === "text" ? "text" : type === "highlight" ? "highlight" : "stroke"; }

    syncProps() {
      const o = this.selObj();
      const type = o ? o.type : this.tool === "draw" ? "ink" : this.tool;
      this.colorIn.value = o && o.color ? o.color : this.props[this.colorKey(type)];
      this.widthSel.value = String(o && o.width ? o.width : this.props.width);
      this.fontSel.value = o && o.font ? o.font : this.props.font;
      this.sizeIn.value = String(o && o.fontSize ? o.fontSize : this.props.fontSize);
    }

    onProp(name, value) {
      const o = this.selObj();
      if (name === "color") this.props[this.colorKey(o ? o.type : this.tool === "draw" ? "ink" : this.tool)] = value;
      else this.props[name] = value;
      if (!o) return;
      const ok = {
        color: ["text", "highlight", "ink", "rect", "ellipse", "line", "arrow", "stamp"],
        width: ["ink", "rect", "ellipse", "line", "arrow"],
        font: ["text"],
        fontSize: ["text", "note"],
      }[name];
      if (!ok.includes(o.type)) return;
      this.pushUndo();
      o[name] = value;
      this.refreshObj(o, true);
      this.setDirty(true);
    }

    toUnits(e, P) {
      const r = P.layer.getBoundingClientRect();
      return [clamp((e.clientX - r.left) / this.cs, 0, P.vp.width / this.cs), clamp((e.clientY - r.top) / this.cs, 0, P.vp.height / this.cs)];
    }

    onScrollerDown(e) {
      if (this.tool === "hand" && e.button === 0) {
        e.preventDefault();
        const sx = e.clientX, sy = e.clientY, st = this.scroller.scrollTop, sl = this.scroller.scrollLeft;
        this.scroller.setPointerCapture(e.pointerId);
        this.root.classList.add("pdfe-grabbing");
        const move = (ev) => { this.scroller.scrollTop = st - (ev.clientY - sy); this.scroller.scrollLeft = sl - (ev.clientX - sx); };
        const up = () => { this.scroller.removeEventListener("pointermove", move); this.scroller.removeEventListener("pointerup", up); this.scroller.removeEventListener("pointercancel", up); this.root.classList.remove("pdfe-grabbing"); };
        this.scroller.addEventListener("pointermove", move);
        this.scroller.addEventListener("pointerup", up);
        this.scroller.addEventListener("pointercancel", up);
        return;
      }
      if (this.tool === "select" && this.selId && !e.target.closest(".pdfe-obj")) this.selectObj(null);
    }

    onLayerDown(e, P) {
      const tool = this.tool;
      if (e.button !== 0 || tool === "select" || tool === "hand") return;
      e.preventDefault();
      e.stopPropagation();
      const [x, y] = this.toUnits(e, P);
      const pw = P.vp.width / this.cs, ph = P.vp.height / this.cs;
      if (tool === "text") {
        const fs = this.props.fontSize;
        this.addObj({ page: P.index, type: "text", x, y: clamp(y - fs * 0.6, 0, ph - fs * 1.2), w: Math.max(40, Math.min(240, pw - x)), h: fs * 1.2, text: "", font: this.props.font, fontSize: fs, color: this.props.text }, { edit: true });
        this.setTool("select");
        this.selectObj(this.objs[this.objs.length - 1].id);
        this.hint("Type your text. Esc finishes editing; drag the ✥ handle or use arrow keys to move the box.");
        return;
      }
      if (tool === "note") {
        this.addObj({ page: P.index, type: "note", x: Math.min(x, pw - 160), y: Math.min(y, ph - 90), w: 160, h: 90, text: "", fontSize: 10 }, { edit: true });
        this.setTool("select");
        this.selectObj(this.objs[this.objs.length - 1].id);
        this.hint("Type the note. Esc finishes editing.");
        return;
      }
      const cs = this.cs;
      const color = tool === "highlight" ? this.props.highlight : this.props.stroke;
      const width = this.props.width;
      const svgNS = "http://www.w3.org/2000/svg";
      const svg = document.createElementNS(svgNS, "svg");
      svg.setAttribute("class", "pdfe-preview");
      svg.setAttribute("width", P.vp.width);
      svg.setAttribute("height", P.vp.height);
      const shape = document.createElementNS(svgNS, tool === "draw" ? "polyline" : tool === "ellipse" ? "ellipse" : tool === "line" || tool === "arrow" ? "line" : "rect");
      shape.setAttribute("fill", tool === "highlight" ? hexToRgba(color, 0.4) : tool === "whiteout" ? "#ffffff" : "none");
      shape.setAttribute("stroke", tool === "highlight" ? "none" : tool === "whiteout" ? "#9e9e9e" : color);
      shape.setAttribute("stroke-width", tool === "whiteout" ? 1 : width * cs);
      shape.setAttribute("stroke-linecap", "round");
      shape.setAttribute("stroke-linejoin", "round");
      if (tool === "whiteout") shape.setAttribute("stroke-dasharray", "4 3");
      svg.append(shape);
      P.layer.append(svg);
      const pts = [[x, y]];
      let end = [x, y];
      const draw = () => {
        const [x0, y0] = pts[0];
        const [x1, y1] = end;
        if (tool === "draw") shape.setAttribute("points", pts.map(([a, b]) => `${a * cs},${b * cs}`).join(" "));
        else if (tool === "line" || tool === "arrow") { shape.setAttribute("x1", x0 * cs); shape.setAttribute("y1", y0 * cs); shape.setAttribute("x2", x1 * cs); shape.setAttribute("y2", y1 * cs); }
        else if (tool === "ellipse") { shape.setAttribute("cx", (x0 + x1) / 2 * cs); shape.setAttribute("cy", (y0 + y1) / 2 * cs); shape.setAttribute("rx", Math.abs(x1 - x0) / 2 * cs); shape.setAttribute("ry", Math.abs(y1 - y0) / 2 * cs); }
        else { shape.setAttribute("x", Math.min(x0, x1) * cs); shape.setAttribute("y", Math.min(y0, y1) * cs); shape.setAttribute("width", Math.abs(x1 - x0) * cs); shape.setAttribute("height", Math.abs(y1 - y0) * cs); }
      };
      draw();
      P.layer.setPointerCapture(e.pointerId);
      const move = (ev) => { end = this.toUnits(ev, P); if (tool === "draw") pts.push(end); draw(); };
      const up = () => {
        P.layer.removeEventListener("pointermove", move);
        P.layer.removeEventListener("pointerup", up);
        P.layer.removeEventListener("pointercancel", up);
        svg.remove();
        this.finishShape(P, tool, pts, end, color, width);
      };
      P.layer.addEventListener("pointermove", move);
      P.layer.addEventListener("pointerup", up);
      P.layer.addEventListener("pointercancel", up);
    }

    finishShape(P, tool, pts, end, color, width) {
      const pw = P.vp.width / this.cs, ph = P.vp.height / this.cs;
      const [x0, y0] = pts[0], [x1, y1] = end;
      if (tool === "highlight" || tool === "whiteout" || tool === "rect" || tool === "ellipse") {
        const w = Math.abs(x1 - x0), h = Math.abs(y1 - y0);
        if (w < 3 || h < 3) { this.hint("Drag on the page to draw the area."); return; }
        const o = { page: P.index, type: tool, x: Math.min(x0, x1), y: Math.min(y0, y1), w, h };
        if (tool !== "whiteout") o.color = color;
        if (tool === "rect" || tool === "ellipse") o.width = width;
        this.addObj(o);
        return;
      }
      const all = tool === "draw" ? pts : [pts[0], end];
      if (tool !== "draw" && Math.hypot(x1 - x0, y1 - y0) < 4) { this.hint("Drag on the page to draw the line."); return; }
      if (tool === "draw" && all.length < 2) return;
      const pad = tool === "arrow" ? 10 + width * 3 : width / 2 + 1;
      let bx0 = Math.min(...all.map((p) => p[0])) - pad, by0 = Math.min(...all.map((p) => p[1])) - pad;
      let bx1 = Math.max(...all.map((p) => p[0])) + pad, by1 = Math.max(...all.map((p) => p[1])) + pad;
      bx0 = Math.max(0, bx0); by0 = Math.max(0, by0); bx1 = Math.min(pw, bx1); by1 = Math.min(ph, by1);
      const w = Math.max(1, bx1 - bx0), h = Math.max(1, by1 - by0);
      this.addObj({ page: P.index, type: tool === "draw" ? "ink" : tool, x: bx0, y: by0, w, h, color, width, pts: all.map(([a, b]) => [(a - bx0) / w, (b - by0) / h]) });
    }

    // ----- overlay objects -----
    selObj() { return this.selId == null ? null : this.objs.find((o) => o.id === this.selId) || null; }

    objNode(id) { return this.root.querySelector(`.pdfe-obj[data-id="${id}"]`); }

    addObj(o, opts = {}) {
      this.pushUndo();
      o.id = this.nextObj++;
      this.objs.push(o);
      this.renderLayer(o.page);
      this.selectObj(o.id);
      const node = this.objNode(o.id);
      if (node) {
        const ta = node.querySelector("textarea");
        if (opts.edit && ta) { node.dataset.fresh = "1"; ta.focus(); } else node.focus({ preventScroll: true });
      }
      this.setDirty(true);
      if (!opts.edit) this.hint(`${TYPE_NAMES[o.type] || "Item"} added. Arrow keys move it, Alt+arrows resize, Delete removes, Ctrl+Z undoes.`);
    }

    deleteObj(o) {
      this.pushUndo();
      this.objs = this.objs.filter((x) => x !== o);
      this.selId = null;
      this.renderLayer(o.page);
      this.setDirty(true);
      this.scroller.focus({ preventScroll: true });
      this.hint(`${TYPE_NAMES[o.type] || "Item"} deleted. Ctrl+Z to undo.`);
    }

    selectObj(id) {
      this.selId = id;
      this.root.querySelectorAll(".pdfe-obj").forEach((n) => n.classList.toggle("selected", n.dataset.id === String(id)));
      this.syncProps();
    }

    renderLayer(i) {
      const P = this.pages[i];
      if (!P || !P.layer) return;
      P.layer.querySelectorAll(".pdfe-obj").forEach((n) => n.remove());
      for (const o of this.objs) if (o.page === i) P.layer.append(this.objEl(o));
    }

    refreshObj(o, keepFocus) {
      const old = this.objNode(o.id);
      if (!old) return this.renderLayer(o.page);
      const hadFocus = keepFocus || old.contains(document.activeElement);
      const n = this.objEl(o);
      old.replaceWith(n);
      if (hadFocus) n.focus({ preventScroll: true });
    }

    position(d, o) {
      const cs = this.cs;
      d.style.left = o.x * cs + "px";
      d.style.top = o.y * cs + "px";
      d.style.width = Math.max(1, o.w * cs) + "px";
      d.style.height = Math.max(1, o.h * cs) + "px";
    }

    paintObj(cv, o) {
      const dpr = window.devicePixelRatio || 1, k = this.cs * dpr;
      cv.width = Math.max(1, Math.round(o.w * k));
      cv.height = Math.max(1, Math.round(o.h * k));
      const ctx = cv.getContext("2d");
      ctx.clearRect(0, 0, cv.width, cv.height);
      drawObj(ctx, o, cv.width, cv.height, k);
      if (o.type === "image" && !(imgCache.get(o.src) || {}).naturalWidth) loadImage(o.src).then(() => this.paintObj(cv, o)).catch(() => {});
    }

    describe(o) {
      const name = TYPE_NAMES[o.type] || "Item";
      const what = o.type === "image" && o.sig ? "Signature" : o.type === "stamp" ? `Stamp ${o.text}` : name;
      const extra = o.type === "text" || o.type === "note" ? " Enter to edit text," : "";
      return `${what} on page ${o.page + 1}.${extra} arrow keys move, Alt+arrows resize, Delete removes.`;
    }

    objEl(o) {
      const cs = this.cs;
      const d = el("div", {
        class: "pdfe-obj pdfe-" + o.type + (o.id === this.selId ? " selected" : ""), tabindex: "0", "data-id": String(o.id),
        role: "group", "aria-roledescription": "annotation", "aria-label": this.describe(o),
      });
      this.position(d, o);
      if (o.type === "text" || o.type === "note") {
        const ta = el("textarea", { class: "pdfe-ta", "aria-label": o.type === "text" ? "Text box content" : "Note text", spellcheck: "true" });
        ta.value = o.text || "";
        if (o.type === "text") {
          ta.style.font = `${o.fontSize * cs}px/1.2 ${FONT_CSS[o.font] || FONT_CSS.Helvetica}`;
          ta.style.color = o.color;
        } else {
          ta.style.font = `${(o.fontSize || 10) * cs}px/1.2 ${FONT_CSS.Helvetica}`;
          ta.style.padding = 4 * cs + "px";
        }
        const autosize = () => {
          if (o.type !== "text") return;
          ta.style.height = "0px";
          const h = Math.max(o.fontSize * 1.2, ta.scrollHeight / this.cs);
          ta.style.height = "";
          if (Math.abs(h - o.h) > 0.5) { o.h = h; this.position(d, o); }
        };
        ta.addEventListener("beforeinput", () => {
          if (!d.dataset.typing) { d.dataset.typing = "1"; if (!d.dataset.fresh) this.pushUndo(); }
          delete d.dataset.fresh;
        });
        ta.addEventListener("input", () => { o.text = ta.value; autosize(); this.setDirty(true); });
        ta.addEventListener("focus", () => { delete d.dataset.typing; });
        ta.addEventListener("blur", () => {
          if (o.type === "text" && !o.text.trim() && this.objs.includes(o)) {
            if (d.dataset.fresh) this.undoStack.pop();
            this.objs = this.objs.filter((x) => x !== o);
            if (this.selId === o.id) this.selId = null;
            d.remove();
            this.updateUndoUi();
            if (!this.undoStack.length && !this.hasChanges) this.setDirty(false);
          }
        });
        ta.addEventListener("keydown", (e) => {
          if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); d.focus(); this.hint("Editing finished. " + this.describe(o)); }
        });
        const grip = el("div", { class: "pdfe-grip", "aria-hidden": "true", title: "Drag to move", text: "✥" });
        d.append(grip, ta);
        requestAnimationFrame(autosize);
      } else {
        const cv = el("canvas", { "aria-hidden": "true" });
        this.paintObj(cv, o);
        d.append(cv);
      }
      d.append(el("div", { class: "pdfe-handle", "aria-hidden": "true", title: "Drag to resize" }));
      d.addEventListener("pointerdown", (e) => this.onObjDown(e, o, d));
      d.addEventListener("keydown", (e) => this.onObjKey(e, o, d));
      d.addEventListener("focusin", () => { if (this.selId !== o.id) this.selectObj(o.id); });
      return d;
    }

    onObjDown(e, o, d) {
      if (this.tool !== "select" || e.button !== 0) return;
      const handle = e.target.classList.contains("pdfe-handle");
      if (!handle && e.target.tagName === "TEXTAREA") { this.selectObj(o.id); return; }
      e.preventDefault();
      e.stopPropagation();
      d.focus({ preventScroll: true });
      this.selectObj(o.id);
      const sx = e.clientX, sy = e.clientY, o0 = { x: o.x, y: o.y, w: o.w, h: o.h };
      const P = this.pages[o.page], pw = P.vp.width / this.cs, ph = P.vp.height / this.cs;
      let moved = false;
      d.setPointerCapture(e.pointerId);
      const move = (ev) => {
        const dx = (ev.clientX - sx) / this.cs, dy = (ev.clientY - sy) / this.cs;
        if (!moved) { if (Math.abs(dx) + Math.abs(dy) < 1) return; this.pushUndo(); moved = true; }
        if (handle) {
          let w = Math.max(8, o0.w + dx), h = Math.max(8, o0.h + dy);
          if (o.type === "image" && !ev.shiftKey) h = (w * o0.h) / o0.w;
          o.w = Math.min(w, pw - o.x);
          o.h = Math.min(h, ph - o.y);
        } else {
          o.x = clamp(o0.x + dx, 0, Math.max(0, pw - o.w));
          o.y = clamp(o0.y + dy, 0, Math.max(0, ph - o.h));
        }
        this.position(d, o);
      };
      const up = () => {
        d.removeEventListener("pointermove", move);
        d.removeEventListener("pointerup", up);
        d.removeEventListener("pointercancel", up);
        if (moved) { this.setDirty(true); if (handle) this.refreshObj(o, true); }
      };
      d.addEventListener("pointermove", move);
      d.addEventListener("pointerup", up);
      d.addEventListener("pointercancel", up);
    }

    onObjKey(e, o, d) {
      if (e.target !== d) return;
      const P = this.pages[o.page], pw = P.vp.width / this.cs, ph = P.vp.height / this.cs;
      const arrows = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
      if (arrows[e.key] && !e.ctrlKey && !e.metaKey) {
        const step = e.shiftKey ? 10 : 1;
        const [dx, dy] = arrows[e.key];
        const now = Date.now();
        if (!this.lastKey || this.lastKey.id !== o.id || now - this.lastKey.t > 1000) this.pushUndo();
        this.lastKey = { id: o.id, t: now };
        if (e.altKey) {
          o.w = clamp(o.w + dx * step, 8, pw - o.x);
          o.h = o.type === "text" ? o.h : clamp(o.h + dy * step, 8, ph - o.y);
          this.refreshObj(o, true);
          this.hint(`Size ${Math.round(o.w)} × ${Math.round(o.h)} pt`);
        } else {
          o.x = clamp(o.x + dx * step, 0, Math.max(0, pw - o.w));
          o.y = clamp(o.y + dy * step, 0, Math.max(0, ph - o.h));
          this.position(d, o);
          this.hint(`Position ${Math.round(o.x)}, ${Math.round(o.y)} pt`);
        }
        this.setDirty(true);
      } else if (e.key === "Delete" || e.key === "Backspace") {
        this.deleteObj(o);
      } else if (e.key === "Enter") {
        const ta = d.querySelector("textarea");
        if (!ta) return;
        ta.focus();
      } else if (e.key === "Escape") {
        this.selectObj(null);
        this.scroller.focus({ preventScroll: true });
      } else return;
      e.preventDefault();
      e.stopPropagation();
    }

    placeBox(o, w, h) {
      if (this.viewRot) { App.toast("Reset the view rotation before adding items.", "error"); return; }
      const P = this.pages[this.cur];
      const pw = P.vp.width / this.cs, ph = P.vp.height / this.cs;
      const s = Math.min(1, (pw * 0.8) / w, (ph * 0.8) / h);
      w *= s; h *= s;
      const top = Math.max(0, (this.scroller.scrollTop - P.el.offsetTop) / this.cs);
      const bottom = Math.min(ph, (this.scroller.scrollTop + this.scroller.clientHeight - P.el.offsetTop) / this.cs);
      const cy = bottom > top ? (top + bottom) / 2 : ph / 2;
      Object.assign(o, { page: P.index, x: (pw - w) / 2, y: clamp(cy - h / 2, 0, Math.max(0, ph - h)), w, h });
      this.addObj(o);
    }

    async insertImage() {
      if (this.viewRot) { App.toast("Reset the view rotation before adding items.", "error"); return; }
      const path = await App.pickFile({ mode: "open", extensions: [".png", ".jpg", ".jpeg"] });
      if (!path) return;
      const p = Array.isArray(path) ? path[0] : path;
      const bytes = await App.readFile(p);
      const src = await blobToDataURL(new Blob([bytes], { type: /\.png$/i.test(p) ? "image/png" : "image/jpeg" }));
      const img = await loadImage(src);
      this.setTool("select");
      this.placeBox({ type: "image", src }, img.naturalWidth * 0.75, img.naturalHeight * 0.75);
    }

    async insertSignature() {
      if (this.viewRot) { App.toast("Reset the view rotation before adding items.", "error"); return; }
      const url = await signatureDialog();
      if (!url) return;
      const img = await loadImage(url);
      const w = 160, h = Math.min(70, (160 * img.naturalHeight) / img.naturalWidth);
      this.setTool("select");
      this.placeBox({ type: "image", src: url, sig: true }, h < 70 ? w : (70 * img.naturalWidth) / img.naturalHeight, h);
    }

    async insertStamp() {
      const k = this.stampSel.value, s = STAMPS[k];
      const text = s.text || new Date().toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" }).toUpperCase();
      this.setTool("select");
      this.placeBox({ type: "stamp", text, color: s.color }, 30 + text.length * 13, 42);
    }

    // ----- undo / redo -----
    snapshot() { return { bytes: this.bytes, objs: JSON.stringify(this.objs), pw: this.savePassword }; }
    pushUndo() {
      this.undoStack.push(this.snapshot());
      if (this.undoStack.length > 40) this.undoStack.shift();
      this.redoStack = [];
      this.hasChanges = true;
      this.updateUndoUi();
    }
    updateUndoUi() {
      if (!this.undoBtn) return;
      this.undoBtn.disabled = !this.undoStack.length;
      this.redoBtn.disabled = !this.redoStack.length;
    }
    async undo() {
      if (!this.undoStack.length) { this.hint("Nothing to undo."); return; }
      this.redoStack.push(this.snapshot());
      await this.restore(this.undoStack.pop());
      this.hint("Undone.");
    }
    async redo() {
      if (!this.redoStack.length) { this.hint("Nothing to redo."); return; }
      this.undoStack.push(this.snapshot());
      await this.restore(this.redoStack.pop());
      this.hint("Redone.");
    }
    async restore(s) {
      this.objs = JSON.parse(s.objs);
      this.savePassword = s.pw;
      this.selId = null;
      if (s.bytes !== this.bytes) await this.load(s.bytes, { keepScroll: true });
      else this.pages.forEach((P) => this.renderLayer(P.index));
      this.updateUndoUi();
      this.updateProtUi();
      this.setDirty(true);
    }

    setDirty(v) {
      this.dirty = !!v;
      App.setDirty(this, this.dirty);
    }

    // ----- flattening -----
    objPng(o, rot) {
      let k = 2;
      if (o.type === "image") { const img = imgCache.get(o.src); if (img && img.naturalWidth) k = clamp(img.naturalWidth / o.w, 2, 4); }
      while (o.w * k * o.h * k > 16e6 && k > 0.5) k /= 1.5;
      const D = document.createElement("canvas");
      D.width = Math.max(1, Math.round(o.w * k));
      D.height = Math.max(1, Math.round(o.h * k));
      drawObj(D.getContext("2d"), o, D.width, D.height, k);
      if (!rot) return D.toDataURL("image/png");
      // the page is shown rotated clockwise by rot; draw the image rotated back so it looks right
      const C = document.createElement("canvas");
      const swap = rot === 90 || rot === 270;
      C.width = swap ? D.height : D.width;
      C.height = swap ? D.width : D.height;
      const ctx = C.getContext("2d");
      if (rot === 90) { ctx.translate(0, C.height); ctx.rotate(-Math.PI / 2); }
      else if (rot === 180) { ctx.translate(C.width, C.height); ctx.rotate(Math.PI); }
      else if (rot === 270) { ctx.translate(C.width, 0); ctx.rotate(Math.PI / 2); }
      ctx.drawImage(D, 0, 0);
      return C.toDataURL("image/png");
    }

    annotItems() {
      const items = [];
      for (const o of this.objs) {
        const P = this.pages[o.page];
        if (!P) continue;
        if ((o.type === "text" || o.type === "note") && !(o.text || "").trim() && o.type === "text") continue;
        const vp = P.page.getViewport({ scale: 1, rotation: this.rot(P) });
        const [ax, ay] = vp.convertToPdfPoint(o.x, o.y);
        const [bx, by] = vp.convertToPdfPoint(o.x + o.w, o.y + o.h);
        const v = P.page.view;
        const box = { x: Math.min(ax, bx) - v[0], y: Math.min(ay, by) - v[1], w: Math.abs(bx - ax), h: Math.abs(by - ay) };
        if (o.type === "text" && vp.rotation === 0) {
          const fs = o.fontSize;
          measureCtx.font = `${fs}px ${FONT_CSS[o.font] || FONT_CSS.Helvetica}`;
          const lines = wrapText(measureCtx, o.text, o.w);
          // server puts the first baseline 0.8*fs below the box top; the editor shows it at BASELINE*fs
          const drop = ((BASELINE[o.font] || BASELINE.Helvetica) - 0.8) * fs;
          items.push({ page: o.page + 1, type: "text", x: box.x, y: box.y, w: box.w, h: box.h - drop, text: lines.join("\n"), fontSize: fs, color: o.color, font: o.font });
        } else {
          items.push({ page: o.page + 1, type: "image", x: box.x, y: box.y, w: box.w, h: box.h, png: this.objPng(o, vp.rotation) });
        }
      }
      return items;
    }

    // PDF bytes with all overlays merged in (does not change the editor state)
    async flattenedBytes() {
      if (document.activeElement && this.root.contains(document.activeElement) && document.activeElement.tagName === "TEXTAREA") document.activeElement.blur();
      const items = this.annotItems();
      if (!items.length) return this.bytes;
      this.hint("Applying added items to the pages…");
      return this.post("/api/pdf/annotate", { items: JSON.stringify(items) });
    }

    async flatten(record) {
      if (!this.objs.length) return;
      if (record) this.pushUndo();
      const out = await this.flattenedBytes();
      this.objs = [];
      this.selId = null;
      await this.load(out, { keepScroll: true });
      this.setDirty(true);
      if (record) this.hint("Added items were applied to the PDF. Ctrl+Z to undo.");
    }

    // ----- server operations -----
    async post(endpoint, fields, files, kind) {
      const fd = new FormData();
      if (files) for (const f of files) fd.append("file", f.blob, f.name || "file");
      else fd.append("file", new Blob([this.bytes], { type: "application/pdf" }), "document.pdf");
      for (const [k, v] of Object.entries(fields || {})) fd.append(k, String(v));
      const res = await App.api.postForm(endpoint, fd);
      if (kind === "json") return res.json();
      if (kind === "blob") return res.blob();
      return new Uint8Array(await res.arrayBuffer());
    }

    async pageOp(endpoint, fields, msg) {
      this.hint(msg || "Working…");
      this.pushUndo();
      if (this.objs.length) {
        this.bytes = await this.flattenedBytes();
        this.objs = [];
        this.selId = null;
      }
      const out = await this.post(endpoint, fields);
      await this.load(out, { keepScroll: true });
      this.setDirty(true);
    }

    targetPages() { return this.sel.size ? [...this.sel].sort((a, b) => a - b) : [this.cur]; }

    async rotatePages(deg) {
      const t = this.targetPages();
      for (const i of t) if (this.pages[i]) this.pages[i].thumbDone = false;
      await this.pageOp("/api/pdf/rotate", { pages: rangeStr(t), degrees: deg }, "Rotating pages…");
      this.hint(`Rotated ${pagesLabel(t)} ${deg === 90 ? "clockwise" : "counter-clockwise"}. Ctrl+Z to undo.`);
    }

    async deletePages() {
      const t = this.targetPages();
      if (t.length >= this.pages.length) { App.toast("A PDF must keep at least one page.", "error"); return; }
      if (!(await App.confirm(`Delete ${pagesLabel(t)}? You can undo this with Ctrl+Z.`))) return;
      await this.pageOp("/api/pdf/delete", { pages: rangeStr(t) }, "Deleting pages…");
      this.sel.clear();
      this.updateThumbSel();
      this.hint(`Deleted ${pagesLabel(t)}. Ctrl+Z to undo.`);
    }

    async moveSel(dir) {
      const t = this.targetPages();
      if (dir < 0 && t[0] === 0) { this.hint("Already at the start."); return; }
      if (dir > 0 && t[t.length - 1] === this.pages.length - 1) { this.hint("Already at the end."); return; }
      await this.movePages(t, dir < 0 ? t[0] - 1 : t[t.length - 1] + 2);
    }

    async movePages(moving, before) {
      const n = this.pages.length, set = new Set(moving);
      const rest = [];
      for (let i = 0; i < n; i++) if (!set.has(i)) rest.push(i);
      let pos = 0;
      for (const r of rest) if (r < before) pos++;
      const order = rest.slice();
      order.splice(pos, 0, ...moving);
      if (order.every((v, i) => v === i)) return;
      await this.pageOp("/api/pdf/reorder", { order: order.map((x) => x + 1).join(",") }, "Moving pages…");
      this.sel = new Set(moving.map((_, k) => pos + k));
      this.anchor = pos;
      this.thumbFocus = pos;
      this.goto(pos);
      this.hint(`Moved ${pagesLabel(moving)} to position ${pos + 1}. Ctrl+Z to undo.`);
    }

    async insertDocs(blobs, what) {
      const after = this.targetPages().slice(-1)[0];
      const n0 = this.pages.length;
      this.hint(`Inserting ${what}…`);
      this.pushUndo();
      const base = await this.flattenedBytes();
      this.objs = [];
      let merged = await this.post("/api/pdf/merge", {}, [{ blob: new Blob([base], { type: "application/pdf" }), name: "document.pdf" }, ...blobs]);
      const tmp = await pdfjsLib.getDocument({ data: merged.slice(), isEvalSupported: false }).promise;
      const total = tmp.numPages;
      tmp.destroy();
      const m = total - n0;
      if (after < n0 - 1 && m > 0) {
        const order = [];
        for (let i = 1; i <= after + 1; i++) order.push(i);
        for (let i = n0 + 1; i <= total; i++) order.push(i);
        for (let i = after + 2; i <= n0; i++) order.push(i);
        merged = await this.post("/api/pdf/reorder", { order: order.join(",") }, [{ blob: new Blob([merged], { type: "application/pdf" }), name: "document.pdf" }]);
      }
      await this.load(merged, { keepScroll: true });
      this.setDirty(true);
      this.sel = new Set();
      for (let k = 0; k < m; k++) this.sel.add(after + 1 + k);
      this.updateThumbSel();
      this.goto(after + 1);
      this.hint(`Inserted ${m} page${m === 1 ? "" : "s"} after page ${after + 1}. Ctrl+Z to undo.`);
    }

    async insertPdfFiles() {
      let paths = await App.pickFile({ mode: "open", extensions: [".pdf"], multiple: true });
      if (!paths) return;
      if (!Array.isArray(paths)) paths = [paths];
      if (!paths.length) return;
      const blobs = [];
      for (const p of paths) blobs.push({ blob: new Blob([await App.readFile(p)], { type: "application/pdf" }), name: baseName(p) });
      await this.insertDocs(blobs, paths.length === 1 ? baseName(paths[0]) : paths.length + " files");
    }

    async insertImagePages() {
      let paths = await App.pickFile({ mode: "open", extensions: [".png", ".jpg", ".jpeg"], multiple: true });
      if (!paths) return;
      if (!Array.isArray(paths)) paths = [paths];
      if (!paths.length) return;
      const files = [];
      for (const p of paths) files.push({ blob: new Blob([await App.readFile(p)], { type: /\.png$/i.test(p) ? "image/png" : "image/jpeg" }), name: baseName(p) });
      const pdf = await this.post("/api/pdf/images-to-pdf", {}, files);
      await this.insertDocs([{ blob: new Blob([pdf], { type: "application/pdf" }), name: "images.pdf" }], "images");
    }

    async extractPages() {
      const t = this.targetPages();
      this.hint("Extracting pages…");
      const base = await this.flattenedBytes();
      const out = await this.post("/api/pdf/extract", { pages: rangeStr(t) }, [{ blob: new Blob([base], { type: "application/pdf" }), name: "document.pdf" }]);
      const name = `${stem(this.doc.name)} (pages ${rangeStr(t)}).pdf`;
      if (App.openBytes) { await App.openBytes(name, "", out); return; }
      await this.saveBytes(out, name, ".pdf");
    }

    async splitPages() {
      this.hint("Splitting into single pages…");
      const base = await this.flattenedBytes();
      const zip = await this.post("/api/pdf/split", {}, [{ blob: new Blob([base], { type: "application/pdf" }), name: "document.pdf" }], "blob");
      await this.saveBytes(zip, stem(this.doc.name) + "-pages.zip", ".zip");
    }

    async saveBytes(data, suggestedName, ext) {
      const p = await App.pickFile({ mode: "save", extensions: [ext], suggestedName });
      if (!p) { this.hint("Cancelled."); return; }
      await App.writeFile(p, data);
      App.toast("Saved " + baseName(p), "success");
      this.hint("Saved " + baseName(p) + ".");
    }

    async setPassword() {
      const pw = await passwordDialog("Set password", "People will need this password to open the saved PDF. It is applied when you save.", true);
      if (pw == null) return;
      this.pushUndo();
      this.savePassword = pw;
      this.setDirty(true);
      this.updateProtUi();
      App.toast("Password set. Save the document to apply it.", "success");
    }

    async removePassword() {
      if (!this.savePassword) { App.toast("This document is not password protected.", "info"); return; }
      if (!(await App.confirm("Save this PDF without a password? Anyone will be able to open it."))) return;
      this.pushUndo();
      this.savePassword = null;
      this.setDirty(true);
      this.updateProtUi();
      App.toast("Password removed. Save the document to apply it.", "success");
    }

    updateProtUi() { this.stProt.hidden = !this.savePassword; }

    async watermark() {
      const text = this.wmText.value.trim();
      if (!text) { App.toast("Enter the watermark text first.", "error"); this.wmText.focus(); return; }
      const op = clamp(parseFloat(this.wmOpacity.value) || 0.3, 0.05, 1);
      await this.pageOp("/api/pdf/watermark", { text, opacity: op }, "Adding watermark…");
      for (const P of this.pages) P.thumbDone = false;
      this.hint("Watermark added to every page. Ctrl+Z to undo.");
    }

    async compress() {
      const before = this.bytes.length;
      await this.pageOp("/api/pdf/compress", {}, "Compressing…");
      const after = this.bytes.length;
      const kb = (n) => (n / 1024).toFixed(n < 10240 ? 1 : 0) + " KB";
      App.toast(after < before ? `Compressed from ${kb(before)} to ${kb(after)}.` : `The file is already optimized (${kb(after)}).`, "success");
    }

    // ----- side panels -----
    openSide(title, content) {
      this.sideTitle.textContent = title;
      this.sideBody.replaceChildren(content);
      this.side.hidden = false;
      this.sideTitle.focus();
    }
    closeSide() {
      this.side.hidden = true;
      this.sideBody.textContent = "";
      this.scroller.focus({ preventScroll: true });
    }

    async showForms() {
      this.hint("Reading form fields…");
      const fields = await this.post("/api/pdf/form-fields", {}, null, "json");
      if (!fields || !fields.length) {
        this.openSide("Form fields", el("p", { text: "This PDF has no fillable form fields. Use Edit > Add text to type on the page instead." }));
        this.hint("No form fields found.");
        return;
      }
      const form = el("form", { class: "pdfe-form" });
      const controls = [];
      fields.forEach((f, i) => {
        const id = this.id("ff-" + i);
        const key = f.name || f.id;
        const type = (f.type || "").toLowerCase();
        const label = (f.name || f.id || "Field " + (i + 1)) + (f.locked ? " (read only)" : "");
        let input;
        if (type.includes("check")) {
          input = el("input", { type: "checkbox", id, checked: /^(true|yes|on|1|x)$/i.test(f.value || ""), disabled: !!f.locked });
          form.append(el("div", { class: "pdfe-form-row pdfe-check" }, input, el("label", { for: id, text: label })));
        } else {
          const opts = (f.options || "").split(/[,;|]/).map((s) => s.trim()).filter(Boolean);
          if (opts.length) {
            if (f.value && !opts.includes(f.value)) opts.unshift(f.value);
            input = el("select", { id, disabled: !!f.locked }, el("option", { value: "", text: "(none)" }), opts.map((o) => el("option", { value: o, text: o, selected: o === f.value })));
          } else {
            input = el("input", { type: type === "date" ? "text" : "text", id, value: f.value || "", disabled: !!f.locked, placeholder: type === "date" ? "Date" : null });
          }
          form.append(el("div", { class: "pdfe-form-row" }, el("label", { for: id, text: label }), input));
        }
        controls.push({ key, type, input, locked: !!f.locked });
      });
      form.append(el("div", { class: "pdfe-row" }, el("button", { type: "submit", class: "btn primary", text: "Apply to PDF" })));
      form.addEventListener("submit", (e) => {
        e.preventDefault();
        const values = {};
        for (const c of controls) {
          if (c.locked) continue;
          if (c.type.includes("check")) values[c.key] = c.input.checked;
          else if (c.type.includes("list")) values[c.key] = c.input.value ? [c.input.value] : [];
          else values[c.key] = c.input.value;
        }
        this.run(async () => {
          await this.pageOp("/api/pdf/form-fill", { values: JSON.stringify(values) }, "Filling form…");
          App.toast("Form fields filled.", "success");
          this.hint("Form fields filled. Ctrl+Z to undo.");
        });
      });
      this.openSide(`Form fields (${fields.length})`, form);
      this.hint(`${fields.length} form fields found.`);
    }

    async showInfo() {
      const info = await this.post("/api/pdf/info", {}, null, "json");
      const sizes = new Map();
      for (const s of info.pageSizes || []) {
        const key = `${Math.round(s.w)} × ${Math.round(s.h)} pt (${Math.round(s.w / 72 * 25.4)} × ${Math.round(s.h / 72 * 25.4)} mm)`;
        sizes.set(key, (sizes.get(key) || 0) + 1);
      }
      const kb = this.bytes.length / 1024;
      const rows = [
        ["File name", this.doc.name],
        ["Location", this.doc.path || "Not saved yet"],
        ["File size", kb < 1024 ? kb.toFixed(1) + " KB" : (kb / 1024).toFixed(2) + " MB"],
        ["Pages", String(info.pages != null ? info.pages : this.pages.length)],
        ["Title", info.title || "—"],
        ["Author", info.author || "—"],
        ["Password on save", this.savePassword ? "Yes" : "No"],
        ["Page sizes", [...sizes].map(([k, c]) => (sizes.size > 1 ? `${c} × ` : "") + k).join("; ") || "—"],
      ];
      this.openSide("Document properties", el("dl", { class: "pdfe-props" }, rows.map(([k, v]) => [el("dt", { text: k }), el("dd", { text: v })])));
    }

    // ----- find -----
    async buildTextIdx(i) {
      if (this.textIdx[i]) return this.textIdx[i];
      const tc = await this.getText(i);
      let full = "";
      const starts = [], lens = [];
      for (const it of tc.items) {
        starts.push(full.length);
        lens.push((it.str || "").length);
        full += it.str || "";
        if (it.hasEOL) full += " ";
      }
      return (this.textIdx[i] = { full: full.toLowerCase(), starts, lens });
    }

    async runFind(dir, silent) {
      const q = this.findInput.value.trim().toLowerCase();
      if (!q) { this.clearFind(); return; }
      if (q !== this.find.q || silent) {
        this.find.q = q;
        this.findCount.textContent = "Searching…";
        const matches = [];
        for (let i = 0; i < this.pages.length; i++) {
          const t = await this.buildTextIdx(i);
          let pos = t.full.indexOf(q);
          while (pos !== -1) { matches.push({ page: i, start: pos, end: pos + q.length }); pos = t.full.indexOf(q, pos + q.length); }
        }
        this.find.matches = matches;
        if (silent) { this.find.idx = Math.min(this.find.idx, matches.length - 1); }
        else {
          const k = matches.findIndex((m) => m.page >= this.cur);
          this.find.idx = dir < 0 ? (k <= 0 ? matches.length - 1 : k - 1) : (k < 0 ? 0 : k);
        }
      } else if (this.find.matches.length) {
        const n = this.find.matches.length;
        this.find.idx = (this.find.idx + (dir || 1) + n) % n;
      }
      const n = this.find.matches.length;
      this.findCount.textContent = n ? `${this.find.idx + 1} of ${n}` : "No matches";
      if (!n) { this.refreshHits(); return; }
      const m = this.find.matches[this.find.idx];
      if (!silent) {
        this.find.scrollTo = this.find.idx;
        if (m.page !== this.cur || !this.visible.has(m.page)) this.goto(m.page);
      }
      this.refreshHits();
    }

    clearFind() {
      this.find = { q: "", matches: [], idx: -1, scrollTo: -1 };
      this.findCount.textContent = "";
      this.refreshHits();
    }

    refreshHits() { for (const P of this.pages) this.applyHits(P.index); }

    applyHits(i) {
      const P = this.pages[i];
      if (!P || !P.textDivs || !P.textEl) return;
      P.textEl.querySelectorAll(".pdfe-hit").forEach((n) => n.classList.remove("pdfe-hit", "pdfe-hit-cur"));
      const t = this.textIdx[i];
      if (!t || !this.find.matches.length) return;
      let curDiv = null;
      this.find.matches.forEach((m, mi) => {
        if (m.page !== i) return;
        for (let j = 0; j < t.starts.length; j++) {
          if (t.starts[j] < m.end && t.starts[j] + t.lens[j] > m.start && P.textDivs[j]) {
            P.textDivs[j].classList.add("pdfe-hit");
            if (mi === this.find.idx) { P.textDivs[j].classList.add("pdfe-hit-cur"); curDiv = curDiv || P.textDivs[j]; }
          }
        }
      });
      if (curDiv && this.find.scrollTo === this.find.idx) {
        this.find.scrollTo = -1;
        curDiv.scrollIntoView({ block: "center", inline: "nearest" });
      }
    }

    // ----- print -----
    async print() {
      this.hint("Preparing pages for printing…");
      const k = 150 / 72;
      const urls = [];
      try {
        for (const P of this.pages) {
          const vp = P.page.getViewport({ scale: k, rotation: this.rot(P) });
          const c = document.createElement("canvas");
          c.width = Math.floor(vp.width);
          c.height = Math.floor(vp.height);
          const ctx = c.getContext("2d");
          ctx.fillStyle = "#ffffff";
          ctx.fillRect(0, 0, c.width, c.height);
          await P.page.render({ canvasContext: ctx, viewport: vp }).promise;
          for (const o of this.objs) {
            if (o.page !== P.index) continue;
            ctx.save();
            ctx.translate(o.x * k, o.y * k);
            drawObj(ctx, o, o.w * k, o.h * k, k);
            ctx.restore();
          }
          const blob = await new Promise((res) => c.toBlob(res, "image/png"));
          urls.push(URL.createObjectURL(blob));
          this.hint(`Preparing page ${P.index + 1} of ${this.pages.length} for printing…`);
        }
      } catch (err) { urls.forEach((u) => URL.revokeObjectURL(u)); throw err; }
      const frame = el("iframe", { class: "pdfe-print-frame", "aria-hidden": "true", tabindex: "-1", title: "Print preview" });
      document.body.append(frame);
      const d = frame.contentDocument;
      d.open();
      d.write(`<!doctype html><html><head><title>${this.doc.name.replace(/[<&>"]/g, "")}</title><style>@page{margin:0}html,body{margin:0;padding:0;background:#fff}img{display:block;width:100%;height:auto;max-height:100vh;object-fit:contain;break-after:page;page-break-after:always}img:last-child{break-after:auto;page-break-after:auto}</style></head><body>${urls.map((u, i) => `<img src="${u}" alt="Page ${i + 1}">`).join("")}</body></html>`);
      d.close();
      await Promise.all([...d.images].map((img) => (img.complete ? null : new Promise((r) => { img.onload = img.onerror = r; }))));
      const cleanup = () => { frame.remove(); urls.forEach((u) => URL.revokeObjectURL(u)); };
      frame.contentWindow.addEventListener("afterprint", () => setTimeout(cleanup, 100));
      setTimeout(() => { if (frame.isConnected) cleanup(); }, 120000);
      this.hint("");
      frame.contentWindow.focus();
      frame.contentWindow.print();
    }

    // ----- save -----
    async save() {
      if (!this.doc.path || !/\.pdf$/i.test(this.doc.path)) return this.saveAs();
      return this.writeTo(this.doc.path);
    }

    async saveAs() {
      const p = await App.pickFile({ mode: "save", extensions: [".pdf"], suggestedName: stem(this.doc.name) + ".pdf" });
      if (!p) return false;
      return this.writeTo(Array.isArray(p) ? p[0] : p);
    }

    async writeTo(path) {
      if (this.busy) throw new Error("Please wait for the current operation to finish.");
      this.busy = true;
      this.root.setAttribute("aria-busy", "true");
      try {
        this.hint("Saving…");
        if (this.objs.length) await this.flatten(true);
        let out = this.bytes;
        if (this.savePassword) out = await this.post("/api/pdf/encrypt", { password: this.savePassword });
        const saved = await App.writeFile(path, out);
        this.doc.path = typeof saved === "string" && saved ? saved : path;
        this.doc.name = baseName(this.doc.path);
        App.setTitle(this, this.doc.name);
        this.setDirty(false);
        if (App.addRecent) { try { App.addRecent(this.doc.path); } catch (_) { /* ignore */ } }
        App.toast("Saved " + this.doc.name, "success");
        this.hint("Saved.");
        return true;
      } finally {
        this.busy = false;
        this.root.removeAttribute("aria-busy");
      }
    }

    isDirty() { return this.dirty; }

    focus() { this.scroller.focus({ preventScroll: true }); }

    destroy() {
      if (this.io) this.io.disconnect();
      if (this.thumbIo) this.thumbIo.disconnect();
      if (this.ro) this.ro.disconnect();
      if (this.pdf) { try { this.pdf.destroy(); } catch (_) { /* ignore */ } }
      this.pdf = null;
      this.pages = [];
      this.root.remove();
    }

    // ----- keyboard -----
    onKey(e) {
      const ctrl = e.ctrlKey || e.metaKey;
      const k = e.key;
      const typing = isTyping(e.target);
      if (ctrl && !e.altKey) {
        const lk = k.toLowerCase();
        if (lk === "f") { e.preventDefault(); this.showTab("home"); this.findInput.focus(); this.findInput.select(); }
        else if (k === "=" || k === "+") { e.preventDefault(); this.zoomStep(1); }
        else if (k === "-") { e.preventDefault(); this.zoomStep(-1); }
        else if (k === "0") { e.preventDefault(); this.setZoom(1); }
        else if (lk === "p") { e.preventDefault(); this.run(() => this.print()); }
        else if (!typing && lk === "z" && !e.shiftKey) { e.preventDefault(); this.run(() => this.undo()); }
        else if (!typing && (lk === "y" || (lk === "z" && e.shiftKey))) { e.preventDefault(); this.run(() => this.redo()); }
        return;
      }
      if (typing || e.altKey) return;
      if (k === "PageDown") { e.preventDefault(); this.goto(this.cur + 1); return; }
      if (k === "PageUp") { e.preventDefault(); this.goto(this.cur - 1); return; }
      if (k === "Escape") {
        if (this.tool !== "select") { this.setTool("select"); e.preventDefault(); }
        else if (this.selId) { this.selectObj(null); e.preventDefault(); }
        return;
      }
      if (k.length !== 1 || e.shiftKey || e.target.closest(".ribbon-tabs,[role=tablist]")) return;
      const lk = k.toLowerCase();
      const t = TOOLS.find((x) => x.key.toLowerCase() === lk);
      if (t) { e.preventDefault(); this.setTool(t.id); this.focus(); return; }
      if (lk === "i") { e.preventDefault(); this.run(() => this.insertImage()); }
      else if (lk === "s") { e.preventDefault(); this.run(() => this.insertSignature()); }
      else if (lk === "p") { e.preventDefault(); this.run(() => this.insertStamp()); }
    }
  }

  App.registerEditor("pdf", {
    label: "PDF",
    icon: "📕",
    extensions: [".pdf"],
    async create(container, doc) {
      const ed = new PdfEditor(container, doc);
      try { await ed.open(); } catch (err) { ed.destroy(); throw err; }
      return ed;
    },
  });
})();
