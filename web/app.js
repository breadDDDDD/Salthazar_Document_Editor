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
  function iconFor(name) {
    const e = extOf(name);
    if (e === ".pdf") return "📕";
    if (IMAGE_EXTS.includes(e)) return "🖼";
    if (e === ".docx") return "📘";
    return "📄";
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
    const t = el("div", { class: "toast " + kind },
      el("span", { class: "toast-msg", text: String(message) }),
      el("button", { type: "button", "aria-label": "Dismiss notification", title: "Dismiss", text: "✕", onclick: () => t.remove() }));
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
          el("button", { type: "button", class: "icon-btn", "aria-label": "Close dialog", title: "Close (Esc)", text: "✕", onclick: () => close(opts.cancelValue === undefined ? null : opts.cancelValue) })),
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
    const upBtn = el("button", { type: "button", class: "btn", title: "Up one folder (Backspace)", "aria-label": "Up one folder", text: "⬆ Up" });
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
          el("span", { "aria-hidden": "true", text: e.isDir ? "📁" : iconFor(e.name) }),
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
      const add = (name, path, icon) => places.append(el("button", { type: "button", title: path, onclick: () => navigate(path, true) }, icon + " " + name));
      add("Home", h.path, "🏠");
      (h.places || []).forEach((p) => add(p.name, p.path, p.name === "Downloads" ? "⬇" : p.name === "Desktop" ? "🖥" : "📁"));
      if (/^[a-zA-Z]:/.test(h.path || "")) add("This PC", "", "💻");
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
          el("span", { class: "recent-icon", "aria-hidden": "true", text: iconFor(r.name) }),
          el("span", { class: "recent-name", text: r.name }),
          el("span", { class: "recent-date", text: fmtDate(r.time) }),
          el("span", { class: "recent-path", text: r.path })),
        el("button", { type: "button", class: "icon-btn", "aria-label": "Remove " + r.name + " from recent files", title: "Remove from list", text: "✕", onclick: () => removeRecent(r.path) })));
    }
  }

  // ---------- editors & documents ----------
  function registerEditor(kind, def) {
    editors[kind] = Object.assign({ label: kind, extensions: [], icon: "📄" }, def);
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
    }, el("span", { class: "doc-tab-icon", "aria-hidden": "true", text: def.icon }), nameEl, dirtyMark);
    const closeBtn = el("button", { type: "button", class: "doc-tab-close", tabindex: "-1", "aria-label": "Close " + name, title: "Close (Ctrl+W)", text: "✕" });
    const wrap = el("div", { class: "doc-tab-wrap", role: "presentation" }, tab, closeBtn);
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
    const ext = def && def.extensions && def.extensions[0] ? def.extensions[0] : ".docx";
    return openDocument(kind, { name: "Document " + untitled++ + ext, path: "", bytes: null });
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
    $("tile-new-doc").addEventListener("click", () => newDocument("word"));
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
  };

  if (document.readyState === "complete") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
