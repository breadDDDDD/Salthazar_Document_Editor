/* DocStudio - Word-like editor for .docx, .txt, .md, .html documents.
 * Plain ES2020, no dependencies. Registers the "word" editor with window.App. */
(function () {
  'use strict';

  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const FONTS = ['Calibri', 'Arial', 'Cambria', 'Georgia', 'Times New Roman', 'Garamond', 'Verdana',
    'Segoe UI', 'Tahoma', 'Trebuchet MS', 'Courier New', 'Consolas', 'Comic Sans MS'];
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72];
  const ZOOMS = [50, 75, 90, 100, 110, 125, 150, 175, 200];
  const SPECIAL_CHARS = ['—', '–', '…', '©', '®', '™', '§', '¶',
    '†', '‡', '•', '·', '“', '”', '‘', '’', '«', '»',
    '¡', '¿', '°', ' ', '­', '€', '£', '¥', '¢', '‰',
    'é', 'è', 'à', 'ç', 'ñ', 'ü', 'ö', 'ä', 'ß', 'ø'];
  const SYMBOLS = ['±', '×', '÷', '≠', '≈', '≤', '≥', '∞',
    '√', '∑', '∏', '∫', '∂', '∆', '∈', '∩', '∪', '¹',
    '²', '³', '½', '¼', '¾', 'α', 'β', 'γ', 'δ', 'ε',
    'θ', 'λ', 'μ', 'π', 'σ', 'ω', 'Ω', '←', '↑', '→',
    '↓', '↔', '⇒', '✓', '✗', '★', '☆', '♥', '☺', '☎'];
  const CHAR_NAMES = { ' ': 'Non-breaking space', '­': 'Soft hyphen', '—': 'Em dash',
    '–': 'En dash', '…': 'Ellipsis', '©': 'Copyright', '®': 'Registered',
    '™': 'Trademark', '§': 'Section', '¶': 'Paragraph', '•': 'Bullet', '€': 'Euro',
    '£': 'Pound', '°': 'Degree', '±': 'Plus minus', '×': 'Multiply', '÷': 'Divide',
    '≠': 'Not equal', '≤': 'Less or equal', '≥': 'Greater or equal', '∞': 'Infinity',
    '→': 'Right arrow', '←': 'Left arrow', '✓': 'Check mark', 'π': 'Pi' };
  const PAGE_SIZES = { letter: { w: 8.5, h: 11, css: 'letter' }, a4: { w: 8.27, h: 11.69, css: 'A4' } };
  const MARGINS = { normal: { y: 1, x: 1 }, narrow: { y: 0.5, x: 0.5 }, wide: { y: 1, x: 2 } };
  const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,blockquote,pre,li,div,td,th';
  const SAVE_EXTS = ['.docx', '.txt', '.md', '.html'];
  let editorCounter = 0;

  /* ---------- small helpers ---------- */
  function h(tag, props, ...children) {
    const el = document.createElement(tag);
    if (props) {
      for (const [k, v] of Object.entries(props)) {
        if (v == null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'text') el.textContent = v;
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
  function extOf(name) {
    const m = /\.[^./\\]+$/.exec(name || '');
    return m ? m[0].toLowerCase() : '';
  }
  function baseName(p) { return (p || '').split(/[\\/]/).pop(); }
  function stripExt(n) { return n.replace(/\.[^.]+$/, ''); }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function safeUrl(u) {
    const s = String(u || '').trim();
    return /^\s*(javascript|vbscript|data:text\/html)/i.test(s) ? '#' : s;
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

  /* ---------- HTML sanitising ---------- */
  const DROP_TAGS = 'script,style,iframe,frame,frameset,object,embed,link,meta,base,noscript,template,form,input,button,select,textarea,svg,math,title';
  function sanitizeHtml(html) {
    const docu = new DOMParser().parseFromString(String(html || ''), 'text/html');
    docu.querySelectorAll(DROP_TAGS).forEach(n => n.remove());
    docu.querySelectorAll('*').forEach(el => {
      for (const a of Array.from(el.attributes)) {
        const n = a.name.toLowerCase();
        if (n.startsWith('on') || n === 'contenteditable' || n === 'id' || n === 'tabindex') el.removeAttribute(a.name);
        else if ((n === 'href' || n === 'src' || n === 'xlink:href' || n === 'action') && safeUrl(a.value) === '#') el.removeAttribute(a.name);
      }
    });
    // MS Office comments / conditional junk
    const walker = docu.createTreeWalker(docu.body, NodeFilter.SHOW_COMMENT);
    const comments = [];
    while (walker.nextNode()) comments.push(walker.currentNode);
    comments.forEach(c => c.remove());
    return docu.body.innerHTML;
  }
  function textToHtml(text) {
    const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
    if (lines.length && lines[lines.length - 1] === '') lines.pop();
    if (!lines.length) return '<p><br></p>';
    return lines.map(l => l === '' ? '<p><br></p>' : '<p>' + escapeHtml(l) + '</p>').join('');
  }

  /* ---------- Markdown -> HTML ---------- */
  function mdInline(src) {
    let t = escapeHtml(src);
    const codes = [];
    t = t.replace(/`([^`]+)`/g, (m, c) => { codes.push(c); return '\u0000' + (codes.length - 1) + '\u0000'; });
    t = t.replace(/!\[([^\]]*)\]\(([^)\s]+)(?:\s+&quot;[^&]*&quot;)?\)/g, (m, a, u) => '<img alt="' + a + '" src="' + safeUrl(u) + '">');
    t = t.replace(/\[([^\]]+)\]\(([^)\s]+)(?:\s+&quot;[^&]*&quot;)?\)/g, (m, x, u) => '<a href="' + safeUrl(u) + '">' + x + '</a>');
    t = t.replace(/&lt;(https?:\/\/[^\s&]+)&gt;/g, '<a href="$1">$1</a>');
    t = t.replace(/(\*\*|__)(?=\S)(.+?)\1/g, '<strong>$2</strong>');
    t = t.replace(/\*(?=\S)(.+?)\*/g, '<em>$1</em>');
    t = t.replace(/(^|[^\w])_(?=\S)(.+?)_(?!\w)/g, '$1<em>$2</em>');
    t = t.replace(/~~(.+?)~~/g, '<s>$1</s>');
    t = t.replace(/ {2,}$/g, '<br>');
    t = t.replace(/\u0000(\d+)\u0000/g, (m, i) => '<code>' + codes[+i] + '</code>');
    return t;
  }
  function mdTableCells(line) {
    let s = line.trim();
    if (s.startsWith('|')) s = s.slice(1);
    if (s.endsWith('|')) s = s.slice(0, -1);
    return s.split('|').map(c => c.trim());
  }
  function mdToHtml(md) {
    const lines = String(md).replace(/\r\n?/g, '\n').split('\n');
    const out = [];
    let para = [];
    const listRe = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/;
    const flush = () => {
      if (para.length) out.push('<p>' + para.map(mdInline).join(' ').replace(/<br> /g, '<br>') + '</p>');
      para = [];
    };
    let i = 0;
    while (i < lines.length) {
      const line = lines[i];
      let m;
      if (/^\s*$/.test(line)) { flush(); i++; continue; }
      if ((m = /^\s*(```|~~~)/.exec(line))) {
        flush();
        const fence = m[1];
        const code = [];
        i++;
        while (i < lines.length && !lines[i].trim().startsWith(fence)) code.push(lines[i++]);
        i++;
        out.push('<pre><code>' + escapeHtml(code.join('\n')) + '</code></pre>');
        continue;
      }
      if ((m = /^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$/.exec(line))) {
        flush();
        out.push('<h' + m[1].length + '>' + mdInline(m[2]) + '</h' + m[1].length + '>');
        i++; continue;
      }
      if (/^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(line)) { flush(); out.push('<hr>'); i++; continue; }
      if (/^\s{0,3}>/.test(line)) {
        flush();
        const q = [];
        while (i < lines.length && /^\s{0,3}>/.test(lines[i])) q.push(lines[i++].replace(/^\s{0,3}>\s?/, ''));
        out.push('<blockquote>' + mdToHtml(q.join('\n')) + '</blockquote>');
        continue;
      }
      if (line.includes('|') && i + 1 < lines.length && /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/.test(lines[i + 1]) && lines[i + 1].includes('-')) {
        flush();
        const head = mdTableCells(line);
        i += 2;
        let html = '<table><thead><tr>' + head.map(c => '<th>' + mdInline(c) + '</th>').join('') + '</tr></thead><tbody>';
        while (i < lines.length && lines[i].includes('|') && lines[i].trim()) {
          html += '<tr>' + mdTableCells(lines[i]).map(c => '<td>' + mdInline(c) + '</td>').join('') + '</tr>';
          i++;
        }
        out.push(html + '</tbody></table>');
        continue;
      }
      if (listRe.test(line)) {
        flush();
        const items = [];
        while (i < lines.length && (m = listRe.exec(lines[i]))) {
          items.push({ ind: m[1].replace(/\t/g, '    ').length, ord: /\d/.test(m[2]), text: m[3] });
          i++;
          // lazy continuation lines
          while (i < lines.length && lines[i].trim() && !listRe.test(lines[i]) && /^\s+/.test(lines[i])) {
            items[items.length - 1].text += ' ' + lines[i].trim();
            i++;
          }
        }
        let html = '';
        const stack = [];
        for (const it of items) {
          while (stack.length && it.ind < stack[stack.length - 1].ind) html += '</li></' + stack.pop().tag + '>';
          const top = stack[stack.length - 1];
          if (!top || it.ind > top.ind) {
            const tag = it.ord ? 'ol' : 'ul';
            html += '<' + tag + '>';
            stack.push({ ind: it.ind, tag });
          } else html += '</li>';
          html += '<li>' + mdInline(it.text);
        }
        while (stack.length) html += '</li></' + stack.pop().tag + '>';
        out.push(html);
        continue;
      }
      para.push(line.trim() + (/ {2,}$/.test(line) ? '  ' : ''));
      i++;
    }
    flush();
    return out.join('\n') || '<p><br></p>';
  }

  /* ---------- HTML -> Markdown ---------- */
  function htmlToMd(root) {
    function wrap(t, mark) {
      const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(t);
      return m[2] ? m[1] + mark + m[2] + mark + m[3] : t;
    }
    function kids(n, d) { return Array.from(n.childNodes).map(c => conv(c, d)).join(''); }
    function conv(n, depth) {
      if (n.nodeType === 3) return n.nodeValue.replace(/\s+/g, ' ').replace(/([*_`\\])/g, '\\$1');
      if (n.nodeType !== 1) return '';
      const tag = n.tagName.toLowerCase();
      const st = n.style || {};
      switch (tag) {
        case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
          return '\n\n' + '#'.repeat(+tag[1]) + ' ' + kids(n, depth).trim() + '\n\n';
        case 'p': case 'div': return '\n\n' + kids(n, depth).trim() + '\n\n';
        case 'br': return '  \n';
        case 'b': case 'strong': return wrap(kids(n, depth), '**');
        case 'i': case 'em': return wrap(kids(n, depth), '*');
        case 's': case 'strike': case 'del': return wrap(kids(n, depth), '~~');
        case 'code': return n.closest('pre') ? n.textContent : '`' + n.textContent + '`';
        case 'a': return '[' + kids(n, depth).trim() + '](' + (n.getAttribute('href') || '') + ')';
        case 'img': return '![' + (n.getAttribute('alt') || '') + '](' + (n.getAttribute('src') || '') + ')';
        case 'hr': return '\n\n---\n\n';
        case 'pre': return '\n\n```\n' + n.textContent.replace(/\n$/, '') + '\n```\n\n';
        case 'blockquote': {
          const inner = kids(n, depth).replace(/\n{3,}/g, '\n\n').trim();
          return '\n\n' + inner.split('\n').map(l => '> ' + l).join('\n') + '\n\n';
        }
        case 'ul': case 'ol': {
          let idx = 0;
          const items = Array.from(n.children).filter(c => c.tagName === 'LI').map(li => {
            idx++;
            const marker = tag === 'ol' ? idx + '. ' : '- ';
            const body = kids(li, depth + 1).replace(/\n{2,}/g, '\n').trim();
            return '  '.repeat(depth) + marker + body;
          });
          return (depth ? '\n' : '\n\n') + items.join('\n') + (depth ? '' : '\n\n');
        }
        case 'table': {
          const rows = Array.from(n.querySelectorAll('tr')).map(tr =>
            '| ' + Array.from(tr.cells).map(c => kids(c, depth).replace(/\s+/g, ' ').trim().replace(/\|/g, '\\|')).join(' | ') + ' |');
          if (!rows.length) return '';
          const cols = n.querySelector('tr').cells.length;
          rows.splice(1, 0, '|' + ' --- |'.repeat(cols));
          return '\n\n' + rows.join('\n') + '\n\n';
        }
        case 'span': case 'font': {
          let t = kids(n, depth);
          if (/bold|[6-9]00/.test(st.fontWeight || '')) t = wrap(t, '**');
          if (st.fontStyle === 'italic') t = wrap(t, '*');
          return t;
        }
        default: return kids(n, depth);
      }
    }
    return conv(root, 0).replace(/[ \t]+\n/g, m => (m.startsWith('  ') ? '  \n' : '\n')).replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }

  /* ---------- editor ---------- */
  function createWordEditor(container, doc) {
    const uid = 'we' + (++editorCounter);
    const listeners = [];
    const on = (target, type, fn, opts) => { target.addEventListener(type, fn, opts); listeners.push([target, type, fn, opts]); };

    let docPath = doc.path || '';
    let docName = doc.name || (docPath ? baseName(docPath) : 'Document1.docx');
    let dirty = false;
    let savedRange = null;
    let zoom = 100;
    let layout = { margin: 'normal', orient: 'portrait', size: 'letter' };
    let loading = true;
    let readMode = false;
    let selectedImg = null;
    let uiTimer = 0;
    let statusTimer = 0;
    const toggles = [];   // {cmd, btn}
    let editor;           // public object

    /* ---- DOM ---- */
    const root = h('div', { class: 'word-editor', id: uid });
    const body = h('div', {
      class: 'word-body', contenteditable: 'false', role: 'textbox', 'aria-multiline': 'true',
      'aria-label': 'Document body', spellcheck: 'true', tabindex: '0'
    });
    const page = h('div', { class: 'word-page' }, body);
    const canvas = h('div', { class: 'word-canvas' }, page);
    const exitRead = h('button', { type: 'button', class: 'tool-btn word-exit-read', hidden: true, title: 'Exit read mode (Esc)', onclick: () => setReadMode(false) }, 'Exit read mode');
    canvas.appendChild(exitRead);

    function btn(o) {
      const title = o.title || o.label;
      const b = h('button', {
        type: 'button', class: 'tool-btn' + (o.big ? ' big' : '') + (o.cls ? ' ' + o.cls : ''),
        title: title + (o.keys ? ' (' + o.keys + ')' : ''),
        'aria-label': o.iconOnly ? o.label : null,
        'aria-keyshortcuts': o.aria || null,
        'aria-pressed': o.toggle || o.pressable ? 'false' : null
      },
      o.icon ? h('span', { class: 'ico', 'aria-hidden': 'true', style: o.iconStyle || null }, o.icon) : null,
      o.iconOnly ? null : h('span', { class: 'lbl' }, o.label));
      b.addEventListener('mousedown', e => e.preventDefault()); // keep the document selection
      b.addEventListener('click', () => { Promise.resolve().then(() => o.run()).catch(err => App.toast((o.errPrefix || '') + (err.message || String(err)), 'error')); });
      if (o.toggle) toggles.push({ cmd: o.toggle, btn: b });
      return b;
    }
    function sel(label, options, onChange, opts = {}) {
      const s = h('select', { class: 'tool-select', 'aria-label': label, title: label + (opts.keys ? ' (' + opts.keys + ')' : '') },
        options.map(o => h('option', { value: o.value, text: o.label })));
      s.addEventListener('change', () => { onChange(s.value); });
      return s;
    }
    function group(label, ...items) {
      return h('div', { class: 'tool-group', role: 'group', 'aria-label': label },
        h('div', { class: 'tool-group-items' }, items),
        h('div', { class: 'tool-group-label', 'aria-hidden': 'true' }, label));
    }
    function colorPair(label, cmd, initial, icon) {
      const input = h('input', { type: 'color', class: 'tool-input word-color', value: initial, 'aria-label': label + ' - choose color', title: label + ' - choose color' });
      const b = btn({ icon, label: label, iconOnly: true, iconStyle: 'border-bottom:4px solid ' + initial, run: () => applyColor() });
      function applyColor() {
        const v = input.value;
        if (cmd === 'hiliteColor') { restoreSel(); exec('styleWithCSS', true); if (!document.execCommand('hiliteColor', false, v)) document.execCommand('backColor', false, v); afterEdit(); }
        else cmd2(cmd, v, true);
      }
      input.addEventListener('change', () => { b.querySelector('.ico').style.borderBottomColor = input.value; applyColor(); });
      return [b, input];
    }

    /* ---- ribbon tabs ---- */
    const tabsEl = h('div', { class: 'ribbon-tabs', role: 'tablist', 'aria-label': 'Document ribbon' });
    const ribbon = h('div', { class: 'ribbon word-ribbon' }, tabsEl);
    const tabs = [];
    function addTab(id, label, content, contextual) {
      const tabId = uid + '-tab-' + id, panelId = uid + '-panel-' + id;
      const tab = h('button', { type: 'button', role: 'tab', id: tabId, 'aria-controls': panelId, 'aria-selected': 'false', tabindex: '-1', class: 'ribbon-tab' + (contextual ? ' ribbon-tab-contextual' : '') }, label);
      const toolbar = h('div', { role: 'toolbar', 'aria-label': label + ' tools', class: 'word-toolbar' }, content);
      const panel = h('div', { role: 'tabpanel', id: panelId, 'aria-labelledby': tabId, class: 'ribbon-panel', hidden: true }, toolbar);
      tab.addEventListener('mousedown', e => e.preventDefault());
      tab.addEventListener('click', () => selectTab(id, true));
      toolbar.addEventListener('keydown', toolbarArrows);
      if (contextual) tab.hidden = true;
      tabsEl.appendChild(tab);
      ribbon.appendChild(panel);
      tabs.push({ id, tab, panel });
    }
    function selectTab(id, focusTab) {
      for (const t of tabs) {
        const on = t.id === id;
        t.tab.setAttribute('aria-selected', String(on));
        t.tab.tabIndex = on ? 0 : -1;
        t.panel.hidden = !on;
        if (on && focusTab) t.tab.focus();
      }
    }
    function activeTab() { return tabs.find(t => t.tab.getAttribute('aria-selected') === 'true'); }
    tabsEl.addEventListener('keydown', e => {
      const vis = tabs.filter(t => !t.tab.hidden);
      let i = vis.findIndex(t => t.tab === document.activeElement);
      if (i < 0) return;
      if (e.key === 'ArrowRight') i = (i + 1) % vis.length;
      else if (e.key === 'ArrowLeft') i = (i - 1 + vis.length) % vis.length;
      else if (e.key === 'Home') i = 0;
      else if (e.key === 'End') i = vis.length - 1;
      else if (e.key === 'ArrowDown') { e.preventDefault(); const f = firstFocusable(vis[i].panel); if (f) f.focus(); return; }
      else return;
      e.preventDefault();
      selectTab(vis[i].id, true);
    });
    function focusables(el) {
      return Array.from(el.querySelectorAll('button,select,input')).filter(x => !x.disabled && x.offsetParent !== null);
    }
    function firstFocusable(el) { return focusables(el)[0]; }
    function toolbarArrows(e) {
      if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
      if (e.target.tagName === 'SELECT' || e.target.tagName === 'INPUT' && e.target.type !== 'color' && e.target.type !== 'checkbox') return;
      const items = focusables(e.currentTarget);
      const i = items.indexOf(document.activeElement);
      if (i < 0) return;
      e.preventDefault();
      items[(i + (e.key === 'ArrowRight' ? 1 : items.length - 1)) % items.length].focus();
    }

    /* FILE */
    addTab('file', 'File', [
      group('Save',
        btn({ icon: '\u{1F4BE}', label: 'Save', keys: 'Ctrl+S', big: true, errPrefix: 'Save failed: ', run: () => editor.save() }),
        btn({ icon: '\u{1F4C4}', label: 'Save As', keys: 'Ctrl+Shift+S', big: true, errPrefix: 'Save failed: ', run: () => editor.saveAs() })),
      group('Export',
        btn({ icon: '\u{1F5A8}', label: 'Print / Save as PDF', keys: 'Ctrl+P', big: true, run: () => printDoc() }),
        btn({ icon: '⬇', label: 'Download copy', big: true, run: () => downloadCopy() })),
      group('Info',
        btn({ icon: '#', label: 'Word count', big: true, run: () => showWordCount() }))
    ]);

    /* HOME */
    const fontSel = sel('Font', FONTS.map(f => ({ value: f, label: f })), v => { cmd2('fontName', v, true); });
    fontSel.classList.add('word-font-select');
    Array.from(fontSel.options).forEach(o => { o.style.fontFamily = o.value; });
    const sizeSel = sel('Font size', SIZES.map(s => ({ value: s, label: String(s) })), v => setFontSize(+v));
    sizeSel.value = '11';
    const spacingSel = sel('Line spacing', [{ value: '', label: 'Spacing' }, { value: '1', label: '1.0' }, { value: '1.15', label: '1.15' }, { value: '1.5', label: '1.5' }, { value: '2', label: '2.0' }], v => { if (v) setLineSpacing(v); });
    const styleBtns = {};
    const styleDefs = [['p', 'Normal', 'Normal paragraph'], ['title', 'Title', 'Title'], ['h1', 'Heading 1', 'Heading 1'], ['h2', 'Heading 2', 'Heading 2'], ['h3', 'Heading 3', 'Heading 3'], ['blockquote', 'Quote', 'Quote']];
    const styleGroupItems = styleDefs.map(([k, label, t]) => {
      const b = btn({ label, title: 'Style: ' + t, cls: 'word-style-btn word-style-' + k, pressable: true, run: () => applyStyle(k) });
      styleBtns[k] = b;
      return b;
    });
    addTab('home', 'Home', [
      group('Undo',
        btn({ icon: '↶', label: 'Undo', keys: 'Ctrl+Z', iconOnly: true, run: () => cmd2('undo') }),
        btn({ icon: '↷', label: 'Redo', keys: 'Ctrl+Y', iconOnly: true, run: () => cmd2('redo') })),
      group('Clipboard',
        btn({ icon: '✂', label: 'Cut', keys: 'Ctrl+X', iconOnly: true, run: () => clip('cut') }),
        btn({ icon: '⎘', label: 'Copy', keys: 'Ctrl+C', iconOnly: true, run: () => clip('copy') }),
        btn({ icon: '\u{1F4CB}', label: 'Paste as plain text', keys: 'Ctrl+Shift+V', iconOnly: true, run: () => pastePlain() })),
      group('Font',
        h('div', { class: 'tool-row' }, fontSel, sizeSel,
          btn({ icon: 'A+', label: 'Grow font', iconOnly: true, run: () => stepFont(1) }),
          btn({ icon: 'A-', label: 'Shrink font', iconOnly: true, run: () => stepFont(-1) })),
        h('div', { class: 'tool-row' },
          btn({ icon: 'B', label: 'Bold', keys: 'Ctrl+B', iconOnly: true, toggle: 'bold', iconStyle: 'font-weight:700', run: () => cmd2('bold') }),
          btn({ icon: 'I', label: 'Italic', keys: 'Ctrl+I', iconOnly: true, toggle: 'italic', iconStyle: 'font-style:italic;font-family:serif', run: () => cmd2('italic') }),
          btn({ icon: 'U', label: 'Underline', keys: 'Ctrl+U', iconOnly: true, toggle: 'underline', iconStyle: 'text-decoration:underline', run: () => cmd2('underline') }),
          btn({ icon: 'S', label: 'Strikethrough', iconOnly: true, toggle: 'strikeThrough', iconStyle: 'text-decoration:line-through', run: () => cmd2('strikeThrough') }),
          btn({ icon: 'x²', label: 'Superscript', iconOnly: true, toggle: 'superscript', run: () => cmd2('superscript') }),
          btn({ icon: 'x₂', label: 'Subscript', iconOnly: true, toggle: 'subscript', run: () => cmd2('subscript') }),
          colorPair('Text color', 'foreColor', '#c00000', 'A'),
          colorPair('Highlight color', 'hiliteColor', '#ffff00', '\u{1F58D}'),
          btn({ icon: '⨯', label: 'Clear formatting', iconOnly: true, run: () => clearFormatting() }))),
      group('Paragraph',
        h('div', { class: 'tool-row' },
          btn({ icon: '•≡', label: 'Bullets', iconOnly: true, toggle: 'insertUnorderedList', run: () => cmd2('insertUnorderedList') }),
          btn({ icon: '1≡', label: 'Numbering', iconOnly: true, toggle: 'insertOrderedList', run: () => cmd2('insertOrderedList') }),
          btn({ icon: '⇤', label: 'Decrease indent', keys: 'Shift+Tab in lists', iconOnly: true, run: () => indent(-1) }),
          btn({ icon: '⇥', label: 'Increase indent', keys: 'Tab in lists', iconOnly: true, run: () => indent(1) }),
          spacingSel),
        h('div', { class: 'tool-row' },
          btn({ icon: ' ', label: 'Align left', keys: 'Ctrl+L', iconOnly: true, toggle: 'justifyLeft', cls: 'word-al word-al-left', run: () => cmd2('justifyLeft', null, true) }),
          btn({ icon: ' ', label: 'Center', keys: 'Ctrl+E', iconOnly: true, toggle: 'justifyCenter', cls: 'word-al word-al-center', run: () => cmd2('justifyCenter', null, true) }),
          btn({ icon: ' ', label: 'Align right', keys: 'Ctrl+R', iconOnly: true, toggle: 'justifyRight', cls: 'word-al word-al-right', run: () => cmd2('justifyRight', null, true) }),
          btn({ icon: ' ', label: 'Justify', keys: 'Ctrl+J', iconOnly: true, toggle: 'justifyFull', cls: 'word-al word-al-justify', run: () => cmd2('justifyFull', null, true) }))),
      group('Styles', h('div', { class: 'word-styles' }, styleGroupItems)),
      group('Editing',
        btn({ icon: '\u{1F50D}', label: 'Find', keys: 'Ctrl+F', run: () => openFind(false) }),
        btn({ icon: '⇄', label: 'Replace', keys: 'Ctrl+H', run: () => openFind(true) }),
        btn({ icon: '▣', label: 'Select all', keys: 'Ctrl+A', run: () => cmd2('selectAll') }))
    ]);

    /* INSERT */
    addTab('insert', 'Insert', [
      group('Tables', btn({ icon: '▦', label: 'Table', big: true, run: () => insertTableDialog() })),
      group('Illustrations',
        btn({ icon: '\u{1F5BC}', label: 'Picture', title: 'Insert picture from a file', big: true, run: () => insertPicture() }),
        btn({ icon: '\u{1F4C1}', label: 'Picture (browse device)', title: 'Insert picture using the system file chooser', big: true, run: () => insertPictureNative() })),
      group('Links',
        btn({ icon: '\u{1F517}', label: 'Link', keys: 'Ctrl+K', big: true, run: () => insertLink() }),
        btn({ icon: '⛓', label: 'Remove link', big: true, run: () => cmd2('unlink') })),
      group('Breaks',
        btn({ icon: '―', label: 'Horizontal line', big: true, run: () => cmd2('insertHorizontalRule') }),
        btn({ icon: '⎘', label: 'Page break', keys: 'Ctrl+Enter', big: true, run: () => insertPageBreak() })),
      group('Text',
        btn({ icon: '\u{1F4C5}', label: 'Date & time', big: true, run: () => insertDate() })),
      group('Symbols',
        btn({ icon: 'Ω', label: 'Symbol', big: true, run: () => charDialog('Insert symbol', SYMBOLS) }),
        btn({ icon: '§', label: 'Special character', big: true, run: () => charDialog('Insert special character', SPECIAL_CHARS) }))
    ]);

    /* LAYOUT */
    const marginSel = sel('Margins', [{ value: 'normal', label: 'Normal (1")' }, { value: 'narrow', label: 'Narrow (0.5")' }, { value: 'wide', label: 'Wide (2" sides)' }], v => { layout.margin = v; applyLayout(); });
    const orientSel = sel('Orientation', [{ value: 'portrait', label: 'Portrait' }, { value: 'landscape', label: 'Landscape' }], v => { layout.orient = v; applyLayout(); });
    const sizePageSel = sel('Page size', [{ value: 'letter', label: 'Letter (8.5" × 11")' }, { value: 'a4', label: 'A4 (210 × 297 mm)' }], v => { layout.size = v; applyLayout(); });
    addTab('layout', 'Layout', [
      group('Page setup',
        h('label', { class: 'word-field' }, h('span', { text: 'Margins' }), marginSel),
        h('label', { class: 'word-field' }, h('span', { text: 'Orientation' }), orientSel),
        h('label', { class: 'word-field' }, h('span', { text: 'Size' }), sizePageSel)),
      group('Paragraph',
        btn({ icon: '⇤', label: 'Decrease indent', big: true, run: () => indent(-1) }),
        btn({ icon: '⇥', label: 'Increase indent', big: true, run: () => indent(1) }))
    ]);
    [marginSel, orientSel, sizePageSel].forEach(s => s.removeAttribute('aria-label'));

    /* VIEW */
    const zoomSel = sel('Zoom', ZOOMS.map(z => ({ value: z, label: z + '%' })), v => setZoom(+v), { keys: 'Ctrl+= / Ctrl+-' });
    zoomSel.value = '100';
    const readBtn = btn({ icon: '\u{1F4D6}', label: 'Read mode', big: true, pressable: true, run: () => setReadMode(!readMode) });
    const darkBtn = btn({ icon: '\u{1F319}', label: 'Dark page', title: 'Dark page (focus mode)', big: true, pressable: true, run: () => { const on = !root.classList.contains('word-dark-page'); root.classList.toggle('word-dark-page', on); darkBtn.setAttribute('aria-pressed', String(on)); } });
    addTab('view', 'View', [
      group('Views', readBtn, darkBtn),
      group('Zoom',
        btn({ icon: '−', label: 'Zoom out', keys: 'Ctrl+-', iconOnly: true, run: () => stepZoom(-1) }),
        zoomSel,
        btn({ icon: '+', label: 'Zoom in', keys: 'Ctrl+=', iconOnly: true, run: () => stepZoom(1) }),
        btn({ icon: '1:1', label: '100%', iconOnly: false, run: () => setZoom(100) })),
      group('Show', btn({ icon: '#', label: 'Word count', big: true, run: () => showWordCount() }))
    ]);

    /* contextual TABLE */
    addTab('table', 'Table', [
      group('Rows',
        btn({ icon: '⤒', label: 'Insert row above', big: true, run: () => tableOp('rowAbove') }),
        btn({ icon: '⤓', label: 'Insert row below', big: true, run: () => tableOp('rowBelow') })),
      group('Columns',
        btn({ icon: '⇤', label: 'Insert column left', big: true, run: () => tableOp('colLeft') }),
        btn({ icon: '⇥', label: 'Insert column right', big: true, run: () => tableOp('colRight') })),
      group('Delete',
        btn({ icon: '✖', label: 'Delete row', big: true, run: () => tableOp('delRow') }),
        btn({ icon: '✖', label: 'Delete column', big: true, run: () => tableOp('delCol') }),
        btn({ icon: '✖', label: 'Delete table', big: true, run: () => tableOp('delTable') }))
    ], true);

    /* contextual PICTURE */
    const imgWidthSel = sel('Picture width', [{ value: '', label: 'Original size' }, { value: '25%', label: '25% of page' }, { value: '50%', label: '50% of page' }, { value: '75%', label: '75% of page' }, { value: '100%', label: 'Full width' }], v => {
      if (!selectedImg) return;
      selectedImg.style.width = v;
      selectedImg.style.height = v ? 'auto' : '';
      selectedImg.removeAttribute('width'); selectedImg.removeAttribute('height');
      markDirty();
    });
    addTab('picture', 'Picture', [
      group('Size', h('label', { class: 'word-field' }, h('span', { text: 'Width' }), imgWidthSel)),
      group('Accessibility', btn({ icon: 'Alt', label: 'Alt text', big: true, run: () => editAltText() })),
      group('Arrange',
        btn({ icon: ' ', label: 'Align left', big: true, cls: 'word-al word-al-left', run: () => alignImage('left') }),
        btn({ icon: ' ', label: 'Center', big: true, cls: 'word-al word-al-center', run: () => alignImage('center') }),
        btn({ icon: ' ', label: 'Align right', big: true, cls: 'word-al word-al-right', run: () => alignImage('right') }),
        btn({ icon: '✖', label: 'Delete picture', big: true, run: () => { if (selectedImg) { selectedImg.remove(); selectImage(null); markDirty(); } } }))
    ], true);
    imgWidthSel.removeAttribute('aria-label');
    selectTab('home');

    /* ---- find panel ---- */
    const findInput = h('input', { type: 'search', class: 'word-input', id: uid + '-find' });
    const replInput = h('input', { type: 'text', class: 'word-input', id: uid + '-repl' });
    const caseBox = h('input', { type: 'checkbox', id: uid + '-case' });
    const wordBox = h('input', { type: 'checkbox', id: uid + '-whole' });
    const findStatus = h('p', { class: 'word-find-status', role: 'status', 'aria-live': 'polite' });
    const replaceRow = h('div', { class: 'word-find-replace' },
      h('label', { for: uid + '-repl', text: 'Replace with' }), replInput,
      h('div', { class: 'word-find-buttons' },
        btn({ label: 'Replace', run: () => replaceOne() }),
        btn({ label: 'Replace all', run: () => replaceAll() })));
    const findPanel = h('aside', { class: 'panel-side right word-find', 'aria-label': 'Find and replace', hidden: true },
      h('div', { class: 'word-find-head' },
        h('h2', { text: 'Find and replace' }),
        btn({ icon: '✕', label: 'Close find panel', keys: 'Esc', iconOnly: true, run: () => closeFind() })),
      h('label', { for: uid + '-find', text: 'Find' }), findInput,
      h('div', { class: 'word-find-opts' },
        h('label', null, caseBox, ' Match case'),
        h('label', null, wordBox, ' Whole words only')),
      h('div', { class: 'word-find-buttons' },
        btn({ icon: '↑', label: 'Previous', keys: 'Shift+Enter', run: () => findStep(-1) }),
        btn({ icon: '↓', label: 'Next', keys: 'Enter', run: () => findStep(1) })),
      findStatus, replaceRow);

    /* ---- status bar ---- */
    const stPage = h('span', { class: 'word-st' }), stWords = h('span', { class: 'word-st' }), stChars = h('span', { class: 'word-st' });
    const stZoom = h('span', { class: 'word-st' }, 'Zoom 100%'), stSave = h('span', { class: 'word-st word-st-save' });
    const statusbar = h('div', { class: 'statusbar word-statusbar', role: 'group', 'aria-label': 'Document status' },
      stPage, stWords, stChars, h('span', { class: 'word-st-spacer' }), stSave, stZoom);

    root.append(ribbon, h('div', { class: 'word-main' }, canvas, findPanel), statusbar);
    container.appendChild(root);
    body.innerHTML = '<p>Loading…</p>';
    applyLayout();

    /* ---- selection handling ---- */
    function inBody(node) { return node && (node === body || body.contains(node)); }
    function restoreSel() {
      const s = window.getSelection();
      if (document.activeElement !== body) body.focus({ preventScroll: true });
      if (savedRange && inBody(savedRange.startContainer)) {
        s.removeAllRanges();
        s.addRange(savedRange);
      } else if (!s.rangeCount || !inBody(s.anchorNode)) {
        const r = document.createRange();
        r.selectNodeContents(body);
        r.collapse(false);
        s.removeAllRanges();
        s.addRange(r);
      }
    }
    function currentRange() {
      const s = window.getSelection();
      if (s.rangeCount && inBody(s.anchorNode)) return s.getRangeAt(0);
      return savedRange && inBody(savedRange.startContainer) ? savedRange : null;
    }
    function exec(name, val) { return document.execCommand(name, false, val); }
    function cmd2(name, value, css) {
      if (loading || readMode) return;
      restoreSel();
      exec('styleWithCSS', !!css);
      exec(name, value == null ? null : value);
      afterEdit();
    }
    function afterEdit() { markDirty(); saveSel(); scheduleUI(); }
    function saveSel() {
      const s = window.getSelection();
      if (s.rangeCount && inBody(s.anchorNode)) savedRange = s.getRangeAt(0).cloneRange();
    }
    on(document, 'selectionchange', () => {
      const s = window.getSelection();
      if (s.rangeCount && inBody(s.anchorNode)) { savedRange = s.getRangeAt(0).cloneRange(); scheduleUI(); }
    });

    function closestIn(node, selector) {
      let el = node && (node.nodeType === 1 ? node : node.parentElement);
      el = el && el.closest(selector);
      return el && body.contains(el) ? el : null;
    }
    function selectedBlocks() {
      const r = currentRange();
      if (!r) return [];
      const all = Array.from(body.querySelectorAll(BLOCK_SEL)).filter(b => r.intersectsNode(b));
      let blocks = all.filter(b => !all.some(o => o !== b && b.contains(o)));
      if (!blocks.length) {
        const c = closestIn(r.startContainer, BLOCK_SEL);
        if (c) blocks = [c];
        else { exec('formatBlock', '<p>'); const c2 = closestIn(window.getSelection().anchorNode, BLOCK_SEL); if (c2) blocks = [c2]; }
      }
      return blocks;
    }

    /* ---- formatting commands ---- */
    function setFontSize(pt) {
      if (loading || readMode) return;
      restoreSel();
      const s = window.getSelection();
      if (s.isCollapsed) {
        const r = s.getRangeAt(0);
        const span = document.createElement('span');
        span.style.fontSize = pt + 'pt';
        const zw = document.createTextNode('​');
        span.appendChild(zw);
        r.insertNode(span);
        const c = document.createRange();
        c.setStart(zw, 1); c.collapse(true);
        s.removeAllRanges(); s.addRange(c);
        afterEdit();
        return;
      }
      exec('styleWithCSS', false);
      exec('fontSize', '7');
      const fonts = Array.from(body.querySelectorAll('font[size="7"]'));
      let first = null, last = null;
      fonts.forEach(f => {
        const span = document.createElement('span');
        span.style.fontSize = pt + 'pt';
        if (f.getAttribute('color')) span.style.color = f.getAttribute('color');
        if (f.getAttribute('face')) span.style.fontFamily = f.getAttribute('face');
        while (f.firstChild) span.appendChild(f.firstChild);
        span.querySelectorAll('[style]').forEach(x => { x.style.fontSize = ''; if (!x.getAttribute('style')) x.removeAttribute('style'); });
        f.replaceWith(span);
        first = first || span; last = span;
      });
      // Chrome may emit spans with xxx-large when styleWithCSS sticks
      body.querySelectorAll('span[style*="xxx-large"]').forEach(sp => { sp.style.fontSize = pt + 'pt'; first = first || sp; last = sp; });
      if (first) {
        const r = document.createRange();
        r.setStartBefore(first); r.setEndAfter(last);
        s.removeAllRanges(); s.addRange(r);
      }
      afterEdit();
    }
    function currentFontPt() {
      const r = currentRange();
      const el = r && (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement);
      if (!el) return 11;
      return Math.round(parseFloat(getComputedStyle(el).fontSize) * 0.75 * 2) / 2;
    }
    function stepFont(dir) {
      const cur = currentFontPt();
      let next = dir > 0 ? SIZES.find(s => s > cur) : SIZES.slice().reverse().find(s => s < cur);
      if (next == null) next = dir > 0 ? SIZES[SIZES.length - 1] : SIZES[0];
      setFontSize(next);
    }
    function setLineSpacing(v) {
      if (loading || readMode) return;
      restoreSel();
      selectedBlocks().forEach(b => { b.style.lineHeight = v; });
      afterEdit();
    }
    function indent(dir) {
      if (loading || readMode) return;
      restoreSel();
      const r = currentRange();
      if (r && closestIn(r.startContainer, 'li')) { exec('styleWithCSS', false); exec(dir > 0 ? 'indent' : 'outdent'); afterEdit(); return; }
      selectedBlocks().forEach(b => {
        const cur = parseFloat(b.style.marginLeft) || 0;
        const next = Math.max(0, cur + dir * 0.5);
        b.style.marginLeft = next ? next + 'in' : '';
        if (!b.getAttribute('style')) b.removeAttribute('style');
      });
      afterEdit();
    }
    function applyStyle(k) {
      if (loading || readMode) return;
      restoreSel();
      exec('styleWithCSS', false);
      if (k === 'title') {
        exec('formatBlock', '<h1>');
        selectedBlocks().forEach(b => { if (b.tagName === 'H1') b.classList.add('title'); });
      } else {
        exec('formatBlock', '<' + k + '>');
        if (k === 'h1') selectedBlocks().forEach(b => { b.classList.remove('title'); if (!b.className) b.removeAttribute('class'); });
      }
      afterEdit();
    }
    function clearFormatting() {
      if (loading || readMode) return;
      restoreSel();
      exec('removeFormat');
      selectedBlocks().forEach(b => { b.removeAttribute('style'); b.removeAttribute('class'); });
      afterEdit();
    }
    async function clip(kind) {
      restoreSel();
      if (!exec(kind)) {
        App.toast('Use Ctrl+' + (kind === 'cut' ? 'X' : 'C') + ' to ' + kind + ' (blocked by the browser).', 'info');
        return;
      }
      if (kind === 'cut') afterEdit();
    }
    async function pastePlain() {
      let text = null;
      try { text = await navigator.clipboard.readText(); } catch (e) { text = null; }
      if (text == null) { App.toast('Clipboard access was blocked. Press Ctrl+Shift+V in the document to paste as plain text.', 'info'); return; }
      restoreSel();
      exec('insertText', text);
      afterEdit();
    }

    /* ---- insert ---- */
    function insertHtmlAtCaret(html) {
      if (loading || readMode) return;
      restoreSel();
      exec('styleWithCSS', false);
      exec('insertHTML', html);
      afterEdit();
    }
    async function insertTableDialog() {
      const rows = h('input', { type: 'number', name: 'rows', min: '1', max: '100', value: '3', required: true, id: uid + '-trows' });
      const cols = h('input', { type: 'number', name: 'cols', min: '1', max: '20', value: '3', required: true, id: uid + '-tcols' });
      const res = await openDialog('Insert table', [
        h('div', { class: 'word-dialog-grid' },
          h('label', { for: uid + '-trows', text: 'Rows' }), rows,
          h('label', { for: uid + '-tcols', text: 'Columns' }), cols),
        h('label', { class: 'word-check' }, h('input', { type: 'checkbox', name: 'header', checked: true }), ' First row is a header row')
      ], 'Insert table');
      if (!res || res.value !== 'ok') return;
      const R = Math.min(100, Math.max(1, parseInt(res.data.get('rows'), 10) || 1));
      const C = Math.min(20, Math.max(1, parseInt(res.data.get('cols'), 10) || 1));
      const header = res.data.get('header') != null;
      let html = '<table><tbody>';
      for (let r = 0; r < R; r++) {
        html += '<tr>';
        for (let c = 0; c < C; c++) html += (header && r === 0 ? '<th><br></th>' : '<td><br></td>');
        html += '</tr>';
      }
      html += '</tbody></table><p><br></p>';
      insertHtmlAtCaret(html);
      App.toast('Table inserted with ' + R + ' rows and ' + C + ' columns. Use Tab to move between cells.', 'success');
    }
    function currentCell() {
      const r = currentRange();
      return r ? closestIn(r.startContainer, 'td,th') : null;
    }
    function caretInto(el) {
      const r = document.createRange();
      r.selectNodeContents(el);
      r.collapse(true);
      const s = window.getSelection();
      body.focus({ preventScroll: true });
      s.removeAllRanges(); s.addRange(r);
      savedRange = r.cloneRange();
    }
    function newCell(tag) { const c = document.createElement(tag); c.innerHTML = '<br>'; return c; }
    function tableOp(op) {
      if (readMode) return;
      const cell = currentCell();
      if (!cell) { App.toast('Place the cursor in a table first.', 'info'); return; }
      const row = cell.parentElement, table = cell.closest('table');
      const idx = cell.cellIndex;
      let focusCell = cell;
      switch (op) {
        case 'rowAbove': case 'rowBelow': {
          const nr = document.createElement('tr');
          Array.from(row.cells).forEach(() => nr.appendChild(newCell('td')));
          row.parentElement.insertBefore(nr, op === 'rowAbove' ? row : row.nextSibling);
          focusCell = nr.cells[Math.min(idx, nr.cells.length - 1)];
          break;
        }
        case 'colLeft': case 'colRight': {
          Array.from(table.rows).forEach(tr => {
            const ref = tr.cells[idx];
            const c = newCell(ref && ref.tagName === 'TH' ? 'th' : 'td');
            if (!ref) tr.appendChild(c);
            else tr.insertBefore(c, op === 'colLeft' ? ref : ref.nextSibling);
            if (tr === row) focusCell = c;
          });
          break;
        }
        case 'delRow': {
          const rowsArr = Array.from(table.rows);
          if (rowsArr.length <= 1) return tableOp('delTable');
          const i = rowsArr.indexOf(row);
          row.remove();
          const rest = Array.from(table.rows);
          focusCell = rest[Math.min(i, rest.length - 1)].cells[0];
          break;
        }
        case 'delCol': {
          Array.from(table.rows).forEach(tr => { if (tr.cells[idx]) tr.cells[idx].remove(); });
          if (!table.querySelector('td,th')) return tableOp('delTable');
          focusCell = row.cells[Math.min(idx, row.cells.length - 1)] || table.querySelector('td,th');
          break;
        }
        case 'delTable': {
          const p = h('p', null, h('br'));
          table.replaceWith(p);
          focusCell = p;
          break;
        }
      }
      if (focusCell) caretInto(focusCell);
      markDirty();
      scheduleUI();
    }
    async function insertImageBytes(bytes, name) {
      const url = await bytesToDataUrl(bytes, imageMime(name));
      const alt = await App.prompt('Describe the picture for people who cannot see it (alt text):', stripExt(baseName(name)));
      if (alt === null) return;
      insertHtmlAtCaret('<img src="' + url + '" alt="' + escapeHtml(alt) + '" style="max-width:100%">');
    }
    async function insertPicture() {
      saveSel();
      const p = await App.pickFile({ mode: 'open', extensions: ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp'] });
      if (!p) return;
      const path = Array.isArray(p) ? p[0] : p;
      const bytes = await App.readFile(path);
      await insertImageBytes(bytes, path);
    }
    function insertPictureNative() {
      saveSel();
      const input = h('input', { type: 'file', accept: 'image/*', hidden: true });
      input.addEventListener('change', async () => {
        const f = input.files && input.files[0];
        input.remove();
        if (f) {
          try { await insertImageBytes(new Uint8Array(await f.arrayBuffer()), f.name); } catch (e) { App.toast(e.message, 'error'); }
        }
      });
      root.appendChild(input);
      input.click();
    }
    async function insertLink() {
      if (loading || readMode) return;
      saveSel();
      const r = currentRange();
      const existing = r && closestIn(r.startContainer, 'a');
      const url = await App.prompt('Link address (URL):', existing ? existing.getAttribute('href') : 'https://');
      if (!url || !url.trim() || url.trim() === 'https://') return;
      const href = safeUrl(url.trim());
      if (existing) { existing.setAttribute('href', href); markDirty(); return; }
      restoreSel();
      if (window.getSelection().isCollapsed) exec('insertHTML', '<a href="' + escapeHtml(href) + '">' + escapeHtml(url.trim()) + '</a>');
      else exec('createLink', href);
      afterEdit();
    }
    function insertPageBreak() { insertHtmlAtCaret('<hr class="page-break"><p><br></p>'); }
    function insertDate() {
      if (loading || readMode) return;
      restoreSel();
      exec('insertText', new Date().toLocaleString(undefined, { dateStyle: 'long', timeStyle: 'short' }));
      afterEdit();
    }
    async function charDialog(title, chars) {
      saveSel();
      const cols = 8;
      const grid = h('div', { class: 'word-char-grid', role: 'group', 'aria-label': 'Characters' },
        chars.map(c => h('button', {
          type: 'submit', value: 'ch:' + c, class: 'tool-btn word-char',
          'aria-label': CHAR_NAMES[c] || ('Character ' + c + ' U+' + c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')),
          title: (CHAR_NAMES[c] || c) + ' (U+' + c.codePointAt(0).toString(16).toUpperCase().padStart(4, '0') + ')'
        }, c === ' ' ? 'nbsp' : c === '­' ? 'shy' : c)));
      grid.addEventListener('keydown', e => {
        const items = Array.from(grid.children);
        const i = items.indexOf(document.activeElement);
        if (i < 0) return;
        const d = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: cols, ArrowUp: -cols }[e.key];
        if (!d) return;
        e.preventDefault();
        const n = i + d;
        if (n >= 0 && n < items.length) items[n].focus();
      });
      const res = await openDialog(title, [h('p', { class: 'word-hint', text: 'Use arrow keys to move, Enter to insert.' }), grid], null);
      if (!res || !res.value.startsWith('ch:')) return;
      restoreSel();
      exec('insertText', res.value.slice(3));
      afterEdit();
    }

    /* pictures: selection + contextual tab */
    function selectImage(img) {
      if (selectedImg) selectedImg.classList.remove('word-img-selected');
      selectedImg = img;
      if (img) {
        img.classList.add('word-img-selected');
        const w = img.style.width;
        imgWidthSel.value = ['25%', '50%', '75%', '100%'].includes(w) ? w : '';
      }
      setContextTab('picture', !!img);
    }
    async function editAltText() {
      if (!selectedImg) return;
      const img = selectedImg;
      const alt = await App.prompt('Alt text (describe the picture):', img.getAttribute('alt') || '');
      if (alt === null) return;
      img.setAttribute('alt', alt);
      markDirty();
    }
    function alignImage(where) {
      if (!selectedImg) return;
      const img = selectedImg;
      const r = document.createRange();
      r.selectNode(img);
      savedRange = r;
      cmd2(where === 'left' ? 'justifyLeft' : where === 'center' ? 'justifyCenter' : 'justifyRight', null, true);
    }
    function setContextTab(id, show) {
      const t = tabs.find(x => x.id === id);
      if (!t || t.tab.hidden === !show) return;
      t.tab.hidden = !show;
      const act = activeTab();
      if (!show && act && act.id === id) selectTab('home');
    }

    /* ---- layout / view ---- */
    function applyLayout() {
      const ps = PAGE_SIZES[layout.size], mg = MARGINS[layout.margin];
      const land = layout.orient === 'landscape';
      page.style.setProperty('--page-w', (land ? ps.h : ps.w) + 'in');
      page.style.setProperty('--page-h', (land ? ps.w : ps.h) + 'in');
      page.style.setProperty('--pad-y', mg.y + 'in');
      page.style.setProperty('--pad-x', mg.x + 'in');
      scheduleStatus();
    }
    function setZoom(z) {
      zoom = Math.max(50, Math.min(200, z));
      page.style.zoom = zoom / 100;
      const nearest = ZOOMS.includes(zoom) ? zoom : null;
      zoomSel.value = nearest ? String(nearest) : zoomSel.value;
      stZoom.textContent = 'Zoom ' + zoom + '%';
    }
    function stepZoom(dir) {
      const next = dir > 0 ? ZOOMS.find(z => z > zoom) : ZOOMS.slice().reverse().find(z => z < zoom);
      if (next) setZoom(next);
    }
    function setReadMode(onState) {
      readMode = onState;
      root.classList.toggle('word-read-mode', onState);
      ribbon.hidden = onState;
      exitRead.hidden = !onState;
      body.setAttribute('contenteditable', onState || loading ? 'false' : 'true');
      body.setAttribute('aria-readonly', String(onState));
      readBtn.setAttribute('aria-pressed', String(onState));
      if (onState) { closeFind(); exitRead.focus(); App.toast('Read mode. Press Esc to return to editing.', 'info'); }
      else restoreSel();
    }

    /* ---- stats / status bar ---- */
    function stats() {
      const text = body.innerText || '';
      const words = (text.match(/[^\s​]+/g) || []).length;
      const chars = text.replace(/[\n​]/g, '').length;
      const charsNoSpace = text.replace(/[\s​]/g, '').length;
      const paras = Array.from(body.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,blockquote,pre')).filter(p => p.textContent.trim()).length;
      const ps = PAGE_SIZES[layout.size], mg = MARGINS[layout.margin];
      const land = layout.orient === 'landscape';
      const contentPx = ((land ? ps.w : ps.h) - 2 * mg.y) * 96;
      const breaks = body.querySelectorAll('hr.page-break').length;
      const pages = Math.max(1, Math.ceil(body.scrollHeight / contentPx)) + breaks;
      return { words, chars, charsNoSpace, paras, pages, contentPx };
    }
    function updateStatus() {
      const s = stats();
      let cur = 1;
      const r = currentRange();
      if (r) {
        let rect = r.getBoundingClientRect();
        if (!rect || (!rect.top && !rect.height)) { const el = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement; rect = el.getBoundingClientRect(); }
        const top = (rect.top - body.getBoundingClientRect().top) / (zoom / 100);
        let before = 0;
        body.querySelectorAll('hr.page-break').forEach(hr => { if (hr.compareDocumentPosition(r.startContainer) & Node.DOCUMENT_POSITION_FOLLOWING) before++; });
        cur = Math.min(s.pages, Math.max(1, Math.floor(top / s.contentPx) + 1 + before));
      }
      stPage.textContent = 'Page ' + cur + ' of ' + s.pages;
      stWords.textContent = s.words + (s.words === 1 ? ' word' : ' words');
      stChars.textContent = s.chars + (s.chars === 1 ? ' character' : ' characters');
      stSave.textContent = loading ? 'Loading…' : dirty ? 'Unsaved changes' : docPath ? 'All changes saved' : 'Not saved yet';
    }
    function scheduleStatus() { clearTimeout(statusTimer); statusTimer = setTimeout(updateStatus, 150); }
    async function showWordCount() {
      const s = stats();
      const dl = h('dl', { class: 'word-stats' },
        [['Pages (estimate)', s.pages], ['Words', s.words], ['Characters (no spaces)', s.charsNoSpace], ['Characters (with spaces)', s.chars], ['Paragraphs', s.paras]]
          .map(([k, v]) => [h('dt', { text: k }), h('dd', { text: String(v) })]));
      await openDialog('Word count', [dl], null, 'Close');
    }

    /* ---- ribbon state ---- */
    function scheduleUI() {
      if (uiTimer) return;
      uiTimer = requestAnimationFrame(() => { uiTimer = 0; updateUI(); });
      scheduleStatus();
    }
    function updateUI() {
      const r = currentRange();
      if (!r || loading) return;
      for (const t of toggles) {
        let st = false;
        try { st = document.queryCommandState(t.cmd); } catch (e) { st = false; }
        t.btn.setAttribute('aria-pressed', String(!!st));
      }
      let fn = '';
      try { fn = (document.queryCommandValue('fontName') || '').replace(/["']/g, '').split(',')[0].trim(); } catch (e) { fn = ''; }
      const match = FONTS.find(f => f.toLowerCase() === fn.toLowerCase());
      if (match) fontSel.value = match;
      else if (fn) {
        let o = fontSel.querySelector('option[data-extra]');
        if (!o) { o = h('option', { 'data-extra': '1' }); fontSel.appendChild(o); }
        o.value = fn; o.textContent = fn; fontSel.value = fn;
      }
      const pt = currentFontPt();
      let so = Array.from(sizeSel.options).find(o => +o.value === pt);
      if (!so) {
        so = sizeSel.querySelector('option[data-extra]');
        if (!so) { so = h('option', { 'data-extra': '1' }); sizeSel.appendChild(so); }
        so.value = String(pt); so.textContent = String(pt);
      }
      sizeSel.value = String(pt);
      const block = closestIn(r.startContainer, 'p,h1,h2,h3,h4,h5,h6,blockquote,pre,li');
      let key = block ? block.tagName.toLowerCase() : 'p';
      if (key === 'h1' && block.classList.contains('title')) key = 'title';
      if (block && block.closest('blockquote') && body.contains(block.closest('blockquote'))) key = 'blockquote';
      for (const k in styleBtns) styleBtns[k].setAttribute('aria-pressed', String(k === key));
      const lh = block && block.style.lineHeight;
      spacingSel.value = ['1', '1.15', '1.5', '2'].includes(lh) ? lh : '';
      setContextTab('table', !!closestIn(r.startContainer, 'table'));
      const img = r.startContainer.nodeType === 1 && r.startContainer.childNodes[r.startOffset] && r.startContainer.childNodes[r.startOffset].tagName === 'IMG' && !r.collapsed ? r.startContainer.childNodes[r.startOffset] : null;
      if (img !== selectedImg && !(selectedImg && document.activeElement !== body)) selectImage(img);
    }

    /* ---- dirty ---- */
    function markDirty() {
      if (loading) return;
      if (!dirty) { dirty = true; App.setDirty(editor, true); }
      scheduleStatus();
      if (!findPanel.hidden && findInput.value) scheduleFindRefresh();
    }
    function setClean() { dirty = false; App.setDirty(editor, false); scheduleStatus(); }

    /* ---- dialogs ---- */
    function openDialog(title, content, okLabel, cancelLabel) {
      return new Promise(resolve => {
        const titleId = uid + '-dlg-' + Date.now();
        const form = h('form', { method: 'dialog', class: 'word-dialog-form' }, content,
          h('div', { class: 'word-dialog-actions' },
            okLabel ? h('button', { type: 'submit', value: 'ok', class: 'tool-btn word-primary' }, okLabel) : null,
            h('button', { type: 'button', class: 'tool-btn', onclick: () => dlg.close('cancel') }, cancelLabel || 'Cancel')));
        const dlg = h('dialog', { class: 'word-dialog', 'aria-labelledby': titleId }, h('h2', { id: titleId, text: title }), form);
        const prevFocus = document.activeElement;
        dlg.addEventListener('close', () => {
          const data = new FormData(form);
          const value = dlg.returnValue || '';
          dlg.remove();
          if (prevFocus && prevFocus.focus && document.contains(prevFocus)) prevFocus.focus();
          resolve({ value, data });
        });
        root.appendChild(dlg);
        dlg.showModal();
        const f = dlg.querySelector('input,.word-char,button[type=submit]');
        if (f) f.focus();
      });
    }

    /* ---- find & replace ---- */
    let matches = [];
    let matchIdx = -1;
    let findTimer = 0;
    const hasHighlights = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight === 'function';
    function openFind(withReplace) {
      if (loading) return;
      if (readMode) setReadMode(false);
      saveSel();
      findPanel.hidden = false;
      replaceRow.hidden = !withReplace;
      const r = currentRange();
      const selText = r && !r.collapsed ? r.toString() : '';
      if (selText && selText.length < 100 && !selText.includes('\n')) findInput.value = selText;
      findInput.select();
      (withReplace && findInput.value ? replInput : findInput).focus();
      refreshFind(true);
    }
    function closeFind() {
      if (findPanel.hidden) return;
      findPanel.hidden = true;
      clearHighlights();
      if (matchIdx >= 0 && matches[matchIdx]) savedRange = matches[matchIdx].cloneRange();
      restoreSel();
    }
    function clearHighlights() {
      if (hasHighlights) { CSS.highlights.delete('word-find'); CSS.highlights.delete('word-find-current'); }
    }
    function buildRegex() {
      const q = findInput.value;
      if (!q) return null;
      let src = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      if (wordBox.checked) src = '(?<![\\p{L}\\p{N}_])' + src + '(?![\\p{L}\\p{N}_])';
      return new RegExp(src, 'gu' + (caseBox.checked ? '' : 'i'));
    }
    function collectMatches() {
      const re = buildRegex();
      if (!re) return [];
      const segs = [];
      let text = '', lastBlock = null;
      const w = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
      while (w.nextNode()) {
        const n = w.currentNode;
        const b = closestIn(n, BLOCK_SEL) || body;
        if (lastBlock && b !== lastBlock) text += '\u0001';
        lastBlock = b;
        segs.push({ node: n, start: text.length });
        text += n.nodeValue;
      }
      const locate = (off, isEnd) => {
        for (let i = 0; i < segs.length; i++) {
          const s = segs[i], len = s.node.nodeValue.length;
          if (off >= s.start && (isEnd ? off <= s.start + len : off < s.start + len)) return [s.node, off - s.start];
        }
        return null;
      };
      const out = [];
      let m;
      while ((m = re.exec(text))) {
        if (!m[0].length) { re.lastIndex++; continue; }
        const a = locate(m.index, false), b = locate(m.index + m[0].length, true);
        if (!a || !b) continue;
        const r = document.createRange();
        r.setStart(a[0], a[1]); r.setEnd(b[0], b[1]);
        out.push(r);
      }
      return out;
    }
    function paintMatches() {
      if (!hasHighlights) return;
      clearHighlights();
      if (matches.length) CSS.highlights.set('word-find', new Highlight(...matches));
      if (matches[matchIdx]) CSS.highlights.set('word-find-current', new Highlight(matches[matchIdx]));
    }
    function refreshFind(fromCaret) {
      matches = collectMatches();
      if (!matches.length) matchIdx = -1;
      else if (fromCaret || matchIdx < 0) {
        const r = savedRange;
        matchIdx = 0;
        if (r) {
          const i = matches.findIndex(m => m.compareBoundaryPoints(Range.START_TO_START, r) >= 0);
          matchIdx = i < 0 ? 0 : i;
        }
      } else matchIdx = Math.min(matchIdx, matches.length - 1);
      paintMatches();
      reportFind();
      if (matches[matchIdx]) scrollToRange(matches[matchIdx]);
    }
    function scheduleFindRefresh() { clearTimeout(findTimer); findTimer = setTimeout(() => refreshFind(false), 250); }
    function reportFind() {
      if (!findInput.value) findStatus.textContent = 'Type text to find.';
      else if (!matches.length) findStatus.textContent = 'No matches found.';
      else findStatus.textContent = 'Match ' + (matchIdx + 1) + ' of ' + matches.length + '.';
    }
    function scrollToRange(r) {
      const el = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      if (el && el.scrollIntoView) el.scrollIntoView({ block: 'center', behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
    }
    function findStep(dir) {
      if (!matches.length) { refreshFind(true); if (!matches.length) return; }
      else matchIdx = (matchIdx + dir + matches.length) % matches.length;
      paintMatches();
      reportFind();
      scrollToRange(matches[matchIdx]);
      savedRange = matches[matchIdx].cloneRange();
    }
    function replaceOne() {
      if (readMode || !matches[matchIdx]) { findStep(1); return; }
      const keep = document.activeElement;
      savedRange = matches[matchIdx].cloneRange();
      restoreSel();
      exec('insertText', replInput.value);
      afterEdit();
      matches = collectMatches();
      const r = window.getSelection().rangeCount ? window.getSelection().getRangeAt(0) : null;
      matchIdx = -1;
      if (matches.length) {
        const i = r ? matches.findIndex(m => m.compareBoundaryPoints(Range.START_TO_START, r) >= 0) : 0;
        matchIdx = i < 0 ? 0 : i;
      }
      paintMatches();
      reportFind();
      if (keep && keep.focus) keep.focus();
    }
    function replaceAll() {
      if (readMode) return;
      const list = collectMatches();
      if (!list.length) { findStatus.textContent = 'No matches found.'; return; }
      const keep = document.activeElement;
      body.focus({ preventScroll: true });
      const s = window.getSelection();
      for (let i = list.length - 1; i >= 0; i--) {
        s.removeAllRanges();
        s.addRange(list[i]);
        exec('insertText', replInput.value);
      }
      afterEdit();
      matches = collectMatches();
      matchIdx = matches.length ? 0 : -1;
      paintMatches();
      findStatus.textContent = 'Replaced ' + list.length + (list.length === 1 ? ' occurrence.' : ' occurrences.');
      App.toast('Replaced ' + list.length + ' occurrence' + (list.length === 1 ? '' : 's') + '.', 'success');
      if (keep && keep.focus) keep.focus();
    }
    findInput.addEventListener('input', () => refreshFind(true));
    caseBox.addEventListener('change', () => refreshFind(true));
    wordBox.addEventListener('change', () => refreshFind(true));
    findInput.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); findStep(e.shiftKey ? -1 : 1); } });
    replInput.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); replaceOne(); } });
    findPanel.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeFind(); } });

    /* ---- serialisation ---- */
    function cleanHtml() {
      const clone = body.cloneNode(true);
      clone.querySelectorAll('.word-img-selected').forEach(i => { i.classList.remove('word-img-selected'); if (!i.className) i.removeAttribute('class'); });
      clone.querySelectorAll('font').forEach(f => {
        const span = document.createElement('span');
        if (f.getAttribute('color')) span.style.color = f.getAttribute('color');
        if (f.getAttribute('face')) span.style.fontFamily = f.getAttribute('face');
        const sz = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 }[f.getAttribute('size')];
        if (sz) span.style.fontSize = sz + 'pt';
        if (f.getAttribute('style')) span.style.cssText += ';' + f.getAttribute('style');
        while (f.firstChild) span.appendChild(f.firstChild);
        f.replaceWith(span);
      });
      // CSS keyword font sizes -> pt
      const kw = { 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24, 'xxx-large': 36 };
      clone.querySelectorAll('[style*="font-size"]').forEach(el => {
        const v = el.style.fontSize;
        if (kw[v]) el.style.fontSize = kw[v] + 'pt';
        else if (/px$/.test(v)) el.style.fontSize = Math.round(parseFloat(v) * 0.75 * 2) / 2 + 'pt';
      });
      clone.querySelectorAll('span').forEach(sp => { if (!sp.attributes.length) sp.replaceWith(...sp.childNodes); });
      // zero-width spaces used as caret anchors
      const w = document.createTreeWalker(clone, NodeFilter.SHOW_TEXT);
      const texts = [];
      while (w.nextNode()) texts.push(w.currentNode);
      texts.forEach(t => { if (t.nodeValue.includes('​')) t.nodeValue = t.nodeValue.replace(/​/g, ''); });
      // bare text / inline nodes at top level -> paragraphs
      let pending = null;
      Array.from(clone.childNodes).forEach(n => {
        const isBlock = n.nodeType === 1 && /^(P|H[1-6]|BLOCKQUOTE|PRE|UL|OL|TABLE|HR|DIV)$/.test(n.tagName);
        if (isBlock) { pending = null; return; }
        if (n.nodeType === 3 && !n.nodeValue.trim() && !pending) { n.remove(); return; }
        if (!pending) { pending = document.createElement('p'); clone.insertBefore(pending, n); }
        pending.appendChild(n);
      });
      return { html: clone.innerHTML, root: clone };
    }
    function formatFor(path) {
      const e = extOf(path);
      if (e === '.txt') return 'txt';
      if (e === '.md' || e === '.markdown') return 'md';
      if (e === '.html' || e === '.htm') return 'html';
      return 'docx';
    }
    async function serialize(fmt) {
      const { html, root: clone } = cleanHtml();
      if (fmt === 'docx') {
        const res = await App.api.postBytes('/api/docx/from-html', new TextEncoder().encode(JSON.stringify({ html })), 'application/json');
        return { bytes: new Uint8Array(await res.arrayBuffer()), mime: DOCX_MIME };
      }
      if (fmt === 'txt') {
        const tmp = h('div', { class: 'word-offscreen', 'aria-hidden': 'true' });
        tmp.innerHTML = html;
        root.appendChild(tmp);
        const text = tmp.innerText.replace(/\n{3,}/g, '\n\n').replace(/\n+$/, '') + '\n';
        tmp.remove();
        return { bytes: new TextEncoder().encode(text.replace(/\n/g, '\r\n')), mime: 'text/plain' };
      }
      if (fmt === 'md') return { bytes: new TextEncoder().encode(htmlToMd(clone)), mime: 'text/markdown' };
      const title = escapeHtml(stripExt(docName));
      const full = '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<title>' + title + '</title>\n' +
        '<style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.4;max-width:8.5in;margin:1in auto;padding:0 1rem;color:#000;background:#fff}' +
        'table{border-collapse:collapse}td,th{border:1px solid #000;padding:4px 8px}img{max-width:100%}hr.page-break{border:0;page-break-after:always;break-after:page}h1.title{font-size:28pt;font-weight:400}' +
        'blockquote{margin:0 2em;font-style:italic;color:#404040}</style>\n</head>\n<body>\n' + html + '\n</body>\n</html>\n';
      return { bytes: new TextEncoder().encode(full), mime: 'text/html' };
    }
    async function writeTo(path) {
      stSave.textContent = 'Saving…';
      const { bytes } = await serialize(formatFor(path));
      const saved = await App.writeFile(path, bytes);
      return saved || path;
    }
    function fail(err) {
      const e = err instanceof Error ? err : new Error(String(err));
      updateStatus();
      throw e;
    }
    async function save() {
      if (loading) return false;
      if (!docPath) return saveAs();
      try {
        docPath = await writeTo(docPath);
        doc.path = docPath;
        setClean();
        App.toast('Saved ' + docName, 'success');
        return true;
      } catch (err) { return fail(err); }
    }
    async function saveAs() {
      if (loading) return false;
      saveSel();
      const ext = extOf(docName);
      const suggested = (SAVE_EXTS.includes(ext) || ext === '.htm') ? docName : stripExt(docName || 'Document1') + '.docx';
      let p = await App.pickFile({ mode: 'save', extensions: SAVE_EXTS, suggestedName: suggested });
      if (!p) return false;
      if (Array.isArray(p)) p = p[0];
      if (![...SAVE_EXTS, '.htm', '.markdown'].includes(extOf(p))) p += '.docx';
      const fmt = formatFor(p);
      if ((fmt === 'txt' || fmt === 'md') && body.querySelector('img,table,span[style],font,[style*="color"]')) {
        const ok = await App.confirm(fmt === 'txt'
          ? 'Plain text (.txt) keeps only the words. Formatting, pictures and tables will be lost. Save anyway?'
          : 'Markdown (.md) keeps headings, lists, bold, italic, links and tables, but colors, fonts and sizes will be lost. Save anyway?');
        if (!ok) return false;
      }
      try {
        docPath = await writeTo(p);
        doc.path = docPath;
        docName = baseName(docPath);
        doc.name = docName;
        App.setTitle(editor, docName);
        setClean();
        App.toast('Saved as ' + docName, 'success');
        return true;
      } catch (err) { return fail(err); }
    }
    async function downloadCopy() {
      try {
        const name = (SAVE_EXTS.includes(extOf(docName)) || extOf(docName) === '.htm') ? docName : stripExt(docName) + '.docx';
        const { bytes, mime } = await serialize(formatFor(name));
        App.download(new Blob([bytes], { type: mime }), name);
      } catch (err) { App.toast('Could not create a copy: ' + err.message, 'error'); }
    }
    function printDoc() {
      const { html } = cleanHtml();
      const ps = PAGE_SIZES[layout.size], mg = MARGINS[layout.margin];
      const pr = h('div', { id: 'word-print-root', class: 'word-print-root' });
      pr.innerHTML = html;
      const style = h('style', { id: 'word-print-style' }, '@page { size: ' + ps.css + ' ' + layout.orient + '; margin: ' + mg.y + 'in ' + mg.x + 'in; }');
      document.getElementById('word-print-root')?.remove();
      document.getElementById('word-print-style')?.remove();
      document.head.appendChild(style);
      document.body.appendChild(pr);
      document.body.classList.add('printing-word');
      const cleanup = () => {
        document.body.classList.remove('printing-word');
        pr.remove(); style.remove();
        window.removeEventListener('afterprint', cleanup);
      };
      window.addEventListener('afterprint', cleanup);
      setTimeout(() => { window.print(); setTimeout(cleanup, 1000); }, 50);
    }

    /* ---- keyboard ---- */
    function focusCycle(dir) {
      const zones = [];
      if (!ribbon.hidden) zones.push(activeTab().tab);
      zones.push(body);
      if (!findPanel.hidden) zones.push(findInput);
      zones.push(statusbar);
      const a = document.activeElement;
      let i = zones.findIndex(z => z === a || z.contains(a));
      if (i < 0 && ribbon.contains(a)) i = 0;
      const next = zones[(i + dir + zones.length) % zones.length];
      if (next === statusbar) { statusbar.tabIndex = -1; }
      if (next === body) restoreSel(); else next.focus();
    }
    on(root, 'keydown', e => {
      const ctrl = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      if (e.key === 'F6') { e.preventDefault(); focusCycle(e.shiftKey ? -1 : 1); return; }
      if (e.key === 'Escape' && readMode) { e.preventDefault(); setReadMode(false); return; }
      if (e.key === 'Escape' && !findPanel.hidden) { e.preventDefault(); closeFind(); return; }
      if (!ctrl || e.altKey) return;
      if (k === 'f' && !e.shiftKey) { e.preventDefault(); openFind(false); }
      else if (k === 'h' && !e.shiftKey) { e.preventDefault(); openFind(true); }
      else if (k === 'p' && !e.shiftKey) { e.preventDefault(); printDoc(); }
      else if (k === '=' || k === '+') { e.preventDefault(); stepZoom(1); }
      else if (k === '-' || k === '_') { e.preventDefault(); stepZoom(-1); }
      else if (k === '0' && !e.shiftKey) { e.preventDefault(); setZoom(100); }
    });
    on(body, 'keydown', e => {
      const ctrl = e.ctrlKey || e.metaKey;
      const k = e.key.toLowerCase();
      if (readMode) return;
      if (ctrl && !e.altKey) {
        const map = { b: 'bold', i: 'italic', u: 'underline' };
        if (map[k] && !e.shiftKey) { e.preventDefault(); cmd2(map[k]); return; }
        if (k === 'k') { e.preventDefault(); insertLink(); return; }
        if (k === 'l' && !e.shiftKey) { e.preventDefault(); cmd2('justifyLeft', null, true); return; }
        if (k === 'e' && !e.shiftKey) { e.preventDefault(); cmd2('justifyCenter', null, true); return; }
        if (k === 'r' && !e.shiftKey) { e.preventDefault(); cmd2('justifyRight', null, true); return; }
        if (k === 'j' && !e.shiftKey) { e.preventDefault(); cmd2('justifyFull', null, true); return; }
        if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); cmd2('redo'); return; }
        if (e.key === 'Enter') { e.preventDefault(); insertPageBreak(); return; }
        if (e.key === ']') { e.preventDefault(); stepFont(1); return; }
        if (e.key === '[') { e.preventDefault(); stepFont(-1); return; }
      }
      if (e.key === 'Tab' && !ctrl && !e.altKey) {
        const r = currentRange();
        const cell = r && closestIn(r.startContainer, 'td,th');
        if (cell) {
          e.preventDefault();
          const table = cell.closest('table');
          const cells = Array.from(table.querySelectorAll('td,th'));
          let i = cells.indexOf(cell) + (e.shiftKey ? -1 : 1);
          if (i >= cells.length) { tableOp('rowBelow'); return; }
          if (i < 0) i = 0;
          const target = cells[i];
          const rr = document.createRange();
          rr.selectNodeContents(target);
          const s = window.getSelection();
          s.removeAllRanges(); s.addRange(rr);
          return;
        }
        if (r && closestIn(r.startContainer, 'li')) { e.preventDefault(); indent(e.shiftKey ? -1 : 1); }
        // otherwise Tab leaves the document (no keyboard trap); use Increase indent for paragraphs
      }
    });
    on(body, 'input', () => { markDirty(); scheduleUI(); });
    on(body, 'click', e => {
      if (e.target.tagName === 'IMG') {
        const r = document.createRange();
        r.selectNode(e.target);
        const s = window.getSelection();
        s.removeAllRanges(); s.addRange(r);
        selectImage(e.target);
      } else if (selectedImg) selectImage(null);
      if (e.target.closest('a') && (e.ctrlKey || readMode)) {
        e.preventDefault();
        const href = e.target.closest('a').getAttribute('href');
        if (href && /^https?:/i.test(href)) window.open(href, '_blank', 'noopener');
      }
    });
    on(body, 'paste', e => {
      if (readMode || loading) return;
      const cd = e.clipboardData;
      if (!cd) return;
      const files = Array.from(cd.files || []).filter(f => f.type.startsWith('image/'));
      if (files.length) {
        e.preventDefault();
        Promise.all(files.map(f => bytesToDataUrl(f, f.type))).then(urls => {
          insertHtmlAtCaret(urls.map(u => '<img src="' + u + '" alt="" style="max-width:100%">').join(''));
          App.toast('Picture pasted. Use the Picture tab to add alt text.', 'info');
        });
        return;
      }
      const htmlData = cd.getData('text/html');
      if (htmlData) {
        e.preventDefault();
        const clean = sanitizeHtml(htmlData.replace(/<!--StartFragment-->|<!--EndFragment-->/g, ''));
        const tmp = document.createElement('div');
        tmp.innerHTML = clean;
        tmp.querySelectorAll('[class]').forEach(n => n.removeAttribute('class'));
        tmp.querySelectorAll('[style]').forEach(n => {
          const keep = ['font-weight', 'font-style', 'text-decoration', 'color', 'background-color', 'text-align'];
          const st = n.style;
          const kept = keep.map(p => st.getPropertyValue(p) ? p + ':' + st.getPropertyValue(p) : '').filter(Boolean).join(';');
          if (kept) n.setAttribute('style', kept); else n.removeAttribute('style');
        });
        exec('insertHTML', tmp.innerHTML);
        afterEdit();
      }
    });
    on(body, 'dragover', e => {
      if (e.dataTransfer && Array.from(e.dataTransfer.items || []).some(i => i.kind === 'file' && i.type.startsWith('image/'))) { e.preventDefault(); e.stopPropagation(); }
    });
    on(body, 'drop', e => {
      const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []).filter(f => f.type.startsWith('image/'));
      if (!files.length || readMode) return;
      e.preventDefault(); e.stopPropagation();
      if (document.caretRangeFromPoint) {
        const r = document.caretRangeFromPoint(e.clientX, e.clientY);
        if (r && inBody(r.startContainer)) savedRange = r;
      }
      Promise.all(files.map(f => bytesToDataUrl(f, f.type))).then(urls => {
        insertHtmlAtCaret(urls.map(u => '<img src="' + u + '" alt="" style="max-width:100%">').join(''));
      });
    });
    on(body, 'focus', () => { try { exec('defaultParagraphSeparator', 'p'); } catch (e) { /* ignore */ } });

    /* ---- load ---- */
    async function load() {
      const ext = extOf(docName || docPath);
      let html = '<p><br></p>';
      try {
        if (doc.bytes && doc.bytes.length) {
          if (ext === '.docx') {
            const res = await App.api.postBytes('/api/docx/to-html', doc.bytes, DOCX_MIME);
            const j = await res.json();
            html = sanitizeHtml(j.html || '') || '<p><br></p>';
          } else {
            const text = new TextDecoder('utf-8').decode(doc.bytes).replace(/^﻿/, '');
            if (ext === '.md' || ext === '.markdown') html = sanitizeHtml(mdToHtml(text));
            else if (ext === '.html' || ext === '.htm') html = sanitizeHtml(text) || '<p><br></p>';
            else html = textToHtml(text);
          }
        }
        body.innerHTML = html;
        if (!body.firstChild) body.innerHTML = '<p><br></p>';
        loading = false;
        body.setAttribute('contenteditable', 'true');
        updateStatus();
      } catch (err) {
        body.innerHTML = '';
        body.appendChild(h('p', { text: 'This document could not be opened: ' + (err.message || err) }));
        App.toast('Could not open ' + docName + ': ' + (err.message || err), 'error');
        loading = false;
        stSave.textContent = 'Open failed';
      }
    }

    editor = {
      save,
      saveAs,
      isDirty: () => dirty,
      focus() {
        if (readMode) { exitRead.focus(); return; }
        restoreSel();
      },
      destroy() {
        listeners.forEach(([t, ty, fn, o]) => t.removeEventListener(ty, fn, o));
        listeners.length = 0;
        clearTimeout(statusTimer); clearTimeout(findTimer);
        if (uiTimer) cancelAnimationFrame(uiTimer);
        clearHighlights();
        root.remove();
      }
    };
    if (!doc.name) queueMicrotask(() => App.setTitle(editor, docName));
    updateStatus();
    load().then(() => { if (document.activeElement === document.body || !document.activeElement) restoreSel(); });
    return editor;
  }

  function register() {
    window.App.registerEditor('word', {
      label: 'Document',
      icon: '\u{1F4C4}',
      extensions: ['.docx', '.txt', '.md', '.html', '.htm'],
      create: createWordEditor
    });
  }
  if (window.App && typeof window.App.registerEditor === 'function') register();
  else window.addEventListener('DOMContentLoaded', () => { if (window.App) register(); else console.error('word.js: window.App is missing'); });

  // exposed for testing
  window.DocStudioWord = { mdToHtml, htmlToMd, sanitizeHtml, textToHtml };
})();
