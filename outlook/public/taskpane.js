/* LS.md for Outlook — LS-md pane behaviour + speedDF markdown projection. */
(function () {
  "use strict";

  const STORAGE_KEY = "lsmd.outlook.draft.v2";
  const THEME_KEY = "lsmd.outlook.previewTheme";
  const ROOT_ID = "lsmd-root";

  const LIGHT = {
    paper: "#ffffff",
    text: "#0f172a",
    secondary: "#475569",
    muted: "#64748b",
    border: "#cbd5e1",
    borderSubtle: "#e2e8f0",
    accent: "#0891b2",
    accentText: "#0e7490",
    hover: "#e2e8f0",
    input: "#f1f5f9",
    kw: "#7c3aed",
    str: "#15803d",
    num: "#c2410c",
    comm: "#64748b",
    title: "#0e7490",
  };

  const TOOLS = [
    { id: "h1", label: "H1", title: "Heading 1" },
    { id: "h2", label: "H2", title: "Heading 2" },
    { id: "h3", label: "H3", title: "Heading 3" },
    { id: "bold", label: "B", title: "Bold  Ctrl+B" },
    { id: "italic", label: "I", title: "Italic  Ctrl+I" },
    { id: "strike", label: "S", title: "Strikethrough" },
    { id: "code", label: "</>", title: "Inline code" },
    { id: "link", label: "🔗", title: "Link  Ctrl+K" },
    { id: "ul", label: "•", title: "Bulleted list" },
    { id: "ol", label: "1.", title: "Numbered list" },
    { id: "check", label: "☑", title: "Checklist" },
    { id: "quote", label: "“", title: "Quote" },
    { id: "fence", label: "{ }", title: "Code block" },
    { id: "table", label: "⊞", title: "Table" },
    { id: "hr", label: "—", title: "Rule" },
  ];

  function lsGet(key) {
    try {
      return window.localStorage.getItem(key);
    } catch (err) {
      return null;
    }
  }

  function lsSet(key, value) {
    try {
      window.localStorage.setItem(key, value);
    } catch (err) {}
  }

  const state = {
    host: "web",
    live: false,
    mode: "block",
    previewTheme: lsGet(THEME_KEY) || "light",
    applied: "",
    syncing: false,
    timer: 0,
  };

  const els = {};

  function $(id) {
    return document.getElementById(id);
  }

  function wordCount(text) {
    const parts = (text || "").trim().split(/\s+/);
    return parts[0] ? parts.length : 0;
  }

  function setStatus(text, kind) {
    els.status.textContent = text;
    els.status.className = kind === "err" ? "err" : "";
  }

  function wrapSelection(value, start, end, before, after) {
    after = after == null ? before : after;
    const selected = value.slice(start, end) || "text";
    const next = value.slice(0, start) + before + selected + after + value.slice(end);
    return { value: next, start: start + before.length, end: start + before.length + selected.length };
  }

  function togglePrefix(value, start, end, prefix) {
    const lineStart = value.lastIndexOf("\n", Math.max(0, start - 1)) + 1;
    const foundEnd = value.indexOf("\n", end);
    const lineEnd = foundEnd === -1 ? value.length : foundEnd;
    const block = value.slice(lineStart, lineEnd);
    const lines = block.split("\n");
    const allPrefixed = lines.every((line) => line.startsWith(prefix) || line.length === 0);
    const nextLines = lines.map((line) => {
      if (!line) return line;
      if (allPrefixed) return line.startsWith(prefix) ? line.slice(prefix.length) : line;
      return line.startsWith(prefix) ? line : prefix + line;
    });
    const joined = nextLines.join("\n");
    return { value: value.slice(0, lineStart) + joined + value.slice(lineEnd), start: lineStart, end: lineStart + joined.length };
  }

  function insertSnippet(value, start, end, snippet) {
    const next = value.slice(0, start) + snippet + value.slice(end);
    const caret = start + snippet.length;
    return { value: next, start: caret, end: caret };
  }

  function applyPatch(patch) {
    const editor = els.editor;
    editor.value = patch.value;
    editor.focus();
    editor.setSelectionRange(patch.start, patch.end);
    onDraftChange();
  }

  function currentSel() {
    return { start: els.editor.selectionStart, end: els.editor.selectionEnd, value: els.editor.value };
  }

  function runTool(id) {
    const { value, start, end } = currentSel();
    if (id === "h1") return applyPatch(togglePrefix(value, start, end, "# "));
    if (id === "h2") return applyPatch(togglePrefix(value, start, end, "## "));
    if (id === "h3") return applyPatch(togglePrefix(value, start, end, "### "));
    if (id === "bold") return applyPatch(wrapSelection(value, start, end, "**"));
    if (id === "italic") return applyPatch(wrapSelection(value, start, end, "*"));
    if (id === "strike") return applyPatch(wrapSelection(value, start, end, "~~"));
    if (id === "code") return applyPatch(wrapSelection(value, start, end, "`"));
    if (id === "link") return applyPatch(wrapSelection(value, start, end, "[", "](https://)"));
    if (id === "ul") return applyPatch(togglePrefix(value, start, end, "- "));
    if (id === "ol") return applyPatch(togglePrefix(value, start, end, "1. "));
    if (id === "check") return applyPatch(togglePrefix(value, start, end, "- [ ] "));
    if (id === "quote") return applyPatch(togglePrefix(value, start, end, "> "));
    if (id === "fence") return applyPatch(insertSnippet(value, start, end, "\n```ts\n\n```\n"));
    if (id === "table") return applyPatch(insertSnippet(value, start, end, "\n| Column | Column |\n| --- | --- |\n|  |  |\n"));
    if (id === "hr") return applyPatch(insertSnippet(value, start, end, "\n\n---\n\n"));
  }

  function configureMarked() {
    if (!window.marked) return;
    const highlightMark = {
      name: "highlightMark",
      level: "inline",
      start: (src) => src.indexOf("=="),
      tokenizer(src) {
        const match = /^==([^=\n]+)==/.exec(src);
        if (!match) return;
        return { type: "highlightMark", raw: match[0], text: match[1] };
      },
      renderer(token) {
        return "<mark>" + token.text + "</mark>";
      },
    };
    if (typeof marked.use === "function") {
      marked.use({
        gfm: true,
        breaks: false,
        extensions: [highlightMark],
        renderer: {
          code(token) {
            const text = typeof token === "string" ? token : token.text || "";
            const lang = (typeof token === "string" ? arguments[1] : token.lang) || "";
            return renderFence(text, lang);
          },
        },
      });
    } else {
      marked.setOptions({ gfm: true, breaks: false });
    }
  }

  function escapeHtml(text) {
    return String(text == null ? "" : text)
      .replace(/&/g, "\u0026amp;")
      .replace(/</g, "\u0026lt;")
      .replace(/>/g, "\u0026gt;")
      .replace(/"/g, "\u0026quot;");
  }

  function renderFence(code, lang) {
    const info = String(lang || "").trim().split(/\s+/)[0];
    let body = escapeHtml(code);
    if (window.hljs) {
      try {
        if (info && hljs.getLanguage(info)) body = hljs.highlight(code, { language: info, ignoreIllegals: true }).value;
        else body = hljs.highlightAuto(code).value;
      } catch (_) {
        body = escapeHtml(code);
      }
    }
    const data = info ? ' data-lang="' + escapeHtml(info) + '"' : "";
    const cls = info ? " language-" + escapeHtml(info) : "";
    return "<pre" + data + "><code class=\"hljs" + cls + "\">" + body + "</code></pre>\n";
  }

  function parseMarkdown(source) {
    if (!source) return "";
    if (!window.marked) return "<p>" + escapeHtml(source).replace(/\n/g, "<br>") + "</p>";
    const html = marked.parse(source, { async: false });
    if (typeof html !== "string") return "";
    return html
      .replace(/<\/table>(?:\s*<p>\s*<\/p>)*/gi, "</table><p class=\"md-gap\">&nbsp;</p>")
      .replace(/<\/pre>(?:\s*<p>\s*<\/p>)*/gi, "</pre><p class=\"md-gap\">&nbsp;</p>");
  }

  function sanitize(html) {
    const doc = new DOMParser().parseFromString('<div id="root">' + html + "</div>", "text/html");
    const root = doc.getElementById("root");
    if (!root) return "";
    root.querySelectorAll("script,iframe,object,embed,link,meta,form,style").forEach((n) => n.remove());
    root.querySelectorAll("*").forEach((el) => {
      [...el.attributes].forEach((attr) => {
        const name = attr.name.toLowerCase();
        if (name.startsWith("on") || name === "srcdoc") el.removeAttribute(attr.name);
        if ((name === "href" || name === "src") && /^\s*(javascript|vbscript):/i.test(attr.value)) {
          el.removeAttribute(attr.name);
        }
      });
    });
    root.querySelectorAll("a").forEach((a) => {
      a.setAttribute("target", "_blank");
      a.setAttribute("rel", "noopener noreferrer");
    });
    return root.innerHTML;
  }

  function previewHtml(source) {
    const raw = parseMarkdown(source);
    return sanitize(raw) || '<p class="empty">Empty document</p>';
  }

  function styleAttr(map) {
    return Object.keys(map)
      .map((k) => k + ":" + map[k])
      .join(";");
  }

  function outlookHtml(source) {
    const c = LIGHT;
    const raw = sanitize(parseMarkdown(source));
    const doc = new DOMParser().parseFromString("<div id='frag'>" + raw + "</div>", "text/html");
    const frag = doc.getElementById("frag");

    const heading = (size, extra) =>
      styleAttr(
        Object.assign(
          {
            "font-family": "Segoe UI,system-ui,Arial,sans-serif",
            color: c.text,
            "font-weight": "700",
            "letter-spacing": "-0.02em",
            "line-height": "1.22",
            margin: extra.margin,
            "font-size": size,
          },
          extra.more || {},
        ),
      );

    frag.querySelectorAll("h1").forEach((el) => {
      el.setAttribute("style", heading("26px", { margin: "0 0 12px 0", more: { "padding-bottom": "8px", "border-bottom": "1px solid " + c.border } }));
    });
    frag.querySelectorAll("h2").forEach((el) => {
      el.setAttribute("style", heading("20px", { margin: "22px 0 8px 0", more: { "padding-bottom": "6px", "border-bottom": "1px solid " + c.borderSubtle } }));
    });
    frag.querySelectorAll("h3").forEach((el) => {
      el.setAttribute("style", heading("17px", { margin: "18px 0 6px 0" }));
    });
    frag.querySelectorAll("h4,h5,h6").forEach((el) => {
      el.setAttribute("style", heading("15px", { margin: "14px 0 6px 0" }));
    });
    frag.querySelectorAll("p").forEach((el) => {
      el.setAttribute("style", "margin:0 0 12px 0;font-family:Segoe UI,system-ui,Arial,sans-serif;font-size:15px;line-height:1.7;color:" + c.text + ";");
    });
    frag.querySelectorAll("li").forEach((el) => {
      el.setAttribute("style", "margin:4px 0;font-family:Segoe UI,system-ui,Arial,sans-serif;font-size:15px;line-height:1.7;color:" + c.text + ";");
    });
    frag.querySelectorAll("ul,ol").forEach((el) => {
      el.setAttribute("style", "margin:8px 0 14px 22px;padding:0;");
    });
    frag.querySelectorAll("blockquote").forEach((el) => {
      el.setAttribute(
        "style",
        "margin:14px 0;padding:8px 12px;border-left:3px solid " + c.accent + ";color:" + c.secondary + ";background:#e0f2fe;",
      );
    });
    frag.querySelectorAll("a").forEach((el) => {
      el.setAttribute("style", "color:" + c.accentText + ";text-decoration:underline;");
    });
    frag.querySelectorAll("hr").forEach((el) => {
      el.setAttribute("style", "border:0;border-top:1px solid " + c.border + ";margin:22px 0;");
    });
    frag.querySelectorAll("table").forEach((el) => {
      if (el.getAttribute("data-lsmd-codewrap")) return;
      el.setAttribute("border", "1");
      el.setAttribute("cellpadding", "6");
      el.setAttribute("cellspacing", "0");
      el.setAttribute(
        "style",
        "border-collapse:collapse;width:auto;max-width:100%;font-size:14px;margin:14px 0 22px 0;border:1px solid " + c.border + ";",
      );
    });
    frag.querySelectorAll("th").forEach((el) => {
      el.setAttribute("style", "border:1px solid " + c.border + ";padding:6px 10px;text-align:left;background:" + c.hover + ";font-weight:600;color:" + c.text + ";");
    });
    frag.querySelectorAll("td").forEach((el) => {
      el.setAttribute("style", "border:1px solid " + c.border + ";padding:6px 10px;text-align:left;color:" + c.text + ";");
    });
    frag.querySelectorAll("p.md-gap").forEach((el) => {
      el.setAttribute("style", "margin:0 0 14px 0;line-height:1;font-size:12px;");
      if (!el.innerHTML.trim()) el.innerHTML = "&nbsp;";
    });
    const tokenColor = {
      "hljs-keyword": c.kw,
      "hljs-built_in": c.kw,
      "hljs-type": c.kw,
      "hljs-selector-tag": c.kw,
      "hljs-string": c.str,
      "hljs-attr": c.str,
      "hljs-attribute": c.str,
      "hljs-comment": c.comm,
      "hljs-quote": c.comm,
      "hljs-number": c.num,
      "hljs-literal": c.num,
      "hljs-title": c.title,
      "hljs-function": c.title,
      "hljs-section": c.title,
      "hljs-variable": c.accentText,
      "hljs-params": c.accentText,
      "hljs-name": c.accentText,
      "hljs-meta": c.muted,
    };
    frag.querySelectorAll("code span[class]").forEach((span) => {
      const hit = String(span.className || "")
        .split(/\s+/)
        .find((cls) => tokenColor[cls]);
      if (hit) span.setAttribute("style", "color:" + tokenColor[hit] + ";");
    });
    frag.querySelectorAll("code").forEach((el) => {
      if (el.parentElement && el.parentElement.tagName === "PRE") {
        el.setAttribute("style", "font-family:Consolas,Courier New,monospace;font-size:13px;color:" + c.text + ";");
      } else {
        el.setAttribute("style", "font-family:Consolas,Courier New,monospace;font-size:13px;background:" + c.input + ";padding:1px 5px;border:1px solid " + c.borderSubtle + ";");
      }
    });
    frag.querySelectorAll("pre").forEach((el) => {
      el.setAttribute(
        "style",
        "margin:0;padding:12px 14px;background:" +
          c.input +
          ";border:0;font-family:Consolas,Courier New,monospace;font-size:13px;line-height:1.55;white-space:pre-wrap;",
      );
      const wrap = doc.createElement("table");
      wrap.setAttribute("data-lsmd-codewrap", "1");
      wrap.setAttribute("cellpadding", "0");
      wrap.setAttribute("cellspacing", "0");
      wrap.setAttribute(
        "style",
        "border-collapse:separate;width:auto;max-width:100%;margin:14px 0 22px 0;border:1px solid " + c.border + ";background:" + c.input + ";",
      );
      const row = doc.createElement("tr");
      const cell = doc.createElement("td");
      cell.setAttribute("style", "padding:0;border:0;");
      el.parentNode.insertBefore(wrap, el);
      cell.appendChild(el);
      row.appendChild(cell);
      wrap.appendChild(row);
    });
    frag.querySelectorAll("mark").forEach((el) => {
      el.setAttribute("style", "background:#fdba74;padding:0 2px;");
    });
    frag.querySelectorAll('input[type="checkbox"]').forEach((el) => {
      const box = doc.createTextNode(el.checked ? "☑ " : "☐ ");
      el.parentNode.replaceChild(box, el);
    });

    const inner = frag.innerHTML;
    const wrapStyle = styleAttr({
      "font-family": "Segoe UI,system-ui,Arial,sans-serif",
      "font-size": "15px",
      "line-height": "1.7",
      color: c.text,
    });
    return '<div id="' + ROOT_ID + '" data-lsmd="1" style="' + wrapStyle + '">' + inner + "</div><!--lsmd-end-->";
  }

  const BLOCK_RE = /<div[^>]*(?:id=["']lsmd-root["']|data-lsmd=["']1["'])[^>]*>[\s\S]*?<\/div>\s*(?:<!--lsmd-end-->)?/i;

  function upsertBlock(existing, block) {
    if (BLOCK_RE.test(existing)) return existing.replace(BLOCK_RE, block);
    const reply = existing.search(/<div[^>]+id=["']divRplyFwdMsg["']/i);
    const sig = existing.search(/<(?:div|p)[^>]+id=["'](?:Signature|sig)["']/i);
    let idx = -1;
    if (reply >= 0 && sig >= 0) idx = Math.min(reply, sig);
    else idx = reply >= 0 ? reply : sig;
    if (idx >= 0) return existing.slice(0, idx) + block + existing.slice(idx);
    if (!String(existing).replace(/<[^>]+>/g, "").trim()) return block;
    return block + existing;
  }

  function htmlToMarkdown(html) {
    const cleaned = String(html || "")
      .replace(/<xml[\s\S]*?<\/xml>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<!--\[if[\s\S]*?<!\[endif\]-->/gi, "")
      .replace(/<!--[\s\S]*?-->/g, "")
      .replace(/<\/?o:p[^>]*>/gi, "");
    const doc = new DOMParser().parseFromString(cleaned, "text/html");
    function esc(t) {
      return t.replace(/([\\`*_[\]#])/g, "\\$1");
    }
    function ser(node) {
      if (node.nodeType === 3) return esc(node.textContent || "");
      if (node.nodeType !== 1) return "";
      const el = node;
      const tag = el.tagName.toLowerCase();
      const inner = [...el.childNodes].map(ser).join("");
      if (tag === "h1") return "\n\n# " + inner.trim() + "\n\n";
      if (tag === "h2") return "\n\n## " + inner.trim() + "\n\n";
      if (tag === "h3") return "\n\n### " + inner.trim() + "\n\n";
      if (tag === "h4") return "\n\n#### " + inner.trim() + "\n\n";
      if (tag === "p" || tag === "div") return "\n\n" + inner + "\n\n";
      if (tag === "br") return "  \n";
      if (tag === "strong" || tag === "b") return "**" + inner + "**";
      if (tag === "em" || tag === "i") return "*" + inner + "*";
      if (tag === "s" || tag === "del") return "~~" + inner + "~~";
      if (tag === "code") return el.closest("pre") ? inner : "`" + inner + "`";
      if (tag === "pre") return "\n\n```\n" + (el.textContent || "").replace(/\n$/, "") + "\n```\n\n";
      if (tag === "blockquote") return "\n\n" + inner.trim().split("\n").map((l) => "> " + l).join("\n") + "\n\n";
      if (tag === "a") return "[" + inner + "](" + (el.getAttribute("href") || "") + ")";
      if (tag === "ul") {
        return "\n" + [...el.children].filter((c) => c.tagName === "LI").map((li) => "- " + ser(li).trim()).join("\n") + "\n";
      }
      if (tag === "ol") {
        return "\n" + [...el.children].filter((c) => c.tagName === "LI").map((li, i) => i + 1 + ". " + ser(li).trim()).join("\n") + "\n";
      }
      if (tag === "hr") return "\n\n---\n\n";
      if (tag === "table") {
        const rows = [...el.querySelectorAll("tr")];
        if (!rows.length) return "";
        const lines = [""];
        rows.forEach((row, i) => {
          const cells = [...row.querySelectorAll("th,td")].map((cell) => ser(cell).replace(/\n+/g, " ").trim());
          lines.push("| " + cells.join(" | ") + " |");
          if (i === 0) lines.push("| " + cells.map(() => "---").join(" | ") + " |");
        });
        lines.push("");
        return lines.join("\n");
      }
      return inner;
    }
    return ser(doc.body).replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
  }

  function renderPreview() {
    els.preview.setAttribute("data-theme", state.previewTheme);
    els.preview.innerHTML = previewHtml(els.editor.value);
    els.words.textContent = wordCount(els.editor.value) + " words";
  }

  function persist() {
    try {
      lsSet(STORAGE_KEY, els.editor.value);
    } catch (_) {}
  }

  function onDraftChange() {
    persist();
    renderPreview();
    if (els.editor.value !== state.applied) setStatus(state.live ? "Live pending" : "Unapplied");
    if (state.live && state.host === "outlook") scheduleApply();
  }

  function scheduleApply() {
    window.clearTimeout(state.timer);
    state.timer = window.setTimeout(() => {
      void applyToOutlook();
    }, 450);
  }

  function officeItem() {
    return window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item;
  }

  function promisify(fn) {
    return new Promise((resolve, reject) => {
      fn((result) => {
        if (!result || result.status === "failed" || (Office.AsyncResultStatus && result.status === Office.AsyncResultStatus.Failed)) {
          reject(new Error((result && result.error && result.error.message) || "Outlook call failed"));
          return;
        }
        resolve(result.value);
      });
    });
  }

  function getBodyHtml() {
    const item = officeItem();
    return promisify((cb) => item.body.getAsync(Office.CoercionType.Html, cb));
  }

  function setBodyHtml(html) {
    const item = officeItem();
    const options = { coercionType: Office.CoercionType.Html };
    if (Office.MailboxEnums && Office.MailboxEnums.BodyMode) {
      options.bodyMode = Office.MailboxEnums.BodyMode.HostConfig;
    }
    return promisify((cb) => item.body.setAsync(html, options, cb));
  }

  function setSelectedHtml(html) {
    const item = officeItem();
    return promisify((cb) => item.body.setSelectedDataAsync(html, { coercionType: Office.CoercionType.Html }, cb));
  }

  async function applyToOutlook() {
    const md = els.editor.value;
    const block = outlookHtml(md);
    if (state.host !== "outlook") {
      state.applied = md;
      setStatus("Preview only · not in Outlook");
      return;
    }
    state.syncing = true;
    setStatus("Updating…");
    try {
      if (state.mode === "insert") {
        await setSelectedHtml(block);
      } else if (state.mode === "replace") {
        await setBodyHtml(block);
      } else {
        const existing = await getBodyHtml();
        await setBodyHtml(upsertBlock(existing || "", block));
      }
      state.applied = md;
      setStatus("In sync");
    } catch (err) {
      setStatus(err.message || "Could not write body", "err");
    } finally {
      state.syncing = false;
    }
  }

  async function pullFromOutlook() {
    if (state.host !== "outlook") {
      setStatus("Pull is only available inside Outlook", "err");
      return;
    }
    try {
      setStatus("Reading…");
      const html = await getBodyHtml();
      const match = BLOCK_RE.exec(html || "");
      els.editor.value = htmlToMarkdown(match ? match[0] : html || "");
      state.applied = els.editor.value;
      persist();
      renderPreview();
      setStatus("Pulled from email");
    } catch (err) {
      setStatus(err.message || "Could not read body", "err");
    }
  }

  function setLayout(view) {
    if (view === "split") view = "write";
    document.querySelector(".workspace").setAttribute("data-layout", view);
    document.querySelectorAll(".tab").forEach((btn) => {
      btn.classList.toggle("is-on", btn.getAttribute("data-view") === view);
    });
  }

  function setMode(mode) {
    state.mode = mode;
    document.querySelectorAll(".seg-btn").forEach((btn) => {
      btn.classList.toggle("is-on", btn.getAttribute("data-mode") === mode);
    });
  }

  function bootOffice() {
    return new Promise((resolve) => {
      let settled = false;
      const done = (host) => {
        if (settled) return;
        settled = true;
        resolve(host);
      };
      const timer = window.setTimeout(() => done("web"), 4000);
      const attach = () => {
        if (!window.Office) return false;
        Office.initialize = Office.initialize || function () {};
        if (typeof Office.onReady !== "function") {
          window.clearTimeout(timer);
          done("web");
          return true;
        }
        Office.onReady((info) => {
          window.clearTimeout(timer);
          const host = String((info && info.host) || "").toLowerCase();
          done(host.indexOf("outlook") >= 0 ? "outlook" : "web");
        });
        return true;
      };
      if (!attach()) {
        const poll = window.setInterval(() => {
          if (attach()) window.clearInterval(poll);
        }, 50);
        window.setTimeout(() => window.clearInterval(poll), 4000);
      }
    });
  }

  function buildToolbar() {
    const bar = $("toolbar");
    TOOLS.forEach((tool) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "tool";
      if (tool.kind) btn.setAttribute("data-kind", tool.kind);
      btn.title = tool.title;
      btn.setAttribute("aria-label", tool.title);
      btn.textContent = tool.label;
      btn.addEventListener("mousedown", (e) => e.preventDefault());
      btn.addEventListener("click", () => runTool(tool.id));
      bar.appendChild(btn);
    });
  }

  async function init() {
    els.editor = $("editor");
    els.preview = $("preview");
    els.status = $("status-line");
    els.words = $("word-count");
    els.host = $("host-pill");

    try {
      configureMarked();
    } catch (err) {
      console.warn("marked unavailable", err);
    }
    try {
      buildToolbar();
    } catch (err) {
      console.warn("toolbar", err);
    }

    const saved = lsGet(STORAGE_KEY);
    els.editor.value = saved && saved.trim() ? saved : "";
    renderPreview();

    document.querySelectorAll(".tab").forEach((btn) => {
      btn.addEventListener("click", () => setLayout(btn.getAttribute("data-view")));
    });
    document.querySelectorAll(".seg-btn").forEach((btn) => {
      btn.addEventListener("click", () => setMode(btn.getAttribute("data-mode")));
    });
    $("live-toggle").addEventListener("click", () => {
      state.live = !state.live;
      $("live-toggle").setAttribute("aria-checked", state.live ? "true" : "false");
      if (state.live) scheduleApply();
    });
    $("apply-btn").addEventListener("click", () => void applyToOutlook());
    $("pull-btn").addEventListener("click", () => void pullFromOutlook());
    $("theme-btn").addEventListener("click", () => {
      state.previewTheme = state.previewTheme === "light" ? "dark" : "light";
      lsSet(THEME_KEY, state.previewTheme);
      renderPreview();
    });
    els.editor.addEventListener("input", onDraftChange);
    els.editor.addEventListener("keydown", (e) => {
      const meta = e.ctrlKey || e.metaKey;
      if (meta && e.key.toLowerCase() === "b") {
        e.preventDefault();
        runTool("bold");
      } else if (meta && e.key.toLowerCase() === "i") {
        e.preventDefault();
        runTool("italic");
      } else if (meta && e.key.toLowerCase() === "k") {
        e.preventDefault();
        runTool("link");
      } else if (meta && e.key.toLowerCase() === "s") {
        e.preventDefault();
        void applyToOutlook();
      }
    });

    state.host = await bootOffice();
    els.host.textContent = state.host === "outlook" ? "Outlook" : "Browser";
    setStatus(state.host === "outlook" ? "Ready" : "Browser preview · sideload the manifest in Outlook");
  }

  function showBootError(err) {
    const pill = document.getElementById("host-pill");
    const status = document.getElementById("status-line");
    if (pill) pill.textContent = "Error";
    if (status) status.textContent = String(err && err.message ? err.message : err);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => void init().catch(showBootError));
  } else {
    void init().catch(showBootError);
  }
})();
