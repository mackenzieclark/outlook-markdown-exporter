/* Copy as Markdown — Outlook add-in task pane.
 * Reads the open message, converts HTML body to Markdown (Turndown + GFM),
 * prefixes YAML frontmatter, copies to clipboard.
 */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var current = "";
  // Read back off the ?v= that build.sh stamps onto this script's URL, so the
  // version number lives in exactly one place: <Version> in the manifest.
  // Shown in the footer, which makes a stale cached script obvious at a glance.
  var VERSION = (function () {
    var m = /[?&]v=([^&]*)/.exec((document.currentScript && document.currentScript.src) || "");
    return m ? decodeURIComponent(m[1]) : "dev";
  })();

  function setStatus(msg, cls) {
    var el = $("status");
    el.textContent = msg;
    el.className = cls || "";
  }

  // ---------- metadata ----------
  function addr(e) {
    if (!e) return "";
    if (e.displayName && e.emailAddress && e.displayName !== e.emailAddress) {
      return e.displayName + " <" + e.emailAddress + ">";
    }
    return e.emailAddress || e.displayName || "";
  }
  function addrs(list) { return (list || []).map(addr).filter(Boolean); }
  function yamlStr(s) { return JSON.stringify(String(s == null ? "" : s)); }
  function yamlList(arr) {
    if (!arr.length) return "[]";
    return "\n" + arr.map(function (a) { return "  - " + yamlStr(a); }).join("\n");
  }

  function buildFrontmatter(item) {
    var lines = [
      "---",
      "subject: " + yamlStr(item.subject),
      "from: " + yamlStr(addr(item.from)),
      "to:" + yamlList(addrs(item.to)),
    ];
    var cc = addrs(item.cc);
    if (cc.length) lines.push("cc:" + yamlList(cc));
    var d = item.dateTimeCreated;
    if (d) lines.push("date: " + yamlStr(d.toISOString()));
    if (item.internetMessageId) lines.push("message_id: " + yamlStr(item.internetMessageId));
    if (item.conversationId) lines.push("conversation_id: " + yamlStr(item.conversationId));
    var att = (item.attachments || []).filter(function (a) { return !a.isInline; })
      .map(function (a) { return a.name; });
    if (att.length) lines.push("attachments:" + yamlList(att));
    lines.push("---", "");
    return lines.join("\n");
  }

  // ---------- HTML cleanup ----------
  function cleanHtml(html) {
    var doc = new DOMParser().parseFromString(html, "text/html");
    var kill = "style, script, head, meta, link, title, o\\:p, xml, " +
      "img[src^='cid:'], [style*='display:none'], [style*='display: none']";
    doc.querySelectorAll(kill).forEach(function (n) { n.remove(); });
    // Conditional comments / comments
    var walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_COMMENT);
    var comments = [];
    while (walker.nextNode()) comments.push(walker.currentNode);
    comments.forEach(function (c) { c.remove(); });
    // Word/Outlook puts paragraphs in <p class=MsoNormal> with &nbsp; fillers
    doc.querySelectorAll("p").forEach(function (p) {
      if (p.textContent.replace(/ /g, "").trim() === "" && !p.querySelector("img")) p.remove();
    });
    return doc.body;
  }

  // Split the body at quoted-reply boundaries. Outlook marks these with
  // <div id="divRplyFwdMsg">, <div id="appendonsend">, or a "From:" header block
  // inside a border-top div. Returns array of {headerText, body} where the first
  // entry has no header.
  function splitThread(body) {
    var sections = [];
    var markers = Array.prototype.slice.call(body.querySelectorAll(
      "div#divRplyFwdMsg, div#appendonsend, div.gmail_quote, blockquote[type='cite'], " +
      "div[style*='border-top'][style*='solid']"
    ));
    // Keep only top-level-ish markers (not nested inside an earlier marker)
    markers = markers.filter(function (m, i) {
      return !markers.slice(0, i).some(function (prev) { return prev.contains(m); });
    });
    if (!markers.length) return [{ header: "", node: body }];

    var cursor = body;
    var head = document.createElement("div");
    // Move everything before the first marker into head
    var first = markers[0];
    var range = document.createRange();
    range.setStartBefore(body.firstChild);
    range.setEndBefore(first);
    head.appendChild(range.extractContents());
    sections.push({ header: "", node: head });

    markers.forEach(function (m, i) {
      var seg = document.createElement("div");
      var r = document.createRange();
      r.setStartBefore(m);
      if (markers[i + 1]) r.setEndBefore(markers[i + 1]); else r.setEndAfter(body.lastChild);
      seg.appendChild(r.extractContents());
      // Try to pull "From: ... Sent: ... To: ... Subject: ..." header line
      var txt = seg.textContent.replace(/ /g, " ");
      var hm = txt.match(/From:\s*(.+?)\s*(?:Sent|Date):\s*(.+?)\s*(?:To|Cc|Subject):/i);
      var header = hm ? (hm[1].trim() + " — " + hm[2].trim()) : "Quoted message " + (i + 1);
      sections.push({ header: header, node: seg });
    });
    return sections;
  }

  // ---------- conversion ----------
  function makeTurndown() {
    var td = new TurndownService({
      headingStyle: "atx",
      codeBlockStyle: "fenced",
      bulletListMarker: "-",
      emDelimiter: "_",
    });
    td.use(turndownPluginGfm.gfm);
    td.addRule("dropEmptyLinks", {
      filter: function (n) { return n.nodeName === "A" && !n.textContent.trim(); },
      replacement: function () { return ""; },
    });
    // Safe-links unwrap (Defender rewrites hrefs)
    td.addRule("safelinks", {
      filter: function (n) {
        return n.nodeName === "A" && /safelinks\.protection\.outlook\.com/.test(n.getAttribute("href") || "");
      },
      replacement: function (content, n) {
        var href = n.getAttribute("href");
        try { href = new URL(href).searchParams.get("url") || href; } catch (e) {}
        href = decodeURIComponent(href);
        return content.trim() === href ? "<" + href + ">" : "[" + content + "](" + href + ")";
      },
    });
    return td;
  }

  function tidy(md) {
    return md
      .replace(/ /g, " ")
      .replace(/[ \t]+$/gm, "")
      .replace(/^(\s*)([-*]|\d+\.) {2,}/gm, "$1$2 ")
      .replace(/^\* \* \*$/gm, "---")
      .replace(/\n{3,}/g, "\n\n")
      .trim() + "\n";
  }

  function convert(item, html) {
    var td = makeTurndown();
    var body = cleanHtml(html);
    var out = $("frontmatter").checked ? buildFrontmatter(item) : "";
    out += "# " + (item.subject || "(no subject)") + "\n\n";
    if ($("splitThread").checked) {
      splitThread(body).forEach(function (s) {
        var md = tidy(td.turndown(s.node.innerHTML));
        if (!md.replace(/^(\* \* \*|---)$/gm, "").trim()) return; // empty / hr-only section
        if (s.header) out += "\n---\n\n## " + s.header + "\n\n";
        out += md + "\n";
      });
    } else {
      out += tidy(td.turndown(body.innerHTML));
    }
    return tidy(out);
  }

  // ---------- theme ----------
  // Outlook's theme can differ from the OS theme that prefers-color-scheme
  // reports, so prefer Office's own body colour when the host exposes it.
  function applyTheme() {
    try {
      var theme = Office.context && Office.context.officeTheme;
      var bg = theme && theme.bodyBackgroundColor;
      var m = /^#?([0-9a-f]{6})$/i.exec(String(bg || "").trim());
      if (!m) return;
      var n = parseInt(m[1], 16);
      // Rec. 601 luma; below half means the host is running a dark theme.
      var luma = 0.299 * (n >> 16 & 255) + 0.587 * (n >> 8 & 255) + 0.114 * (n & 255);
      document.documentElement.setAttribute("data-theme", luma < 128 ? "dark" : "light");
    } catch (e) { /* the prefers-color-scheme rules already cover this */ }
  }

  // ---------- footer ----------
  function showBuild() {
    $("version").textContent = "v" + VERSION;
    try {
      var d = Office.context.mailbox.diagnostics;
      $("host").textContent = [d.hostName, d.hostVersion].filter(Boolean).join(" ");
    } catch (e) { /* diagnostics are not worth failing over */ }
  }

  // ---------- actions ----------
  function copy() {
    if (!current) return;
    var done = function () { setStatus("Copied " + current.length + " chars to clipboard.", "ok"); };
    var legacy = function () {
      try {
        $("out").select();
        if (document.execCommand("copy")) return done();
      } catch (e) { /* fall through */ }
      setStatus("Clipboard blocked — select the text below and copy manually.", "err");
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(current).then(done, legacy);
    } else {
      legacy();
    }
  }

  function download() {
    if (!current) return;
    var item = Office.context.mailbox.item;
    var d = item.dateTimeCreated ? item.dateTimeCreated.toISOString().slice(0, 10) : "undated";
    var name = d + " " + (item.subject || "email").replace(/[\\/:*?"<>|]+/g, "_").slice(0, 80) + ".md";
    var blob = new Blob([current], { type: "text/markdown;charset=utf-8" });
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name; a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  function render(auto) {
    var item = Office.context.mailbox.item;
    setStatus("Reading message…");
    item.body.getAsync(Office.CoercionType.Html, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Failed to read body: " + res.error.message, "err");
        return;
      }
      try {
        current = convert(item, res.value);
        $("out").value = current;
        if (auto) {
          // Clipboard writes need a transient user activation, so on open
          // there is no gesture to spend — render and wait for the button.
          setStatus(current.length + " chars ready — click Copy Markdown.");
        } else {
          copy();
        }
      } catch (e) {
        setStatus("Conversion error: " + e.message, "err");
        console.error(e);
      }
    });
  }

  Office.onReady(function () {
    $("copy").addEventListener("click", function () { copy(); });
    $("download").addEventListener("click", function () { download(); });
    $("frontmatter").addEventListener("change", function () { render(false); });
    $("splitThread").addEventListener("change", function () { render(false); });
    applyTheme();
    showBuild();
    render(true);
  });
})();
