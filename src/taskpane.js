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
  function cleanHtml(html, strip) {
    var doc = new DOMParser().parseFromString(html, "text/html");
    var kill = "style, script, head, meta, link, title, o\\:p, xml, " +
      "img[src^='cid:'], [style*='display:none'], [style*='display: none']";
    doc.querySelectorAll(kill).forEach(function (n) { n.remove(); });
    // Conditional comments / comments
    var walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_COMMENT);
    var comments = [];
    while (walker.nextNode()) comments.push(walker.currentNode);
    comments.forEach(function (c) { c.remove(); });
    // Boilerplate goes first: what identifies a signature is the little stack of
    // short lines inside a layout table, and normalizeTables is about to take
    // that table apart.
    if (strip) removeBoilerplate(doc.body);
    normalizeTables(doc.body);
    // Word/Outlook puts paragraphs in <p class=MsoNormal> with &nbsp; fillers
    doc.querySelectorAll("p").forEach(function (p) {
      if (p.textContent.replace(/ /g, "").trim() === "" && !p.querySelector("img")) p.remove();
    });
    return doc.body;
  }

  // ---------- tables ----------
  // turndown-plugin-gfm converts a <table> only when every cell of its first row
  // is a <th>. Everything else it hands to keep(), which emits the element's raw
  // outerHTML — and Turndown does not descend into a kept element, so the
  // safelinks and empty-link rules never ran inside one either. Outlook builds
  // every signature, logo strip and two-column arrangement out of borderless
  // <td> tables, so in practice that put <table style="…"> blobs in the output
  // and took the link handling down with them. The two shapes need two different
  // answers, and cleanHtml settles both before Turndown ever sees the element.

  var CELL_BLOCK = "p, div, table, ul, ol, blockquote, pre, hr, h1, h2, h3, h4, h5, h6";
  var BLOCK_NAME = /^(P|DIV|TD|TH|TR|TABLE|THEAD|TBODY|TFOOT|LI|UL|OL|DL|DT|DD|H[1-6]|BLOCKQUOTE|PRE|HR|SECTION|ARTICLE|HEADER|FOOTER)$/;

  function list(nodes) { return Array.prototype.slice.call(nodes); }
  function flat(s) { return String(s == null ? "" : s).replace(/[\s ]+/g, " ").trim(); }
  function textOf(n) { return flat(n.textContent); }

  // Direct rows and cells only. table.rows would be ambiguous about a table
  // nested inside a cell, and telling those two apart is the whole job here.
  function tableRows(t) {
    var out = [];
    var push = function (r) { if (r.nodeName === "TR") out.push(r); };
    list(t.childNodes).forEach(function (n) {
      if (n.nodeName === "TR") push(n);
      else if (/^(THEAD|TBODY|TFOOT)$/.test(n.nodeName)) list(n.childNodes).forEach(push);
    });
    return out;
  }

  function rowCells(tr) {
    return list(tr.childNodes).filter(function (n) {
      return n.nodeName === "TD" || n.nodeName === "TH";
    });
  }

  function unwrapNode(n) {
    var p = n.parentNode;
    while (n.firstChild) p.insertBefore(n.firstChild, n);
    p.removeChild(n);
  }

  // Outlook writes a header row as bold text in ordinary <td>s far more often
  // than as <th>, so both count as one.
  function isBoldCell(c) {
    var t = textOf(c).replace(/\s/g, "");
    var seen = "";
    if (!t) return false;
    list(c.querySelectorAll("b, strong")).forEach(function (b) { seen += textOf(b); });
    if (seen.replace(/\s/g, "").indexOf(t) >= 0) return true;
    return /font-weight\s*:\s*(bold|[6-9]00)/i.test(c.getAttribute("style") || "");
  }

  function headerRow(t) {
    var first = tableRows(t)[0];
    var cells = first ? rowCells(first) : [];
    if (!cells.length) return null;
    if (cells.every(function (c) { return c.nodeName === "TH"; })) return first;
    if (cells.every(isBoldCell)) return first;
    return null;
  }

  function tableShape(t) {
    var s = { rows: 0, cols: 0, uniform: true, complex: false, nested: false };
    var rows = tableRows(t);
    s.rows = rows.length;
    rows.forEach(function (tr, i) {
      var cs = rowCells(tr);
      if (i === 0) s.cols = cs.length;
      else if (cs.length !== s.cols) s.uniform = false;
      cs.forEach(function (c) {
        if (c.querySelector("table")) s.nested = true;
        if (c.querySelectorAll(CELL_BLOCK).length > 1) s.complex = true;
        if (c.getAttribute("colspan") || c.getAttribute("rowspan")) s.complex = true;
      });
    });
    return s;
  }

  // A grid, or an arrangement? A data table is at least two rows by two columns,
  // has the same number of cells in every row, and holds nothing but simple cell
  // content — and has no table inside a cell, which is the giveaway that the
  // outer one is scaffolding. A drawn border rescues the awkward middle: a table
  // Outlook gave visible rules to is a grid even when its cells are busy.
  function isDataTable(t) {
    var s = tableShape(t);
    if (s.rows < 2 || s.cols < 2) return false;
    if (s.nested || !s.uniform) return false;
    if (headerRow(t)) return true;
    if (s.complex) return /^[1-9]/.test(t.getAttribute("border") || "");
    return true;
  }

  // gfm reads rows[0] and asks isHeadingRow(); a <thead> answers yes whatever
  // colgroups or stray sections Outlook left lying around inside the table.
  function headTable(t, tr) {
    var head = t.ownerDocument.createElement("thead");
    head.appendChild(tr);
    t.insertBefore(head, t.firstChild);
  }

  function toTh(td) {
    var th = td.ownerDocument.createElement("th");
    var align = td.getAttribute("align");
    if (align) th.setAttribute("align", align);
    while (td.firstChild) th.appendChild(td.firstChild);
    td.parentNode.replaceChild(th, td);
  }

  // gfm works out a cell's column from its index among the row's childNodes, so
  // the newlines Outlook leaves between </td> and <td> would shift every column
  // by one and lose the leading pipe.
  function tidyRow(tr) {
    list(tr.childNodes).forEach(function (n) {
      if (n.nodeType === 3 && !flat(n.nodeValue)) tr.removeChild(n);
    });
  }

  function makeConvertible(t) {
    var hr = headerRow(t);
    var tr, i, cols;
    tableRows(t).forEach(tidyRow);
    if (hr) {
      rowCells(hr).forEach(function (c) { if (c.nodeName !== "TH") toTh(c); });
      // It is a header row now; leaving the bold on would render it as ** **.
      list(hr.querySelectorAll("b, strong")).forEach(unwrapNode);
      headTable(t, hr);
      return;
    }
    // Nothing to promote, so synthesize an empty header. A label/value table has
    // no header row at all, and eating its first row to make one would silently
    // drop a pair.
    cols = tableShape(t).cols;
    tr = t.ownerDocument.createElement("tr");
    for (i = 0; i < cols; i++) tr.appendChild(t.ownerDocument.createElement("th"));
    headTable(t, tr);
  }

  // Replace a layout table with its cells' contents as blocks, in document
  // order, so the text survives and every inner rule gets to run on it.
  function unwrapTable(t) {
    var d = t.ownerDocument;
    var frag = d.createDocumentFragment();
    tableRows(t).forEach(function (tr) {
      rowCells(tr).forEach(function (c) {
        var div;
        if (!textOf(c) && !c.querySelector("img, hr")) return; // spacer cell
        div = d.createElement("div");
        while (c.firstChild) div.appendChild(c.firstChild);
        frag.appendChild(div);
      });
    });
    t.parentNode.replaceChild(frag, t);
  }

  function normalizeTables(body) {
    // Innermost first. A nested table always follows its container in document
    // order, so walking the list backwards settles the children before the
    // parent, and the parent is then classified on what its cells became. Real
    // signatures nest three deep, which is exactly this case.
    list(body.querySelectorAll("table")).reverse().forEach(function (t) {
      if (!body.contains(t)) return;
      if (!tableRows(t).length) { t.remove(); return; }
      if (isDataTable(t)) makeConvertible(t); else unwrapTable(t);
    });
  }

  // ---------- signatures and boilerplate ----------
  var CLIENT_FOOTER = /^(sent from my [\w' ]{1,24}|sent from mail for windows|sent from outlook( for (ios|android))?|get outlook for (ios|android))[.!]?$/i;
  var DISCLAIMER = /(intended recipient|confidentiality notice|privileged and confidential|received this (e-?mail |message |transmission )?in error|this e-?mail and any attachments|notify the sender (immediately )?and delete)/i;
  var BOILER_SEL = "p, div, td, th, table, blockquote, span, font, section";
  var EMPTY_SEL = "p, div, span, font, table, blockquote";

  // Lines as the reader sees them: <br> and every block boundary start a new
  // one. A signature is a stack of these; a paragraph of prose is one long one.
  function textLines(el) {
    var out = [];
    var buf = "";
    var flush = function () { var s = flat(buf); if (s) out.push(s); buf = ""; };
    var walk = function (n) {
      list(n.childNodes).forEach(function (c) {
        if (c.nodeType === 3) buf += c.nodeValue;
        else if (c.nodeName === "BR") flush();
        else if (BLOCK_NAME.test(c.nodeName)) { flush(); walk(c); flush(); }
        else walk(c);
      });
    };
    walk(el);
    flush();
    return out;
  }

  function hasAddress(s) { return new RegExp(EMAIL_SRC).test(s); }

  // With the table handling above in place the signal is reliable: a layout
  // table whose text is a handful of short lines and carries an address is a
  // signature — name, job title, phone, address, company site. An address on its
  // own is not enough, because prose quotes addresses too; the short-line shape
  // has to be there as well.
  function isSignatureTable(t) {
    var lines = textLines(t);
    var text = lines.join(" ");
    if (lines.length < 2 || lines.length > 10) return false;
    if (text.length > 400 || !hasAddress(text)) return false;
    if (isDataTable(t)) return false;
    return lines.every(function (l) { return l.length <= 80; });
  }

  function tinyImg(im) {
    var s = im.getAttribute("style") || "";
    var dim = function (prop) {
      var m = new RegExp(prop + "\\s*:\\s*([\\d.]+)", "i").exec(s);
      var v = m ? parseFloat(m[1]) : parseFloat(im.getAttribute(prop));
      return isNaN(v) ? Infinity : v;
    };
    return dim("width") <= 2 || dim("height") <= 2;
  }

  // A parent absorbs the matched block only if it holds nothing else: no loose
  // text of its own, and every block inside it is disclaimer too. That is what
  // stops the climb from swallowing a short message whose last paragraph happens
  // to be the legal footer.
  function onlyDisclaimer(el) {
    var own = "";
    var kids = list(el.children);
    if (!kids.length || textOf(el).length > 1500) return false;
    if (el.querySelector("div#divRplyFwdMsg, div#appendonsend")) return false;
    list(el.childNodes).forEach(function (c) { if (c.nodeType === 3) own += c.nodeValue; });
    if (flat(own)) return false;
    return kids.every(function (c) { return !textOf(c) || DISCLAIMER.test(textOf(c)); });
  }

  // A disclaimer is usually several paragraphs inside one wrapper, so climb from
  // the block that matched up to the wrapper — and no further.
  function disclaimerRoot(n, body) {
    var p = n.parentNode;
    while (p && p !== body && p.nodeType === 1 && !p.id && onlyDisclaimer(p)) {
      n = p; p = n.parentNode;
    }
    return n;
  }

  function removeBoilerplate(body) {
    // Tracking pixels and layout spacers.
    list(body.querySelectorAll("img")).forEach(function (im) {
      if (tinyImg(im)) im.remove();
    });
    // Social icon strips: an anchor wrapping only an image has no text to keep.
    list(body.querySelectorAll("a")).forEach(function (a) {
      if (!textOf(a)) a.remove();
    });
    // Signatures, outermost first so a nested one goes with its container.
    list(body.querySelectorAll("table")).forEach(function (t) {
      if (body.contains(t) && isSignatureTable(t)) t.remove();
    });
    // Client footers: the whole block is the footer, or it is not one at all.
    list(body.querySelectorAll(BOILER_SEL)).forEach(function (n) {
      if (body.contains(n) && CLIENT_FOOTER.test(textOf(n))) n.remove();
    });
    // Legal disclaimers, innermost first, so the climb starts at the match and
    // not at some wrapper that happens to contain the whole message.
    list(body.querySelectorAll(BOILER_SEL)).reverse().forEach(function (n) {
      if (body.contains(n) && DISCLAIMER.test(textOf(n))) disclaimerRoot(n, body).remove();
    });
    // Whatever the removals emptied out. Children first, so a container left
    // holding nothing but empty children goes too. Anything with an id stays:
    // splitThread uses empty #appendonsend / #divRplyFwdMsg divs as markers.
    list(body.querySelectorAll(EMPTY_SEL)).reverse().forEach(function (n) {
      if (!body.contains(n) || n.id) return;
      if (textOf(n) || n.querySelector("img, hr")) return;
      n.remove();
    });
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

  // ---------- people ----------
  // "Strip email addresses to names" rewrites every place a person can appear
  // in the finished Markdown: frontmatter from/to/cc, the "## Sender — date"
  // thread headers, the **From:/To:/Cc:** blocks quoted inside the body, bare
  // addresses in prose, and mailto: links.
  //
  // Aliases are keyed on the lowercased ADDRESS, never on the display name.
  // One person turns up as "Bob", "Bob Smith" and "bob@example.com" in the same
  // thread, and two different people can share a display name, so the address
  // is the identity. Numbers are handed out in order of first appearance in the
  // document — no hashing, no object key order — so the same message always
  // produces the same aliases. A first pass links display names to addresses
  // before any number is assigned, which is what lets a name-only mention in a
  // quoted header ("**From:** Bob") resolve to the same alias as the addressed
  // copies, even when the name-only one comes first.

  var EMAIL_SRC = "[A-Za-z0-9._%+\\-]+@[A-Za-z0-9\\-]+(?:\\.[A-Za-z0-9\\-]+)*\\.[A-Za-z]{2,}";
  // Inside a header field the entire value is people, so a display name may be
  // anything up to the angle bracket. The same pattern in running prose would
  // swallow the sentence in front of the address ("write to <bob@…>"), so there
  // a name has to look like one: up to four capitalised words.
  var FIELD_NAME_SRC = "(?:\"([^\"\\r\\n]*)\"|([^<>;,\\r\\n]*?))";
  var TEXT_NAME_SRC = "(?:\"([^\"\\r\\n]{1,80})\"|" +
    "((?:[A-Z\\u00c0-\\u024f][^\\s<>;,]*)(?:[ \\u00a0][A-Z\\u00c0-\\u024f][^\\s<>;,]*){0,3})?)";
  var FM_PERSON_KEY = /^(from|to|cc|bcc|reply_to)$/;
  var QUOTE_LABEL = /^\*\*[A-Za-z-]+\s*:\*\*$/;
  var QUOTE_PERSON_LABEL = /^\*\*(from|to|cc|bcc|reply-to)\s*:\*\*$/i;

  // Alternation, in match order: [text](mailto:…) | mailto:… | Name <addr> |
  // bare addr. Group layout is fixed so one scanner handles both name patterns:
  // 1 link text, 2 link address, 3 mailto address, 4/5 quoted/plain name,
  // 6 bracketed address, 7 the character in front of a bare address, 8 it.
  function addrSrc(nameSrc) {
    return "\\[([^\\]\\r\\n]*)\\]\\(\\s*mailto:(" + EMAIL_SRC + ")[^)\\r\\n]*\\)" +
      "|<?\\s*mailto:(" + EMAIL_SRC + ")\\s*>?" +
      "|" + nameSrc + "\\s*<\\s*(" + EMAIL_SRC + ")\\s*>" +
      "|(^|[^A-Za-z0-9._%+\\-@/])(" + EMAIL_SRC + ")";
  }

  function has(o, k) { return Object.prototype.hasOwnProperty.call(o, k); }
  function cleanName(s) {
    return String(s == null ? "" : s)
      .replace(/[\s ]+/g, " ")
      .replace(/^[\s"']+|[\s"',;]+$/g, "");
  }
  function normName(s) { return cleanName(s).toLowerCase(); }

  function newPeople() {
    return {
      byName: {},  // display name -> address
      byFirst: {}, // first word of a display name -> address, "" once ambiguous
      nameOf: {},  // address -> display name as first seen
      alias: {},   // identity key -> "UserN"
      count: 0,
    };
  }

  function learnPerson(p, name, email) {
    var a = String(email || "").toLowerCase();
    var n = normName(name);
    var f;
    if (!a || !n || n === a) return;
    if (!has(p.byName, n)) p.byName[n] = a;
    if (!has(p.nameOf, a)) p.nameOf[a] = cleanName(name);
    f = n.split(" ")[0];
    if (!has(p.byFirst, f)) p.byFirst[f] = a;
    else if (p.byFirst[f] !== a) p.byFirst[f] = "";
  }

  function personKey(p, name, email) {
    var n = normName(name);
    var f;
    if (email) return "a:" + String(email).toLowerCase();
    if (!n) return "";
    if (has(p.byName, n)) return "a:" + p.byName[n];
    f = n.split(" ")[0];
    if (has(p.byFirst, f) && p.byFirst[f]) return "a:" + p.byFirst[f];
    return "n:" + n; // nothing to tie this name to; still stable within the run
  }

  function aliasFor(p, key) {
    if (!has(p.alias, key)) { p.count += 1; p.alias[key] = "User" + p.count; }
    return p.alias[key];
  }

  function displayFor(p, name, email) {
    var a = String(email || "").toLowerCase();
    var disp = cleanName(name);
    var local;
    if (!disp || disp.toLowerCase() === a) disp = a && has(p.nameOf, a) ? p.nameOf[a] : "";
    if (!disp && a) {
      // The message never spells this address out as a name, so the local part
      // is the only thing left to call the person.
      local = a.split("@")[0].split(/[._+\-]/)[0];
      disp = local ? local.charAt(0).toUpperCase() + local.slice(1) : "";
    }
    return disp;
  }

  function firstNameOf(disp) {
    var s = disp.indexOf(",") >= 0 ? disp.slice(disp.lastIndexOf(",") + 1) : disp; // "Smith, Bob"
    return cleanName(s).split(" ")[0];
  }

  function personLabel(p, name, email, mode) {
    var key, disp;
    if (mode === "alias") {
      key = personKey(p, name, email);
      return key ? aliasFor(p, key) : cleanName(name);
    }
    disp = displayFor(p, name, email);
    if (mode === "first") disp = firstNameOf(disp);
    return disp || cleanName(name);
  }

  // Rewrite every address-shaped match in s, handing the gaps between them to
  // gap() — verbatim in prose, person-by-person inside a header field.
  function scanAddrs(s, nameSrc, gap, visit) {
    var re = new RegExp(addrSrc(nameSrc), "g");
    var out = "";
    var last = 0;
    var m, name, email, pre;
    while ((m = re.exec(s))) {
      pre = "";
      if (m[2] != null) { name = m[1]; email = m[2]; }
      else if (m[3] != null) { name = ""; email = m[3]; }
      else if (m[6] != null) {
        name = m[4] != null ? m[4] : (m[5] != null ? m[5] : "");
        email = m[6];
        // An unquoted display name absorbs the space that separated it from
        // "**From:**" or from the previous recipient; put it back.
        if (m[4] == null) pre = /^\s*/.exec(name)[0];
      } else { name = ""; email = m[8]; pre = m[7]; }
      out += gap(s.slice(last, m.index), visit) + pre + visit(name, email);
      last = m.index + m[0].length;
    }
    return out + gap(s.slice(last), visit);
  }

  function keepGap(g) { return g; }

  // A gap inside a header field is the recipients that carry no address, e.g.
  // "**To:** Bob" or "Alice; Bob <bob@example.com>".
  function looksLikeName(s) {
    if (!/[A-Za-zÀ-ɏ]/.test(s)) return false;
    if (s.length > 60) return false;
    if (/[*|<>@:]/.test(s)) return false;
    if (/\d{4}/.test(s)) return false; // a date, not a person
    return s.split(/\s+/).length <= 5;
  }

  function nameGap(g, visit) {
    if (!g) return "";
    return g.split(/([;,])/).map(function (part) {
      var m;
      if (part === ";" || part === ",") return part;
      m = /^(\s*)([\s\S]*?)(\s*)$/.exec(part);
      if (!m[2] || !looksLikeName(m[2])) return part;
      return m[1] + visit(m[2], "") + m[3];
    }).join("");
  }

  function fieldText(s, visit) { return scanAddrs(s, FIELD_NAME_SRC, nameGap, visit); }
  function proseText(s, visit) { return scanAddrs(s, TEXT_NAME_SRC, keepGap, visit); }

  // Frontmatter values are JSON strings; unwrap, rewrite, re-quote.
  function fmValue(raw, visit) {
    var parsed = null;
    var quoted = false;
    var v = raw;
    try { parsed = JSON.parse(raw); } catch (e) { /* bare scalar, e.g. [] */ }
    if (typeof parsed === "string") { v = parsed; quoted = true; }
    v = fieldText(v, visit);
    return quoted ? yamlStr(v) : v;
  }

  function bodyLine(line, visit) {
    // "## Bob <bob@example.com> — Tuesday…": everything before the em dash is
    // the sender, the tail is a date.
    var m = /^(##\s+)([\s\S]*?)(\s—\s[\s\S]*)$/.exec(line);
    var parts, res, isPerson, j;
    if (m) return m[1] + fieldText(m[2], visit) + proseText(m[3], visit);
    // Quoted header block. Splitting on the bold labels keeps "**Sent:**" and
    // "**Subject:**" out of the person handling even when Outlook runs the
    // whole block onto one line.
    if (/\*\*[A-Za-z-]+\s*:\*\*/.test(line)) {
      parts = line.split(/(\*\*[A-Za-z-]+\s*:\*\*)/);
      res = "";
      isPerson = false;
      for (j = 0; j < parts.length; j++) {
        if (QUOTE_LABEL.test(parts[j])) {
          isPerson = QUOTE_PERSON_LABEL.test(parts[j]);
          res += parts[j];
        } else {
          res += isPerson ? fieldText(parts[j], visit) : proseText(parts[j], visit);
        }
      }
      return res;
    }
    return proseText(line, visit);
  }

  // One walk of the document, in reading order, calling visit(name, email) at
  // every person. Used twice: once to learn, once to rewrite.
  function walkPeople(md, visit) {
    var lines = md.split("\n");
    var fm = lines[0] === "---";
    var key = "";
    var out = [];
    var i, line, m;
    for (i = 0; i < lines.length; i++) {
      line = lines[i];
      if (fm && i > 0) {
        if (line === "---") { fm = false; out.push(line); continue; }
        m = /^([A-Za-z_][A-Za-z0-9_]*):[ \t]*([\s\S]*)$/.exec(line);
        if (m) {
          key = m[1].toLowerCase();
          // message_id and conversation_id are addresses in shape only.
          out.push(FM_PERSON_KEY.test(key) && m[2] ? m[1] + ": " + fmValue(m[2], visit) : line);
          continue;
        }
        m = /^([ \t]*-[ \t]+)([\s\S]*)$/.exec(line);
        if (m && FM_PERSON_KEY.test(key)) { out.push(m[1] + fmValue(m[2], visit)); continue; }
        out.push(line);
        continue;
      }
      out.push(bodyLine(line, visit));
    }
    return out.join("\n");
  }

  function stripAddresses(md, mode) {
    var p = newPeople();
    // Pass one only records who is who; its output is thrown away so that no
    // alias number is assigned before every name-to-address link is known.
    walkPeople(md, function (name, email) { learnPerson(p, name, email); return ""; });
    return walkPeople(md, function (name, email) { return personLabel(p, name, email, mode); });
  }

  function nameMode() {
    if ($("nameAlias").checked) return "alias";
    if ($("nameFirst").checked) return "first";
    return "display";
  }

  function syncNameMode() { $("nameModeGroup").hidden = !$("stripEmails").checked; }

  // ---------- link protectors ----------
  // Mail gateways rewrite every href in the body to point at themselves. Two
  // kinds, and they need different treatment. Defender keeps the real URL in the
  // ?url= parameter, so it can be recovered exactly. Mimecast and its like hash
  // it into an opaque path — nothing in the href can be decoded back. For those
  // the visible text is the only thing left, and it is a safe fallback in the
  // one case that matters: a protector wrapped a link whose text is itself a
  // URL, so the text is the link. Anything else keeps the rewritten href rather
  // than guess.
  var SAFELINK_HOST = /(^|\.)safelinks\.protection\.outlook\.com$/i;
  var PROTECTOR_HOST = /(^|\.)(mimecastprotect\.com|mimecast\.com|urldefense\.com|urldefense\.proofpoint\.com|protect\.checkpoint\.com|clicktime\.symantec\.com|linkprotect\.cudasvc\.com|protection\.inkyphishfence\.com)$/i;

  function hostOf(href) {
    try { return new URL(href).hostname; } catch (e) { return ""; }
  }

  function param(href, name) {
    try { return new URL(href).searchParams.get(name) || ""; } catch (e) { return ""; }
  }

  function decodeOnce(s) {
    try { return decodeURIComponent(s); } catch (e) { return s; }
  }

  // A plain URL standing on its own, which is what an unstyled pasted link looks
  // like once a gateway has rewritten the href out from under it.
  function textUrl(s) {
    var t = flat(s).replace(/[)>.,;]+$/, "");
    if (/^https?:\/\/[^\s]+$/i.test(t)) return t;
    if (/^www\.[^\s/]+\.[A-Za-z]{2,}(\/[^\s]*)?$/.test(t)) return "https://" + t;
    return "";
  }

  function isProtectedLink(n) {
    var host = n.nodeName === "A" ? hostOf(n.getAttribute("href") || "") : "";
    return !!host && (SAFELINK_HOST.test(host) || PROTECTOR_HOST.test(host));
  }

  function realHref(n, content) {
    var raw = n.getAttribute("href") || "";
    var url = param(raw, "url") || param(raw, "u");
    if (SAFELINK_HOST.test(hostOf(raw))) return decodeOnce(url || raw);
    if (/^https?:/i.test(url)) return decodeOnce(url);
    return textUrl(content) || raw;
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
    // A GFM table is line-oriented, but Outlook wraps every cell's content in a
    // <p class=MsoNormal> and Turndown quite reasonably renders that as a blank
    // line — which tears the row apart. Flatten whatever comes back out of a
    // cell, and escape any pipe in it. Added after gfm, so it shadows the
    // plugin's own tableCell rule.
    td.addRule("tableCell", {
      filter: ["th", "td"],
      replacement: function (content, n) {
        var text = content.replace(/\|/g, "\\|").replace(/\s+/g, " ").trim();
        return (rowCells(n.parentNode)[0] === n ? "| " : " ") + text + " |";
      },
    });
    td.addRule("dropEmptyLinks", {
      filter: function (n) { return n.nodeName === "A" && !n.textContent.trim(); },
      replacement: function () { return ""; },
    });
    // Unwrap gateway-rewritten hrefs back to the real URL (Defender Safe Links,
    // Mimecast and the rest — see the link protectors section above).
    td.addRule("safelinks", {
      filter: isProtectedLink,
      replacement: function (content, n) {
        // Rules added later win, so this one shadows dropEmptyLinks; an
        // image-only protector link (a social icon) has to be dropped here too.
        var href = realHref(n, content);
        if (!n.textContent.trim()) return "";
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
    var body = cleanHtml(html, $("stripBoilerplate").checked);
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
    out = tidy(out);
    // Runs over the finished Markdown so that one pass covers frontmatter,
    // section headers, quoted header blocks and prose alike — and so that
    // "order of first appearance" means order in the document the user gets.
    if ($("stripEmails").checked) out = tidy(stripAddresses(out, nameMode()));
    return out;
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

  // ---------- settings ----------
  // The task pane is torn down and rebuilt every time it opens, so without
  // this every control would snap back to its markup default on each message.
  // roamingSettings persists per user in the mailbox, so preferences also
  // follow them to Outlook on the web and to other machines.
  //
  // Not secure storage — other services, Microsoft Graph included, can read
  // it. Only ever put UI preferences here, never message content.
  var SETTINGS_KEY = "options";

  function roaming() {
    try {
      return Office.context.roamingSettings || null;
    } catch (e) {
      return null; // host does not offer it; markup defaults stand
    }
  }

  function currentOptions() {
    return {
      frontmatter: $("frontmatter").checked,
      splitThread: $("splitThread").checked,
      stripBoilerplate: $("stripBoilerplate").checked,
      stripEmails: $("stripEmails").checked,
      nameMode: nameMode(),
    };
  }

  var BOOL_OPTS = ["frontmatter", "splitThread", "stripBoilerplate", "stripEmails"];
  var MODE_BOX = { display: "nameDisplay", first: "nameFirst", alias: "nameAlias" };

  function loadSettings() {
    var rs = roaming();
    var saved, i, k;
    if (!rs) return;
    try { saved = rs.get(SETTINGS_KEY); } catch (e) { return; }
    if (!saved || typeof saved !== "object") return;
    for (i = 0; i < BOOL_OPTS.length; i++) {
      k = BOOL_OPTS[i];
      // Only a real boolean overrides the markup, so a partial or older blob
      // leaves the remaining controls at their shipped defaults.
      if (typeof saved[k] === "boolean") $(k).checked = saved[k];
    }
    if (has(MODE_BOX, saved.nameMode)) $(MODE_BOX[saved.nameMode]).checked = true;
  }

  function saveSettings() {
    var rs = roaming();
    if (!rs) return;
    try {
      rs.set(SETTINGS_KEY, currentOptions());
      // Fire and forget: a failed write costs the user one preference next
      // time, which is not worth interrupting the conversion over.
      rs.saveAsync(function () {});
    } catch (e) { /* preferences just will not stick */ }
  }

  function optionChanged() { saveSettings(); render(false); }

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
    $("frontmatter").addEventListener("change", optionChanged);
    $("splitThread").addEventListener("change", optionChanged);
    $("stripBoilerplate").addEventListener("change", optionChanged);
    $("stripEmails").addEventListener("change", function () { syncNameMode(); optionChanged(); });
    $("nameDisplay").addEventListener("change", optionChanged);
    $("nameFirst").addEventListener("change", optionChanged);
    $("nameAlias").addEventListener("change", optionChanged);
    loadSettings();
    syncNameMode();
    applyTheme();
    showBuild();
    render(true);
  });
})();
