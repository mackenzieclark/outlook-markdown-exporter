const assert = require("assert");
const { JSDOM } = require("jsdom");
const fs = require("fs");

// Build the pane from the real taskpane.html rather than a hand-copied stub, so
// the harness exercises the shipped markup (ids, defaults, the hidden child
// group) instead of drifting from it. jsdom does not run the <script> tags.
let dom = new JSDOM(fs.readFileSync(__dirname + "/../src/taskpane.html", "utf8"));
let w = dom.window;
let doc = w.document;
for (const k of ["document", "DOMParser", "NodeFilter", "Node", "Blob", "URL"]) global[k] = w[k];
global.window = w;
global.navigator = {};
global.TurndownService = require("turndown");
global.turndownPluginGfm = require("turndown-plugin-gfm");

// ---------- fixtures ----------
// Public repo: every identity here is Alice/Bob/Carol at example.com.

// A: the synthetic Outlook/Word reply chain, single-word display names.
const A = {
  html: fs.readFileSync(__dirname + "/sample.html", "utf8"),
  item: {
    subject: "RE: Q3 Deliverables",
    from: { displayName: "Alice", emailAddress: "alice@example.com" },
    to: [{ displayName: "Bob", emailAddress: "bob@example.com" }],
    cc: [],
    dateTimeCreated: new Date("2026-08-20T15:30:00Z"),
    internetMessageId: "<abc@example.com>",
    conversationId: "AAQk",
    attachments: [{ name: "plan.xlsx", isInline: false }, { name: "image001.png", isInline: true }],
  },
};

// B: exercises the awkward cases — multi-word display names (so "Display name"
// and "First name only" differ), a bare address and a mailto: link in prose, a
// quoted header whose sender has NO address and appears *before* any addressed
// copy of that person, and two different people sharing the display name
// "Alice Smith" (which must not collide).
const B = {
  html: [
    '<html><body><div class="WordSection1">',
    '<p class="MsoNormal">Alice, please copy carol@example.com on the next one.</p>',
    '<p class="MsoNormal">Ping <a href="mailto:bob@example.com">Bob Smith</a> if it slips.</p>',
    '<div id="divRplyFwdMsg" dir="ltr"><b>From:</b> Carol Jones<br><b>Sent:</b> Monday, August 17, 2026 8:00 AM',
    '<br><b>To:</b> Bob Smith &lt;bob@example.com&gt;<br><b>Subject:</b> Budget</div>',
    '<p class="MsoNormal">Numbers attached.</p>',
    '<div style="border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in"><p class="MsoNormal">',
    '<b>From:</b> Carol Jones &lt;carol@example.com&gt;<br><b>Sent:</b> Friday, August 14, 2026 4:10 PM',
    '<br><b>To:</b> Alice Smith &lt;alice@example.com&gt;; Alice Smith &lt;alice.smith@example.com&gt;',
    '<br><b>Subject:</b> Budget</p></div>',
    '<p class="MsoNormal">First draft.</p>',
    "</div></body></html>",
  ].join("\n"),
  item: {
    subject: "Budget",
    from: { displayName: "Bob Smith", emailAddress: "bob@example.com" },
    to: [
      { displayName: "Alice Smith", emailAddress: "alice@example.com" },
      { displayName: "Carol Jones", emailAddress: "carol@example.com" },
    ],
    cc: [{ displayName: "Alice Smith", emailAddress: "alice.smith@example.com" }],
    dateTimeCreated: new Date("2026-08-18T09:00:00Z"),
    internetMessageId: "<budget@example.com>",
    conversationId: "AAQl",
    attachments: [],
  },
};

// C: the shape Outlook emits around a signature — a borderless MsoNormalTable
// nesting three deep, an icon strip of image-only links, a tracking pixel, a
// client footer and a legal disclaimer. Invented from scratch: Bob Jones of
// Example Corp does not exist, and neither does any URL below.
const C = {
  html: [
    '<html xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="utf-8">',
    "<style>p.MsoNormal{margin:0}</style></head>",
    '<body lang="EN-US"><div class="WordSection1">',
    '<p class="MsoNormal">Hi Carol,<o:p></o:p></p>',
    '<p class="MsoNormal">The Q4 plan is attached. Shout if anything looks off before Friday.</p>',
    '<p class="MsoNormal">Thanks,<br>Bob</p>',
    '<table class="MsoNormalTable x_sig" border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse">',
    '<tr><td style="padding:0in 5.4pt 0in 5.4pt">',
    '<table class="MsoNormalTable" border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse">',
    '<tr><td style="padding:0in"><p class="x_MsoNormal"><b>Bob Jones</b></p></td></tr>',
    '<tr><td style="padding:0in"><p class="x_MsoNormal">Widget Lead, Example Corp</p></td></tr>',
    '<tr><td style="padding:0in"><p class="x_MsoNormal">+1 555 0100</p></td></tr>',
    '<tr><td style="padding:0in"><p class="x_MsoNormal"><a href="mailto:bob@example.com">bob@example.com</a></p></td></tr>',
    '<tr><td style="padding:0in"><p class="x_MsoNormal"><a href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fwww.example.com%2F&amp;data=05%7C01">www.example.com</a></p></td></tr>',
    "</table>",
    '</td><td style="padding:0in 5.4pt 0in 5.4pt">',
    '<table class="MsoNormalTable" border="0"><tr>',
    '<td><a href="https://protect-eu.mimecastprotect.com/s/AbC1"><img src="https://cdn.example.com/x_a.png" width="24" height="24" alt=""></a></td>',
    '<td><a href="https://protect-eu.mimecastprotect.com/s/DeF2"><img src="https://cdn.example.com/x_b.png" width="24" height="24" alt=""></a></td>',
    "</tr></table>",
    "</td></tr></table>",
    '<p class="MsoNormal">Sent from my iPhone</p>',
    '<div style="font-size:8.0pt;color:#888888">',
    '<p class="MsoNormal">CONFIDENTIALITY NOTICE: this e-mail and any attachments are for the named recipient only.</p>',
    '<p class="MsoNormal">If you have received this in error, please notify the sender and delete it.</p>',
    "</div>",
    '<img src="https://track.example.com/o.gif" width="1" height="1" style="width:1.0pt;height:1.0pt">',
    "</div></body></html>",
  ].join("\n"),
  item: {
    subject: "Q4 plan",
    from: { displayName: "Bob Jones", emailAddress: "bob@example.com" },
    to: [{ displayName: "Carol Jones", emailAddress: "carol@example.com" }],
    cc: [],
    dateTimeCreated: new Date("2026-08-21T11:00:00Z"),
    internetMessageId: "<q4@example.com>",
    conversationId: "AAQm",
    attachments: [],
  },
};

// D: tables that ARE data, plus the two protector shapes. Table one has a bold
// first row (a header Outlook wrote as <td><b>); table two is a label/value
// grid with no header at all, which must keep every one of its rows.
const D = {
  html: [
    '<html><body><div class="WordSection1">',
    '<p class="MsoNormal">Status below.</p>',
    '<table class="MsoNormalTable" border="0" cellpadding="0" cellspacing="0" style="border-collapse:collapse">',
    '<tr><td><p class="MsoNormal"><b>Env</b></p></td><td><p class="MsoNormal"><b>Status</b></p></td></tr>',
    '<tr><td><p class="MsoNormal">staging</p></td><td><p class="MsoNormal">green</p></td></tr>',
    '<tr><td><p class="MsoNormal">prod</p></td><td><p class="MsoNormal">amber</p></td></tr>',
    "</table>",
    '<p class="MsoNormal">And the ticket fields:</p>',
    '<table border="0" style="width:400pt">',
    "<tr><td>Owner</td><td>Alice</td></tr>",
    "<tr><td>Due</td><td>2026-09-01</td></tr>",
    "<tr><td>Priority</td><td>High</td></tr>",
    "</table>",
    '<p class="MsoNormal">Runbook: <a href="https://protect-eu.mimecastprotect.com/s/Zz9?domain=example.com">https://docs.example.com/runbook</a></p>',
    '<p class="MsoNormal">Plan: <a href="https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fexample.com%2Fplan&amp;data=05%7C01">the plan</a></p>',
    '<p class="MsoNormal">Wiki: <a href="https://protect-eu.mimecastprotect.com/s/Yy8?domain=example.com">the wiki page</a></p>',
    "</div></body></html>",
  ].join("\n"),
  item: {
    subject: "Release status",
    from: { displayName: "Alice", emailAddress: "alice@example.com" },
    to: [{ displayName: "Bob", emailAddress: "bob@example.com" }],
    cc: [],
    dateTimeCreated: new Date("2026-08-22T08:00:00Z"),
    internetMessageId: "<rel@example.com>",
    conversationId: "AAQn",
    attachments: [],
  },
};

let fixture = A;
const clone = (v) => (v === undefined ? undefined : JSON.parse(JSON.stringify(v)));
let roamingStored = {};   // what has actually been persisted
let roamingDraft = {};    // what set() has staged but not yet saved
let roamingSaves = 0;

global.Office = {
  onReady: (cb) => cb(),
  CoercionType: { Html: "html" },
  AsyncResultStatus: { Succeeded: "succeeded" },
  context: {
    roamingSettings: {
      get: (k) => clone(roamingStored[k]),
      set: (k, v) => { roamingDraft[k] = clone(v); },
      remove: (k) => { delete roamingDraft[k]; },
      saveAsync: (cb) => {
        Object.assign(roamingStored, roamingDraft);
        roamingSaves += 1;
        if (cb) cb({ status: "succeeded" });
      },
    },
    mailbox: {
      get item() {
        return Object.assign({}, fixture.item, {
          body: { getAsync: (t, cb) => cb({ status: "succeeded", value: fixture.html }) },
        });
      },
    },
  },
};

require("../src/taskpane.js");

// The first render happens inside Office.onReady, with the pane's own defaults.
const DEFAULT_OUT = doc.getElementById("out").value;
const DEFAULT_STATUS = doc.getElementById("status").textContent;

// ---------- driver ----------
const MODE_ID = { display: "nameDisplay", first: "nameFirst", alias: "nameAlias" };

// Close and reopen the pane: a fresh DOM from the shipped markup and a fresh
// evaluation of taskpane.js. roamingSettings is read once, at load, so this is
// the only way to prove a preference actually survives.
function reopen() {
  dom = new JSDOM(fs.readFileSync(__dirname + "/../src/taskpane.html", "utf8"));
  w = dom.window;
  doc = w.document;
  for (const k of ["document", "DOMParser", "NodeFilter", "Node", "Blob", "URL"]) global[k] = w[k];
  global.window = w;
  delete require.cache[require.resolve("../src/taskpane.js")];
  require("../src/taskpane.js");
  return doc;
}
const boxes = () => ({
  frontmatter: doc.getElementById("frontmatter").checked,
  splitThread: doc.getElementById("splitThread").checked,
  stripBoilerplate: doc.getElementById("stripBoilerplate").checked,
  stripEmails: doc.getElementById("stripEmails").checked,
  nameMode: doc.getElementById("nameAlias").checked ? "alias"
    : doc.getElementById("nameFirst").checked ? "first" : "display",
});

// Drives the pane the way a user does: flip the controls, fire "change", let the
// add-in's own listeners re-render.
function run(opts) {
  opts = opts || {};
  fixture = opts.fixture || A;
  doc.getElementById("frontmatter").checked = opts.frontmatter !== false;
  doc.getElementById("splitThread").checked = opts.splitThread !== false;
  doc.getElementById("stripBoilerplate").checked = opts.boiler !== false;
  doc.getElementById("stripEmails").checked = !!opts.strip;
  Object.keys(MODE_ID).forEach((m) => { doc.getElementById(MODE_ID[m]).checked = false; });
  doc.getElementById(MODE_ID[opts.mode || "display"]).checked = true;
  doc.getElementById("stripEmails").dispatchEvent(new w.Event("change"));
  return doc.getElementById("out").value;
}

// message_id is an address in shape only and is deliberately left alone, so it
// is excluded when asserting that no addresses survive.
const withoutIds = (s) => s.replace(/^(message_id|conversation_id):.*$/gm, "");

let checks = 0;
function check(name, fn) {
  fn();
  checks += 1;
  console.log("  ok  " + name);
}

function show(title, md) {
  console.log("\n=== " + title + " " + "=".repeat(Math.max(0, 58 - title.length)));
  console.log(md.trimEnd());
}

// ---------- output ----------
show("A: defaults (strip off) — unchanged behaviour", DEFAULT_OUT);
console.log("\nSTATUS: " + DEFAULT_STATUS);
console.log("BUILD: " + doc.getElementById("version").textContent);

const aDisplay = run({ strip: true, mode: "display" });
const aFirst = run({ strip: true, mode: "first" });
const aAlias = run({ strip: true, mode: "alias" });
const bDisplay = run({ fixture: B, strip: true, mode: "display" });
const bFirst = run({ fixture: B, strip: true, mode: "first" });
const bAlias = run({ fixture: B, strip: true, mode: "alias" });

show("A: strip on — Alias all names", aAlias);
show("B: strip on — Display name", bDisplay);
show("B: strip on — First name only", bFirst);
show("B: strip on — Alias all names", bAlias);

const cKeep = run({ fixture: C, boiler: false });
const cStrip = run({ fixture: C });
const cKeepDisplay = run({ fixture: C, boiler: false, strip: true, mode: "display" });
const cStripAlias = run({ fixture: C, strip: true, mode: "alias" });
const dOut = run({ fixture: D });

show("C: nested signature table — boilerplate strip OFF (Part 1 only)", cKeep);
show("C: boilerplate strip ON — the shipped default", cStrip);
show("D: data tables and link protectors", dOut);

// Nothing Turndown kept as raw HTML may reach the Markdown. `keep` emits the
// element's outerHTML, so a leaked table brings its inline styles with it —
// which makes style= the cheapest tripwire for the whole class of bug.
const RAW_HTML = /<\/?(table|thead|tbody|tfoot|tr|td|th)\b|style=/i;
function noRawHtml(md, where) {
  const m = RAW_HTML.exec(md);
  assert.ok(!m, "raw HTML leaked into " + where + ": " + (m && md.slice(m.index, m.index + 60)));
}

// ---------- tests ----------
console.log("\n=== tests " + "=".repeat(53));

console.log("\ncontrols");
check("child radio group is hidden while the parent checkbox is off", () => {
  run({});
  assert.strictEqual(doc.getElementById("nameModeGroup").hidden, true);
});
check("child radio group is revealed when the parent checkbox is on", () => {
  run({ strip: true, mode: "display" });
  assert.strictEqual(doc.getElementById("nameModeGroup").hidden, false);
});
check("pane ships with strip off and Display name preselected", () => {
  const fresh = new JSDOM(fs.readFileSync(__dirname + "/../src/taskpane.html", "utf8")).window.document;
  assert.strictEqual(fresh.getElementById("stripEmails").checked, false);
  assert.strictEqual(fresh.getElementById("nameDisplay").checked, true);
  assert.strictEqual(fresh.getElementById("nameFirst").checked, false);
  assert.strictEqual(fresh.getElementById("nameAlias").checked, false);
  assert.ok(fresh.getElementById("nameModeGroup").hasAttribute("hidden"));
});

console.log("\ndefault behaviour is untouched");
check("addresses survive with the checkbox off", () => {
  const off = run({});
  assert.strictEqual(off, DEFAULT_OUT);
  assert.ok(off.includes('from: "Alice <alice@example.com>"'));
  assert.ok(off.includes('  - "Bob <bob@example.com>"'));
  assert.ok(off.includes("## Bob <bob@example.com> — Tuesday, August 19, 2026 3:02 PM"));
  assert.ok(off.includes("**From:** Bob <bob@example.com>"));
});

console.log("\nalias stability (regression)");
// These are the assertions that break if aliases are ever handed out by hashing,
// by object key order, or keyed on the display name instead of the address.
check("frontmatter alias matches the same person's thread section header", () => {
  const from = /^from: "(User\d+)"$/m.exec(aAlias);
  const to = /^ {2}- "(User\d+)"$/m.exec(aAlias);
  assert.ok(from && to, "frontmatter should be aliased");
  // Alice sent the top message and the Monday quote; Bob sent the Tuesday one.
  assert.strictEqual(/^## (User\d+) — Tuesday/m.exec(aAlias)[1], to[1]);
  assert.strictEqual(/^## (User\d+) — Monday/m.exec(aAlias)[1], from[1]);
  assert.notStrictEqual(from[1], to[1], "two addresses must never share an alias");
});
check("the same address is the same alias in every quoted header block", () => {
  const bob = /^from: "(User\d+)"$/m.exec(bAlias)[1];
  assert.strictEqual(/\*\*To:\*\* (User\d+)$/m.exec(bAlias)[1], bob);
  assert.ok(bAlias.includes("Ping " + bob + " if it slips."), "mailto: link must use the same alias");
});
check("a name-only sender reuses the alias of the addressed copy", () => {
  // "**From:** Carol Jones" (no address) appears BEFORE "Carol Jones
  // <carol@example.com>" in the document, so a naive one-pass scan would mint a
  // second alias for her.
  const nameOnly = /^\*\*From:\*\* (User\d+)$/m.exec(bAlias)[1];
  const addressed = /^## (User\d+) — Friday/m.exec(bAlias)[1];
  const fmAliases = bAlias.match(/^ {2}- "(User\d+)"$/gm).map((s) => /User\d+/.exec(s)[0]);
  assert.strictEqual(nameOnly, addressed);
  assert.ok(fmAliases.includes(nameOnly), "and it is the alias she got in the frontmatter");
  assert.ok(bAlias.includes("please copy " + nameOnly + " on the next one."));
});
check("two people sharing a display name get different aliases", () => {
  const pair = /\*\*To:\*\* (User\d+); (User\d+)/.exec(bAlias);
  assert.ok(pair, "both Alice Smiths should be aliased");
  assert.notStrictEqual(pair[1], pair[2]);
});
check("exactly four distinct people are found in fixture B", () => {
  assert.strictEqual(new Set(bAlias.match(/User\d+/g)).size, 4);
});
check("aliases are numbered by first appearance, not by hash or key order", () => {
  assert.ok(bAlias.includes('from: "User1"'));
  assert.ok(bAlias.includes('  - "User2"'));
  assert.ok(bAlias.includes('  - "User3"'));
  assert.ok(bAlias.includes('  - "User4"'));
});
check("repeat conversions are byte-identical", () => {
  assert.strictEqual(run({ fixture: B, strip: true, mode: "alias" }), bAlias);
  assert.strictEqual(run({ fixture: B, strip: true, mode: "alias" }), bAlias);
  assert.strictEqual(run({ strip: true, mode: "alias" }), aAlias);
});

console.log("\nDisplay name");
check("frontmatter, headers and quoted blocks keep the full name", () => {
  assert.ok(bDisplay.includes('from: "Bob Smith"'));
  assert.ok(bDisplay.includes('  - "Alice Smith"'));
  assert.ok(bDisplay.includes('  - "Carol Jones"'));
  assert.ok(bDisplay.includes("## Carol Jones — Monday, August 17, 2026 8:00 AM"));
  assert.ok(bDisplay.includes("**To:** Bob Smith"));
  assert.ok(bDisplay.includes("**To:** Alice Smith; Alice Smith"));
});
check("mailto link and bare address become the person's name", () => {
  assert.ok(bDisplay.includes("Ping Bob Smith if it slips."));
  assert.ok(bDisplay.includes("please copy Carol Jones on the next one."));
});
check("no address survives", () => assert.ok(!/@example\.com/.test(withoutIds(bDisplay))));

console.log("\nFirst name only");
check("Alice Smith -> Alice everywhere", () => {
  assert.ok(bFirst.includes('from: "Bob"'));
  assert.ok(bFirst.includes('  - "Alice"'));
  assert.ok(bFirst.includes('  - "Carol"'));
  assert.ok(bFirst.includes("## Carol — Monday, August 17, 2026 8:00 AM"));
  assert.ok(bFirst.includes("**To:** Alice; Alice"));
  assert.ok(bFirst.includes("Ping Bob if it slips."));
  assert.ok(bFirst.includes("please copy Carol on the next one."));
});
check("no address survives", () => assert.ok(!/@example\.com/.test(withoutIds(bFirst))));
check("single-word names are unaffected by the split", () => {
  assert.ok(aFirst.includes('from: "Alice"'));
  assert.ok(aFirst.includes("## Bob — Tuesday, August 19, 2026 3:02 PM"));
});

console.log("\nAlias all names");
check("no address survives, and every person site is aliased", () => {
  assert.ok(!/@example\.com/.test(withoutIds(bAlias)));
  const sites = bAlias.split("\n").filter((l) =>
    /^(from|to|cc):/.test(l) || /^ {2}- "/.test(l) || /^## /.test(l) ||
    /^\*\*(From|To|Cc):\*\*/.test(l));
  assert.ok(sites.length >= 10, "expected to find the person sites, got " + sites.length);
  sites.forEach((l) => assert.ok(!/Alice|Bob|Carol|Smith|Jones/.test(l), "unaliased name in: " + l));
});
check("names in prose are left alone — documented limitation", () => {
  // Only addresses and the person fields are rewritten; replacing every word
  // that happens to be a name in running text would be guesswork.
  assert.ok(bAlias.includes("Alice, please copy User3 on the next one."));
});
check("Display name mode does NOT alias", () => {
  assert.ok(!/User\d/.test(bDisplay));
  assert.ok(!/User\d/.test(bFirst));
});

console.log("\ninteraction with the other toggles");
check("works with frontmatter off", () => {
  const md = run({ fixture: B, strip: true, mode: "alias", frontmatter: false });
  assert.ok(!md.includes("---\nsubject:"));
  assert.ok(!/@example\.com/.test(md));
  assert.ok(/^## User\d+ — Monday/m.test(md));
});
check("works with thread splitting off", () => {
  const md = run({ fixture: B, strip: true, mode: "alias", splitThread: false });
  assert.ok(!md.includes("\n## "));
  assert.ok(!/@example\.com/.test(withoutIds(md)));
  assert.ok(md.includes('from: "User1"'));
});
check("message_id is left alone — it is not a person", () => {
  assert.ok(bAlias.includes('message_id: "<budget@example.com>"'));
  assert.ok(bDisplay.includes('message_id: "<budget@example.com>"'));
});

console.log("\ntables never reach the Markdown as raw HTML");
check("no fixture leaks a <table>, a cell, or an inline style, either way up", () => {
  [A, B, C, D].forEach((f, i) => {
    const name = "ABCD"[i];
    noRawHtml(run({ fixture: f }), name + " (strip on)");
    noRawHtml(run({ fixture: f, boiler: false }), name + " (strip off)");
    noRawHtml(run({ fixture: f, boiler: false, strip: true, mode: "alias" }), name + " (aliased)");
  });
});
check("a signature's nested layout tables are unwrapped, not dumped", () => {
  // Same fixture, strip OFF: the text has to survive as ordinary blocks.
  assert.ok(cKeep.includes("Bob Jones"));
  assert.ok(cKeep.includes("Widget Lead, Example Corp"));
  assert.ok(cKeep.includes("+1 555 0100"));
  assert.ok(!cKeep.includes("MsoNormalTable"));
});
check("Turndown's own rules now run inside what used to be a kept table", () => {
  // Both of these live inside the signature table, so before the fix Turndown
  // never descended far enough to see either one.
  assert.ok(cKeep.includes("[www.example.com](https://www.example.com/)"), "safelink unwrapped");
  assert.ok(!/mimecastprotect/.test(cKeep), "image-only icon links dropped");
});
check("a bold first row is promoted to a real header row", () => {
  assert.ok(dOut.includes("| Env | Status |\n| --- | --- |\n| staging | green |"));
  assert.ok(dOut.includes("| prod | amber |"));
  assert.ok(!dOut.includes("**Env**"), "the bold is the header now, not emphasis");
});
check("a label/value table keeps every row under a synthesized header", () => {
  assert.ok(dOut.includes("|  |  |\n| --- | --- |\n| Owner | Alice |"));
  assert.ok(dOut.includes("| Due | 2026-09-01 |"));
  assert.ok(dOut.includes("| Priority | High |"), "no data row may be eaten as a header");
});
check("a table that already had <th> converts exactly as before", () => {
  assert.ok(DEFAULT_OUT.includes("| Env | Status |\n| --- | --- |\n| prod | green |"));
});

console.log("\nlink protectors");
check("Defender safe links still unwrap to the real URL", () => {
  assert.ok(DEFAULT_OUT.includes("[the doc](https://example.sharepoint.com/sites/x/doc.docx)"));
  assert.ok(dOut.includes("[the plan](https://example.com/plan)"));
});
check("a protector that hides the URL falls back to the anchor text", () => {
  // Mimecast hashes the target into the path; the visible text is all there is.
  assert.ok(dOut.includes("<https://docs.example.com/runbook>"));
});
check("a protected link whose text is not a URL keeps the rewritten href", () => {
  assert.ok(dOut.includes("[the wiki page](https://protect-eu.mimecastprotect.com/s/Yy8?domain=example.com)"));
});

console.log("\nstrip signatures and boilerplate");
check("the checkbox ships checked, between split and strip-emails", () => {
  const fresh = new JSDOM(fs.readFileSync(__dirname + "/../src/taskpane.html", "utf8")).window.document;
  const ids = [...fresh.querySelectorAll("body > label > input")].map((i) => i.id);
  assert.deepStrictEqual(ids, ["frontmatter", "splitThread", "stripBoilerplate", "stripEmails"]);
  assert.strictEqual(fresh.getElementById("stripBoilerplate").checked, true);
  // The radio group must still hang off stripEmails, not off the new checkbox.
  const group = fresh.getElementById("nameModeGroup");
  assert.strictEqual(group.previousElementSibling.querySelector("input").id, "stripEmails");
});
check("the signature block goes", () => {
  const body = cStrip.slice(cStrip.indexOf("# Q4 plan"));
  assert.ok(!body.includes("Widget Lead"));
  assert.ok(!body.includes("+1 555 0100"));
  assert.ok(!body.includes("www.example.com"));
  assert.ok(!body.includes("bob@example.com"));
});
check("client footers go", () => assert.ok(!cStrip.includes("Sent from my iPhone")));
check("confidentiality disclaimers go, every paragraph of them", () => {
  assert.ok(!/CONFIDENTIALITY NOTICE/i.test(cStrip));
  assert.ok(!/named recipient/i.test(cStrip));
  assert.ok(!/received this in error/i.test(cStrip));
});
check("tracking pixels, spacers and social icon rows go", () => {
  assert.ok(!cStrip.includes("track.example.com"), "1x1 pixel");
  assert.ok(!cStrip.includes("cdn.example.com"), "icon images");
  assert.ok(!/mimecastprotect/.test(cStrip), "image-only icon links");
});
check("prose and the sign-off survive", () => {
  assert.ok(cStrip.includes("Hi Carol,"));
  assert.ok(cStrip.includes("The Q4 plan is attached. Shout if anything looks off before Friday."));
  assert.ok(/Thanks,\nBob/.test(cStrip), "a sign-off is not a signature block");
});
check("with the checkbox off the signature stays, merely unwrapped", () => {
  assert.ok(cKeep.includes("Widget Lead, Example Corp"));
  assert.ok(cKeep.includes("Sent from my iPhone"));
  assert.ok(/CONFIDENTIALITY NOTICE/.test(cKeep));
});
check("real data tables and real links are left alone by the strip", () => {
  assert.strictEqual(run({ fixture: D }), run({ fixture: D, boiler: false }));
  assert.ok(dOut.includes("| Owner | Alice |"));
  assert.ok(dOut.includes("[the plan](https://example.com/plan)"));
});
check("stripping boilerplate does not disturb the reply-split markers", () => {
  const md = run({ fixture: A });
  assert.strictEqual(md, DEFAULT_OUT);
  assert.ok(md.includes("## Bob <bob@example.com> — Tuesday, August 19, 2026 3:02 PM"));
  assert.ok(md.includes("## Alice — Monday, August 18, 2026 9:00 AM"));
});

console.log("\nboilerplate x strip-email-addresses");
check("with both on, the signature is gone and no address is left to strip", () => {
  assert.ok(!/@example\.com/.test(withoutIds(cStripAlias)));
  assert.ok(!cStripAlias.includes("Widget Lead"));
  assert.ok(cStripAlias.includes('from: "User1"'));
  assert.strictEqual(new Set(cStripAlias.match(/User\d+/g)).size, 2);
});
check("with boilerplate kept, the signature's address is still rewritten", () => {
  assert.ok(cKeepDisplay.includes("Widget Lead, Example Corp"), "signature kept");
  assert.ok(!/@example\.com/.test(withoutIds(cKeepDisplay)), "but its address is not");
  assert.ok(cKeepDisplay.includes("Bob Jones"));
});
check("the two checkboxes are independent", () => {
  const alias = run({ fixture: C, boiler: false, strip: true, mode: "alias" });
  assert.ok(alias.includes("Widget Lead, Example Corp"));
  assert.ok(!/@example\.com/.test(withoutIds(alias)));
  assert.ok(/User\d+/.test(alias));
});
check("repeat conversions stay byte-identical with the new toggle", () => {
  assert.strictEqual(run({ fixture: C }), cStrip);
  assert.strictEqual(run({ fixture: C, boiler: false }), cKeep);
});


console.log("\nsettings persistence");
check("changing an option writes it through to storage", () => {
  const before = roamingSaves;
  run({ fixture: A, frontmatter: false, boiler: false, strip: true, mode: "alias" });
  assert.ok(roamingSaves > before, "saveAsync was called");
  assert.deepStrictEqual(roamingStored.options, {
    frontmatter: false,
    splitThread: true,
    stripBoilerplate: false,
    stripEmails: true,
    nameMode: "alias",
  });
});
check("set() alone does not persist — only saveAsync commits", () => {
  const stored = clone(roamingStored.options);
  Office.context.roamingSettings.set("options", {
    frontmatter: true, splitThread: true, stripBoilerplate: true,
    stripEmails: false, nameMode: "display",
  });
  assert.deepStrictEqual(roamingStored.options, stored, "staged, not written");
});
check("every option survives closing and reopening the pane", () => {
  reopen();
  assert.deepStrictEqual(boxes(), {
    frontmatter: false,
    splitThread: true,
    stripBoilerplate: false,
    stripEmails: true,
    nameMode: "alias",
  });
});
check("the reopened pane renders with the restored options, not the defaults", () => {
  const out = doc.getElementById("out").value;
  assert.ok(!out.startsWith("---"), "frontmatter stayed off");
  assert.ok(/User\d+/.test(out), "alias mode stayed on");
});
check("the child radio group is re-revealed to match a restored checkbox", () => {
  assert.strictEqual(doc.getElementById("nameModeGroup").hidden, false);
});
check("with nothing stored, the shipped markup defaults win", () => {
  roamingStored = {};
  roamingDraft = {};
  reopen();
  assert.deepStrictEqual(boxes(), {
    frontmatter: true,
    splitThread: true,
    stripBoilerplate: true,
    stripEmails: false,
    nameMode: "display",
  });
  assert.strictEqual(doc.getElementById("nameModeGroup").hidden, true);
});
check("a partial or older blob leaves unknown options at their defaults", () => {
  roamingStored = { options: { stripEmails: true, nameMode: "first" } };
  reopen();
  const st = boxes();
  assert.strictEqual(st.stripEmails, true, "what was stored is applied");
  assert.strictEqual(st.nameMode, "first");
  assert.strictEqual(st.frontmatter, true, "what was not stored keeps its default");
  assert.strictEqual(st.stripBoilerplate, true);
});
check("a corrupt blob is ignored rather than throwing", () => {
  roamingStored = { options: "not an object" };
  assert.doesNotThrow(reopen);
  assert.strictEqual(boxes().frontmatter, true);
  roamingStored = { options: { frontmatter: "yes please", nameMode: "nonsense" } };
  assert.doesNotThrow(reopen);
  assert.strictEqual(boxes().frontmatter, true, "a non-boolean is not applied");
  assert.strictEqual(boxes().nameMode, "display", "an unknown mode is not applied");
});


// E: the same person in two quoted headers — once as a plain address, once
// linkified by Outlook into [addr](mailto:addr). Both must strip identically;
// the linkified form used to leave the display name stranded outside the match,
// leaking real names into alias mode.
const E = {
  html: [
    '<html><body><div class="WordSection1">',
    '<p class="MsoNormal">Body text.</p>',
    '<div id="divRplyFwdMsg" dir="ltr">',
    '<b>From:</b> Alice Smith &lt;alice@example.com&gt;<br>',
    '<b>Sent:</b> Monday, August 17, 2026 8:00 AM<br>',
    '<b>To:</b> Bob Jones &lt;bob@example.com&gt;<br>',
    '<b>Subject:</b> Budget</div>',
    '<p class="MsoNormal">First reply.</p>',
    '<div style="border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in"><p class="MsoNormal">',
    '<b>From:</b> Alice Smith &lt;<a href="mailto:alice@example.com">alice@example.com</a>&gt;<br>',
    '<b>Sent:</b> Friday, August 14, 2026 4:10 PM<br>',
    '<b>To:</b> Bob Jones &lt;<a href="mailto:bob@example.com">bob@example.com</a>&gt;<br>',
    '<b>Subject:</b> Budget</p></div>',
    '<p class="MsoNormal">Original.</p>',
    "</div></body></html>",
  ].join("\n"),
  item: {
    subject: "Budget",
    from: { displayName: "Alice Smith", emailAddress: "alice@example.com" },
    to: [{ displayName: "Bob Jones", emailAddress: "bob@example.com" }],
    cc: [],
    dateTimeCreated: new Date("2026-08-20T15:30:00Z"),
    internetMessageId: "<abc@example.com>",
    conversationId: "AAQk",
    attachments: [],
  },
};

// F: automated mail with no To: recipients at all — the empty-sequence case.
const F = {
  html: '<html><body><p class="MsoNormal">Automated notice.</p></body></html>',
  item: {
    subject: "Notice",
    from: { displayName: "Alice", emailAddress: "noreply@example.com" },
    to: [],
    cc: [],
    dateTimeCreated: new Date("2026-08-20T15:30:00Z"),
    internetMessageId: "<n@example.com>",
    conversationId: "AAQk",
    attachments: [],
  },
};

console.log("\nlinkified addresses in quoted headers");
check("a linkified header strips the same as a plain one", () => {
  const out = run({ fixture: E, strip: true, mode: "alias", boiler: false });
  assert.ok(!/Alice Smith|Bob Jones/.test(withoutIds(out)), "no display name survives");
  assert.ok(!/@example\.com/.test(withoutIds(out)), "no address survives");
});
check("one person is one alias across both header forms", () => {
  const out = withoutIds(run({ fixture: E, strip: true, mode: "alias", boiler: false }));
  assert.strictEqual(new Set(out.match(/User\d+/g)).size, 2, "exactly two people");
  // Every **From:** in this fixture is the same person, so they must agree.
  const froms = out.match(/\*\*From:\*\* (\S+)/g);
  assert.strictEqual(new Set(froms).size, 1, "same sender, same alias: " + froms);
});
check("display mode keeps the name and drops only the address", () => {
  const out = run({ fixture: E, strip: true, mode: "display", boiler: false });
  assert.ok(out.includes("Alice Smith"), "name kept");
  assert.ok(!/@example\.com/.test(withoutIds(out)), "address gone");
  assert.ok(!/\]\(mailto:/.test(out), "no mailto link left behind");
});
check("a mailto link in prose is still left as a link", () => {
  const out = run({ fixture: B, boiler: false });
  assert.ok(/\[Bob Smith\]\(mailto:bob@example\.com\)/.test(out), "prose link intact");
});

console.log("\nfrontmatter is valid YAML");
check("an empty recipient list keeps a space after the colon", () => {
  const out = run({ fixture: F });
  assert.ok(/^to: \[\]$/m.test(out), "got: " + (out.match(/^to:.*$/m) || [])[0]);
  assert.ok(!/^to:\[\]$/m.test(out), "to:[] is not a mapping entry and breaks the block");
});
check("every frontmatter line parses as a mapping entry", () => {
  const out = run({ fixture: F });
  const fm = out.split("---")[1].trim().split("\n");
  fm.forEach((line) => {
    if (/^\s*-\s/.test(line)) return; // sequence item
    assert.ok(/^[A-Za-z_][A-Za-z0-9_]*: \S/.test(line), "not a mapping entry: " + line);
  });
});


// G: the boundary that carries no marker of its own — a bare <hr> and then the
// quoted header block, no id, no class, no border-top anywhere. Outlook mobile
// and most gateways write the chain this way, and it is the single most common
// shape in real threaded mail. Also carries a decorative <hr> inside the newest
// message, which must NOT split, and a Spanish header block for the locale
// table.
const G = {
  html: [
    '<html><body><div class="WordSection1">',
    '<p class="MsoNormal">Latest reply, and then the checklist:</p>',
    "<hr>",
    '<p class="MsoNormal">Step one, then step two.</p>',
    "<hr>",
    "<div><b>From:</b> Bob &lt;bob@example.com&gt;<br><b>Sent:</b> Tuesday, August 19, 2026 3:02 PM",
    "<br><b>To:</b> Alice &lt;alice@example.com&gt;<br><b>Subject:</b> Q3 Deliverables</div>",
    '<p class="MsoNormal">Second message body.</p>',
    "<hr>",
    "<div><b>De:</b> Carol &lt;carol@example.com&gt;<br><b>Enviado:</b> lunes, 17 de agosto de 2026 8:00",
    "<br><b>Para:</b> Bob &lt;bob@example.com&gt;<br><b>Asunto:</b> Presupuesto</div>",
    '<p class="MsoNormal">Tercer mensaje.</p>',
    "</div></body></html>",
  ].join("\n"),
  item: {
    subject: "RE: Q3 Deliverables",
    from: { displayName: "Alice", emailAddress: "alice@example.com" },
    to: [{ displayName: "Bob", emailAddress: "bob@example.com" }],
    cc: [],
    dateTimeCreated: new Date("2026-08-20T15:30:00Z"),
    internetMessageId: "<hr@example.com>",
    conversationId: "AAQo",
    attachments: [],
  },
};

// H: a Gmail chain. Every quoted message sits INSIDE the one that quotes it, so
// four messages means four levels of <blockquote class="gmail_quote"> — a real
// export reached twenty-one. Each level is a <div class="gmail_quote"> holding
// the "On … wrote:" attribution and then the blockquote itself; the two are one
// boundary, but the level below is a separate message and must get its own
// section.
const H = {
  html: [
    '<html><body><div dir="ltr">Newest note, from Alice.</div>',
    '<div class="gmail_quote gmail_quote_container">',
    '<div dir="ltr" class="gmail_attr">On Thu, Aug 20, 2026 at 9:15 AM Bob &lt;bob@example.com&gt; wrote:<br></div>',
    '<blockquote class="gmail_quote" style="margin:0 0 0 .8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">',
    '<div dir="ltr">Bob answers the second question.</div>',
    '<div class="gmail_quote gmail_quote_container">',
    '<div dir="ltr" class="gmail_attr">On Wed, Aug 19, 2026 at 4:40 PM Carol &lt;carol@example.com&gt; wrote:<br></div>',
    '<blockquote class="gmail_quote" style="margin:0 0 0 .8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">',
    '<div dir="ltr">Carol asks the second question.</div>',
    // This level has NO container div — just the attribution and then the
    // blockquote, which is what is left after another client re-emits a Gmail
    // quote. div.gmail_quote does not match it and neither does type='cite'.
    '<div dir="ltr">On Tue, Aug 18, 2026 at 11:05 AM Alice &lt;alice@example.com&gt; wrote:</div>',
    '<blockquote class="gmail_quote" style="margin:0 0 0 .8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">',
    '<div dir="ltr">Alice answers the first question.</div>',
    '<div class="gmail_quote gmail_quote_container">',
    '<div dir="ltr" class="gmail_attr">On Mon, Aug 17, 2026 at 8:00 AM Bob &lt;bob@example.com&gt; wrote:<br></div>',
    '<blockquote class="gmail_quote" style="margin:0 0 0 .8ex;border-left:1px solid rgb(204,204,204);padding-left:1ex">',
    '<div dir="ltr">Bob asks the first question.</div>',
    "</blockquote></div>",
    "</blockquote>",
    "</blockquote></div>",
    "</blockquote></div>",
    "</body></html>",
  ].join("\n"),
  item: {
    subject: "Re: Two questions",
    from: { displayName: "Alice", emailAddress: "alice@example.com" },
    to: [{ displayName: "Bob", emailAddress: "bob@example.com" }],
    cc: [{ displayName: "Carol", emailAddress: "carol@example.com" }],
    dateTimeCreated: new Date("2026-08-21T07:00:00Z"),
    internetMessageId: "<gm@example.com>",
    conversationId: "AAQp",
    attachments: [],
  },
};

// I: Apple Mail. Same nesting as Gmail, but the quote is a bare
// <blockquote type="cite"> and the attribution is a SIBLING in front of it
// rather than a wrapper around it — so the cut has to be made before the
// attribution, or the only header the quote has ends up at the foot of the
// previous section.
const I = {
  html: [
    "<html><body><div>Thanks, that works for me.</div>",
    "<div><br></div>",
    "<div>On Aug 20, 2026, at 9:15 AM, Bob &lt;bob@example.com&gt; wrote:</div>",
    '<blockquote type="cite">',
    "<div>Bob confirms the date.</div>",
    "<div>On Aug 19, 2026, at 4:40 PM, Carol &lt;carol@example.com&gt; wrote:</div>",
    '<blockquote type="cite"><div>Carol proposes the date.</div></blockquote>',
    "</blockquote>",
    "</body></html>",
  ].join("\n"),
  item: {
    subject: "Re: Kickoff date",
    from: { displayName: "Alice", emailAddress: "alice@example.com" },
    to: [{ displayName: "Bob", emailAddress: "bob@example.com" }],
    cc: [],
    dateTimeCreated: new Date("2026-08-21T09:00:00Z"),
    internetMessageId: "<ap@example.com>",
    conversationId: "AAQq",
    attachments: [],
  },
};

const gOut = run({ fixture: G });
const hOut = run({ fixture: H });
const iOut = run({ fixture: I });

show("G: bare <hr> boundaries, one of them Spanish", gOut);
show("H: a four-deep Gmail chain", hOut);
show("I: Apple Mail, attribution outside the blockquote", iOut);

// Section headings, in order, for a converted message.
const headings = (md) => (md.match(/^## .*$/gm) || []).map((s) => s.slice(3));
// Everything above the first section heading: the message the user has open.
const newest = (md) => md.split(/^## /m)[0];

console.log("\nbare <hr> boundaries");
check("an <hr> followed by a header block is a boundary", () => {
  assert.deepStrictEqual(headings(gOut), [
    "Bob <bob@example.com> — Tuesday, August 19, 2026 3:02 PM",
    "Carol <carol@example.com> — lunes, 17 de agosto de 2026 8:00",
  ]);
  assert.ok(gOut.includes("Second message body."));
  assert.ok(gOut.includes("Tercer mensaje."));
});
check("an <hr> with no header block after it does NOT split", () => {
  // A horizontal rule is ordinary punctuation inside a message; splitting on
  // every one of them would shred a normal message into bogus sections.
  const top = newest(gOut);
  assert.ok(top.includes("Latest reply, and then the checklist:"));
  assert.ok(top.includes("Step one, then step two."), "the rule kept its own message whole");
  assert.ok(/^---$/m.test(top), "and the rule itself survives as a thematic break");
});
check("the introducing rule is not left under the heading it introduces", () => {
  // Every section, not just the first: a boundary partway into a wrapper is
  // extracted inside a clone of that wrapper, so the rule is not necessarily
  // the fragment's first child and cannot simply be trimmed off afterwards.
  gOut.split(/^## .*$/m).slice(1).forEach((part) => {
    assert.ok(!/^\s*---/.test(part), "section opens with a stray rule: " + part.slice(0, 60));
  });
});
check("a Spanish header block reads like an English one", () => {
  // De/Enviado/Para/Asunto. The word lists are the extension point for any
  // further locale; nothing else in the splitter is language-specific.
  assert.ok(gOut.includes("## Carol <carol@example.com> — lunes, 17 de agosto de 2026 8:00"));
});
check("the <hr> path strips addresses like every other boundary", () => {
  const md = run({ fixture: G, strip: true, mode: "alias" });
  assert.ok(!/@example\.com/.test(withoutIds(md)));
  assert.ok(/^## User\d+ — lunes, 17 de agosto de 2026 8:00$/m.test(md));
});

console.log("\nGmail and Apple Mail quotes");
check("a Gmail chain gives one section per quoted message, not one for the lot", () => {
  // The whole rest of the thread sits inside the FIRST blockquote, so a filter
  // that keeps only outermost markers returns exactly one section here however
  // long the chain is.
  assert.deepStrictEqual(headings(hOut), [
    "Bob <bob@example.com> — Thu, Aug 20, 2026 at 9:15 AM",
    "Carol <carol@example.com> — Wed, Aug 19, 2026 at 4:40 PM",
    "Alice <alice@example.com> — Tue, Aug 18, 2026 at 11:05 AM",
    "Bob <bob@example.com> — Mon, Aug 17, 2026 at 8:00 AM",
  ]);
});
check("each Gmail section holds its own message and no other", () => {
  const parts = hOut.split(/^## .*$/m);
  assert.ok(parts[0].includes("Newest note, from Alice."));
  assert.ok(parts[1].includes("Bob answers the second question."));
  assert.ok(!parts[1].includes("Carol asks the second question."));
  assert.ok(parts[2].includes("Carol asks the second question."));
  assert.ok(parts[3].includes("Alice answers the first question."));
  assert.ok(parts[4].includes("Bob asks the first question."));
});
check("the oldest message is not buried under four levels of '>'", () => {
  // Cutting a nested chain leaves each section wrapped in clones of every
  // blockquote it sat inside; twenty-one of those turn the oldest message into
  // twenty-one '>' prefixes on every line, and say nothing the heading has not.
  assert.ok(!/^\s*>/m.test(hOut), "no quote prefixes: " + (hOut.match(/^\s*>.*/m) || [])[0]);
});
check("the container div and the blockquote in it are ONE boundary", () => {
  // Gmail's <div class="gmail_quote"> holds the attribution and then the
  // <blockquote class="gmail_quote">. Counting both would put an empty section
  // titled after the attribution in front of every real one.
  assert.strictEqual(headings(hOut).length, 4, "four quoted messages, four sections");
  assert.ok(!/^## Quoted message/m.test(hOut), "no section fell back to a number");
});
check("an Apple Mail blockquote[type=cite] still splits, and nests", () => {
  assert.deepStrictEqual(headings(iOut), [
    "Bob <bob@example.com> — Aug 20, 2026, at 9:15 AM",
    "Carol <carol@example.com> — Aug 19, 2026, at 4:40 PM",
  ]);
  assert.ok(!/^\s*>/m.test(iOut));
});
check("the attribution line goes with the quote it introduces", () => {
  // It sits outside the blockquote, so cutting at the blockquote would leave
  // "On … wrote:" dangling at the end of the message above it.
  assert.ok(!/wrote:/.test(newest(iOut)), "not stranded: " + newest(iOut));
  assert.ok(newest(iOut).includes("Thanks, that works for me."));
});
check("Gmail and Apple sections strip addresses like the Outlook ones", () => {
  [H, I].forEach((f) => {
    const md = run({ fixture: f, strip: true, mode: "alias" });
    assert.ok(!/@example\.com/.test(withoutIds(md)), "no address survives");
    assert.ok(/^## User\d+ — /m.test(md), "the heading is aliased too");
  });
});

console.log("\nthe new boundaries behave like the old ones");
check("splitting off puts every fixture back in one piece", () => {
  [G, H, I].forEach((f) => {
    const md = run({ fixture: f, splitThread: false });
    assert.ok(!/^## /m.test(md));
  });
});
check("no new fixture leaks raw HTML, either way up", () => {
  [G, H, I].forEach((f, i) => {
    const name = "GHI"[i];
    noRawHtml(run({ fixture: f }), name + " (strip on)");
    noRawHtml(run({ fixture: f, boiler: false }), name + " (strip off)");
    noRawHtml(run({ fixture: f, boiler: false, strip: true, mode: "alias" }), name + " (aliased)");
  });
});
check("repeat conversions stay byte-identical", () => {
  assert.strictEqual(run({ fixture: G }), gOut);
  assert.strictEqual(run({ fixture: H }), hOut);
  assert.strictEqual(run({ fixture: I }), iOut);
});

console.log("\n" + checks + " checks passed.");
