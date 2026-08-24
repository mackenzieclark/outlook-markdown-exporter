const assert = require("assert");
const { JSDOM } = require("jsdom");
const fs = require("fs");

// Build the pane from the real taskpane.html rather than a hand-copied stub, so
// the harness exercises the shipped markup (ids, defaults, the hidden child
// group) instead of drifting from it. jsdom does not run the <script> tags.
const dom = new JSDOM(fs.readFileSync(__dirname + "/../src/taskpane.html", "utf8"));
const w = dom.window;
const doc = w.document;
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

let fixture = A;
global.Office = {
  onReady: (cb) => cb(),
  CoercionType: { Html: "html" },
  AsyncResultStatus: { Succeeded: "succeeded" },
  context: {
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

// Drives the pane the way a user does: flip the controls, fire "change", let the
// add-in's own listeners re-render.
function run(opts) {
  opts = opts || {};
  fixture = opts.fixture || A;
  doc.getElementById("frontmatter").checked = opts.frontmatter !== false;
  doc.getElementById("splitThread").checked = opts.splitThread !== false;
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

console.log("\n" + checks + " checks passed.");
