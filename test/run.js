const { JSDOM } = require("jsdom");
const fs = require("fs");
const dom = new JSDOM("<!doctype html><body><input type=checkbox id=frontmatter checked><input type=checkbox id=splitThread checked><textarea id=out></textarea><div id=status></div><button id=copy></button><button id=download></button><span id=version></span><span id=host></span>");
const w = dom.window;
for (const k of ["document","DOMParser","NodeFilter","Node","Blob","URL"]) global[k] = w[k];
global.window = w; global.navigator = {};
global.TurndownService = require("turndown");
global.turndownPluginGfm = require("turndown-plugin-gfm");
const html = fs.readFileSync(__dirname + "/sample.html", "utf8");
let captured;
global.Office = {
  onReady: (cb) => cb(),
  CoercionType: { Html: "html" },
  AsyncResultStatus: { Succeeded: "succeeded" },
  context: { mailbox: { item: {
    subject: "RE: Q3 Deliverables",
    from: { displayName: "Alice", emailAddress: "alice@example.com" },
    to: [{ displayName: "Bob", emailAddress: "bob@example.com" }],
    cc: [],
    dateTimeCreated: new Date("2026-08-20T15:30:00Z"),
    internetMessageId: "<abc@example.com>",
    conversationId: "AAQk",
    attachments: [{ name: "plan.xlsx", isInline: false }, { name: "image001.png", isInline: true }],
    body: { getAsync: (t, cb) => cb({ status: "succeeded", value: html }) },
  } } },
};
require("../src/taskpane.js");
console.log(w.document.getElementById("out").value);
console.log("STATUS:", w.document.getElementById("status").textContent);
console.log("BUILD:", w.document.getElementById("version").textContent);
