# outlook-markdown-exporter

One-button **Copy as Markdown** for Outlook. Converts the open email — including
its quoted reply chain — to clean Markdown with YAML frontmatter and puts it on
the clipboard, ready to paste into an LLM agent, Obsidian, a ticket, anywhere.

Works in **New Outlook**, **Outlook on the web**, and classic Outlook for
Windows/Mac (anything that runs Office.js add-ins). No VBA, no Pandoc, no
Python — a static HTML page plus a manifest.

```markdown
---
subject: "RE: Q3 Deliverables"
from: "Alice <alice@example.com>"
to:
  - "Bob <bob@example.com>"
date: "2026-08-20T15:30:00.000Z"
message_id: "<abc@example.com>"
conversation_id: "AAQk…"
attachments:
  - "plan.xlsx"
---
# RE: Q3 Deliverables

Here is the **update** on the _Q3_ deliverables. See [the doc](https://example.sharepoint.com/…).

- Item one
- Item two

| Env | Status |
| --- | --- |
| prod | green |

---

## Bob <bob@example.com> — Tuesday, August 19, 2026 3:02 PM

Alice, can you send the status?
```

## What it does

- Reads the message via Office.js (`item.body.getAsync(Html)` + item metadata).
- Strips Outlook/Word HTML noise (`<o:p>`, conditional comments, styles,
  inline `cid:` images, empty `MsoNormal` paragraphs).
- Unwraps Defender Safe Links back to the real URL.
- Converts with [Turndown](https://github.com/mixmark-io/turndown) + GFM plugin
  (tables, strikethrough, task lists).
- Splits the quoted history at Outlook's reply boundaries (`divRplyFwdMsg`,
  `appendonsend`, Gmail quotes, `border-top` header blocks) into `##` sections
  titled by sender and date. Toggle off to get the body verbatim.
- Converts as soon as the pane opens; **Copy Markdown** puts it on the
  clipboard (the clipboard API only accepts a real click, so the copy waits
  for one) and **Save .md** downloads it.

## Install (sideload)

Office add-ins are fetched over HTTPS every time the task pane opens — the
add-in model has no fully offline install, even for desktop Outlook. Two ways
to satisfy that; neither needs Node (it's only used by the test harness).

### Option A: install straight from this repo — no clone, no hosting

Two manifests are committed, both already pointing at this repo through
[jsDelivr](https://www.jsdelivr.com/), which serves public GitHub repos with
correct MIME types. Nothing to build or deploy — pick the one that matches
your client:

| Client | Download | Manifest type |
| --- | --- | --- |
| Any Outlook: Windows (new **and** classic), Mac, web | [`manifest.xml`](manifest.xml) | add-in only |
| new Outlook on Windows, Outlook on the web | [`outlook-markdown-exporter.zip`](outlook-markdown-exporter.zip) | [unified manifest for Microsoft 365](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview) |

**If you are not sure, take `manifest.xml`.** It works everywhere:

1. Download [`manifest.xml`](manifest.xml) (Raw → Save).
2. Sideload: open <https://aka.ms/olksideload> → *My add-ins* →
   *Add a custom add-in* → *Add from file* → pick the downloaded
   `manifest.xml`. (Classic Outlook: *Get Add-ins* → *My add-ins* → same
   dialog.)
3. Open any message; the **Copy as Markdown** button is on the message
   toolbar (New Outlook: the `…` overflow menu on the message, until you
   pin it).

The `.zip` is the same add-in packaged in Microsoft's newer manifest format.
It is the forward-looking option and the one AppSource submissions are moving
to, but it **does not support Outlook on Mac or perpetual (non-subscription)
Office**, and it does not install through the *Add from file* dialog above —
see [sideloading with the unified
manifest](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-add-in-with-unified-manifest).

Both manifests track `@main`, which jsDelivr caches for up to ~12 hours, so
updates arrive on that delay. Forks and self-hosters can repoint everything at
once:

```sh
./build.sh https://cdn.jsdelivr.net/gh/you/your-fork@main https://github.com/you/your-fork
```

### Option B: local server — private fork / no public hosting

Any static HTTPS server works. [Caddy](https://caddyserver.com) is a single
binary that mints a localhost certificate and installs its CA into the
system trust store itself (which is what makes Outlook's WebView accept it):

1. Install Caddy: `winget install CaddyServer.Caddy` (or scoop/brew/apt).
2. In the repo root, create a `Caddyfile`:
   ```
   https://localhost:3000 {
       root * .
       file_server
       # Outlook caches add-in files in a store that ignores Cache-Control,
       # but this keeps the browser out of the way while developing.
       header Cache-Control "no-store"
       tls internal
   }
   ```
3. `caddy run` from the repo root; approve the one-time trust prompt.
4. `./build.sh https://localhost:3000 https://localhost:3000 dist`, then
   sideload `dist/manifest.xml` as in Option A.

The server must be running whenever you use the button — `caddy start` from
a login task (e.g. a shortcut in `shell:startup`) makes that permanent.
Outlook runs the task pane on your machine either way; the email content
never leaves it.

### Versioning

The version lives in exactly one place: `<Version>` in
`manifest.template.xml`. `build.sh` stamps it into both manifests and into
the task pane's `taskpane.js?v=` URL, and the task pane shows it in the
footer. That query string matters — Outlook caches add-in subresources in a
store that ignores `Cache-Control`, so bumping the version is what actually
delivers updated JavaScript to an existing install. Note that AppSource
rejects versions below `1.0.0`.

### Tests

```sh
cd test && npm i && node run.js
```
Runs the conversion against `test/sample.html` (a synthetic Outlook/Word reply
chain) under jsdom with a stubbed `Office` object and prints the Markdown.

## Limitations

- **Thread = quoted history.** Office.js has no conversation API, so "thread"
  means whatever is quoted inside the message you have open. Open the newest
  message to get the whole chain. Splitting true sibling messages would need
  Microsoft Graph (`/me/messages?$filter=conversationId eq …`) and an Entra app
  registration with tenant consent — planned as an optional v2.
- Header detection relies on Outlook's `From: / Sent: / To:` block; other
  clients' quote formats fall back to `Quoted message N`.
- Inline images are dropped; linked images become `![alt](url)`.
- Tables that Outlook emits without a real header row (`MsoNormalTable`, all
  `<td>`) are passed through as raw HTML, because the GFM converter only
  handles tables whose first row is `<th>`.
- Requires Mailbox requirement set 1.8 (any current Outlook).

## Privacy

The add-in reads only the message you have open, converts it in the task pane,
and sends nothing anywhere. See [PRIVACY.md](PRIVACY.md).

## License

MIT. Icon: [Teenyicons](https://github.com/teenyicons/teenyicons) markdown
glyph (MIT, © 2020 Anja van Staden), recolored. `src/vendor/` contains
unmodified builds of [Turndown](https://github.com/mixmark-io/turndown) 7.2.0
and turndown-plugin-gfm 1.0.2 (both MIT).
