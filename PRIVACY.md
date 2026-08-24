# Privacy policy

**Copy as Markdown does not collect, transmit, or store your email.**

The add-in runs entirely inside Outlook's task pane on your own machine. It
reads the message you have open through the Office.js API, converts it to
Markdown in the page, and hands the result to your clipboard or a downloaded
`.md` file. The message body, its metadata, and the quoted thread never leave
the task pane. There is no server component, no account, no telemetry, no
analytics, and no local storage.

## What the add-in requests

`ReadItem` (`MailboxItem.Read.User`) — read-only access to the message
currently open. The add-in cannot read your other mail, send mail, or modify
anything in your mailbox.

## Network requests

Office add-ins are web pages, so the task pane's own files are fetched over
HTTPS each time it opens:

- `appsforoffice.microsoft.com` — the Office.js library, served by Microsoft.
- Whichever host the manifest points at (by default
  `mackenzieclark.github.io`, this repository's GitHub Pages site) — the
  task pane's HTML, JavaScript, and icons.

Those hosts see an ordinary web request for a static file, of the kind any web
page produces: IP address, timestamp, user agent. They receive nothing about
your mailbox or the message being converted, because no message data is ever
included in those requests. If you would rather not rely on someone else's
host, host the files yourself — see Option B in the [README](README.md).

## Changes

Any change to this policy will be a commit in this repository; its history is
the changelog.

## Contact

Open an issue at
<https://github.com/mackenzieclark/outlook-markdown-exporter/issues>.
