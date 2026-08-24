# Privacy policy

**Copy as Markdown does not collect, transmit, or store your email.**

The add-in runs entirely inside Outlook's task pane on your own machine. It
reads the message you have open through the Office.js API, converts it to
Markdown in the page, and hands the result to your clipboard or a downloaded
`.md` file. The message body, its metadata, and the quoted thread never leave
the task pane. There is no server component, no account, no telemetry, no
analytics, and no tracking of any kind. The only thing the add-in stores is
your own checkbox settings — see "Your settings" below.

## What the add-in requests

`ReadItem` (`MailboxItem.Read.User`) — read-only access to the message
currently open. The add-in cannot read your other mail, send mail, or modify
anything in your mailbox.

## Your settings

The task pane is rebuilt every time it opens, so your choice of options would
otherwise reset on every message. To avoid that, the add-in saves the state of
its checkboxes and radio buttons — and nothing else — using Office's
`roamingSettings` API, which stores them **in your own mailbox**. That is why
the settings follow you to Outlook on the web and to other machines.

What is stored is five values: whether each option is on, and which name mode
is selected. No message content, no addresses, no names. Microsoft notes that
roaming settings are not secure storage and can be read by other services with
access to your mailbox, such as Microsoft Graph, which is precisely why
nothing sensitive is put there. Clearing them is a matter of removing the
add-in.

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
