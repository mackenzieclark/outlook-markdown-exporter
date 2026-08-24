# Contributing

Issues and pull requests are welcome. This is a small utility with no build
step, so getting started is quick.

## Before anything else: never paste real email

**Every identifier must be removed from anything you post in an issue, a pull
request, or a test fixture.** This project's whole subject matter is other
people's mail, and an issue is a permanent, public, search-indexed page. A
pasted conversion that "looks harmless" routinely carries a real customer name,
an internal project code, or a link that still resolves.

Strip all of the following, from both the input HTML and the converted
Markdown:

- names of people, including in prose and sign-offs
- email addresses, and the domains inside them
- company, product, customer, and project names
- subject lines
- phone numbers, postal addresses, job titles
- URLs — including the encoded original inside a Safe Links or Mimecast
  wrapper, which usually still points somewhere real
- `message_id` and `conversation_id` from the frontmatter; these identify a
  specific message in a specific mailbox
- invoice, ticket, order, and case numbers, and monetary amounts

Replace them using the scheme the tests already use: people are **Alice**,
**Bob**, and **Carol** at `example.com`, organisations are **Example Corp**,
and links point at `example.com`.

**Keep the structure — that is the part we actually need.** Outlook's markup
is what causes the bugs, so leave these exactly as they are:

- class and id names: `MsoNormal`, `MsoNormalTable`, `WordSection1`,
  `divRplyFwdMsg`, `appendonsend`, `x_` and `x_x_` prefixes
- the shape of the markup: nesting depth, empty cells, `<br>` vs `<p>`,
  inline `style` attributes
- link-protector hostnames such as `safelinks.protection.outlook.com` or
  `mimecastprotect.com` — just swap the wrapped URL for `example.com`

A report that says "a three-deep `MsoNormalTable` whose innermost cell holds a
`<p class=MsoNormal>` came out as X, expected Y" is more useful than a real
transcript, and costs no one their privacy.

### The add-in's own options are not a sanitiser

**Strip email addresses to names** with **Alias all names** will replace
addressed people with `User1`, `User2`, and so on. That is a convenience, not a
guarantee. By design it does **not** touch names in running prose, and it never
touches company names, subject lines, URLs, or message IDs. Read the output
yourself before posting it.

If you post real data by accident, say so in the issue and we will delete it —
but note that GitHub retains edit history, so deletion of the whole issue is
usually the only real fix.

### Security issues

Do not open a public issue for a vulnerability. Use GitHub's private
[security advisory](https://github.com/mackenzieclark/outlook-markdown-exporter/security/advisories/new)
form instead.

## Filing a good bug

The task pane footer shows the add-in version on the left and your Outlook
client and version on the right. Both matter — most conversion bugs are
specific to how one client emits HTML.

Include:

1. Version and client, from that footer.
2. Which options were on (the four checkboxes and the name mode).
3. A **sanitised** minimal fragment of the source HTML, if you can get it.
4. What you got, and what you expected.

The most valuable bug report is a fragment small enough to become a test
fixture. If you can reduce it to a dozen lines that still misbehave, you have
done most of the work of fixing it.

## Working on the code

There is no build step for the add-in itself. `src/taskpane.html` and
`src/taskpane.js` are served as-is.

```sh
cd test && npm i && node run.js
```

That runs the whole suite. It builds the pane from the real
`src/taskpane.html`, so the tests exercise the shipped markup rather than a
copy that can drift, and it drives the pane through its own change listeners
the way a user does.

To try changes in Outlook, see **Option B** in the [README](README.md) — a
local HTTPS server plus a manifest pointing at it.

### Conventions

- **ES5 only.** `var`, `function`, no arrow functions, no `const`/`let`, no
  template literals. The task pane ships unbundled and untranspiled.
- **No dependencies.** Turndown and its GFM plugin are vendored unmodified in
  `src/vendor/`; the only other packages are the test harness's, and the
  add-in must never require a build to run.
- **Colours come from CSS custom properties** (`--fg`, `--muted`, `--border`,
  `--accent`). The pane follows Outlook's light and dark themes; never
  hardcode a colour.
- **Comment the why, not the what.** Most of the tricky code here exists
  because of a specific Outlook behaviour — say which one.

### Versioning and releases

`<Version>` in `manifest.template.xml` is the single source of truth. Run:

```sh
./build.sh https://you.github.io/your-fork https://github.com/you/your-fork
```

That stamps the version into both manifests, rebuilds the app package, and
updates the `taskpane.js?v=` query string in `src/taskpane.html`.

**Bump the version for any change to `src/`.** Outlook caches add-in
subresources in a store that ignores `Cache-Control`, so that query string is
the only reliable way an existing install ever receives updated JavaScript. A
fix that ships without a bump will not reach anyone.

Note also that the two manifest formats disagree on version arity — the
add-in only manifest uses `a.b.c.d`, the unified manifest requires `n.n.n` —
and that AppSource rejects anything below `1.0.0.0`. `build.sh` handles the
conversion; you only edit the one number.

## Pull requests

- Add a test. The suite is plain `assert` with no framework, and every fix so
  far has a regression test that fails without it — please check that yours
  does too, by temporarily reverting your fix.
- Keep fixtures in the Alice/Bob/Carol scheme.
- Update the README if you change behaviour, including its **Limitations**
  list if you add or remove one.

## License

By contributing you agree that your contributions are licensed under the
[MIT License](LICENSE) that covers this project.
