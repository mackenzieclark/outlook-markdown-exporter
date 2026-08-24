# Terms of use

**Copy as Markdown is free software. Use it for anything, including
commercial work. There is nothing to sign up for and nothing to pay.**

These terms cover your use of the add-in as distributed. The code itself is
MIT licensed, and the [LICENSE](LICENSE) file governs it — you may use,
copy, modify, merge, publish, distribute, sublicense, and sell copies of the
software, subject only to the MIT license's requirement that the copyright
notice and permission notice travel with it. Nothing here adds a restriction
on top of that.

## What the add-in is

A static web page plus a manifest. It runs entirely inside Outlook's task pane
on your own machine, reads only the message you currently have open, and
converts it to Markdown in the page. There is no server component, no account,
and no data collection. See [PRIVACY.md](PRIVACY.md) for how data is handled.

## Hosted files

Office add-ins are web pages, so the task pane's own static files — HTML,
JavaScript, icons — are fetched over HTTPS from a CDN or host each time the
pane opens. That hosting is provided as a convenience, with no guarantee of
availability. It may change, move, or stop working at any time, without
notice, and no support is promised for it. If you need the add-in to keep
working on your own terms, host the files yourself — see Option B in the
[README](README.md).

## No warranty

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.

In plain English: the conversion is best-effort. It may drop, mangle, or
misattribute content. Check the output before you rely on it, and keep the
original message.

## No liability

IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE
OR OTHER DEALINGS IN THE SOFTWARE.

## Changes

Any change to these terms will be a commit in this repository; its history is
the changelog.

## Contact

Open an issue at
<https://github.com/mackenzieclark/outlook-markdown-exporter/issues>.
