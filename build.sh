#!/usr/bin/env bash
# Usage: ./build.sh <base-url> [repo-url] [out-dir]
#
# Builds both sideloadable manifests from the templates and stamps the version
# into the task pane. <Version> in manifest.template.xml is the single source of
# truth for the version number; nothing else needs editing to cut a release.
#
#   ./build.sh https://cdn.jsdelivr.net/gh/you/repo@main https://github.com/you/repo
#   ./build.sh https://localhost:3000 https://localhost:3000 dist   # local testing
#
# Writes <out-dir>/manifest.xml, <out-dir>/manifest.json and the zipped app
# package that unified-manifest hosts install from.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
base="${1:?base url}"; base="${base%/}"
repo="${2:-$base}"; repo="${repo%/}"
out="${3:-$here}"

version="$(sed -n 's:.*<Version>\([^<]*\)</Version>.*:\1:p' "$here/manifest.template.xml" | head -1)"
[ -n "$version" ] || { echo "no <Version> found in manifest.template.xml" >&2; exit 1; }

stamp() {
  sed -e "s|__BASE_URL__|$base|g" -e "s|__REPO_URL__|$repo|g" -e "s|__VERSION__|$version|g" "$1"
}

mkdir -p "$out"
# Must be absolute: the packaging step below runs from a staging directory, so a
# relative out-dir would resolve against that instead of the caller's cwd.
out="$(cd "$out" && pwd)"
stamp "$here/manifest.template.xml"  > "$out/manifest.xml"
stamp "$here/manifest.template.json" > "$out/manifest.json"

# The task pane's cache-busting query string has to match the version: Outlook
# caches add-in subresources in a store that ignores Cache-Control, so a changed
# URL is the only reliable way to get an updated taskpane.js to clients.
sed -i "s|taskpane\.js?v=[^\"]*|taskpane.js?v=$version|" "$here/src/taskpane.html"

# Unified-manifest hosts install from a zip holding the manifest and its two
# package-relative icons, not from a bare manifest file.
pkg="$out/outlook-markdown-exporter.zip"
rm -f "$pkg"
staging="$(mktemp -d)"
trap 'rm -rf "$staging"' EXIT
cp "$out/manifest.json" "$here/assets/color.png" "$here/assets/outline.png" "$staging/"
( cd "$staging" && 7z a -tzip -bso0 -bsp0 "$pkg" manifest.json color.png outline.png >/dev/null )

echo "version  $version"          >&2
echo "xml      $out/manifest.xml" >&2
echo "json     $out/manifest.json" >&2
echo "package  $pkg"              >&2
