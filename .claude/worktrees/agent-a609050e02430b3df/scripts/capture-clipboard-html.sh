#!/bin/sh
# Captures the macOS clipboard's HTML flavour as one of SPEC-125's external
# paste fixtures, byte for byte:
#
#   scripts/capture-clipboard-html.sh <name>
#
# Select A1:I1 of the capture sheet (specs/125 "Capture fixtures";
# e2e/fixtures/spec125-clipboard/NOTES.txt) in the producer, copy it with the
# producer's own Copy, then run this with the fixture's name — excel, sheets,
# libreoffice or numbers. It reads the clipboard's HTML with
# `osascript -e 'the clipboard as «class HTML»'`, hex-decodes the
# «data HTML…» result back to the original bytes, writes them to
# e2e/fixtures/spec125-clipboard/<name>.html and prints the byte count. Then
# record the producer, its version and the date in NOTES.txt. When the
# clipboard holds no HTML it writes nothing and exits non-zero. macOS only
# (osascript), so no automated test runs it.
set -eu

usage() {
  echo "usage: scripts/capture-clipboard-html.sh <name>"
  echo "  Writes the macOS clipboard's HTML flavour, byte for byte, to"
  echo "  e2e/fixtures/spec125-clipboard/<name>.html — name: excel, sheets, libreoffice or numbers."
}

if [ "$#" -ne 1 ] || [ -z "$1" ]; then
  usage
  exit 2
fi
name=$1
case "$name" in
  *[!A-Za-z0-9_-]*)
    echo "A fixture name holds only letters, digits, - and _" >&2
    exit 2
    ;;
esac
root=$(cd "$(dirname "$0")/.." && pwd)
out="$root/e2e/fixtures/spec125-clipboard/$name.html"

# AppleScript prints the flavour as «data HTML3C6D657461…»: the bytes in hex
# after the four-letter type code.
raw=$(osascript -e 'the clipboard as «class HTML»' 2>/dev/null) || raw=""
hex=$(printf '%s' "$raw" | LC_ALL=C sed -n 's/.*HTML\([0-9A-Fa-f]*\).*/\1/p')
if [ -z "$hex" ]; then
  echo "The clipboard holds no HTML" >&2
  exit 1
fi

tmp=$(mktemp)
if command -v xxd >/dev/null 2>&1; then
  printf '%s' "$hex" | xxd -r -p >"$tmp"
else
  printf '%s' "$hex" | perl -ne 'print pack("H*", $_)' >"$tmp"
fi
mkdir -p "$(dirname "$out")"
mv "$tmp" "$out"
bytes=$(wc -c <"$out" | tr -d ' ')
echo "$bytes bytes written to e2e/fixtures/spec125-clipboard/$name.html"
