#!/usr/bin/env bash
# render.sh DECK.pptx OUTDIR — PPTX → PDF (LibreOffice) → OUTDIR/slide-N.png (poppler), for pptx_live_preview.
set -euo pipefail
PPTX="$1"; OUTDIR="$2"
mkdir -p "$OUTDIR"
# Find LibreOffice.
SOFFICE="${SOFFICE:-}"
if [ -z "$SOFFICE" ]; then
  for c in soffice libreoffice /Applications/LibreOffice.app/Contents/MacOS/soffice; do
    if command -v "$c" >/dev/null 2>&1; then SOFFICE="$c"; break; fi
    if [ -x "$c" ]; then SOFFICE="$c"; break; fi
  done
fi
if [ -z "$SOFFICE" ]; then echo "render.sh: LibreOffice not found (brew install --cask libreoffice, or set SOFFICE=...)" >&2; exit 1; fi
if ! command -v pdftoppm >/dev/null 2>&1; then echo "render.sh: pdftoppm not found (brew install poppler)" >&2; exit 1; fi
TMP="$(mktemp -d)"
# A private profile dir keeps headless conversion from colliding with a running LibreOffice.
"$SOFFICE" -env:UserInstallation="file://$TMP/profile" --headless --convert-to pdf --outdir "$TMP" "$PPTX" >/dev/null 2>&1
PDF="$TMP/$(basename "${PPTX%.*}").pdf"
[ -f "$PDF" ] || { echo "render.sh: PDF conversion failed" >&2; exit 1; }
rm -f "$OUTDIR"/slide-*.png
pdftoppm -png -r "${RENDER_DPI:-110}" "$PDF" "$TMP/slide" 2>/dev/null
# pdftoppm zero-pads (slide-01.png); the viewer wants slide-1.png.
for f in "$TMP"/slide-*.png; do
  n="$(basename "$f" .png)"; n="${n#slide-}"; n="$((10#$n))"
  mv "$f" "$OUTDIR/slide-$n.png"
done
cp "$PDF" "$OUTDIR/../deck.pdf" 2>/dev/null || true
rm -rf "$TMP"
echo "rendered $(ls "$OUTDIR"/slide-*.png | wc -l | tr -d ' ') slides into $OUTDIR"
