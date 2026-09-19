#!/usr/bin/env bash
# Rasterise assets/icon.svg into every size the manifest references.
# Rendering all sizes from one vector keeps the ribbon icon (16/32/80) and the
# install-dialog icon (64/128) visually identical.
set -euo pipefail

cd "$(dirname "$0")/.."
src="assets/icon.svg"
out="public/outlook-addin/assets"

command -v rsvg-convert >/dev/null || {
  echo "rsvg-convert not found. Install it with: brew install librsvg" >&2
  exit 1
}

mkdir -p "$out"
for size in 16 32 64 80 128; do
  rsvg-convert -w "$size" -h "$size" "$src" -o "$out/icon-$size.png"
  echo "wrote $out/icon-$size.png"
done
