#!/usr/bin/env bash
# Renders build_mode_strategy.pptx to one JPEG per slide in preview/ for visual QA.
# Requires LibreOffice (soffice) and Poppler (pdftoppm).
set -euo pipefail
cd "$(dirname "$0")/.."
mkdir -p preview
soffice --headless --convert-to pdf --outdir preview build_mode_strategy.pptx >/dev/null
pdftoppm -jpeg -r 110 preview/build_mode_strategy.pdf preview/slide
rm -f preview/build_mode_strategy.pdf
ls -1 preview
