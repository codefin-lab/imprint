#!/usr/bin/env bash
# Install what Imprint needs to build documents and decks. Safe to re-run.
#   ./setup.sh
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
ENGINE="$HERE/plugins/imprint/engine"

echo "Python packages"
python3 -m pip install --quiet -r "$ENGINE/requirements.txt"
python3 -c "import docx, pptx, yaml, matplotlib, PIL, lxml" && echo "  ok"

missing=0
check() {  # name, test command, how to install
  if eval "$2" >/dev/null 2>&1; then echo "  ok       $1"; else echo "  MISSING  $1: $3"; missing=1; fi
}
echo "Tools"
check "LibreOffice (PDF and table of contents)" \
  "command -v soffice || test -x /Applications/LibreOffice.app/Contents/MacOS/soffice" \
  "brew install --cask libreoffice   (Linux: apt install libreoffice)"
check "pdftoppm and pdftotext (reviewing PDFs)" "command -v pdftoppm && command -v pdftotext" \
  "brew install poppler   (Linux: apt install poppler-utils)"
echo "Fonts used by the default theme"
check "Sarabun" "fc-list | grep -qi sarabun || ls ~/Library/Fonts /Library/Fonts 2>/dev/null | grep -qi sarabun" \
  "https://fonts.google.com/specimen/Sarabun"
check "Anuphan" "fc-list | grep -qi anuphan || ls ~/Library/Fonts /Library/Fonts 2>/dev/null | grep -qi anuphan" \
  "https://fonts.google.com/specimen/Anuphan"

cat <<'EOF'

Then, in Claude Code:
  /plugin marketplace add codefin-lab/imprint
  /plugin install imprint@imprint
  /plugin install diagram-design@imprint
EOF
exit $missing
