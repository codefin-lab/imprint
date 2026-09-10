#!/usr/bin/env bash
# Make sure the imprint engine this skill was written for is installed.
# Safe to run before every build: it does nothing when a new enough engine is there.
# scripts/release.py sets REQUIRED; the skill and the engine release together.
set -euo pipefail
REQUIRED="0.2.3"
SPEC="git+https://github.com/codefin-lab/imprint@v${REQUIRED}#subdirectory=plugins/imprint/engine"

export PATH="$HOME/.local/bin:$PATH"
current=""
if command -v imprint >/dev/null 2>&1; then current="$(imprint --version 2>/dev/null | awk '{print $NF}')"; fi
if [ -n "$current" ] && [ "$(printf '%s\n%s\n' "$REQUIRED" "$current" | sort -V | tail -1)" = "$current" ]; then
  echo "imprint $current (this skill needs $REQUIRED or later): ok"
  exit 0
fi

echo "installing imprint $REQUIRED (found: ${current:-none})"
if command -v uv >/dev/null 2>&1; then
  uv tool install --force --quiet "$SPEC"
elif command -v pipx >/dev/null 2>&1; then
  pipx install --force "$SPEC"
else
  python3 -m pip install --user --upgrade --quiet "$SPEC"
fi

if ! command -v imprint >/dev/null 2>&1; then
  echo "installed, but 'imprint' is not on PATH: add \$HOME/.local/bin to PATH" >&2
  exit 1
fi
imprint --version
