#!/usr/bin/env python3
"""Set a new version everywhere it lives, so the skills and the engine move together.

    python3 scripts/release.py 0.3.0

Updates the engine's __version__, the plugin manifest, and REQUIRED in every skill's
scripts/ensure-engine.sh. Then:

    1. add the release to CHANGELOG.md
    2. python3 plugins/imprint/engine/tests/run.py
    3. git commit -am "Release 0.3.0" && git tag v0.3.0 && git push --follow-tags

A skill updated with `npx skills update` asks for the new engine, and its ensure-engine.sh
installs that tag on the next build.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
PLUGIN = REPO / "plugins/imprint"
INIT = PLUGIN / "engine/imprint/__init__.py"
MANIFEST = PLUGIN / ".claude-plugin/plugin.json"
ENSURE = sorted(PLUGIN.glob("skills/*/scripts/ensure-engine.sh"))
SEMVER = re.compile(r"^\d+\.\d+\.\d+$")


def versions() -> dict[str, str]:
    """Every place a version is written, and what it says."""
    found = {str(INIT.relative_to(REPO)): re.search(r'__version__ = "([^"]+)"', INIT.read_text()).group(1),
             str(MANIFEST.relative_to(REPO)): json.loads(MANIFEST.read_text())["version"]}
    for f in ENSURE:
        found[str(f.relative_to(REPO))] = re.search(r'^REQUIRED="([^"]+)"', f.read_text(), re.M).group(1)
    return found


def main() -> int:
    if len(sys.argv) != 2 or not SEMVER.match(sys.argv[1]):
        print(__doc__.strip())
        return 2
    new = sys.argv[1]
    INIT.write_text(re.sub(r'__version__ = "[^"]+"', f'__version__ = "{new}"', INIT.read_text()))
    manifest = json.loads(MANIFEST.read_text())
    manifest["version"] = new
    MANIFEST.write_text(json.dumps(manifest, indent=2) + "\n")
    for f in ENSURE:
        f.write_text(re.sub(r'^REQUIRED="[^"]+"', f'REQUIRED="{new}"', f.read_text(), flags=re.M))
    for where, v in versions().items():
        print(f"  {v}  {where}")
    print(f"\nnext: CHANGELOG.md, tests, then git tag v{new}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
