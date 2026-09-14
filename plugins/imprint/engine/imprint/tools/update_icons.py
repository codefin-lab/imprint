"""Refresh the bundled icon set from a Lucide release.

    npm pack lucide-static && tar xzf lucide-static-*.tgz
    python -m imprint.tools.update_icons package/

Writes imprint/icons/lucide/{icons.json.gz, tags.json.gz, LICENSE, VERSION}. The gzip
files carry no timestamp, so the same release gives the same bytes.
"""
from __future__ import annotations

import gzip
import json
import shutil
import sys
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "icons" / "lucide"


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        print(__doc__.strip())
        return 2
    src = Path(argv[0])
    nodes = json.loads((src / "icon-nodes.json").read_text(encoding="utf-8"))
    tags = json.loads((src / "tags.json").read_text(encoding="utf-8"))
    version = json.loads((src / "package.json").read_text(encoding="utf-8"))["version"]
    OUT.mkdir(parents=True, exist_ok=True)
    for name, data in (("icons.json.gz", nodes), ("tags.json.gz", {k: tags.get(k, []) for k in nodes})):
        blob = json.dumps(dict(sorted(data.items())), separators=(",", ":"), ensure_ascii=False).encode()
        (OUT / name).write_bytes(gzip.compress(blob, compresslevel=9, mtime=0))
    shutil.copyfile(src / "LICENSE", OUT / "LICENSE")
    (OUT / "VERSION").write_text(f"lucide-static {version}\n", encoding="utf-8")
    print(f"{len(nodes)} icons from lucide-static {version} into {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
