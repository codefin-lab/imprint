#!/usr/bin/env python3
"""Imprint's regression suite: build every template and the smoke test, twice.

    python3 plugins/imprint/engine/tests/run.py            # docx and pptx, no PDF
    python3 plugins/imprint/engine/tests/run.py --pdf      # also render PDFs (needs LibreOffice)

Fails on a build error, a verifier error, two builds that differ, or a company-specific or
client name anywhere in the repository.
"""
from __future__ import annotations

import argparse
import hashlib
import re
import subprocess
import sys
import zipfile
from pathlib import Path

ENGINE = Path(__file__).resolve().parent.parent
PLUGIN = ENGINE.parent
REPO = PLUGIN.parent.parent
OUT = ENGINE / "tests" / "out"

TEMPLATES = ENGINE / "imprint" / "templates"
DOCS = sorted(TEMPLATES.glob("*.md"))
DECKS = [TEMPLATES / "presentation.md"]
SMOKE = ENGINE / "tests" / "smoke.md"

# names that belong in a brand plugin, never in the public engine
LEAKS = re.compile(r"(?i)\b(daol|tisco|vahalla|valhalla|cimb|investifi|zuellig|silom|"
                   r"land and houses|codefin company|codefin co\.)\b")
SCAN = {".py", ".md", ".yaml", ".yml", ".json", ".sh", ".txt", ".html", ""}


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with zipfile.ZipFile(path) as z:
        for name in sorted(z.namelist()):
            h.update(name.encode())
            h.update(z.read(name))
    return h.hexdigest()


def build(kind: str, md: Path, out: Path, pdf: bool) -> tuple[int, str]:
    # the package in this checkout, not an installed one
    cmd = [sys.executable, "-m", "imprint.cli", kind, str(md), "-o", str(out)]
    if pdf:
        cmd.append("--pdf")
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=ENGINE)
    return r.returncode, r.stdout + r.stderr


def new_from_template(out: Path) -> tuple[int, str]:
    dest = out / "new" / "proposal.md"
    if dest.exists():
        dest.unlink()
    r = subprocess.run([sys.executable, "-m", "imprint.cli", "new", "proposal", str(dest)],
                       capture_output=True, text=True, cwd=ENGINE)
    ok = r.returncode == 0 and dest.exists() and (dest.parent / "example-architecture.png").exists()
    return (0 if ok else 1), r.stdout + r.stderr


def leaks() -> list[str]:
    found = []
    for path in REPO.rglob("*"):
        if ".git" in path.parts or "out" in path.parts or path.suffix not in SCAN or not path.is_file():
            continue
        if path.name == "LICENSE":            # names the copyright holder, as it must
            continue
        if path.resolve() == Path(__file__).resolve():   # this file holds the list itself
            continue
        for n, line in enumerate(path.read_text(encoding="utf-8", errors="ignore").splitlines(), 1):
            if LEAKS.search(line):
                found.append(f"{path.relative_to(REPO)}:{n}: {line.strip()[:100]}")
    return found


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--pdf", action="store_true")
    args = ap.parse_args()
    OUT.mkdir(parents=True, exist_ok=True)
    failures = []

    jobs = [("docx", md, ".docx") for md in DOCS if md.name != "presentation.md"]
    jobs += [("docx", SMOKE, ".docx")]
    jobs += [("pptx", md, ".pptx") for md in DECKS]
    for script, md, ext in jobs:
        first, second = OUT / f"{md.stem}{ext}", OUT / f"{md.stem}.again{ext}"
        code, log = build(script, md, first, args.pdf)
        status = "ok"
        if code != 0 or "VERIFY FAILED" in log:
            status = "FAILED"
            failures.append(f"{md.name}: build failed\n{log}")
        else:
            build(script, md, second, False)
            if digest(first) != digest(second):
                status = "NOT DETERMINISTIC"
                failures.append(f"{md.name}: two builds differ")
        warnings = [l.strip() for l in log.splitlines() if "warning" in l]
        print(f"  {status:18} {md.name:18} {ext}" + (f"  ({len(warnings)} warning(s))" if warnings else ""))

    sys.path.insert(0, str(REPO / "scripts"))
    from release import ENSURE, versions
    written = versions()
    same = len(set(written.values())) == 1 and len({f.read_text() for f in ENSURE}) == 1
    print(f"  {'ok' if same else 'MISMATCH':18} versions {sorted(set(written.values()))}")
    if not same:
        failures.append("versions disagree or ensure-engine.sh copies differ: " + str(written))

    code, log = new_from_template(OUT)
    print(f"  {'ok' if code == 0 else 'FAILED':18} imprint new copies a template and its picture")
    if code:
        failures.append(f"imprint new: {log}")

    found = leaks()
    print(f"  {'ok' if not found else 'LEAK':18} brand and client names")
    failures += [f"leak: {f}" for f in found]

    for f in failures:
        print(f"\nFAIL {f}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
