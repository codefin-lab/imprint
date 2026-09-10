#!/usr/bin/env python3
"""Build a .docx (and optionally a PDF) from Markdown.

    imprint docx proposal.md -o proposal.docx
    imprint docx proposal.md --theme default --pdf

Design lives in themes/<name>/ (theme.yaml + base.docx), never in this code.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

from . import Theme, build
from .office import convert, find_soffice
from .verify import check

HERE = Path(__file__).resolve().parent


def to_pdf(soffice: str, docx_path: Path) -> Path:
    return convert(soffice, docx_path)


def heading_pages(pdf: Path, headings: list[str], skip_until: int = 1) -> dict:
    """Which page each heading actually landed on, read back off the PDF.

    A heading also appears in the table of contents itself, so a page is only
    accepted once past the contents — `skip_until` is the last page the contents
    occupy.  The match is against a whole line, not a substring, or "1. SCOPE"
    would match "1. SCOPE OF WORK" in a running header.
    """
    text = subprocess.run(["pdftotext", "-layout", str(pdf), "-"],
                          capture_output=True, text=True).stdout
    pages = text.split("\f")
    found = {}
    for n, page in enumerate(pages, 1):
        if n <= skip_until:
            continue
        lines = {" ".join(l.split()) for l in page.splitlines()}
        for h in headings:
            if h not in found and " ".join(h.split()) in lines:
                found[h] = n
    return found


def toc_last_page(pdf: Path, marker: str = "TABLE OF CONTENTS") -> int:
    text = subprocess.run(["pdftotext", str(pdf), "-"], capture_output=True, text=True).stdout
    for n, page in enumerate(text.split("\f"), 1):
        if marker.lower() in page.lower():
            return n
    return 1


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="imprint docx", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("markdown", type=Path)
    ap.add_argument("-o", "--out", type=Path, help="output .docx (default: alongside the input)")
    ap.add_argument("--theme", default="default", help="theme name, or a path to a theme folder")
    ap.add_argument("--pdf", action="store_true", help="also render a PDF preview (needs LibreOffice)")
    ap.add_argument("--no-verify", action="store_true")
    ap.add_argument("--strict", action="store_true",
                    help="fail if any {{placeholder}} is still unfilled")
    args = ap.parse_args(argv)

    if not args.markdown.exists():
        print(f"not found: {args.markdown}", file=sys.stderr)
        return 1

    theme = Theme.load(args.theme)
    out = args.out or args.markdown.with_suffix(".docx")
    stats = build(args.markdown, out, theme)

    # A table of contents needs page numbers, and page numbers only exist once
    # something has laid the document out.  So: build with placeholder numbers,
    # render, read the real ones back, rebuild.  The placeholder pass writes the
    # same number of TOC lines, so pagination does not move between passes.
    soffice = find_soffice()
    if stats.get("has_toc") and stats.get("headings"):
        if not soffice:
            print("  warning     table of contents left blank — needs LibreOffice "
                  "to find the page numbers")
        else:
            first = to_pdf(soffice, out)
            pages = heading_pages(first, stats["headings"], toc_last_page(first))
            stats = build(args.markdown, out, theme, toc_pages=pages)
            # NB: not `check` — that name is the imported verify function, and
            # assigning it here would make it local to main() for the whole
            # function, breaking the verify call further down.
            confirmed = heading_pages(to_pdf(soffice, out), stats["headings"],
                                      toc_last_page(out.with_suffix(".pdf")))
            moved = {h: (pages.get(h), confirmed.get(h)) for h in stats["headings"]
                     if pages.get(h) != confirmed.get(h)}
            if moved:
                print(f"  warning     {len(moved)} contents entry(s) shifted between "
                      f"passes: {list(moved)[:3]}")
            else:
                entries = stats.get("toc_entries", [])
                missing = [h for h in entries if h not in pages]
                print(f"  contents    {len(entries) - len(missing)}/{len(entries)} "
                      f"entries, page numbers verified"
                      + (f" — NOT FOUND: {missing}" if missing else ""))
    print(f"{out}")
    print(f"  theme       {theme.name}")
    print(f"  blocks      {stats['blocks']}")
    print(f"  tables      {stats['tables']}")
    print(f"  placeholders {stats['placeholders']} filled")
    for w in dict.fromkeys(stats.get("warnings", [])):
        print(f"  warning     {w}")

    exit_code = 0
    if not args.no_verify:
        errors, warnings = check(out, theme)
        for w in warnings:
            print(f"  warning     {w}")
        if errors:
            print("  VERIFY FAILED")
            for e in errors:
                print(f"    - {e}")
            return 2
        if warnings and args.strict:
            print("  VERIFY FAILED (--strict)")
            return 2
        print("  verify      ok")

    if args.pdf:
        soffice = find_soffice()
        if not soffice:
            print("  pdf         skipped (LibreOffice not found)")
        else:
            print(f"  pdf         {convert(soffice, out)}")
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
