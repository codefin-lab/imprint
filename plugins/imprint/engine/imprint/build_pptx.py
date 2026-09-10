#!/usr/bin/env python3
"""Build a .pptx deck (and optionally a PDF) from Markdown.

    imprint pptx deck.md -o deck.pptx
    imprint pptx deck.md --theme default --pdf

Design lives in themes/<name>/ (theme.yaml under slides:, plus base.pptx made by
`imprint make-base-pptx`), never in this code.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

from . import Theme
from .office import convert, find_soffice
from .slides import build


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="imprint pptx", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("markdown", type=Path)
    ap.add_argument("-o", "--out", type=Path, help="output .pptx (default: alongside the input)")
    ap.add_argument("--theme", default="default", help="theme name, or a path to a theme folder")
    ap.add_argument("--pdf", action="store_true", help="also render a PDF preview (needs LibreOffice)")
    ap.add_argument("--strict", action="store_true", help="fail on any warning")
    args = ap.parse_args(argv)

    if not args.markdown.exists():
        print(f"not found: {args.markdown}", file=sys.stderr)
        return 1
    theme = Theme.load(args.theme)
    out = args.out or args.markdown.with_suffix(".pptx")
    stats = build(args.markdown, out, theme)
    print(f"{out}")
    print(f"  theme       {theme.name}")
    print(f"  slides      {stats['slides']}")
    for w in stats["warnings"]:
        print(f"  warning     {w}")
    if args.pdf:
        soffice = find_soffice()
        if not soffice:
            print("  pdf         skipped (LibreOffice not found)")
        else:
            print(f"  pdf         {convert(soffice, out)}")
    if stats["warnings"] and args.strict:
        print("  FAILED (--strict)")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
