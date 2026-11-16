"""Page numbers for the table of contents, read back off a rendered PDF.

    imprint toc report.docx            re-read the page numbers and write them into the contents
    imprint toc report.docx --pdf      ... and leave the refreshed PDF beside it

`imprint docx` already fills the contents.  Run `imprint toc` after anything that edits the
.docx afterwards and can move a heading to another page (a project script that restyles
cells, adds rows, swaps a picture): it renders the file as it is now, finds where each
heading landed, and rewrites the numbers.

Matching ignores whitespace and case, and accepts a heading that wraps onto several lines.
It still has to fill whole lines, so "1. Scope" never matches "1. Scope of work" in a
running header.  Whitespace matters for Thai: pdftotext often splits a Thai word with a
space ("ฟี เจอร์").
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

DASH = "—"          # what an entry reads when its heading was not found
WRAP = 4                 # a heading may wrap onto at most this many lines
_ENTRY_TAIL = re.compile(r"[.\-_·…]*(\d+|" + DASH + r")")


def norm(text: str) -> str:
    return re.sub(r"\s+", "", text).casefold()


def page_texts(pdf: Path) -> list[str]:
    out = subprocess.run(["pdftotext", "-layout", str(pdf), "-"],
                         capture_output=True, text=True, check=True).stdout
    return out.split("\f")


def match_pages(pages: list[str], headings: list[str], skip_until: int = 0) -> dict:
    """{heading: page number} for each heading found on a page after `skip_until`.

    A heading matches one line, or a run of up to WRAP consecutive lines, that holds exactly
    its text once whitespace is removed.  The first page it is found on wins.
    """
    keys = {h: norm(h) for h in headings}
    found: dict = {}
    for n, page in enumerate(pages, 1):
        if n <= skip_until:
            continue
        lines = [l for l in (norm(x) for x in page.splitlines()) if l]
        runs = set()
        for i in range(len(lines)):
            acc = ""
            for line in lines[i:i + WRAP]:
                acc += line
                runs.add(acc)
        for h, k in keys.items():
            if h not in found and k in runs:
                found[h] = n
    return found


def _is_entry_line(line: str, keys: list[str]) -> bool:
    return any(line.startswith(k) and _ENTRY_TAIL.fullmatch(line[len(k):]) for k in keys)


def contents_pages(pages: list[str], title: str | None, entries: list[str]) -> tuple[int, int]:
    """(first, last) page of the contents; (0, 0) when it cannot be found.

    The first page is the one carrying the contents title (the heading written above
    `<!-- toc -->`, or "Table of contents").  The contents continues over every following
    page that still holds an entry line: an entry's text, a leader, then a number or a dash.
    """
    markers = {norm(t) for t in (title, "Table of contents", "Contents") if t}
    keys = [norm(e) for e in entries]
    for n, page in enumerate(pages, 1):
        if markers & {norm(l) for l in page.splitlines()}:
            last = n
            while last < len(pages) and any(_is_entry_line(norm(l), keys)
                                            for l in pages[last].splitlines()):
                last += 1
            return n, last
    return 0, 0


def heading_pages(pdf: Path, headings: list[str], title: str | None = None,
                  entries: list[str] | None = None) -> dict:
    pages = page_texts(pdf)
    _, last = contents_pages(pages, title, entries if entries is not None else headings)
    return match_pages(pages, headings, last)


# ---------------------------------------------------------------- rewriting a built .docx

def _contents_paragraphs(document):
    """The contents entries of a document Imprint built: consecutive paragraphs with one
    right-aligned leader tab whose last run is a page number or a dash."""
    from docx.oxml.ns import qn
    block, title = [], None
    for p in document.paragraphs:
        tabs = p._p.pPr.findall(qn("w:tabs") + "/" + qn("w:tab")) if p._p.pPr is not None else []
        entry = (len(tabs) == 1 and tabs[0].get(qn("w:val")) == "right"
                 and tabs[0].get(qn("w:leader")) is not None and "\t" in p.text
                 and p.runs and re.fullmatch(r"\d+|" + DASH, p.runs[-1].text.strip() or "x"))
        if entry:
            block.append(p)
        elif block:
            break
        elif p.text.strip():
            title = p.text.strip()
    return title, block


def refill(docx_path: Path, pdf_path: Path) -> tuple[dict, list[str]]:
    """Write the page numbers found in `pdf_path` into the contents of `docx_path`.
    Returns ({entry: page}, [entries not found])."""
    import docx
    document = docx.Document(str(docx_path))
    title, paras = _contents_paragraphs(document)
    entries = [p.text.rsplit("\t", 1)[0] for p in paras]
    pages = heading_pages(pdf_path, entries, title, entries)
    for p, e in zip(paras, entries):
        runs = p.runs
        runs[-1].text = str(pages[e]) if e in pages else DASH
    document.save(str(docx_path))
    return pages, [e for e in entries if e not in pages]


def main(argv=None) -> int:
    from .office import convert, find_soffice
    ap = argparse.ArgumentParser(prog="imprint toc", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx", type=Path)
    ap.add_argument("--pdf", action="store_true", help="keep the refreshed PDF beside the .docx")
    args = ap.parse_args(argv)
    if not args.docx.exists():
        print(f"not found: {args.docx}", file=sys.stderr)
        return 1
    soffice = find_soffice()
    if not soffice:
        print("imprint toc needs LibreOffice to lay the document out", file=sys.stderr)
        return 1
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        pages, missing = refill(args.docx, convert(soffice, args.docx, outdir=Path(tmp)))
        # a wider number could in theory push a heading over; read it all back once more
        again, _ = refill(args.docx, convert(soffice, args.docx, outdir=Path(tmp)))
    moved = [e for e in pages if again.get(e) != pages[e]]
    pdf = convert(soffice, args.docx) if args.pdf else None
    for e, n in again.items():
        print(f"  {n:>4}  {e}")
    if moved:
        print(f"  warning     {len(moved)} entry(s) moved after rewriting: {moved[:3]}")
    if missing:
        print(f"  warning     no page found for: {missing}")
        return 1
    print(f"  contents    {len(again)}/{len(again)} entries" + (f", {pdf}" if pdf else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
