#!/usr/bin/env python3
"""Rebuild a theme's base.docx after the design changes.

The base document carries everything the generator does NOT create itself:
the cover page, paragraph styles, header/footer and page setup.

Workflow when the design changes
--------------------------------
1. edit the design in Google Docs (or Word)
2. export it as .docx
3. python3 tools/make_base.py export.docx --theme default
4. rebuild a proposal and eyeball the PDF

Everything after the cut marker is dropped — the body is generated from
Markdown, so only the front matter of the document is kept.

    python3 tools/make_base.py export.docx --cut-at "INTERNAL"
    python3 tools/make_base.py export.docx --keep-all      # no cover, styles only
"""
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
from ..office import convert, find_soffice  # noqa: E402
from ..theme import THEMES_DIR as THEMES  # noqa: E402

DECIMAL_ATTR = re.compile(rb'(\sw:[a-zA-Z]+=")(-?\d+\.\d+)(")')


def paragraph_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def trim_after(document, marker: str) -> int | None:
    """Drop every body child from the first paragraph containing `marker`."""
    body = document.element.body
    cut = None
    for i, child in enumerate(body):
        if child.tag == qn("w:p") and marker.lower() in paragraph_text(child).lower():
            cut = i
            break
    if cut is None:
        return None
    sectPr = body.find(qn("w:sectPr"))
    for child in list(body)[cut:]:
        if child is not sectPr:
            body.remove(child)
    return cut


def name_all_styles(document) -> list[str]:
    """Word 'repairs' any file whose styles lack w:name — Google Docs omits them."""
    fixed = []
    for style in document.styles.element.findall(qn("w:style")):
        if style.find(qn("w:name")) is None:
            sid = style.get(qn("w:styleId"))
            name = OxmlElement("w:name")
            name.set(qn("w:val"), sid)
            style.insert(0, name)
            fixed.append(sid)
    return fixed


def round_decimal_attrs(path: Path) -> int:
    """Google Docs writes 863.9999 into integer-only attributes; Word dislikes it."""
    src = zipfile.ZipFile(path)
    tmp = path.with_suffix(".tmp.docx")
    total = 0
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename.endswith(".xml"):
                data, n = DECIMAL_ATTR.subn(
                    lambda m: m.group(1) + str(round(float(m.group(2)))).encode() + m.group(3), data)
                total += n
            out.writestr(item, data)
    src.close()
    shutil.move(tmp, path)
    return total


NUMBERING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
NUMBERING_CT = ("application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.numbering+xml")


def install_numbering(path: Path, numbering_xml: Path) -> str:
    """Put the theme's numbering.xml into the package.

    LibreOffice drops numbering.xml when the trimmed base has no lists left,
    and Google Docs numbers its abstractNumIds arbitrarily.  Installing the
    theme's own file makes the ids in theme.yaml stable forever.
    """
    src = zipfile.ZipFile(path)
    names = set(src.namelist())
    tmp = path.with_suffix(".tmp.docx")
    body = numbering_xml.read_bytes()

    rels_name = "word/_rels/document.xml.rels"
    rels = src.read(rels_name).decode("utf-8")
    if "numbering.xml" not in rels:
        rid = "rId%d" % (max([int(m) for m in re.findall(r'Id="rId(\d+)"', rels)] or [0]) + 1)
        rels = rels.replace(
            "</Relationships>",
            f'<Relationship Id="{rid}" Type="{NUMBERING_REL}" Target="numbering.xml"/>'
            "</Relationships>")

    types_name = "[Content_Types].xml"
    types = src.read(types_name).decode("utf-8")
    if "numbering+xml" not in types:
        types = types.replace(
            "</Types>",
            f'<Override PartName="/word/numbering.xml" ContentType="{NUMBERING_CT}"/></Types>')

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            if item.filename == "word/numbering.xml":
                continue
            data = src.read(item.filename)
            if item.filename == rels_name:
                data = rels.encode("utf-8")
            elif item.filename == types_name:
                data = types.encode("utf-8")
            out.writestr(item, data)
        out.writestr("word/numbering.xml", body)
    src.close()
    shutil.move(tmp, path)
    return "replaced" if "word/numbering.xml" in names else "added"


def normalize_with_libreoffice(path: Path) -> bool:
    """Rewrite the package with LibreOffice so the OOXML is schema-clean."""
    soffice = find_soffice()
    if not soffice:
        return False
    outdir = path.parent / "_lo"
    outdir.mkdir(exist_ok=True)
    produced = convert(soffice, path, "docx:MS Word 2007 XML", outdir)
    if produced.exists():
        shutil.move(produced, path)
    shutil.rmtree(outdir, ignore_errors=True)
    return True


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("export", type=Path, help=".docx exported from Google Docs or Word")
    ap.add_argument("--theme", default="default")
    ap.add_argument("--cut-at", default="INTERNAL",
                    help="drop the body from the first paragraph containing this text")
    ap.add_argument("--keep-all", action="store_true", help="keep the whole body")
    ap.add_argument("--no-libreoffice", action="store_true")
    args = ap.parse_args(argv)

    if not args.export.exists():
        print(f"not found: {args.export}", file=sys.stderr)
        return 1

    target = THEMES / args.theme / "base.docx"
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        backup = target.with_suffix(".docx.bak")
        shutil.copy(target, backup)
        print(f"previous base kept at {backup}")

    shutil.copy(args.export, target)
    document = docx.Document(str(target))

    if args.keep_all:
        print("body kept in full")
    else:
        cut = trim_after(document, args.cut_at)
        if cut is None:
            print(f"marker {args.cut_at!r} not found — nothing trimmed; "
                  f"pass --cut-at or --keep-all", file=sys.stderr)
        else:
            print(f"body trimmed from child {cut} (marker {args.cut_at!r})")

    named = name_all_styles(document)
    document.save(str(target))
    print(f"styles given a w:name: {len(named)}")

    rounded = round_decimal_attrs(target)
    print(f"decimal attributes rounded: {rounded}")

    if not args.no_libreoffice:
        print("normalized with LibreOffice" if normalize_with_libreoffice(target)
              else "LibreOffice not found — skipped normalization")

    numbering = target.parent / "numbering.xml"
    if numbering.exists():
        print(f"theme numbering.xml {install_numbering(target, numbering)}")
    else:
        print(f"no {numbering} — lists will not render", file=sys.stderr)

    check = docx.Document(str(target))
    print(f"\n{target}")
    print(f"  paragraphs {len(check.paragraphs)}  tables {len(check.tables)}")
    print(f"  sections   {len(check.sections)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
