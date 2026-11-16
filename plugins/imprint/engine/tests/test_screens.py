#!/usr/bin/env python3
"""Chips, checkboxes, alert text, pictures in cells, two-row headers, screen specs and galleries.

    python3 plugins/imprint/engine/tests/test_screens.py

Run by tests/run.py. Builds one small document with generated pictures and checks the .docx XML.
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ENGINE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ENGINE))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from PIL import Image  # noqa: E402

from imprint.markdown import parse, rich_spans  # noqa: E402
from imprint.render import build  # noqa: E402
from imprint.theme import Theme  # noqa: E402
from imprint.verify import check  # noqa: E402

DOC = """\
---
cover: false
---

# SCREENS

Status [[Active]], [[Draft]], [[Hold|amber]] and !!no going back!!.

Table: Fields {header=2}
| Fields | > | Required | > |
| --- | --- | --- | --- |
| Label | EN | Draft | Example |
| Name | [x] | [[Y]] | ![](pics/wide.png) |
| Port | [ ] | [[N]] | text<br>![](pics/wide.png) |

::: screen pics/phone.png
### Alert screen

| TH | EN |
| --- | --- |
| แจ้งเตือน | Alert |

Remark:
- Tap Home to go back
:::

::: screens {per_row=2}
![One](pics/phone.png)
![Two](pics/phone.png)
![Three](pics/phone.png)
:::
"""

failures: list[str] = []


def expect(ok: bool, what: str):
    print(f"  {'ok' if ok else 'FAILED':8} {what}")
    if not ok:
        failures.append(what)


def pictures(el) -> list:
    return el.findall(".//" + qn("w:drawing"))


def extent_in(drawing) -> float:
    ext = drawing.find(".//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent")
    return int(ext.get("cx")) / 914400


def main() -> int:
    spans = list(rich_spans("[[Active]] [[N/A|gray]] !!red!! [x] [ ] **b**"))
    marks = [m for *_, m in spans if m]
    expect(marks == [("chip", None), ("chip", "gray"), ("alert", None), ("box", None), ("box", None)],
           "inline marks: chip, chip with colour, alert, two checkboxes")
    expect([t for t, *_ in spans if t in ("☑", "☐")] == ["☑", "☐"],
           "[x] and [ ] become ballot boxes")
    blocks = parse(DOC)
    kinds = [k for k, _ in blocks]
    expect("screen" in kinds and "screens" in kinds, "::: screen and ::: screens parse as blocks")
    screen = next(p for k, p in blocks if k == "screen")
    expect(screen[0] == "pics/phone.png" and [k for k, _ in screen[1]] == ["h", "table", "p", "ul"],
           "a screen keeps its picture and the blocks inside")
    shots = next(p for k, p in blocks if k == "screens")
    expect(len(shots[0]) == 3 and shots[1] == {"per_row": "2"}, "a gallery lists its pictures and options")

    theme = Theme.load()
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        (tmp / "pics").mkdir()
        Image.new("RGB", (1200, 400), "#3366cc").save(tmp / "pics" / "wide.png")
        Image.new("RGB", (600, 1300), "#222222").save(tmp / "pics" / "phone.png")
        md = tmp / "s.md"
        md.write_text(DOC, encoding="utf-8")
        out = tmp / "s.docx"
        warnings = build(md, out, theme).get("warnings", [])
        expect(not [w for w in warnings if "image" in str(w)], f"every picture found ({warnings})")
        errors, _ = check(out, theme)
        expect(not errors, f"verify passes ({'; '.join(errors) or 'no errors'})")
        d = docx.Document(str(out))

        runs = [r for p in d.paragraphs for r in p.runs]
        chip = next(r for r in runs if r.text == " Active ")
        shd = chip._r.find(f"{qn('w:rPr')}/{qn('w:shd')}")
        expect(shd is not None and shd.get(qn("w:fill")) == "DCF2E3", "Active is a green chip")
        hold = next(r for r in runs if r.text == " Hold ")
        expect(hold._r.find(f"{qn('w:rPr')}/{qn('w:shd')}").get(qn("w:fill")) == "FEF3C7",
               "[[Hold|amber]] takes the colour it names")
        alert = next(r for r in runs if r.text == "no going back")
        expect(str(alert.font.color.rgb) == "C00000", "!!text!! is red")

        fields, screen_t, gallery = d.tables
        head = [fields.rows[1], fields.rows[2]]
        expect(all(r._tr.find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is not None for r in head),
               "both header rows repeat on a new page")
        expect(len(fields.rows[1]._tr.findall(qn("w:tc"))) == 2, "the first header row merges into two groups")
        expect(fields.cell(3, 1).text == "☑" and fields.cell(4, 1).text == "☐",
               "checkboxes in cells")
        pic = pictures(fields.cell(3, 3)._tc)
        grid = [int(g.get(qn("w:w"))) for g in fields._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))]
        expect(len(pic) == 1 and extent_in(pic[0]) <= grid[3] / 1440, "a picture in a cell fits its column")
        expect(len(pictures(fields.cell(4, 3)._tc)) == 1 and fields.cell(4, 3).text.startswith("text"),
               "text and a picture share a cell across <br>")

        left, right = screen_t.rows[0].cells
        expect(len(pictures(left._tc)) == 1, "the screen's picture sits in the left cell")
        expect(len(right.tables) == 1 and right.tables[0].cell(1, 1).text == "Alert",
               "the localisation table nests in the right cell")
        expect(any(p.text.startswith("-") or "Tap Home" in p.text for p in right.paragraphs),
               "the remark list follows in the right cell")

        expect(len(gallery.rows) == 4 and len(gallery.columns) == 2, "three shots two to a row: two caption rows")
        expect(gallery.cell(0, 0).text == "One" and gallery.cell(2, 0).text == "Three", "captions above pictures")
        expect(len(pictures(gallery.cell(1, 1)._tc)) == 1 and not pictures(gallery.cell(3, 1)._tc),
               "an empty slot stays empty")

    if failures:
        print(f"\n{len(failures)} failed")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
