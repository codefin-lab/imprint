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

| Field | State |
| --- | --- |
| Name | [[Signed off|green]] |

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


A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"


def chip_shapes(el) -> dict:
    """label -> (fill, corner adj) for every chip shape under el."""
    out = {}
    for dr in el.iter(qn("w:drawing")):
        geom = dr.find(f".//{A}prstGeom")
        if geom is None or geom.get("prst") != "roundRect":
            continue
        label = dr.find(f".//{WP}docPr").get("descr")
        fill = dr.find(f".//{A}solidFill/{A}srgbClr").get("val")
        adj = geom.find(f".//{A}gd").get("fmla").split()[-1]
        out[label] = (fill, adj)
    return out


def pictures(el) -> list:
    """Pictures (not chip shapes) under el."""
    return [d for d in el.findall(".//" + qn("w:drawing"))
            if d.find(f".//{A}prstGeom[@prst='roundRect']") is None]


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
        def shade(run):
            el = run._r.find(f"{qn('w:rPr')}/{qn('w:shd')}")
            return None if el is None else el.get(qn("w:fill"))
        active = next((r for r in runs if r.text == "\u00a0Active\u00a0"), None)
        expect(active is not None and shade(active) == "DCF2E3" and active.bold,
               "a chip is one shaded run by default, padded with no-break spaces: Active on green")
        hold = next((r for r in runs if r.text == "\u00a0Hold\u00a0"), None)
        expect(hold is not None and shade(hold) == "FEF3C7", "[[Hold|amber]] takes the colour it names")
        cell_runs = [r for t in d.tables for row in t.rows for c in row.cells for p in c.paragraphs for r in p.runs]
        expect(any(r.text.strip("\u00a0") == "Signed off" and shade(r) for r in cell_runs),
               "a [[chip|colour]] in a table cell keeps its pipe out of the columns")
        expect(not chip_shapes(d.element.body), "no drawn shapes for chips by default")
        alert = next(r for r in runs if r.text == "no going back")
        expect(str(alert.font.color.rgb) == "C00000", "!!text!! is red")

        _state, fields, screen_t, gallery = d.tables
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

        # a document adjusts the theme in its own front matter
        md2 = tmp / "o.md"
        md2.write_text("---\ncover: false\nchips: {style: box, radius: 0.2}\ntable: {header_fill: DDEBF7}\n---\n\n"
                       "[[Active]]\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n", encoding="utf-8")
        out2 = tmp / "o.docx"
        stats = build(md2, out2, theme)
        errors, _ = check(out2, stats["theme"])
        expect(not errors and stats["theme"].table["header_fill"] == "DDEBF7"
               and theme.table["header_fill"] != "DDEBF7",
               "front matter table: overrides this document only, and verify uses it")
        d2 = docx.Document(str(out2))
        boxes = chip_shapes(d2.element.body)
        expect(boxes.get("Active", (None, None))[1] == "20000",
               f"chips: {{style: box, radius: 0.2}} draws a rounded box at that radius ({boxes})")
        grows = d2.element.body.find(".//" + A + "spAutoFit")
        expect(grows is not None, "a box chip grows with its text when edited in Word")

    if failures:
        print(f"\n{len(failures)} failed")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
