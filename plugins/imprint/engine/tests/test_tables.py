#!/usr/bin/env python3
"""Table title bands, form tables with merged cells, and callout boxes.

    python3 plugins/imprint/engine/tests/test_tables.py

Run by tests/run.py. Builds one small document and checks the parsed blocks and the .docx XML.
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ENGINE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ENGINE))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from imprint.markdown import parse  # noqa: E402
from imprint.render import _merge_regions, build  # noqa: E402
from imprint.theme import Theme  # noqa: E402
from imprint.verify import check  # noqa: E402

DOC = """\
---
cover: false
---

# TABLES

Table: Version history
| Version | Author |
| :-- | :-- |
| 0.1 | Analyst |

Table: Organization {layout=form}
| **Project name** |  | > | > |
| --- | --- | --- | --- |
| **Name** |  | **Mailing address** |  |
| **Phone** |  | ^ | ^ |
| **Date** |  | **Author** |  |

| Plain | Table |
| --- | --- |
| a > b | ^ not alone |

> [!WARNING] ข้อควรระวัง
> ห้ามส่งข้อมูลส่วนบุคคลทางอีเมล

> [!NOTE]
> Plain note.

> Just a quote.

Table: Details {layout=sections tone=amber}
| Overview | What it is |  |
| --- | --- | --- |
|  |  | 1. First point |
|  |  | 2. Second point |
| Risks | What could go wrong | Too few testers |

Table: Servers {layout=form tone=blue height=0.5}
| **Application servers** | > |
| --- | --- |
| **Operating system** |  |
"""

failures: list[str] = []


def expect(ok: bool, what: str):
    print(f"  {'ok' if ok else 'FAILED':8} {what}")
    if not ok:
        failures.append(what)


def tcs(tr):
    return tr.findall(qn("w:tc"))


def span(tc) -> int:
    g = tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
    return int(g.get(qn("w:val"))) if g is not None else 1


def vmerge(tc) -> str | None:
    v = tc.find(f"{qn('w:tcPr')}/{qn('w:vMerge')}")
    return None if v is None else (v.get(qn("w:val")) or "continue")


def fill(tc) -> str:
    s = tc.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
    return (s.get(qn("w:fill")) or "").lower() if s is not None else ""


def main() -> int:
    blocks = parse(DOC)
    tables = [p for k, p in blocks if k == "table"]
    expect(tables[0][2] == {"title": "Version history"}, "Table: line gives the title")
    expect(tables[1][2] == {"title": "Organization", "layout": "form"}, "Table: attributes are read")
    expect(tables[2][2] == {}, "a table without a Table: line has no options")
    expect(not any(k == "p" and "Table:" in str(p) for k, p in blocks), "the Table: line is not printed")
    callouts = [p for k, p in blocks if k == "callout"]
    expect(callouts == [("warning", "ข้อควรระวัง", "ห้ามส่งข้อมูลส่วนบุคคลทางอีเมล"),
                        ("note", None, "Plain note.")], "callouts parse with kind, title and text")
    expect(("quote", "Just a quote.") in blocks, "a plain quote stays a quote")
    expect(sorted(_merge_regions(tables[1][0], 4)) == [(0, 1, 0, 3), (1, 2, 2, 2), (1, 3, 2, 3)],
           "> widens a cell and ^ under its whole width makes it taller")
    expect(_merge_regions(tables[2][0], 2) == [], "> or ^ inside other text merges nothing")

    theme = Theme.load()
    with tempfile.TemporaryDirectory() as tmp:
        md = Path(tmp) / "t.md"
        md.write_text(DOC, encoding="utf-8")
        out = Path(tmp) / "t.docx"
        build(md, out, theme)
        errors, _ = check(out, theme)
        expect(not errors, f"verify passes ({'; '.join(errors) or 'no errors'})")
        d = docx.Document(str(out))
        t1, t2, t3, c1, c2, sec, srv = d.tables

        band = tcs(t1.rows[0]._tr)
        expect(len(band) == 1 and span(band[0]) == 2 and fill(band[0]) == "1f2937",
               "the title band is one dark cell across the table")
        expect(t1.rows[0].cells[0].text == "VERSION HISTORY", "the title band is in capitals")
        expect(t1.rows[0]._tr.find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is not None
               and t1.rows[1]._tr.find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is not None,
               "title band and header both repeat on a new page")

        rows = [tcs(r._tr) for r in t2.rows]
        expect([span(tc) for tc in rows[1]] == [1, 3], "Project name's value spans three columns")
        expect(vmerge(rows[2][2]) == "restart" and vmerge(rows[3][2]) == "continue",
               "Mailing address runs down two rows")
        expect(span(rows[2][2]) == 1 and vmerge(rows[2][3]) == "restart",
               "the address label and its value merge down separately")
        expect(fill(rows[1][0]) == "efefef" and fill(rows[1][1]) == "",
               "a form shades its labels, not its values")
        expect(t2.rows[1]._tr.find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is None,
               "a form's first row is not a repeating header")
        expect(len(t2.rows[1].cells[1].paragraphs) == 1, "a merged cell gains no empty lines")

        expect(t3.rows[1].cells[0].text == "a > b", "text containing > is left alone")

        bar = c1._tbl.tblPr.find(f"{qn('w:tblBorders')}/{qn('w:left')}")
        expect(bar is not None and bar.get(qn("w:color")) == "D97706", "a warning has an amber bar")
        expect(c1.cell(0, 0).paragraphs[0].text == "ข้อควรระวัง", "a callout title replaces the label")
        expect(c2.cell(0, 0).paragraphs[0].text == "Note", "a callout without a title shows its kind")

        texts = [r.cells[0].text for r in sec.rows]
        expect(texts == ["DETAILS", "OVERVIEW   |   What it is", "1. First point", "2. Second point",
                         "RISKS   |   What could go wrong", "Too few testers"],
               f"sections: band, then boxes; an empty box before more boxes is dropped ({texts})")
        expect(fill(tcs(sec.rows[1]._tr)[0]) == "f8d97c" and fill(tcs(sec.rows[2]._tr)[0]) == "fdf1d0",
               "tone=amber colours the bands and boxes")
        srow = [tcs(r._tr) for r in srv.rows]
        expect(len(srow[1]) == 1 and fill(srow[1][0]) == "adb9ca", "a merged bold row in a form is a group row")
        expect(len(srow[2]) == 2 and fill(srow[2][0]) == "d6dce4" and fill(srow[2][1]) == "f1f4fc",
               "a label beside an empty value stays a label row, toned")
        h = srv.rows[2]._tr.find(f"{qn('w:trPr')}/{qn('w:trHeight')}")
        expect(h is not None and h.get(qn("w:val")) == "720", "height=0.5 sets the least row height")

    if failures:
        print(f"\n{len(failures)} failed")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
