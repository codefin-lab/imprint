#!/usr/bin/env python3
"""Text watermarks: front-matter forms, every header, first page and off/on runs.

    python3 plugins/imprint/engine/tests/test_watermark.py

Run by tests/run.py. Builds small documents and checks the header XML.
"""
from __future__ import annotations

import re
import sys
import tempfile
import zipfile
from pathlib import Path

ENGINE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ENGINE))

import docx  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from imprint.markdown import parse  # noqa: E402
from imprint.render import build  # noqa: E402
from imprint.theme import Theme  # noqa: E402
from imprint.verify import check  # noqa: E402

failures: list[str] = []


def expect(ok: bool, what: str):
    print(f"  {'ok' if ok else 'FAILED':8} {what}")
    if not ok:
        failures.append(what)


def make(tmp: Path, name: str, front: str, body: str):
    md = tmp / f"{name}.md"
    md.write_text(f"---\n{front}\n---\n\n{body}\n", encoding="utf-8")
    out = tmp / f"{name}.docx"
    stats = build(md, out, Theme.load())
    return out, stats


def marked(section, kind="default") -> bool | None:
    """Whether a section's header of `kind` carries the watermark (None: no such header)."""
    ref = section._sectPr.find(f"{qn('w:headerReference')}[@{qn('w:type')}='{kind}']")
    if ref is None:
        return None
    part = section.part.rels[ref.get(qn("r:id"))].target_part
    return "PowerPlusWaterMarkObject" in part.blob.decode() if hasattr(part, "blob") and not hasattr(part, "element") \
        else "PowerPlusWaterMarkObject" in part.element.xml


def main() -> int:
    expect(("watermark", "off") in parse("x\n\n<!-- watermark: off -->\n\ny")
           and ("watermark", "on") in parse("<!-- watermark on -->"), "off / on markers parse")
    theme = Theme.load()
    t, _ = theme.with_overrides({"watermark": "DRAFT"})
    expect(t.watermark.get("text") == "DRAFT", "watermark: DRAFT is the text alone")
    t2, _ = t.with_overrides({"watermark": False})
    expect(not t2.watermark.get("text"), "watermark: false turns it off")
    t3, _ = theme.with_overrides({"watermark": {"text": "X", "angle": -30}})
    expect(t3.watermark.get("angle") == -30, "a watermark mapping sets its options")

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        out, stats = make(tmp, "plain", "cover: false", "# A\n\ntext")
        expect("PowerPlusWaterMarkObject" not in zipfile.ZipFile(out).read("word/header1.xml").decode()
               if "word/header1.xml" in zipfile.ZipFile(out).namelist() else True,
               "no watermark unless asked for")

        body = "# ONE\n\na\n\n<!-- landscape -->\n\n# WIDE\n\nb\n\n<!-- portrait -->\n\n# THREE\n\nc"
        out, stats = make(tmp, "all", "watermark: Confidential", body)
        errors, _ = check(out, stats["theme"])
        expect(not errors, f"verify passes ({errors})")
        d = docx.Document(str(out))
        expect(all(marked(s) for s in d.sections), "every section's running header carries it, landscape too")
        expect(marked(d.sections[0], "first") is True, "and the cover's first-page header")
        xml = d.sections[0].header._element.xml
        m = re.search(r'string="([^"]+)"', xml)
        expect(m and m.group(1) == "Confidential", "the text is the word art's string")

        out, _ = make(tmp, "custom", "cover: false\nwatermark: {text: ฉบับร่าง, color: C00000, opacity: 0.2, angle: -30}",
                      "# A\n\ntext")
        xml = docx.Document(str(out)).sections[0].header._element.xml
        expect("rotation:330" in xml, "angle -30 is rotation 330")
        expect('fillcolor="#F2CCCC"' in xml, "C00000 at 0.2 is blended over white")
        expect('string="ฉบับร่าง"' in xml, "Thai text")

        body = ("# ONE\n\na\n\n\\pagebreak\n\n# TWO\n\nb\n\n<!-- watermark: off -->\n\n# APPENDIX\n\nc\n\n"
                "<!-- landscape -->\n\n# WIDE APPENDIX\n\nd\n\n<!-- portrait -->\n\n<!-- watermark: on -->\n\n# LAST\n\ne")
        out, stats = make(tmp, "runs", "cover: false\nwatermark: {text: DRAFT, first: false}", body)
        errors, _ = check(out, stats["theme"])
        expect(not errors, f"verify passes with first: false and off/on ({errors})")
        d = docx.Document(str(out))
        s = d.sections
        expect(s[0].different_first_page_header_footer and marked(s[0], "first") is False,
               "first: false gives the first page a header without it")
        expect(s[0].first_page_footer._element.xml.count("<w:p") > 0, "the first page keeps a footer")
        expect(marked(s[0]) is True, "pages after the first in that section keep it")
        flags = [marked(x) for x in s]
        expect(flags == [True, False, False, True],
               f"off covers the appendix, landscape included; on brings it back, with no empty section ({flags})")
        texts = [x.header._element.xml.count("<w:t") for x in s]
        expect(texts[1] == texts[0], "a header without the watermark keeps the running header's text")

    if failures:
        print(f"\n{len(failures)} failed")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
