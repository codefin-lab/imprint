#!/usr/bin/env python3
"""Tests for the table-of-contents page numbers (imprint.toc).

    python3 plugins/imprint/engine/tests/test_toc.py          # matching rules, no LibreOffice
    python3 plugins/imprint/engine/tests/test_toc.py --pdf    # also build and render real files

The rendered tests build tests/toc.md (long headings that wrap, Thai headings, a Thai
contents title), check every entry got its real page, then insert a page into the .docx, run
`imprint toc` and check the numbers followed.
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

ENGINE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ENGINE))

from imprint import toc  # noqa: E402

FAILS: list[str] = []


def expect(name: str, got, want) -> None:
    ok = got == want
    print(f"  {'ok' if ok else 'FAILED':18} {name}")
    if not ok:
        FAILS.append(f"{name}: got {got!r}, want {want!r}")


def unit() -> None:
    header = "Example Client            1. SCOPE OF WORK"
    pages = [
        "cover",
        "Contents\n1. Scope.........3\n2. A heading long enough to wrap onto a second line.....4",
        f"{header}\n1. Scope\nbody text",
        f"{header}\n2. A heading long enough to wrap onto a\nsecond line\nbody",
        "5. ชั้นที่ 3: ปรับแต่งตาม Requirement (12 ฟี เจอร์)\nเนื้อหา",
        "APPENDIX A. OPTIONAL MODULES\ntext",
    ]
    heads = ["1. Scope", "2. A heading long enough to wrap onto a second line",
             "5. ชั้นที่ 3: ปรับแต่งตาม Requirement (12 ฟีเจอร์)", "Appendix A. Optional modules",
             "9. Not in the document"]
    got = toc.match_pages(pages, heads, skip_until=2)
    expect("one-line heading", got.get(heads[0]), 3)
    expect("heading that wraps onto two lines", got.get(heads[1]), 4)
    expect("Thai heading with a space pdftotext added", got.get(heads[2]), 5)
    expect("heading set in capitals", got.get(heads[3]), 6)
    expect("missing heading stays missing", heads[4] in got, False)
    expect("a running header is not a heading",
           toc.match_pages([header], ["1. Scope of work"]).get("1. Scope of work"), None)
    expect("part of a line is not a heading",
           toc.match_pages(["1. Scope of work"], ["1. Scope"]), {})

    contents = ["cover", "สารบัญ\n1. ภาพรวม.............4\n2. เนื้อหา..........—",
                "3. ภาคผนวก   9\n4. อื่นๆ ........ 10", "1. ภาพรวม\nเนื้อหา", "body"]
    entries = ["1. ภาพรวม", "2. เนื้อหา", "3. ภาคผนวก", "4. อื่นๆ"]
    expect("contents under a Thai title, over two pages",
           toc.contents_pages(contents, "สารบัญ", entries), (2, 3))
    expect("contents with the default title",
           toc.contents_pages(["x", "TABLE OF CONTENTS\n1. ภาพรวม....4", "1. ภาพรวม"], None,
                              ["1. ภาพรวม"]), (2, 2))
    expect("no contents", toc.contents_pages(["a", "b"], "สารบัญ", entries), (0, 0))


TOC_MD = """---
PROJECT_NAME: "Example Project"
CLIENT_LEGAL_NAME: "Example Client"
VERSION: "0.1"
---

# สารบัญ

<!-- toc -->

# 1. ภาพรวมของระบบ

{para}

# 2. ชั้นที่ 1: ฐานระบบแบบ open source ที่รวมความสามารถหลักไว้ทั้งหมด (19 ฟีเจอร์ ใช้งานได้ทันที)

{para}

# 3. A deliberately long English heading that cannot fit on one line of the page (with notes)

{para}

# 4. Short

{para}
"""


def run(*cmd, cwd=ENGINE) -> str:
    r = subprocess.run([sys.executable, "-m", "imprint.cli", *map(str, cmd)],
                       capture_output=True, text=True, cwd=cwd)
    if r.returncode not in (0, 1):
        FAILS.append(f"{' '.join(map(str, cmd))} exited {r.returncode}:\n{r.stdout}{r.stderr}")
    return r.stdout + r.stderr


def page_of(pdf: Path, heading: str) -> int | None:
    pages = toc.page_texts(pdf)
    _, last = toc.contents_pages(pages, "สารบัญ", [heading])
    return toc.match_pages(pages, [heading], last).get(heading)


def contents_numbers(docx_path: Path) -> dict:
    import docx
    _, paras = toc._contents_paragraphs(docx.Document(str(docx_path)))
    return {p.text.rsplit("\t", 1)[0]: p.text.rsplit("\t", 1)[1] for p in paras}


def rendered() -> None:
    import docx
    out = ENGINE / "tests" / "out" / "toc"
    out.mkdir(parents=True, exist_ok=True)
    para = " ".join(["ย่อหน้าตัวอย่างสำหรับทดสอบสารบัญ Example text for the contents test."] * 40)
    md = out / "toc.md"
    md.write_text(TOC_MD.format(para=para), encoding="utf-8")
    docx_path, pdf = out / "toc.docx", out / "toc.pdf"
    log = run("docx", md, "-o", docx_path, "--pdf")
    expect("imprint docx reports every entry", "4/4 entries" in log and "no page found" not in log, True)

    nums = contents_numbers(docx_path)
    expect("no entry reads a dash", any(v == toc.DASH for v in nums.values()), False)
    for h, n in nums.items():
        expect(f"page of '{h[:40]}'", n, str(page_of(pdf, h)))

    # edit the .docx the way a project script would: push section 3 onto a later page
    d = docx.Document(str(docx_path))
    target = next(p for p in d.paragraphs                 # the heading, not its contents entry
                  if p.text.startswith("3. A deliberately") and "\t" not in p.text)
    for _ in range(2):
        br = target.insert_paragraph_before("")
        br.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
    d.save(str(docx_path))
    before = int(nums[next(h for h in nums if h.startswith("3."))])
    log = run("toc", docx_path, "--pdf")
    after = contents_numbers(docx_path)
    h3 = next(h for h in after if h.startswith("3."))
    expect("imprint toc follows a heading that moved", int(after[h3]) > before, True)
    for h, n in after.items():
        expect(f"after imprint toc, page of '{h[:40]}'", n, str(page_of(pdf, h)))


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--pdf", action="store_true", help="also build and render real files")
    args = ap.parse_args()
    unit()
    if args.pdf:
        rendered()
    for f in FAILS:
        print(f"\nFAIL {f}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    raise SystemExit(main())
