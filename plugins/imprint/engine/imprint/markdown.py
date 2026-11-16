"""A deliberately small Markdown reader — only what proposals actually use.

Returns a flat list of (kind, payload) blocks:

    ('h',         (level, text))
    ('p',         (indent_level, [(text, forced_break), ...]))  one entry per
                  source line; indent_level counts leading whitespace
    ('ul' | 'ol', [(level, 'ul' | 'ol', item, number), ...])  level 0 = top;
                  nest by indenting.  `number` is what an ordered item was
                  numbered in the source — a list starting at 2 continues at 2.
    ('table',     ([[cell, ...], ...], [align, ...], opts))  first row is the header;
                  one align per column from the separator row: 'left' (:--),
                  'right' (--:), 'center' (:-:) or None (---, the style's own).
                  opts comes from a `Table: Title {layout=form}` line right above:
                  {'title': 'Title', 'layout': 'form'}, or {} without one.
                  A cell that is only `>` joins the cell to its left, only `^` the
                  cell above
    ('gantt',     spec text)
    ('image',     (src, caption))   a line that is only ![caption](file.png)
    ('quote',     text)             consecutive lines starting with >, joined by newlines
    ('callout',   (kind, title, text))  a quote whose first line is [!NOTE], [!TIP],
                  [!IMPORTANT], [!WARNING] or [!CAUTION]; kind is lower case,
                  title is any text after the marker or None
    ('screen',    (src, [block, ...]))  `::: screen file.png` up to `:::`: a screenshot
                  beside the blocks written inside (paragraphs, tables, lists)
    ('screens',   ([(src, caption), ...], opts))  `::: screens {per_row=3}` up to
                  `:::`, one ![caption](file.png) per line: screenshots side by side
    ('orient',    'landscape' | 'portrait')
    ('hr',        None)
    ('pagebreak', None)
    ('toc',       None)

Keeping the grammar this small is the point: what the team writes maps
one-to-one onto what comes out, with no surprises to debug.
"""
from __future__ import annotations

import re

import yaml

TABLE_SEP = re.compile(r"^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?$")
HEADING = re.compile(r"^(#{1,6})\s+(.*)$")


def _column_aligns(separator: str) -> list[str | None]:
    """Read each column's alignment off a table's separator row, the standard
    Markdown way: `:--` left, `--:` right, `:-:` center, `---` unset."""
    out = []
    for cell in separator.strip("|").split("|"):
        cell = cell.strip()
        left, right = cell.startswith(":"), cell.endswith(":")
        out.append("center" if left and right else "right" if right else "left" if left else None)
    return out
# •, ·, ▪, ◦, ‣ come in with content pasted out of Word or Google Docs, where
# the bullet is a literal character rather than a list.  Treating them as
# markers is what stops that content rendering as flat, unindented paragraphs.
BULLET = re.compile(r"^[-*+\u2022\u00b7\u25aa\u25e6\u2023]\s+")
ORDERED = re.compile(r"^(\d+)[.)]\s+")
RULE = re.compile(r"^(-{3,}|\*{3,}|_{3,})$")
# ขึ้นหน้าใหม่ — เขียนแบบไหนก็ได้:  \pagebreak  \newpage  <!-- pagebreak -->  <!-- pb -->
ORIENT = re.compile(r"^(?:\\(?P<b>landscape|portrait)|<!--\s*(?P<c>landscape|portrait)\s*-->)$", re.I)
PAGEBREAK = re.compile(r"^(?:\\(?:pagebreak|newpage)|<!--\s*(?:pagebreak|pb|page-break)\s*-->)$", re.I)
TOC = re.compile(r"^(?:\\toc|<!--\s*toc\s*-->)$", re.I)
# `Table: Version history` on the line right above a table: a title band across its top
TABLE_TITLE = re.compile(r"^Table:\s+(?P<title>\S.*)$")
# GitHub's alert syntax: the first line of a quote names the kind of box
CALLOUT = re.compile(r"^\[!(?P<kind>note|tip|important|warning|caution)\]\s*(?P<title>.*)$", re.I)
CALLOUT_KINDS = ("note", "tip", "important", "warning", "caution")

FENCE = re.compile(r"^```+\s*(\w+)?\s*$")
# A picture on a line of its own.  The alt text is the caption — the figure's
# title belongs to the document, never baked into the image.
IMAGE = re.compile(r'^!\[(?P<alt>[^\]]*)\]\((?P<src>[^)\s]+)(?:\s+"[^"]*")?\)$')

# [[Active]] or [[Active|green]] is a chip, !!text!! is alert-red text, [x] / [ ] a checkbox
INLINE = re.compile(r"(\[\[[^\]\n]+?\]\]|!![^!\n]+?!!|\[[ xX]\]|\*\*.+?\*\*|(?<!\*)\*[^*]+?\*|`[^`]+?`)", re.S)
# a fenced div, Pandoc style: `::: screen shot.png` ... `:::`
DIV_OPEN = re.compile(r"^:{3,}\s*(?P<kind>screens?)\b\s*(?P<rest>.*)$", re.I)
DIV_CLOSE = re.compile(r"^:{3,}\s*$")
# a picture written inside a table cell, alone or between <br>s
CELL_IMAGE = re.compile(r'^!\[(?P<alt>[^\]]*)\]\((?P<src>[^)\s]+)(?:\s+"[^"]*")?\)$')
CHECKBOX = {"[x]": "\u2611", "[X]": "\u2611", "[ ]": "\u2610"}


ATTRS = re.compile(r"\s*\{((?:\s*[A-Za-z_][\w-]*=[^\s}]+)+)\s*\}\s*$")


_CELL_SPLIT = re.compile(r"\[\[[^\]\n]*\]\]|\\\||\|")


def _cells(row: str) -> list[str]:
    """Split a table row on `|`, except inside a [[chip|colour]] or written `\|`."""
    row = row.strip()
    row = row[1:] if row.startswith("|") else row
    row = row[:-1] if row.endswith("|") and not row.endswith("\\|") else row
    cells, start = [], 0
    for m in _CELL_SPLIT.finditer(row):
        if m.group() == "|":
            cells.append(row[start:m.start()])
            start = m.end()
    cells.append(row[start:])
    return [c.strip().replace("\\|", "|") for c in cells]


def split_attrs(text: str) -> tuple[str, dict]:
    """`Title {icon=wallet}` into ("Title", {"icon": "wallet"}): Pandoc-style attributes on a heading."""
    m = ATTRS.search(text or "")
    if not m:
        return text, {}
    pairs = dict(kv.split("=", 1) for kv in m.group(1).split())
    return text[: m.start()].rstrip(), {k.lower(): v.strip("\"'") for k, v in pairs.items()}


def split_front_matter(text: str) -> tuple[dict, str]:
    """Pull the leading `---` YAML block off, if there is one."""
    if text.startswith("---\n"):
        end = text.find("\n---", 4)
        if end != -1:
            return yaml.safe_load(text[4:end]) or {}, text[end + 4:].lstrip("\n")
    return {}, text


def parse(md: str, *, indent_spaces: int = 4) -> list[tuple]:
    lines = md.replace("\r\n", "\n").split("\n")
    blocks: list[tuple] = []
    para: list[str] = []
    para_indent = 0
    table_opts: dict = {}
    i = 0

    def flush():
        nonlocal para_indent
        if para:
            blocks.append(("p", (para_indent, list(para))))
            para.clear()
            para_indent = 0

    while i < len(lines):
        raw = lines[i]
        line = raw.strip()

        if m := FENCE.match(line):
            flush()
            lang = (m.group(1) or "").lower()
            i += 1
            content = []
            while i < len(lines) and not FENCE.match(lines[i].strip()):
                content.append(lines[i])
                i += 1
            i += 1                       # closing fence
            # a chart or widget language keeps its own kind; every other fence, with or
            # without a language tag (json, bash, http), is a code block
            blocks.append((lang if lang in ("gantt", "markwhen", "widget") else "code", "\n".join(content)))
            continue
        elif m := DIV_OPEN.match(line):
            flush()
            kind = m.group("kind").lower()
            rest, opts = split_attrs(m.group("rest").strip())
            i += 1
            inner = []
            while i < len(lines) and not DIV_CLOSE.match(lines[i].strip()):
                inner.append(lines[i])
                i += 1
            i += 1                       # closing :::
            if kind == "screen":
                blocks.append(("screen", (rest.strip(), parse("\n".join(inner), indent_spaces=indent_spaces))))
            else:
                shots = [(im.group("src"), im.group("alt").strip())
                         for im in (IMAGE.match(x.strip()) for x in inner) if im]
                blocks.append(("screens", (shots, opts)))
            continue
        elif not line:
            flush()
        elif TOC.match(line):
            flush()
            blocks.append(("toc", None))
        elif PAGEBREAK.match(line):
            flush()
            blocks.append(("pagebreak", None))
        elif m := ORIENT.match(line):
            flush()
            blocks.append(("orient", (m.group("b") or m.group("c")).lower()))
        elif RULE.match(line):
            flush()
            blocks.append(("hr", None))
        elif m := HEADING.match(line):
            flush()
            blocks.append(("h", (len(m.group(1)), m.group(2).strip())))
        elif ((m := TABLE_TITLE.match(line)) and not para and i + 2 < len(lines)
              and lines[i + 1].strip().startswith("|") and TABLE_SEP.match(lines[i + 2].strip())):
            title, attrs = split_attrs(m.group("title").strip())
            table_opts = {**attrs, "title": title or None}
        elif line.startswith("|") and i + 1 < len(lines) and TABLE_SEP.match(lines[i + 1].strip()):
            flush()
            rows = []
            aligns = _column_aligns(lines[i + 1].strip())
            while i < len(lines) and lines[i].strip().startswith("|"):
                row = lines[i].strip()
                if not TABLE_SEP.match(row):
                    rows.append(_cells(row))
                i += 1
            blocks.append(("table", (rows, aligns, table_opts)))
            table_opts = {}
            continue
        elif line.startswith(">"):
            flush()
            quote = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                quote.append(lines[i].strip()[1:].strip())
                i += 1
            if quote and (m := CALLOUT.match(quote[0])):
                blocks.append(("callout", (m.group("kind").lower(), m.group("title").strip() or None,
                                           "\n".join(quote[1:]).strip("\n"))))
            else:
                blocks.append(("quote", "\n".join(quote)))
            continue
        elif m := IMAGE.match(line):
            flush()
            blocks.append(("image", (m.group("src"), m.group("alt").strip())))
        elif BULLET.match(line) or ORDERED.match(line):
            flush()
            blocks.append(_parse_list(lines, i))
            i = _list_end(lines, i)
            continue
        else:
            # "two trailing spaces" and "trailing backslash" are the standard
            # Markdown ways to force a break; honour them in either mode
            if not para:                       # indent comes from the first line
                lead = raw[: len(raw) - len(raw.lstrip(" \t"))]
                para_indent = len(lead.expandtabs(4)) // indent_spaces
            forced = raw.rstrip("\n").endswith("  ") or line.endswith("\\")
            para.append((line.rstrip("\\").rstrip(), forced))
        i += 1

    flush()
    return blocks


def _list_marker(raw: str):
    """(indent, kind, text, number) for a list line, or None.

    `number` is what the author actually typed for an ordered item.  A list that
    starts at something other than 1 continues from there, which is how you carry
    numbering across a table or a paragraph that interrupts the list.
    """
    stripped = raw.lstrip(" \t")
    if BULLET.match(stripped):
        kind, text, number = "ul", BULLET.sub("", stripped), None
    elif m := ORDERED.match(stripped):
        kind, text, number = "ol", ORDERED.sub("", stripped), int(m.group(1))
    else:
        return None
    indent = len(raw[: len(raw) - len(stripped)].expandtabs(4))
    return indent, kind, text.strip(), number


def _list_end(lines, i):
    while i < len(lines) and _list_marker(lines[i]):
        i += 1
    return i


def _parse_list(lines, i):
    """Consume consecutive list lines. Nesting follows indentation: every
    deeper indent opens a level, returning to a shallower indent closes it.
    Bullets and numbers may be mixed across levels."""
    items = []
    stack: list[int] = []
    while i < len(lines) and (m := _list_marker(lines[i])):
        indent, kind, text, number = m
        while stack and indent < stack[-1]:
            stack.pop()
        if not stack or indent > stack[-1]:
            stack.append(indent)
        items.append((len(stack) - 1, kind, text, number))
        i += 1
    return items[0][1], items


def rich_spans(text: str):
    """Yield (text, bold, italic, mono, mark) for each span.  `mark` is None for
    plain text, ('chip', colour or None) for [[Label]] / [[Label|colour]],
    ('alert', None) for !!text!!, and ('box', None) for a [x] / [ ] checkbox,
    whose text is then the ballot-box character."""
    for piece in INLINE.split(text):
        if not piece:
            continue
        if piece.startswith("[[") and piece.endswith("]]"):
            label, _, colour = piece[2:-2].partition("|")
            yield label.strip(), False, False, False, ("chip", colour.strip().lower() or None)
        elif piece in CHECKBOX:
            yield CHECKBOX[piece], False, False, False, ("box", None)
        elif piece.startswith("!!") and piece.endswith("!!") and len(piece) > 4:
            yield piece[2:-2], False, False, False, ("alert", None)
        elif piece.startswith("**") and piece.endswith("**"):
            yield piece[2:-2], True, False, False, None
        elif piece.startswith("*") and piece.endswith("*"):
            yield piece[1:-1], False, True, False, None
        elif piece.startswith("`") and piece.endswith("`"):
            yield piece[1:-1], False, False, True, None
        else:
            yield piece, False, False, False, None


def inline_spans(text: str):
    """Yield (text, bold, italic, mono) for each formatted span; chips, alert
    text and checkboxes come through as their plain text."""
    for chunk, b, i, mono, _mark in rich_spans(text):
        yield chunk, b, i, mono
