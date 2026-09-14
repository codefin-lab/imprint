"""Slide layouts beyond text and a visual.

    statement   one sentence, large, beside an accent bar          (a > quote)
    cards       three to six short blocks in a grid                 (### groups)
    compare     two columns under headers                           (two ### groups)
    steps       boxes joined by arrows                              (a list, or ### groups)
    numbered    a numbered list, with a callout beside it           (a list, then a > quote)
    stats       two to four large numbers with labels               (- **42%** label)
    media       text on one side, a picture on the other            (media, or media-left)
    closing     a large closing line and contact details
    gallery     two to six pictures in a row, captioned             (pictures only)
    logos       a grid of logos on light tiles

A slide gets a layout from `<!-- layout: name -->` under its heading, or, when the
content has an unmistakable shape, from `detect`. `<!-- tone: dark -->` puts any of
them on the theme's dark slide. Every colour and size comes from the theme's
`slides:` section (`components`, `dark`, `size`), never from here.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path

from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt

LAYOUTS = ("statement", "cards", "compare", "steps", "numbered", "stats", "media", "media-left",
           "closing", "gallery", "logos")
STAT = re.compile(r"^\*\*([^*]+)\*\*\s*(.*)$")
NUMBERED_TITLE = re.compile(r"^(\d+)[.)]?\s+(.*)$")
EMU_IN = 914400


def _s():
    from . import slides     # slides imports this module; import back lazily
    return slides


# ------------------------------------------------------------------ tone

@dataclass
class Tone:
    dark: bool
    fg: str
    sub: str
    card_fill: str
    card_border: str
    accent: str
    on_accent: str
    accents: list = field(default_factory=list)
    radius: float = 0.08
    sizes: dict = field(default_factory=dict)

    def strip(self, i: int) -> str | None:
        return self.accents[i % len(self.accents)] if self.accents else None


def tone_for(deck, dark: bool) -> Tone:
    s = deck.theme.slides or {}
    comp = s.get("components", {}) or {}
    dk = s.get("dark", {}) or {}
    size = s.get("size", {}) or {}
    sizes = {"statement": float(size.get("statement", 30)), "stat": float(size.get("stat", 40)),
             "card_title": float(size.get("card_title", 16)), "card_body": float(size.get("card_body", 13)),
             "step": float(size.get("step", 14)), "closing": float(size.get("closing", 40))}
    accents = [str(a).lstrip("#") for a in (comp.get("accents") or [])]
    radius = float(comp.get("card_radius_in", 0.08))
    if dark:
        return Tone(True, str(dk.get("text", deck.paper)), str(dk.get("sub", deck.section_sub)),
                    str(dk.get("card_fill", "2D3748")), str(dk.get("card_border", "4A5568")),
                    str(dk.get("accent", deck.paper)), str(dk.get("on_accent", deck.ink)),
                    accents, radius, sizes)
    return Tone(False, deck.ink, deck.soft, str(comp.get("card_fill", "F7FAFC")),
                str(comp.get("card_border", "E2E8F0")), str(comp.get("accent", deck.ink)), deck.paper,
                accents, radius, sizes)


# ------------------------------------------------------------------ reading the content

def split_groups(blocks):
    """(intro, [(title, blocks)]) when the content is ### groups, else None."""
    first = next((i for i, (k, p) in enumerate(blocks) if k == "h" and p[0] == 3), None)
    if first is None:
        return None
    intro = blocks[:first]
    if any(k not in ("p", "quote") for k, _ in intro):
        return None
    groups = []
    for kind, payload in blocks[first:]:
        if kind == "h" and payload[0] == 3:
            groups.append((payload[1], []))
        elif kind == "h":
            return None
        else:
            groups[-1][1].append((kind, payload))
    return intro, groups


def stat_items(list_block):
    """[(value, label)] when a list is two to four `**value** label` items with a digit in each value."""
    if list_block[0] not in ("ul", "ol"):
        return None
    items = [it for it in list_block[1] if it[0] == 0]
    out = []
    for _, _, text, _ in items:
        m = STAT.match(text)
        if not m or not re.search(r"\d", m.group(1)):
            return None
        out.append((m.group(1).strip(), m.group(2).strip()))
    return out if 2 <= len(out) <= 4 else None


def detect(blocks) -> str | None:
    """A layout for content whose shape leaves no doubt; None keeps the plain slide."""
    if not blocks:
        return None
    kinds = [k for k, _ in blocks]
    if "quote" in kinds and all(k in ("quote", "p") for k in kinds) and kinds[0] == "quote":
        return "statement"
    grouped = split_groups(blocks)
    if grouped is not None:
        n = len(grouped[1])
        if n == 2:
            return "compare"
        if 3 <= n <= 6:
            return "cards"
    lists = [b for b in blocks if b[0] in ("ul", "ol")]
    if len(lists) == 1 and all(k in ("ul", "ol", "p") for k in kinds) and stat_items(lists[0]):
        return "stats"
    if kinds.count("image") >= 2 and all(k in ("image", "p") for k in kinds):
        return "gallery"
    return None


def _para_text(blocks) -> str:
    out = []
    for kind, payload in blocks:
        if kind == "p":
            out.append(" ".join(t.strip() for t, _ in payload[1] if t.strip()))
        elif kind == "quote":
            out.append(payload)
    return "\n".join(t for t in out if t)


def _lead_split(text: str) -> tuple[str, str]:
    """`**Title** description` into its parts; plain text is all title."""
    m = STAT.match(text)
    return (m.group(1).strip(), m.group(2).strip()) if m else (text.strip(), "")


# ------------------------------------------------------------------ drawing

def _plain(shape):
    """No theme style: LibreOffice would add its effects (a shadow) to a filled shape."""
    style = shape._element.find(qn("p:style"))
    if style is not None:
        shape._element.remove(style)
    return shape


def _box(slide, x, y, w, h, *, fill=None, line=None, radius=None, line_w=1.0):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shp = _plain(slide.shapes.add_shape(kind, int(x), int(y), int(w), int(h)))
    S = _s()
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = S._rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = S._rgb(line)
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    if radius:
        shp.adjustments[0] = max(0.0, min(0.5, Inches(radius) / max(1, min(int(w), int(h)))))
    return shp


def _lines(text: str, size_pt: float, width_emu: int) -> int:
    per_line = max(4, (width_emu / EMU_IN) * 72 / (size_pt * 0.5))
    return sum(max(1, math.ceil(_s()._visible_len(re.sub(r"[*`]", "", part)) / per_line))
               for part in (text.split("\n") or [""]))


def _height(paras, width_emu: int) -> int:
    """EMU for [(text, size, ...)] paragraphs set in `width_emu`."""
    return int(sum(_lines(p[0], p[1], width_emu) * p[1] * 1.28 / 72 + 4 / 72 for p in paras) * EMU_IN)


def _text(slide, x, y, w, h, paras, deck, *, anchor=MSO_ANCHOR.TOP, align=None, space=4):
    """paras: [(text, size_pt, colour, bold, font or None)]"""
    S = _s()
    tb = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for k, (text, size, colour, bold, font) in enumerate(paras):
        for j, part in enumerate(text.split("\n")):
            p = tf.paragraphs[0] if (k == 0 and j == 0) else tf.add_paragraph()
            S._para_format(p, space_after_pt=space, align=align)
            S._runs(p, part, deck, size, colour, bold=bold, font=font)
    return tb


def _items_box(slide, x, y, w, h, items, deck, tone, size):
    """Paragraphs and list items from _text_items, in `size`, in the tone's colour."""
    S = _s()
    tb = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for k, (kind, text, level, number) in enumerate(items):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        if kind == "ul":
            S._para_format(p, level=level, bullet=deck.bullet, space_after_pt=4,
                           bullet_font=deck.body_font, bullet_color=tone.fg)
        elif kind == "ol":
            S._para_format(p, level=level, number=number or 1, space_after_pt=4,
                           bullet_font=deck.body_font, bullet_color=tone.fg)
        else:
            S._para_format(p, space_after_pt=6)
        S._runs(p, text, deck, size, tone.fg, bold=kind == "lead")
    return tb


def _items_height(items, width_emu, size) -> int:
    return _height([(t, size) for _, t, _, _ in items], width_emu - Inches(0.32))


def _arrow(slide, x1, y1, x2, y2, colour):
    S = _s()
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, int(x1), int(y1), int(x2), int(y2))
    _plain(conn)
    conn.line.color.rgb = S._rgb(colour)
    conn.line.width = Pt(1.5)
    ln = conn.line._get_or_add_ln()
    tail = OxmlElement("a:tailEnd")
    tail.set("type", "triangle")
    ln.append(tail)
    return conn


def _intro(slide, box, intro_blocks, deck, tone, warnings, n, title):
    """An optional sentence above the layout; returns the box that is left."""
    x, y, w, h = box
    text = _para_text(intro_blocks)
    if not text:
        return box
    size = deck.levels_pt[0]
    need = min(_height([(text, size)], w), int(h * 0.3))
    _text(slide, x, y, w, need, [(text, size, tone.fg, False, None)], deck)
    gap = Inches(0.25)
    return x, y + need + gap, w, h - need - gap


def _warn_fit(warnings, n, title, what, need, have):
    if need > have:
        warnings.append(f"slide {n} ({title}): {what} needs about {need / EMU_IN:.1f} in, "
                        f"has {have / EMU_IN:.1f}; shorten or split")


# ------------------------------------------------------------------ the layouts

def statement(ctx):
    slide, deck, tone, (x, y, w, h) = ctx["slide"], ctx["deck"], ctx["tone"], ctx["box"]
    blocks = ctx["blocks"]
    quote = "\n".join(p for k, p in blocks if k == "quote")
    source = _para_text([b for b in blocks if b[0] == "p"])
    size = tone.sizes["statement"]
    tw = int(w * 0.82)
    qh = _height([(quote, size)], tw)
    sh = _height([(source, deck.levels_pt[0] - 2)], tw) if source else 0
    total = qh + (sh + Inches(0.25) if source else 0)
    top = y + max(0, (h - total) // 2)
    _box(slide, x, top, Inches(0.08), qh, fill=tone.accent)
    _text(slide, x + Inches(0.4), top, tw, qh, [(quote, size, tone.fg, False, deck.heading_font)], deck, space=6)
    if source:
        _text(slide, x + Inches(0.4), top + qh + Inches(0.25), tw, sh,
              [(source, deck.levels_pt[0] - 2, tone.sub, False, None)], deck)
    _warn_fit(ctx["warnings"], ctx["n"], ctx["title"], "the statement", total, h)


def cards(ctx):
    slide, deck, tone = ctx["slide"], ctx["deck"], ctx["tone"]
    grouped = split_groups(ctx["blocks"])
    if not grouped:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): cards need ### headings, one per card")
        return
    intro, groups = grouped
    x, y, w, h = _intro(slide, ctx["box"], intro, deck, tone, ctx["warnings"], ctx["n"], ctx["title"])
    count = len(groups)
    if count > 6:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): {count} cards; six is the most that reads")
    cols = count if count <= 4 else 3
    rows = math.ceil(count / cols)
    gap = deck.gutter
    cw = (w - gap * (cols - 1)) // cols
    pad = Inches(0.22)
    tw = cw - 2 * pad
    S = _s()
    numbered = all(NUMBERED_TITLE.match(t) for t, _ in groups)
    sq = Inches(0.36)
    # a number tile carries the card's colour, so it replaces the strip
    head = (sq + Inches(0.12)) if numbered else (Inches(0.2) if tone.accents else 0)
    prepared = []
    for title, body in groups:
        num = None
        if numbered:
            num, title = NUMBERED_TITLE.match(title).groups()
        items = S._text_items(body)
        th = _height([(title, tone.sizes["card_title"])], tw)
        bh = _items_height(items, tw, tone.sizes["card_body"]) if items else 0
        prepared.append((num, title, items, th, bh))
    need = max(pad * 2 + head + th + Inches(0.08) + bh for _, _, _, th, bh in prepared)
    room = (h - gap * (rows - 1)) // rows
    ch = min(room, max(need, Inches(1.5)))
    if need > room:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): card text needs about "
                               f"{need / EMU_IN:.1f} in, a card has {room / EMU_IN:.1f}; shorten or split")
    top = y + max(0, (h - (ch * rows + gap * (rows - 1))) // 2)
    for i, (num, title, items, th, bh) in enumerate(prepared):
        cx, cy = x + (i % cols) * (cw + gap), top + (i // cols) * (ch + gap)
        _box(slide, cx, cy, cw, ch, fill=tone.card_fill, line=tone.card_border, radius=tone.radius, line_w=0.75)
        inner_y = cy + pad
        strip = tone.strip(i)
        tx = cx + pad
        if numbered:
            _box(slide, tx, inner_y, sq, sq, fill=strip or tone.accent, radius=0.04)
            _text(slide, tx, inner_y, sq, sq, [(num, tone.sizes["card_title"] - 2,
                  "FFFFFF" if strip else tone.on_accent, True, deck.heading_font)],
                  deck, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, space=0)
            inner_y += sq + Inches(0.12)
        elif strip:
            _box(slide, tx, inner_y, Inches(0.5), Inches(0.06), fill=strip)
            inner_y += Inches(0.2)
        _text(slide, tx, inner_y, tw, th, [(title, tone.sizes["card_title"], tone.fg, True, deck.heading_font)], deck)
        if items:
            by = inner_y + th + Inches(0.08)
            _items_box(slide, tx, by, tw, cy + ch - pad - by, items, deck, tone, tone.sizes["card_body"])


def compare(ctx):
    slide, deck, tone = ctx["slide"], ctx["deck"], ctx["tone"]
    grouped = split_groups(ctx["blocks"])
    if not grouped or len(grouped[1]) != 2:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): compare needs exactly two ### headings")
        return
    intro, groups = grouped
    x, y, w, h = _intro(slide, ctx["box"], intro, deck, tone, ctx["warnings"], ctx["n"], ctx["title"])
    gap = Inches(0.5)
    cw = (w - gap) // 2
    hh = Inches(0.6)
    size = max(12.0, deck.levels_pt[0] - 2)
    S = _s()
    for i, (title, body) in enumerate(groups):
        cx = x + i * (cw + gap)
        edge = tone.strip(i) or (tone.accent if i == 0 else tone.sub)
        _box(slide, cx, y, cw, hh, line=edge, radius=tone.radius, line_w=2)
        _text(slide, cx, y, cw, hh, [(title, tone.sizes["card_title"] + 2, tone.fg, True, deck.heading_font)],
              deck, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, space=0)
        items = S._text_items(body)
        top = y + hh + Inches(0.3)
        _items_box(slide, cx + Inches(0.1), top, cw - Inches(0.2), y + h - top, items, deck, tone, size)
        _warn_fit(ctx["warnings"], ctx["n"], ctx["title"], f"column {i + 1}",
                  hh + Inches(0.3) + _items_height(items, cw, size), h)


def _step_items(blocks):
    grouped = split_groups(blocks)
    if grouped:
        return grouped[0], [(t, _para_text(b) or " ".join(it[2] for it in
                             next((p for k, p in b if k in ("ul", "ol")), []))) for t, b in grouped[1]]
    lists = [p for k, p in blocks if k in ("ul", "ol")]
    if not lists:
        return [], []
    intro = [b for b in blocks if b[0] in ("p", "quote")]
    return intro, [_lead_split(it[2]) for it in lists[0] if it[0] == 0]


def steps(ctx):
    slide, deck, tone = ctx["slide"], ctx["deck"], ctx["tone"]
    intro, items = _step_items(ctx["blocks"])
    if not 2 <= len(items) <= 6:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): steps need two to six items, got {len(items)}")
        if not items:
            return
        items = items[:6]
    x, y, w, h = _intro(slide, ctx["box"], intro, deck, tone, ctx["warnings"], ctx["n"], ctx["title"])
    count = len(items)
    gap = Inches(0.5)
    bw = (w - gap * (count - 1)) // count
    size = tone.sizes["step"]
    bh = max(Inches(0.8), max(_height([(t, size)], bw - Inches(0.3)) for t, _ in items) + Inches(0.3))
    desc_size = max(11.0, size - 2)
    dh = max((_height([(d, desc_size)], bw) for _, d in items if d), default=0)
    band = bh + (dh + Inches(0.2) if dh else 0)
    top = y + max(0, (h - band) // 2)
    for i, (title, desc) in enumerate(items):
        bx = x + i * (bw + gap)
        fill = tone.strip(i) or tone.accent
        text_colour = "FFFFFF" if tone.strip(i) else tone.on_accent
        _box(slide, bx, top, bw, bh, fill=fill, radius=tone.radius)
        _text(slide, bx + Inches(0.15), top, bw - Inches(0.3), bh, [(title, size, text_colour, True, deck.heading_font)],
              deck, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, space=0)
        if desc:
            _text(slide, bx, top + bh + Inches(0.2), bw, dh, [(desc, desc_size, tone.sub, False, None)],
                  deck, align=PP_ALIGN.CENTER)
        if i < count - 1:
            _arrow(slide, bx + bw + Inches(0.08), top + bh // 2, bx + bw + gap - Inches(0.08), top + bh // 2, tone.sub)
    _warn_fit(ctx["warnings"], ctx["n"], ctx["title"], "the steps", band, h)


def numbered(ctx):
    slide, deck, tone, (x, y, w, h) = ctx["slide"], ctx["deck"], ctx["tone"], ctx["box"]
    blocks = ctx["blocks"]
    lists = [p for k, p in blocks if k in ("ul", "ol")]
    if not lists:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): numbered needs a list")
        return
    items = [_lead_split(it[2]) for it in lists[0] if it[0] == 0]
    first_list = next(i for i, b in enumerate(blocks) if b[0] in ("ul", "ol"))
    callout = _para_text([b for b in blocks[first_list + 1:] if b[0] in ("p", "quote")])
    intro = _para_text([b for b in blocks[:first_list] if b[0] in ("p", "quote")])
    if intro:
        x, y, w, h = _intro(slide, (x, y, w, h), [("p", (0, [(intro, False)]))], deck, tone,
                            ctx["warnings"], ctx["n"], ctx["title"])
    lw = int(w * 0.62) if callout else w
    count = len(items)
    if count > 6:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): {count} numbered items; keep to six")
    row = min(Inches(1.0), h // max(1, count))
    sq = Inches(0.42)
    size = tone.sizes["card_title"]
    for i, (title, desc) in enumerate(items):
        ry = y + i * row
        fill = tone.strip(i) or tone.accent
        _box(slide, x, ry, sq, sq, fill=fill, radius=0.04)
        _text(slide, x, ry, sq, sq, [(str(i + 1), size - 2, "FFFFFF" if tone.strip(i) else tone.on_accent, True,
              deck.heading_font)], deck, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, space=0)
        paras = [(title, size, tone.fg, True, deck.heading_font)]
        if desc:
            paras.append((desc, tone.sizes["card_body"], tone.sub, False, None))
        _text(slide, x + sq + Inches(0.2), ry - Inches(0.02), lw - sq - Inches(0.2), row, paras, deck, space=2)
    if callout:
        cx = x + lw + Inches(0.4)
        cw = w - lw - Inches(0.4)
        csize = tone.sizes["card_title"]
        chh = _height([(callout, csize)], cw - Inches(0.5)) + Inches(0.5)
        edge = tone.strip(len(items)) or tone.accent
        _box(slide, cx, y, cw, chh, line=edge, radius=tone.radius, line_w=2)
        _text(slide, cx + Inches(0.25), y, cw - Inches(0.5), chh, [(callout, csize, tone.fg, True, deck.heading_font)],
              deck, anchor=MSO_ANCHOR.MIDDLE, space=4)
    _warn_fit(ctx["warnings"], ctx["n"], ctx["title"], "the list", row * count, h)


def stats(ctx):
    slide, deck, tone = ctx["slide"], ctx["deck"], ctx["tone"]
    blocks = ctx["blocks"]
    lst = next((b for b in blocks if b[0] in ("ul", "ol")), None)
    values = stat_items(lst) if lst else None
    if not values:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): stats need two to four `**number** label` items")
        return
    first = blocks.index(lst)
    x, y, w, h = _intro(slide, ctx["box"], blocks[:first], deck, tone, ctx["warnings"], ctx["n"], ctx["title"])
    count = len(values)
    gap = deck.gutter
    cw = (w - gap * (count - 1)) // count
    vs, ls = tone.sizes["stat"], max(12.0, tone.sizes["card_body"] + 1)
    lh = max(_height([(label, ls)], cw - Inches(0.5)) for _, label in values)
    ch = min(h, int(vs * 1.3 / 72 * EMU_IN) + lh + Inches(0.9))
    top = y + max(0, (h - ch) // 2)
    for i, (value, label) in enumerate(values):
        cx = x + i * (cw + gap)
        _box(slide, cx, top, cw, ch, fill=tone.card_fill, line=tone.card_border, radius=tone.radius, line_w=0.75)
        strip = tone.strip(i)
        if strip:
            _box(slide, cx, top, cw, Inches(0.07), fill=strip)
        vh = int(vs * 1.3 / 72 * EMU_IN)
        _text(slide, cx + Inches(0.25), top + Inches(0.35), cw - Inches(0.5), vh,
              [(value, vs, strip or tone.accent, True, deck.heading_font)], deck, space=0)
        _text(slide, cx + Inches(0.25), top + Inches(0.35) + vh + Inches(0.1), cw - Inches(0.5), lh,
              [(label, ls, tone.sub, False, None)], deck)


def media(ctx, image_left=False):
    slide, deck, tone, (x, y, w, h) = ctx["slide"], ctx["deck"], ctx["tone"], ctx["box"]
    S = _s()
    blocks = ctx["blocks"]
    image = next((p for k, p in blocks if k == "image"), None)
    items = S._text_items([b for b in blocks if b[0] in ("p", "ul", "ol", "h", "quote")])
    if image is None:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): media needs a picture")
        return
    gap = Inches(0.5)
    tw = int(w * 0.42)
    iw = w - tw - gap
    tx, ix = (x + iw + gap, x) if image_left else (x, x + tw + gap)
    size = deck.levels_pt[0]
    need = _items_height(items, tw, size)
    ty = y + max(0, (h - need) // 2)
    _items_box(slide, tx, ty, tw, h - (ty - y), items, deck, tone, size)
    src, caption = image
    path = (ctx["md_dir"] / Path(src).expanduser()).resolve()
    if path.exists():
        S._add_picture(slide, path, (ix, y, iw, h), caption, deck, caption_color=tone.sub)
    else:
        ctx["warnings"].append(f"slide {ctx['n']}: image not found: {src}")
    _warn_fit(ctx["warnings"], ctx["n"], ctx["title"], "the text", need, h)


def gallery(ctx):
    slide, deck, tone = ctx["slide"], ctx["deck"], ctx["tone"]
    S = _s()
    blocks = ctx["blocks"]
    images = [p for k, p in blocks if k == "image"]
    intro = [b for b in blocks if b[0] in ("p", "quote")]
    x, y, w, h = _intro(slide, ctx["box"], intro, deck, tone, ctx["warnings"], ctx["n"], ctx["title"])
    count = len(images)
    if not 2 <= count <= 6:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): a gallery reads best with two to six pictures")
    cols = min(count, 5) or 1
    rows = math.ceil(count / cols) if count else 1
    gap = deck.gutter
    cw = (w - gap * (cols - 1)) // cols
    ch = (h - gap * (rows - 1)) // rows
    for i, (src, caption) in enumerate(images):
        cx, cy = x + (i % cols) * (cw + gap), y + (i // cols) * (ch + gap)
        path = (ctx["md_dir"] / Path(src).expanduser()).resolve()
        if path.exists():
            S._add_picture(slide, path, (cx, cy, cw, ch), caption, deck, caption_color=tone.sub)
        else:
            ctx["warnings"].append(f"slide {ctx['n']}: image not found: {src}")


def logos(ctx):
    slide, deck, tone = ctx["slide"], ctx["deck"], ctx["tone"]
    blocks = ctx["blocks"]
    images = [p for k, p in blocks if k == "image"]
    intro = [b for b in blocks if b[0] in ("p", "quote")]
    x, y, w, h = _intro(slide, ctx["box"], intro, deck, tone, ctx["warnings"], ctx["n"], ctx["title"])
    count = len(images)
    if not count:
        ctx["warnings"].append(f"slide {ctx['n']} ({ctx['title']}): logos need pictures")
        return
    cols = count if count <= 4 else (4 if count <= 8 else 6)
    rows = math.ceil(count / cols)
    gap = Inches(0.25)
    cw = (w - gap * (cols - 1)) // cols
    ch = min(Inches(1.4), (h - gap * (rows - 1)) // rows)
    top = y + max(0, (h - (ch * rows + gap * (rows - 1))) // 2)
    S = _s()
    for i, (src, caption) in enumerate(images):
        cx, cy = x + (i % cols) * (cw + gap), top + (i // cols) * (ch + gap)
        # logos keep a light tile on a dark slide: most are drawn for white paper
        _box(slide, cx, cy, cw, ch, fill=deck.paper, line=tone.card_border if not tone.dark else None,
             radius=tone.radius, line_w=0.75)
        path = (ctx["md_dir"] / Path(src).expanduser()).resolve()
        if not path.exists():
            ctx["warnings"].append(f"slide {ctx['n']}: image not found: {src}")
            continue
        pad_x, pad_y = int(cw * 0.18), int(ch * 0.22)
        pw, ph = S._picture_size(path)
        scale = min((cw - 2 * pad_x) / pw, (ch - 2 * pad_y) / ph)
        iw, ih = int(pw * scale), int(ph * scale)
        pic = slide.shapes.add_picture(str(path), cx + (cw - iw) // 2, cy + (ch - ih) // 2, iw, ih)
        if caption:
            pic._element.nvPicPr.cNvPr.set("descr", caption)


RENDERERS = {"statement": statement, "cards": cards, "compare": compare, "steps": steps,
             "numbered": numbered, "stats": stats, "media": media,
             "media-left": lambda ctx: media(ctx, image_left=True), "closing": None,
             "gallery": gallery, "logos": logos}


def render(name, spec, prs, deck, dark, md_dir, warnings, n, footer):
    """Add one slide in layout `name`; returns the slide."""
    S = _s()
    tone = tone_for(deck, dark)
    if name == "closing":
        return _closing(spec, prs, deck, dark, warnings, n, footer)
    layout = S._layout(prs, "Title Only Dark" if dark else "Title Only")
    slide = prs.slides.add_slide(layout)
    title = spec.get("title", "")
    keep = set()
    if title and name != "closing":
        S._set_title(slide, title, deck)
        keep.add(0)
    box = S._content_area(slide, prs, deck)
    S._drop_empty_placeholders(slide, keep)
    ctx = {"slide": slide, "prs": prs, "deck": deck, "tone": tone, "box": box, "blocks": spec["blocks"],
           "md_dir": md_dir, "warnings": warnings, "n": n, "title": title or "untitled"}
    RENDERERS[name](ctx)
    S._footer(slide, prs, footer, deck, inverted=dark)
    return slide

def _closing(spec, prs, deck, dark, warnings, n, footer):
    """The last slide mirrors the first: the section slide when dark, the title slide
    when light, with the heading as its title and every line under it as the subtitle."""
    S = _s()
    slide = prs.slides.add_slide(S._layout(prs, "Section Header" if dark else "Title Slide"))
    S._set_title(slide, spec.get("title", ""), deck)
    lines = [re.sub(r"[*`]", "", text) for _, text, _, _ in S._text_items(spec["blocks"])]
    sub = next((ph for ph in slide.placeholders if ph.placeholder_format.idx == 1), None)
    if sub is not None and lines:
        sub.text_frame.text = "\n".join(lines)
        S._drop_empty_placeholders(slide, {0, 1})
    else:
        S._drop_empty_placeholders(slide, {0})
    if len(lines) > 5:
        warnings.append(f"slide {n} ({spec.get('title')}): {len(lines)} closing lines; keep to five")
    if dark and deck.section_inverted and deck.footer_badge_invert:
        S._footer(slide, prs, footer, deck, inverted=True)
    return slide
