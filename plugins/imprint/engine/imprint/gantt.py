"""Render a ```markwhen (or ```gantt) block into a PNG, styled from the theme.

The timeline is written in **markwhen** — https://markwhen.com — which is the
plain-text timeline language the team already knows:

    title: Offshore Equity Trading

    group Design
    2026-01/2026-02: Gathering requirements #design
    2026-01-15/3 weeks: UX/UI design #design
    endGroup

    group Build
    2026-03/2026-06: Implementation — mobile app #build
    endGroup

    2026-07-01: Go-live #deploy

We speak markwhen's syntax but draw the chart ourselves.  The official
`@markwhen/parser` is a Node package (a Python pipeline would have to shell out
to it), it silently drops `group`/`endGroup` in the version on npm, and its
renderer is a Vue web app — turning that into a PNG for a .docx would mean
driving a headless browser, which is neither cheap nor deterministic.  The
subset below covers what proposal timelines actually use and renders through the
same matplotlib path as before, so builds stay byte-identical run to run.

Date grammar (dates only; times are ignored):

    2026            2026-01           2026-01-15         a point or a bound
    A / B           both bounds, end rounded up to the end of its granularity
    A / 3 months    a bound plus a duration
    3 months: ...   a duration alone starts where the previous event ended

A lone day-granularity date (`2026-07-01: Go-live`) is a milestone; add
`#milestone` to any event to force one.  A `#tag` picks the bar colour from
`theme.yaml → gantt.colors`.

The older `name | start | duration | kind` spec still parses — a block is read
as markwhen unless it contains one of those pipe rows.

matplotlib is imported lazily so the rest of docgen keeps working without it.
"""
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

from .theme import Theme

# --- legacy ``name | start | duration | kind`` spec ------------------------
AXIS = re.compile(r"^axis\s*:\s*(?P<fmt>\S+)\s+(?P<lo>-?[\d.]+)\s*\.\.\s*(?P<hi>-?[\d.]+)\s*$", re.I)
MILESTONE = re.compile(r"^milestone\s*:\s*(?P<at>-?[\d.]+)\s+(?P<label>.+)$", re.I)
TITLE = re.compile(r"^title\s*:\s*(?P<text>.+)$", re.I)
LEGACY_ROW = re.compile(r"^[^|:]+\|\s*-?[\d.]+\s*\|\s*-?[\d.]+")

# --- markwhen -------------------------------------------------------------
GROUP_OPEN = re.compile(r"^(?:group|section)\b\s*(?P<label>.*)$", re.I)
GROUP_CLOSE = re.compile(r"^(?:endGroup|endSection)\s*$", re.I)
HEADER_KEY = re.compile(r"^(?P<key>[A-Za-z_][\w-]*)\s*:\s*(?P<value>.*)$")
EVENT = re.compile(r"^(?P<when>[^:]+?)\s*:\s*(?P<text>.+)$")
TAG = re.compile(r"#([\w-]+)")
DURATION = re.compile(r"^(?P<n>\d+(?:\.\d+)?)\s*(?P<unit>day|days|week|weeks|month|months|year|years)$", re.I)
YMD = re.compile(r"^(?P<y>\d{4})(?:-(?P<m>\d{1,2})(?:-(?P<d>\d{1,2}))?)?$")

DAY = 1.0
UNIT_DAYS = {"day": 1, "week": 7, "month": 30.4375, "year": 365.25}


@dataclass
class Task:
    name: str
    start: float
    duration: float
    kind: str = "default"
    depth: int = 0
    label: str | None = None      # duration as the author wrote it, if they did


@dataclass
class Group:
    label: str
    depth: int = 0


@dataclass
class GanttSpec:
    """Rows in draw order, plus everything the renderer needs to place them.

    `rows` interleaves Group headers and Tasks so a chart keeps the shape the
    author wrote.  Positions are always numbers: axis units in legacy mode,
    `date.toordinal()` in markwhen mode.
    """
    rows: list = field(default_factory=list)
    milestones: list[tuple[float, str]] = field(default_factory=list)
    axis_format: str = "T+{i}"
    axis_lo: float | None = None
    axis_hi: float | None = None
    title: str | None = None
    date_mode: bool = False
    ticks: list[tuple[float, str]] = field(default_factory=list)

    @property
    def tasks(self) -> list[Task]:
        return [r for r in self.rows if isinstance(r, Task)]

    @property
    def lo(self) -> float:
        if self.axis_lo is not None:
            return self.axis_lo
        return min([t.start for t in self.tasks] + [m[0] for m in self.milestones] + ([] if self.date_mode else [0]))

    @property
    def hi(self) -> float:
        if self.axis_hi is not None:
            return self.axis_hi
        ends = [t.start + t.duration for t in self.tasks] + [m[0] for m in self.milestones]
        return max(ends + ([] if self.date_mode else [1]))


# ------------------------------------------------------------------ parsing

def parse_spec(text: str, theme: Theme | None = None) -> GanttSpec:
    """markwhen unless the block still uses the old pipe rows."""
    if any(LEGACY_ROW.match(l.strip()) for l in text.splitlines()):
        return _parse_legacy(text)
    unit = (theme.gantt.get("axis_unit", "auto") if theme else "auto")
    return parse_markwhen(text, axis_unit=str(unit))


def _parse_legacy(text: str) -> GanttSpec:
    spec = GanttSpec()
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or line.startswith("//"):
            continue
        if m := AXIS.match(line):
            spec.axis_format = m.group("fmt")
            spec.axis_lo, spec.axis_hi = float(m.group("lo")), float(m.group("hi"))
        elif m := MILESTONE.match(line):
            spec.milestones.append((float(m.group("at")), m.group("label").strip()))
        elif m := TITLE.match(line):
            spec.title = m.group("text").strip()
        elif "|" in line:
            parts = [c.strip() for c in line.split("|")]
            if len(parts) < 3:
                continue
            kind = parts[3] if len(parts) > 3 and parts[3] else "default"
            try:
                spec.rows.append(Task(parts[0], float(parts[1]), float(parts[2]), kind))
            except ValueError:
                continue
    return spec


def _bound(token: str, *, end: bool) -> date | None:
    """A markwhen date bound.  A start rounds down to its granularity, an end
    rounds up past it: `2026-01` as an end means the last day of January."""
    m = YMD.match(token.strip())
    if not m:
        return None
    y = int(m.group("y"))
    mo, d = m.group("m"), m.group("d")
    if not end:
        return date(y, int(mo or 1), int(d or 1))
    if d:
        return date(y, int(mo), int(d)) + timedelta(days=1)
    if mo:
        mo = int(mo)
        return date(y + (mo == 12), 1 if mo == 12 else mo + 1, 1)
    return date(y + 1, 1, 1)


def _granularity(token: str) -> str | None:
    m = YMD.match(token.strip())
    if not m:
        return None
    return "day" if m.group("d") else ("month" if m.group("m") else "year")


def _shift(start: date, n: float, unit: str) -> date:
    """Add a duration.  Whole months and years land on the same day-of-month
    rather than drifting by 30.44 days, which is what a plan means by them."""
    unit = unit.rstrip("s").lower()
    if unit in ("month", "year") and float(n).is_integer():
        months = int(n) * (12 if unit == "year" else 1)
        total = start.month - 1 + months
        y, mo = start.year + total // 12, total % 12 + 1
        return date(y, mo, min(start.day, calendar.monthrange(y, mo)[1]))
    return start + timedelta(days=round(float(n) * UNIT_DAYS[unit]))


def _parse_when(when: str, previous_end: date | None):
    """(start, end, is_point, spelled) for a markwhen date expression, or None.

    `spelled` is the duration the way the author wrote it (`2 months` -> "2 mo").
    A written duration beats a derived one: Jan 1 + 2 months is 59 days, and
    printing "1.9 mo" next to a bar the author called two months is just wrong.
    """
    parts = [p.strip() for p in when.split("/")]
    if len(parts) == 1:
        token = parts[0]
        if m := DURATION.match(token):                 # `3 months:` — relative
            start = previous_end or date.today()
            return start, _shift(start, float(m.group("n")), m.group("unit")), False, _spelled(m)
        start = _bound(token, end=False)
        if start is None:
            return None
        return start, _bound(token, end=True), _granularity(token) == "day", None
    start = _bound(parts[0], end=False)
    if start is None:
        return None
    tail = parts[1]
    if m := DURATION.match(tail):
        return start, _shift(start, float(m.group("n")), m.group("unit")), False, _spelled(m)
    end = _bound(tail, end=True)
    if end is None:
        return None
    return start, end, False, None


SHORT_UNIT = {"day": "d", "week": "wk", "month": "mo", "year": "y"}


def _spelled(m: re.Match) -> str:
    n = float(m.group("n"))
    return f"{n:g} {SHORT_UNIT[m.group('unit').rstrip('s').lower()]}"


def parse_markwhen(text: str, *, axis_unit: str = "auto") -> GanttSpec:
    spec = GanttSpec(date_mode=True)
    spec.axis_format = axis_unit
    depth = 0
    previous_end: date | None = None
    seen_event = False

    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("//"):
            continue
        if GROUP_CLOSE.match(line):
            depth = max(0, depth - 1)
            continue
        if m := GROUP_OPEN.match(line):
            label = TAG.sub("", m.group("label")).strip()
            if label:
                spec.rows.append(Group(label, depth))
            depth += 1
            continue

        parsed = None
        if m := EVENT.match(line):
            parsed = _parse_when(m.group("when"), previous_end)

        if parsed is None:
            # header keys are only header keys before the first event
            if not seen_event and (h := HEADER_KEY.match(line)):
                key, value = h.group("key").lower(), h.group("value").strip()
                if key == "title":
                    spec.title = value
                elif key == "axis":
                    spec.axis_format = value          # month | quarter | year | auto
            continue

        seen_event = True
        start, end, is_point, spelled = parsed
        previous_end = end
        text_part = m.group("text")
        tags = [t.lower() for t in TAG.findall(text_part)]
        name = TAG.sub("", text_part).strip()
        x0, x1 = float(start.toordinal()), float(end.toordinal())

        if is_point or "milestone" in tags:
            spec.milestones.append((x0, name))
        else:
            kind = next((t for t in tags if t != "milestone"), "default")
            spec.rows.append(Task(name, x0, x1 - x0, kind, depth, spelled))

    if spec.date_mode and (spec.tasks or spec.milestones):
        spec.axis_lo, spec.axis_hi, spec.ticks = _date_axis(spec)
    return spec


def _date_axis(spec: GanttSpec):
    """Pad the range out to whole periods and lay ticks on their boundaries."""
    lo = date.fromordinal(int(min([t.start for t in spec.tasks] + [m[0] for m in spec.milestones])))
    hi = date.fromordinal(int(max([t.start + t.duration for t in spec.tasks] + [m[0] for m in spec.milestones])))
    span = (hi - lo).days
    unit = spec.axis_format.lower() if spec.axis_format in ("month", "quarter", "year", "week") else None
    if unit is None:
        unit = "week" if span <= 70 else "month" if span <= 500 else "quarter" if span <= 1500 else "year"

    ticks: list[tuple[float, str]] = []
    if unit == "week":
        cur = lo - timedelta(days=lo.weekday())
        while cur <= hi + timedelta(days=7):
            ticks.append((float(cur.toordinal()), cur.strftime("%-d %b")))
            cur += timedelta(days=7)
    elif unit in ("month", "quarter"):
        step = 3 if unit == "quarter" else 1
        cur = date(lo.year, lo.month - (lo.month - 1) % step, 1)
        while cur <= _shift(date(hi.year, hi.month, 1), step, "month"):
            label = f"Q{(cur.month - 1) // 3 + 1} {cur:%Y}" if unit == "quarter" else cur.strftime("%b %y")
            ticks.append((float(cur.toordinal()), label))
            cur = _shift(cur, step, "month")
    else:
        for y in range(lo.year, hi.year + 2):
            ticks.append((float(date(y, 1, 1).toordinal()), str(y)))

    axis_lo = min(ticks[0][0], float(lo.toordinal()))
    axis_hi = max(ticks[-1][0], float(hi.toordinal()))
    return axis_lo, axis_hi, ticks


def _duration_label(days: float, date_mode: bool) -> str:
    if not date_mode:
        return f"{days:g} mo"
    if days >= 350:
        return f"{days / 365.25:.1f}".rstrip("0").rstrip(".") + " y"
    if days >= 25:
        return f"{days / 30.4375:.1f}".rstrip("0").rstrip(".") + " mo"
    if days >= 7 and days % 7 == 0:
        return f"{days / 7:g} wk"
    return f"{days:g} d"


# ----------------------------------------------------------------- drawing

def render_png(spec: GanttSpec, theme: Theme, out_path: Path, *, landscape: bool = False) -> Path:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    from matplotlib.patches import FancyBboxPatch

    g = theme.gantt
    colors = {k: "#" + v.lstrip("#") for k, v in (g.get("colors") or {}).items()}
    default = colors.get("default", "#2D3748")
    label_color = "#" + str(g.get("label_color", "4A5568")).lstrip("#")
    milestone_color = "#" + str(g.get("milestone_color", "1A202C")).lstrip("#")
    font_pt = g.get("font_pt", 9)

    installed = {f.name for f in font_manager.fontManager.ttflist}
    family = theme.body_font if theme.body_font in installed else "DejaVu Sans"
    plt.rcParams.update({"font.family": family, "font.size": font_pt})

    rows = spec.rows
    n = max(len(rows), 1)
    # Render at the proportions of the page it is going on.  A chart drawn for a
    # portrait column is too tall for a landscape page once the heading and the
    # notes underneath it have taken their share of the height.
    fig_w = g.get("landscape_width_in", 10.4) if landscape else g.get("width_in", 9.6)
    per_task = (g.get("landscape_height_per_task_in", 0.32) if landscape
                else g.get("height_per_task_in", 0.42))
    height = per_task * n + (0.9 if spec.milestones else 0.5)
    fig, ax = plt.subplots(figsize=(fig_w, height), dpi=g.get("dpi", 300))

    lo, hi = spec.lo, spec.hi
    span = max(hi - lo, 1.0)
    bar_h = g.get("bar_height", 0.56)

    bars = []   # drawn after the layout is final — the corner depends on the axes' real size
    for i, row in enumerate(rows):
        y = n - i - 1
        if isinstance(row, Group):
            continue
        bars.append((row, y))
        if g.get("show_duration", True):
            ax.text(row.start + row.duration + (span * 0.012 if spec.date_mode else 0.12), y,
                    row.label or _duration_label(row.duration, spec.date_mode),
                    va="center", ha="left", fontsize=font_pt - 1.5, color=label_color)

    base = -0.95 if spec.milestones else -0.5
    for at, label in spec.milestones:
        ax.plot(at, base, marker="D", markersize=g.get("milestone_size", 7),
                color=milestone_color, clip_on=False, zorder=5)
        ax.text(at, base - 0.5, label, ha="center", va="top",
                fontsize=font_pt - 1, fontweight="bold", color=milestone_color)

    ax.set_yticks(range(n))
    ax.set_yticklabels([r.label if isinstance(r, Group) else ("    " * r.depth + r.name)
                        for r in reversed(rows)], fontsize=font_pt - 0.5)
    for tick, row in zip(ax.get_yticklabels(), reversed(rows)):
        if isinstance(row, Group):
            tick.set_fontweight("bold")

    if spec.date_mode:
        ax.set_xticks([t for t, _ in spec.ticks])
        ax.set_xticklabels([lbl for _, lbl in spec.ticks], fontsize=font_pt - 0.5)
        pad = span * 0.085
    else:
        ticks = list(range(int(lo), int(hi) + 1))
        ax.set_xticks(ticks)
        ax.set_xticklabels([spec.axis_format.format(i=i) for i in ticks], fontsize=font_pt - 0.5)
        pad = 0.6
    ax.set_xlim(lo - (span * 0.008 if spec.date_mode else 0.05), hi + pad)
    ax.set_ylim(base - (0.6 if spec.milestones else 0.1), n - 0.35)
    ax.xaxis.set_ticks_position("top")
    ax.xaxis.set_label_position("top")
    for side in ("top", "right", "left", "bottom"):
        ax.spines[side].set_visible(False)
    ax.tick_params(length=0)
    ax.grid(axis="x", color="#" + str(g.get("grid_color", "E2E8F0")).lstrip("#"), linewidth=0.8)
    ax.set_axisbelow(True)
    if spec.title:
        # the axis ticks sit on TOP of the chart, so the title needs more
        # clearance than matplotlib's default or it crowds the dates
        ax.set_title(spec.title, loc="left", fontsize=font_pt + 1,
                     fontweight="bold", pad=g.get("title_pad_pt", 24))

    fig.tight_layout()

    # rounding_size is in DATA units and applies to both axes, but the x axis
    # (months or days) and the y axis (rows) have very different scales, so one
    # radius comes out as an ellipse — stretched, lopsided bar ends.
    # mutation_aspect squeezes y before the corner is cut and stretches it back
    # after; set to the axes' real x/y inches-per-unit ratio, measured after
    # tight_layout has sized the axes around the task labels, the corner is a
    # true circle.  The radius is a fraction of the bar height (gantt.corner_ratio).
    pos = ax.get_position()
    (x0, x1), (y0, y1) = ax.get_xlim(), ax.get_ylim()
    in_per_x = pos.width * fig.get_figwidth() / (x1 - x0)
    in_per_y = pos.height * fig.get_figheight() / (y1 - y0)
    aspect = in_per_x / in_per_y
    ratio = min(float(g.get("corner_ratio", 0.25)), 0.5)
    for row, y in bars:
        corner = min(ratio * bar_h / aspect, row.duration / 2)
        ax.add_patch(FancyBboxPatch(
            (row.start, y - bar_h / 2), row.duration, bar_h,
            boxstyle=f"round,pad=0,rounding_size={corner}",
            mutation_aspect=aspect,
            linewidth=0, facecolor=colors.get(row.kind, default)))

    fig.savefig(out_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return out_path
