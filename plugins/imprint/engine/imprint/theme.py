"""Design tokens for a document theme, loaded from themes/<name>/theme.yaml."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

THEMES_DIR = Path(__file__).resolve().parent.parent / "themes"
DEFAULT_THEME = "default"


def _as_line_heights(value) -> dict:
    """`line_height: 1.45` is shorthand for `line_height: {body: 1.45}`."""
    if value is None:
        return {}
    if isinstance(value, dict):
        return value
    return {"body": value}


@dataclass(frozen=True)
class Theme:
    name: str
    base: Path
    font: dict
    size: dict
    heading_style: dict
    list: dict
    table: dict
    rule: dict
    page: dict = field(default_factory=dict)
    gantt: dict = field(default_factory=dict)
    image: dict = field(default_factory=dict)
    text: dict = field(default_factory=dict)
    space: dict = field(default_factory=dict)
    cover: dict = field(default_factory=dict)
    line_height: dict = field(default_factory=dict)
    toc: dict = field(default_factory=dict)
    labels: dict = field(default_factory=dict)
    defaults: dict = field(default_factory=dict)
    slides: dict = field(default_factory=dict)
    source: Path = field(repr=False, default=None)

    # --- convenience accessors used by the renderer -----------------------
    @property
    def body_font(self) -> str: return self.font["body"]

    @property
    def heading_font(self) -> str: return self.font["heading"]

    @property
    def mono_font(self) -> str: return self.font.get("mono", "Courier New")

    @property
    def body_pt(self) -> int: return self.size["body"]

    def heading_pt(self, level: int) -> int:
        return self.size.get(f"h{level}", self.size["body"])

    @property
    def literal_numbers(self) -> bool:
        """literal = the number written in the Markdown is printed as text.

        native (default) uses Word's own list numbering, so Word renumbers when
        someone edits the document; literal is what-you-typed-is-what-you-get and
        cannot drift, but nothing renumbers on edit.
        """
        return str(self.list.get("ordered_numbers", "native")).lower() == "literal"

    def line_twips(self, kind: str) -> int | None:
        """`w:line` for one kind of text, or None to leave the style alone.

        w:line with lineRule="auto" counts 240ths of a line, so 1.45 -> 348.
        """
        ratio = self.line_height.get(kind)
        return None if not ratio else int(round(float(ratio) * 240))

    def space_twips(self, key: str, default_pt: float = 0) -> int:
        """Vertical gap in twips for one of the theme's `space:` keys.

        base.docx sets `after: 0` on Normal, so without these every paragraph,
        list item and table butts straight into the next one — a blank line in
        the Markdown produces no visible gap at all.
        """
        return int(round(float(self.space.get(key, default_pt)) * 20))

    @property
    def hard_line_breaks(self) -> bool:
        """hard = a newline in the Markdown is a newline in the document."""
        return str(self.text.get("line_breaks", "hard")).lower() == "hard"

    @property
    def landscape_gantt(self) -> bool:
        return bool(self.gantt.get("landscape", False))

    @property
    def break_before_h1(self) -> bool:
        return bool(self.page.get("break_before_h1", False))

    def style_for(self, level: int) -> str:
        return self.heading_style.get(level, self.heading_style[max(self.heading_style)])

    @classmethod
    def load(cls, name_or_path: str | Path = DEFAULT_THEME) -> "Theme":
        path = Path(name_or_path)
        directory = path if path.is_dir() else THEMES_DIR / str(name_or_path)
        spec = directory / "theme.yaml"
        if not spec.exists():
            available = ", ".join(sorted(p.name for p in THEMES_DIR.iterdir() if p.is_dir()))
            raise FileNotFoundError(f"no theme at {spec} (available: {available})")

        data = yaml.safe_load(spec.read_text(encoding="utf-8")) or {}
        base = directory / data.get("base", "base.docx")
        if not base.exists():
            raise FileNotFoundError(f"theme '{data.get('name', directory.name)}' has no base document at {base}")

        # yaml gives int keys for 1/2/3 already, but be forgiving
        headings = {int(k): v for k, v in (data.get("heading_style") or {}).items()}
        return cls(
            name=data.get("name", directory.name),
            base=base,
            font=data.get("font", {}),
            size=data.get("size", {}),
            heading_style=headings,
            list=data.get("list", {}),
            table=data.get("table", {}),
            rule=data.get("rule", {}),
            page=data.get("page", {}) or {},
            gantt=data.get("gantt", {}) or {},
            image=data.get("image", {}) or {},
            text=data.get("text", {}) or {},
            space=data.get("space", {}) or {},
            cover=data.get("cover", {}) or {},
            line_height=_as_line_heights(data.get("line_height")),
            toc=data.get("toc", {}) or {},
            labels=data.get("labels", {}) or {},
            defaults={str(k): str(v) for k, v in (data.get("defaults") or {}).items()},
            slides=data.get("slides", {}) or {},
            source=spec,
        )
