"""Design tokens for a document theme, loaded from themes/<name>/theme.yaml."""
from __future__ import annotations

from dataclasses import dataclass, field, fields, replace
from pathlib import Path

import yaml

THEMES_DIR = Path(__file__).resolve().parent / "themes"


def theme_dir(name_or_path) -> Path:
    """A theme folder: a path to one, or the name of a theme that ships with the engine."""
    path = Path(name_or_path)
    return path if path.is_dir() else THEMES_DIR / str(name_or_path)
DEFAULT_THEME = "default"


def _as_line_heights(value) -> dict:
    """`line_height: 1.45` is shorthand for `line_height: {body: 1.45}`."""
    if value is None:
        return {}
    if isinstance(value, dict):
        return value
    return {"body": value}


def _merge(base: dict, over: dict) -> dict:
    """`over` laid on `base`, nested dicts merged rather than replaced."""
    out = dict(base)
    for k, v in over.items():
        out[k] = _merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


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
    callout: dict = field(default_factory=dict)
    chips: dict = field(default_factory=dict)
    screen: dict = field(default_factory=dict)
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

    # theme sections a document may adjust in its own front matter
    OVERRIDABLE = ("font", "size", "list", "table", "rule", "page", "gantt", "image", "text",
                   "space", "cover", "line_height", "toc", "labels", "callout", "chips", "screen",
                   "slides")

    def with_overrides(self, meta: dict) -> tuple["Theme", dict]:
        """(theme, rest of meta): a front-matter key named after a theme section
        and holding a mapping adjusts that section for this one document, e.g.

            chips: {radius: 0.2, padding_x: 5}
            table: {header_fill: DDEBF7}

        Scalars stay placeholders as before; only mappings are theme settings."""
        changes, rest = {}, {}
        names = {f.name for f in fields(self)}
        for k, v in meta.items():
            key = str(k)
            if key in self.OVERRIDABLE and key in names and isinstance(v, (dict, float, int)) \
                    and (isinstance(v, dict) or key == "line_height"):
                current = getattr(self, key)
                v = _as_line_heights(v) if key == "line_height" else v
                changes[key] = _merge(current or {}, v)
            else:
                rest[k] = v
        return (replace(self, **changes) if changes else self), rest

    @classmethod
    def load(cls, name_or_path: str | Path = DEFAULT_THEME) -> "Theme":
        path = Path(name_or_path)
        directory = theme_dir(name_or_path)
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
            callout=data.get("callout", {}) or {},
            chips=data.get("chips", {}) or {},
            screen=data.get("screen", {}) or {},
            source=spec,
        )
