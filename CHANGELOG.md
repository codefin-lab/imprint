# Changelog

Versions follow Semantic Versioning. The skills and the engine share one version: a skill
asks for the engine release with the same number, and installs it when it is missing.

## 0.2.4 (2026-09-11)

- Slides: dark section slides can carry line art (`section_art`), recoloured by the engine
  (`section_art_color`, default white), faint (`section_art_alpha`), in a corner of its own
  (`section_art_corner`), so it need not repeat the title slide's.

## 0.2.3 (2026-09-11)

- Slides: `footer_bottom_in` sets how far a footer badge and its lines sit above the bottom
  edge (default 0.2 in, as before).

## 0.2.2 (2026-09-11)

- PDF conversion runs LibreOffice in Imprint's own profile (`~/.cache/imprint/libreoffice`).
  A LibreOffice window that is already open is left alone and no longer takes the conversion
  over, so a font installed while it was open is used at once and two conversions do not
  collide.
- Slides: the rule beside titles has no shadow; LibreOffice drew the shape style's effect.
- Slides: dark section slides can carry the footer inverted (`footer_badge_invert`).
- Slides: the footer's lines sit a little high beside the badge, with a small gap between them.

## 0.2.1 (2026-09-11)

- Slides: the logo can sit beside every title with a rule between them
  (`logo_position: title`, `title_rule`), with the title slide taking its own `cover_logo`.
- Slides: the footer can lead with a round badge (`footer_badge`), and its last line can be
  set in ink (`footer_last_line: ink`) for a label-then-name footer.
- Slides: inverted section slides can carry a light logo (`section_logo`).
- Slides: a list item that opens with `code` keeps a normal number or bullet; LibreOffice drew
  it in the code's grey monospace.
- Slides: code blocks take their size from the theme (`size.code`, default 14 pt).
- Slides: the title slide's line art stays clear of a two-line title.

## 0.2.0 (2026-09-11)

- The engine is an installable command, `imprint`: `imprint new`, `imprint docx`,
  `imprint pptx`, `imprint themes`, `imprint make-base-pptx`, `imprint doctor`.
- Install the skills with `npx skills add codefin-lab/imprint`, or as a Claude Code plugin.
- Each skill carries `scripts/ensure-engine.sh`, which installs or upgrades the engine to the
  version the skill was written for. `npx skills update` therefore upgrades both.
- Templates and the default theme ship inside the engine package.
- `imprint new --templates <folder>` starts a document from a brand's own templates.
- `build_docx.py` and `build_pptx.py` still work, as wrappers around the command.

## 0.1.0 (2026-09-11)

- First release: Markdown to .docx and .pptx, a brand-free default theme, guidance and
  templates for proposals, BRDs, technical and API specifications, minutes, reports and
  presentations.
