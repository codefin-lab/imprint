# Changelog

Versions follow Semantic Versioning. The skills and the engine share one version: a skill
asks for the engine release with the same number, and installs it when it is missing.

## Unreleased

- Tables: `Table: Title` on the line above a table puts a dark title band across its top,
  repeated with the header on a new page.
- Tables: `{layout=form}` on that line makes a form: no header row, and every all-bold cell is a
  shaded label, with label columns sized to their labels.
- Tables: a cell that is only `>` merges into the cell on its left, only `^` into the cell above.
- Tables: `{layout=sections}` draws headed boxes (a band with the heading and italic guidance over
  a box to write in), the shape of a project-details or design form.
- Tables: `tone=gray|light|amber|blue` colours bands, labels and boxes (a theme can add tones under
  `table.tones`); `height=` sets the least height of rows to write in. In a form, `| **Group** | > |`
  is a group row and `| text | > |` a full-width description row.
- Callouts: `> [!NOTE]`, `[!TIP]`, `[!IMPORTANT]`, `[!WARNING]` and `[!CAUTION]` draw a boxed note
  with a coloured bar; text after the marker replaces the label. The theme's `callout:` sets
  colours, labels and padding. On a slide a callout is a paragraph led by its label.
- Inline: `[[Active]]` is a chip coloured by its label (theme `chips.labels`), `[[Hold|amber]]`
  names its colour; `!!text!!` is alert red; `[x]` and `[ ]` are ballot boxes.
- Tables: `![](shot.png)` in a cell, alone or between `<br>`s, is a picture fitted to the column.
  `{header=2}` on the `Table:` line makes the first two rows the header, merges included.
- Screens: `::: screen shot.png` ... `:::` puts the screenshot on a dark panel beside what is
  written inside (paragraphs, tables, lists, callouts); `::: screens {per_row=3}` lays
  `![caption](file.png)` lines side by side under their captions. Slides show the pictures and
  the text in turn.
- The BRD template uses them: a project form, titled tables, a priority scale and references.

## 0.6.0 (2026-09-14)

- Devices: the drawn phone is now an iPhone Pro to its real proportions, with a metal band,
  black border, Dynamic Island and side buttons, in `iphone`, `iphone-black`, `iphone-white` and
  `iphone-desert` (the new `frame` field).
- Devices: frame pictures. Any PNG with a transparent screen works; Imprint finds the screen and
  places the screenshot under the frame. Frames are looked up in the theme's
  `slides.device_frames`, then in `~/.imprint/device-frames` (or `IMPRINT_DEVICE_FRAMES`), for
  frames such as Apple's product bezels that each person installs and may not share.
- `imprint devices` lists the frames available.

## 0.5.0 (2026-09-14)

- Icons: Lucide 1.46.0 (1,838 icons, ISC licence) bundled with the engine and drawn as native
  PowerPoint freeform shapes: sharp at any size, recolourable, the same in LibreOffice.
  `imprint icons <word>` searches names and meanings. A theme can add SVG icons with
  `slides.icons`.
- Widget items take `icon` (kpi, progress, timeline, cycle, hub, features) and `image` (devices).
- New widgets: `features` (icon, title and a line per item, in a grid) and `devices`
  (screenshots in phone frames drawn from shapes).
- Layouts: `### Heading {icon=name}` puts an icon on a card or a compare header. Documents drop
  the attribute.
- `imprint new` also copies pictures a template names in a widget's `image:` field.
- `tools/update_icons.py` refreshes the icon set from a Lucide release.

## 0.4.0 (2026-09-14)

- Widgets: charts and infographics written as a ` ```widget ` YAML block with one standard set
  of fields (`widget`, `title`, `note`, `items` or `categories` and `series`, `prefix`, `suffix`,
  `decimals`, `max`, `highlight`, `colors`, `labels`, `legend`, `columns`, `center`, `x`, `y`,
  `quadrants`). Unknown widgets and fields are reported by name.
- Native PowerPoint charts, data editable in Excel: column, bar, stacked-column, stacked-bar,
  line, area, pie, doughnut and waterfall. Their embedded workbooks are pinned, so builds stay
  byte-identical.
- Infographics from shapes: kpi (with change), progress, rings, funnel, timeline (done, now,
  next), cycle, hub, nested (market size), waffle and matrix.
- A widget sits wherever a picture does, beside text or at full width; in a document it becomes a
  table of its data.
- `imprint widgets` lists them, prints an example, or lists the fields. Template `widgets`, a deck
  using every widget. Themes: `slides.widgets` colours, `size.widget_value`, `size.widget_label`.
- The slides skill: `references/widgets.md` (the standard, choosing a widget, how good data slides
  look), deck shapes by purpose and choosing a layout from the content.
- The engine now needs XlsxWriter (pinned), which python-pptx uses for chart data.

## 0.3.0 (2026-09-14)

- Slides: ten layouts beyond text and a picture: statement, stats, cards, compare, steps,
  numbered with a callout, media (picture right or left), gallery, logos and closing. Content
  with an unmistakable shape picks its own; `<!-- layout: name -->` chooses one.
- Slides: `<!-- tone: dark -->` puts any slide on a dark twin of Title Only ("Title Only Dark"
  in the master, with `logo_dark`); `tone: dark` in the front matter makes it the default.
- Themes: `components` (card fill, border, radius, accent, `accents` in order), `dark`
  colours, and sizes for the layouts. Regenerate a theme's master with `imprint make-base-pptx`.
- Markdown: `>` blockquotes. Documents set them as an indented italic paragraph.
- Templates: `showcase`, a deck that uses every layout.

## 0.2.6 (2026-09-11)

- The slides skill documents every `slides:` theme key in one table, and its checks now cover
  titles that wrap beside a logo, line art that stops in mid-slide, and fonts (`pdffonts`).
  No engine change.

## 0.2.5 (2026-09-11)

- Slides: `section_art_flip` (vertical, horizontal or both) turns the section slides' line art
  so the edges it was cropped on land on the slide's edges; its lines run off the slide
  instead of stopping in mid-air.

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
