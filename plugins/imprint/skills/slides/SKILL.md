---
name: slides
description: "Build a presentation deck (.pptx and PDF) from Markdown with the same syntax as docgen, on a theme's slide master, following presentation best practice (assertion titles, one idea per slide). Use when the user wants a presentation, pitch, deck or slides, or wants to turn a document into slides."
---

# slides: Markdown to .pptx

The same Markdown as docgen, laid out on slides. The output is a real PowerPoint file:
native layouts, editable text, a slide master from the theme. Same Markdown, same bytes.
Read `references/presentation.md` before writing a deck, and `references/widgets.md` before
putting a chart or infographic on a slide.

**Before the first build in a session**, make sure the engine is installed and new enough:

```bash
bash <this skill's folder>/scripts/ensure-engine.sh
```

```bash
imprint new presentation <project>/deck.md        # then edit it
imprint pptx <project>/deck.md --pdf
```

`--theme <name or folder>` picks the design; a brand skill names its own. `--strict` fails
on any warning. `imprint doctor` checks the installation.

## How Markdown becomes slides

| Markdown | Slide |
| :-- | :-- |
| front matter `TITLE`, `SUBTITLE`, `PRESENTER`, `DATE` | the title slide (`cover: false` leaves it out) |
| `<!-- toc -->` | an agenda slide listing the sections that follow |
| `# Heading` | a section divider; the first paragraph under it is its subtitle |
| `## Heading` | a content slide; write the heading as a full-sentence assertion |
| `### Heading` | a bold lead line inside the slide |
| paragraphs, bullets, numbered lists | the slide's text |
| table, `![caption](file.png)`, ` ```markwhen `, code block | a visual |
| ` ```widget ` with YAML | a chart or infographic (see Widgets below) |
| `---` or `\pagebreak` | a new slide that keeps the current title |
| `<!-- notes: ... -->` | speaker notes |

Layout is chosen for you: text alone fills the slide; text and a visual share it, text on the
left; a wide visual (a wide picture, a table of four or more columns, a timeline, code) goes
under the text at full width; a visual alone takes the whole area.

## Layouts

Most slides pick their layout from the shape of their content. Put `<!-- layout: name -->`
under the `##` heading to choose one yourself.

| Layout | Write | Picked automatically |
| :-- | :-- | :-- |
| `statement` | a `>` quote, then an optional source line | when the slide is only that |
| `stats` | two to four `- **42%** label` items | yes |
| `cards` | three to six `### Heading` blocks; `### 1 Heading` adds number tiles, `### Heading {icon=wallet}` an icon | yes |
| `compare` | exactly two `### Heading` blocks | yes |
| `steps` | a list of `**Step** description`, two to six items | name it |
| `numbered` | a list of `**Title** description`, then a `>` line for the callout | name it |
| `media`, `media-left` | text and one picture, the picture right or left | name it |
| `gallery` | two to six pictures; the alt text is the caption | when the slide is only pictures |
| `logos` | one picture per logo | name it |
| `closing` | the last heading and up to five lines under it | name it |

Any slide can sit on the theme's dark slide with `<!-- tone: dark -->`, and `tone: dark` in the
front matter makes dark the default. `imprint new showcase` starts a deck that uses every
layout, with the Markdown for each.

The build warns past each layout's limits: six cards, six steps, four numbers, six numbered
items, five closing lines. More than that is two slides.

## Widgets

Charts and infographics are written as a ` ```widget ` block of YAML with one standard set of
fields, and drawn from PowerPoint's own charts and shapes, so the reader can edit them.
`references/widgets.md` has the standard, a guide to choosing, and how good data slides look.

````markdown
## The pilot moved every number that matters

```widget
widget: kpi
items:
  - {label: Accounts opened, value: "12,400", delta: +18%, note: vs 2025}
  - {label: Days to open, value: 1, delta: -4, note: was 5}
```
````

| Kind | Widgets |
| :-- | :-- |
| Native charts, data editable in Excel | `column`, `bar`, `stacked-column`, `stacked-bar`, `line`, `area`, `pie`, `doughnut`, `waterfall` |
| Infographics from shapes | `kpi`, `progress`, `rings`, `funnel`, `timeline`, `cycle`, `hub`, `nested`, `waffle`, `matrix`, `features`, `devices` |

**Icons**: 1,838 Lucide icons (ISC licence) drawn as native shapes. Widget items take
`icon: name`; a `###` heading on a cards or compare slide takes `{icon=name}` at its end.
`imprint icons <word>` finds names; https://lucide.dev/icons shows them all.

**Phone frames**: `devices` draws an iPhone Pro from shapes, or uses a frame picture installed
in the theme or in `~/.imprint/device-frames` (`imprint devices` lists them). Apple's real
bezels may be installed there by each person, never shared; see `references/widgets.md`.

`imprint widgets` lists them, `imprint widgets <name>` prints an example, and
`imprint new widgets` starts a deck that uses every one. A mistake in the YAML (an unknown
widget or field, a value that is not a number) becomes a warning naming the slide, and the
widget is left out; build with `--strict` to stop instead. In a Word document a widget
becomes a table of its data.

## Diagrams on slides

Ask the diagram-design skill for the `slide-16x9` preset (1280 by 720, presentation type
ramp) and embed the PNG. A diagram drawn for a document page has type too small to project.

## Checks

The build warns when a slide has more bullets than the theme allows, when text or a table
probably overflows, and when a placeholder is left unfilled. It cannot see the rendered slide,
so **render the PDF and look at every slide**:

```bash
pdftoppm -png -r 50 deck.pdf /tmp/deck/s    # then read the images
```

Look for:

- a title that wraps, or breaks a word across lines; titles beside a logo have less room, so
  keep them to one line and shorten the words rather than the type
- text running out of its box, anything covering the logo or footer
- pictures that are blurred or too small, line art that stops in mid-slide
- a font that is not the theme's: check with `pdffonts deck.pdf`. The theme's fonts must be
  installed on the machine that builds and on the one that shows the deck

The PDF is made by LibreOffice in Imprint's own profile, so a LibreOffice window you have open
neither blocks it nor lends it the fonts it had when it started.

## Themes

A theme's `slides:` section in `theme.yaml` holds the whole slide design; the engine never
does. After changing anything there, regenerate the master and rebuild:

```bash
imprint make-base-pptx --theme <name or folder>
```

| Key | What it sets |
| :-- | :-- |
| `font.heading`, `font.body` | Latin and Thai typefaces (titles are set bold) |
| `size.title`, `size.body`, `size.levels`, `size.code`, `size.footer` | type sizes in points |
| `colors`, `accents` | ink, paper, greys, and the palette PowerPoint offers for charts |
| `logo`, `logo_position` | `top-right`, `bottom-left`, or `title`: beside every content title |
| `title_rule` | with `logo_position: title`, a vertical rule between logo and title |
| `cover_logo`, `cover_art` | the title slide's own logo and line art (the art keeps clear of the title) |
| `section_inverted` | dark section slides, with light text |
| `section_logo` | a light logo on the dark section slides |
| `section_art` | line art on the dark section slides, recoloured by the engine |
| `section_art_color`, `section_art_alpha` | its line colour and strength (0 to 1) |
| `section_art_corner`, `section_art_flip` | its corner, and a flip so the edges it was cropped on meet the slide's edges |
| `footer` | footer text; `\n` starts a second line |
| `footer_align`, `footer_last_line` | left or right; `ink` sets the last line darker than the rest |
| `footer_badge`, `footer_badge_invert` | a round badge before the footer, and its version for dark slides |
| `footer_bottom_in`, `footer_h_in` | how far the footer sits above the edge, and the space kept clear for it |
| `size.statement`, `size.stat`, `size.card_title`, `size.card_body`, `size.step` | type sizes inside the layouts |
| `components` | card fill, border and corner radius; `accent` for numbers and step boxes; `accents` to colour cards, steps and stats in order |
| `dark`, `logo_dark` | the colours of `tone: dark` slides, and the light logo they carry |
| `icons` | a folder of the theme's own SVG icons, added to the bundled set |
| `device_frames` | a folder of frame pictures (PNG, transparent screen) the theme may share |
| `widgets`, `size.widget_value`, `size.widget_label` | widget colours (palette, highlight, muted, track, grid, positive, negative, and a `dark` set) and type |
| `slide_number`, `max_bullets`, `bullet_char` | slide numbers, the bullet warning limit, the bullet |

Line art is usually a crop of a larger drawing, so its lines stop at the edges it was cut on.
Place it so those edges sit on the slide's edges (a corner, with `section_art_flip` when the
corner differs from the crop), or the lines end in mid-air.
