---
name: slides
description: "Build a presentation deck (.pptx and PDF) from Markdown with the same syntax as docgen, on a theme's slide master, following presentation best practice (assertion titles, one idea per slide). Use when the user wants a presentation, pitch, deck or slides, or wants to turn a document into slides."
---

# slides: Markdown to .pptx

The same Markdown as docgen, laid out on slides. The output is a real PowerPoint file:
native layouts, editable text, a slide master from the theme. Same Markdown, same bytes.
Read `references/presentation.md` before writing a deck.

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
| `---` or `\pagebreak` | a new slide that keeps the current title |
| `<!-- notes: ... -->` | speaker notes |

Layout is chosen for you: text alone fills the slide; text and a visual share it, text on the
left; a wide visual (a wide picture, a table of four or more columns, a timeline, code) goes
under the text at full width; a visual alone takes the whole area.

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
| `slide_number`, `max_bullets`, `bullet_char` | slide numbers, the bullet warning limit, the bullet |

Line art is usually a crop of a larger drawing, so its lines stop at the edges it was cut on.
Place it so those edges sit on the slide's edges (a corner, with `section_art_flip` when the
corner differs from the crop), or the lines end in mid-air.
