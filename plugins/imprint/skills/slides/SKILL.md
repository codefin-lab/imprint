---
name: slides
description: Build a presentation deck (.pptx and PDF) from Markdown with the same syntax as docgen, on a theme's slide master, following presentation best practice (assertion titles, one idea per slide). Use when the user wants a presentation, pitch, deck or slides, or wants to turn a document into slides.
---

# slides: Markdown to .pptx

The same Markdown as docgen, laid out on slides. The output is a real PowerPoint file:
native layouts, editable text, a slide master from the theme. Same Markdown, same bytes.
Read `references/presentation.md` before writing a deck.

```bash
cp "${CLAUDE_PLUGIN_ROOT}/templates/presentation.md" <project>/deck.md    # then edit
python3 "${CLAUDE_PLUGIN_ROOT}/engine/build_pptx.py" <project>/deck.md -o <project>/deck.pptx --pdf
```

`--theme <name or folder>` picks the design; a brand plugin names its own. `--strict` fails
on any warning.

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

Ask the diagram-design plugin for the `slide-16x9` preset (1280 by 720, presentation type
ramp) and embed the PNG. A diagram drawn for a document page has type too small to project.

## Checks

The build warns when a slide has more bullets than the theme allows, when text or a table
probably overflows, and when a placeholder is left unfilled. It cannot see the rendered slide,
so **render the PDF and look at every slide**:

```bash
pdftoppm -png -r 50 deck.pdf /tmp/deck/s    # then read the images
```

Look for text running out of its box, anything covering the logo or footer, and pictures that
are blurred or too small.

## Themes

A theme's `slides:` section in `theme.yaml` sets fonts (Latin and Thai), sizes, colours, the
accent palette, the logo and its position, the footer, and whether section slides are inverted.
After changing it, regenerate the master:

```bash
python3 "${CLAUDE_PLUGIN_ROOT}/engine/tools/make_base_pptx.py" --theme <name or folder>
```
