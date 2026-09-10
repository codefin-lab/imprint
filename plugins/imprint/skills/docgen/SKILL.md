---
name: docgen
description: "Write and build client-ready documents (business proposal, BRD, technical specification, API specification, minutes of meeting, technical report) as .docx and PDF from Markdown, following general international standards for each document type, in a theme's house design. Use when the user wants to write, build, rebuild or restyle a document, turn notes or minutes into a deliverable, put a diagram into a document, or mentions docgen or Imprint."
---

# docgen: Markdown to .docx

One deterministic pipeline for every document: same Markdown, same bytes, every run. The
engine is the `imprint` command.

**Before the first build in a session**, run this skill's installer; it installs or upgrades
the engine to the version this skill was written for, and does nothing if it is already there:

```bash
bash <this skill's folder>/scripts/ensure-engine.sh
```

If a build misbehaves, `imprint doctor` checks the engine, LibreOffice, poppler and fonts, and
says whether a newer release exists.

## Pick the document type

Read the reference for the type before writing; it gives the structure and the rules.

| Document | Start it with | Reference |
| :-- | :-- | :-- |
| Business proposal | `imprint new proposal <path>` | `references/proposal.md` |
| Business requirements document | `imprint new brd <path>` | `references/brd.md` |
| Technical specification | `imprint new tech-spec <path>` | `references/tech-spec.md` |
| API specification | `imprint new api-spec <path>` | `references/api-spec.md` |
| Minutes of meeting | `imprint new mom <path>` | `references/mom-and-report.md` |
| Technical report | `imprint new report <path>` | `references/mom-and-report.md` |
| Presentation | use the `slides` skill | |

`references/writing.md` applies to all of them: plain language, words that bind, dates and
units, tables, figures, document control.

If a brand skill is installed (for example a company's own Imprint skill), it names the theme
and templates to use instead; follow it.

## Build

```bash
imprint new brd <project>/docs/BRD-Example.md          # then edit it
imprint docx <project>/docs/BRD-Example.md --pdf
```

**A document lives with its project**, next to the `.docx` and `.pdf` it builds.
`--theme <name or folder>` picks a design (default: `default`); `--strict` fails on any
unfilled `{{placeholder}}`; `-o` sets the output path.

## Front matter decides the document's shape

```yaml
document_kind: Technical Specification   # the kind on the cover and in the running header
cover: false                  # no cover page (minutes, reports, memos)
header_left: Minutes of Meeting   # retitle only the running header
page_break_before_h1: false   # flow instead of one section per page
line_breaks: soft             # standard Markdown wrapping instead of literal newlines
ANY_KEY: value                # fills {{ANY_KEY}} in body, cover, header and footer
```

A theme's `defaults:` fill the keys a document leaves out, such as the company name and
address.

## Writing the Markdown

| Syntax | Result |
| :-- | :-- |
| `# `, `## `, `### ` | section, sub-head, minor head |
| `- ` or `* ` | bullet; indent 4 spaces to nest |
| `1. ` | numbered; write the number you want, it survives an interrupting table |
| pipe table | shaded repeating header, borders, columns sized by content |
| `:--`, `--:`, `:-:` in the separator row | column aligned left, right or centre, header included |
| `\| **Group** \|  \|  \|` | full-width shaded group row (label must be bold) |
| fenced code block (any language) | monospace block, for API examples and commands |
| ` ```markwhen ` | timeline or Gantt chart in markwhen syntax, real dates |
| `![caption](file.png)` | a picture on its own line, fitted, caption beneath |
| `<!-- landscape -->` up to `<!-- portrait -->` | a section on landscape pages |
| `<!-- toc -->` | table of contents with real page numbers |
| `\pagebreak` | page break |

A newline in the source is a newline on the page; a blank line starts a new paragraph.

**Align every table column by what it holds**: text left, quantities and money right with
the same decimals and the unit in the header, codes and IDs left, short equal values such as
Yes/No centred.

**Plain text only**: no section signs, emoji or decorative symbols. Write "Section 3", "to",
"Yes".

## Diagrams go in as pictures

Draw diagrams with the **diagram-design** skill and embed the exported PNG:

```markdown
## Architecture

![Figure 1: Solution architecture](diagrams/architecture.png)
```

- The heading and caption belong to the document; export the diagram without its own title.
- PNG or JPEG only; Word cannot embed SVG.
- Draw at document size or it prints too small: a portrait page is 6.9 in wide, so ask
  diagram-design for `doc-inline` (960 wide); anything wider or dense goes on a landscape page
  (`print-a4-landscape`).
- Check a diagram with diagram-design's own checks, and look at it, before embedding.

## Themes

A theme is a folder: `theme.yaml` (fonts, sizes, spacing, tables, cover, labels, defaults,
slides) plus `base.docx` (cover, styles, header and footer) and `base.pptx` (slide master).
Design changes go in the theme, never in the engine. To make a brand, copy the default theme
(`imprint themes` prints its folder), change it, and pass `--theme <folder>`. `labels:` must
match the literal text on the new `base.docx`'s cover.

## Before saying it is done

`verify` runs on every build and fails on structural problems. That is not enough: **render
the PDF and look at every page.**

```bash
pdftoppm -png -r 80 out.pdf /tmp/review/p    # then read the images
```

Check: no near-empty pages, no table row cut at a page break, no heading stranded at the foot
of a page, every figure readable, every placeholder filled.

## Traps

- The table of contents needs LibreOffice: the build lays the document out, reads each
  heading's page back from the PDF, and rebuilds. If LibreOffice is busy it can fail for one
  run; the build warns, and a rebuild fixes it.
- A heading right after `<!-- landscape -->` must not also get `\pagebreak`; the section break
  already starts a page. Put the landscape marker before the heading.
- Fonts must be installed on the machine that opens the file, or Word substitutes.
