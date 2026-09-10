# Imprint

Documents, diagrams and slides from Markdown, in your own house style.

Imprint is a set of Claude Code plugins. You write, or ask Claude to write, a Markdown file;
Imprint builds a Word document or a PowerPoint deck from it in your company's design, the same
bytes every time. Each document type comes with guidance based on general international
standards, so a BRD reads like a BRD and an API specification like one.

| Plugin | What it does |
| :-- | :-- |
| `imprint` | Markdown to .docx and .pptx; templates and guidance per document type |
| `diagram-design` | architecture, flow, sequence, state, ER, AWS and other diagrams, checked for geometry |

## Document types

| Type | Built on |
| :-- | :-- |
| Business proposal | common professional practice, plain language (ISO 24495-1) |
| Business requirements document | ISO/IEC/IEEE 29148, IIBA BABOK v3, ISO/IEC 25010 qualities |
| Technical specification | arc42, ISO/IEC/IEEE 42010, the C4 model |
| API specification | OpenAPI 3.1, RFC 9110, RFC 9457 |
| Minutes of meeting, technical report | decisions and actions with owners; findings first |
| Presentation | assertion-evidence slides, WCAG 2.2 contrast |

## Install

```bash
git clone https://github.com/codefin-lab/imprint.git && cd imprint && ./setup.sh
```

Then in Claude Code:

```text
/plugin marketplace add codefin-lab/imprint
/plugin install imprint@imprint
/plugin install diagram-design@imprint
```

Ask for what you need: "write a BRD for customer onboarding", "turn these notes into minutes",
"make a ten-slide deck from this proposal".

## Without Claude

```bash
python3 plugins/imprint/engine/build_docx.py plugins/imprint/templates/brd.md --pdf
python3 plugins/imprint/engine/build_pptx.py plugins/imprint/templates/presentation.md --pdf
```

## Your own brand

A theme is a folder: `theme.yaml` (fonts, sizes, colours, cover labels, default company
details, slide settings), `base.docx` (cover page, styles, header and footer) and `base.pptx`
(slide master, generated from `theme.yaml`). Copy `plugins/imprint/engine/themes/default`,
change it, regenerate the master with `tools/make_base_pptx.py`, and build with
`--theme <folder>`.

To give a team its house style, templates and conventions in one install, put the theme in a
plugin of its own with a skill that names it. Keep that plugin private if your brand is.

## Layout

```text
plugins/imprint/
  skills/docgen/      document skill and per-type references
  skills/slides/      deck skill and presentation reference
  engine/             build_docx.py, build_pptx.py, the imprint package, tools, themes, tests
  templates/          one Markdown template per document type
```

## Contributing

See `CONTRIBUTING.md`. Run `python3 plugins/imprint/engine/tests/run.py` before a pull request.

## License

MIT. See `LICENSE`.
