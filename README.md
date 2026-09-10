# Imprint

Documents, diagrams and slides from Markdown, in your own house style.

Imprint is a set of agent skills and the engine behind them. You write, or ask Claude to
write, a Markdown file; Imprint builds a Word document or a PowerPoint deck from it in your
company's design, the same bytes every time. Each document type comes with guidance based on
general international standards, so a BRD reads like a BRD and an API specification like one.

| Skill | What it does |
| :-- | :-- |
| `docgen` | Markdown to .docx: proposals, BRDs, technical and API specifications, minutes, reports |
| `slides` | Markdown to .pptx on a generated slide master |
| `diagram-design` | editorial diagrams, by [Cathryn Lavery](https://github.com/cathrynlavery/diagram-design) |

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

With [skills](https://github.com/vercel-labs/skills), for Claude Code and other agents:

```bash
npx skills add codefin-lab/imprint -g
npx skills add cathrynlavery/diagram-design -g
```

The first time a skill builds something, its `scripts/ensure-engine.sh` installs the engine
(the `imprint` command) with `uv`, `pipx` or `pip`. LibreOffice (for PDFs), poppler and the
default theme's fonts (Sarabun, Anuphan) are yours to install; `imprint doctor` lists what is
missing.

Or as a Claude Code plugin:

```text
/plugin marketplace add codefin-lab/imprint
/plugin install imprint@imprint
/plugin install diagram-design@imprint
```

Then ask for what you need: "write a BRD for customer onboarding", "turn these notes into
minutes", "make a ten-slide deck from this proposal".

## Upgrade

```bash
npx skills update -g
```

That fetches the new skills; each one pins the engine version it was written for, so the next
build upgrades the engine to match. Plugin users: `/plugin marketplace update imprint`.
`imprint doctor` says when a newer release exists. Releases and what changed:
`CHANGELOG.md`.

## The command

```bash
imprint new brd docs/BRD-Onboarding.md      # start from a template (imprint new lists them)
imprint docx docs/BRD-Onboarding.md --pdf
imprint new presentation deck.md && imprint pptx deck.md --pdf
imprint doctor
```

Without the skills: `uv tool install "git+https://github.com/codefin-lab/imprint@v0.2.0#subdirectory=plugins/imprint/engine"`.

## Your own brand

A theme is a folder: `theme.yaml` (fonts, sizes, colours, cover labels, default company
details, slide settings), `base.docx` (cover page, styles, header and footer) and `base.pptx`
(slide master, generated with `imprint make-base-pptx`). Copy the default theme (`imprint
themes` prints where it is), change it, and build with `--theme <folder>`.

To give a team its house style, templates and conventions in one install, put the theme and
templates in a skill of its own that names them, and pin the engine in its own
`scripts/ensure-engine.sh`. Keep that repository private if your brand is.

## Layout

```text
plugins/imprint/
  skills/docgen/      document skill, per-type references, ensure-engine.sh
  skills/slides/      deck skill, presentation reference, ensure-engine.sh
  engine/             the imprint package: command, themes, templates, tests
scripts/release.py    set a new version everywhere at once
setup.sh              install the engine from this checkout, check LibreOffice and fonts
```

## Contributing

See `CONTRIBUTING.md` and `RELEASING.md`. Run `python3 plugins/imprint/engine/tests/run.py`
before a pull request.

## License

MIT. See `LICENSE`.
