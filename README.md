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

## Requirements

| Needed for | What | macOS | Debian or Ubuntu |
| :-- | :-- | :-- | :-- |
| the engine | Python 3.11 or later, and uv (or pipx or pip) | `brew install python uv` | `apt install python3 pipx` |
| PDFs and the table of contents | LibreOffice | `brew install --cask libreoffice` | `apt install libreoffice` |
| reviewing PDFs | poppler | `brew install poppler` | `apt install poppler-utils` |
| the default theme | the Sarabun and Anuphan fonts | from Google Fonts | from Google Fonts |
| installing skills | Node.js, for `npx` | `brew install node` | `apt install nodejs npm` |

The engine's own Python packages are pinned and installed with it; nothing else to install.

## Install

**With [skills](https://github.com/vercel-labs/skills)**, for Claude Code and other agents:

```bash
npx skills add codefin-lab/imprint -g
npx skills add cathrynlavery/diagram-design -g
```

`-g` installs for your user; leave it out to install into the current project only. The first
time a skill builds something, its `scripts/ensure-engine.sh` installs the engine (the
`imprint` command) at the version the skill needs. To install it straight away:

```bash
bash ~/.claude/skills/docgen/scripts/ensure-engine.sh
```

**As a Claude Code plugin** instead:

```text
/plugin marketplace add codefin-lab/imprint
/plugin install imprint@imprint
/plugin install diagram-design@imprint
```

**The command only**, without any skills:

```bash
uv tool install "git+https://github.com/codefin-lab/imprint@v0.2.0#subdirectory=plugins/imprint/engine"
```

## Check the installation

```bash
imprint --version
imprint doctor
```

`imprint doctor` lists anything missing (LibreOffice, poppler, fonts) with how to install it,
and says whether a newer release exists. If `imprint` is not found, add `~/.local/bin` to your
`PATH`.

Then ask your agent for what you need: "write a BRD for customer onboarding", "turn these
notes into minutes", "make a ten-slide deck from this proposal".

## Upgrade

```bash
npx skills update -g
```

That fetches the new skills; each one pins the engine version it was written for, so the next
build upgrades the engine to match. Plugin users: `claude plugin marketplace update imprint`,
then `claude plugin update imprint@imprint`, and restart. What changed in each release:
`CHANGELOG.md`.

## Uninstall

```bash
npx skills remove docgen slides -g
uv tool uninstall imprint-engine        # or: pipx uninstall imprint-engine
```

Plugin users: `claude plugin uninstall imprint@imprint`.

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
