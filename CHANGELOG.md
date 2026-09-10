# Changelog

Versions follow Semantic Versioning. The skills and the engine share one version: a skill
asks for the engine release with the same number, and installs it when it is missing.

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
