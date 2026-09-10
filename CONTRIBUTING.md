# Contributing

## Before a pull request

```bash
python3 plugins/imprint/engine/tests/run.py
```

It builds every template as .docx and .pptx with the default theme, runs the verifier, checks
that two builds are identical, and fails if a company-specific or client name appears in the
repository. Then render the PDFs of anything you changed and look at them; most layout
problems are only visible on the page.

## Rules

- **Design lives in themes, never in the engine.** A colour, size, label or company detail in
  Python code is a bug.
- **Deterministic output.** No timestamps, random names or unordered iteration in anything
  written to a file.
- **No client or company data.** Examples use "Example Project" and "Example Client". Never
  commit a real client's document, name or diagram, not even in history.
- **Plain text** in templates, skills and documentation: no section signs, emoji or decorative
  symbols.
- **Standards cited correctly.** When guidance rests on a standard, name it with its number
  and year, and say only what it says.
