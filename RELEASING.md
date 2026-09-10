# Releasing

One version number covers the skills and the engine. A release is a git tag; nothing is
published to a package index.

## Steps

```bash
python3 scripts/release.py 0.3.0          # engine __version__, plugin.json, every ensure-engine.sh
# add the release to CHANGELOG.md
python3 plugins/imprint/engine/tests/run.py --pdf
git commit -am "Release 0.3.0"
git tag -a v0.3.0 -m "Release 0.3.0"     # annotated: --follow-tags pushes only these
git push --follow-tags
```

The test suite fails if the version numbers disagree, so a release cannot go out half-bumped.

## How users get it

| Installed with | They run | What happens |
| :-- | :-- | :-- |
| `npx skills` | `npx skills update -g` | new SKILL.md files; the next build's `ensure-engine.sh` installs engine `v0.3.0` |
| Claude Code plugin | `/plugin marketplace update imprint`, then reinstall the plugin | the same, through the plugin |
| the command only | `uv tool install --force "git+https://github.com/codefin-lab/imprint@v0.3.0#subdirectory=plugins/imprint/engine"` | the engine only |

`imprint doctor` tells anyone on an older version that a newer release exists.

## What counts as a breaking change

Semantic Versioning, judged from the Markdown a user writes and the theme folder a brand
maintains:

- **Major**: existing Markdown or an existing theme builds differently or fails.
- **Minor**: new syntax, commands, templates or theme keys; old documents build the same.
- **Patch**: fixes that change no correct output.

A brand plugin pins the engine in its own `ensure-engine.sh`; raise it when the brand needs a
new feature, not on every engine release.
