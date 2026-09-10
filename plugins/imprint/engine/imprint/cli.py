"""The imprint command.

    imprint new <type> [path]        start a document from a template (no type: list them)
    imprint docx <file.md> [...]     build a .docx, and a PDF with --pdf
    imprint pptx <file.md> [...]     build a .pptx, and a PDF with --pdf
    imprint themes                   list the themes that ship with the engine, with their folders
    imprint make-base-pptx --theme <name or folder>    regenerate a theme's slide master
    imprint make-base-docx ...       rebuild a theme's base.docx from a Word export
    imprint doctor                   check the installation and look for a newer release
    imprint --version
"""
from __future__ import annotations

import re
import shutil
import subprocess
import sys
from pathlib import Path

from . import __version__

PKG = Path(__file__).resolve().parent
TEMPLATES = PKG / "templates"
REPO = "https://github.com/codefin-lab/imprint"
IMAGE = re.compile(r"!\[[^\]]*\]\(([^)\s]+)\)")


def _new(argv: list[str]) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="imprint new", description="Start a document from a template.")
    ap.add_argument("type", nargs="?", help="template name; leave out to list them")
    ap.add_argument("path", nargs="?", type=Path, help="where to write it (default: ./<type>.md)")
    ap.add_argument("--templates", type=Path, default=TEMPLATES,
                    help="another template folder, such as a brand skill's templates")
    ap.add_argument("--force", action="store_true", help="overwrite an existing file")
    args = ap.parse_args(argv)
    names = sorted(p.stem for p in args.templates.glob("*.md"))
    if not args.type:
        print("\n".join(names))
        return 0
    if args.type not in names:
        print(f"no template '{args.type}' in {args.templates}; choose from: {', '.join(names)}", file=sys.stderr)
        return 1
    src = args.templates / f"{args.type}.md"
    dest = args.path or Path(f"{args.type}.md")
    if dest.is_dir():
        dest = dest / f"{args.type}.md"
    if dest.exists() and not args.force:
        print(f"{dest} exists; pass --force to overwrite", file=sys.stderr)
        return 1
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(src, dest)
    # a template's example picture travels with it, so the first build works
    for ref in IMAGE.findall(src.read_text(encoding="utf-8")):
        image, target = args.templates / ref, dest.parent / ref
        if image.is_file() and not target.exists():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(image, target)
    print(dest)
    return 0


def _themes(argv: list[str]) -> int:
    from .theme import THEMES_DIR
    for d in sorted(p for p in THEMES_DIR.iterdir() if (p / "theme.yaml").exists()):
        print(f"{d.name:12} {d}")
    return 0


def _latest_release() -> str | None:
    """The newest vX.Y.Z tag, from the GitHub API: plain HTTPS, so no git credentials
    (a stale token for another account would turn even a public read into a 403)."""
    import json
    import urllib.request
    api = REPO.replace("https://github.com/", "https://api.github.com/repos/") + "/tags?per_page=100"
    try:
        with urllib.request.urlopen(urllib.request.Request(api, headers={"User-Agent": "imprint"}),
                                    timeout=10) as r:
            names = [t["name"] for t in json.load(r)]
    except (OSError, ValueError, KeyError):
        return None
    tags = [n[1:] for n in names if re.fullmatch(r"v\d+\.\d+\.\d+", n)]
    return max(tags, key=lambda v: tuple(map(int, v.split(".")))) if tags else ""


def _doctor(argv: list[str]) -> int:
    problems = 0

    def line(ok: bool, what: str, fix: str = ""):
        nonlocal problems
        problems += 0 if ok else 1
        print(f"  {'ok     ' if ok else 'MISSING'}  {what}" + ("" if ok else f": {fix}"))

    print(f"imprint {__version__}  ({PKG})")
    for mod, pkg in (("docx", "python-docx"), ("pptx", "python-pptx"), ("yaml", "PyYAML"),
                     ("matplotlib", "matplotlib"), ("PIL", "pillow"), ("lxml", "lxml")):
        try:
            __import__(mod)
            line(True, pkg)
        except ImportError:
            line(False, pkg, "reinstall the engine")
    soffice = shutil.which("soffice") or (
        "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").exists() else None)
    line(bool(soffice), "LibreOffice (PDF and table of contents)",
         "brew install --cask libreoffice, or apt install libreoffice")
    line(bool(shutil.which("pdftoppm")), "poppler (reviewing PDFs)", "brew install poppler, or apt install poppler-utils")
    fonts = ""
    if shutil.which("fc-list"):
        fonts = subprocess.run(["fc-list"], capture_output=True, text=True).stdout.lower()
    for d in (Path.home() / "Library/Fonts", Path("/Library/Fonts")):
        if d.is_dir():
            fonts += " ".join(p.name.lower() for p in d.iterdir())
    for font in ("Sarabun", "Anuphan"):
        line(font.lower() in fonts, f"font {font} (default theme)", f"https://fonts.google.com/specimen/{font}")
    latest = _latest_release()
    if latest is None:
        print("  ?        could not reach GitHub to check for a newer release")
    elif not latest:
        print("  ok       no release published yet")
    elif tuple(map(int, latest.split("."))) > tuple(map(int, __version__.split("."))):
        print(f"  NEWER    imprint {latest} is available. Update the skills (npx skills update -g), and the\n"
              f"           next build upgrades the engine; or run the skill's scripts/ensure-engine.sh")
    else:
        print(f"  ok       latest release is {latest}")
    return 1 if problems else 0


def _run(module: str, argv: list[str]) -> int:
    import importlib
    return importlib.import_module(f"imprint.{module}").main(argv)


COMMANDS = {
    "new": _new,
    "docx": lambda a: _run("build_docx", a),
    "pptx": lambda a: _run("build_pptx", a),
    "themes": _themes,
    "make-base-pptx": lambda a: _run("tools.make_base_pptx", a),
    "make-base-docx": lambda a: _run("tools.make_base", a),
    "boost-cover-art": lambda a: _run("tools.boost_cover_art", a),
    "redraw-cover-art": lambda a: _run("tools.redraw_cover_art", a),
    "doctor": _doctor,
}


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if not argv or argv[0] in ("-h", "--help", "help"):
        print(__doc__.strip())
        return 0
    if argv[0] in ("-V", "--version", "version"):
        print(f"imprint {__version__}")
        return 0
    command = COMMANDS.get(argv[0])
    if command is None:
        print(f"unknown command '{argv[0]}'; run imprint --help", file=sys.stderr)
        return 2
    return command(argv[1:]) or 0


if __name__ == "__main__":
    raise SystemExit(main())
