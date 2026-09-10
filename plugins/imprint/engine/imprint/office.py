"""LibreOffice, run in a profile of its own.

A conversion started while LibreOffice is already open, as a window or another headless run,
is handed to that instance. It keeps the fonts it loaded when it started, so a font installed
since then is silently replaced, and two conversions can collide. A dedicated user profile gives
every conversion a fresh process of its own and leaves any open LibreOffice alone.
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

MAC_APP = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
PROFILE = Path.home() / ".cache" / "imprint" / "libreoffice"


def find_soffice() -> str | None:
    exe = shutil.which("soffice")
    if exe:
        return exe
    return str(MAC_APP) if MAC_APP.exists() else None


def convert(soffice: str, path: Path, fmt: str = "pdf", outdir: Path | None = None) -> Path:
    """Convert in Imprint's own LibreOffice profile, next to the source unless `outdir` is
    given. `fmt` is a LibreOffice target such as "pdf" or "docx:MS Word 2007 XML"."""
    PROFILE.mkdir(parents=True, exist_ok=True)
    target = Path(outdir) if outdir else path.parent
    subprocess.run([soffice, f"-env:UserInstallation={PROFILE.as_uri()}", "--headless",
                    "--convert-to", fmt, "--outdir", str(target), str(path)],
                   check=True, capture_output=True)
    return target / f"{path.stem}.{fmt.split(':')[0]}"
