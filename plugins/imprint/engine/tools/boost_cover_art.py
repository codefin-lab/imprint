#!/usr/bin/env python3
"""Restore a theme's cover line art from the pristine export, and fix its z-order.

    python3 tools/boost_cover_art.py --theme default
    python3 tools/boost_cover_art.py --theme default --peak 235 --grow 5

By default both geometric drawings are put back exactly as the Google Docs
export drew them — thin, faint hairlines — and the only change made is z-order:
the bottom-right one ships in FRONT of the text and crosses the disclaimer and
the address, so every floating drawing is sent behind. That leaves the artwork
looking untouched while nothing sits on top of the logo.

`--peak` (with optional `--grow`) is the heavier pass: it widens each stroke and
lifts the midtones with a gamma curve. Scaling alpha linearly barely shows —
the strokes are antialiased, so most pixels sit near the original median of 53
rather than near the peak — which is why the curve is there and not a multiply.

Either way the tool reads from `base.docx.bak` when one is present, so repeated
runs never compound. Re-run it after `make_base.py`, which replaces base.docx
with a fresh export.
"""
from __future__ import annotations

import argparse
import io
import zipfile
from pathlib import Path

import docx
import numpy as np
from PIL import Image, ImageFilter
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent.parent

# filename -> peak alpha.  The top-left drawing has clean white space to itself
# and carries the geometric character; the bottom-right one sits under the logo
# and the address, so it stays a whisper or it reads as clutter.
ART = ("word/media/image1.png", "word/media/image2.png")
GAMMA = 0.55


def _drop_drawings(base: Path, media: set[str]) -> int:
    """Delete the runs that place `media`, and the images themselves."""
    if not media:
        return 0
    d = docx.Document(str(base))
    part = d.part
    doomed = {rid for rid, rel in part.rels.items()
              if rel.reltype.endswith("/image")
              and "word/" + rel.target_part.partname.lstrip("/").split("word/")[-1] in media
              or (rel.reltype.endswith("/image") and str(rel.target_part.partname) in
                  {"/" + m for m in media})}
    removed = 0
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    for run in list(d.element.body.iter(qn("w:r"))):
        blips = [b for b in run.iter(qn("a:blip"))]
        if any(b.get(R + "embed") in doomed for b in blips):
            parent = run.getparent()
            if parent is not None:
                parent.remove(run)
                removed += 1
    if removed:
        d.save(str(base))
    return removed


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--theme", default="default")
    ap.add_argument("--peak", type=int, default=None,
                    help="target peak alpha; omit to keep the artwork exactly as exported")
    ap.add_argument("--grow", type=int, default=1,
                    help="MaxFilter size used with --peak; 1 leaves the stroke width alone")
    args = ap.parse_args()

    base = HERE / "themes" / args.theme / "base.docx"
    if not base.exists():
        print(f"not found: {base}")
        return 1
    pristine = base.with_suffix(".docx.bak")
    source = pristine if pristine.exists() else base

    with zipfile.ZipFile(source) as z:
        original = {i.filename: z.read(i.filename) for i in z.infolist()}
    with zipfile.ZipFile(base) as z:
        items = list(z.infolist())
        current = {i.filename: z.read(i.filename) for i in items}

    for name in ART:
        if name not in original:
            print(f"  skip {name} (not in {source.name})")
            continue
        if args.peak is None:                      # put the export's own art back
            current[name] = original[name]
            a = np.array(Image.open(io.BytesIO(original[name])).convert("RGBA"))
            nz = a[..., 3][a[..., 3] > 0]
            print(f"  {name}: restored as exported (peak alpha {int(nz.max())})")
            continue
        a = np.array(Image.open(io.BytesIO(original[name])).convert("RGBA"))
        alpha = Image.fromarray(a[..., 3])
        if args.grow > 1:
            alpha = alpha.filter(ImageFilter.MaxFilter(args.grow))
        al = np.array(alpha).astype(np.float32)
        was = al.max()
        al = args.peak * (al / was) ** GAMMA
        a[..., 3] = np.clip(al, 0, 255).astype(np.uint8)
        a[..., :3] = 0
        buf = io.BytesIO()
        Image.fromarray(a, "RGBA").save(buf, "PNG", optimize=True)
        current[name] = buf.getvalue()
        print(f"  {name}: peak alpha {was:.0f} -> {args.peak}")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for info in items:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            z.writestr(zi, current[info.filename])
    base.write_bytes(out.getvalue())

    d = docx.Document(str(base))
    moved = 0
    for anchor in d.element.body.iter(qn("wp:anchor")):
        if anchor.get("behindDoc") != "1":
            anchor.set("behindDoc", "1")
            anchor.set("layoutInCell", "1")
            moved += 1
    if moved:
        d.save(str(base))
    print(f"  anchors sent behind the text: {moved}")
    print(f"{base}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
