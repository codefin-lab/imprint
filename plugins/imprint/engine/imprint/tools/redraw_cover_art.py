#!/usr/bin/env python3
"""Redraw the cover's geometric line art at a single, even stroke weight.

    python3 tools/redraw_cover_art.py --theme default
    python3 tools/redraw_cover_art.py --theme default --width-pt 0.6 --alpha 0.45

The art that came out of the Google Docs export is a rasterised PNG, so every
stroke's weight depends on its angle: measured across image1, a horizontal run
through a stroke is anywhere from 1 to 18 pixels wide (median 3) and its peak
alpha ranges 87-119.  The two drawings do not even agree with each other —
image1 peaks at alpha 119, image2 at 51, less than half.  That is what reads as
uneven; no amount of post-processing fixes it, because the unevenness is baked
into the pixels.

So the geometry is recovered once (Hough over the original, endpoints snapped
so shared vertices meet) and redrawn as real strokes: one width, one colour,
one opacity, for both drawings.  Composition and placement are unchanged.
"""
from __future__ import annotations

from ..theme import theme_dir

import argparse
import io
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent

# Normalised (x, y) in [0,1] of the source PNGs, recovered from the export by
# Hough transform.  Each line's extent is the CONTIGUOUS run of ink along it,
# not merely "ink near the infinite line" — a stroke that happens to graze the
# line would otherwise drag the endpoint out with it — and the fragments Hough
# returns are clustered onto one line first, since deduplicating them by
# (angle, offset) is what left lines broken off short.  Verified against the
# original: 100% of its ink lies on one of these, and 100% of their length has
# ink under it.  Placed size sets the aspect ratio.
DRAWINGS = {
    "word/media/image1.png": {
        "size_in": (2.88, 3.72),
        "segments": [
            ((0.3165, 0.9904), (0.9851, 0.3863)),
            ((0.1513, 0.5824), (0.7291, -0.0004)),
            ((0.1496, 0.582), (0.985, 0.3865)),
            ((0.1524, 0.5802), (0.3187, 0.9901)),
            ((0.8122, -0.0001), (0.9845, 0.3884)),
            ((0.3974, -0.0), (0.4286, 0.3046)),
            ((-0.0002, 0.9417), (0.3192, 0.9895)),
            ((-0.0001, 0.5581), (0.1566, 0.5815)),
        ],
    },
    "word/media/image2.png": {
        "size_in": (3.15, 2.40),
        "segments": [
            ((0.0501, 0.0003), (0.6548, 0.7137)),
            ((0.0502, 0.0), (0.2626, 0.9983)),
            ((0.6464, -0.0014), (0.9984, 0.7053)),
            ((0.0503, 0.0), (0.649, 0.0)),
            ((0.6465, 0.7125), (0.6459, -0.0)),
            ((0.6434, 0.7053), (0.9986, 0.7031)),
            ((0.6388, 0.6886), (0.7764, 0.998)),
            ((0.5348, 0.9985), (0.6498, 0.6956)),
        ],
    },
}


def _snap(segments, tol=0.03):
    """Put every shared corner on the intersection of the lines that meet there.

    Hough returns each stroke independently, so a corner arrives as two or three
    endpoints a little apart.  Averaging them (the obvious fix) drags an endpoint
    off its own line — up to 3.75pt on image2 — and the stroke then runs past the
    corner instead of stopping at it.  Solving for the point closest to all the
    lines meeting there lands exactly on the intersection, so the join is clean
    from both sides.
    """
    import numpy as np

    lines = []                      # (point on line, unit direction) per segment
    for (x1, y1), (x2, y2) in segments:
        a = np.array([x1, y1], float)
        d = np.array([x2 - x1, y2 - y1], float)
        d /= np.hypot(*d) or 1.0
        lines.append((a, d))

    ends = [(i, e, np.array(segments[i][e], float))
            for i in range(len(segments)) for e in (0, 1)]
    clusters: list[list[tuple]] = []
    for item in ends:
        for c in clusters:
            centre = np.mean([q[2] for q in c], axis=0)
            if np.hypot(*(item[2] - centre)) <= tol:
                c.append(item)
                break
        else:
            clusters.append([item])

    fixed = [list(seg) for seg in segments]
    for c in clusters:
        members = {i for i, _, _ in c}
        centre = np.mean([q[2] for q in c], axis=0)
        if len(members) >= 2:
            # least squares over the lines: minimise the distance to each one
            A = np.zeros((2, 2))
            b = np.zeros(2)
            for i in members:
                _, d = lines[i]
                n = np.array([-d[1], d[0]])
                A += np.outer(n, n)
                b += n * (n @ lines[i][0])
            try:
                centre = np.linalg.solve(A, b)
            except np.linalg.LinAlgError:
                pass
        v = (round(float(centre[0]), 4), round(float(centre[1]), 4))
        for i, e, _ in c:
            fixed[i][e] = v
    return [tuple(seg) for seg in fixed]


def _bleed(segments, margin, tol=0.004):
    """Run a free endpoint that sits on an edge out past it, along its own line.

    Only a FREE endpoint — one no other segment shares.  A corner that happens
    to sit on the edge is a corner: extending the lines through it makes each
    one overshoot the other, which is exactly the little overhang you see at a
    join that should be mitred.
    """
    from collections import Counter
    shared = Counter(p for seg in segments for p in seg)

    out = []
    for a, b in segments:
        (x1, y1), (x2, y2) = a, b
        dx, dy = x2 - x1, y2 - y1
        length = (dx * dx + dy * dy) ** 0.5 or 1.0
        ux, uy = dx / length, dy / length
        if shared[a] == 1 and min(x1, 1 - x1, y1, 1 - y1) <= tol:
            x1, y1 = x1 - ux * margin * 2, y1 - uy * margin * 2
        if shared[b] == 1 and min(x2, 1 - x2, y2, 1 - y2) <= tol:
            x2, y2 = x2 + ux * margin * 2, y2 + uy * margin * 2
        out.append(((x1, y1), (x2, y2)))
    return out


def render(spec, width_pt: float, colour: str, alpha: float, dpi: int,
           margin: float = 0.006) -> bytes:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    w_in, h_in = spec["size_in"]
    fig = plt.figure(figsize=(w_in, h_in), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    # A stroke lying exactly on the canvas edge gets half of itself cut off and
    # renders at half weight — the top edge of image2 was doing precisely that.
    # A hair of margin keeps every stroke whole; endpoints that sat on an edge
    # are pushed out past it so the drawing still bleeds off the page.
    ax.set_xlim(-margin, 1 + margin)
    ax.set_ylim(1 + margin, -margin)
    ax.axis("off")
    for (x1, y1), (x2, y2) in _bleed(_snap(spec["segments"]), margin):
        ax.plot([x1, x2], [y1, y2], color=colour, alpha=alpha,
                linewidth=width_pt, solid_capstyle="round", clip_on=False,
                antialiased=True)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True, pad_inches=0, dpi=dpi)
    plt.close(fig)
    return buf.getvalue()


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--theme", default="default")
    ap.add_argument("--width-pt", type=float, default=0.35, help="stroke weight in points")
    ap.add_argument("--alpha", type=float, default=1.0)
    ap.add_argument("--colour", default="#BFC3C7")
    ap.add_argument("--dpi", type=int, default=1200)
    ap.add_argument("--out-dir", type=Path, default=None,
                    help="write the PNGs here instead of into base.docx")
    args = ap.parse_args(argv)

    base = theme_dir(args.theme) / "base.docx"
    if not base.exists():
        print(f"not found: {base}")
        return 1

    fresh = {name: render(spec, args.width_pt, args.colour, args.alpha, args.dpi)
             for name, spec in DRAWINGS.items()}

    if args.out_dir:
        args.out_dir.mkdir(parents=True, exist_ok=True)
        for name, data in fresh.items():
            path = args.out_dir / Path(name).name
            path.write_bytes(data)
            print(f"  {path} ({len(data)} bytes)")
        return 0

    with zipfile.ZipFile(base) as z:
        items = list(z.infolist())
        current = {i.filename: z.read(i.filename) for i in items}
    for name, data in fresh.items():
        if name not in current:
            print(f"  skip {name} (not in base.docx)")
            continue
        print(f"  {name}: redrawn at {args.width_pt}pt / alpha {args.alpha} "
              f"({len(current[name])} -> {len(data)} bytes)")
        current[name] = data

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for info in items:
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            z.writestr(zi, current[info.filename])
    base.write_bytes(out.getvalue())
    print(f"{base}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
