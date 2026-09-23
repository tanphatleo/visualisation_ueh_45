"""
A stylised top-view brain drawn as coloured, winding paths with dot terminals.

Generated, not traced.

The structure that makes the reference style work is not "smooth random curves".
It is a MAZE: every path is a self-avoiding walk on a square lattice, which is
why the curves run in evenly spaced parallel tracks, turn through tight hairpins,
and almost never cross each other. Rounding the lattice corners afterwards is
what hides the grid and leaves the hand-drawn, gyrus-like look.

  1. hemisphere_mask()  which lattice cells fall inside the hemisphere
  2. walk()             self-avoiding walk, refusing any cell that touches an
                        occupied one — that one rule creates the even spacing
  3. chaikin()          corner cutting; a lattice zig-zag becomes a rounded curve
  4. colour_regions()   nearest-seed assignment, so hues come in contiguous
                        blobs instead of scattering
  5. clipping           every artist is clipped to the hemisphere patch

Run:  python build_brain.py            # brain.png, transparent
      python build_brain.py 7          # a different brain from seed 7
"""

import sys
from pathlib import Path as FilePath

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Circle, PathPatch
from matplotlib.path import Path as MPath

HERE = FilePath(__file__).resolve().parent
OUT = HERE / "brain.png"

SEED = int(sys.argv[1]) if len(sys.argv) > 1 else 5

# Bright, high-chroma palette. Deliberately NOT the deck's chart palette — this
# is an illustration, not a data graphic, so style_chart.md's one-accent rule
# does not apply.
PALETTE = [
    "#F0B323",   # amber
    "#7B2D8E",   # purple
    "#C0499A",   # magenta
    "#E8225C",   # crimson
    "#8DC63F",   # lime
    "#1E9E4A",   # green
    "#12A5A0",   # teal
    "#1F7A9E",   # steel blue
    "#1D4E9E",   # navy
    "#F07C2B",   # orange
    "#5B2E91",   # deep violet
]

STEP = 0.058         # lattice pitch — sets the spacing of the parallel tracks
JITTER = 0.24        # per-cell offset, as a fraction of STEP
LINE_W = 3.0
DOT_R = 0.026
GAP = 0.030          # half-width of the longitudinal fissure
STRAIGHT_BIAS = 1.7  # how much a walk prefers to carry on rather than turn

# Outer profile of the right hemisphere, frontal pole down to occipital.
OUTER = np.array([
    (0.00, 1.34), (0.20, 1.33), (0.42, 1.25), (0.60, 1.10),
    (0.75, 0.90), (0.86, 0.66), (0.93, 0.40), (0.98, 0.12),
    (0.99, -0.16), (0.96, -0.42), (0.89, -0.64), (0.79, -0.84),
    (0.66, -1.01), (0.50, -1.15), (0.32, -1.25), (0.15, -1.31),
    (0.05, -1.30),
])


def chaikin(pts, iters=3, closed=False):
    """Corner cutting. Turns the lattice zig-zag into rounded, drawn-looking curves."""
    p = np.asarray(pts, dtype=float)
    for _ in range(iters):
        if closed:
            a, b = p, np.roll(p, -1, axis=0)
        else:
            a, b = p[:-1], p[1:]
        q = 0.75 * a + 0.25 * b
        r = 0.25 * a + 0.75 * b
        cut = np.empty((len(q) + len(r), 2))
        cut[0::2], cut[1::2] = q, r
        p = cut if closed else np.vstack([p[0], cut, p[-1]])
    return p


def hemisphere_outline(side=1):
    """Closed hemisphere: down the outer edge, back up the fissure."""
    y = np.linspace(OUTER[-1, 1], OUTER[0, 1], 30)
    t = np.linspace(0, 1, 30)
    inner = np.column_stack([GAP + 0.016 * np.sin(t * 4.5), y])
    poly = np.vstack([OUTER, inner])
    poly[:, 0] *= side
    smooth = chaikin(poly, iters=3, closed=True)
    split = int(len(smooth) * len(OUTER) / len(poly))
    return smooth, split


def hemisphere_mask(rng, mpath, side):
    """
    Lattice cells whose centre sits inside the hemisphere, with a small inset.

    Each cell's drawing coordinate is JITTERED off the lattice. The walk still
    runs on a clean grid — that is what keeps the tracks evenly spaced and
    non-crossing — but the drawn points no longer line up on it, so after corner
    cutting nothing looks axis-aligned. Without this the output reads as a
    circuit board rather than as something organic.
    """
    xs = np.arange(GAP + STEP * 0.6, 1.02, STEP) * side
    ys = np.arange(-1.30, 1.34, STEP)
    cells, coords = set(), {}
    for i, x in enumerate(xs):
        for j, y in enumerate(ys):
            # Inset by shrinking toward the hemisphere's own axis, so paths keep
            # clear of the outline instead of running along underneath it.
            probe = [(x, y), (x + STEP * 0.55 * side, y), (x - STEP * 0.55 * side, y),
                     (x, y + STEP * 0.55), (x, y - STEP * 0.55)]
            if all(mpath.contains_point(p) for p in probe):
                cells.add((i, j))
                coords[(i, j)] = (x + rng.uniform(-1, 1) * JITTER * STEP,
                                  y + rng.uniform(-1, 1) * JITTER * STEP)
    return cells, coords


DIRS = [(1, 0), (-1, 0), (0, 1), (0, -1)]


def walk(rng, owner, cells, start, pid, max_len):
    """
    Self-avoiding walk that must not sit NEXT to a DIFFERENT path.

    Two rules, and the distinction between them is the whole trick:
      * no cell may be entered twice, by anyone;
      * a cell may not touch another path's cells, but MAY touch its own.

    Allowing self-contact is what permits the tight hairpin: turn twice and the
    new cell lands diagonally beside the cell two steps back. Forbid that and
    every path comes out dead straight — which is exactly what the first version
    of this script drew.
    """
    def free(c, pid):
        if c not in cells or c in owner:
            return False
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                n = (c[0] + dx, c[1] + dy)
                if owner.get(n, pid) != pid:
                    return False
        return True

    if not free(start, pid):
        return []
    path = [start]
    owner[start] = pid
    heading = DIRS[rng.integers(4)]
    for _ in range(max_len):
        cur = path[-1]
        # Weight straight-ahead heavily; the rest is an even chance of turning.
        opts, wts = [], []
        for d in DIRS:
            if (d[0] == -heading[0] and d[1] == -heading[1]):
                continue                       # never double back
            nxt = (cur[0] + d[0], cur[1] + d[1])
            if free(nxt, pid):
                opts.append((d, nxt))
                wts.append(STRAIGHT_BIAS if d == heading else 1.0)
        if not opts:
            break
        k = rng.choice(len(opts), p=np.array(wts) / sum(wts))
        heading, nxt = opts[k]
        path.append(nxt)
        owner[nxt] = pid
    return path


def colour_regions(rng, coords_list, n=len(PALETTE)):
    """
    Assign hues by nearest seed point, so colours arrive in contiguous blobs.

    Colouring at random scatters every hue everywhere and the illustration turns
    to confetti; the reference style groups them by area.
    """
    pts = np.array(coords_list)
    idx = rng.choice(len(pts), size=min(n, len(pts)), replace=False)
    seeds = pts[idx]
    order = rng.permutation(len(PALETTE))

    def colour_of(xy):
        d = np.hypot(*(seeds - np.asarray(xy)).T)
        return PALETTE[order[int(np.argmin(d)) % len(PALETTE)]]
    return colour_of


def draw_hemisphere(ax, rng, side, max_paths=95):
    poly, split = hemisphere_outline(side)
    mpath = MPath(poly, closed=True)
    clip = PathPatch(mpath, facecolor="none", edgecolor="none")
    ax.add_patch(clip)

    cells, coords = hemisphere_mask(rng, mpath, side)
    colour_of = colour_regions(rng, list(coords.values()))

    # --- outline, coloured by the same regions so it belongs to the drawing ---
    for idx in np.array_split(np.arange(len(poly)), 22):
        seg = poly[np.append(idx, (idx[-1] + 1) % len(poly))]
        ax.plot(seg[:, 0], seg[:, 1], color=colour_of(seg.mean(axis=0)),
                lw=LINE_W, solid_capstyle="round", zorder=3)

    # --- the maze ---------------------------------------------------------
    owner = {}
    order = list(cells)
    rng.shuffle(order)
    drawn = 0
    for start in order:
        if drawn >= max_paths:
            break
        cellpath = walk(rng, owner, cells, start, drawn,
                        max_len=rng.integers(12, 60))
        if len(cellpath) < 3:
            continue
        pts = np.array([coords[c] for c in cellpath])
        curve = chaikin(pts, iters=5)
        colour = colour_of(pts.mean(axis=0))
        line, = ax.plot(curve[:, 0], curve[:, 1], color=colour, lw=LINE_W,
                        solid_capstyle="round", solid_joinstyle="round", zorder=4)
        line.set_clip_path(clip)
        for e in (pts[0], pts[-1]):
            if rng.random() < 0.82:            # most ends get a dot, not all
                dot = Circle(e, DOT_R, color=colour, zorder=5)
                ax.add_patch(dot)
                dot.set_clip_path(clip)
        drawn += 1


def main():
    rng = np.random.default_rng(SEED)
    fig, ax = plt.subplots(figsize=(8, 8), dpi=200)
    fig.patch.set_alpha(0)
    ax.patch.set_alpha(0)

    draw_hemisphere(ax, rng, side=+1)
    draw_hemisphere(ax, rng, side=-1)

    ax.set_xlim(-1.12, 1.12)
    ax.set_ylim(-1.45, 1.45)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.subplots_adjust(0, 0, 1, 1)
    fig.savefig(OUT, transparent=True, dpi=200,
                bbox_inches="tight", pad_inches=0.05)
    print(f"wrote {OUT}  (seed {SEED})")


if __name__ == "__main__":
    main()
