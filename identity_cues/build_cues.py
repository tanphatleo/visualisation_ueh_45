"""
The four visual channels that identify a mark: spatial region, colour hue,
motion, shape.

One row per channel — a label on the left, and on the right the smallest set of
marks that shows what that channel can carry. The rows are deliberately flat
illustrations rather than a chart: nothing here encodes a number, so there is no
axis, no scale and no legend.

Motion is the awkward one. It is the only channel that does not survive being
printed, so it is drawn the way comics draw it: a mark plus a rotation arc. The
arc is a glyph standing in for the channel, not a path anything travels.

Run:  python build_cues.py            # identity_cues.png, transparent
      python build_cues.py --bg white
"""

import argparse
from pathlib import Path as FilePath

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Circle, FancyArrowPatch, Rectangle, RegularPolygon
from matplotlib.path import Path

HERE = FilePath(__file__).resolve().parent

INK = "#3F4142"          # the neutral mark colour, used wherever hue is NOT the point
LABEL = "#404040"
CONTEXT_LINE = "#8C8C8C"  # the source line — same grey as the effectiveness figure
HUES = ["#EFC31B", "#D8532B", "#22BC8A", "#2E86C8"]

# The drawing lives in a 100 x 58 box, equal-aspect, so every size below is in
# the same units and rows stay comparable.
W, H = 100.0, 58.0
ROW_Y = [50.0, 36.5, 23.0, 9.5]      # centre line of each row, top to bottom
LABEL_X = 1.0
FONT_PT = 17
SOURCE_PT = 10.5
# Anchored by its TOP, just under the shape row (which bottoms out near y = 5.6),
# and wrapped by hand. Set on one line it runs wider than the marks do, and
# bbox_inches="tight" then widens the whole canvas to fit it, which leaves the
# marks stranded in the middle of a letterbox.
SOURCE_Y = 3.4


def square(ax, x, y, size, colour):
    """A square centred on (x, y) — the centring is what keeps rows aligned."""
    ax.add_patch(Rectangle((x - size / 2, y - size / 2), size, size,
                           facecolor=colour, edgecolor="none"))


def plus(ax, x, y, size, colour, thick=0.34):
    """A cross, as two overlapping bars. Same colour and no edge, so no seam."""
    t = size * thick
    ax.add_patch(Rectangle((x - size / 2, y - t / 2), size, t,
                           facecolor=colour, edgecolor="none"))
    ax.add_patch(Rectangle((x - t / 2, y - size / 2), t, size,
                           facecolor=colour, edgecolor="none"))


def spin(ax, x, y, r, colour, t0=-35, t1=245, lw=1.3, head=7):
    """
    A rotation arc with an arrowhead — the conventional glyph for motion.

    Drawn as a FancyArrowPatch over an explicit arc path rather than as an Arc
    plus a separate triangle: that way the head is placed on the curve's own
    tangent and stays attached at any radius.
    """
    t = np.radians(np.linspace(t0, t1, 60))
    path = Path(np.column_stack([x + r * np.cos(t), y + r * np.sin(t)]))
    ax.add_patch(FancyArrowPatch(path=path, arrowstyle="-|>", mutation_scale=head,
                                 lw=lw, color=colour, shrinkA=0, shrinkB=0,
                                 joinstyle="round", capstyle="round"))


def row_spatial(ax, y):
    """Same mark, three sizes, staggered — position and extent, not colour."""
    for x, dy, s in [(57.5, 3.4, 4.6), (66.0, -1.2, 7.4), (78.5, 0.4, 11.0)]:
        square(ax, x, y + dy, s, INK)


def row_hue(ax, y):
    """Identical squares; hue is the only thing separating them."""
    size, gap, x = 8.0, 4.4, 55.0
    for colour in HUES:
        square(ax, x + size / 2, y, size, colour)
        x += size + gap


def row_motion(ax, y):
    """Dots of one size and colour — two of them spinning."""
    dots = [(60, 1.8), (68.5, -1.6), (73.5, -2.4), (80, -1.2), (86, -1.0), (93, -0.6)]
    for x, dy in dots:
        ax.add_patch(Circle((x, y + dy), 1.2, facecolor=INK, edgecolor="none"))
    spin(ax, 60, y + 1.8, 3.1, INK)
    spin(ax, 86, y - 1.0, 3.1, INK)


def row_shape(ax, y):
    """One size, one colour; only the outline differs."""
    plus(ax, 57.5, y, 7.8, INK)
    ax.add_patch(Circle((69.5, y), 3.8, facecolor=INK, edgecolor="none"))
    square(ax, 81.0, y, 7.4, INK)
    # radius, not side — nudged up so the triangle reads the same visual weight
    ax.add_patch(RegularPolygon((92.5, y - 0.4), 3, radius=4.5,
                                facecolor=INK, edgecolor="none"))


ROWS = [
    ("Spatial region", row_spatial),
    ("Color hue", row_hue),
    ("Motion", row_motion),
    ("Shape", row_shape),
]


def build(bg):
    fig, ax = plt.subplots(figsize=(6.0, 6.0 * H / W), dpi=300)
    for (text, draw), y in zip(ROWS, ROW_Y):
        ax.text(LABEL_X, y, text, ha="left", va="center",
                fontsize=FONT_PT, color=LABEL)
        draw(ax, y)

    ax.text(LABEL_X, SOURCE_Y,
            "After Munzner, Visualization Analysis and Design (2014);\n"
            "ranking after Cleveland & McGill (1984)",
            ha="left", va="top", linespacing=1.5,
            fontsize=SOURCE_PT, color=CONTEXT_LINE)

    ax.set_xlim(0, W)
    ax.set_ylim(0, H)
    ax.set_aspect("equal")
    ax.axis("off")
    if bg != "none":
        fig.patch.set_facecolor(bg)
    fig.subplots_adjust(0, 0, 1, 1)
    return fig


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--out", default="identity_cues.png")
    ap.add_argument("--bg", default="none", help="'none' for transparent, or a colour")
    a = ap.parse_args()

    fig = build(a.bg)
    fig.savefig(HERE / a.out, dpi=300, transparent=(a.bg == "none"),
                bbox_inches="tight", pad_inches=0.04)
    print(f"wrote {a.out}")


if __name__ == "__main__":
    main()
