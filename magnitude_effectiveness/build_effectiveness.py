"""
The effectiveness ranking of the magnitude channels — redrawn from
Munzner, Visualization Analysis and Design (2014), which builds on
Cleveland & McGill (1984).

A schematic, not a data chart: every row is a picture of one encoding channel,
ordered from the one the eye decodes most accurately (position on a common
scale) down to the least (volume). Colours are the roles from style_chart.md §2.1.

    python build_effectiveness.py     ->  images/effectiveness.png
"""

import colorsys
import math
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import Polygon, Rectangle

# --- colour roles — style_chart.md §2.1 ---------------------------------
INK = "#262A33"           # the marks themselves, and the row labels
TEXT_2 = "#595959"        # the effectiveness axis and its labels
CONTEXT_LINE = "#8C8C8C"  # brackets, source line
CONTEXT_FILL = "#BDBDBD"  # the shaded face of a cube
SURFACE = "#FFFFFF"
ALERT = "#990F3D"         # the hue the saturation ramp runs to

# --- geometry -----------------------------------------------------------
# x and y are in the same units (the aspect is equal), so a square is a square.
MX0, MX1 = 54.0, 112.0    # the column the marks live in
MW = MX1 - MX0
ROW_H = 9.0
TOP_Y = 94.0
LABEL_X = 0.0

ROWS = [
    "Position on common scale",
    "Position on unaligned scale",
    "Length (1D size)",
    "Tilt/angle",
    "Area (2D size)",
    "Depth (3D position)",
    "Color luminance",
    "Color saturation",
    "Curvature",
    "Volume (3D size)",
]

# Luminance ramp: no hue at all, four steps of lightness.
LUMINANCE = ["#F7F7F7", "#BDBDBD", "#8C8C8C", "#262A33"]


def saturation_ramp(hex_colour, steps=(0.12, 0.40, 0.70, 1.00)):
    """Four swatches of one hue, from a near-white tint to the hue at full.

    One hue throughout is the point of the row — what changes is how much of
    it there is. (A constant-lightness ramp would be the purer demonstration,
    but its pale end reads as grey, which is the row above.)
    """
    r, g, b = (int(hex_colour[i:i + 2], 16) / 255 for i in (1, 3, 5))
    return [(1 - f + f * r, 1 - f + f * g, 1 - f + f * b) for f in steps]


SATURATION = saturation_ramp(ALERT)


def row_y(i):
    return TOP_Y - i * ROW_H


# --- the marks, one function per row ------------------------------------
def scale_mark(ax, x0, x1, y, dot_frac):
    """A measured line with end caps and a point on it."""
    ax.plot([x0, x1], [y, y], color=INK, lw=1.2, solid_capstyle="butt", zorder=2)
    for x in (x0, x1):
        ax.plot([x, x], [y - 1.1, y + 1.1], color=INK, lw=1.2, zorder=2)
    ax.plot([x0 + dot_frac * (x1 - x0)], [y], marker="o", ms=6,
            color=INK, zorder=3)


def position_common(ax, y):
    # Both lines share one scale: the ends line up, so the two dots can be
    # compared directly.
    scale_mark(ax, MX0, MX1, y + 2.4, 0.61)
    scale_mark(ax, MX0, MX1, y - 2.4, 0.18)


def position_unaligned(ax, y):
    # The same marks on two scales that do not line up — the comparison now
    # costs the reader an extra step.
    scale_mark(ax, MX0, MX0 + 0.42 * MW, y + 2.4, 0.44)
    scale_mark(ax, MX0 + 0.30 * MW, MX1, y - 2.4, 0.48)


def length(ax, y):
    for f0, f1 in [(0.00, 0.13), (0.20, 0.41), (0.50, 1.00)]:
        ax.plot([MX0 + f0 * MW, MX0 + f1 * MW], [y, y],
                color=INK, lw=1.6, solid_capstyle="butt", zorder=2)


def tilt(ax, y):
    seg = 7.0
    for i, deg in enumerate([90, 60, 30, 6]):
        cx = MX0 + (0.10 + 0.26 * i) * MW
        dx = seg / 2 * math.cos(math.radians(deg))
        dy = seg / 2 * math.sin(math.radians(deg))
        ax.plot([cx - dx, cx + dx], [y - dy, y + dy],
                color=INK, lw=1.6, solid_capstyle="butt", zorder=2)


def area(ax, y):
    # Side lengths chosen so the *areas* step 1 : 4 : 9 : 16.
    for i, s in enumerate([2.05, 4.1, 6.15, 8.2]):
        cx = MX0 + (0.06 + 0.29 * i) * MW
        ax.add_patch(Rectangle((cx - s / 2, y - s / 2), s, s,
                               facecolor=INK, edgecolor="none", zorder=2))


def depth(ax, y):
    # Distance along the line of sight — read off a foreshortened axis.
    for x0, x1 in [(MX0, MX0 + 0.26 * MW), (MX0 + 0.38 * MW, MX1)]:
        ax.plot([x0, x0], [y - 1.3, y + 1.3], color=INK, lw=1.2, zorder=2)
        ax.annotate("", xy=(x1 - 0.9, y), xytext=(x0, y),
                    arrowprops=dict(arrowstyle="-|>", color=INK, lw=1.2,
                                    shrinkA=0, shrinkB=0, mutation_scale=9))
        ax.plot([x1], [y], marker="o", ms=6, color=INK, zorder=3)


def swatches(ax, y, colours):
    s = 6.4
    for i, c in enumerate(colours):
        cx = MX0 + (0.08 + 0.28 * i) * MW
        ax.add_patch(Rectangle((cx - s / 2, y - s / 2), s, s, facecolor=c,
                               edgecolor=CONTEXT_LINE, lw=0.5, zorder=2))


def curvature(ax, y):
    arc = 7.0
    for i, turn in enumerate([0.0, 0.9, 1.8, 2.7]):
        cx = MX0 + (0.08 + 0.27 * i) * MW
        if turn == 0:
            xs, ys = [cx, cx], [y - arc / 2, y + arc / 2]
        else:
            r = arc / turn                      # same arc length, tighter bend
            ts = [(-turn / 2) + turn * k / 40 for k in range(41)]
            xs = [cx - r * (1 - math.cos(t)) for t in ts]
            ys = [y + r * math.sin(t) for t in ts]
        ax.plot(xs, ys, color=INK, lw=1.6, solid_capstyle="butt", zorder=2)


def cube(ax, cx, cy, s):
    d = s * 0.36                                # the receding edge, at 45°
    x0, y0 = cx - (s + d) / 2, cy - (s + d) / 2
    faces = [
        ([(x0, y0), (x0 + s, y0), (x0 + s, y0 + s), (x0, y0 + s)], INK),
        ([(x0, y0 + s), (x0 + d, y0 + s + d),
          (x0 + s + d, y0 + s + d), (x0 + s, y0 + s)], SURFACE),
        ([(x0 + s, y0), (x0 + s + d, y0 + d),
          (x0 + s + d, y0 + s + d), (x0 + s, y0 + s)], CONTEXT_FILL),
    ]
    for pts, fill in faces:
        ax.add_patch(Polygon(pts, closed=True, facecolor=fill, edgecolor=INK,
                             lw=0.6, joinstyle="miter", zorder=2))


def volume(ax, y):
    # Side lengths chosen so the *volumes* step 1 : 8 : 27 : 64.
    for i, s in enumerate([1.55, 3.1, 4.65, 6.2]):
        cube(ax, MX0 + (0.06 + 0.29 * i) * MW, y, s)


DRAW = [position_common, position_unaligned, length, tilt, area, depth,
        lambda ax, y: swatches(ax, y, LUMINANCE),
        lambda ax, y: swatches(ax, y, SATURATION),
        curvature, volume]


# --- the two right-hand annotations -------------------------------------
def same_bracket(ax, first_row, last_row):
    """Rows the eye cannot rank against each other are bracketed as equal."""
    x, tick = 116.0, 1.6
    y_top = row_y(first_row) + ROW_H / 2 - 0.5
    y_bot = row_y(last_row) - ROW_H / 2 + 0.5
    ax.plot([x, x], [y_bot, y_top], color=CONTEXT_LINE, lw=1.0, zorder=2)
    for yy in (y_top, y_bot):
        ax.plot([x - tick, x], [yy, yy], color=CONTEXT_LINE, lw=1.0, zorder=2)
    ax.text(x + 1.4, (y_top + y_bot) / 2, "Same", rotation=90,
            ha="center", va="center", fontsize=17, color=TEXT_2)


def effectiveness_axis(ax):
    x = 126.0
    y_bot, y_top = row_y(9) - 3.0, row_y(0) + 3.0
    ax.plot([x, x], [y_bot, y_top], color=TEXT_2, lw=0.9, zorder=2)
    ax.plot([x], [y_top + 1.2], marker="^", ms=6, color=TEXT_2, zorder=2)
    ax.plot([x], [y_bot - 1.2], marker="v", ms=6, color=TEXT_2, zorder=2)
    for y, label in [(y_top - 7, "Best"),
                     ((y_top + y_bot) / 2, "Effectiveness"),
                     (y_bot + 8, "Least")]:
        ax.text(x - 2.2, y, label, rotation=90, ha="center", va="center",
                fontsize=19, color=TEXT_2)


def build(out_path):
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Aptos", "Calibri", "Segoe UI",
                                       "DejaVu Sans", "Arial"]

    xlim, ylim = (-2.0, 132.0), (-4.5, 100.0)
    width = 9.7
    height = width * (ylim[1] - ylim[0]) / (xlim[1] - xlim[0])

    fig, ax = plt.subplots(figsize=(width, height), dpi=200)
    fig.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.01)
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal")
    ax.axis("off")

    for i, (name, draw) in enumerate(zip(ROWS, DRAW)):
        y = row_y(i)
        ax.text(LABEL_X, y, name, ha="left", va="center",
                fontsize=22, color=INK)
        draw(ax, y)

    same_bracket(ax, 6, 7)   # luminance and saturation rank together
    same_bracket(ax, 8, 9)   # so do curvature and volume
    effectiveness_axis(ax)

    ax.text(LABEL_X, -1.5,
            "After Munzner, Visualization Analysis and Design (2014); "
            "ranking after Cleveland & McGill (1984)",
            ha="left", va="center", fontsize=14, color=CONTEXT_LINE)

    fig.savefig(out_path, facecolor=SURFACE)
    plt.close(fig)
    print(f"wrote {out_path}")


if __name__ == "__main__":
    build(Path(__file__).with_name("images") / "effectiveness.png")
