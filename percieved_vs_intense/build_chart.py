"""
Stevens's power law as a deck figure: perceived sensation against intensity.

The same six curves as perceived_vs_intensity.xlsx — psi = I^a — drawn for the
slide rather than for reading values off. Which is why there are no gridlines
and no legend: the chart is read for the SHAPE of each curve and for where it
sits relative to the diagonal, and every curve carries its own name.

The exponent is the whole story:

    a > 1   sensation outruns the stimulus (shock, saturation)
    a = 1   veridical — what you encode is what the reader perceives (length)
    a < 1   sensation is compressed (area, depth, brightness)

Labels ride along their curves at the curve's own on-screen angle, which is what
lets six of them sit in a crowded plot without a legend and without collisions.

Run:  python build_chart.py                # images/perceived_vs_intensity.png
      python build_chart.py --bg white
"""

import argparse
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

HERE = Path(__file__).resolve().parent

INK = "#262A33"
TEXT_2 = "#595959"
CONTEXT_LINE = "#8C8C8C"

X_MAX, Y_MAX = 5.0, 5.0

# name, exponent, colour, x to label at, perpendicular offset in points.
#
# Five exponents are Stevens's own (Stevens 1957, Table 1). DEPTH IS NOT:
# 0.67 is his LOUDNESS exponent and his table has no depth entry. It is kept
# because depth's rank is well established in the visualization literature —
# and the figure's own caption says so rather than implying Stevens.
CHANNELS = [
    # name              a      colour      label x   offset
    # The two steep curves are all but vertical where they are labelled, so
    # their perpendicular offset is essentially a horizontal one: shrinking the
    # magnitude walks the label back towards its own line, to the LEFT.
    ("Electric Shock", 3.5,  "#1B9AD6",      1.42,   -24),
    ("Saturation",     1.7,  "#2E9E4F",      2.15,   -20),
    ("Length",         1.0,  "#7B3F9E",      4.30,    16),
    ("Area",           0.7,  "#C0392B",      4.25,    18),
    ("Depth",          0.67, "#0D7680",      4.25,   -18),
    ("Brightness",     0.5,  "#8C8C8C",      3.95,   -19),
]

LINE_W = 3.8
LABEL_PT = 23
AXIS_PT = 25
TICK_PT = 21
CITE_PT = 13
CITE_Y = 0.048        # top of the credit block, in figure fractions
CITE_LEAD = 1.35      # line pitch as a multiple of CITE_PT


def screen_angle(ax, x, a, dx=1e-3):
    """
    The curve's slope in DEGREES ON SCREEN at x, not in data units.

    A label has to follow what the reader sees, and the two differ whenever the
    axes are not equally scaled — which here they are not, and the steep curves
    would be labelled at visibly the wrong angle if this used dy/dx directly.
    """
    x0, x1 = max(x - dx, 1e-9), x + dx
    (px0, py0), (px1, py1) = ax.transData.transform([(x0, x0 ** a), (x1, x1 ** a)])
    return np.degrees(np.arctan2(py1 - py0, px1 - px0))


def label_on_curve(ax, x, a, text, colour, offset):
    """
    Put the series' name on its own line, rotated to match and nudged clear.

    The offset is PERPENDICULAR to the curve, so a label sits the same visual
    distance off its line whether that line is near-flat or near-vertical — a
    plain dy offset would drift right across a steep curve and land on it.
    """
    ang = screen_angle(ax, x, a)
    rad = np.radians(ang)
    dx, dy = -np.sin(rad) * offset, np.cos(rad) * offset
    ax.annotate(text, xy=(x, x ** a), xytext=(dx, dy),
                textcoords="offset points", rotation=ang, rotation_mode="anchor",
                ha="center", va="center", fontsize=LABEL_PT,
                color=colour, fontweight="bold", clip_on=False)


def build(bg):
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Aptos", "Calibri", "Segoe UI",
                                       "DejaVu Sans", "Arial"]

    fig, ax = plt.subplots(figsize=(8.6, 9.0), dpi=200)
    fig.subplots_adjust(left=0.122, right=0.979, top=0.983, bottom=0.164)

    x = np.linspace(0, X_MAX, 600)
    for name, a, colour, _, _ in CHANNELS:
        ax.plot(x, x ** a, color=colour, lw=LINE_W,
                solid_capstyle="round", zorder=3)

    ax.set_xlim(0, X_MAX)
    ax.set_ylim(0, Y_MAX)          # the steep curves simply run off the top,
    ax.set_xticks(range(0, 6))     # exactly as they do in the reference
    ax.set_yticks(range(0, 6))
    # Belt and braces: the figure is already sized to give a square box, and
    # this pins one data unit to one display unit so Length stays at 45 even
    # if the margins are retuned later.
    ax.set_aspect("equal", adjustable="box")
    ax.tick_params(labelsize=TICK_PT, colors=TEXT_2, length=4, width=1.0)
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_color(TEXT_2)
        ax.spines[side].set_linewidth(1.0)

    ax.set_xlabel("Physical intensity", fontsize=AXIS_PT, color=INK, labelpad=10)
    ax.set_ylabel("Perceived sensation", fontsize=AXIS_PT, color=INK, labelpad=10)

    # Labels are placed last: screen_angle needs the axes limits already set,
    # or every rotation is computed against matplotlib's provisional scaling.
    for name, a, colour, lx, off in CHANNELS:
        label_on_curve(ax, lx, a, f"{name} ({a:g})", colour, off)

    # One text block, not two: the caveat belongs to the credit above it, so the
    # gap between them is a leading multiple that stays put when CITE_PT or the
    # block position changes. As two separate fig.text calls the spacing was a
    # difference between two y-values, and drifted every time either moved.
    fig.text(0.122, CITE_Y,
             "Exponents after Stevens, “On the psychophysical law”, "
             "Psychological Review 64(3), 1957.\n"
             "Depth (0.67) is not in Stevens’s table — it is carried "
             "from the visualization literature.",
             ha="left", va="top", linespacing=CITE_LEAD,
             fontsize=CITE_PT, color=CONTEXT_LINE)

    if bg != "none":
        fig.patch.set_facecolor(bg)
    return fig


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--out", default="perceived_vs_intensity.png")
    ap.add_argument("--bg", default="none", help="'none' for transparent, or a colour")
    ap.add_argument("--outdir", default="images")
    a = ap.parse_args()

    outdir = HERE / a.outdir
    outdir.mkdir(exist_ok=True)
    fig = build(a.bg)
    fig.savefig(outdir / a.out, dpi=200, transparent=(a.bg == "none"))
    plt.close(fig)
    print(f"wrote {a.outdir}/{a.out}")


if __name__ == "__main__":
    main()
