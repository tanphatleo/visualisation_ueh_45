"""
The same table, the same library, two prompts -- and the anatomy of the good one.

Three figures for the AI section of the deck:

    chart_no_style.png   what you get from "make me a bar chart of this"
    chart_styled.png     what you get when the prompt carries the style guide
    chart_anatomy.png    the styled chart with every element named and specced

The two charts are DELIBERATELY built from the same numbers with the same
library. Nothing about the data changed between them -- only the instructions.
The exact prompts are in PROMPTS.md next to this file, and are quoted on the
slides, because the prompt is the lesson, not the picture.

chart_no_style is written the way a default matplotlib call comes out, on
purpose: rainbow bars, a legend of one, gridlines both ways, a full box, a
label for a title, rotated category names. Do not "fix" it.

    python build_charts.py
"""

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
OUT = HERE / "images"

# Revenue by channel, FY2026, billion VND. Marketplace beats retail and
# wholesale COMBINED (21.8 + 18.6 = 40.4) -- that is the takeaway the styled
# chart is built to make, and the reason marketplace is the one in colour.
DATA = [
    ("Marketplace", 42.3),
    ("Retail",      21.8),
    ("Wholesale",   18.6),
    ("Direct",      12.4),
    ("Export",       7.1),
    ("Affiliate",    3.2),
]

# style_chart.md 2.1 role colours
ACCENT = "#0F5499"
CONTEXT_FILL = "#BDBDBD"
INK = "#262A33"
TEXT_2 = "#595959"
CONTEXT_LINE = "#8C8C8C"

# style_chart.md 3, the type scale, in points
PT_TITLE, PT_SUB, PT_LABEL, PT_SOURCE = 20, 14, 11, 9

FONTS = ["Aptos", "Calibri", "Segoe UI", "DejaVu Sans", "Arial"]

TITLE = "Marketplace now earns more than retail and wholesale combined"
SUBTITLE = "Revenue by channel, FY2026 \u2014 billion VND"
SOURCE = "Source: finance data warehouse, FY2026 close. Revenue net of returns."


def no_style():
    """Defaults all the way down -- this is the control, not a strawman."""
    plt.style.use("default")
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["DejaVu Sans"]
    fig, ax = plt.subplots(figsize=(9.0, 5.4), dpi=170)
    names = [n for n, _ in DATA]
    vals = [v for _, v in DATA]
    # A different colour per bar: what you get when the tool colours by
    # category because nobody told it the categories are not a series.
    colours = plt.cm.tab10.colors[:len(names)]
    ax.bar(names, vals, color=colours, label="Revenue")
    ax.set_title("Revenue by Channel")
    ax.set_xlabel("Channel")
    ax.set_ylabel("Revenue")
    ax.grid(True)
    ax.set_axisbelow(True)
    ax.legend()
    plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
    fig.tight_layout()
    return fig


def styled_axes(fig, rect):
    """The chart block itself: sorted bars, one accent, labels on the bars."""
    ax = fig.add_axes(rect)
    rows = sorted(DATA, key=lambda r: r[1])          # barh draws bottom-up
    names = [n for n, _ in rows]
    vals = [v for _, v in rows]
    top = max(vals)
    colours = [ACCENT if v == top else CONTEXT_FILL for v in vals]

    ax.barh(names, vals, height=0.62, color=colours, zorder=3)
    for y, (name, v) in enumerate(rows):
        ax.text(v + top * 0.015, y, "%.1f" % v,
                va="center", ha="left", fontsize=PT_LABEL,
                color=ACCENT if v == top else TEXT_2,
                fontweight="semibold" if v == top else "normal")

    # style_chart.md 4: no gridlines on a bar chart that carries direct labels,
    # no value axis, no y-axis line, no tick marks, no box.
    ax.set_xlim(0, top * 1.12)
    ax.set_xticks([])
    for side in ("top", "right", "bottom", "left"):
        ax.spines[side].set_visible(False)
    ax.tick_params(axis="y", length=0, labelsize=PT_LABEL, colors=TEXT_2)
    return ax


def styled(annotated=False):
    plt.style.use("default")
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = FONTS

    if annotated:
        fig = plt.figure(figsize=(14.0, 6.6), dpi=170)
        rect, left = [0.115, 0.145, 0.46, 0.60], 0.045
    else:
        fig = plt.figure(figsize=(9.0, 5.4), dpi=170)
        rect, left = [0.155, 0.145, 0.795, 0.60], 0.055

    styled_axes(fig, rect)

    # Rule 1 and section 3: a sentence, then the description and the units.
    fig.text(left, 0.945, TITLE, fontsize=PT_TITLE, fontweight="semibold",
             color=INK, ha="left", va="top")
    fig.text(left, 0.845, SUBTITLE, fontsize=PT_SUB, color=TEXT_2,
             ha="left", va="top")
    fig.text(left, 0.055, SOURCE, fontsize=PT_SOURCE, color=CONTEXT_LINE,
             ha="left", va="bottom")
    return fig


# number, x, y (figure fractions), badge sits on a dark bar?, heading, body
NOTES = [
    (1, 0.030, 0.930, False, "Takeaway title",
     "A sentence carrying the point, not \u201cRevenue by channel\u201d.\n"
     "20 pt semibold, ink #262A33, left-aligned."),
    (2, 0.030, 0.838, False, "Subtitle",
     "What is measured, the period, the unit.\n14 pt regular, #595959."),
    (3, 0.030, 0.470, False, "Category labels",
     "11 pt, #595959, horizontal. Never rotated \u2014 if they do not fit,\n"
     "the chart turns on its side, as this one has."),
    (4, 0.220, 0.690, True, "One accent",
     "#0F5499 on the bar the title is about, #BDBDBD on every other.\n"
     "Colour marks the message and nothing else."),
    (5, 0.565, 0.690, False, "Data labels",
     "11 pt at the end of each bar, semibold on the accent.\n"
     "One decimal; the unit is stated once, in the subtitle."),
    (6, 0.350, 0.200, False, "No gridlines, no value axis, no box",
     "The labels already carry the numbers, so a scale up the side\n"
     "would only be a second copy of them."),
    (7, 0.030, 0.062, False, "Source line",
     "9 pt #8C8C8C at the foot. Where the numbers came from,\n"
     "and what they leave out."),
]


def badge(fig, n, x, y, on_dark):
    """
    A numbered dot, keyed to the note of the same number on the right.

    Drawn as a text bbox, not a Circle patch: a circle in FIGURE coordinates
    comes out an ellipse on any figure that is not square, and this one is
    more than twice as wide as it is tall.
    """
    face = "#FFFFFF" if on_dark else ACCENT
    fig.text(x, y, str(n), ha="center", va="center", zorder=11,
             fontsize=10, fontweight="bold",
             color=ACCENT if on_dark else "#FFFFFF",
             bbox=dict(boxstyle="circle,pad=0.34", facecolor=face,
                       edgecolor="none"))


def anatomy():
    fig = styled(annotated=True)
    # Badges sit on the chart; the notes stack down the right-hand column,
    # each opening with the same number.
    ny = 0.925
    for n, x, y, on_dark, head, body in NOTES:
        badge(fig, n, x, y, on_dark)
        badge(fig, n, 0.648, ny, False)
        fig.text(0.670, ny + 0.014, head, fontsize=11.5, fontweight="bold",
                 color=INK, ha="left", va="top")
        fig.text(0.670, ny - 0.013, body, fontsize=10.5, color=TEXT_2,
                 ha="left", va="top", linespacing=1.4)
        ny -= 0.132
    return fig


def main():
    OUT.mkdir(exist_ok=True)
    for name, fn in [("chart_no_style.png", no_style),
                     ("chart_styled.png", styled),
                     ("chart_anatomy.png", anatomy)]:
        fig = fn()
        fig.savefig(OUT / name, dpi=170, facecolor="white")
        plt.close(fig)
        print("wrote images/%s" % name)


if __name__ == "__main__":
    main()
