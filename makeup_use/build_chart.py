"""
Daily makeup use, 2024 vs 2019 — an overlapped paired bar chart.

Every category fell. That is the whole point of the chart, and it is why the
two years OVERLAP rather than sit as two separate bars in a cluster: the grey
2019 bar is drawn behind and the claret 2024 bar on top of it, so the drop is
the exposed grey tail on the right. A reader measures one gap per row instead
of comparing two bar lengths across a gutter.

The claret bar carries the number that matters and is labelled larger; the grey
one states where it came from, in the sliver it still owns.

Values are transcribed from a supplied chart image — read twice, once off the
printed labels and once by measuring each bar's pixel length, which agreed. The
originating survey is NOT known here: set SOURCE below before this goes in front
of anyone, or leave it empty and the figure simply carries no credit line rather
than a made-up one.

Run:  python build_chart.py               # images/makeup_use.png
      python build_chart.py --bg white
"""

import argparse
import csv
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent

CLARET = "#941D52"        # 2024 — sampled from the reference
GREY = "#828282"          # 2019
INK = "#1A1A1A"
LABEL = "#3C3C3C"

# Left empty on purpose. Naming a publication I have not verified would be
# worse than carrying no credit at all.
SOURCE = ""

TITLE_PT = 25
SUB_PT = 12
CAT_PT = 13
VAL_PT = 15               # the 2024 number
VAL_SMALL_PT = 11         # the 2019 number, in the sliver it still owns

BAR_H = 0.50              # height of each bar, in category units
OFFSET = 0.095            # how far each is nudged off the row centre
X_MAX = 80


def load():
    with open(HERE / "daily_makeup_use.csv", encoding="utf-8") as fh:
        return [(r["Type"], int(r["Pct2024"]), int(r["Pct2019"]))
                for r in csv.DictReader(fh)]


def draw_runs(fig, x, y, runs, size):
    """
    Lay coloured text fragments end to end, so the two years in the title can
    be the colour of the bars they name.

    matplotlib has no rich text, so each run is its own Text and the pen is
    advanced by what that run actually measured — guessing character widths
    puts visible gaps between the fragments at any size but one.
    """
    r = fig.canvas.get_renderer()
    for text, colour, weight in runs:
        t = fig.text(x, y, text, color=colour, fontsize=size, fontweight=weight,
                     ha="left", va="baseline", family="serif")
        x += t.get_window_extent(renderer=r).width / fig.bbox.width
    return x


def build(bg):
    plt.rcParams["font.family"] = "serif"
    plt.rcParams["font.serif"] = ["Georgia", "Times New Roman", "DejaVu Serif"]

    rows = load()
    n = len(rows)
    fig, ax = plt.subplots(figsize=(8.6, 7.4), dpi=200)
    fig.subplots_adjust(left=0.20, right=0.99, top=0.80, bottom=0.03)

    for i, (name, v24, v19) in enumerate(rows):
        y = n - 1 - i                      # first row at the top
        # 2019 behind, nudged up; 2024 in front, nudged down. Drawing order is
        # the whole trick — the drop reads as the grey tail the claret leaves.
        ax.barh(y + OFFSET, v19, height=BAR_H, color=GREY, zorder=2)
        ax.barh(y - OFFSET, v24, height=BAR_H, color=CLARET, zorder=3)

        ax.text(v24 - 1.2, y - OFFSET, f"{v24}%", ha="right", va="center",
                color="white", fontsize=VAL_PT, fontweight="bold", zorder=4)
        ax.text(v19 - 1.0, y + OFFSET + 0.16, f"{v19}%", ha="right", va="center",
                color="white", fontsize=VAL_SMALL_PT, fontweight="bold", zorder=4)
        ax.text(-1.5, y, name, ha="right", va="center",
                color=INK, fontsize=CAT_PT)

    ax.set_xlim(0, X_MAX)
    ax.set_ylim(-0.7, n - 0.3)
    ax.axis("off")

    # --- title, with each year in the colour of its own bars --------------
    x = draw_runs(fig, 0.055, 0.925, [
        ("Daily makeup use: ", INK, "normal"),
        ("2024", CLARET, "bold"),
        (" vs. ", INK, "normal"),
        ("2019", GREY, "bold"),
    ], TITLE_PT)

    fig.text(0.20, 0.855, "TYPE  |  % REPORTING DAILY USE", ha="left",
             va="baseline", color=LABEL, fontsize=SUB_PT, family="serif")

    if SOURCE:
        fig.text(0.055, 0.012, SOURCE, ha="left", va="bottom",
                 color=GREY, fontsize=9, family="serif")

    if bg != "none":
        fig.patch.set_facecolor(bg)
    return fig


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--out", default="makeup_use.png")
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
