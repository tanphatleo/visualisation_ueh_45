"""
Write inventory_turnover.ipynb — read the CSV, redraw both charts in matplotlib.

The notebook's code cells live here as strings so they can be executed straight
out of this file before being written into the .ipynb:

    C:\\Python313\\python.exe build_notebook.py --run     # exec every code cell, then write
    C:\\Python313\\python.exe build_notebook.py           # just write the notebook

nbformat is not installed on this machine, so the .ipynb JSON is assembled by hand
(nbformat 4.5 — every cell needs an "id").
"""

import json
import sys
from pathlib import Path

HERE = Path(__file__).parent
NB = HERE / "inventory_turnover.ipynb"

MD_INTRO = """\
# Annual inventory turnover — one dataset, two renderings

`data/inventory_turnover.csv` holds six fiscal years of inventory turns for us and for the
industry benchmark. This notebook draws it twice:

1. **The replica** — a pixel-faithful rebuild of `image.png`: blue hero line, dashed black
   benchmark, no legend, both series labelled at their right-hand end.
2. **Warehouse Stand-Up** — the same numbers under `style_chart.md` v1.0: cream paper, one
   tangerine accent, chunky rounded line, punchline title, and a dashed event line on the year
   the story turns.

`build_xlsx.py` builds the Excel version of both — same data, same geometry, same colours.
The two renderings of a chart should never disagree.
"""

CODE_SETUP = """\
from pathlib import Path

import matplotlib as mpl
import matplotlib.pyplot as plt
import pandas as pd

HERE = Path.cwd()                      # run the notebook from its own folder
CSV = HERE / "data" / "inventory_turnover.csv"
IMAGES = HERE / "images"
IMAGES.mkdir(exist_ok=True)

df = pd.read_csv(CSV)
years = df["fiscal_year"]
ours = df["our_company"]
bench = df["industry_benchmark"]
df
"""

MD_REPLICA = """\
## 1. The replica

Every number below was measured out of `image.png`: the y-axis sits at x = 58 px, the baseline at
y = 252 px, the "10" tick at y = 85 px, and 2025 at x = 308.5 px on a 402 x 301 canvas. Pinning the
axes to those fractions at 100 dpi reproduces the original's geometry, not just its style.

The blue is `#357DBF`, sampled from the PNG.
"""

CODE_REPLICA = """\
BLUE, DARK, TICK, AXIS = "#357DBF", "#404040", "#595959", "#B5B5B5"
FONT = "Segoe UI"

fig = plt.figure(figsize=(4.02, 3.01), dpi=100, facecolor="white")
# left, bottom, width, height as fractions of the 402 x 301 canvas
ax = fig.add_axes([58 / 402, 1 - 252 / 301, (308.5 - 58) / 402, (252 - 85) / 301],
                  facecolor="white")

ax.plot(years, ours, color=BLUE, lw=2.2, solid_joinstyle="round", zorder=3)
ax.plot(years, bench, color=DARK, lw=1.2, ls=(0, (4, 3)), zorder=2)

# 2020 sits ON the axis, 2025 on the right edge: the "on tick marks" placement
ax.set_xlim(2020, 2025)
ax.set_ylim(0, 10)
ax.set_xticks(years)
ax.set_yticks(range(0, 11, 2))
for side in ("top", "right"):
    ax.spines[side].set_visible(False)
for side in ("left", "bottom"):
    ax.spines[side].set_color(AXIS)
    ax.spines[side].set_linewidth(0.9)
ax.tick_params(colors=AXIS, labelcolor=TICK, labelsize=8, length=3, width=0.9, direction="out")
for label in ax.get_xticklabels() + ax.get_yticklabels():
    label.set_fontfamily(FONT)

ax.set_ylabel("# OF INVENTORY TURNS", fontsize=6.5, color=DARK, fontfamily=FONT, labelpad=6)
fig.text(58 / 402, 0.04, "FISCAL YEAR ENDING 12/31", fontsize=6.5, color=DARK, fontfamily=FONT)
fig.text(0.058, 0.915, "Annual inventory turnover", fontsize=12, color=DARK,
         fontfamily=FONT, va="center")

# direct labels instead of a legend — the words sit on the line they describe
ax.annotate("OUR COMPANY", xy=(2025, ours.iloc[-1]), xytext=(7, 0),
            textcoords="offset points", color=BLUE, fontsize=7.5, fontweight="bold",
            fontfamily=FONT, va="center")
ax.annotate("INDUSTRY\\nBENCHMARK", xy=(2025, bench.iloc[-1]), xytext=(7, 0),
            textcoords="offset points", color=DARK, fontsize=7.5,
            fontfamily=FONT, va="center", linespacing=1.4)

fig.savefig(IMAGES / "replica_matplotlib.png", dpi=100, facecolor="white")
plt.show()
"""

MD_STYLE = """\
## 2. Warehouse Stand-Up — `style_chart.md` v1.0

The style guide in this folder, applied literally:

* **§2** cream `#FFF9F0` paper, one tangerine `#EF5B23` hero, warm grey `#6E6259` benchmark.
  Tangerine is 3.2:1 on cream — legal for a line, illegal for text — so the end label is written in
  `#C2400F` instead.
* **§3** Verdana Bold for the punchline, Segoe UI for everything else. No Comic Sans: the chart is
  the comedian, the font is the straight man.
* **§4** four bands — punchline, straight-man subtitle, plot, source — and 24 % of the width kept
  free on the right for the end labels.
* **§5** 3 pt line, round caps, hollow round markers so the reader can count the six years.
* **§7** the story has a cause, so it gets a vertical dashed event line with the cause written on
  it. In Excel that is a scatter series on a hidden 0–1 axis; here it is `axvline` + `text`, same
  rules: dashed, alert colour, caption at the top.
* **§1** one emoji, one wisecrack, both spent in that caption.
"""

CODE_STYLE = """\
# --- style_chart.md §8, pasted verbatim -----------------------------------
STYLE = {
    "figure.facecolor": "#FFF9F0", "axes.facecolor": "#FFF9F0",
    "figure.figsize": (8.0, 6.0), "figure.dpi": 150,
    "font.family": ["Segoe UI", "Calibri", "DejaVu Sans"], "font.size": 9,
    "text.color": "#2F2A26",
    "axes.edgecolor": "#E7D8C3", "axes.linewidth": 1.0,
    "axes.labelcolor": "#5C5149", "axes.labelsize": 8,
    "axes.titlelocation": "left", "axes.titlesize": 13, "axes.titleweight": "bold",
    "axes.titlecolor": "#2F2A26",
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "axes.grid.axis": "y",
    "grid.color": "#E7D8C3", "grid.linewidth": 0.9,
    "xtick.color": "#5C5149", "ytick.color": "#5C5149",
    "xtick.labelsize": 9, "ytick.labelsize": 9,
    "xtick.direction": "out", "ytick.direction": "out",
    "lines.linewidth": 3.0, "lines.solid_capstyle": "round",
    "lines.dash_capstyle": "round", "lines.markersize": 7,
    "legend.frameon": False,
    "savefig.facecolor": "#FFF9F0", "savefig.bbox": "standard",
}
mpl.rcParams.update(STYLE)

SURFACE, WASH, GRID = "#FFF9F0", "#FBEFDD", "#E7D8C3"
CONTEXT, TEXT_2, INK = "#6E6259", "#5C5149", "#2F2A26"
ACCENT, ACCENT_TEXT, GOOD, BAD = "#EF5B23", "#C2400F", "#0B7A72", "#B8352C"
HEAD, BODY = "Verdana", "Segoe UI"
EMOJI = [BODY, "Segoe UI Emoji"]       # font fallback, or the emoji renders as a box

PUNCHLINE = ("Four years of losing to the Joneses —\\n"
             "then someone found the warehouse keys.")
SUBTITLE = ("Inventory turns per year, fiscal years ending 12/31.\\n"
            "Higher is better; the dashed line is the industry benchmark.")
SOURCE = ("Source: data/inventory_turnover.csv  |  Turns = COGS \\u00f7 average inventory. "
          "No inventory was harmed in the making of this chart.")
EVENT_YEAR, EVENT_TEXT = 2024, "New WMS goes live \\U0001F389"
"""

CODE_STANDUP = """\
fig, ax = plt.subplots(figsize=(8.0, 6.0), dpi=150)
# band fractions straight out of style_chart.md §4: title 14 %, subtitle 8 %,
# visual 72 % (plot + tick labels + axis title), source 6 %
fig.subplots_adjust(left=0.10, right=0.76, top=0.755, bottom=0.155)

ax.plot(years, ours, color=ACCENT, lw=3.0, marker="o", ms=7, mfc=SURFACE, mec=ACCENT,
        mew=2.0, solid_capstyle="round", solid_joinstyle="round", zorder=4)
ax.plot(years, bench, color=CONTEXT, lw=1.5, ls=(0, (5, 3)), zorder=3)

ax.set_xlim(2020, 2025)
ax.set_ylim(0, 10)
ax.set_xticks(years)
ax.set_yticks(range(0, 11, 2))
ax.set_ylabel("# OF INVENTORY TURNS", labelpad=8)
ax.set_xlabel("FISCAL YEAR ENDING 12/31", loc="left", labelpad=10)
ax.tick_params(length=3, width=1.0, color=GRID)
ax.set_axisbelow(True)

# §7 — the event line, built at the period it explains, caption riding on top
ax.axvline(EVENT_YEAR, ymin=0, ymax=0.95, color=BAD, lw=1.0, ls=(0, (4, 3)), zorder=2)
ax.text(EVENT_YEAR, 9.65, EVENT_TEXT, color=BAD, fontsize=9, style="italic",
        ha="center", va="bottom", fontfamily=EMOJI)

# §11 — no legend; every series labelled at its end
ax.annotate("US", xy=(2025, ours.iloc[-1]), xytext=(10, 0), textcoords="offset points",
            color=ACCENT_TEXT, fontsize=9, fontweight="bold", va="center")
ax.annotate("THE JONESES", xy=(2025, bench.iloc[-1]), xytext=(10, 0),
            textcoords="offset points", color=CONTEXT, fontsize=9, fontweight="bold",
            va="center")

# §4 — the four bands: punchline, straight man, plot, source
fig.text(0.02, 0.985, PUNCHLINE, fontsize=16, fontweight="bold", fontfamily=HEAD,
         color=INK, va="top", linespacing=1.35)
fig.text(0.02, 0.855, SUBTITLE, fontsize=11, color=TEXT_2, va="top", linespacing=1.35)
fig.text(0.02, 0.03, SOURCE, fontsize=9, color=TEXT_2, va="bottom")

fig.savefig(IMAGES / "standup_matplotlib.png", dpi=150, facecolor=SURFACE)
plt.show()
"""

MD_OUTRO = """\
## 3. The Excel twin

```
C:\\Python313\\python.exe build_xlsx.py
```

writes `inventory_turnover.xlsx` with the same two charts as native Excel line charts — sheet
`replica` and sheet `standup`. Both read the table on their sheet rather than baked-in numbers, so
editing a cell moves the chart. The event line there is a two-point scatter series on a hidden
0–1 secondary axis, which is the Excel equivalent of the `axvline` above.

Checklist before calling either version done: `style_chart.md` §11.
"""

CELLS = [
    ("markdown", MD_INTRO),
    ("code", CODE_SETUP),
    ("markdown", MD_REPLICA),
    ("code", CODE_REPLICA),
    ("markdown", MD_STYLE),
    ("code", CODE_STYLE),
    ("code", CODE_STANDUP),
    ("markdown", MD_OUTRO),
]


def as_lines(src):
    """Notebook JSON stores source as a list of lines, each keeping its newline."""
    lines = src.splitlines(keepends=True)
    return lines


def notebook():
    cells = []
    for i, (kind, src) in enumerate(CELLS, start=1):
        cell = {"cell_type": kind, "id": f"cell{i:02d}", "metadata": {},
                "source": as_lines(src)}
        if kind == "code":
            cell["execution_count"] = None
            cell["outputs"] = []
        cells.append(cell)
    return {
        "cells": cells,
        "metadata": {
            "kernelspec": {"display_name": "Python 3", "language": "python",
                           "name": "python3"},
            "language_info": {"name": "python", "version": "3.13"},
        },
        "nbformat": 4,
        "nbformat_minor": 5,
    }


def run_cells():
    """Execute the code cells in order, in one shared namespace, from HERE."""
    import os
    os.chdir(HERE)
    import matplotlib
    matplotlib.use("Agg")
    ns = {"__name__": "__notebook__"}
    for kind, src in CELLS:
        if kind == "code":
            exec(compile(src, "<cell>", "exec"), ns)
    print("all code cells ran")


if __name__ == "__main__":
    if "--run" in sys.argv:
        run_cells()
    NB.write_text(json.dumps(notebook(), indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"wrote {NB}")
