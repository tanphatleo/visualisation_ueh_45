"""
Write playfair_trade_balance.ipynb, and optionally execute every cell first.

    python build_notebook.py          # write the notebook
    python build_notebook.py --run    # exec every code cell here, then write it

nbformat is not installed in this environment, so the notebook JSON is authored by
hand (nbformat 4.5 — every cell needs an `id`). `--run` is the smoke test: it runs
the cells in one namespace so a broken figure fails here rather than in Jupyter.
"""
import json
import sys
import uuid
from pathlib import Path

HERE = Path(__file__).resolve().parent
OUT = HERE / "playfair_trade_balance.ipynb"

# ---------------------------------------------------------------------------
MD_INTRO = """\
# Playfair's trade balance, 1700–1780

In 1786 William Playfair published *The Commercial and Political Atlas* in London. It
contains the first time-series charts ever printed — 43 of them — and one bar chart, and
it is where statistical graphics starts. The plate rebuilt here is the most quoted of the
set: **Exports and Imports to and from DENMARK & NORWAY from 1700 to 1780**.

Two curves, eighty-one years, and one idea that had never been drawn before: **shade the
gap between them and the trade balance becomes a shape you can see**. Playfair labelled
the two shapes *BALANCE AGAINST* and *BALANCE in FAVOUR of ENGLAND*, which is a takeaway
title two centuries early.

*Source:* Jeremy Norman, *History of Information*, entry 2929 —
<https://www.historyofinformation.com/detail.php?entryid=2929>

## Where the numbers come from

Playfair published the plate, not the table. The CSV in `data/` was **traced from the
engraving**: each curve was isolated by its ink colour (ochre for imports, crimson for
exports), sampled at every year, and converted to values through the plate's own
gridlines. `trace_from_image.py` does it and writes a verification overlay to
`reference/trace_verification.png`.

So treat the numbers as **± 1 (thousand pounds)**. They are a faithful reading of the
engraving, not Playfair's ledger figures.

## Three charts from one table

| Figure | What it is | Style |
| --- | --- | --- |
| 1 | A replica of the 1786 plate | Playfair's, deliberately not the house style |
| 2 | Exports vs imports, one accent | `style/style_chart.md` v1.0 |
| 3 | The balance itself, by favourability | `style/style_chart.md` v1.0 |

Figure 1 exists to prove the data is right. Figures 2 and 3 are what you would actually
put in front of someone today.
"""

CELL_SETUP = '''\
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

HERE = Path.cwd()
sys.path.insert(0, str(HERE / "style"))

import finance_style as fs          # noqa: E402  - needs the path above

fs.apply()                          # loads style/finance.mplstyle

# Wrapped by hand into three lines: at 9 pt a one-line source runs off a 9-inch
# figure, and savefig.bbox is "standard", not "tight", so nothing rescues it.
SOURCE = ("Source: William Playfair, The Commercial and Political Atlas, London 1786, plate\\n"
          "'Exports and Imports to and from Denmark & Norway from 1700 to 1780'"
          " · via historyofinformation.com entry 2929\\n"
          "Values traced from the engraving to ±1, not transcribed from ledgers"
          " · Traced 18 Aug 2026")
print(fs.ACCENT, fs.CONTEXT_LINE, fs.GOOD, fs.BAD)
'''

CELL_READ = '''\
df = pd.read_csv(HERE / "data" / "playfair_denmark_norway.csv")
df["balance"] = df["exports"] - df["imports"]          # recomputed, never trusted

year = df["year"].to_numpy()
imports = df["imports"].to_numpy()
exports = df["exports"].to_numpy()
balance = df["balance"].to_numpy()

# The year the balance turns: the first year exports go ahead AND stay ahead. Not
# simply the first positive year - a one-off crossing is not a turning point.
turn = int(next(y for i, y in enumerate(year) if (balance[i:] > 0).all()))
print(f"{len(df)} years, {year[0]}-{year[-1]}  ·  balance turns in {turn}")
df.head()
'''

MD_FIG1 = """\
## Figure 1 — the replica

Everything here breaks `style_chart.md` on purpose: a boxed plot, gridlines both ways,
nineteen labelled ticks on the **right**-hand axis, two saturated fills, a heavy rule at
£100,000. That is the artefact, and the point of drawing it is that the traced numbers
have to land on Playfair's own curves for the replica to look right.

Colours are measured off the engraving — the median of the ink pixels in each stroke and
each band — not guessed.
"""

CELL_FIG1 = '''\
# Colours sampled from the engraving itself (median of the stroke / fill pixels).
PF_IMPORT, PF_EXPORT = "#D48D50", "#B05058"       # ochre and crimson lines
PF_AGAINST, PF_FAVOUR = "#F7DAD7", "#E7D9B5"      # pink and buff bands
PF_RULE, PF_FRAME = "#8A837C", "#4A443E"

fig, ax = plt.subplots(figsize=(11.0, 7.2))

# fill_between is what Excel cannot do: shade between two curves directly. In the
# workbook the same two bands cost three stacked-area series and a hidden base.
# interpolate=True makes the band close exactly at the crossing, not at the next year.
ax.fill_between(year, imports, exports, where=imports >= exports,
                color=PF_AGAINST, interpolate=True)
ax.fill_between(year, imports, exports, where=exports >= imports,
                color=PF_FAVOUR, interpolate=True)
ax.axhline(100, color=PF_FRAME, linewidth=1.4)     # Playfair's heavy 100,000 rule
ax.plot(year, imports, color=PF_IMPORT, linewidth=2.4)
ax.plot(year, exports, color=PF_EXPORT, linewidth=2.4)

ax.set_xlim(1700, 1780)
ax.set_ylim(0, 200)
ax.set_xticks(range(1700, 1781, 10))
ax.set_yticks(range(0, 201, 10))                   # "the Right hand line into L10,000 each"

# The house style strips all of this. The replica puts it back, one line each, so the
# diff between the two figures is legible rather than buried in rcParams.
ax.grid(True, axis="both", color=PF_RULE, linewidth=0.6)
ax.set_axisbelow(True)
for side in ("top", "right", "bottom", "left"):
    ax.spines[side].set_visible(True)
    ax.spines[side].set_color(PF_FRAME)
    ax.spines[side].set_linewidth(1.6)
ax.yaxis.tick_right()
ax.tick_params(axis="both", length=0, colors=PF_FRAME, labelsize=8)
ax.set_title("Exports and Imports to and from DENMARK & NORWAY from 1700 to 1780",
             fontsize=13, color=PF_FRAME, loc="center", weight="normal", pad=14)

# Playfair's own four captions, in his own words.
style = dict(color=PF_FRAME, style="italic", ha="center", va="center")
ax.text(1737, 76, "BALANCE AGAINST", weight="bold", fontsize=12, **style)
ax.text(1768, 128, "BALANCE in FAVOUR of\\nENGLAND", weight="bold", fontsize=12, **style)
ax.text(1721, 106, "Line of Imports", fontsize=9, **style)
ax.text(1731, 68, "Line of Exports", fontsize=9, **style)
ax.text(1777, 103, "100,000", fontsize=9, **style)

fig.text(0.06, 0.015, "The Bottom line is divided into Years, the Right hand line into "
                      "L10,000 each. — W. Playfair, 1786",
         size=9, color=PF_FRAME, style="italic", ha="left", va="bottom")
fig.subplots_adjust(left=0.04, right=0.955, top=0.92, bottom=0.10)
fig.savefig(HERE / "reference" / "fig1_replica.png", dpi=200)
plt.show()
'''

MD_FIG2 = """\
## Figure 2 — the same data, house style

`style_chart.md` v1.0, applied literally:

* **Takeaway title**, not a label of the contents (§1 rule 1).
* Grey is the default; **one accent** on the series the title is about (§1 rule 2, §2.3).
* **Direct labels** at each line's own end — no legend (§1 rule 5).
* Horizontal gridlines only, no y-axis line, no tick marks, no frame (§4).
* The value axis runs from **0**, so the slopes are honest (§1 rule 4).
* An **event line** on the year the balance turns, with the cause written on the chart
  (non-negotiable 11). In matplotlib that is `axvline` plus `ax.text`; in the workbook it
  is a two-point scatter series on a hidden 0–1 axis, because a drawn shape does not move
  when the data does.

Units go in the axis title and the numbers stay bare (§5). Currency is stated as
*pounds sterling* rather than the guide's `₫` for the obvious reason — noted in
`README.md` under deviations.
"""

CELL_FIG2 = '''\
fig, ax = fs.new_chart(
    title=f"Exports overtook imports in {turn} and never fell back",
    subtitle="Annual exports and imports, thousands of pounds sterling, 1700-1780",
    source=SOURCE,
    right=0.855,            # room outside the plot for the two direct labels
)

ax.plot(year, imports, color=fs.CONTEXT_LINE, linewidth=fs.LINE_CONTEXT)
ax.plot(year, exports, color=fs.ACCENT, linewidth=fs.LINE_HIGHLIGHT)

# 1698, not 1700: with the first category ON the axis the "1700" tick label lands on
# top of the "0" value label. checks/check_chart_text.py catches it; the eye catches it
# faster. Two years of inset moves the tick clear and costs nothing.
ax.set_xlim(1698, 1780)
ax.set_ylim(0, 200)
ax.set_xticks(range(1700, 1781, 10))
ax.set_yticks(range(0, 201, 50))
ax.tick_params(axis="x", pad=5)
ax.grid(axis="y", color=fs.GRID, linewidth=0.75)
ax.set_axisbelow(True)
for side in ("top", "right", "left"):
    ax.spines[side].set_visible(False)
ax.spines["bottom"].set_color(fs.TEXT_2)
ax.spines["bottom"].set_linewidth(0.75)
ax.tick_params(axis="both", length=0, colors=fs.TEXT_2, labelsize=fs.SZ_AXIS_LABEL)
# No y-axis title: the subtitle already states the unit, and the guide wants it stated
# ONCE. In the workbook it goes the other way round - an Excel chart object has no
# subtitle, so there the axis title is the only place the unit can live.

# Direct labels: name and value at the series' own last point, in the series' colour.
for label, values, colour in (("Exports", exports, fs.ACCENT),
                              ("Imports", imports, fs.CONTEXT_LINE)):
    ax.text(1781.5, values[-1], f"{label}, {fs.vn(values[-1])}",
            color=colour, size=fs.SZ_DATA_LABEL, weight="semibold", va="center")

# The event line, and the reason the line bends written next to it.
ax.axvline(turn, color=fs.ALERT, linewidth=0.75, linestyle=(0, (4, 3)))
ax.text(turn - 1.5, 178, f"Balance turns in\\nEngland's favour, {turn}",
        color=fs.ALERT, size=fs.SZ_ANNOTATION, ha="right", va="top", linespacing=1.3)

fig.savefig(HERE / "reference" / "fig2_house_style.png", dpi=200)
plt.show()
'''

MD_FIG3 = """\
## Figure 3 — the balance on its own

Figure 2 shows two quantities and asks the reader to subtract them by eye. This one plots
the subtraction: **exports minus imports**, one column a year, centred on zero.

Two house rules do the work here:

* **Colour by favourability, not by direction** (§1 rule 10). Teal is a balance in
  England's favour, claret a balance against — Playfair's own two categories.
* Teal and claret sit only 12.3 L\\* apart, so in greyscale they are near-identical.
  Direction therefore never depends on colour alone: **every label carries its sign**
  (§2.6). The midpoint is stated in the subtitle.

Fifty-five years of deficit, then twenty-five of surplus, and the turn is a single
crossing rather than a wobble.
"""

CELL_FIG3 = '''\
def gbp_k(value):
    """Round to whole thousands - the traced data is only good to +/- 1 - then vi-VN."""
    return "£" + fs.vn(round(value) * 1000)


fig, ax = fs.new_chart(
    title=(f"A {gbp_k(-balance[0])} deficit in {year[0]} became a "
           f"{gbp_k(balance[-1])} surplus in {year[-1]}"),
    # Wrapped: one line of this at 14 pt overruns a 9-inch figure.
    subtitle=("Exports minus imports, thousands of pounds sterling.\\n"
              "Midpoint is zero: teal is a balance in England's favour, claret against."),
    source=SOURCE,
    left=0.075, right=0.975,
)

colours = [fs.GOOD if b >= 0 else fs.BAD for b in balance]
ax.bar(year, balance, width=fs.BAR_WIDTH, color=colours, linewidth=0)

ax.set_xlim(1699, 1781)
ax.set_ylim(-62, 118)
ax.set_xticks(range(1700, 1781, 10))
ax.set_yticks(range(-50, 101, 50))
ax.grid(axis="y", color=fs.GRID, linewidth=0.75)
ax.set_axisbelow(True)
# The zero rule is drawn as a line, NOT by moving the bottom spine to y=0: moving the
# spine drags the year labels up onto the rule with it. This is the same problem Excel
# solves with tickLblPos = "low".
for side in ("top", "right", "left", "bottom"):
    ax.spines[side].set_visible(False)
ax.axhline(0, color=fs.TEXT_2, linewidth=1.0, zorder=3)
ax.tick_params(axis="both", length=0, colors=fs.TEXT_2, labelsize=fs.SZ_AXIS_LABEL)
# Unit is in the subtitle, so no axis title here either.

# Signed labels on three points only: first, the turn, and last. A sign is what makes
# this readable in greyscale.
for x in (year[0], turn, year[-1]):
    i = int((year == x).nonzero()[0][0])
    b = balance[i]
    # The 1700 label is centred on the leftmost bar, which puts it on top of the
    # "-50" tick label. Left-align that one so it runs into the plot instead.
    edge = x == year[0]
    ax.text(x - 0.4 if edge else x, b + (5 if b >= 0 else -5),
            f"{'+' if b >= 0 else chr(8722)}{fs.vn(abs(b))}",
            color=fs.GOOD if b >= 0 else fs.BAD, size=fs.SZ_DATA_LABEL,
            weight="semibold", ha="left" if edge else "center",
            va="bottom" if b >= 0 else "top")

fig.savefig(HERE / "reference" / "fig3_deviation.png", dpi=200)
plt.show()
'''

MD_CLOSE = """\
## What this case is for

**The technique.** `fill_between(..., where=..., interpolate=True)` shades between two
curves in one line. Excel has no such thing: the same two bands in
`playfair_trade_balance.xlsx` cost three stacked-area series — an invisible base at
`MIN(imports, exports)`, then the deficit gap, then the surplus gap — with the two real
lines drawn on top as a combo line chart. Worth knowing before promising a client a
shaded band in a workbook.

**The history.** Playfair had no convention to follow, and he still landed on a takeaway
title (*BALANCE in FAVOUR of ENGLAND*), direct labels on the curves (*Line of Imports*), a
stated unit (*the Right hand line into L10,000 each*) and a reference line at a round
number. Four of the rules in `style/style_chart.md` are in a chart from 1786.

**Where he would fail a review today.** Nineteen labelled gridlines and a full box are
noise by any modern standard; the two saturated fills give the reader no place to rest;
and there is no source line. Figure 2 is the same argument with 90% of the ink removed —
compare them side by side, then decide which one you would present.

**And a caution.** The numbers here were traced off an engraving. That is a legitimate way
to recover a historical series and a terrible way to report a current one. Every artefact
in this folder says so, in the source line, on the chart.
"""


# ---------------------------------------------------------------------------
def md(text):
    return {"cell_type": "markdown", "id": str(uuid.uuid4())[:8],
            "metadata": {}, "source": text.splitlines(keepends=True)}


def code(text):
    return {"cell_type": "code", "id": str(uuid.uuid4())[:8], "metadata": {},
            "execution_count": None, "outputs": [],
            "source": text.rstrip("\n").splitlines(keepends=True)}


CODE_CELLS = [CELL_SETUP, CELL_READ, CELL_FIG1, CELL_FIG2, CELL_FIG3]

CELLS = [md(MD_INTRO), code(CELL_SETUP), code(CELL_READ),
         md(MD_FIG1), code(CELL_FIG1),
         md(MD_FIG2), code(CELL_FIG2),
         md(MD_FIG3), code(CELL_FIG3),
         md(MD_CLOSE)]


def run_cells():
    """Execute every code cell in one namespace — the notebook's smoke test."""
    import matplotlib
    matplotlib.use("Agg")
    ns = {"__name__": "__main__"}
    for i, src in enumerate(CODE_CELLS, 1):
        exec(compile(src, f"<cell {i}>", "exec"), ns)
        print(f"cell {i} ok")
    return ns


def main():
    if "--run" in sys.argv:
        import os
        os.chdir(HERE)
        run_cells()
    nb = {"cells": CELLS,
          "metadata": {"kernelspec": {"display_name": "Python 3", "language": "python",
                                      "name": "python3"},
                       "language_info": {"name": "python", "version": "3.13"}},
          "nbformat": 4, "nbformat_minor": 5}
    OUT.write_text(json.dumps(nb, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
