# Mini case — Playfair's trade balance, 1700–1780

Rebuild of the chart that invented the shaded time series: **"Exports and Imports to and
from DENMARK & NORWAY from 1700 to 1780"**, plate from William Playfair's *The Commercial
and Political Atlas*, London 1786.

Brief: *History of Information*, entry 2929 —
<https://www.historyofinformation.com/detail.php?entryid=2929>

Same data three ways — Playfair's own layout, the house style, and the balance on its own —
in a workbook and a notebook, from one CSV.

---

## Files

| Path | What it is |
| --- | --- |
| `data/playfair_denmark_norway.csv` | 81 rows: `year, imports, exports, balance`, thousands of pounds |
| `trace_from_image.py` | recovers the CSV from the engraving; `--write` rewrites it |
| `build_xlsx.py` | builds `playfair_trade_balance.xlsx` — 4 sheets, 3 charts |
| `build_notebook.py` | writes `playfair_trade_balance.ipynb`; `--run` executes every cell first |
| `playfair_trade_balance.xlsx` | the workbook (Excel-verified, 0 collisions) |
| `playfair_trade_balance.ipynb` | the notebook, three matplotlib figures |
| `reference/playfair_original.png` | the source scan |
| `reference/trace_verification.png` | traced points drawn back onto the original |
| `reference/png/*.png` | each Excel chart as Excel actually renders it |
| `reference/fig*.png` | each notebook figure |
| `style/` | `style_chart.md` v1.0 + `finance_style.py` + `finance.mplstyle` — the house style |

Rebuild:

```powershell
python trace_from_image.py --write     # optional: re-derive the CSV
python build_xlsx.py
python build_notebook.py --run
```

---

## Where the data came from, and how good it is

**Playfair published the plate, not the table.** The numbers here were traced off the
engraving:

1. Find the plate's own gridlines — rows and columns whose ink fraction clears 0.55.
2. Fit `value = m·y + c` through the 19 horizontal rules, whose values are known (190 down
   to 10, £10.000 apart; the plate says so in its own caption). Max residual **0,48 units**.
3. Map years to x **piecewise** between the decade rules. They sit 213–227 px apart, not
   evenly — a copper plate engraved by hand — and a single linear fit is visibly off by
   the 1750s.
4. Separate the two inks. Both are `R − G > 45`; what splits them is `G` against `B` —
   the ochre imports line is warm (`G − B > 14`), the crimson exports line is cool
   (`G − B < 9`). Take the largest contiguous run per column, which is the stroke rather
   than speckle from the wash.
5. **Draw the sampled points back onto the plate and look at them** —
   `reference/trace_verification.png`. This is the only step that can tell you the trace
   locked onto the right line; the arithmetic above cannot.

The plate is 5,75 px per thousand pounds, so a stroke centre is good to about ±0,3.
After the ink and the antialiasing, **treat the CSV as ±1 (thousand pounds)**. Every
artefact in this folder says so in its source line. This is a legitimate way to recover a
historical series and a terrible way to report a current one.

Cross-checks that make the trace credible: exports start at 33,5 and end at 186,7 against
the plate's 190 rule; imports peak at 102,8 just over the heavy £100.000 line, in 1723–26,
where Playfair drew the peak; and the two lines cross between 1754 and 1755, which is where
the pink band closes on the original.

---

## The three charts

**1 · `Playfair_1786` — the replica.** Boxed plot, gridlines both ways, nineteen labelled
ticks on the **right**-hand axis, two saturated bands, the heavy rule at £100.000, and all
four of Playfair's captions in his own words. It breaks the house style in almost every
respect on purpose: it is the artefact, and if the traced numbers were wrong the replica
would not sit on the original.

**2 · `Balance` — the house style.** `style_chart.md` v1.0 applied literally: takeaway
title, grey default with one accent on exports, direct labels at each line's own end, no
legend, horizontal gridlines only, value axis from 0, and an **event line** on 1755 with
the cause written on the chart.

**3 · `Deviation` — the balance on its own.** Chart 2 asks the reader to subtract two lines
by eye; this one plots the subtraction. Coloured by *favourability* rather than direction
(§1 rule 10) — teal in England's favour, claret against — with signed labels at the first
year, the turn and the last, because teal and claret are only 12,3 L\* apart and greyscale
must still read (§2.6).

Playfair, with no convention to follow, already had a takeaway title
(*BALANCE in FAVOUR of ENGLAND*), direct labels on the curves (*Line of Imports*), the unit
stated once (*the Right hand line into L10,000 each*) and a reference line at a round
number. Where he would fail a review today: nineteen labelled gridlines and a full box are
noise, the two saturated fills leave the eye nowhere to rest, and there is no source line.
Chart 2 is the same argument with about 90 % of the ink removed.

---

## Deviations from `style_chart.md`, and why

| Rule | What was done | Why |
| --- | --- | --- |
| §5 currency `₫` after the number | `£` **before** the number | A 1786 sterling series cannot be denominated in đồng. Number *formatting* still follows §5: dot thousands, comma decimal, U+2212 minus. |
| §5 vi-VN separators in Excel | Excel number-format codes are written locale-neutrally (`#,##0`) | `gotchas.md`: a format code with a hard-coded comma decimal breaks on another machine. Excel renders separators from the workbook locale, and that is not addressable from openpyxl. The vi-VN convention is enforced where it can be — every string this code writes. |
| §4 "no frame, no border" | Sheet 1 has a full box frame | Sheet 1 is a replica of a 1786 plate, not a house chart. Sheets 2 and 3 are the house version of the same data. |
| §2.2 "line charts: max 3 series" | Sheet 1 draws 3 lines + 2 bands | Same reason. Sheets 2 and 3 stay inside the limit. |
| §3 axis title carries the unit | Excel charts have an axis title; the notebook figures do not | The unit is stated **once**. An Excel chart object has no subtitle, so there the axis title is the only place it can live; in matplotlib the subtitle is part of the figure, so an axis title would repeat it. |
| Language | English labels | The subject is English trade history and Playfair's own wording is quoted throughout. |

---

## What broke, and what fixed it

Six things that were invisible from Python and only showed up once Excel or a renderer had
drawn the chart.

**1 · Data-label number formats were silently ignored.** openpyxl writes
`<numFmt formatCode="…"/>` and stops; Excel reads a missing `sourceLinked` as *true*, uses
the **cell's** format instead, and the label format does nothing. Passing a `NumFmt` object
does not help — the descriptor calls `str()` on it. Fixed by patching the saved package
(`fix_label_number_formats` in `build_xlsx.py`). Symptom before the fix: `187` asked for,
`186.7` drawn.

**2 · `showVal=False` at *list* level is load-bearing.** Setting `dLbl=[one_label]` and
leaving `DataLabelList.showVal` unset made Excel label **all 81 points** on both curves —
757 reported collisions from one missing keyword.

**3 · A category axis has no `majorGridlines` object at all.** The value axis gets one by
default; `x.majorGridlines.spPr = …` on a category axis is an `AttributeError` until you
assign a `ChartLines()` first.

**4 · An unrotated value-axis title lands on the tick labels.** `rot=0` has to go on the
title's own `bodyPr` (setting it on the axis does nothing), and even then Excel centres the
box vertically on the axis, straight through the tick labels. It needs a `ManualLayout`
with **`w` as well as `x`/`y`** — without `w`, Excel sizes the box to the plot's left
margin and wraps a short unit label onto two lines, and the second line lands on the top
gridline.

**5 · `check_chart_overlap.ps1` had no vertical test for value-axis titles.** It assumed
the title sits to the *left* of the tick-label band, which is true only for a rotated
title. A horizontal title parked above the plot clears `InsideLeft` by construction and was
reported as a collision every time. The detector now requires vertical overlap with the
plot floor as well — patched in the skill, not worked around here.

**6 · In matplotlib, `1700` sat on top of `0`.** With the first category on the axis, the
x tick label and the value label share the bottom-left corner. `checks/check_chart_text.py`
(from `hoa_gao_case`) caught it; two years of x-axis inset fixed it. Both checkers now
report clean:

```
3 charts checked, 0 overlap(s)          # check_chart_overlap.ps1
checked 3 charts / no clipped text, no overlapping text   # check_chart_text.py
```

---

## The technique worth keeping

**Shading between two curves is one line in matplotlib and three series in Excel.**

```python
ax.fill_between(year, imports, exports, where=imports >= exports,
                color=PF_AGAINST, interpolate=True)
```

`interpolate=True` closes the band exactly at the crossing rather than at the next year.
Excel has no equivalent, so the same two bands cost a **stacked area chart with an
invisible base**:

| column | formula | role |
| --- | --- | --- |
| `E` | `=MIN(C,D)` | invisible base — lifts the stack to the lower curve |
| `F` | `=MAX(C-D,0)` | the deficit gap, pink |
| `G` | `=MAX(D-C,0)` | the surplus gap, buff |

Only one of `F` and `G` is ever non-zero, so the stack tops out at `MAX(imports, exports)`
and the two real lines are drawn on top of it as a combo line chart. Add the base series
**first**: a later series draws in front.

Worth knowing before promising anyone a shaded band in a workbook.
