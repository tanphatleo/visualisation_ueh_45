"""
Anscombe's quartet as a live Excel workbook — style/style_chart.md v1.0.

The point of the quartet is that the summary statistics agree and the pictures
do not, so the workbook has to make both halves visible at once:

  * the four datasets sit in one table, as numbers a reader can edit;
  * two summary blocks underneath are LIVE FORMULAS over that table, never
    typed constants — change a y and watch every statistic move:
      - block 1 describes each of the eight columns on its own terms, the
        ordinary descriptive run anyone would do before plotting;
      - block 2 describes the fitted line, dataset by dataset;
  * the four scatters are small multiples on a shared axis (§6, "many
    entities, same measure"), so the only thing that differs between the
    panels is the data.

§6 forbids a trend line without stating R² and n. Both are in each panel's
descriptive subtitle, and block 2 shows the full working.

Run:  python build_xlsx.py
"""

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.trendline import Trendline
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "style"))
import xl_style as xs  # noqa: E402

OUT = HERE / "anscombe_quartet.xlsx"

# --------------------------------------------------------------------------
# The data — Anscombe, F. J. (1973), Table 1.
# --------------------------------------------------------------------------
QUARTET = [
    # (x, y) for datasets I, II, III, IV
    [(10.0, 8.04), (8.0, 6.95), (13.0, 7.58), (9.0, 8.81), (11.0, 8.33),
     (14.0, 9.96), (6.0, 7.24), (4.0, 4.26), (12.0, 10.84), (7.0, 4.82),
     (5.0, 5.68)],
    [(10.0, 9.14), (8.0, 8.14), (13.0, 8.74), (9.0, 8.77), (11.0, 9.26),
     (14.0, 8.10), (6.0, 6.13), (4.0, 3.10), (12.0, 9.13), (7.0, 7.26),
     (5.0, 4.74)],
    [(10.0, 7.46), (8.0, 6.77), (13.0, 12.74), (9.0, 7.11), (11.0, 7.81),
     (14.0, 8.84), (6.0, 6.08), (4.0, 5.39), (12.0, 8.15), (7.0, 6.42),
     (5.0, 5.73)],
    [(8.0, 6.58), (8.0, 5.76), (8.0, 7.71), (8.0, 8.84), (8.0, 8.47),
     (8.0, 7.04), (8.0, 5.25), (19.0, 12.50), (8.0, 5.56), (8.0, 7.91),
     (8.0, 6.89)],
]
N = len(QUARTET[0])
SERIES_HEADERS = ["x₁", "y₁", "x₂", "y₂", "x₃", "y₃", "x₄", "y₄"]
ROMAN = ["I", "II", "III", "IV"]

# One sentence per panel — §1 rule 1. The descriptive line underneath carries
# n and R², which §6 makes mandatory for any chart with a trend line.
PANELS = [
    ("A line genuinely fits Dataset I",      "Noisy, but linear"),
    ("Dataset II is a curve, not a line",    "A perfect relationship, just not a straight one"),
    ("One outlier tilts Dataset III's line", "Ten collinear points plus one high y"),
    ("One point invents Dataset IV's line",  "x is 8 for ten of the eleven points"),
]

# Shared axis window for all four panels. Data spans x 4–19, y 3,1–12,74; the
# window is rounded outwards and IDENTICAL across panels, because a small
# multiple whose axes move is not a comparison.
X_MIN, X_MAX, X_UNIT = 2, 20, 2
Y_MIN, Y_MAX, Y_UNIT = 2, 14, 2

# --------------------------------------------------------------------------
# Layout. Column B is the label gutter for the summary blocks; the data table
# starts at C so both blocks line up column-for-column underneath it.
# --------------------------------------------------------------------------
LABEL_COL = 2                      # B
TBL_COL = 3                        # C — x₁
TBL_TOP = 8                        # header row of the data table
DATA_FIRST, DATA_LAST = TBL_TOP + 1, TBL_TOP + N
DESC_TOP = DATA_LAST + 3           # note row of the per-series block
CHART_ANCHORS = ["L8", "T8", "L24", "T24"]
PANEL_W, PANEL_H = 12.7, 8.2       # cm — small multiples, not the 22.9x11.4 block

# Every function below is a LEGACY name — COUNT, AVERAGE, MEDIAN, STDEV,
# QUARTILE, VAR, CORREL, SLOPE, INTERCEPT, RSQ. openpyxl writes those without
# an _xlfn prefix and Excel evaluates them on open; STDEV.S and QUARTILE.INC
# would need the prefix and return the same numbers.
DESCRIPTIVE = [
    ("n",               "=COUNT({s})",          '0'),
    ("Mean",            "=AVERAGE({s})",        '0.000'),
    ("Median",          "=MEDIAN({s})",         '0.000'),
    ("Std deviation",   "=STDEV({s})",          '0.000'),
    ("Sample variance", "=VAR({s})",            '0.000'),
    ("Minimum",         "=MIN({s})",            '0.00'),
    ("Lower quartile",  "=QUARTILE({s},1)",     '0.00'),
    ("Upper quartile",  "=QUARTILE({s},3)",     '0.00'),
    ("Maximum",         "=MAX({s})",            '0.00'),
    ("Range",           "=MAX({s})-MIN({s})",   '0.00'),
]

FIT = [
    ("Correlation of x and y", "=CORREL({x},{y})",    '0.000'),
    ("Slope",                  "=SLOPE({y},{x})",     '0.000'),
    ("Intercept",              "=INTERCEPT({y},{x})", '0.000'),
    ("R²",                     "=RSQ({y},{x})",       '0.000'),
    ("n",                      "=COUNT({y})",         '0'),
]


def _range(col_offset):
    """The A1 range of one data column, e.g. 'C9:C19'."""
    letter = get_column_letter(TBL_COL + col_offset)
    return f"{letter}{DATA_FIRST}:{letter}{DATA_LAST}"


def build_table(ws):
    """The four datasets side by side — the only typed numbers in the workbook."""
    row = xs.write_table_header(ws, TBL_TOP, SERIES_HEADERS, col=TBL_COL)
    for i in range(N):
        for d, pairs in enumerate(QUARTET):
            for j, v in enumerate(pairs[i]):
                c = ws.cell(row + i, TBL_COL + d * 2 + j, v)
                c.font = Font(name=xs.FONT, size=11, color=xs.INK)
                c.number_format = '0.00'
                c.alignment = Alignment(horizontal="right")
    return row + N


def _block(ws, top, note, headers, rows, bold_label=None):
    """A labelled block of live formulas: note, header row, one row per statistic."""
    row = xs.note(ws, top, note, col=LABEL_COL)
    row = xs.write_table_header(ws, row, ["Statistic"] + headers, col=LABEL_COL)
    for k, (label, cells, fmt) in enumerate(rows):
        c = ws.cell(row + k, LABEL_COL, label)
        c.font = Font(name=xs.FONT, size=11, color=xs.TEXT_2)
        for j, formula in enumerate(cells):
            c = ws.cell(row + k, LABEL_COL + 1 + j, formula)
            c.font = Font(name=xs.FONT, size=11, color=xs.INK,
                          bold=(label == bold_label))
            c.number_format = fmt
            c.alignment = Alignment(horizontal="right")
    return row + len(rows)


def build_descriptive(ws):
    """Block 1 — each of the eight columns described on its own terms."""
    rows = [(label, [f.format(s=_range(i)) for i in range(8)], fmt)
            for label, f, fmt in DESCRIPTIVE]
    return _block(
        ws, DESC_TOP,
        "Descriptive statistics, series by series — live formulas over the table "
        "above. Mean, standard deviation and variance agree across all four "
        "datasets; the median and the quartiles do not, and that is the first "
        "hint that the numbers are hiding something.",
        SERIES_HEADERS, rows)


def build_fit(ws, top):
    """Block 2 — the fitted line, dataset by dataset. The numbers §6 demands."""
    rows = [(label,
             [f.format(x=_range(d * 2), y=_range(d * 2 + 1)) for d in range(4)],
             fmt)
            for label, f, fmt in FIT]
    return _block(
        ws, top,
        "The fitted straight line, dataset by dataset. Same correlation, same "
        "slope, same intercept, same R² — and four different pictures.",
        ROMAN, rows, bold_label="R²")


def panel(ws, idx):
    """One small multiple: the points in accent, the fitted line in context grey."""
    takeaway, descriptive = PANELS[idx]
    xcol = TBL_COL + idx * 2

    chart = ScatterChart()
    chart.scatterStyle = "marker"        # markers only; no connecting line

    xref = Reference(ws, min_col=xcol, min_row=DATA_FIRST, max_row=DATA_LAST)
    yref = Reference(ws, min_col=xcol + 1, min_row=DATA_FIRST, max_row=DATA_LAST)
    s = Series(yref, xref, title=f"Dataset {ROMAN[idx]}")

    # The points are the story, so they carry the one accent (§2.3). The line
    # is identical in all four panels — that is exactly why it stays grey.
    s.marker = Marker(
        symbol="circle", size=7,
        spPr=GraphicalProperties(
            solidFill=xs.ACCENT,
            ln=LineProperties(solidFill=xs.SURFACE, w=xs.W_THIN)))
    s.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    # dispEq / dispRSqr must be written explicitly as false. Leave them out and
    # Excel draws its own equation-and-R² caption across the middle of the plot
    # — the XML openpyxl wrote contained neither element, and Excel supplied
    # its own default anyway.
    s.trendline = Trendline(
        trendlineType="linear", dispEq=False, dispRSqr=False,
        spPr=GraphicalProperties(
            ln=LineProperties(solidFill=xs.CONTEXT_LINE, w=xs.W_THIN)))
    chart.series.append(s)

    # R² and n on the face of the chart — §6 makes both mandatory next to a
    # trend line, and a chart pasted into a deck leaves the worksheet behind.
    r2 = rsq([p[0] for p in QUARTET[idx]], [p[1] for p in QUARTET[idx]])
    chart.title = xs.chart_message(
        takeaway,
        f"{descriptive}  ·  n = {N}  ·  R² = {r2:.2f}".replace(".", ","),
        size_pt=10, sub_pt=8)
    # §1 rule 7 — the message sits top-left. Excel centres a chart title over
    # the plot area whatever the paragraph alignment says, so the only way to
    # move it is a manual layout pinned to the chart-area edge.
    chart.title.layout = Layout(
        manualLayout=ManualLayout(xMode="edge", yMode="edge", x=0.01, y=0.02))

    xs.declutter(chart, gridlines=True, x_line=True)
    xs.no_legend(chart)
    chart.width, chart.height = PANEL_W, PANEL_H

    # Shared window, stated on every panel so the four are directly comparable.
    for ax, (lo, hi, unit) in ((chart.x_axis, (X_MIN, X_MAX, X_UNIT)),
                               (chart.y_axis, (Y_MIN, Y_MAX, Y_UNIT))):
        ax.scaling.min, ax.scaling.max = lo, hi
        ax.majorUnit = unit
        ax.delete = False
        ax.numFmt = '0'
        ax.number_format = '0'
    xs.axis_title(chart.x_axis, f"x{chr(0x2080 + idx + 1)}")
    xs.axis_title(chart.y_axis, f"y{chr(0x2080 + idx + 1)}")

    ws.add_chart(chart, CHART_ANCHORS[idx])
    return chart


def rsq(xs_, ys):
    """Plain least-squares R², for the subtitle text only — Excel recomputes its own."""
    n = len(xs_)
    mx, my = sum(xs_) / n, sum(ys) / n
    sxy = sum((a - mx) * (b - my) for a, b in zip(xs_, ys))
    sxx = sum((a - mx) ** 2 for a in xs_)
    syy = sum((b - my) ** 2 for b in ys)
    return (sxy * sxy) / (sxx * syy)


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Anscombe"
    ws.sheet_view.showGridLines = False

    xs.write_header(
        ws, 2,
        "Four datasets with the same statistics look nothing alike",
        "Anscombe's quartet · n = 11 each · identical mean, variance, "
        "correlation and fitted line to 2 dp",
        col=LABEL_COL)

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 24
    xs.widths(ws, [9] * 8, start=TBL_COL)

    build_table(ws)
    end = build_descriptive(ws)
    end = build_fit(ws, end + 2)

    xs.write_source(
        ws, end + 2,
        "Source: Anscombe, F. J. (1973), ‘Graphs in Statistical Analysis’, "
        "The American Statistician 27(1), pp. 17–21",
        col=LABEL_COL)

    for i in range(4):
        panel(ws, i)

    wb.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
