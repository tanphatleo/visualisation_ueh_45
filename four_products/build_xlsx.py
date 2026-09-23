"""
Four products, one average — a time-series twin of Anscombe's quartet.
style/style_chart.md v1.0.

Twelve months of sales for four products. Every product has EXACTLY the same
mean (8,00 MM) and the same standard deviation (3,00 MM), and the four monthly
patterns could hardly be less alike:

    A  growing, with one bad month
    B  an S-curve that has already flattened
    C  flat for ten months, then a step change
    D  no trend at all — it just oscillates

How the data was made (see DERIVATION below): each product starts as a hand-drawn
shape, is standardised to mean 0 / sd 1, and is mapped onto mean 8 / sd 3. Rounding
to two decimals would spoil the match, so the values are then nudged on the 0,01
grid until the two constraints hold exactly:

    sum(v) = 96   and   sum(v²) = 11·3² + 12·8² = 867

Both summary blocks are live formulas over the table, so a reader who edits a
month watches the equality break — which is the point.

Run:  python build_xlsx.py
"""

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "style"))
import xl_style as xs  # noqa: E402

OUT = HERE / "four_products.xlsx"

# --------------------------------------------------------------------------
# The data. Monthly sales, MM. Constructed — see the module docstring.
# Verified: every column has mean 8,000000000000 and sample sd 3,000000000000.
# --------------------------------------------------------------------------
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
PRODUCTS = ["Product A", "Product B", "Product C", "Product D"]
SALES = {
    "Product A": [4.81, 5.95, 6.18, 6.27, 6.82, 2.78, 8.75, 9.09, 10.66, 10.21, 12.17, 12.31],
    "Product B": [4.09, 4.28, 4.71, 5.38, 6.23, 7.42, 8.71, 9.89, 10.77, 11.29, 11.56, 11.67],
    "Product C": [6.18, 6.27, 6.38, 6.51, 6.71, 6.81, 6.91, 7.05, 7.12, 7.32, 13.99, 14.75],
    "Product D": [8.96, 8.18, 13.40, 7.91, 6.74, 12.35, 6.48, 5.95, 10.14, 4.24, 3.21, 8.44],
}
N = len(MONTHS)
MEAN, SD = 8.0, 3.0

# One sentence per panel — §1 rule 1.
PANELS = [
    ("Product A is growing, one bad month aside", "June collapsed to 2,8, then recovered"),
    ("Product B's growth has already flattened",  "An S-curve: fast to September, flat since"),
    ("Product C did nothing until November",      "Ten flat months, then a step change"),
    # Verified against the data: Mar 13,40 · Jun 12,35 · Sep 10,14 · Dec 8,44 —
    # each one a local peak, each lower than the last, over a level that falls
    # from 8,96 to 3,21. Slope of the twelve months is −0,39 per month.
    ("Product D is declining, but does well at every quarter end",
     "Each spike is lower than the last"),
]

# Shared y window for all four panels — a small multiple whose axes move is not
# a comparison. Data runs 2,78 to 14,75.
Y_MIN, Y_MAX, Y_UNIT = 0, 16, 2

# --------------------------------------------------------------------------
# Layout. Column B doubles as the Month column and the label gutter for the
# summary blocks, so the blocks line up under the products.
# --------------------------------------------------------------------------
LABEL_COL = 2                       # B — "Month", then "Statistic"
TBL_COL = 3                         # C — Product A
DELTA_COL = TBL_COL + 5             # H — month-on-month changes
TBL_TOP = 8
DATA_FIRST, DATA_LAST = TBL_TOP + 1, TBL_TOP + N
DESC_TOP = DATA_LAST + 3
CHART_ANCHORS = ["S8", "AC8", "S24", "AC24"]
PANEL_W, PANEL_H = 15.5, 8.2        # cm — wide, so the months breathe

# Legacy function names only — AVERAGE, MEDIAN, STDEV, VAR, QUARTILE, COUNT —
# so openpyxl writes them without an _xlfn prefix and Excel evaluates them on
# open. STDEV.S and QUARTILE.INC would need the prefix for the same numbers.
DESCRIPTIVE = [
    ("n",                 "=COUNT({s})",        '0'),
    ("Mean",              "=AVERAGE({s})",      '0.00'),
    ("Std deviation",     "=STDEV({s})",        '0.00'),
    ("Sample variance",   "=VAR({s})",          '0.00'),
    ("Median",            "=MEDIAN({s})",       '0.00'),
    ("Minimum",           "=MIN({s})",          '0.00'),
    ("Lower quartile",    "=QUARTILE({s},1)",   '0.00'),
    ("Upper quartile",    "=QUARTILE({s},3)",   '0.00'),
    ("Maximum",           "=MAX({s})",          '0.00'),
    ("Range",             "=MAX({s})-MIN({s})", '0.00'),
]

# What the summary misses. These are the numbers that actually separate the
# four, and none of them appear in a mean-and-sd summary.
SHAPE = [
    ("First month",            "={f}",                              '0.00'),
    ("Last month",             "={l}",                              '0.00'),
    ("Change, first to last",  "={l}-{f}",                          '+0.00;−0.00'),
    ("Biggest month-on-month rise", "=MAX({d})",                    '+0.00;−0.00'),
    ("Biggest month-on-month fall", "=MIN({d})",                    '+0.00;−0.00'),
    # {d} is the delta helper block, not an inline range subtraction:
    # =MAX(C10:C20-C9:C19) is an ARRAY formula, and openpyxl writes it as an
    # ordinary one, so Excel returns #VALUE! on open. A visible helper column
    # is both correct and better teaching than a hidden CSE formula.
    ("Months above the mean",  "=COUNTIF({s},\">\"&AVERAGE({s}))",  '0'),
]


def _col(offset):
    """A1 range of one product column, e.g. 'C9:C20'."""
    letter = get_column_letter(TBL_COL + offset)
    return f"{letter}{DATA_FIRST}:{letter}{DATA_LAST}"


def build_table(ws):
    """Months down the side, one column per product — the only typed numbers here."""
    row = xs.write_table_header(ws, TBL_TOP, ["Month"] + PRODUCTS, col=LABEL_COL)
    for i, m in enumerate(MONTHS):
        c = ws.cell(row + i, LABEL_COL, m)
        c.font = Font(name=xs.FONT, size=11, color=xs.TEXT_2)
        c.alignment = Alignment(horizontal="left")
        for d, p in enumerate(PRODUCTS):
            c = ws.cell(row + i, TBL_COL + d, SALES[p][i])
            c.font = Font(name=xs.FONT, size=11, color=xs.INK)
            c.number_format = '0.00'
            c.alignment = Alignment(horizontal="right")
    return row + N


def build_delta_helper(ws):
    """
    Month-on-month change, one column per product. Feeds the biggest rise and
    biggest fall rows in the shape block — which would otherwise need an array
    formula that openpyxl cannot write.
    """
    xs.note(ws, TBL_TOP - 1,
            "Helper — month-on-month change. Feeds the biggest rise and biggest "
            "fall in the shape block below.", col=DELTA_COL)
    for d, p in enumerate(PRODUCTS):
        c = ws.cell(TBL_TOP, DELTA_COL + d, f"Δ {p[-1]}")
        c.font = Font(name=xs.FONT, size=10, color=xs.TEXT_2, italic=True)
        letter = get_column_letter(TBL_COL + d)
        for i in range(1, N):        # month 1 has no previous month
            c = ws.cell(DATA_FIRST + i, DELTA_COL + d,
                        f"={letter}{DATA_FIRST + i}-{letter}{DATA_FIRST + i - 1}")
            c.font = Font(name=xs.FONT, size=10, color=xs.CONTEXT_LINE)
            c.number_format = '+0.00;−0.00'


def _block(ws, top, note, rows, bold_label=None):
    """A labelled block of live formulas: note, header row, one row per statistic."""
    row = xs.note(ws, top, note, col=LABEL_COL)
    row = xs.write_table_header(ws, row, ["Statistic"] + PRODUCTS, col=LABEL_COL)
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
    """Block 1 — the summary a reader would run before plotting anything."""
    rows = [(label, [f.format(s=_col(d)) for d in range(4)], fmt)
            for label, f, fmt in DESCRIPTIVE]
    return _block(
        ws, DESC_TOP,
        "Descriptive statistics — live formulas over the table above. Mean and "
        "standard deviation are identical for all four products; the median and "
        "the quartiles already disagree, and the charts disagree completely.",
        rows, bold_label="Std deviation")


def build_shape(ws, top):
    """Block 2 — the numbers that actually tell the four apart."""
    rows = []
    for label, f, fmt in SHAPE:
        cells = []
        for d in range(4):
            letter = get_column_letter(TBL_COL + d)
            delta = get_column_letter(DELTA_COL + d)
            cells.append(f.format(
                s=_col(d),
                f=f"{letter}{DATA_FIRST}",
                l=f"{letter}{DATA_LAST}",
                d=f"{delta}{DATA_FIRST + 1}:{delta}{DATA_LAST}"))
        rows.append((label, cells, fmt))
    return _block(
        ws, top,
        "What the summary above cannot see: where each product started, where it "
        "finished and how it moved. This is the shape — and only the chart shows "
        "it at a glance.",
        rows)


def panel(ws, idx):
    """One small multiple: the product's line in accent, its mean as a dashed rule."""
    takeaway, descriptive = PANELS[idx]

    chart = LineChart()
    cats = Reference(ws, min_col=LABEL_COL, min_row=DATA_FIRST, max_row=DATA_LAST)

    val = Reference(ws, min_col=TBL_COL + idx,
                    min_row=DATA_FIRST, max_row=DATA_LAST)
    s = Series(val, title=PRODUCTS[idx])
    xs.fill(s, xs.ACCENT, line=True)          # 2,5 pt accent line
    # xs.fill turns markers off; put them back as white discs with an accent
    # ring, so each of the twelve months is a readable, countable point.
    s.marker = Marker(
        symbol="circle", size=7,
        spPr=GraphicalProperties(
            solidFill=xs.SURFACE,
            ln=LineProperties(solidFill=xs.ACCENT, w=xs.W_CONTEXT)))
    chart.series.append(s)
    chart.set_categories(cats)

    # Takeaway only — no descriptive second line, at the deck's request.
    chart.title = xs.chart_message(takeaway, None, size_pt=20)
    # §1 rule 7 — Excel centres a chart title over the plot area whatever the
    # paragraph alignment says; a manual layout is the only way to pin it left.
    chart.title.layout = Layout(
        manualLayout=ManualLayout(xMode="edge", yMode="edge", x=0.01, y=0.02))

    xs.declutter(chart, gridlines=True, x_line=True)
    # declutter strips the category gridlines (§4 asks for horizontal only);
    # this deck wants both, so put the verticals back afterwards.
    chart.x_axis.majorGridlines = ChartLines(
        spPr=GraphicalProperties(ln=LineProperties(solidFill=xs.GRID, w=xs.W_THIN)))
    # By default a category axis puts each month in a BAND and the tick (so the
    # gridline) on the band edge, i.e. between two months. crossBetween="midCat"
    # — set on the VALUE axis, which is where OOXML keeps it — puts the ticks
    # through the middle of each month instead, under its label and through its
    # data point.
    chart.y_axis.crossBetween = "midCat"
    xs.no_legend(chart)
    chart.width, chart.height = PANEL_W, PANEL_H

    chart.y_axis.scaling.min, chart.y_axis.scaling.max = Y_MIN, Y_MAX
    chart.y_axis.majorUnit = Y_UNIT
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    # Without this the axis inherits the cells' '0.00' and reads 16.00 / 12.00.
    chart.y_axis.numFmt = '0'
    chart.y_axis.number_format = '0'
    # Both axis titles off: Jan-Dec explains itself, and the y-axis title was
    # dropped at the deck's request. Units (MM) therefore live only on the
    # worksheet header — worth restating in the slide caption.
    chart.y_axis.title = None
    chart.x_axis.title = None

    ws.add_chart(chart, CHART_ANCHORS[idx])
    return chart


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "Four products"
    ws.sheet_view.showGridLines = False

    xs.write_header(
        ws, 2,
        "Four products with the same average and the same spread",
        "Monthly sales, MM · 12 months · every product has mean 8,00 and "
        "standard deviation 3,00",
        col=LABEL_COL)

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 24
    xs.widths(ws, [12] * 4, start=TBL_COL)
    xs.widths(ws, [10] * 4, start=DELTA_COL)

    build_table(ws)
    build_delta_helper(ws)
    end = build_descriptive(ws)
    end = build_shape(ws, end + 2)

    xs.write_source(
        ws, end + 2,
        "Source: illustrative dataset, constructed so that all four products "
        "share a mean of 8,00 MM and a standard deviation of 3,00 MM. Not real "
        "sales.", col=LABEL_COL)

    for i in range(4):
        panel(ws, i)

    wb.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
