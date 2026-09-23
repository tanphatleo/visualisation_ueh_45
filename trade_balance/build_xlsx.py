"""
Playfair's trade-balance chart, rebuilt in Excel.

    python build_xlsx.py

Source of the original:
    Jeremy Norman, History of Information, entry 2929 — "William Playfair Founds
    Statistical Graphics, and Invents the Line Chart and Bar Chart".
    https://www.historyofinformation.com/detail.php?entryid=2929
    Plate: "Exports and Imports to and from DENMARK & NORWAY from 1700 to 1780",
    The Commercial and Political Atlas, London, 1786.

The data is traced from the engraving, not transcribed from a ledger — see
`trace_from_image.py` and README.md. Treat it as +/- 1 (thousand pounds).

Three sheets, three jobs:

    Playfair_1786   a faithful replica of the 1786 plate. Deliberately breaks the
                    house style: box frame, both sets of gridlines, right-hand
                    value axis, two filled bands. That IS the artefact.
    Balance         the same data under style/style_chart.md v1.0 — one accent,
                    two lines, direct labels, horizontal gridlines, an event line
                    on the year the balance turns.
    Deviation       the balance itself, coloured by favourability (§1 rule 10).

Everything above and below each plot is worksheet cells: an Excel chart object has
no subtitle, no source line and no annotation layer.
"""
import csv
import re
import shutil
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (CharacterProperties, Paragraph,
                                   ParagraphProperties, RegularTextRun,
                                   RichTextProperties)
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# Style roles — style/style_chart.md section 2.1. Never a raw hex at a call site.
# ---------------------------------------------------------------------------
SURFACE, WASH, GRID = "FFFFFF", "F7F7F7", "E8E8E8"
CONTEXT_FILL, CONTEXT_LINE = "BDBDBD", "8C8C8C"
TEXT_2, INK = "595959", "262A33"
ACCENT, ALERT, GOOD, BAD = "0F5499", "990F3D", "0D7680", "990F3D"
# Light arms of the diverging ramp (section 2.6) — band fills, never lines.
GOOD_TINT, BAD_TINT = "C8DEE0", "E4C3CE"

FONT = "Aptos"
PT = 12700                      # EMU per point, for LineProperties.w
W_HAIR, W_THIN = int(0.75 * PT), int(1.5 * PT)
W_BOLD, W_FRAME = int(2.5 * PT), int(2.25 * PT)

# Colours measured off the engraving itself (median of the stroke / fill pixels).
PF_IMPORT, PF_EXPORT = "D48D50", "B05058"      # ochre and crimson lines
PF_AGAINST, PF_FAVOUR = "F7DAD7", "E7D9B5"     # pink and buff bands
PF_RULE, PF_FRAME = "8A837C", "4A443E"

SOURCE = ("Source: William Playfair, The Commercial and Political Atlas, London 1786, "
          "plate 'Exports and Imports to and from Denmark & Norway from 1700 to 1780', "
          "via historyofinformation.com entry 2929 · Values traced from the engraving "
          "to ±1, not transcribed from ledgers · Traced 18 Aug 2026")

TURN_YEAR = 1755                # the year the traced lines cross
HDR = 8                         # table header row
R0 = HDR + 1                    # first data row


# ---------------------------------------------------------------------------
# Small text helpers
# ---------------------------------------------------------------------------
def rich(size_pt, colour, bold=False, italic=False):
    cp = CharacterProperties(sz=int(size_pt * 100), b=bold, i=italic,
                             solidFill=colour, latin=None)
    return RichText(bodyPr=RichTextProperties(),
                    p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def chart_title(text, size_pt=13, colour=INK, bold=True, italic=False, left=True):
    """
    A chart title box. algn='l' left-aligns the text INSIDE the box; it does not
    move the box, which Excel centres. Pinning the box with a manual layout is
    what actually puts a takeaway title top-left (see the skill's slope-charts.md).
    """
    cp = CharacterProperties(sz=int(size_pt * 100), b=bold, i=italic, solidFill=colour)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp, algn="l" if left else "ctr"),
                     r=[RegularTextRun(rPr=cp, t=text)])
    t = Title(tx=Text(rich=RichText(p=[para])), overlay=False)
    if left:
        t.layout = Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge",
                                                    x=0.01, y=0.015))
    return t


def axis_title(text, colour=TEXT_2, size_pt=11, x=0.005, y=0.105):
    """
    An axis title that stays HORIZONTAL. Excel rotates a value-axis title 90 degrees
    by default, and the house style forbids rotated text, so rot/vert are forced on
    the title's own bodyPr - setting them on the axis does nothing.
    """
    cp = CharacterProperties(sz=int(size_pt * 100), b=False, solidFill=colour)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp, algn="l"),
                     r=[RegularTextRun(rPr=cp, t=text)])
    t = Title(tx=Text(rich=RichText(bodyPr=RichTextProperties(rot=0, vert="horz"),
                                    p=[para])), overlay=False)
    # Unrotating the title is not enough: Excel still centres the box vertically on
    # the axis, where it lands on the tick labels. Pin it above the plot instead.
    # w matters as much as x/y: without it Excel sizes the box to the plot's left
    # margin and wraps a perfectly short unit label onto two lines.
    t.layout = Layout(manualLayout=ManualLayout(xMode="edge", yMode="edge",
                                                x=x, y=y, w=0.45, h=0.06))
    return t


def header(ws, title, subtitle):
    ws["B2"] = title
    ws["B2"].font = Font(name=FONT, size=16, bold=True, color=INK)
    ws["B3"] = subtitle
    ws["B3"].font = Font(name=FONT, size=11, color=TEXT_2)


def source_at(ws, cell):
    ws[cell] = SOURCE
    ws[cell].font = Font(name=FONT, size=9, color=CONTEXT_LINE)
    ws[cell].alignment = Alignment(vertical="top")


def note(ws, cell, text, colour=ALERT):
    ws[cell] = text
    ws[cell].font = Font(name=FONT, size=9, italic=True, color=colour)


def table_header(ws, row, col, labels, fill=ACCENT):
    for j, label in enumerate(labels):
        c = ws.cell(row, col + j, label)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="right" if j else "left")


def declutter_value_axis(axis, gridlines=True, labels=True, fmt="#,##0"):
    """Horizontal gridlines only, no axis line, no tick marks (section 4)."""
    axis.delete = False                     # never leave this unset - see gotchas.md
    axis.majorTickMark = axis.minorTickMark = "none"
    axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    axis.txPr = rich(9, TEXT_2)
    axis.numFmt = fmt
    if gridlines:
        axis.majorGridlines.spPr = GraphicalProperties(
            ln=LineProperties(solidFill=GRID, w=W_HAIR))
    else:
        axis.majorGridlines = None
    if not labels:
        axis.numFmt = ";;;"


def declutter_category_axis(axis, skip=10, baseline=True):
    axis.delete = False
    axis.majorGridlines = None
    axis.majorTickMark = axis.minorTickMark = "none"
    axis.tickLblSkip = skip
    axis.tickMarkSkip = skip
    axis.txPr = rich(9, TEXT_2)
    axis.spPr = GraphicalProperties(
        ln=LineProperties(solidFill=TEXT_2, w=W_HAIR) if baseline
        else LineProperties(noFill=True))


def end_label(series, idx, colour, show_name=True, fmt="#,##0"):
    """Direct labelling: name and value at the series' own last point (rule 5)."""
    off = dict(showCatName=False, showLegendKey=False, showPercent=False,
               showBubbleSize=False)
    lbl = DataLabel(idx=idx, showSerName=show_name, showVal=True, dLblPos="r", **off)
    lbl.txPr = rich(11, colour, bold=True)
    lbl.numFmt = fmt
    series.dLbls = DataLabelList(dLbl=[lbl], showVal=False, showSerName=False, **off)


# ---------------------------------------------------------------------------
# Data
# ---------------------------------------------------------------------------
def read_data():
    path = HERE / "data" / "playfair_denmark_norway.csv"
    with path.open(encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    return [(int(r["year"]), float(r["imports"]), float(r["exports"])) for r in rows]


DATA = read_data()
N = len(DATA)
LAST = R0 + N - 1
IDX_TURN = TURN_YEAR - DATA[0][0] + 1        # 1-based category position


def write_data_sheet(wb):
    ws = wb.create_sheet("Data_Trade")
    ws.sheet_view.showGridLines = False
    header(ws, "Playfair's Denmark & Norway plate, traced back to numbers",
           "Thousands of pounds sterling. One row per year, 1700-1780. "
           "Every chart in this workbook reads this table through formulas.")
    source_at(ws, "B5")
    note(ws, "B6",
         "These are TRACED values: each curve was found by colour in the engraving, "
         "sampled at every year and converted through the plate's own gridlines. "
         "Accuracy is about ±1. They are not Playfair's ledger figures.")
    table_header(ws, HDR, 2, ["Year", "Imports", "Exports", "Balance"])
    for i, (year, imp, exp) in enumerate(DATA):
        r = R0 + i
        ws.cell(r, 2, year).font = Font(name=FONT, size=10, color=INK)
        for j, v in enumerate((imp, exp)):
            c = ws.cell(r, 3 + j, v)
            c.font = Font(name=FONT, size=10, color=TEXT_2)
            c.number_format = "#,##0.0"
        c = ws.cell(r, 5, f"=D{r}-C{r}")
        c.font = Font(name=FONT, size=10, color=TEXT_2)
        c.number_format = "+#,##0.0;-#,##0.0;0.0"
    for col, w in zip("BCDE", (10, 11, 11, 11)):
        ws.column_dimensions[col].width = w
    return ws


def live_table(ws, cols, extra=None):
    """
    Mirror Data_Trade into this sheet through formulas, so a reader can change a
    number there and watch every chart move. `extra` adds helper columns.
    """
    table_header(ws, HDR, 2, cols)
    for i in range(N):
        r, src = R0 + i, R0 + i
        ws.cell(r, 2, f"=Data_Trade!B{src}").font = Font(name=FONT, size=10, color=INK)
        for j, letter in enumerate("CD"):
            c = ws.cell(r, 3 + j, f"=Data_Trade!{letter}{src}")
            c.font = Font(name=FONT, size=10, color=TEXT_2)
            c.number_format = "#,##0.0"
        for j, formula in enumerate(extra or []):
            c = ws.cell(r, 5 + j, formula.format(r=r))
            c.font = Font(name=FONT, size=10, color=TEXT_2)
            c.number_format = "#,##0.0"
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 11


def years_ref(ws):
    return Reference(ws, min_col=2, min_row=R0, max_row=LAST)


# ---------------------------------------------------------------------------
# Sheet 1 — the replica
# ---------------------------------------------------------------------------
def build_replica(ws):
    ws.sheet_view.showGridLines = False
    header(ws, "Replica: the 1786 plate as Playfair engraved it",
           "Exports and Imports to and from Denmark & Norway, 1700-1780, "
           "thousands of pounds. Deliberately NOT house style - see the Balance sheet for that.")
    source_at(ws, "B5")
    note(ws, "B6",
         "Helper columns E:G exist only to fake the two coloured bands. Excel cannot fill "
         "between two lines, so the bands are a STACKED AREA chart: an invisible base at "
         "MIN(imports, exports), then the deficit gap, then the surplus gap. Only one of the "
         "two gaps is ever non-zero, so the stack tops out at MAX(imports, exports) and the "
         "two real lines are drawn on top of it as a line chart.")

    # The column headers ARE the chart's text: every caption on the plate is a series
    # name shown as one data label, so Playfair's wording lives in a cell, not a shape.
    live_table(ws, ["Year", "Line of Imports", "Line of Exports", "Base (hidden)",
                    "BALANCE AGAINST", "BALANCE in FAVOUR of ENGLAND", "100"],
               extra=["=MIN(C{r},D{r})", "=MAX(C{r}-D{r},0)", "=MAX(D{r}-C{r},0)", "=100"])

    # --- the bands: stacked area, invisible base first -----------------------
    area = AreaChart()
    area.grouping = "stacked"
    area.add_data(Reference(ws, min_col=5, max_col=7, min_row=HDR, max_row=LAST),
                  titles_from_data=True)
    area.set_categories(years_ref(ws))
    off = dict(showCatName=False, showLegendKey=False,
               showPercent=False, showBubbleSize=False)
    for series, fill, at in zip(area.series, (None, PF_AGAINST, PF_FAVOUR),
                                (None, 1737 - 1700, 1768 - 1700)):
        series.graphicalProperties = GraphicalProperties(
            noFill=True) if fill is None else GraphicalProperties(
            solidFill=fill, ln=LineProperties(noFill=True))
        if at is None:
            continue
        # dLblPos is deliberately NOT set: an area chart accepts only "ctr", and any
        # other value is a repair prompt. Omit it and Excel centres the label itself.
        lbl = DataLabel(idx=at, showSerName=True, showVal=False, **off)
        lbl.txPr = rich(11, PF_FRAME, bold=True, italic=True)
        series.dLbls = DataLabelList(dLbl=[lbl], showSerName=False, showVal=False, **off)

    # --- the two lines, on the same primary axes -----------------------------
    line = LineChart()
    # The heavy 100,000 rule goes in FIRST so it draws under the two curves - a
    # horizontal reference line is just a series of one constant value.
    line.add_data(Reference(ws, min_col=8, min_row=HDR, max_row=LAST),
                  titles_from_data=True)
    line.add_data(Reference(ws, min_col=3, max_col=4, min_row=HDR, max_row=LAST),
                  titles_from_data=True)
    line.set_categories(years_ref(ws))
    for series, colour, width, at in (
            (line.series[0], PF_FRAME, int(1.75 * PT), None),
            (line.series[1], PF_IMPORT, W_BOLD, 1722 - 1700),
            (line.series[2], PF_EXPORT, W_BOLD, 1730 - 1700)):
        series.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=colour, w=width))
        series.marker = Marker(symbol="none")
        series.smooth = False
        if at is None:
            # "100,000" at the right-hand end, spelled out as Playfair spelled it.
            # dLblPos="r" put this on top of the axis's own "100" tick label.
            # "l" tucks it inside the plot at the right-hand end instead.
            lbl = DataLabel(idx=N - 1, showVal=True, showSerName=False, dLblPos="l", **off)
            lbl.numFmt = '#,##0",000"'
        else:
            lbl = DataLabel(idx=at, showSerName=True, showVal=False, dLblPos="t", **off)
        lbl.txPr = rich(10, PF_FRAME, italic=True)
        # showVal=False at LIST level is load-bearing: leave it unset and Excel labels
        # all 81 points, not the one dLbl entry you asked for.
        series.dLbls = DataLabelList(dLbl=[lbl], showSerName=False, showVal=False, **off)
    # Sub-charts keep openpyxl's default axIds (10 / 100), so both draw against the
    # primary pair. Change one and Excel invents a secondary axis.
    area += line

    area.title = chart_title(
        "Exports and Imports to and from DENMARK & NORWAY from 1700 to 1780",
        size_pt=12, colour=PF_FRAME, bold=False, left=False)
    area.width, area.height = 24.0, 15.5
    area.legend = None
    area.roundedCorners = False
    area.style = None
    area.graphical_properties = GraphicalProperties(
        solidFill=SURFACE, ln=LineProperties(solidFill=PF_FRAME, w=W_FRAME))
    # Playfair boxed the plot itself, not just the sheet of paper.
    area.plot_area.graphicalProperties = GraphicalProperties(
        solidFill=SURFACE, ln=LineProperties(solidFill=PF_FRAME, w=int(1.25 * PT)))

    y = area.y_axis
    y.delete = False
    y.scaling.min, y.scaling.max = 0, 200
    y.majorUnit, y.minorUnit = 10, 10
    y.numFmt = "0"
    y.txPr = rich(8, PF_FRAME)
    y.majorTickMark = y.minorTickMark = "none"
    # The plate carries its scale on the RIGHT: "the Right hand line into L10,000 each".
    y.tickLblPos = "high"
    y.majorGridlines.spPr = GraphicalProperties(
        ln=LineProperties(solidFill=PF_RULE, w=W_HAIR))
    y.spPr = GraphicalProperties(ln=LineProperties(solidFill=PF_FRAME, w=W_HAIR))

    x = area.x_axis
    x.delete = False
    x.tickLblSkip = x.tickMarkSkip = 10
    x.txPr = rich(9, PF_FRAME)
    x.majorTickMark = x.minorTickMark = "none"
    # A category axis has no majorGridlines object at all until you make one; the
    # value axis gets one by default. Playfair ruled the plate both ways.
    x.majorGridlines = ChartLines()
    x.majorGridlines.spPr = GraphicalProperties(
        ln=LineProperties(solidFill=PF_RULE, w=W_HAIR))
    x.spPr = GraphicalProperties(ln=LineProperties(solidFill=PF_FRAME, w=W_HAIR))

    ws.add_chart(area, "I2")
    ws["I34"] = ("Playfair's own caption: \"The Bottom line is divided into Years, "
                 "the Right hand line into L10,000 each.\"")
    ws["I34"].font = Font(name=FONT, size=9, italic=True, color=PF_FRAME)
    ws["I35"] = ("Pink band = BALANCE AGAINST England (imports exceed exports). "
                 "Buff band = BALANCE IN FAVOUR OF ENGLAND.")
    ws["I35"].font = Font(name=FONT, size=9, italic=True, color=PF_FRAME)
    source_at(ws, "I37")


# ---------------------------------------------------------------------------
# Sheet 2 — the house-style version
# ---------------------------------------------------------------------------
def event_line(ws, row, caption, x_index, top=1.0, col=9):
    """
    A vertical reference line built from data, not drawn as a shape. Two points at
    the same x, 0 and `top`, on a hidden 0-1 secondary axis pair. The caption is the
    SERIES NAME shown as the top point's data label, so the words travel with the line.
    This is the one licensed secondary axis: it encodes nothing and it is hidden.
    """
    letter = get_column_letter(col)
    ws.cell(row, col, caption).font = Font(name=FONT, size=9, italic=True, color=TEXT_2)
    for k, yv in enumerate((0, top)):
        ws.cell(row + k, col + 1, x_index).font = Font(name=FONT, size=9, color=TEXT_2)
        ws.cell(row + k, col + 2, yv).font = Font(name=FONT, size=9, color=TEXT_2)

    s = Series(Reference(ws, min_col=col + 2, min_row=row, max_row=row + 1),
               Reference(ws, min_col=col + 1, min_row=row, max_row=row + 1))
    # Series(title=...) calls str() on its argument, so a SeriesLabel would serialise
    # as "<openpyxl...object>". Attach it afterwards.
    s.tx = SeriesLabel(strRef=StrRef(f"'{ws.title}'!${letter}${row}"))
    s.graphicalProperties = GraphicalProperties(
        ln=LineProperties(solidFill=ALERT, w=W_HAIR, prstDash="dash"))
    s.marker = Marker(symbol="none")

    off = dict(showVal=False, showCatName=False, showLegendKey=False,
               showPercent=False, showBubbleSize=False)
    top_label = DataLabel(idx=1, showSerName=True, dLblPos="t", **off)
    top_label.txPr = rich(9, ALERT)
    s.dLbls = DataLabelList(dLbl=[DataLabel(idx=0, showSerName=False, **off), top_label],
                            showSerName=False, **off)

    sc = ScatterChart()
    sc.series.append(s)
    # The secondary pair must sit OPPOSITE the primary pair and cross at max. Left at
    # openpyxl's defaults it lands on the primary axes and Excel rescales the REAL data
    # to 0-1, flattening it into a hairline.
    sc.x_axis.axId, sc.y_axis.axId = 300, 200
    sc.x_axis.crossAx, sc.y_axis.crossAx = 200, 300
    sc.x_axis.axPos, sc.y_axis.axPos = "t", "r"
    sc.x_axis.crosses = sc.y_axis.crosses = "max"
    sc.x_axis.scaling.min, sc.x_axis.scaling.max = 0.5, N + 0.5
    sc.y_axis.scaling.min, sc.y_axis.scaling.max = 0, 1
    sc.x_axis.majorGridlines = sc.y_axis.majorGridlines = None
    for ax in (sc.x_axis, sc.y_axis):
        # Hidden, NOT deleted: a deleted secondary axis makes Excel reassign the series
        # to the primary axis, which then rescales to 0-1.
        ax.delete = False
        ax.majorTickMark = ax.minorTickMark = "none"
        # tickLblPos="none" is read by openpyxl as "unset" and dropped. ";;;" is
        # Excel's own show-nothing format and does the job.
        ax.numFmt = ";;;"
        ax.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    return sc


def build_balance(ws):
    ws.sheet_view.showGridLines = False
    header(ws, f"England's trade with Denmark & Norway turned from deficit to "
               f"surplus after {TURN_YEAR}",
           "Annual exports and imports, thousands of pounds sterling, 1700-1780. "
           "The value axis runs 0-200 and is not truncated, so the slopes are honest.")
    source_at(ws, "B5")
    note(ws, "B6",
         "I9:K10 is the event-line helper block: two points at the same x, one at 0 and "
         "one at 1, drawn as a scatter line on a hidden 0-1 secondary axis. The caption "
         "is the series NAME, shown as the top point's data label.")

    live_table(ws, ["Year", "Imports", "Exports"])

    chart = LineChart()
    chart.add_data(Reference(ws, min_col=3, max_col=4, min_row=HDR, max_row=LAST),
                   titles_from_data=True)
    chart.set_categories(years_ref(ws))
    chart.title = chart_title(f"Exports overtook imports in {TURN_YEAR} and never "
                              f"looked back", size_pt=13)
    chart.width, chart.height = 22.0, 11.0
    chart.legend = None                     # both lines are labelled at their own end
    chart.roundedCorners = False
    chart.style = None
    chart.graphical_properties = GraphicalProperties(
        solidFill=SURFACE, ln=LineProperties(noFill=True))
    chart.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.05, y=0.16, w=0.79, h=0.72))

    # Grey is the default; the accent marks the one thing the title is about.
    for series, colour, width in ((chart.series[0], CONTEXT_LINE, W_THIN),
                                  (chart.series[1], ACCENT, W_BOLD)):
        series.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=colour, w=width))
        series.marker = Marker(symbol="none")
        series.smooth = False
        end_label(series, N - 1, colour, fmt="#,##0")

    declutter_value_axis(chart.y_axis)
    chart.y_axis.scaling.min, chart.y_axis.scaling.max = 0, 200
    chart.y_axis.majorUnit = 50
    chart.y_axis.title = axis_title("Thousands of pounds sterling")
    declutter_category_axis(chart.x_axis, skip=10)

    # top=0.88 parks the caption in the empty upper-left of the plot. At 0.62 it sat
    # on the rising exports line - only the render shows that.
    chart += event_line(ws, R0, f"Balance turns in England's favour, {TURN_YEAR}",
                        IDX_TURN, top=0.88)

    ws.add_chart(chart, "M2")
    source_at(ws, "M26")
    for col in "IJK":
        ws.column_dimensions[col].width = 12


# ---------------------------------------------------------------------------
# Sheet 3 — the balance as a deviation chart
# ---------------------------------------------------------------------------
def build_deviation(ws):
    ws.sheet_view.showGridLines = False
    header(ws, "Fifty-five years of deficit, then twenty-five of surplus",
           "Exports minus imports, thousands of pounds sterling. Midpoint is zero: "
           "teal is a balance in England's favour, claret a balance against. "
           "Bars carry signs, so the chart still reads in greyscale.")
    source_at(ws, "B5")
    note(ws, "B6",
         "Every bar carries invertIfNegative = False. Without it Excel draws the negative "
         "half hollow - white with an outline - and ignores the fill entirely.")

    live_table(ws, ["Year", "Imports", "Exports", "Balance"],
               extra=["=D{r}-C{r}"])

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.gapWidth = 40                     # section 4: 40% for a single series
    chart.add_data(Reference(ws, min_col=5, min_row=HDR, max_row=LAST),
                   titles_from_data=True)
    chart.set_categories(years_ref(ws))
    # Numbers follow style_chart.md section 5 (dot thousands, U+2212 minus). The
    # currency mark is GBP, not the guide's dong, for the obvious reason - see README.
    chart.title = chart_title("A £38.000 deficit in 1700 became a £94.000 "
                              "surplus in 1780", size_pt=13)
    chart.width, chart.height = 22.0, 10.0
    chart.legend = None
    chart.roundedCorners = False
    chart.style = None
    chart.graphical_properties = GraphicalProperties(
        solidFill=SURFACE, ln=LineProperties(noFill=True))
    chart.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.06, y=0.18, w=0.90, h=0.70))

    series = chart.series[0]
    series.invertIfNegative = False          # or the negative half renders hollow
    series.graphicalProperties = GraphicalProperties(
        solidFill=BAD, ln=LineProperties(noFill=True))
    # Colour by favourability, not by direction (rule 10). One DataPoint per year.
    series.data_points = [
        DataPoint(idx=i, invertIfNegative=False,
                  spPr=GraphicalProperties(solidFill=GOOD if exp >= imp else BAD,
                                           ln=LineProperties(noFill=True)))
        for i, (_, imp, exp) in enumerate(DATA)]

    # Signed labels on three points: teal and claret sit only 12.3 L* apart, so
    # direction must never depend on colour alone (guide section 2.6).
    off = dict(showCatName=False, showLegendKey=False, showPercent=False,
               showBubbleSize=False, showSerName=False)
    marks = []
    for idx, colour in ((0, BAD), (IDX_TURN - 1, TEXT_2), (N - 1, GOOD)):
        lbl = DataLabel(idx=idx, showVal=True,
                        dLblPos="inEnd" if idx == IDX_TURN - 1 else "outEnd", **off)
        lbl.txPr = rich(11, colour, bold=True)
        lbl.numFmt = '+#,##0;"−"#,##0;0'
        marks.append(lbl)
    series.dLbls = DataLabelList(dLbl=marks, showVal=False, **off)

    declutter_value_axis(chart.y_axis, fmt='#,##0;"−"#,##0')
    chart.y_axis.majorUnit = 50
    # One line, or Excel wraps it onto the top tick label. The definition lives in
    # the worksheet subtitle, which has room for it.
    chart.y_axis.title = axis_title("Balance, GBP '000")
    declutter_category_axis(chart.x_axis, skip=10)
    # With negatives present the zero line is drawn and labelled (section 4).
    chart.x_axis.spPr = GraphicalProperties(
        ln=LineProperties(solidFill=TEXT_2, w=int(1.0 * PT)))
    chart.x_axis.tickLblPos = "low"          # keep year labels below the plot, not on zero

    ws.add_chart(chart, "H2")
    source_at(ws, "H24")


# ---------------------------------------------------------------------------
def fix_label_number_formats(path):
    """
    Give every data-label number format an explicit sourceLinked="0".

    openpyxl writes <numFmt formatCode="..."/> and stops. Excel reads a missing
    sourceLinked as TRUE, uses the CELL's format instead, and the label format does
    nothing - silently. Passing a NumFmt object does not help; the descriptor calls
    str() on it. Patching the saved package is the only fix.
    """
    tmp = path.with_suffix(".patching.xlsx")
    n = 0
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(
            tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            blob = src.read(item.filename)
            if re.match(r"xl/charts/chart\d+\.xml$", item.filename):
                text = blob.decode("utf-8")
                patched = re.sub(r'<numFmt formatCode="([^"]*)"/>',
                                 r'<numFmt formatCode="\1" sourceLinked="0"/>', text)
                if patched != text:
                    blob = patched.encode("utf-8")
                    n += 1
            dst.writestr(item, blob)
    shutil.move(tmp, path)
    return n


def main():
    wb = Workbook()
    wb.remove(wb.active)
    write_data_sheet(wb)
    build_replica(wb.create_sheet("Playfair_1786"))
    build_balance(wb.create_sheet("Balance"))
    build_deviation(wb.create_sheet("Deviation"))
    wb.move_sheet("Data_Trade", offset=3)
    out = HERE / "playfair_trade_balance.xlsx"
    wb.save(out)
    print(f"wrote {out}  (label formats patched in {fix_label_number_formats(out)} charts)")


if __name__ == "__main__":
    main()
