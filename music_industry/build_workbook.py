"""
Raw data -> Data visualisation -> Data storytelling, on one dataset.

Three sheets, the same numbers on each, and the only thing that changes is how
much work the reader has to do:

    1 Raw data       the table. Every value present, no claim made.
    2 Visualisation  a stacked column chart. The shape is now visible; the
                     reader still has to decide what it means.
    3 Storytelling   the same data as an area chart, with the takeaway in the
                     title, the crossover marked, the bands labelled where they
                     are widest, and the two totals the title argues from
                     called out. The chart now argues something.

Data: RIAA U.S. Music Revenue Database, revenues adjusted for inflation to
2019 dollars, $ millions. See us_music_revenue.csv (built by prepare_data.py).

Sync is FOLDED INTO DIGITAL rather than carried as its own band. It is only
~2.5% of the total, but RIAA's headline figures include it: drop it and the two
callouts on sheet 3 read $22.4B and $10.8B rather than the $22.4B / $11.1B RIAA
publishes. Folding keeps it in the stack, so the totals still match, and two
bands read faster than three. Note this does overstate "digital" slightly —
synchronisation is licensing for film, TV and advertising, not a digital format.

    python build_workbook.py    ->  US_Music_Revenue.xlsx
"""

import csv
import re
import shutil
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import (AreaChart, BarChart, LineChart, Reference,
                            Series)
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (CharacterProperties, Font as DFont, Paragraph,
                                   ParagraphProperties, RegularTextRun)
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE / "style"))
import xl_style as xs  # noqa: E402

TEAL = xs.GOOD                      # digital — the series the story is about
DARK = xs.INK                       # physical — the one it replaced
GREY = xs.CONTEXT_FILL              # sync — present, and deliberately quiet
# Outside the guide palette, and deliberately: --alert claret (#990F3D) reads
# as maroon at 1px on a projector. 4.35:1 on white, so it still clears the 3:1
# floor for marks.
EVENT_RED = "E8352B"
SOURCE = ("Source: RIAA U.S. Music Revenue Database · revenues adjusted for "
          "inflation to 2019 dollars")
ADOPT_SOURCE = ("Adoption: World Bank / ITU, Individuals using the Internet "
                "(% of US population); Pew Research Center Mobile Fact Sheet, "
                "% of US adults owning a smartphone (from 2011)")
# The adoption lines ride over the dark navy and teal bands, so they need
# hues that hold against both those fills AND white plot background.
NET_AMBER, PHONE_VIOLET = "F2A900", "9B59B6"

# Where the table sits on every sheet. One layout, three sheets.
HDR = 8                             # the column-header row
COL_YEAR, COL_PHYS, COL_DIG, COL_TOT = 2, 3, 4, 5
COL_NET, COL_PHONE = 6, 7           # adoption %, on Data and sheet 4 only
AXIS_B = '$#,##0,"B"'
Y_MAX = 25000                       # $MM — the ceiling both charts share


def load_rows():
    with open(HERE / "us_music_revenue.csv", encoding="utf-8") as fh:
        return [(int(r["Year"]), int(r["Physical"]),
                 int(r["Digital"]) + int(r["Synchronization"]))
                for r in csv.DictReader(fh)]


def load_adoption():
    """
    US internet and smartphone adoption, to sit behind the revenue story.

    Smartphone years before 2011 are BLANK, not zero: Pew's series does not go
    back further, and a zero would draw twenty-one years of flat line along the
    floor for something that was never measured.
    """
    with open(HERE / "us_adoption.csv", encoding="utf-8") as fh:
        return {int(r["Year"]): (float(r["InternetPct"]),
                                 float(r["SmartphonePct"]) if r["SmartphonePct"]
                                 else None)
                for r in csv.DictReader(fh)}


def write_data_sheet(wb, rows):
    """The one place the numbers live. Every other sheet formulas off this."""
    ws = wb.create_sheet("Data")
    xs.write_header(ws, 2, "Source data",
                    "RIAA U.S. music revenue by format, $ millions, "
                    "adjusted to 2019 dollars")
    xs.write_table_header(ws, HDR, ["Year", "Physical", "Digital", "Total",
                                    "Internet %", "Smartphone %"])
    adoption = load_adoption()
    for i, (year, phys, dig) in enumerate(rows):
        r = HDR + 1 + i
        ws.cell(r, COL_YEAR, year).number_format = "0"
        for col, val in ((COL_PHYS, phys), (COL_DIG, dig)):
            ws.cell(r, col, val).number_format = "#,##0"
        ws.cell(r, COL_TOT, "=C%d+D%d" % (r, r)).number_format = "#,##0"
        net, phone = adoption.get(year, (None, None))
        ws.cell(r, COL_NET, net).number_format = "0.0"
        if phone is not None:            # blank, not zero, before 2011
            ws.cell(r, COL_PHONE, phone).number_format = "0"
    xs.write_source(ws, HDR + len(rows) + 2, SOURCE)
    xs.note(ws, HDR + len(rows) + 3, ADOPT_SOURCE)
    xs.widths(ws, [10, 12, 12, 12, 12, 13])
    return ws


def link_table(ws, n, adoption=False, phone_start=None):
    """Mirror the Data sheet into this sheet so the chart reads local cells."""
    heads = ["Year", "Physical", "Digital", "Total"]
    cols = [(COL_PHYS, "C", "#,##0"), (COL_DIG, "D", "#,##0"),
            (COL_TOT, "E", "#,##0")]
    if adoption:
        heads += ["Internet %", "Smartphone %"]
        cols += [(COL_NET, "F", "0.0"), (COL_PHONE, "G", "0")]
    xs.write_table_header(ws, HDR, heads)
    for i in range(n):
        r = HDR + 1 + i
        ws.cell(r, COL_YEAR, "=Data!B%d" % r).number_format = "0"
        for col, src, fmt in cols:
            # Years with no smartphone reading get NO CELL AT ALL. A formula
            # returning "" is not a blank — Excel plots it as zero and the line
            # dives to the floor for every year before 2011. dispBlanksAs="gap"
            # only honours genuinely empty cells.
            if col == COL_PHONE and phone_start is not None and i < phone_start:
                continue
            ws.cell(r, col, "=Data!%s%d" % (src, r)).number_format = fmt
    xs.widths(ws, [10, 12, 12, 12] + ([12, 13] if adoption else []))


def series_refs(ws, n):
    data = Reference(ws, min_col=COL_PHYS, max_col=COL_DIG,
                     min_row=HDR, max_row=HDR + n)
    cats = Reference(ws, min_col=COL_YEAR, min_row=HDR + 1, max_row=HDR + n)
    return data, cats


def paint_bands(ch):
    xs.fill(ch.series[0], DARK)
    xs.fill(ch.series[1], TEAL)


# ---------------------------------------------------------------- sheet 1
def sheet_raw_data(wb, rows):
    ws = wb.create_sheet("1 Raw data")
    xs.write_header(ws, 2, "Raw data",
                    "Everything is here, and nothing is said. Thirty rows is "
                    "already past what a reader will scan.")

    ws.merge_cells(start_row=HDR - 1, start_column=COL_YEAR,
                   end_row=HDR - 1, end_column=COL_TOT)
    banner = ws.cell(HDR - 1, COL_YEAR, "REVENUE ($MM)")
    banner.fill = PatternFill("solid", fgColor=xs.INK)
    banner.font = Font(name=xs.FONT, size=11, bold=True, color="FFFFFF")
    banner.alignment = Alignment(horizontal="center", vertical="center")

    link_table(ws, len(rows))
    for i in range(len(rows)):          # banding, so the eye can hold a row
        if i % 2:
            for col in range(COL_YEAR, COL_TOT + 1):
                ws.cell(HDR + 1 + i, col).fill = PatternFill(
                    "solid", fgColor=xs.WASH)
    xs.write_source(ws, HDR + len(rows) + 2, SOURCE)
    return ws


# ---------------------------------------------------------------- sheet 2
def sheet_visualisation(wb, rows):
    n = len(rows)
    ws = wb.create_sheet("2 Visualisation")
    xs.write_header(ws, 2, "U.S. Music Revenue",
                    "Physical and digital revenue, $ billions, adjusted "
                    "to 2019 dollars. The shape is visible — the point is not "
                    "yet made.")
    link_table(ws, n)

    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "col", "stacked", 100
    data, cats = series_refs(ws, n)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    paint_bands(ch)

    xs.declutter(ch, gap=30, num_fmt=AXIS_B)
    xs.legend_bottom(ch)
    ch.title = xs.chart_title("U.S. Music Revenue")
    # 30 categories will not fit as 30 labels; Excel's answer is to rotate
    # them, which the guide forbids. Label a decade at a time instead.
    ch.x_axis.tickLblSkip = ch.x_axis.tickMarkSkip = 10
    ch.y_axis.scaling.max, ch.y_axis.majorUnit = Y_MAX, 5000
    ch.width, ch.height = 24.0, 11.5
    ws.add_chart(ch, "H8")

    xs.write_source(ws, HDR + n + 2, SOURCE)
    return ws


# ---------------------------------------------------------------- sheet 3
def two_tone_title(lead, lead_colour, rest, subtitle_runs, size=18, sub_size=14):
    """
    A chart title whose opening phrase is the colour of the band it names.

    xl_style.chart_message writes one colour per line. Here the first two words
    ARE the series, so colouring them is the same direct-labelling move as
    naming a band inside itself — applied to the sentence instead of the shape.
    """
    def cp(pt, colour, bold):
        return CharacterProperties(sz=int(pt * 100), b=bold, solidFill=colour,
                                   latin=DFont(typeface=xs.FONT))

    lead_cp, rest_cp = cp(size, lead_colour, True), cp(size, xs.INK, True)
    sub_cp = cp(sub_size, xs.TEXT_2, False)
    paras = [
        Paragraph(pPr=ParagraphProperties(algn="l", defRPr=rest_cp),
                  r=[RegularTextRun(rPr=lead_cp, t=lead),
                     RegularTextRun(rPr=rest_cp, t=rest)]),
        Paragraph(pPr=ParagraphProperties(algn="l", defRPr=sub_cp),
                  r=[RegularTextRun(rPr=cp(sub_size, xs.TEXT_2, bold), t=text)
                     for text, bold in subtitle_runs]),
    ]
    # algn="l" alone does not hold: Excel re-centres a chart title over the
    # plot area regardless. Pinning the title block to the left edge with a
    # manual layout is the only thing that makes §3 (titles left-aligned)
    # actually stick.
    return Title(tx=Text(rich=RichText(p=paras)), overlay=False,
                 layout=Layout(manualLayout=ManualLayout(
                     xMode="edge", yMode="edge", x=0.02, y=0.015)))


def chart_captions(ws, row, scatter, items, col=2):
    """
    Print a stack total at a chosen year, on the chart, as text.

    openpyxl's DataLabel has no rich-text element, so a label can only show a
    value or a series name — and the value it would show is the top band's own
    number, not the stack total. So each callout is a one-point scatter series
    NAMED after the total, riding the hidden secondary axis that event_lines
    already set up, with that name as its only label. The name is a formula
    over the total column, so the caption cannot drift from the data.

    items   [(caption_formula, x_index, y, dLblPos, colour, size, bold)] —
            x_index is 1-based over the categories, matching event_lines; y is
            a fraction of the plot height (0 = the baseline, 1 = the top).
    """
    letter = get_column_letter(col)
    r = row
    for formula, x_index, y, pos, colour, size, bold in items:
        cap = ws.cell(r, col, formula)
        cap.font = Font(name=xs.FONT, size=9, color=xs.TEXT_2, italic=True)
        ws.cell(r, col + 1, x_index).font = Font(name=xs.FONT, size=9)
        ws.cell(r, col + 2, y).font = Font(name=xs.FONT, size=9)

        s = Series(Reference(ws, min_col=col + 2, min_row=r, max_row=r),
                   Reference(ws, min_col=col + 1, min_row=r, max_row=r))
        # Series(title=...) str()s whatever it is given, so the reference is
        # attached afterwards or the caption becomes a repr.
        s.tx = SeriesLabel(strRef=StrRef("'%s'!$%s$%d" % (ws.title, letter, r)))
        s.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
        s.marker = Marker(symbol="none")

        off = dict(showVal=False, showCatName=False, showLegendKey=False,
                   showPercent=False, showBubbleSize=False)
        lbl = DataLabel(idx=0, showSerName=True, dLblPos=pos, **off)
        lbl.txPr = xs._rich(size, colour, bold=bold)
        s.dLbls = DataLabelList(dLbl=[lbl], showSerName=False, **off)
        scatter.series.append(s)
        r += 1
    return r


def sheet_storytelling(wb, rows):
    n = len(rows)
    ws = wb.create_sheet("3 Storytelling")
    xs.write_header(
        ws, 2, "Digital revenue climbed steadily with more internet adoption",
        "However, the music industry shrunk 50% comparing to 1999")
    link_table(ws, n)
    # The area bands are named from these header cells, and inside a filled
    # band a set-caps name reads as a region label rather than as data. Only
    # this sheet: sheet 2's legend keeps sentence case.
    for col, name in ((COL_PHYS, "PHYSICAL"), (COL_DIG, "DIGITAL")):
        ws.cell(HDR, col, name)

    ch = AreaChart()
    ch.grouping = "stacked"
    data, cats = series_refs(ws, n)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    paint_bands(ch)

    # Name each band inside itself, at the point where it is widest — the
    # legend on sheet 2 costs a glance down and back, this costs nothing.
    xs.label_point(ch.series[0], 7, "FFFFFF", size=12, position=None)
    xs.label_point(ch.series[1], 27, "FFFFFF", size=12, position=None)

    xs.declutter(ch, gap=None, num_fmt=AXIS_B, gridlines=False)
    xs.no_legend(ch)
    # as on sheet 2: 30 labels do not fit, and a rotated label is not an
    # acceptable answer. The iTunes caption carries the year that matters.
    ch.x_axis.tickLblSkip = ch.x_axis.tickMarkSkip = 10
    ch.y_axis.scaling.max, ch.y_axis.majorUnit = Y_MAX, 5000
    # no y-axis title: the tick labels already read "$25B", so naming the axis
    # only repeats the units and steals width from the plot.
    ch.title = two_tone_title(
        "Digital revenue", TEAL,
        " climbed steadily with more internet adoption",
        [("However, the music industry ", False), ("shrunk 50%", True),
         (" comparing to 1999", False)])

    # The iTunes line and the two totals are built INTO the chart rather than
    # drawn on top of it, and share one hidden secondary axis pair — a second
    # pair is not available, so the callouts are appended to this scatter.
    # The line stops at $20B rather than running the full height: it is a
    # marker on the timeline, not a series, and a full-height rule reads as a
    # boundary splitting the chart in two.
    events = [(2003 - 1990 + 1, "iTunes Store opens (2003)", 20000 / Y_MAX)]
    scatter, next_row = xs.event_lines(ws, HDR + n + 4, events, n,
                                       colour=EVENT_RED)
    # event_lines writes its caption at 9pt; that is right for a chart on a
    # worksheet and too small for one projected. The caption is the label on
    # the line's top point, so it is the dLbl carrying idx=1.
    for lbl in scatter.series[0].dLbls.dLbl:
        if lbl.idx == 1:
            lbl.txPr = xs._rich(11, EVENT_RED, bold=True)
            # "t" centres the caption over the line, which pushes its left half
            # back across the falling physical curve. "r" hangs it off the top
            # of the line instead, into the empty part of the plot.
            lbl.dLblPos = "r"
    # and thicker than the hairline event_lines draws, for the same reason
    scatter.series[0].graphicalProperties.ln.w = int(2.25 * xs.PT)
    r99, r19 = HDR + 1 + (1999 - 1990), HDR + 1 + (2019 - 1990)
    tot = get_column_letter(COL_TOT)
    money = '=TEXT(%s%%d/1000,"$0.0")&"B"' % tot
    frac = "=%s%%d/%%d" % tot
    next_row = chart_captions(ws, next_row + 1, scatter, [
        # the peak, and where it ended up. Both sit "t" — directly over their
        # own year. "l" hangs the label off to the LEFT of the point, which is
        # what made $11.1B look mis-registered against 2019.
        (money % r99, 1999 - 1990 + 1, frac % (r99, Y_MAX), "t",
         xs.INK, 11, True),
        (money % r19, 2019 - 1990 + 1, frac % (r19, Y_MAX), "t",
         xs.INK, 11, True),
        # tickLblSkip puts year labels on 1990/2000/2010 and there is no way to
        # ask Excel for an extra one, so the final year is written on the
        # baseline as its own caption.
        ('=TEXT(B%d,"0")' % r19, 2019 - 1990 + 1, 0, "b", xs.TEXT_2, 11, False),
    ])
    ch += scatter
    xs.hide_legend_entries(ch, [2, 3, 4, 5])

    ch.width, ch.height = 24.0, 11.5
    ws.add_chart(ch, "H8")

    xs.write_source(ws, next_row + 1, SOURCE)
    return ws


# ---------------------------------------------------------------- sheet 4
def sheet_analyst(wb, rows):
    """
    The same revenue, plus the reason for it, for a reader who wants detail.

    This is the analyst cut of sheet 3: nothing is simplified away and nothing
    is argued. The differences are all deliberate:

      * TWO VALUE AXES. Revenue is $ billions and adoption is a percentage;
        they share no scale, so one axis would either flatten the revenue bands
        or send the percentages off the top. The right axis is pinned 0-100 so
        the lines cannot be made steeper or flatter by their own range.
      * NO LEGEND. Every series names itself where it sits — the two bands from
        the inside, the two lines at their far ends. A legend for four series
        costs a glance down and back for every one of them.
      * A DESCRIPTIVE TITLE, not a takeaway. The analyst is here to form their
        own view; sheet 3 is where the chart argues.

    Correlation is not cause, and the chart does not claim it is: the lines are
    laid alongside the bands so the reader can judge the timing themselves.
    """
    n = len(rows)
    ws = wb.create_sheet("4 Analyst")
    xs.write_header(
        ws, 2, "Digital revenue climbed steadily with more internet adoption",
        "However, the music industry shrunk 50% comparing to 1999")
    adoption = load_adoption()
    phone_start = next(i for i, (y, _, _) in enumerate(rows)
                       if adoption[y][1] is not None)
    link_table(ws, n, adoption=True, phone_start=phone_start)
    # Set caps for the two band names only: inside a filled band a set-caps
    # name reads as a region label rather than as a data point. The line names
    # stay sentence case — they are labels on a line, not regions.
    for col, name in ((COL_PHYS, "PHYSICAL"), (COL_DIG, "DIGITAL")):
        ws.cell(HDR, col, name)

    ch = AreaChart()
    ch.grouping, ch.overlap = "stacked", 100
    # DIGITAL ON THE BASELINE, physical stacked above it — the reverse of
    # sheet 3. In a stacked area only the bottom band has a flat reference to
    # read against; every band above it is distorted by the ones below. This
    # chart is about digital's rise, so digital is the one that has to be
    # readable. The combined outline, and so both callout totals, is unchanged.
    _, cats = series_refs(ws, n)
    for col in (COL_DIG, COL_PHYS):
        ch.add_data(Reference(ws, min_col=col, min_row=HDR, max_row=HDR + n),
                    titles_from_data=True)
    ch.set_categories(cats)
    xs.fill(ch.series[0], TEAL)        # digital, now first = bottom
    xs.fill(ch.series[1], DARK)        # physical
    xs.declutter(ch, gap=None, num_fmt=AXIS_B, gridlines=False)
    ch.title = two_tone_title(
        "Digital revenue", TEAL,
        " climbed steadily with more internet adoption",
        [("However, the music industry ", False), ("shrunk 50%", True),
         (" comparing to 1999", False)])
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, Y_MAX
    ch.y_axis.majorUnit = 5000
    ch.x_axis.tickLblSkip = ch.x_axis.tickMarkSkip = 10
    # Each band named inside itself, where it is widest. series[0] is DIGITAL
    # here, so the indices are the other way round from sheet 3.
    xs.label_point(ch.series[0], 27, "FFFFFF", size=12, position=None)
    xs.label_point(ch.series[1], 7, "FFFFFF", size=12, position=None)

    # --- the second axis ------------------------------------------------
    line = LineChart()
    line.add_data(Reference(ws, min_col=COL_NET, max_col=COL_PHONE,
                            min_row=HDR, max_row=HDR + n),
                  titles_from_data=True)
    line.set_categories(cats)
    # Each line names itself AND states its 2019 value, which is what lets the
    # legend go entirely. openpyxl's DataLabel can show a series name or a value
    # but not both, so the name is a formula that already contains the number —
    # read live off the last row, so the caption cannot drift from the data.
    last = HDR + n
    cap_row = HDR + n + 5
    xs.note(ws, cap_row - 1,
            "Helper: the two line captions. Each is the series name for the "
            "chart, so the end label carries the 2019 value.")
    for i, (series, colour, pos, label, src) in enumerate((
            (line.series[0], NET_AMBER, "r", "Internet",
             get_column_letter(COL_NET)),
            (line.series[1], PHONE_VIOLET, "r", "Smartphone",
             get_column_letter(COL_PHONE)))):
        series.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=colour, w=int(2.25 * xs.PT)))
        series.marker = Marker(symbol="none")
        series.smooth = False
        r = cap_row + i
        ws.cell(r, COL_YEAR,
                '="%s "&TEXT(%s%d,"0")&"%%"' % (label, src, last)
                ).font = Font(name=xs.FONT, size=9, color=xs.TEXT_2, italic=True)
        series.tx = SeriesLabel(strRef=StrRef(
            "'%s'!$%s$%d" % (ws.title, get_column_letter(COL_YEAR), r)))
        xs.label_last_point(series, n, colour, size=11, position=pos)

    # A label for the final year. tickLblSkip=10 puts axis labels on 1990, 2000
    # and 2010, and there is no way to ask Excel for one more at 2019. Sheet 3
    # writes it as a scatter caption, but that hidden axis pair is already spent
    # here on the adoption lines. So: a third line series, blank everywhere
    # except a single 0 at 2019, invisible, carrying a data label that shows its
    # CATEGORY name — which is the year. It sits at 0 on the 0-100 axis, and 0
    # is the same baseline as the dollar axis, so the label lands on the axis.
    tick_col = COL_PHONE + 1
    ws.cell(HDR, tick_col, "Year tick")
    ws.cell(HDR + n, tick_col, 0)
    xs.note(ws, HDR + n + 7,
            "Helper: one zero at the last year, drawn invisibly, so its data "
            "label can print 2019 on the axis.")
    line.add_data(Reference(ws, min_col=tick_col, min_row=HDR, max_row=HDR + n),
                  titles_from_data=True)
    tick = line.series[2]
    tick.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    tick.marker = Marker(symbol="none")
    _off = dict(showVal=False, showSerName=False, showLegendKey=False,
                showPercent=False, showBubbleSize=False)
    _lbl = DataLabel(idx=n - 1, showCatName=True, dLblPos="b", **_off)
    _lbl.txPr = xs._rich(11, xs.TEXT_2)
    tick.dLbls = DataLabelList(dLbl=[_lbl], showCatName=False, **_off)

    # A REAL secondary axis pair needs its OWN category axis, not just its own
    # value axis. Give the line chart only a second valAx and Excel still draws
    # the right-hand scale — but the object model reports one axis group, every
    # series comes back AxisGroup=1, and Axes(xlValue, xlSecondary) throws. The
    # hidden second catAx is what makes it a genuine pair.
    line.y_axis.axId = 200
    line.x_axis.axId = 300
    line.x_axis.delete = True          # hidden: the primary one is the visible one
    line.y_axis.crossAx = 300
    line.x_axis.crossAx = 200
    line.y_axis.axPos = "r"
    line.y_axis.crosses = "max"
    line.y_axis.scaling.min, line.y_axis.scaling.max = 0, 100
    line.y_axis.majorUnit = 20
    line.y_axis.numFmt = line.y_axis.number_format = '0"%"'
    # The axis is HIDDEN, not deleted. delete=True on a secondary axis orphans
    # the series Excel had assigned to it, which re-scales them onto the primary
    # $ axis — 89% against a 25,000 ceiling lies flat on the floor. tickLblPos
    # "none" keeps the axis and its scale and simply draws nothing, which is
    # what the end labels replace.
    line.y_axis.delete = False
    line.y_axis.tickLblPos = "none"
    line.y_axis.majorGridlines = None
    line.y_axis.majorTickMark = line.y_axis.minorTickMark = "none"
    line.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    ch += line

    # Blanks must show as gaps or the smartphone line runs along zero from 1990.
    ch.dispBlanksAs = "gap"

    xs.no_legend(ch)
    # Reserve the right-hand strip for the two end labels. With the % axis
    # labels hidden Excel expands the plot to the chart edge, and a dLblPos="r"
    # label then runs straight off it and is clipped mid-word. A manual plot
    # area is the only way to hand that space back.
    ch.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.055, y=0.15, w=0.80, h=0.80))
    ch.width, ch.height = 24.0, 11.5
    ws.add_chart(ch, "J8")

    xs.write_source(ws, HDR + n + 2, SOURCE)
    xs.note(ws, HDR + n + 3, ADOPT_SOURCE)
    return ws


# ----------------------------------------------------------------- package
def pin_label_formats(path):
    """
    Give every data-label number format an explicit sourceLinked="0".

    openpyxl writes <numFmt formatCode="..."/> and stops. Excel reads a missing
    sourceLinked as TRUE, falls back to the cell's own format and silently
    ignores the one asked for. Same fix as hoa_gao_case/style/xl_chartex.py.
    """
    src = zipfile.ZipFile(path)
    parts = {nm: src.read(nm) for nm in src.namelist()}
    src.close()
    fixed = 0
    for name in list(parts):
        if re.match(r"xl/charts/chart\d+\.xml$", name):
            x = parts[name].decode()
            new = re.sub(r'<numFmt formatCode="([^"]*)"/>',
                         r'<numFmt formatCode="\1" sourceLinked="0"/>', x)
            if new != x:
                parts[name] = new.encode("utf-8")
                fixed += 1
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, blob in parts.items():
            out.writestr(name, blob)
    shutil.move(tmp, str(path))
    return fixed


def hide_tick_labels(path, ax_id=200):
    """
    Hide one axis's tick labels by writing tickLblPos into the saved package.

    openpyxl accepts `axis.tickLblPos = "none"` and then DROPS it on the way
    out — the attribute never reaches the XML, and Excel falls back to drawing
    the labels. Same class of silent loss as the numFmt fix above.

    Deleting the axis instead is not an option: delete=True orphans the series
    Excel had assigned to it and re-scales them onto the primary $ axis, where
    89% against a 25,000 ceiling lies flat on the floor.

    Schema order matters — CT_ValAx wants tickLblPos after minorTickMark and
    before spPr/txPr/crossAx, so it is spliced in at the first of those.
    """
    src = zipfile.ZipFile(path)
    parts = {nm: src.read(nm) for nm in src.namelist()}
    src.close()
    fixed = 0
    for name in list(parts):
        if not re.match(r"xl/charts/chart\d+\.xml$", name):
            continue
        x = parts[name].decode()
        out, pos, changed = [], 0, False
        for m in re.finditer(r"<valAx>.*?</valAx>", x, re.S):
            block = m.group(0)
            if 'axId val="%d"' % ax_id not in block:
                continue
            if "<tickLblPos" in block:
                continue
            for anchor in ("<spPr>", "<spPr ", "<txPr>", "<crossAx"):
                i = block.find(anchor)
                if i != -1:
                    new = block[:i] + '<tickLblPos val="none"/>' + block[i:]
                    out.append(x[pos:m.start()] + new)
                    pos = m.end()
                    changed = True
                    break
        if changed:
            parts[name] = ("".join(out) + x[pos:]).encode("utf-8")
            fixed += 1
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for name, blob in parts.items():
            out_zip.writestr(name, blob)
    shutil.move(tmp, str(path))
    return fixed


def nudge_data_label(path, marker, dx=0.0, dy=0.0):
    """
    Shift one data label by a fraction of the chart, written into the package.

    openpyxl's DataLabel has no `layout` element, so a caption can only sit at
    one of Excel's discrete positions — and here every one of them is wrong.
    "r" leaves it hard against its own event line; "t" centres it over the line,
    which drops its left half onto the dark PHYSICAL band where red on navy is
    unreadable. CT_DLbl allows a manualLayout offset, and with no xMode/yMode it
    defaults to FACTOR mode — an offset from the default position as a fraction
    of the chart, which is exactly the small nudge wanted.

    `marker` is a snippet unique to the target label's XML; its own txPr serves.
    Schema order: layout goes immediately after idx.
    """
    src = zipfile.ZipFile(path)
    parts = {nm: src.read(nm) for nm in src.namelist()}
    src.close()
    lay = ('<layout><manualLayout><x val="%s"/><y val="%s"/></manualLayout>'
           '</layout>' % (dx, dy))
    moved = 0
    for name in list(parts):
        if not re.match(r"xl/charts/chart\d+\.xml$", name):
            continue
        x = parts[name].decode()
        out, pos, changed = [], 0, False
        for m in re.finditer(r"<dLbl>.*?</dLbl>", x, re.S):
            block = m.group(0)
            if marker not in block or "<layout>" in block:
                continue
            i = block.find("/>", block.find("<idx ")) + 2
            out.append(x[pos:m.start()] + block[:i] + lay + block[i:])
            pos = m.end()
            changed = True
            moved += 1
        if changed:
            parts[name] = ("".join(out) + x[pos:]).encode("utf-8")
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for nm, blob in parts.items():
            out_zip.writestr(nm, blob)
    shutil.move(tmp, str(path))
    return moved


def main():
    rows = load_rows()
    wb = Workbook()
    wb.remove(wb.active)
    sheet_raw_data(wb, rows)
    sheet_visualisation(wb, rows)
    sheet_storytelling(wb, rows)
    sheet_analyst(wb, rows)
    write_data_sheet(wb, rows)

    out = HERE / "US_Music_Revenue.xlsx"
    wb.save(out)
    n = pin_label_formats(out)
    hidden = hide_tick_labels(out)
    # the iTunes caption, nudged left off its own event line
    # dx/dy are fractions of the chart, and OOXML's y runs DOWNWARD, so up
    # is negative. ~2 characters right of the previous nudge, ~5px up.
    nudged = nudge_data_label(out, 'val="%s"' % EVENT_RED,
                              dx=-0.011, dy=-0.005)
    print("  secondary tick labels hidden on %d chart(s)" % hidden)
    print("  data labels nudged: %d" % nudged)
    print("wrote %s  (%d rows, label formats pinned on %d charts)"
          % (out, len(rows), n))


if __name__ == "__main__":
    main()
