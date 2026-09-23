"""
Daily makeup use, 2024 vs 2019, as a real Excel chart.

Every category fell, and that is the whole point, so the two years OVERLAP
rather than stand side by side in a cluster. The grey 2019 bar is drawn behind
and the claret 2024 bar on top of it, offset a little, so the drop reads as the
exposed grey tail on the right — one gap per row to measure instead of two bar
lengths to compare across a gutter.

    overlap 82   how much of the pair shares a slot. 100 would make them
                 concentric and lose the offset; 0 puts them side by side and
                 turns the drop back into a two-bar comparison.
    gapWidth 55  space between category slots, tighter than Excel's 150 so the
                 rows group rather than float.

No legend: the years are named in the title, each in the colour of its own bar,
which is the same direct-labelling move applied to a sentence.

Values are transcribed from a supplied chart image — read twice, once off the
printed labels and once by measuring each bar's pixel length, which agreed. The
originating survey is not known here; SOURCE is left empty rather than filled
with a guess.

    python build_xlsx.py    ->  daily_makeup_use.xlsx
"""

import csv
import re
import shutil
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (CharacterProperties, Font as DFont, Paragraph,
                                   ParagraphProperties, RegularTextRun)
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "style"))
import xl_style as xs  # noqa: E402

OUT = HERE / "daily_makeup_use.xlsx"

CLARET = "941D52"          # 2024 — sampled from the reference image
GREY = "828282"            # 2019
FACE = xs.FONT             # Aptos, the deck face — not the reference's serif
SOURCE = ""                # unknown — see the module docstring

HDR = 8
# Slopegraph: which items get a LEFT label. 42/41/42/39 cannot all be read.
LEFT_LABELLED = {0, 1, 2, 5}
# Dot-plot helper columns: rails and connectors as gapped x/y pairs.
COL_RAIL_X, COL_RAIL_Y, COL_CONN_X, COL_CONN_Y = 7, 8, 10, 11
COL_TYPE, COL_24, COL_19, COL_CHG = 2, 3, 4, 5
X_MAX = 80


def load():
    with open(HERE / "daily_makeup_use.csv", encoding="utf-8") as fh:
        return [(r["Type"], int(r["Pct2024"]), int(r["Pct2019"]))
                for r in csv.DictReader(fh)]


def cp(pt, colour, bold=False):
    return CharacterProperties(sz=int(pt * 100), b=bold, solidFill=colour,
                               latin=DFont(typeface=FACE))


YEAR_24, YEAR_19 = ("2024", CLARET, True), ("2019", GREY, True)
LEAD = ("Daily makeup use: ", xs.INK, False)
VS = (" vs. ", xs.INK, False)


def two_tone_title(runs=None, subtitle="TYPE  |  % REPORTING DAILY USE"):
    """
    The title carries the key: each year is set in the colour of its own bars.

    That is what lets the chart drop its legend — the reader meets the mapping
    in the sentence they are already reading, instead of looking it up. The run
    order follows the bar order, so on the column chart, where 2019 stands to
    the left of 2024, the title names them in that order too.
    """
    runs = runs or [LEAD, YEAR_24, VS, YEAR_19]
    line = Paragraph(
        pPr=ParagraphProperties(algn="l", defRPr=cp(18, xs.INK, True)),
        r=[RegularTextRun(rPr=cp(18, c, b), t=t) for t, c, b in runs])
    sub = cp(11, xs.TEXT_2)
    sub_line = Paragraph(
        pPr=ParagraphProperties(algn="l", defRPr=sub),
        r=[RegularTextRun(rPr=sub, t=subtitle)])
    # algn="l" alone loses: Excel re-centres a chart title over the plot area
    # regardless, so the block is pinned to the left edge by hand.
    return Title(tx=Text(rich=RichText(p=[line, sub_line])), overlay=False,
                 layout=Layout(manualLayout=ManualLayout(
                     xMode="edge", yMode="edge", x=0.01, y=0.02)))


def value_labels(series, size):
    """The number rides inside the end of its own bar, in white."""
    off = dict(showSerName=False, showCatName=False, showLegendKey=False,
               showPercent=False, showBubbleSize=False)
    dl = DataLabelList(showVal=True, dLblPos="inEnd", **off)
    dl.txPr = xs._rich(size, "FFFFFF", bold=True)
    dl.numFmt = '0"%"'
    series.dLbls = dl


def pin_label_formats(path):
    """
    Give every data-label number format an explicit sourceLinked="0".

    openpyxl writes <numFmt formatCode="..."/> and stops. Excel reads a missing
    sourceLinked as TRUE, falls back to the cell's own format and silently
    ignores the one asked for — so '0"%"' would come out as a bare number.
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


def sheet_vertical(wb, rows):
    """
    The same comparison as columns — side by side, not overlapped.

    Nothing here is wrong; it is the other honest reading of the same data, and
    the differences from the horizontal version are all consequences of turning
    it:

      * overlap 0, so the pair TOUCHES rather than overlapping. Vertical space
        is cheap here and the categories are short, so the two years can each
        have their own column instead of sharing a slot.
      * the title names 2019 first, because on this chart 2019 stands on the
        left. A key that runs opposite to the bars is worse than no key.
      * category names sit under the columns, so they must stay short enough to
        print horizontally — Excel's answer to crowding is to rotate them, and
        a rotated label is not a label.
    """
    n = len(rows)
    ws = wb.create_sheet("Vertical")
    xs.write_header(ws, 2, "The same comparison as columns",
                    "Side by side rather than overlapped; 2019 on the left")

    xs.write_table_header(ws, HDR, ["Type", "2024", "2019"])
    for i in range(n):
        r = HDR + 1 + i
        ws.cell(r, COL_TYPE, "=Makeup!B%d" % r)
        ws.cell(r, COL_24, "=Makeup!C%d" % r).number_format = '0"%"'
        ws.cell(r, COL_19, "=Makeup!D%d" % r).number_format = '0"%"'
    xs.widths(ws, [14, 9, 9])

    last = HDR + n
    ch = BarChart()
    ch.type, ch.grouping = "col", "clustered"
    # 0 = the pair touches; the gap between CATEGORIES is what separates groups.
    ch.overlap, ch.gapWidth = 0, 80

    cats = Reference(ws, min_col=COL_TYPE, min_row=HDR + 1, max_row=last)
    for col, colour, edge in ((COL_19, GREY, None), (COL_24, CLARET, "FFFFFF")):
        ch.add_data(Reference(ws, min_col=col, min_row=HDR, max_row=last),
                    titles_from_data=True)
        s = ch.series[-1]
        xs.fill(s, colour)
        if edge:
            s.graphicalProperties.ln = LineProperties(solidFill=edge,
                                                      w=int(1.0 * xs.PT))
        value_labels(s, 10)
    ch.set_categories(cats)

    xs.declutter(ch, gap=None, gridlines=False)
    ch.overlap, ch.gapWidth = 0, 80        # declutter resets gapWidth
    ch.title = two_tone_title(runs=[LEAD, YEAR_19, VS, YEAR_24],
                              subtitle="% REPORTING DAILY USE")
    xs.no_legend(ch)
    ch.y_axis.delete = True                # the numbers ride on the columns
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, X_MAX
    ch.x_axis.txPr = xs._rich(11, xs.INK)
    ch.width, ch.height = 19.0, 11.0
    ws.add_chart(ch, "F8")

    ws.sheet_view.showGridLines = False
    return ws


def sheet_dots(wb, rows):
    """
    A dot plot (dumbbell): the two years as two dots, joined by a rule.

    It keeps what the bars do well — position on a common scale, read against a
    shared axis — while spending far less ink, and it makes the GAP the object
    on the page rather than something inferred from two lengths. For "how far
    apart are these two?", it is the better chart.

    Excel has no dumbbell type, so this is the bullet-chart construction: a bar
    chart supplies the CATEGORY axis (the only way to get text labels down the
    side) and a scatter rides a hidden secondary axis pair carrying everything
    that is actually drawn. The bar series itself is all zeros and invisible.

    Slot space is the whole trick: the secondary y axis is pinned to
    0.5 .. N+0.5, so category i (1-based) sits at y = i and scatter points line
    up with bar centres. Both value ranges are pinned to the SAME numbers — the
    bar's value axis and the scatter's x axis are different axes over the same
    quantity, and if either autoscales the dots sit at the wrong x on every row
    while the chart still looks perfectly normal.
    """
    n = len(rows)
    ws = wb.create_sheet("Dot plot")
    xs.write_header(ws, 2, "The gap is the point",
                    "Two dots and the distance between them, rather than two "
                    "bar lengths to difference by eye")

    xs.write_table_header(ws, HDR, ["Type", "2024", "2019", "Slot", "Bar"])
    for i in range(n):
        r = HDR + 1 + i
        ws.cell(r, COL_TYPE, "=Makeup!B%d" % r)
        ws.cell(r, COL_24, "=Makeup!C%d" % r).number_format = '0"%"'
        ws.cell(r, COL_19, "=Makeup!D%d" % r).number_format = '0"%"'
        ws.cell(r, COL_TYPE + 3, i + 1)        # slot: category i sits at y = i
        ws.cell(r, COL_TYPE + 4, 0)            # the invisible bar series
    last = HDR + n

    # --- helper block: rails and connectors, as gapped scatter series -------
    # Three rows per category — start, end, BLANK. The blank breaks the line,
    # so one series draws six separate segments instead of one zig-zag. This is
    # why dispBlanksAs has to be "gap": a blank read as zero would drag every
    # segment back to the axis.
    H0 = last + 3
    xs.note(ws, H0 - 1,
            "Helper: x/y pairs for the row rails and the connectors, one blank "
            "row between segments so a single series draws six separate rules.")
    for i in range(n):
        r0 = H0 + i * 3
        slot = i + 1
        ws.cell(r0, COL_RAIL_X, 0)
        ws.cell(r0, COL_RAIL_Y, slot)
        ws.cell(r0 + 1, COL_RAIL_X, X_MAX)
        ws.cell(r0 + 1, COL_RAIL_Y, slot)
        ws.cell(r0, COL_CONN_X, "=D%d" % (HDR + 1 + i))
        ws.cell(r0, COL_CONN_Y, slot)
        ws.cell(r0 + 1, COL_CONN_X, "=C%d" % (HDR + 1 + i))
        ws.cell(r0 + 1, COL_CONN_Y, slot)
    H1 = H0 + n * 3 - 1

    xs.widths(ws, [14, 9, 9, 6, 6, 3, 8, 8, 3, 8, 8])

    # --- the bar chart: category axis only ---------------------------------
    ch = BarChart()
    ch.type, ch.grouping = "bar", "clustered"
    ch.add_data(Reference(ws, min_col=COL_TYPE + 4, min_row=HDR, max_row=last),
                titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=COL_TYPE, min_row=HDR + 1,
                                max_row=last))
    ch.series[0].graphicalProperties = GraphicalProperties(
        noFill=True, ln=LineProperties(noFill=True))
    xs.declutter(ch, gap=None, gridlines=False)
    ch.title = two_tone_title()
    xs.no_legend(ch)
    ch.x_axis.scaling.orientation = "maxMin"      # first category at the TOP
    ch.x_axis.txPr = xs._rich(11, xs.INK)
    ch.y_axis.delete = False
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, X_MAX
    ch.y_axis.majorUnit = 10
    ch.y_axis.numFmt = ch.y_axis.number_format = '0"%"'
    ch.y_axis.txPr = xs._rich(10, xs.TEXT_2)

    # --- the scatter: everything that is actually drawn ---------------------
    sc = ScatterChart()
    sc.scatterStyle = "lineMarker"

    def seg(xcol, ycol, colour, width):
        sr = Series(Reference(ws, min_col=ycol, min_row=H0, max_row=H1),
                    Reference(ws, min_col=xcol, min_row=H0, max_row=H1))
        sr.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=colour, w=width))
        sr.marker = Marker(symbol="none")
        sr.smooth = False
        sc.series.append(sr)

    seg(COL_RAIL_X, COL_RAIL_Y, xs.GRID, int(1.0 * xs.PT))       # row rails
    seg(COL_CONN_X, COL_CONN_Y, xs.INK, int(1.75 * xs.PT))       # the gap itself

    for vcol, colour in ((COL_19, GREY), (COL_24, CLARET)):
        sr = Series(Reference(ws, min_col=COL_TYPE + 3, min_row=HDR + 1,
                              max_row=last),
                    Reference(ws, min_col=vcol, min_row=HDR + 1, max_row=last))
        sr.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
        sr.marker = Marker(symbol="circle", size=9, spPr=GraphicalProperties(
            solidFill=colour, ln=LineProperties(solidFill="FFFFFF",
                                                w=int(1.0 * xs.PT))))
        sc.series.append(sr)

    # The secondary pair: each axis points at the OTHER, both on the far side.
    sc.x_axis.axId, sc.y_axis.axId = 300, 200
    sc.x_axis.crossAx, sc.y_axis.crossAx = 200, 300
    sc.x_axis.axPos, sc.y_axis.axPos = "t", "r"
    sc.x_axis.crosses = sc.y_axis.crosses = "max"
    sc.x_axis.scaling.min, sc.x_axis.scaling.max = 0, X_MAX   # == the bar axis
    sc.y_axis.scaling.min, sc.y_axis.scaling.max = 0.5, n + 0.5
    sc.y_axis.scaling.orientation = "maxMin"      # match the reversed categories
    for ax in (sc.x_axis, sc.y_axis):
        ax.delete = False                         # HIDE, never delete
        ax.majorGridlines = None
        ax.majorTickMark = ax.minorTickMark = "none"
        ax.spPr = GraphicalProperties(ln=LineProperties(noFill=True))

    ch += sc
    ch.dispBlanksAs = "gap"
    ch.width, ch.height = 17.0, 11.0
    ws.add_chart(ch, "N8")

    ws.sheet_view.showGridLines = False
    return ws


def sheet_slope(wb, rows):
    """
    The same comparison as a slopegraph — Tufte's form, two periods, six items.

    It answers a different question from the bars. Bars ask "how big is each?";
    a slopegraph asks "what MOVED, and did anything move against the rest?" The
    eye reads slope directly, so six falls register as one shape without anyone
    decoding a number.

    The construction is a transpose: each ITEM is a series and the two years are
    the only two categories — `from_rows=True`, the Switch Row/Column button in
    the Excel UI. Then the value axis is deleted (its job, reading magnitudes,
    has been handed to the end labels), gridlines and legend go, and the chart
    is built TALLER THAN WIDE because a wide plot flattens every slope toward
    horizontal and destroys the only thing the form encodes.

    The axis range is a rhetorical choice: 20-80 pads the data without starting
    at zero. Pinned explicitly so a rebuild cannot silently restate the story.
    """
    n = len(rows)
    ws = wb.create_sheet("Slope")
    xs.write_header(ws, 2, "Every category fell between 2019 and 2024",
                    "% reporting daily use; slope is the encoding, so the "
                    "chart is taller than it is wide")

    # 2019 FIRST here: the left-hand period has to be the left-hand column.
    xs.write_table_header(ws, HDR, ["Type", "2019", "2024"])
    cap_col = COL_TYPE + 4
    for i in range(n):
        r = HDR + 1 + i
        ws.cell(r, COL_TYPE, "=Makeup!B%d" % r)
        ws.cell(r, COL_TYPE + 1, "=Makeup!D%d" % r).number_format = '0"%"'
        ws.cell(r, COL_TYPE + 2, "=Makeup!C%d" % r).number_format = '0"%"'
        # The right-hand caption. Excel renders a label's fields in a fixed
        # order — series name, then value — so "68% Mascara" is unreachable by
        # switching flags on. Building the whole string as the SERIES NAME and
        # showing only that is the way round it, and it stays live.
        ws.cell(r, cap_col, '=TEXT(D%d,"0")&"%% "&B%d' % (r, r)).font = Font(
            name=FACE, size=9, color=xs.TEXT_2, italic=True)
    xs.note(ws, HDR + n + 2,
            "Helper: the right-hand end labels. Each is a series name for the "
            "chart, so the caption carries the 2024 value.")
    xs.widths(ws, [14, 9, 9, 3, 18])

    last = HDR + n
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=COL_TYPE, max_col=COL_TYPE + 2,
                          min_row=HDR + 1, max_row=last),
                from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=COL_TYPE + 1, max_col=COL_TYPE + 2,
                                min_row=HDR))

    off = dict(showCatName=False, showLegendKey=False, showPercent=False,
               showBubbleSize=False)
    for i, s in enumerate(ch.series):
        r = HDR + 1 + i
        s.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=xs.INK, w=int(1.5 * xs.PT)))
        s.marker = Marker(symbol="circle", size=6, spPr=GraphicalProperties(
            solidFill=xs.INK, ln=LineProperties(solidFill=xs.INK)))
        s.smooth = False
        s.tx = SeriesLabel(strRef=StrRef("'%s'!$%s$%d" % (
            ws.title, get_column_letter(cap_col), r)))
        right = DataLabel(idx=1, showSerName=True, showVal=False,
                          dLblPos="r", **off)
        labels = [right]
        # The LEFT end is only labelled where a label can actually be read.
        # Four items start within three points of each other (42, 41, 42, 39):
        # no font size or nudge separates those, because it is a property of
        # the data, not the formatting. Printing the top and bottom of the
        # cluster states its range honestly; printing all four states nothing
        # and overprints. The right end names every item, so nothing is lost.
        if i in LEFT_LABELLED:
            labels.insert(0, DataLabel(idx=0, showSerName=False, showVal=True,
                                       dLblPos="l", **off))
        for lbl in labels:
            lbl.txPr = xs._rich(9, xs.INK)
            lbl.numFmt = '0"%"'
        # List-level flags stay False, per-point dLbl entries carry the truth —
        # otherwise Excel labels BOTH ends identically and the name appears twice.
        s.dLbls = DataLabelList(dLbl=labels, showVal=False,
                                showSerName=False, **off)

    ch.legend = None
    ch.title = two_tone_title(runs=[("Daily makeup use", xs.INK, True)],
                              subtitle="% REPORTING DAILY USE")
    # Both end labels live OUTSIDE the plot, so the plot takes only the middle.
    ch.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.13, y=0.15, w=0.55, h=0.76))
    y = ch.y_axis
    y.delete = True                    # sanctioned here: end labels do its job
    y.scaling.min, y.scaling.max = 20, 80
    y.majorGridlines = None
    y.crossBetween = "midCat"          # 2019 sits ON the left edge
    x = ch.x_axis
    x.delete = False
    x.majorGridlines = None
    x.majorTickMark = x.minorTickMark = "none"
    x.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    x.txPr = xs._rich(10, xs.TEXT_2)
    # No declutter() call on this chart — the slopegraph strips more than it
    # does — so the chart-area border has to be cleared by hand.
    ch.graphical_properties = GraphicalProperties(
        solidFill=xs.SURFACE, ln=LineProperties(noFill=True))
    # Sized to match the workbook as it stands — resized by hand in Excel and
    # kept. Still taller than wide, which is the one thing the form requires.
    ch.width, ch.height = 12.1, 14.5
    ws.add_chart(ch, "H8")

    ws.sheet_view.showGridLines = False
    return ws


def sheet_wrong(wb, rows):
    """
    The same six numbers as a line chart — a chart that is WRONG, and not
    because it is ugly. It is drawn competently on purpose: the failure has to
    be the encoding, not the craft, or the lesson lands as "make it prettier".

    A line asserts that its points are CONNECTED — that the space between two
    of them is real and can be travelled. Makeup types are nominal: there is no
    quantity of "between mascara and lipstick", so every slope on this chart is
    an artefact of the order the rows happen to sit in. Sort the categories
    differently and the shape changes completely while the data does not, which
    is the tell: a line chart's shape should never depend on row order.

    Bars have no such claim. They sit side by side and assert nothing about the
    gap, which is why the same data is honest in the chart on the other sheet.
    """
    n = len(rows)
    ws = wb.create_sheet("Wrong chart type")
    xs.write_header(
        ws, 2, "The same numbers as a line chart - and it is wrong",
        "A line says its points are connected. Makeup types have no order, so "
        "every slope here is an artefact of the row order, not a trend.")

    xs.write_table_header(ws, HDR, ["Type", "2024", "2019"])
    for i in range(n):
        r = HDR + 1 + i
        ws.cell(r, COL_TYPE, "=Makeup!B%d" % r)
        ws.cell(r, COL_24, "=Makeup!C%d" % r).number_format = '0"%"'
        ws.cell(r, COL_19, "=Makeup!D%d" % r).number_format = '0"%"'
    xs.widths(ws, [14, 9, 9])

    last = HDR + n
    ch = LineChart()
    cats = Reference(ws, min_col=COL_TYPE, min_row=HDR + 1, max_row=last)
    for col, colour in ((COL_19, GREY), (COL_24, CLARET)):
        ch.add_data(Reference(ws, min_col=col, min_row=HDR, max_row=last),
                    titles_from_data=True)
        s = ch.series[-1]
        s.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=colour, w=int(2.25 * xs.PT)))
        s.marker = Marker(symbol="circle", size=7,
                          spPr=GraphicalProperties(
                              solidFill=colour,
                              ln=LineProperties(solidFill="FFFFFF",
                                                w=int(1.25 * xs.PT))))
        s.smooth = False
    ch.set_categories(cats)

    xs.declutter(ch, gap=None, gridlines=True, num_fmt='0"%"')
    ch.title = two_tone_title()
    xs.legend_bottom(ch)
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, X_MAX
    ch.y_axis.majorUnit = 20
    ch.x_axis.txPr = xs._rich(11, xs.INK)
    ch.width, ch.height = 19.2, 10.9   # matches the workbook, resized in Excel
    ws.add_chart(ch, "F8")

    xs.note(ws, last + 2,
            "Deliberately wrong. Re-sort the rows on the Makeup sheet and this "
            "chart changes shape while the data does not - that is the tell.")
    ws.sheet_view.showGridLines = False
    return ws


def hide_tick_labels(path, ax_ids):
    """
    Hide an axis's tick labels by writing tickLblPos into the saved package.

    openpyxl accepts `axis.tickLblPos = "none"` and then DROPS it on the way
    out, so the dot plot's hidden scatter pair kept printing 0-80 along the
    bottom and 0.5-6.5 down the right.

    Deleting those axes instead is not an option: a deleted secondary axis makes
    Excel reassign the scatter to the PRIMARY pair, which then rescales to the
    slot numbers and collapses the chart.

    Schema order: tickLblPos sits after minorTickMark and before spPr/crossAx.
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
        # Two plain patterns rather than one with a backreference: the axis
        # tag names are known, and a backslash-1 in a non-raw string is an
        # octal escape waiting to happen - which is exactly how this pattern
        # lost its closing tag and silently matched nothing.
        spans = sorted(
            (m for tag in ("valAx", "catAx")
             for m in re.finditer("<%s>.*?</%s>" % (tag, tag), x, re.S)),
            key=lambda m: m.start())
        for m in spans:
            block = m.group(0)
            if not any('axId val="%d"' % a in block for a in ax_ids):
                continue
            if "<tickLblPos" in block:
                continue
            for anchor in ("<spPr>", "<spPr ", "<txPr>", "<crossAx"):
                i = block.find(anchor)
                if i != -1:
                    out.append(x[pos:m.start()]
                               + block[:i] + '<tickLblPos val="none"/>' + block[i:])
                    pos, changed = m.end(), True
                    fixed += 1
                    break
        if changed:
            parts[name] = ("".join(out) + x[pos:]).encode("utf-8")
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for nm, blob in parts.items():
            zout.writestr(nm, blob)
    shutil.move(tmp, str(path))
    return fixed


def nudge_end_labels(path, marker, moves):
    """
    Shift individual end labels, written straight into the saved package.

    openpyxl's DataLabel has no `layout`, so a label can only sit at one of
    Excel's fixed positions. The bottom four items end 34/33/31/28 — a one-point
    gap at 9 pt type is about two points of plot — so two captions want the same
    line. CT_DLbl takes a manualLayout offset in FACTOR mode (a fraction of the
    chart, y running DOWNWARD), which is the only way to prise them apart.

    Targeting is by ORDER, which is stable: the Nth <ser> block, then the <dLbl>
    whose idx matches. `moves` is {(series_index, point_index): (dx, dy)}.
    """
    src = zipfile.ZipFile(path)
    parts = {nm: src.read(nm) for nm in src.namelist()}
    src.close()
    moved = 0
    for name in list(parts):
        if not re.match(r"xl/charts/chart\d+\.xml$", name):
            continue
        x = parts[name].decode()
        if marker not in x:
            continue
        out, pos = [], 0
        for si, ser in enumerate(re.finditer(r"<ser>.*?</ser>", x, re.S)):
            block = ser.group(0)
            new, bpos = [], 0
            for lbl in re.finditer(r"<dLbl>.*?</dLbl>", block, re.S):
                b = lbl.group(0)
                m = re.search(r'<idx val="(\d+)"/>', b)
                key = (si, int(m.group(1))) if m else None
                if key not in moves:
                    continue
                dx, dy = moves[key]
                lay = ('<layout><manualLayout><x val="%s"/><y val="%s"/>'
                       '</manualLayout></layout>' % (dx, dy))
                i = b.find("/>", b.find("<idx ")) + 2
                new.append(block[bpos:lbl.start()] + b[:i] + lay + b[i:])
                bpos = lbl.end()
                moved += 1
            if new:
                out.append(x[pos:ser.start()] + "".join(new) + block[bpos:])
                pos = ser.end()
        if out:
            parts[name] = ("".join(out) + x[pos:]).encode("utf-8")
    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for nm, blob in parts.items():
            zout.writestr(nm, blob)
    shutil.move(tmp, str(path))
    return moved


def build():
    rows = load()
    n = len(rows)
    wb = Workbook()
    ws = wb.active
    ws.title = "Makeup"

    xs.write_header(ws, 2, "Daily makeup use fell in every category",
                    "% reporting daily use, 2024 vs 2019")
    xs.write_table_header(ws, HDR, ["Type", "2024", "2019", "Change (pp)"])
    for i, (name, v24, v19) in enumerate(rows):
        r = HDR + 1 + i
        ws.cell(r, COL_TYPE, name)
        ws.cell(r, COL_24, v24).number_format = '0"%"'
        ws.cell(r, COL_19, v19).number_format = '0"%"'
        ws.cell(r, COL_CHG, "=C%d-D%d" % (r, r)).number_format = '+0;-0'
    xs.widths(ws, [14, 9, 9, 13])

    last = HDR + n
    ch = BarChart()
    ch.type, ch.grouping = "bar", "clustered"
    ch.overlap, ch.gapWidth = 82, 55

    cats = Reference(ws, min_col=COL_TYPE, min_row=HDR + 1, max_row=last)
    # 2019 FIRST so it is drawn behind; 2024 second lands on top of it.
    # The claret bar carries a white edge: it sits directly on top of the grey
    # one, and without a border the two similar-value bars meet as a single
    # shape wherever they overlap. The rule is what keeps the pair legible as
    # two bars rather than one. xs.fill sets ln=noFill, so this comes after.
    for col, colour, size, edge in ((COL_19, GREY, 9, None),
                                    (COL_24, CLARET, 12, "FFFFFF")):
        ch.add_data(Reference(ws, min_col=col, min_row=HDR, max_row=last),
                    titles_from_data=True)
        s = ch.series[-1]
        xs.fill(s, colour)
        if edge:
            # 1pt, not 2: export_images scales line widths with the figure,
            # so anything heavier here reads as a frame at 3x rather than a seam.
            s.graphicalProperties.ln = LineProperties(solidFill=edge,
                                                      w=int(1.0 * xs.PT))
        value_labels(s, size)
    ch.set_categories(cats)

    xs.declutter(ch, gap=None, gridlines=False, x_line=False)
    ch.overlap, ch.gapWidth = 82, 55       # declutter resets gapWidth
    ch.title = two_tone_title()
    xs.no_legend(ch)

    # Category axis: first row at the TOP. Reversing it also sends the value
    # axis to the far side, so the category names are pinned back to the left.
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.txPr = xs._rich(11, xs.INK)
    ch.x_axis.txPr.p[0].pPr.defRPr.latin = DFont(typeface=FACE)
    ch.y_axis.crosses = "max"
    # The numbers are on the bars; a value axis would only repeat them.
    ch.y_axis.delete = True
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, X_MAX

    ch.width, ch.height = 17.0, 12.0
    ws.add_chart(ch, "G8")

    if SOURCE:
        xs.write_source(ws, last + 2, SOURCE)
    xs.note(ws, last + 3,
            "Values transcribed from a supplied chart image; originating "
            "survey not established. Add the citation before publishing.")
    ws.sheet_view.showGridLines = False

    sheet_vertical(wb, rows)
    sheet_dots(wb, rows)
    sheet_slope(wb, rows)
    sheet_wrong(wb, rows)

    wb.save(OUT)
    fixed = pin_label_formats(OUT)
    hidden = hide_tick_labels(OUT, (200, 300))   # the dot plot's hidden pair
    # Prise apart the four end labels that share a two-point band.
    spread = nudge_end_labels(OUT, "'Slope'!", {
        (2, 1): (0, -0.011),      # 34% Foundation, up
        (3, 1): (0, -0.001),      # 33% Powder
        (4, 1): (0, 0.006),       # 31% Eyeshadow, down
        (5, 1): (0, 0.010),       # 28% Eyeliner, down
    })
    print("wrote %s  (%d rows, label formats pinned on %d chart(s), "
          "%d end labels nudged, %d axes hidden)"
          % (OUT.name, n, fixed, spread, hidden))


if __name__ == "__main__":
    build()
