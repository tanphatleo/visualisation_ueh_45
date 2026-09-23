"""
Feature satisfaction survey -- the same 15 rows drawn twice.

    Sheet "1 All six"   every response category, one shade of blue each
    Sheet "2 Top two"   the two positive categories in blue, the rest grey,
                        and moved to the LEFT EDGE so they share a baseline

The second chart is not a restyle of the first. The move that makes it work is
reordering the series: "Completely satisfied" is plotted FIRST, so every bar
starts its positive block at x = 0. In the first chart those blocks float --
each one starts wherever the categories to its left happened to end -- and
comparing them means comparing lengths that begin in fifteen different places,
which is the one thing bar length cannot do.

Feature names are invented (a generic collaboration tool); the percentages are
transcribed from the supplied reference images. Every row sums to 100.

    python build_xlsx.py     ->  feature_satisfaction.xlsx
"""

import csv
import re
import shutil
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, ScatterChart, Series
from openpyxl.chart.label import DataLabel, DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.legend import LegendEntry
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
sys.path.insert(0, str(HERE / "style"))
import xl_style as xs  # noqa: E402

OUT = HERE / "feature_satisfaction.xlsx"

CATS = ["Have not used", "Not satisfied at all", "Not very satisfied",
        "Somewhat satisfied", "Very satisfied", "Completely satisfied"]

# ---- chart 1: one sequential blue ramp, lightest = least satisfied ---------
# Six steps of ONE hue, not six hues. The scale is ordered, so the colours are
# ordered too; a categorical palette would imply the six answers are unrelated
# kinds rather than points on a line.
RAMP = ["E6EEF6", "B8CFE4", "7FA6C9", "3F73A6", "0F5499", "0F3762"]
# Label colour has to follow the fill: white fails on the first three
# (2,9:1 on 7FA6C9) and ink fails on the last three. Computed, not eyeballed.
RAMP_LABEL = [xs.INK, xs.INK, xs.INK, "FFFFFF", "FFFFFF", "FFFFFF"]

# ---- chart 2: ONE blue carries the message, the rest is context ------------
# Both positive answers are the SAME navy, told apart only by the white rule
# between them: the message is "very or completely satisfied" as one quantity,
# and two shades would invite a comparison between them that is not the point.
# The three dissatisfied answers likewise share one grey -- they are context,
# and giving each its own tone spends contrast on the part nobody is reading.
TOP2 = {"Completely satisfied": "0F3762", "Very satisfied": "0F3762"}
GREYS = {"Somewhat satisfied": "BDBDBD", "Not very satisfied": "BDBDBD",
         "Not satisfied at all": "BDBDBD", "Have not used": "EDEDED"}

# ---- chart 3: the two dissatisfied answers, in one orange ------------------
# BF5700, not Office's ED7D31. The bright orange reads better as a colour but
# fails white text at 2,8:1, and these segments are 8-14% wide -- too narrow to
# carry ink on a light fill without the number fighting the bar. 4,6:1 passes.
ORANGE = "BF5700"
DISS = ["Not satisfied at all", "Not very satisfied"]

# The five answers that form a scale, worst to best. "Have not used" is not on
# it -- it is a different question -- and always goes last in the stack.
SCALE = ["Not satisfied at all", "Not very satisfied", "Somewhat satisfied",
         "Very satisfied", "Completely satisfied"]

# ---- chart 4: never used --------------------------------------------------
TEAL = "0D7680"        # white labels at 5,4:1 — the HIGHLIGHT
TEAL_BASE = "129FAE"   # the other fourteen bars, one step lighter
UNUSED = ["Have not used"]
# The satisfied answers drop to a paler grey than the dissatisfied ones. On
# this chart the whole right-hand side is context, and the part of it a reader
# is least likely to want is the part about satisfaction.
UNUSED_GREYS = {"Very satisfied": "D9D9D9", "Completely satisfied": "D9D9D9"}

# ---- chart 5: every answer, annotated -------------------------------------
# The reference palette: cool for unused, warm for unhappy, grey for lukewarm,
# blue for happy. Three families instead of one ramp, because this chart is
# read for THREE separate questions at once and the call-outs name each one.
# EVERY SERIES IS DRAWN PALE, and only the points the call-outs name are given
# the strong colour. That is the difference between a chart that shows six
# categories and one that makes an argument: the shape of all fifteen rows is
# still there to be read, but five points are lit and the eye goes to them.
ANNOTATED = {"Have not used": "A5D5DA", "Not satisfied at all": "F4B183",
             "Not very satisfied": "F4B183", "Somewhat satisfied": "BFBFBF",
             "Very satisfied": "ACCDF2", "Completely satisfied": "ACCDF2"}
# The highlight colour per series, applied to the points in ANNOTATED_KEEP.
ANNOTATED_HIGHLIGHT = {"Have not used": TEAL, "Not satisfied at all": "ED7D31",
                       "Not very satisfied": "ED7D31",
                       "Very satisfied": "0F3762", "Completely satisfied": "0F3762"}
# The SAME index sets drive the fills and the labels, because they are the same
# decision: these five points are the ones the call-outs are about.
#   0,1 Search and Mobile app   9,13 Bulk edit and Guest access   14 API access
ANNOTATED_KEEP = {"Very satisfied": {0, 1}, "Completely satisfied": {0, 1},
                  "Not satisfied at all": {9, 13}, "Not very satisfied": {9, 13},
                  "Have not used": {14}}
# Labels only ever land on highlighted points, and every highlight is dark
# enough for white — so there is no per-series contrast switch to make here.
ANNOTATED_LABEL = {k: "FFFFFF" for k in ANNOTATED}
# Legend keys tinted to their own series, so the key reads as the thing it
# names. "Somewhat satisfied" (index 3) is left default: it is the one answer
# with nothing to say.
# idx -> (colour, bold). Index 3, "Somewhat satisfied", is left default: it is
# the one answer with nothing to say. Index 0 is tinted but not bold.
ANNOTATED_LEGEND = {0: (TEAL_BASE, False), 1: ("F4B183", True),
                    2: ("F4B183", True), 4: ("ACCDF2", True), 5: ("ACCDF2", True)}

CALLOUTS = [
    ("0F3762", "Search and the mobile app continue to top user satisfaction."),
    (ORANGE, "Users are least satisfied with bulk edit and guest access. What "
             "improvements can we make here for a better experience?"),
    (TEAL, "API access is least used — half have never touched it. What can we "
           "do to raise utilisation among existing users?"),
]
FOOTNOTE = ('Responses to "How satisfied have you been with each of these '
            'features?" Context still missing: how many people completed the '
            "survey, and what share of all users do they represent?")

MIN_LABEL = 3          # below this a label does not fit inside its segment

# ---- sheets 6 and 7: the action maps --------------------------------------
# Usage is everything except "have not used", and BOTH rates are taken over
# usage, not over all respondents. That is the point of the pair of columns: a
# feature half the base has never opened should be judged on what its actual
# users think, not diluted by people with no opinion. API access scores 8% raw
# dissatisfaction and 16% among the people who have used it.
C_USAGE, C_SAT, C_DISS = 9, 10, 11        # I J K

# Split lines. Round numbers, not medians -- a median split guarantees four
# non-empty boxes but moves every time the data does, and "85% adoption" is a
# threshold a product team can argue with. Both were checked against the data:
# no feature sits ON a line, which would leave its quadrant ambiguous.
U_SPLIT = 85          # % of respondents who use the feature at all
D_SPLIT = 0.10        # of users, dissatisfied
S_SPLIT = 0.65        # of users, satisfied

DISS_ACTIONS = {          # (x high?, y high?) -> label, colour
    (True, True): ("FIX FIRST", "990F3D"),
    (False, True): ("INVESTIGATE", ORANGE),
    (True, False): ("PROTECT", "0F3762"),
    (False, False): ("PROMOTE", TEAL),
}
# Same four verbs as DISS_ACTIONS wherever the meaning is the same: a feature
# that is heavily used and poorly rated is FIX FIRST on either map, and calling
# it IMPROVE here would make the pair look like two unrelated frameworks.
SAT_ACTIONS = {
    (True, True): ("PROTECT", "0F3762"),
    (False, True): ("PROMOTE", TEAL),
    (True, False): ("FIX FIRST", ORANGE),
    (False, False): ("RECONSIDER", "990F3D"),
}
DISS_RANGE = (40, 108, 0.0, 0.35)
SAT_RANGE = (40, 108, 0.30, 0.95)

# Where each quadrant caption sits, and which side of its anchor the text goes.
# PLACED BY LOOKING AT THE PLOT, not by formula: the obvious choice -- the four
# outer corners -- drops PROTECT straight on top of Search and Mobile app,
# because the quadrant that means "everything is fine here" is the one with
# most of the points in it. Its caption goes to the empty inner-bottom corner
# instead, just right of the split line.
DISS_CAPTIONS = {(False, True): (42, 0.335, "r"), (True, True): (106, 0.335, "l"),
                 (False, False): (42, 0.010, "r"), (True, False): (86.5, 0.010, "r")}
SAT_CAPTIONS = {(False, True): (42, 0.925, "r"), (True, True): (106, 0.925, "l"),
                (False, False): (42, 0.315, "r"), (True, False): (106, 0.315, "l")}

TITLE_1 = "How satisfied have you been with each of these features?"
SUB_1 = "Product X user satisfaction: features — % of respondents"
TITLE_2 = "SEARCH and MOBILE APP has highest user satisfaction"
SUB_2 = "User satisfaction (%) by features"
TITLE_3 = "GUEST ACCESS and BULK EDIT draw the most dissatisfaction"
SUB_3 = "Not very or not at all satisfied (%) by feature"
TITLE_4 = "API ACCESS is barely used — half have never touched it"
SUB_4 = "Have not used the feature (%), by feature"
TITLE_5 = "User satisfaction varies greatly by feature"
SUB_5 = "Product X user satisfaction: features"
TITLE_6 = "BULK EDIT is the one to fix: heavily used and widely disliked"
SUB_6 = "Dissatisfied share of users (%) against reach (% of respondents using it)"
TITLE_7 = "TASK COMMENTS and TIME TRACKING are liked but under-adopted"
SUB_7 = "Satisfied share of users (%) against reach (% of respondents using it)"
SOURCE = ""            # illustrative survey; no outside source to credit

TITLE_PT = 16          # takeaway line
SUB_PT = 16            # descriptive line, same size, separated by colour alone
SUB_INK = "404040"     # Excel's tx1 at 75% luminance, written literally
LEGEND_PT = 12
CAT_PT = 12

# Title and legend are placed by hand; the PLOT AREA IS LEFT AUTOMATIC and
# must stay that way. A plot-area manualLayout carrying only y and w -- no x,
# no h -- does not mean "keep the rest": Excel reads the missing edges as zero,
# floats the plot to the top of the frame and draws the title and legend on
# top of the bars. Either give all four, or give none.
TITLE_XY = (0.00926, 0.01512)
LEGEND_XYWH = (0.001154, 0.13625, 0.9, 0.05167)

HDR = 8
R0 = HDR + 1
C_NAME = 2             # B


def load():
    with open(HERE / "feature_satisfaction.csv", encoding="utf-8") as fh:
        rows = [(r["Feature"], [int(r[c]) for c in CATS])
                for r in csv.DictReader(fh)]
    for name, vals in rows:
        assert sum(vals) == 100, "%s sums to %d, not 100" % (name, sum(vals))
    return rows


_ROWS = None


def rows_cache():
    """load() re-reads the CSV; the per-point label loop asks for it 60 times."""
    global _ROWS
    if _ROWS is None:
        _ROWS = load()
    return _ROWS


def rich(pt, colour, bold=False):
    c = CharacterProperties(sz=None if pt is None else int(pt * 100),
                            b=bold, solidFill=colour, latin=None)
    return RichText(bodyPr=RichTextProperties(),
                    p=[Paragraph(pPr=ParagraphProperties(defRPr=c), endParaRPr=c)])


def two_line_title(headline, subline, headline_colour):
    """
    Takeaway on one line, description under it, both inside the chart title.

    Excel writes this as ONE paragraph with an <a:br> between the runs.
    openpyxl cannot: Paragraph serialises its elements in schema order
    (pPr, r, br, fld, endParaRPr), so every run is emitted before every break
    and the line break lands after both lines instead of between them. Two
    paragraphs give the same two lines and are what this builds.
    """
    def para(text, pt, colour, bold):
        cp = CharacterProperties(sz=int(pt * 100), b=bold, solidFill=colour,
                                 latin=None)
        return Paragraph(pPr=ParagraphProperties(algn="l", defRPr=cp),
                         r=[RegularTextRun(rPr=cp, t=text)])

    return Title(
        tx=Text(rich=RichText(bodyPr=RichTextProperties(), p=[
            para(headline, TITLE_PT, headline_colour, True),
            para(subline, SUB_PT, SUB_INK, False)])),
        overlay=False,
        layout=Layout(manualLayout=ManualLayout(
            xMode="edge", yMode="edge", x=TITLE_XY[0], y=TITLE_XY[1])))


def write_data(ws, rows):
    xs.write_header(ws, 2, TITLE_1, SUB_1)
    xs.note(ws, 5,
            "Feature names are illustrative. Percentages are transcribed from "
            "the reference survey chart; every row sums to 100.")
    xs.write_table_header(ws, HDR, ["Feature"] + CATS)
    for i, (name, vals) in enumerate(rows):
        r = R0 + i
        ws.cell(r, C_NAME, name).font = Font(name=xs.FONT, size=11, color=xs.INK)
        for j, v in enumerate(vals):
            c = ws.cell(r, C_NAME + 1 + j, v)
            c.font = Font(name=xs.FONT, size=11, color=xs.TEXT_2)
            c.alignment = Alignment(horizontal="right")
            c.number_format = '0"%"'
    # --- derived columns, as LIVE FORMULAS ---------------------------------
    # The chart reads the table and the table computes itself, so a reader can
    # change one response and watch both scatter charts move. Hard-coding these
    # would make the workbook a picture of an answer instead of the answer.
    for j, h in enumerate(["Usage", "Satisfied of users", "Dissatisfied of users"]):
        c = ws.cell(HDR, C_USAGE + j, h)
        c.font = Font(name=xs.FONT, size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=xs.ACCENT)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i in range(len(rows)):
        r = R0 + i
        u = ws.cell(r, C_USAGE, "=SUM(D%d:H%d)" % (r, r))      # all but "not used"
        s = ws.cell(r, C_SAT, "=(G%d+H%d)/I%d" % (r, r, r))
        d = ws.cell(r, C_DISS, "=(D%d+E%d)/I%d" % (r, r, r))
        u.number_format = '0"%"'
        s.number_format = d.number_format = "0.0%"
        for c in (u, s, d):
            c.font = Font(name=xs.FONT, size=11, color=xs.TEXT_2)
            c.alignment = Alignment(horizontal="right")

    last = R0 + len(rows) - 1

    tot = ws.cell(last + 2, C_NAME, "Each row totals 100%")
    tot.font = Font(name=xs.FONT, size=10, italic=True, color=xs.TEXT_2)
    if SOURCE:
        xs.write_source(ws, last + 4, SOURCE)
    xs.note(ws, last + 3,
            "Usage = every answer except 'have not used'. Both rates are taken "
            "over USAGE, not over all respondents — a feature half the base has "
            "never opened is judged on what its actual users think.")
    xs.widths(ws, [16] + [13] * len(CATS) + [9, 12, 12])
    ws.sheet_view.showGridLines = False
    return last


def stacked_bar(data_ws, last, order, colours, labels, label_colours,
                headline, subline, bold_legend=(), pin_legend=False,
                headline_colour=None, keep=None, width=26.0, legend_w=None,
                highlight=None, highlight_idx=None, label_pos=None,
                legend_colours=None, src=None, order_rows=None):
    """
    A 100%-wide stacked bar. `order` is the series order LEFT TO RIGHT, which
    is the whole difference between the two charts.
    """
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap = "bar", "stacked", 100
    ch.title = None

    # `src` = (worksheet, header row, last data row, first column). Defaults to
    # the Data table; sheets 3 and 4 point it at their own ranked copy, because
    # a bar chart plots categories in the order its source range holds them.
    src_ws, hdr, src_last, name_col = src or (data_ws, HDR, last, C_NAME)
    rows_here = order_rows or rows_cache()
    cats = Reference(src_ws, min_col=name_col, min_row=hdr + 1, max_row=src_last)
    for name in order:
        col = name_col + 1 + CATS.index(name)
        s = Series(Reference(src_ws, min_col=col, min_row=hdr, max_row=src_last),
                   title_from_data=True)
        # White hairline between segments: with six of them on one bar the
        # boundaries are the only thing separating two neighbouring shades.
        s.graphicalProperties = GraphicalProperties(
            solidFill=colours[name],
            ln=LineProperties(solidFill="FFFFFF", w=int(0.75 * xs.PT)))
        # Per-POINT fills. `highlight` gives the colour and `highlight_idx`
        # the points, falling back to `keep`.
        #
        # THESE ARE TWO DIFFERENT QUESTIONS and only sheet 5 answers them the
        # same way. There, the lit points and the labelled points are the five
        # the call-outs name. On sheet 4 one bar is lit and ALL of them are
        # labelled, so passing `keep` for both silently stripped fourteen
        # labels off a chart that wants them.
        #
        # openpyxl exposes these as Series.data_points; there is no `dPt`
        # element to pass to the constructor, and Series.__elements__ is empty
        # so the attribute is not discoverable from the class either. A point
        # that overrides the fill does NOT inherit the series outline, so each
        # lit point restates the white separator as well as its colour.
        lit = (highlight_idx or keep or {})
        if highlight and name in highlight:
            for idx in sorted(lit.get(name, ())):
                s.data_points.append(DataPoint(
                    idx=idx, spPr=GraphicalProperties(
                        solidFill=highlight[name],
                        ln=LineProperties(solidFill="FFFFFF",
                                          w=int(0.75 * xs.PT)))))
        if name in labels:
            s.dLbls = DataLabelList(
                showVal=True, showSerName=False, showCatName=False,
                showLegendKey=False, showPercent=False, showBubbleSize=False,
                dLblPos=label_pos,
                spPr=GraphicalProperties(noFill=True),
                # No explicit size: the labels take the chart's, so changing
                # the chart size does not leave them behind. No explicit
                # numFmt either -- without one Excel source-links to the cell,
                # and the cells are already formatted 0"%".
                txPr=rich(None, label_colours[name]))
            # A 1% or 2% segment is narrower than its own label, so the text
            # spills over its neighbours. Silence those points individually --
            # there is no "hide small labels" switch.
            #
            # OOXML's <c:dLbl> has a <c:delete> element and OPENPYXL'S DataLabel
            # DOES NOT EXPOSE IT -- DataLabel(delete=True) is a TypeError, the
            # same gap as the dropped tickLblPos. Turning every show* flag off
            # is the reachable equivalent: the label is still emitted, with
            # nothing in it to draw.
            #
            # `keep` names the points worth labelling outright; without it the
            # rule is "label it if it fits", i.e. anything >= MIN_LABEL.
            wanted = None if keep is None else keep.get(name, set())
            s.dLbls.dLbl = [
                DataLabel(idx=i, showVal=False, showSerName=False,
                          showCatName=False, showLegendKey=False,
                          showPercent=False, showBubbleSize=False)
                for i, (_, vals) in enumerate(rows_here)
                if (i not in wanted if wanted is not None
                    else vals[CATS.index(name)] < MIN_LABEL)]
        ch.series.append(s)
    ch.set_categories(cats)

    xs.declutter(ch, gridlines=False, x_line=False, y_labels=False, gap=35)
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 100
    # A horizontal bar chart puts the first category at the BOTTOM. The survey
    # reads top down, so the category axis is reversed -- and because nothing
    # else here rides a second axis, that is the whole fix.
    ch.x_axis.scaling.orientation = "maxMin"
    ch.x_axis.delete = False
    ch.x_axis.txPr = rich(CAT_PT, xs.INK)
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))

    ch.title = two_line_title(headline, subline,
                              headline_colour or TOP2["Completely satisfied"])
    ch.legend.position = "t"
    ch.legend.overlay = False
    ch.legend.txPr = rich(LEGEND_PT, xs.TEXT_2)
    # Pinned only where it is wanted. w=0.9 is narrower than six 12pt entries
    # need, so check_chart_overlap.ps1 reports the legend clipped by the chart
    # edge -- true of the hand-edited original too, and accepted there because
    # the two entries that matter sit at the left. Chart 1 has no highlighted
    # entry to protect, so its legend is left automatic and stays clean.
    if pin_legend:
        ch.legend.layout = Layout(manualLayout=ManualLayout(
            xMode="edge", yMode="edge", x=LEGEND_XYWH[0], y=LEGEND_XYWH[1],
            w=legend_w or LEGEND_XYWH[2], h=LEGEND_XYWH[3]))
    # Bold the legend entries that are actually carrying the message, in their
    # own colour. The key beside "Completely satisfied" is a navy square on
    # white; grey 12pt next to it reads as one more thing to ignore.
    ch.legend.legendEntry = [
        LegendEntry(idx=i, txPr=rich(LEGEND_PT,
                                     headline_colour or TOP2["Completely satisfied"],
                                     bold=True))
        for i in bold_legend] + [
        LegendEntry(idx=i, txPr=rich(LEGEND_PT, c, bold=bold))
        for i, (c, bold) in sorted((legend_colours or {}).items())]

    ch.width, ch.height = width, 14.0
    return ch


def action_map(ws, data_ws, last, y_col, y_split, actions, rng,
               captions, headline, subline, nudge=None, caption_pt=16):
    """
    One point per feature on reach x rate, with the plane cut into four
    named actions.

    ONE SERIES PER FEATURE, fifteen of them, each holding a single point.
    That is not fussiness: a scatter's data labels can show the X value, the Y
    value or the SERIES name, and there is no reachable way to print the
    category name -- openpyxl's DataLabel has no `tx` element, so literal text
    would mean a package rewrite. Make the feature the series name and
    showSerName prints it for free, and each point can then carry its own
    quadrant colour, which a single 15-point series could not.

    The dividers and the four quadrant captions are more of the same trick:
    two-point series drawn as lines, and one-point series with no marker whose
    only job is to place their own name.
    """
    x0, x1, y0, y1 = rng
    nudge = nudge or {}
    ch = ScatterChart()
    # lineMarker, not marker: every series here has its line explicitly set to
    # noFill, so nothing is drawn either way, but it is what Excel writes back
    # and keeping it matching avoids a spurious diff on every round trip.
    ch.scatterStyle = "lineMarker"

    # --- helper block: dividers, then one anchor per quadrant caption ------
    h = last + 20
    xs.note(ws, h - 1,
            "Helper: the two split lines, then one anchor per quadrant caption. "
            "Each is a chart series; none is data.")
    grid = [(U_SPLIT, y0), (U_SPLIT, y1), (None, None),      # vertical rule
            (x0, y_split), (x1, y_split)]                    # horizontal rule
    for i, (xv, yv) in enumerate(grid):
        if xv is None:
            continue
        ws.cell(h + i, C_NAME + 1, xv)
        ws.cell(h + i, C_NAME + 2, yv)
    for i, (rule_from, rule_to) in enumerate(((0, 1), (3, 4))):
        s = Series(Reference(ws, min_col=C_NAME + 2, min_row=h + rule_from,
                             max_row=h + rule_to),
                   Reference(ws, min_col=C_NAME + 1, min_row=h + rule_from,
                             max_row=h + rule_to))
        s.marker = Marker(symbol="none")
        s.graphicalProperties = GraphicalProperties(
            ln=LineProperties(solidFill=xs.GRID, w=int(1.25 * xs.PT)))
        s.smooth = False
        ch.series.append(s)

    # --- one series per feature, coloured by the quadrant it lands in ------
    for i in range(last - R0 + 1):
        r = R0 + i
        usage = data_ws.cell(r, C_USAGE).value
        # the formulas are strings until Excel opens the file, so the quadrant
        # is worked out from the raw answers here, the same arithmetic
        vals = rows_cache()[i][1]
        u = sum(vals[1:])
        rate = ((vals[4] + vals[5]) if y_col == C_SAT else (vals[1] + vals[2])) / u
        _, colour = actions[(u >= U_SPLIT, rate >= y_split)]
        s = Series(Reference(data_ws, min_col=y_col, min_row=r, max_row=r),
                   Reference(data_ws, min_col=C_USAGE, min_row=r, max_row=r))
        # tx IS ASSIGNED, NOT PASSED AS title=. Series(title=SeriesLabel(...))
        # runs the object through str() and the label on the chart comes out as
        # "<openpyxl.chart.series.SeriesLabel object> Parameters: strRef=None,
        # v='Search'" -- printed, in the chart, with no error anywhere.
        s.tx = SeriesLabel(v=rows_cache()[i][0])
        s.marker = Marker(symbol="circle", size=9, spPr=GraphicalProperties(
            solidFill=colour, ln=LineProperties(solidFill="FFFFFF",
                                                w=int(1.0 * xs.PT))))
        s.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
        pos = nudge.get(rows_cache()[i][0], "r")
        s.dLbls = DataLabelList(
            showSerName=True, showVal=False, showCatName=False,
            showLegendKey=False, showPercent=False, showBubbleSize=False,
            dLblPos=pos, spPr=GraphicalProperties(noFill=True),
            txPr=rich(10, xs.INK))
        ch.series.append(s)

    # --- the four quadrant captions ---------------------------------------
    for k, (quad, (name, colour)) in enumerate(actions.items()):
        cx, cy, cpos = captions[quad]
        rr = h + 8 + k
        ws.cell(rr, C_NAME + 1, cx)
        ws.cell(rr, C_NAME + 2, cy)
        s = Series(Reference(ws, min_col=C_NAME + 2, min_row=rr, max_row=rr),
                   Reference(ws, min_col=C_NAME + 1, min_row=rr, max_row=rr))
        s.tx = SeriesLabel(v=name)
        s.marker = Marker(symbol="none")
        s.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
        s.dLbls = DataLabelList(
            showSerName=True, showVal=False, showCatName=False,
            showLegendKey=False, showPercent=False, showBubbleSize=False,
            dLblPos=cpos, spPr=GraphicalProperties(noFill=True),
            txPr=rich(caption_pt, colour, bold=True))
        ch.series.append(s)

    xs.declutter(ch, gridlines=False, x_line=False, gap=None)
    ch.legend = None
    ch.x_axis.scaling.min, ch.x_axis.scaling.max = x0, x1
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = y0, y1
    for ax, fmt in ((ch.x_axis, '0"%"'), (ch.y_axis, "0%")):
        ax.delete = False
        ax.numFmt = ax.number_format = fmt
        ax.majorGridlines = None
        ax.txPr = rich(11, xs.TEXT_2)
        ax.spPr = GraphicalProperties(
            ln=LineProperties(solidFill=xs.CONTEXT_FILL, w=int(0.75 * xs.PT)))
    xs.axis_title(ch.x_axis, "Reach — % of respondents who use the feature")
    ch.title = two_line_title(headline, subline, "0F3762")
    ch.width, ch.height = 26.0, 15.0
    return ch


def sorted_block(ws, rows, key, row0=40, col=C_NAME):
    """
    A re-ordered copy of the table, on the chart's own sheet, as LIVE formulas.

    A bar chart plots its categories in the order the source range holds them,
    so ranking a chart means ranking its source. This writes the same fifteen
    features sorted by `key`, every cell a formula back to Data, so the numbers
    stay live -- edit a response and the bar moves.

    THE ORDER IS FIXED AT BUILD TIME, the values are not. Ranking the rows with
    LARGE/MATCH instead would make the sort live too, and would also break: the
    dissatisfaction column has a four-way tie on 8, and MATCH would return the
    same feature four times. A stated order that is honest beats a formula that
    silently duplicates rows.
    """
    order = sorted(range(len(rows)), key=lambda i: -key(rows[i][1]))
    xs.write_table_header(ws, row0, ["Feature"] + CATS, col=col)
    for j, src in enumerate(order):
        r, sr = row0 + 1 + j, R0 + src
        for k in range(len(CATS) + 1):
            c = ws.cell(r, col + k, "='Data'!%s%d" % (
                get_column_letter(C_NAME + k), sr))
            c.font = Font(name=xs.FONT, size=11,
                          color=xs.INK if k == 0 else xs.TEXT_2)
            if k:
                c.number_format = '0"%"'
                c.alignment = Alignment(horizontal="right")
    return order, row0, row0 + len(rows)


def focus_chart(data_ws, last, focus, colour, headline, subline,
                greys=None, highlight=None, highlight_idx=None,
                keep=None, label_pos=None, src=None, order_rows=None):
    """
    Charts 2, 3 and 4 are one idea three times: name the answers you care
    about, plot them FIRST so every bar starts its block at zero, give them one
    colour, and let everything else fall back to grey. What changes between the
    three is only which answers are the subject.
    """
    # The remaining answers keep running in the direction they were already
    # going: a chart focused on the unhappy end continues worst-to-best, one
    # focused on the happy end continues best-to-worst. Ordering them by
    # whatever CATS happens to hold puts "completely satisfied" next to "not
    # satisfied at all" and the scale stops meaning anything.
    rest = [c for c in SCALE if c not in focus]
    if focus and focus[0] in SCALE and SCALE.index(focus[0]) > 2:
        rest.reverse()
    order = list(focus) + rest + (
        [] if "Have not used" in focus else ["Have not used"])
    colours = {c: GREYS.get(c, "BDBDBD") for c in CATS}
    colours.update(greys or {})
    for c in focus:
        colours[c] = colour
    return stacked_bar(
        data_ws, last, order, colours, set(focus),
        {c: "FFFFFF" for c in focus}, headline, subline,
        bold_legend=(),
        # NOT pinned. Sheet 2 pins its legend because that is what the
        # hand-edited chart does, and pays for it: six 12pt entries do not fit
        # the fixed box and check_chart_overlap.ps1 reports the legend clipped
        # by the chart edge. Widening the box to 0,99 did not help -- a pinned
        # legend cannot reflow, an automatic one can. These two are new charts
        # with nothing to reproduce, so they leave it automatic and come back
        # clean.
        headline_colour=colour, highlight=highlight, keep=keep,
        highlight_idx=highlight_idx, label_pos=label_pos, src=src,
        order_rows=order_rows)


def callouts(ws, first_row, col, width_cols):
    """
    The three call-out boxes beside chart 5, as merged worksheet cells.

    Excel charts have no annotation layer, and openpyxl cannot write a text-box
    shape at all -- so these are cells, which is also the form a reader can
    edit. Each block is 5 rows tall with a blank row between.
    """
    r = first_row
    for colour, text in CALLOUTS:
        ws.merge_cells(start_row=r, start_column=col,
                       end_row=r + 4, end_column=col + width_cols - 1)
        c = ws.cell(r, col, text)
        c.fill = PatternFill("solid", fgColor=colour)
        c.font = Font(name=xs.FONT, size=11, color="FFFFFF")
        c.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="left", indent=1)
        for rr in range(r, r + 5):
            for cc in range(col, col + width_cols):
                ws.cell(rr, cc).fill = PatternFill("solid", fgColor=colour)
        r += 6
    return r


DRAGGED = {
    # chart part -> {series name: (dx, dy)} as a fraction of the chart, from
    # the positions these labels were dragged to by hand in Excel.
    "6 Action map": {"Task comments": (-0.07056, 0.02822)},
    "7 Satisfaction map": {
        "Notifications": (-0.01221, 0.01411),
        "Calendar sync": (-0.02534, -0.04209),
        "Task comments": (-0.11397, -0.01646),
        "Time tracking": (-0.04749, 0.02117),
        "Custom fields": (-0.10458, -0.00494),
    },
}


def drag_labels(path, sheet_of_chart, dragged):
    """
    Move individual data labels off their default position.

    OPENPYXL'S DataLabel HAS NO `layout`. The OOXML <c:dLbl> takes a
    <c:layout><c:manualLayout> with x/y offsets, and openpyxl's class simply
    does not model it -- the same shape of gap as the dropped tickLblPos and
    the missing <c:delete>. So the offsets are spliced into the saved package.

    Each series here holds ONE point, so the dLbl is always idx 0, and the
    offsets are RELATIVE to wherever dLblPos already put the label.
    """
    src = zipfile.ZipFile(path)
    parts = {n: src.read(n) for n in src.namelist()}
    src.close()

    # which chart part belongs to which sheet, via the sheet's drawing
    names = re.findall(r'<sheet [^>]*name="([^"]+)"', parts["xl/workbook.xml"].decode())
    part_of = {}
    for i, nm in enumerate(names, 1):
        rels = parts.get("xl/worksheets/_rels/sheet%d.xml.rels" % i)
        if not rels:
            continue
        m = re.search(r"(drawings/drawing\d+\.xml)", rels.decode())
        if not m:
            continue
        dr = parts["xl/drawings/_rels/%s.rels" % m.group(1).split("/")[-1]].decode()
        for c in re.findall(r"(charts/chart\d+\.xml)", dr):
            part_of[nm] = "xl/" + c

    moved = 0
    for sheet, offsets in dragged.items():
        name = part_of.get(sheet)
        if not name:
            continue
        x = parts[name].decode()
        out = []
        pos = 0
        for m in re.finditer(r"<ser>.*?</ser>", x, re.S):
            block = m.group(0)
            title = re.search(r"<tx>\s*<v>([^<]*)</v>", block, re.S)
            if not title or title.group(1) not in offsets:
                continue
            dx, dy = offsets[title.group(1)]
            i = block.find("<dLbls>")
            if i == -1:
                continue
            # Schema order inside <dLbl>: idx, then layout, tx, numFmt, spPr,
            # txPr, dLblPos, then the show* flags. dLblPos AFTER the layout and
            # BEFORE the flags, or Excel rejects the part.
            lbl = ('<dLbl><idx val="0"/><layout><manualLayout>'
                   '<x val="%s"/><y val="%s"/></manualLayout></layout>'
                   '<dLblPos val="r"/>'
                   '<showLegendKey val="0"/><showVal val="0"/>'
                   '<showCatName val="0"/><showSerName val="1"/>'
                   '<showPercent val="0"/><showBubbleSize val="0"/></dLbl>'
                   % (dx, dy))
            j = i + len("<dLbls>")
            out.append(x[pos:m.start()] + block[:j] + lbl + block[j:])
            pos = m.end()
            moved += 1
        if out:
            parts[name] = ("".join(out) + x[pos:]).encode("utf-8")

    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, blob in parts.items():
            zout.writestr(n, blob)
    shutil.move(tmp, str(path))
    return moved


def main():
    rows = load()
    wb = Workbook()

    data = wb.active
    data.title = "Data"
    last = write_data(data, rows)

    # --- chart 1: every category, one blue ramp ----------------------------
    ws1 = wb.create_sheet("1 All six")
    xs.write_header(ws1, 2, TITLE_1, SUB_1)
    xs.note(ws1, 5,
            "Six ordered answers on one hue. Readable — but only the leftmost "
            "block starts at a common edge, so the two categories that matter "
            "float, and no two bars can be compared by length.")
    ch1 = stacked_bar(data, last, CATS, dict(zip(CATS, RAMP)),
                      set(CATS[2:]), dict(zip(CATS, RAMP_LABEL)),
                      TITLE_1, SUB_1)
    ws1.add_chart(ch1, "B8")
    ws1.sheet_view.showGridLines = False

    # --- chart 2: the two positive answers, pulled to the left edge --------
    ws2 = wb.create_sheet("2 Top two")
    xs.write_header(ws2, 2, TITLE_2, SUB_2)
    xs.note(ws2, 5,
            "Same numbers. 'Completely satisfied' is plotted first, so every "
            "positive block starts at zero and the bars can finally be "
            "compared to each other. Everything else recedes to grey.")
    order2 = ["Completely satisfied", "Very satisfied", "Somewhat satisfied",
              "Not very satisfied", "Not satisfied at all", "Have not used"]
    colours2 = dict(TOP2)
    colours2.update(GREYS)
    ch2 = stacked_bar(data, last, order2, colours2,
                      set(TOP2), {k: "FFFFFF" for k in TOP2},
                      # No bold_legend: the two blues already read as one
                      # block against four greys, and bolding their keys as
                      # well was emphasis spent twice on the same point.
                      TITLE_2, SUB_2, pin_legend=True)
    ws2.add_chart(ch2, "B8")
    ws2.sheet_view.showGridLines = False

    # --- chart 3: the two dissatisfied answers, in orange ------------------
    ws3 = wb.create_sheet("3 Dissatisfied")
    xs.write_header(ws3, 2, TITLE_3, SUB_3)
    xs.note(ws3, 5,
            "The same move as sheet 2, pointed at the other end of the scale. "
            "Both unhappy answers share one orange: the question is how much "
            "dissatisfaction there is, not how it splits between two grades.")
    # Ranked by the thing the chart is about: the two unhappy answers added
    # together. Unranked, a reader has to scan fifteen bars to find the worst;
    # ranked, the answer is the top row.
    order3, h3, l3 = sorted_block(ws3, rows, lambda v: v[1] + v[2])
    xs.note(ws3, h3 - 2,
            "Ranked copy of the table, most dissatisfied first. Live formulas "
            "back to Data — the values follow an edit, the ORDER does not.")
    ch3 = focus_chart(data, last, DISS, ORANGE, TITLE_3, SUB_3,
                      src=(ws3, h3, l3, C_NAME),
                      order_rows=[rows[i] for i in order3])
    ws3.add_chart(ch3, "B8")
    ws3.sheet_view.showGridLines = False

    # --- chart 4: never used, in teal --------------------------------------
    ws4 = wb.create_sheet("4 Not used")
    xs.write_header(ws4, 2, TITLE_4, SUB_4)
    xs.note(ws4, 5,
            "Non-use is not dissatisfaction, so it gets its own colour and its "
            "own chart. A feature half the base has never opened cannot be "
            "judged on its satisfaction score at all. API access is lit; the "
            "other fourteen bars are the same hue one step lighter, so the "
            "ranking is still readable without competing with the point.")
    order4, h4, l4 = sorted_block(ws4, rows, lambda v: v[0])
    xs.note(ws4, h4 - 2,
            "Ranked copy of the table, least used first. Live formulas back to "
            "Data — the values follow an edit, the ORDER does not.")
    # The highlight is an INDEX INTO THE PLOTTED ORDER, not into the CSV, so it
    # has to be looked up after the sort or the wrong bar lights up.
    api = order4.index(next(i for i, (n, _) in enumerate(rows)
                            if n == "API access"))
    ch4 = focus_chart(data, last, UNUSED, TEAL_BASE, TITLE_4, SUB_4,
                      greys=UNUSED_GREYS,
                      highlight={"Have not used": TEAL},
                      highlight_idx={"Have not used": {api}},
                      label_pos="inEnd",
                      src=(ws4, h4, l4, C_NAME),
                      order_rows=[rows[i] for i in order4])
    ws4.add_chart(ch4, "B8")
    ws4.sheet_view.showGridLines = False

    # --- chart 5: all six, annotated ---------------------------------------
    ws5 = wb.create_sheet("5 Annotated")
    xs.write_header(ws5, 2, TITLE_5, SUB_5)
    ch5 = stacked_bar(data, last, CATS, ANNOTATED, set(ANNOTATED_KEEP),
                      ANNOTATED_LABEL, TITLE_5, SUB_5,
                      keep=ANNOTATED_KEEP, width=19.5,
                      highlight=ANNOTATED_HIGHLIGHT,
                      legend_colours=ANNOTATED_LEGEND)
    ws5.add_chart(ch5, "B8")
    # MEASURED, not estimated: at 19,5 cm anchored on B the chart's
    # BottomRightCell is M34, so the call-outs start at O with a column to
    # spare. Excel reports this as Range(TopLeftCell, BottomRightCell) -- check
    # it again if the chart width or the column widths change, because a chart
    # drawn over the call-outs still exports without complaint.
    end = callouts(ws5, 9, 15, 4)          # column O, four columns wide
    xs.write_source(ws5, end + 1, FOOTNOTE, col=2)
    for col in "OPQR":
        ws5.column_dimensions[col].width = 13
    ws5.sheet_view.showGridLines = False

    # --- chart 6: reach vs dissatisfaction, four named actions -------------
    ws6 = wb.create_sheet("6 Action map")
    xs.write_header(ws6, 2, TITLE_6, SUB_6)
    xs.note(ws6, 5,
            "Custom fields and Automations sit on the SAME POINT here (73%, "
            "9.6%) — identical usage and identical dissatisfaction. Their "
            "labels are pushed apart; the markers genuinely coincide.")
    ch6 = action_map(ws6, data, last, C_DISS, D_SPLIT, DISS_ACTIONS, DISS_RANGE,
                     DISS_CAPTIONS, TITLE_6, SUB_6,
                     # Positions chosen against the rendered plot, not guessed.
                     # Every one of these is a pair that collided at the
                     # default "r": the cluster either side of the 10% line is
                     # six features inside four points of each other.
                     nudge={"Custom fields": "l", "Automations": "b",
                            "Time tracking": "t", "Task comments": "r",
                            "Templates": "t", "Dashboards": "b",
                            "Calendar sync": "t", "File sharing": "b",
                            "Mobile app": "t"})
    ws6.add_chart(ch6, "B8")
    ws6.sheet_view.showGridLines = False

    # --- chart 7: reach vs satisfaction ------------------------------------
    ws7 = wb.create_sheet("7 Satisfaction map")
    xs.write_header(ws7, 2, TITLE_7, SUB_7)
    xs.note(ws7, 5,
            "The same plane read from the other side. Note it is NOT the mirror "
            "image: satisfaction and dissatisfaction do not sum to one, because "
            "'somewhat satisfied' belongs to neither.")
    ch7 = action_map(ws7, data, last, C_SAT, S_SPLIT, SAT_ACTIONS, SAT_RANGE,
                     SAT_CAPTIONS, TITLE_7, SUB_7, caption_pt=14,
                     nudge={"Custom fields": "t", "Dashboards": "r",
                            "Calendar sync": "b", "Templates": "b",
                            "File sharing": "t", "Mobile app": "l"})
    ws7.add_chart(ch7, "B8")
    ws7.sheet_view.showGridLines = False

    wb.save(OUT)
    n = drag_labels(OUT, None, DRAGGED)
    print("wrote %s  (%d features, %d categories, %d charts, %d labels nudged)"
          % (OUT.name, len(rows), len(CATS), 7, n))


if __name__ == "__main__":
    main()
