"""
Office -> Team -> Employee, as a real Excel sunburst.

A sunburst is a chartEx type: it is NOT in the original SpreadsheetML chart
schema and openpyxl cannot write one. So openpyxl builds the sheet, and the
saved .xlsx is reopened as a zip and a chartEx part is injected by hand. The
recipe is in .claude/skills/excel-charts-openpyxl/references/chartex-charts.md;
what that reference did not cover, and what Excel's own sunburst turned out to
need, is written up in NOTES.md next to this file.

The short version of the one thing that matters:

    <cx:strDim type="cat"><cx:f>'Sunburst'!$B$9:$D$20</cx:f></cx:strDim>

ONE category dimension over a THREE-COLUMN range. The hierarchy levels are
columns of a single range, not one strDim per ring. Point it at a single column
and Excel draws a doughnut with one ring and no error.

THE HEADCOUNT COLUMN IS ADDED, NOT GIVEN. The source table is three text
columns with no measure in it, and a sunburst has to size its wedges by
something. Each employee counts 1, so the rings read as headcount: 12 people,
6 teams, 3 offices. Nothing is invented -- the column is the row count made
explicit -- but it is a choice and the sheet says so.

    python build_xlsx.py     ->  office_headcount.xlsx
"""

import csv
import re
import shutil
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE / "style"))
import xl_style as xs  # noqa: E402

OUT = HERE / "office_headcount.xlsx"
ASSETS = HERE / "assets"
SHEET = "Sunburst"

# One hue per office on a descending-lightness ramp, largest office darkest --
# an ordered layout deserves an ordered ramp, and a categorical rainbow fights
# it. Not grey-plus-one-accent, because no single office is the story.
#
# All three are dark enough to carry WHITE labels: contrast against white is
# 11,8:1 / 7,6:1 / 5,0:1, so one text colour serves the whole chart and there
# is no per-label colour switch to compute. Going one step lighter (93B2D1,
# 2,1:1) would have forced ink labels on the pale office and white on the rest.
OFFICE_COLOUR = {"Boston": "0F3762", "New York": "0F5499", "Miami": "3F73A6"}

HDR = 8                       # table header row
R0 = HDR + 1                  # first data row
C_OFFICE, C_TEAM, C_EMP, C_HEAD = 2, 3, 4, 5      # B C D E

TITLE = "Boston holds half the company's headcount"
SUBTITLE = "People by office, team and individual — 12 staff in 3 offices"
SOURCE = ""                   # the roster is illustrative; no outside source

# ---------------------------------------------------------------- namespaces
CX = "http://schemas.microsoft.com/office/drawing/2014/chartex"
CX1 = "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
REL_CHARTEX = "http://schemas.microsoft.com/office/2014/relationships/chartEx"
REL_COLORS = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle"
REL_STYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle"
REL_DRAWING = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/drawing")


def load():
    with open(HERE / "employees.csv", encoding="utf-8") as fh:
        return [(r["Office"], r["Team"], r["Employee"], int(r["Headcount"]))
                for r in csv.DictReader(fh)]


# --------------------------------------------------------------- the sheet
def write_sheet(ws, rows):
    xs.write_header(ws, 2, TITLE, SUBTITLE)
    xs.note(ws, 5,
            "Headcount is one per person — the row count made explicit, so the "
            "sunburst has a measure to size its wedges by. The source table is "
            "three text columns and carries no measure of its own.")

    xs.write_table_header(ws, HDR, ["Office", "Team", "Employee", "Headcount"])
    for i, (office, team, emp, head) in enumerate(rows):
        r = R0 + i
        ws.cell(r, C_OFFICE, office).font = Font(name=xs.FONT, size=11, color=xs.INK)
        ws.cell(r, C_TEAM, team).font = Font(name=xs.FONT, size=11, color=xs.INK)
        ws.cell(r, C_EMP, emp).font = Font(name=xs.FONT, size=11, color=xs.INK)
        c = ws.cell(r, C_HEAD, head)
        c.font = Font(name=xs.FONT, size=11, color=xs.TEXT_2)
        c.alignment = Alignment(horizontal="right")
        c.number_format = "0"

    last = R0 + len(rows) - 1
    t = ws.cell(last + 1, C_OFFICE, "Total")
    t.font = Font(name=xs.FONT, size=11, bold=True, color=xs.INK)
    tot = ws.cell(last + 1, C_HEAD, "=SUM(E%d:E%d)" % (R0, last))
    tot.font = Font(name=xs.FONT, size=11, bold=True, color=xs.INK)
    tot.alignment = Alignment(horizontal="right")

    if SOURCE:
        xs.write_source(ws, last + 3, SOURCE)
    xs.widths(ws, [12, 11, 15, 12])
    ws.sheet_view.showGridLines = False
    return last


def wedge_order():
    """
    Every wedge the sunburst will draw, in the order cx:dataPt indexes them.

    THIS IS THE ONE THING ABOUT COLOURING A SUNBURST. `cx:dataPt idx` does NOT
    index the data rows -- it walks EVERY wedge of EVERY ring, depth first:

        0 Boston            <- ring 1
        1 Boston/Sales      <- ring 2
        2 Lisa T.           <- ring 3
        3 Phil B.
        4 Marc G.
        5 Boston/IT         <- back out to ring 2
        ...
       21 Dennis M.         22 wedges from 12 rows

    Index them as 12 leaves (which is what a treemap would want, and what
    `chartex-charts.md` describes) and the colours land on the wrong wedges
    while the chart still renders perfectly happily. Anything past the last
    dataPt falls back to the theme palette, which is how a green Miami appeared
    in a blue chart.
    """
    order, seen_office, seen_team = [], [], []
    for office, team, emp, _ in load():
        if office not in seen_office:
            seen_office.append(office)
            order.append((office, office))
        if (office, team) not in seen_team:
            seen_team.append((office, team))
            order.append((office, team))
        order.append((office, emp))
    return order


# ------------------------------------------------------------- the chartEx
def chartex_xml(last):
    """
    The sunburst part. Modelled on what Excel itself writes for xlChartType
    120, with two deliberate differences: plain ranges instead of hidden
    _xlchart.v1.N defined names, and per-point fills.
    """
    cat = "'%s'!$%s$%d:$%s$%d" % (SHEET, "B", R0, "D", last)
    size = "'%s'!$E$%d:$E$%d" % (SHEET, R0, last)
    name = "'%s'!$E$%d" % (SHEET, HDR)

    pts = []
    for i, (office, _) in enumerate(wedge_order()):
        pts.append(
            '<cx:dataPt idx="%d"><cx:spPr><a:solidFill><a:srgbClr val="%s"/>'
            '</a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="%s"/>'
            '</a:solidFill></a:ln></cx:spPr></cx:dataPt>'
            % (i, OFFICE_COLOUR[office], xs.SURFACE))

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cx:chartSpace xmlns:a="%s" xmlns:r="%s" xmlns:cx="%s">'
        '<cx:chartData><cx:data id="0">'
        '<cx:strDim type="cat"><cx:f>%s</cx:f></cx:strDim>'
        '<cx:numDim type="size"><cx:f>%s</cx:f></cx:numDim>'
        '</cx:data></cx:chartData>'
        '<cx:chart>'
        # A cx:title on the CHART is fine; one on an AXIS makes Excel reject
        # the file. A sunburst has no axes, so the question does not arise here.
        #
        # align="ctr" IS NOT A STYLE CHOICE. cx:title/@align rejects "l" -- and
        # rejecting it means Excel refuses the whole workbook, with the same
        # "unable to get the Open property" that a malformed package gives.
        # Left-align the TEXT with algn="l" on the paragraph below instead;
        # that one is accepted.
        '<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich>'
        '<a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" '
        'vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/>'
        '<a:p><a:pPr algn="l"><a:defRPr sz="1400" b="1">'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:latin typeface="%s"/></a:defRPr></a:pPr>'
        '<a:r><a:rPr lang="en-US" sz="1400" b="1">'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:latin typeface="%s"/></a:rPr><a:t>%s</a:t></a:r>'
        '</a:p></cx:rich></cx:tx></cx:title>'
        '<cx:plotArea><cx:plotAreaRegion>'
        '<cx:series layoutId="sunburst" uniqueId="'
        '{3C1E9A57-2B44-4D0E-9F1A-7A6C5E2D8B41}">'
        '<cx:tx><cx:txData><cx:f>%s</cx:f><cx:v>Headcount</cx:v></cx:txData></cx:tx>'
        '%s'
        '<cx:dataLabels pos="ctr">'
        '<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="0">'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:latin typeface="%s"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/>'
        '</a:p></cx:txPr>'
        '<cx:visibility seriesName="0" categoryName="1" value="0"/>'
        '</cx:dataLabels>'
        '<cx:dataId val="0"/>'
        '</cx:series>'
        '</cx:plotAreaRegion></cx:plotArea>'
        '</cx:chart>'
        '<cx:spPr><a:noFill/><a:ln><a:noFill/></a:ln></cx:spPr>'
        '</cx:chartSpace>'
        % (A_NS, R_NS, CX, cat, size,
           xs.INK, xs.FONT, xs.INK, xs.FONT, TITLE,
           name, "".join(pts), xs.SURFACE, xs.FONT))


FALLBACK_TEXT = ("Sunburst charts need Excel 2016 or later. The same figures are "
                 "in the table to the left: 12 people across 3 offices.")

DRAWING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<xdr:wsDr xmlns:xdr="%s" xmlns:a="%s">'
    '<xdr:twoCellAnchor>'
    '<xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff>'
    '<xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    '<xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff>'
    '<xdr:row>29</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
    # mc:AlternateContent wraps ONLY the graphicFrame and sits INSIDE the
    # anchor. Wrapping the whole twoCellAnchor makes the workbook unopenable.
    '<mc:AlternateContent xmlns:mc="%s">'
    '<mc:Choice xmlns:cx1="%s" Requires="cx1">'
    '<xdr:graphicFrame macro="">'
    '<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Sunburst 1"/>'
    '<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
    '<a:graphic><a:graphicData uri="%s">'
    '<cx:chart xmlns:cx="%s" xmlns:r="%s" r:id="%%s"/>'
    '</a:graphicData></a:graphic>'
    '</xdr:graphicFrame>'
    '</mc:Choice>'
    '<mc:Fallback>'
    '<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/>'
    '<xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>'
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="4572000"/>'
    '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    '<a:ln w="9525"><a:solidFill><a:srgbClr val="BDBDBD"/></a:solidFill></a:ln>'
    '</xdr:spPr>'
    '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square"/>'
    '<a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>%s</a:t></a:r>'
    '</a:p></xdr:txBody></xdr:sp>'
    '</mc:Fallback>'
    '</mc:AlternateContent>'
    '<xdr:clientData/>'
    '</xdr:twoCellAnchor></xdr:wsDr>'
    % (XDR_NS, A_NS, MC_NS, CX1, CX, CX, R_NS, FALLBACK_TEXT))


def inject(path, last):
    """
    Reopen the saved package and add the chartEx chart to it.

    The sheet has no drawing at all -- openpyxl only writes one when it has an
    ordinary chart to put in it -- so the drawing part, its rel, the sheet's
    rel to it and the <drawing> element in the sheet are all created here too.
    """
    style = ASSETS / "chartex_style1.xml"
    colors = ASSETS / "chartex_colors1.xml"
    # style1.xml is MANDATORY for a chartEx part. Without it Excel refuses to
    # open the whole workbook, not just the chart.
    assert style.exists(), "missing %s -- chartEx will not open without it" % style
    assert colors.exists(), "missing %s" % colors

    zin = zipfile.ZipFile(path)
    parts = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    parts["xl/charts/chartEx1.xml"] = chartex_xml(last).encode("utf-8")
    parts["xl/charts/style1.xml"] = style.read_bytes()
    parts["xl/charts/colors1.xml"] = colors.read_bytes()
    parts["xl/charts/_rels/chartEx1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships">'
        '<Relationship Id="rId1" Type="%s" Target="style1.xml"/>'
        '<Relationship Id="rId2" Type="%s" Target="colors1.xml"/>'
        '</Relationships>' % (REL_STYLE, REL_COLORS)).encode("utf-8")

    parts["xl/drawings/drawing1.xml"] = (DRAWING % "rId1").encode("utf-8")
    parts["xl/drawings/_rels/drawing1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships">'
        '<Relationship Id="rId1" Type="%s" Target="../charts/chartEx1.xml"/>'
        '</Relationships>' % REL_CHARTEX).encode("utf-8")

    # which sheet part is ours: sheet order in workbook.xml gives the index
    names = re.findall(r'<sheet [^>]*name="([^"]+)"', parts["xl/workbook.xml"].decode())
    idx = names.index(SHEET) + 1
    sheet_part = "xl/worksheets/sheet%d.xml" % idx
    rels_part = "xl/worksheets/_rels/sheet%d.xml.rels" % idx

    rel = ('<Relationship Id="rIdDrawing" Type="%s" Target="../drawings/'
           'drawing1.xml"/>' % REL_DRAWING)
    if rels_part in parts:
        parts[rels_part] = parts[rels_part].decode().replace(
            "</Relationships>", rel + "</Relationships>").encode("utf-8")
    else:
        parts[rels_part] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            'relationships">%s</Relationships>' % rel).encode("utf-8")

    # <drawing> is near the END of CT_Worksheet's sequence -- after pageMargins,
    # immediately before </worksheet>.
    #
    # xmlns:r IS DECLARED ON THE ELEMENT ITSELF, and that is load-bearing.
    # openpyxl writes the worksheet root as a bare
    #     <worksheet xmlns="...spreadsheetml/2006/main">
    # with no xmlns:r, because a sheet it built has no relationship-bearing
    # element to need one. Write r:id here without declaring the prefix and the
    # part is not well-formed XML, so Excel refuses to open THE WHOLE WORKBOOK
    # with "unable to get the Open property" -- no mention of a chart, and the
    # chartEx part is not the cause. Declaring it locally is valid XML and
    # leaves openpyxl's own output untouched.
    sx = parts[sheet_part].decode()
    assert "<drawing " not in sx, "sheet already has a drawing"
    parts[sheet_part] = sx.replace(
        "</worksheet>",
        '<drawing xmlns:r="%s" r:id="rIdDrawing"/></worksheet>' % R_NS
    ).encode("utf-8")

    ct = parts["[Content_Types].xml"].decode()
    for part, mime in (
            ("/xl/drawings/drawing1.xml",
             "application/vnd.openxmlformats-officedocument.drawing+xml"),
            ("/xl/charts/chartEx1.xml", "application/vnd.ms-office.chartex+xml"),
            ("/xl/charts/style1.xml", "application/vnd.ms-office.chartstyle+xml"),
            ("/xl/charts/colors1.xml",
             "application/vnd.ms-office.chartcolorstyle+xml")):
        if part not in ct:
            ct = ct.replace("</Types>", '<Override PartName="%s" ContentType="%s"/>'
                            "</Types>" % (part, mime))
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, blob in parts.items():
            zout.writestr(n, blob)
    shutil.move(tmp, str(path))


def main():
    rows = load()
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    last = write_sheet(ws, rows)
    wb.save(OUT)
    inject(OUT, last)
    print("wrote %s  (%d rows, %d offices)"
          % (OUT.name, len(rows), len({r[0] for r in rows})))


if __name__ == "__main__":
    main()
