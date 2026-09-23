"""
The four-band chart layout, built as a real Excel chart.

Title 12% / subtitle 8% / field 75% / source line 5% -- the proportions
style/xl_style.py builds every chart in this course to. The diagram is a
single 100% stacked column with four segments, so the bands ARE the numbers:
change a share in the data sheet, rebuild, and the picture follows.

Two things worth knowing before changing this:

  * The labels sit hard left inside each band. Stacked-column data labels can
    only be centred, inside-base or inside-end within their segment -- there is
    no left -- so the four labels are chart TEXT BOXES, placed from the plot
    area's own geometry. Text boxes added to a chart do render above the plot,
    and Chart.Export keeps them.
  * The first attempt drew the bands as coloured worksheet cells and copied the
    range as a picture. Do not go back to that: an empty chart object still
    paints its plot area OVER anything pasted into it, and no amount of
    PlotArea.Interior / Format.Fill / ZOrder makes it stop. A real chart has no
    such problem.

openpyxl writes the data sheet; Excel itself adds the chart, because the text
boxes need the plot area's measured position, which only Excel can give.

    python build_xlsx.py
"""

import subprocess
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

HERE = Path(__file__).resolve().parent
OUT = HERE / "chart_layout.xlsx"

FONT = "Aptos"

# Three steps off the end of style_chart.md's sequential ramp, darkest at the
# title so the fills carry the same hierarchy the type does. The field keeps
# the reference figure's grey: it stands for the data, which is not one of the
# furniture bands and should not read as part of the ramp.
DEEP = 0x0F3762         # darkest step of the ramp
MID = 0x0F5499          # --accent
LIGHT = 0x3F73A6
FIELD_GREY = 0xCAD3D8

# Label ink is per band, because the three blues need white to stay legible
# (#0F3762 against near-black is 3:1 and unreadable at 13pt) while the grey
# field needs the dark.
WHITE = 0xFFFFFF
INK = 0x1A1A1A

# Bottom of the stack first -- that is the order a stacked column draws in.
# name, share of the height, fill, label ink, label point size at scale 1
BANDS = [
    ("Source line",  5, MID,        WHITE, 13),
    ("Field",       75, FIELD_GREY, INK,   14),
    ("Subtitle",     8, LIGHT,      WHITE, 14),
    ("Title",       12, DEEP,       WHITE, 14),
]

# The reference figure is 695 x 723 px; 519 x 540 pt is the same shape.
BASE_W, BASE_H = 519.0, 540.0
# Everything -- frame and type -- is multiplied by this, so the picture is the
# same picture at more pixels. Text in a chart is in points and does NOT follow
# the frame, which is why the sizes are scaled here rather than on export.
SCALE = 2
INDENT_PT = 7.0     # the small gap before the label, as in the reference


def rgb(hexval):
    """Excel's COM colour order is BGR, not RGB."""
    r, g, b = (hexval >> 16) & 255, (hexval >> 8) & 255, hexval & 255
    return r + (g << 8) + (b << 16)


def write_data():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.sheet_view.showGridLines = False
    ws["A1"], ws["B1"] = "Band", "Share"
    for c in ("A1", "B1"):
        ws[c].font = Font(name=FONT, bold=True)
    for i, (name, share, _, _, _) in enumerate(BANDS, start=2):
        ws.cell(row=i, column=1, value=name).font = Font(name=FONT)
        ws.cell(row=i, column=2, value=share / 100).number_format = "0%"
    ws.cell(row=len(BANDS) + 2, column=1, value="Total").font = Font(name=FONT, bold=True)
    t = ws.cell(row=len(BANDS) + 2, column=2, value="=SUM(B2:B%d)" % (len(BANDS) + 1))
    t.number_format = "0%"
    t.font = Font(name=FONT, bold=True)
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 10
    ws["A1"].alignment = Alignment(horizontal="left")
    wb.save(OUT)


PS = r"""
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
# Everything is wrapped so that a failure still QUITS EXCEL. Without this a
# broken run leaves an invisible instance holding the workbook open, and the
# next run cannot even overwrite the file.
try {{
$wb = $xl.Workbooks.Open("{book}")
$ws = $wb.Worksheets.Item("Data")

$W = {w}
$H = {h}
$co = $ws.ChartObjects().Add(220, 10, $W, $H)
$co.Name = "layout"
$ch = $co.Chart
$ch.ChartType = 52                       # xlColumnStacked

{series}

$ch.HasTitle = $false
$ch.HasLegend = $false
$ch.ChartGroups(1).GapWidth = 0
# The axes are HIDDEN, not deleted. Deleting them takes the value scale with
# them, and Excel then auto-picks a maximum of 1.2 for a stack that adds to 1 --
# the bands come out at 83% of the frame with white above, in the right
# proportions to each other but not to the picture. Pinning 0..1 is what makes
# the four percentages the real proportions of the image.
$vax = $ch.Axes(2)
$vax.MinimumScale = 0
$vax.MaximumScale = 1
if ($vax.HasMajorGridlines) {{ $vax.MajorGridlines.Delete() }}
foreach ($ax in @($ch.Axes(1), $vax)) {{
  $ax.TickLabelPosition = -4142          # xlTickLabelPositionNone
  $ax.MajorTickMark = -4142
  $ax.MinorTickMark = -4142
  $ax.Format.Line.Visible = $false
}}
$ch.ChartArea.Format.Line.Visible = $false
$ch.ChartArea.Format.Fill.ForeColor.RGB = 16777215
$ch.ChartArea.Format.Fill.Solid()
$ch.PlotArea.Format.Line.Visible = $false
$ch.PlotArea.Format.Fill.Visible = $false
$co.Border.LineStyle = -4142             # xlLineStyleNone

# With both axes gone the plot area can take the whole frame, so the bands run
# edge to edge and the percentages are the picture's real proportions.
$ch.PlotArea.Left = 0
$ch.PlotArea.Top = 0
$ch.PlotArea.Width = $W
$ch.PlotArea.Height = $H

$pl = $ch.PlotArea.InsideLeft
$pt = $ch.PlotArea.InsideTop
$pw = $ch.PlotArea.InsideWidth
$ph = $ch.PlotArea.InsideHeight

{labels}

$wb.Save()
$wb.Close($false)
Write-Output ("built: " + $W + " x " + $H + " pt")
}} finally {{
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}}
"""

SERIES = """
$s = $ch.SeriesCollection().NewSeries()
$s.Name = "{name}"
$s.Values = $ws.Range("B{row}")   # no $B$2: PowerShell would eat the $B
$s.Format.Fill.ForeColor.RGB = {fill}
$s.Format.Fill.Solid()
$s.Format.Line.Visible = $false
""".strip()

# msoTextOrientationHorizontal = 1. Anchor 3 = middle, align 1 = left.
LABEL = """
$tb = $ch.Shapes.AddTextbox(1, $pl + {left}, $pt + $ph * {top}, $pw - {left}, $ph * {frac})
$tb.Fill.Visible = $false
$tb.Line.Visible = $false
$tf = $tb.TextFrame2
$tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
$tf.VerticalAnchor = 3
$tf.WordWrap = 0
$tr = $tf.TextRange
$tr.Text = "{text}"
$tr.ParagraphFormat.Alignment = 1
$tr.Font.Name = "{font}"
$tr.Font.Size = {size}
$tr.Font.Fill.ForeColor.RGB = {ink}
$tr.Characters(1, {bold_len}).Font.Bold = -1
""".strip()


def main():
    assert sum(b[1] for b in BANDS) == 100, "the four bands must add to 100%"
    write_data()

    w, h = BASE_W * SCALE, BASE_H * SCALE
    series = "\n".join(
        SERIES.format(name=name, row=i, fill=rgb(fill))
        for i, (name, _, fill, _, _) in enumerate(BANDS, start=2))

    # Text boxes are placed from the TOP, so walk the stack downwards: the last
    # band in BANDS is the top one.
    labels, top = [], 0.0
    for name, share, _, ink, size in reversed(BANDS):
        labels.append(LABEL.format(
            left=INDENT_PT * SCALE, top=round(top, 6), frac=share / 100,
            text="%s %d%%" % (name, share), font=FONT, size=size * SCALE,
            ink=rgb(ink), bold_len=len(name)))
        top += share / 100
    assert abs(top - 1.0) < 1e-9

    r = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command",
         PS.format(book=str(OUT).replace("\\", "\\\\"), w=w, h=h,
                   series=series, labels="\n".join(labels))],
        capture_output=True, text=True)
    print(r.stdout.strip() or r.stderr.strip())
    if r.returncode:
        raise SystemExit(r.returncode)
    print("wrote %s" % OUT.name)


if __name__ == "__main__":
    main()
