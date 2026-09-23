"""
Export the workbook's charts as PNGs at N times screen resolution.

Excel's Chart.Export writes at ~96 dpi, sized from the chart frame, so a chart
that looks right on a worksheet comes out around 900px wide — soft the moment
it is projected.

Enlarging the frame alone does not work. Chart text is specified in POINTS and
does not rescale with the frame, and ChartArea.AutoScaleFont does not change
that: you get three times the pixels with all the type still at its original
size, so every label shrinks to a third of its intended size against the plot.

So this scales the chart XML itself before exporting:

  * every  sz="..."  (font size, hundredths of a point) x SCALE
  * every  <a:ln w="...">  (line width, EMU) x SCALE

then Excel opens that copy, sets each frame x SCALE, and exports. Fonts, line
weights and frame all grow together, so the layout is pixel-for-pixel the same
picture — just rendered with SCALE times the detail.

The scaled copy is a throwaway. The real workbook is never modified.

    python export_images.py           # 3x
    python export_images.py --scale 4
"""

import argparse
import re
import shutil
import subprocess
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
BOOK = HERE / "daily_makeup_use.xlsx"
PRINT_COPY = HERE / "_print_copy.xlsx"
OUT = HERE / "images"

# sheet name -> output file
CHARTS = [
    ("Makeup", "makeup_use.png"),
    ("Vertical", "makeup_vertical.png"),
    ("Slope", "makeup_slope.png"),
    ("Dot plot", "makeup_dotplot.png"),
    ("Wrong chart type", "wrong_line_chart.png"),
]
# the raw-data sheet is a cell range, not a chart: captured at a zoom instead
# B:D only — Type, 2024, 2019. Column E is the computed Change (pp), which is
# the chart's job on the slide, not the table's.
TABLE = ("Makeup", "B8:D14", "data_table.png")

CHART_PART = re.compile(r"xl/charts/chart\d+\.xml$")


def scale_chart_xml(src, dst, scale):
    """Copy the package, multiplying every font size and line width by `scale`."""
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    touched = 0
    for name in list(parts):
        if not CHART_PART.match(name):
            continue
        x = parts[name].decode("utf-8")
        # sz is in hundredths of a point, on <a:defRPr>/<a:rPr>/<a:endParaRPr>
        x2 = re.sub(r'\bsz="(\d+)"',
                    lambda m: 'sz="%d"' % (int(m.group(1)) * scale), x)
        # line width in EMU, only on <a:ln ...> so the layout <w val=".8"/> is
        # left alone — that one is a fraction of the chart and must not scale
        x2 = re.sub(r'(<a:ln\b[^>]*?\bw=")(\d+)(")',
                    lambda m: m.group(1) + str(int(m.group(2)) * scale) + m.group(3),
                    x2)
        if x2 != x:
            parts[name] = x2.encode("utf-8")
            touched += 1

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in parts.items():
            zout.writestr(name, blob)
    return touched


PS = r"""
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open("{book}")
$scale = {scale}
{chart_lines}
# the table is a range: raise the window zoom so the screen capture has more
# pixels, then copy it as a picture and paste it into a throwaway chart
$ws = $wb.Worksheets.Item("{tsheet}")
$ws.Activate()
$xl.ActiveWindow.Zoom = 100 * $scale
$rng = $ws.Range("{trange}")
$rng.Select() | Out-Null
$rng.CopyPicture(1, 2)
Start-Sleep -Milliseconds 900
$cho = $ws.ChartObjects().Add(0, 0, $rng.Width, $rng.Height)
$cho.Chart.ChartArea.Border.LineStyle = -4142
$cho.Activate()
$cho.Chart.Paste()
Start-Sleep -Milliseconds 900
$cho.Chart.Export("{tout}") | Out-Null
$cho.Delete()
$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
"""

CHART_LINE = r"""
$co = $wb.Worksheets.Item("{sheet}").ChartObjects(1)
$co.Width = $co.Width * $scale; $co.Height = $co.Height * $scale
$co.Chart.Export("{out}") | Out-Null
Write-Output "exported {out_name}"
""".strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--scale", type=int, default=3)
    a = ap.parse_args()

    OUT.mkdir(exist_ok=True)
    touched = scale_chart_xml(BOOK, PRINT_COPY, a.scale)
    print("scaled type and line weights in %d chart part(s)" % touched)

    lines = "\n".join(
        CHART_LINE.format(sheet=s, out=str(OUT / f).replace("\\", "\\\\"),
                          out_name=f)
        for s, f in CHARTS)
    script = PS.format(
        book=str(PRINT_COPY).replace("\\", "\\\\"),
        scale=a.scale, chart_lines=lines,
        tsheet=TABLE[0], trange=TABLE[1],
        tout=str(OUT / TABLE[2]).replace("\\", "\\\\"))

    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive",
                        "-Command", script],
                       capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode:
        print(r.stderr.strip())
    PRINT_COPY.unlink(missing_ok=True)
    for _, f in CHARTS:
        print("  %-14s %s" % (f, (OUT / f).stat().st_size))


if __name__ == "__main__":
    main()
