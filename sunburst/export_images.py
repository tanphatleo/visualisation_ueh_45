"""
Export the sheet's two charts and the data table as PNGs at N times resolution.

Same problem and same fix as ../music_industry/export_images.py: Excel's
Chart.Export writes at ~96 dpi sized from the chart frame, and enlarging the
frame alone does not help, because chart text is specified in POINTS and does
not rescale with it -- you get three times the pixels with all the type still
at its original size. So the chart XML is scaled first:

  * every  sz="..."       (font size, hundredths of a point)  x SCALE
  * every  <a:ln w="...">  (line width, EMU)                  x SCALE

then Excel opens that copy, enlarges each frame by the same factor, and
exports. Type, rules and frame grow together, so the picture is identical and
simply carries SCALE times the detail.

ONE DIFFERENCE FROM THE ORDINARY-CHART VERSION: the sunburst lives in
xl/charts/chartEx1.xml, not chart1.xml, so the part pattern has to match both.
Miss that and the sunburst exports with third-size labels while the bar chart
next to it comes out right -- which looks like a sunburst problem and is not.

The table is a cell range, not a chart, so it is captured the same way the
music_industry raw table is: raise the window zoom so the screen copy has more
pixels, CopyPicture, and paste into a throwaway chart object.

    python export_images.py           # 3x
    python export_images.py --scale 4
"""

import argparse
import re
import subprocess
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
BOOK = HERE / "office_headcount.xlsx"
PRINT_COPY = HERE / "_print_copy.xlsx"
OUT = HERE / "images"
SHEET = "Sunburst"

# shape index on the sheet -> output file
CHARTS = [(1, "sunburst.png"), (2, "bar.png")]
# the table is a range, captured at a zoom instead
TABLE = ("B8:E20", "table.png")

# chartEx parts as well as ordinary chart parts
CHART_PART = re.compile(r"xl/charts/chart(?:Ex)?\d+\.xml$")


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
        x2 = re.sub(r'\bsz="(\d+)"',
                    lambda m: 'sz="%d"' % (int(m.group(1)) * scale), x)
        # only on <a:ln ...>, so any fractional layout width is left alone
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
# Visible: CopyPicture below races the clipboard when Excel is hidden.
$xl.Visible = $true; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open("{book}")
$ws = $wb.Worksheets.Item("{sheet}")
$scale = {scale}
{chart_lines}
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
Write-Output "exported {tout_name}"
$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
"""

CHART_LINE = r"""
$sh = $ws.Shapes.Item({idx})
$sh.Width = $sh.Width * $scale; $sh.Height = $sh.Height * $scale
$sh.Chart.Export("{out}") | Out-Null
Write-Output "exported {out_name}"
""".strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--scale", type=int, default=3)
    a = ap.parse_args()

    OUT.mkdir(exist_ok=True)
    n = scale_chart_xml(BOOK, PRINT_COPY, a.scale)
    print("scaled type and line weights in %d chart part(s)" % n)
    assert n >= 2, "expected both chart parts to scale -- chartEx1 AND chart1"

    lines = "\n".join(
        CHART_LINE.format(idx=i, out=str(OUT / f).replace("\\", "\\\\"), out_name=f)
        for i, f in CHARTS)
    script = PS.format(book=str(PRINT_COPY).replace("\\", "\\\\"),
                       sheet=SHEET, scale=a.scale, chart_lines=lines,
                       trange=TABLE[0],
                       tout=str(OUT / TABLE[1]).replace("\\", "\\\\"),
                       tout_name=TABLE[1])
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive",
                        "-Command", script], capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode:
        print(r.stderr.strip())
    PRINT_COPY.unlink(missing_ok=True)
    for f in [f for _, f in CHARTS] + [TABLE[1]]:
        p = OUT / f
        print("  %-14s %s" % (f, p.stat().st_size if p.exists() else "MISSING"))


if __name__ == "__main__":
    main()
