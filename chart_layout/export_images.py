"""
Render the layout chart in chart_layout.xlsx to images/chart_layout.png.

build_xlsx.py already sizes the chart and its type for the final picture, so
this does nothing but export and then check what came out -- no scaling here.
Chart text is in points and does not follow the frame, so resizing the chart
object at export time would give more pixels and smaller labels; the scale
lives in build_xlsx.SCALE instead.

    python export_images.py
"""

import subprocess
from pathlib import Path

HERE = Path(__file__).resolve().parent
BOOK = HERE / "chart_layout.xlsx"
OUT = HERE / "images" / "chart_layout.png"

PS = r"""
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open("BOOK")
  $ws = $wb.Worksheets.Item("Data")
  $ws.ChartObjects("layout").Chart.Export("OUT", "PNG") | Out-Null
  $wb.Close($false)
  Write-Output "exported"
} finally {
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
"""


def check(png):
    """Catch a blank or mis-proportioned export here, not on a slide."""
    from PIL import Image
    im = Image.open(png).convert("RGB")
    w, h = im.size
    # Walk down the middle and measure each run of colour: the four bands must
    # come out at 12 / 8 / 75 / 5 percent of the height.
    col, runs, prev = w // 2, [], None
    for y in range(h):
        p = im.getpixel((col, y))
        if p != prev:
            runs.append([p, 0])
            prev = p
        runs[-1][1] += 1
    bands = [(p, n) for p, n in runs if n > h * 0.01]
    return (w, h), [(p, round(100 * n / h)) for p, n in bands]


def main():
    OUT.parent.mkdir(exist_ok=True)
    r = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command",
         PS.replace("BOOK", str(BOOK).replace("\\", "\\\\"))
           .replace("OUT", str(OUT).replace("\\", "\\\\"))],
        capture_output=True, text=True)
    print(r.stdout.strip() or r.stderr.strip())
    if r.returncode:
        raise SystemExit(r.returncode)

    size, bands = check(OUT)
    print("  %s  %sx%s  %d bytes" % (OUT.name, size[0], size[1], OUT.stat().st_size))
    for colour, pct in bands:
        print("    %-18s %3d%%" % (str(colour), pct))
    shares = [pct for _, pct in bands]
    if shares != [12, 8, 75, 5]:
        raise SystemExit("bands measured %s, expected [12, 8, 75, 5]" % shares)


if __name__ == "__main__":
    main()
