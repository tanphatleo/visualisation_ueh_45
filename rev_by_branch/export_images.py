"""
Export the four charts on the 'Answer_' sheet as PNGs at N times screen resolution.

Same approach as ../music_industry/export_images.py: Excel's Chart.Export writes
at ~96 dpi sized from the chart frame, and simply enlarging the frame does not
help because chart text is specified in POINTS and does not rescale with it. So
the chart XML is scaled first -- every sz="" and every <a:ln w=""> -- and Excel
then exports a copy whose frame, type and line weights have all grown together.

The scaled copy is a throwaway. The real workbook is never modified.

    python export_images.py            # 3x
    python export_images.py --scale 4
"""

import argparse
import re
import subprocess
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
BOOK = HERE / "HatGao_Questions_F_Q1.xlsx"
PRINT_COPY = HERE / "_print_copy.xlsx"
OUT = HERE / "images"
SHEET = "Answer_"

# ChartObjects index on the sheet -> output file. Index order follows the
# drawing order, which is the order the four were added: the cluttered
# original first, then each successive pass.
CHARTS = [(1, "rev_01.png"), (2, "rev_02.png"), (3, "rev_03.png"), (4, "rev_04.png")]

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
        x2 = re.sub(r'\bsz="(\d+)"',
                    lambda m: 'sz="%d"' % (int(m.group(1)) * scale), x)
        # only on <a:ln ...> so a manual layout's <w val="0.8"/> is left alone
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
$ws = $wb.Worksheets.Item("{sheet}")
$scale = {scale}
Write-Output ("chart objects: " + $ws.ChartObjects().Count)
{chart_lines}
$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
"""

CHART_LINE = r"""
$co = $ws.ChartObjects({idx})
$co.Width = $co.Width * $scale; $co.Height = $co.Height * $scale
$co.Chart.Export("{out}") | Out-Null
Write-Output "exported {out_name}"
""".strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--scale", type=int, default=3)
    a = ap.parse_args()

    OUT.mkdir(exist_ok=True)
    print("scaled type and line weights in %d chart part(s)"
          % scale_chart_xml(BOOK, PRINT_COPY, a.scale))

    lines = "\n".join(
        CHART_LINE.format(idx=i, out=str(OUT / f).replace("\\", "\\\\"), out_name=f)
        for i, f in CHARTS)
    script = PS.format(book=str(PRINT_COPY).replace("\\", "\\\\"),
                       sheet=SHEET, scale=a.scale, chart_lines=lines)

    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive",
                        "-Command", script], capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode:
        print(r.stderr.strip())
    PRINT_COPY.unlink(missing_ok=True)
    for _, f in CHARTS:
        p = OUT / f
        print("  %-12s %s" % (f, p.stat().st_size if p.exists() else "MISSING"))


if __name__ == "__main__":
    main()
