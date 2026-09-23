"""
Export both charts as PNGs at N times screen resolution.

Same approach as ../music_industry/export_images.py: scale every sz="" and
<a:ln w=""> in the chart XML first, then let Excel enlarge the frame by the
same factor and export. Chart text is in POINTS and does not rescale with the
frame, so without this you get N times the pixels and third-size labels.

    python export_images.py           # 3x
    python export_images.py --scale 4
"""

import argparse
import re
import subprocess
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
BOOK = HERE / "feature_satisfaction.xlsx"
PRINT_COPY = HERE / "_print_copy.xlsx"
OUT = HERE / "images"

CHARTS = [("1 All six", "all_six.png"), ("2 Top two", "top_two.png"),
          ("3 Dissatisfied", "dissatisfied.png"), ("4 Not used", "not_used.png"),
          ("5 Annotated", "annotated_chart.png"),
          ("6 Action map", "action_map.png"),
          ("7 Satisfaction map", "satisfaction_map.png")]
# Sheet 5's call-outs are worksheet cells, not part of the chart, so
# Chart.Export alone loses them. They are DRAWN HERE instead of captured.
#
# WHY NOT CopyPicture: CopyPicture(xlScreen) copies WHAT IS ON SCREEN. Raise
# the zoom to get more pixels and the range stops fitting in the window, so the
# picture comes back silently cropped -- one attempt returned six of fifteen
# bars and no call-outs, another returned a slice of the chart where the
# call-out cells should have been, neither with an error. The text and the
# colours are already in build_xlsx.CALLOUTS, so drawing them is exact at any
# scale and has no clipboard race.
COMPOSITE_OUT = "annotated.png"
FONT_FILE = "C:/Windows/Fonts/segoeui.ttf"
CHART_PART = re.compile(r"xl/charts/chart(?:Ex)?\d+\.xml$")


DLBLS = re.compile(r"<(?:c:)?dLbls>.*?</(?:c:)?dLbls>", re.S)
MARKER = re.compile(r"<(?:c:)?marker>.*?</(?:c:)?marker>", re.S)
MARKER_SIZE = re.compile(r'(<(?:c:)?size val=")(\d+)(")')
DEFRPR_NO_SZ = re.compile(r"<a:defRPr(?![^>]*\bsz=)")


def pin_label_sizes(x, default_pt=10):
    """
    Give every data label an explicit size before anything is scaled.

    A label with no sz="" inherits the chart's default, which is what the
    workbook wants: resize the chart and the labels follow. But this script
    scales type by rewriting sz="" attributes, and AN ATTRIBUTE THAT IS NOT
    THERE CANNOT BE MULTIPLIED -- the labels stay at 10pt while title, legend
    and axis all triple, and the export comes back with unreadable numbers
    inside enormous bars. So the inherited size is written out first, in the
    throwaway copy only.
    """
    return DLBLS.sub(
        lambda m: DEFRPR_NO_SZ.sub('<a:defRPr sz="%d"' % (default_pt * 100),
                                   m.group(0)), x)


def scale_markers(x, scale):
    """
    Scatter marker size is a THIRD unit this script has to scale.

    It is neither a font size nor a line width: <c:size val="9"/> inside
    <c:marker>, in points, capped at 72. Scale sz and a:ln w only and the
    scatter charts come back with type and rules three times bigger around
    9pt dots that now read as specks. Only touched inside a marker block --
    "size" is a common enough element name to be worth the guard.
    """
    return MARKER.sub(
        lambda m: MARKER_SIZE.sub(
            lambda s: s.group(1) + str(min(72, int(s.group(2)) * scale)) + s.group(3),
            m.group(0)), x)


def scale_chart_xml(src, dst, scale):
    zin = zipfile.ZipFile(src)
    parts = {n: zin.read(n) for n in zin.namelist()}
    zin.close()
    touched = 0
    for name in list(parts):
        if not CHART_PART.match(name):
            continue
        x = parts[name].decode("utf-8")
        x2 = pin_label_sizes(x)
        x2 = scale_markers(x2, scale)
        x2 = re.sub(r'\bsz="(\d+)"',
                    lambda m: 'sz="%d"' % (int(m.group(1)) * scale), x2)
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
{lines}
$wb.Close($false); $xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
"""

LINE = r"""
$co = $wb.Worksheets.Item("{sheet}").ChartObjects(1)
$co.Width = $co.Width * $scale; $co.Height = $co.Height * $scale
$co.Chart.Export("{out}") | Out-Null
Write-Output "exported {out_name}"
""".strip()

# The composite is a cell range: raise the window zoom so the screen copy has
# more pixels, CopyPicture, and paste into a throwaway chart object. Excel must
# be VISIBLE or the paste races the clipboard and lands blank.



def compose(chart_png, out_png, scale, gap_frac=0.035, bg=(255, 255, 255)):
    """Chart on the left, the three call-out boxes drawn down the right."""
    import sys
    sys.path.insert(0, str(HERE))
    from build_xlsx import CALLOUTS
    from PIL import Image, ImageDraw, ImageFont

    chart = Image.open(chart_png).convert("RGB")
    gap = round(chart.width * gap_frac)
    box_w = round(chart.width * 0.40)
    pad = round(11 * scale)
    font = ImageFont.truetype(FONT_FILE, round(11.5 * scale))

    out = Image.new("RGB", (chart.width + gap + box_w, chart.height), bg)
    out.paste(chart, (0, 0))
    d = ImageDraw.Draw(out)

    # Wrap to the box width first, so every box is exactly as tall as its own
    # text needs and the three stack without guessing at line counts.
    blocks = []
    for colour, text in CALLOUTS:
        words, lines, cur = text.split(), [], ""
        for w in words:
            trial = (cur + " " + w).strip()
            if d.textlength(trial, font=font) <= box_w - 2 * pad:
                cur = trial
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
        blocks.append((colour, lines))

    lh = round(font.size * 1.45)
    heights = [len(ls) * lh + 2 * pad for _, ls in blocks]
    spacing = max(round(8 * scale),
                  (chart.height - sum(heights)) // max(1, len(blocks) - 1))
    y = 0
    for (colour, lines), h in zip(blocks, heights):
        x0 = chart.width + gap
        d.rectangle([x0, y, x0 + box_w, y + h],
                    fill=tuple(int(colour[i:i+2], 16) for i in (0, 2, 4)))
        ty = y + pad
        for ln in lines:
            d.text((x0 + pad, ty), ln, font=font, fill=(255, 255, 255))
            ty += lh
        y += h + spacing
    out.save(out_png)
    return out.size


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--scale", type=int, default=3)
    a = ap.parse_args()

    OUT.mkdir(exist_ok=True)
    print("scaled %d chart part(s)" % scale_chart_xml(BOOK, PRINT_COPY, a.scale))
    lines = "\n".join(
        LINE.format(sheet=s, out=str(OUT / f).replace("\\", "\\\\"), out_name=f)
        for s, f in CHARTS)
    r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command",
                        PS.format(book=str(PRINT_COPY).replace("\\", "\\\\"),
                                  scale=a.scale, lines=lines)],
                       capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode:
        print(r.stderr.strip())
    PRINT_COPY.unlink(missing_ok=True)
    compose(OUT / "annotated_chart.png", OUT / COMPOSITE_OUT, a.scale)
    for f in [f for _, f in CHARTS] + [COMPOSITE_OUT]:
        pth = OUT / f
        print("  %-20s %s" % (f, pth.stat().st_size if pth.exists() else "MISSING"))


if __name__ == "__main__":
    main()
