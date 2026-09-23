"""
Recover Playfair's numbers from Playfair's engraving.

    python trace_from_image.py            # trace, verify, print
    python trace_from_image.py --write    # also rewrite data/playfair_denmark_norway.csv

Playfair published the plate, not the table, so the series has to be read back off the
picture. The method is the same one that works on any scanned or screenshotted chart:

  1  Find the plate's own gridlines by looking for rows and columns that are mostly ink.
  2  Fit value = m*y + c through the horizontal rules, whose values are known (190 down
     to 10, ten thousand pounds apart - the plate says so in its own caption).
  3  Map years to x PIECEWISE between the decade rules. They are 213-227 px apart, not
     evenly spaced: a copper plate was engraved by hand and a single linear fit is
     visibly wrong by the 1750s.
  4  Isolate each curve by ink colour, take the largest contiguous run in each column
     (the stroke, not stray speckle from the wash), and sample at every year.
  5  Draw the sampled points back onto the original and LOOK at it. Step 5 is the whole
     check - the arithmetic in 1-4 cannot tell you it locked onto the wrong line.

Accuracy: the plate is 5.75 px per thousand pounds, so a stroke centre is good to about
+/- 0.3. Round-tripping through the ink and the antialiasing, treat the output as +/- 1.

Requires pillow and numpy. Input: reference/playfair_original.png (2067 x 1527), the
scan linked from historyofinformation.com entry 2929.
"""
import csv
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw

HERE = Path(__file__).resolve().parent
SRC = HERE / "reference" / "playfair_original.png"
CHECK = HERE / "reference" / "trace_verification.png"
CSV_OUT = HERE / "data" / "playfair_denmark_norway.csv"

a = np.asarray(Image.open(SRC).convert("RGB")).astype(int)
R, G, B = a[:, :, 0], a[:, :, 1], a[:, :, 2]

# --- y calibration: the 19 labelled rules, 190 at the top down to 10 -----------
# Found by scanning for rows whose ink fraction exceeds 0.55 across the plot, then
# taking each run's intensity-weighted centre. Hard-coded here so the trace is
# reproducible without re-running the detector.
GRID_Y = [211.0, 268.0, 325.5, 380.0, 439.5, 498.5, 555.0, 611.5, 668.5, 725.0,
          783.4, 839.0, 896.5, 954.9, 1013.0, 1074.0, 1131.0, 1189.0, 1245.5]
GRID_V = list(range(190, 0, -10))
M, C = np.polyfit(GRID_Y, GRID_V, 1)                 # value = M*y + C


def value_at(y):
    return M * y + C


# --- x calibration: the decade rules, piecewise ------------------------------
GRID_X = [87.0, 304.4, 520.5, 738.0, 964.5, 1178.0, 1405.0, 1626.2, 1848.5]
GRID_YR = list(range(1700, 1790, 10))


def x_of(year):
    return float(np.interp(year, GRID_YR, GRID_X))


# --- the two inks ------------------------------------------------------------
# Ochre imports #D48D50 and crimson exports #B05058 both have R-G > 45. What
# separates them is G vs B: the ochre is warm (G well above B), the crimson is
# cool (G at or below B). Nothing else on the plate is that saturated.
SAT = (R - G > 45) & (R > 110)
MASK_IMPORTS = SAT & (G - B > 14)
MASK_EXPORTS = SAT & (G - B < 9)

Y0, Y1 = 150, 1305                                   # inside the frame
X0, X1 = 88, 1849


def trace(mask):
    """Column -> stroke centre in pixels."""
    out = {}
    for x in range(X0, X1):
        ys = np.nonzero(mask[Y0:Y1, x])[0]
        if len(ys) == 0:
            continue
        ys = ys + Y0
        runs, start = [], 0
        for i in range(1, len(ys) + 1):
            if i == len(ys) or ys[i] - ys[i - 1] > 4:
                runs.append(ys[start:i])
                start = i
        run = max(runs, key=len)
        if len(run) < 3:                              # speckle from the wash
            continue
        out[x] = float(run.mean())
    return out


def sample(stroke, year, half=6):
    """Median of the stroke centre in a +/- 6 px window, converted to a value."""
    x = int(round(x_of(year)))
    vs = [stroke[k] for k in range(x - half, x + half + 1) if k in stroke]
    return None if not vs else value_at(float(np.median(vs)))


def main():
    resid = max(abs(GRID_V[i] - value_at(GRID_Y[i])) for i in range(len(GRID_Y)))
    print(f"y fit: value = {M:.6f}*y + {C:.3f}   max residual {resid:.2f} units")
    print(f"       y at 0 = {(0 - C) / M:.1f}   y at 200 = {(200 - C) / M:.1f}")

    ti, te = trace(MASK_IMPORTS), trace(MASK_EXPORTS)
    span = X1 - X0
    print(f"traced columns: imports {len(ti)}/{span}  exports {len(te)}/{span}")

    rows = [(yr, sample(ti, yr), sample(te, yr)) for yr in range(1700, 1781)]
    missing = [yr for yr, i, e in rows if i is None or e is None]
    if missing:
        print(f"WARNING: no stroke found for {missing}")

    # --- step 5: put the points back on the plate and look at the result -----
    im = Image.open(SRC).convert("RGB")
    d = ImageDraw.Draw(im)
    for yr, imp, exp in rows:
        x = x_of(yr)
        for v, colour in ((imp, (0, 110, 255)), (exp, (0, 170, 0))):
            if v is None:
                continue
            y = (v - C) / M
            d.ellipse([x - 5, y - 5, x + 5, y + 5], outline=colour, width=3)
    im.resize((1240, 916), Image.LANCZOS).save(CHECK)
    print(f"wrote {CHECK.relative_to(HERE)} - open it and check every marker sits on a line")

    for yr, imp, exp in rows:
        print(f"{yr}  imports {imp:6.1f}   exports {exp:6.1f}   balance {exp - imp:+7.1f}")

    if "--write" in sys.argv:
        with CSV_OUT.open("w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["year", "imports", "exports", "balance"])
            for yr, imp, exp in rows:
                w.writerow([yr, f"{imp:.1f}", f"{exp:.1f}", f"{exp - imp:.1f}"])
        print(f"wrote {CSV_OUT.relative_to(HERE)}")
    else:
        print("(--write to rewrite the CSV)")


if __name__ == "__main__":
    main()
