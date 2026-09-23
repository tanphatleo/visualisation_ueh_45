"""
Rebuild a line-art PNG at higher resolution with clean, flat colours.

This is an automated re-trace of an image you supply — no redrawing. It works
because line art on white is a very constrained kind of picture: every pixel is
either white, or one of a handful of stroke colours laid down at some coverage.

The pipeline:

  1. coverage      how much ink is in each pixel, from its luminance. Edge
                   pixels are partly covered, which is what makes them look soft.
  2. unmix         a half-covered orange pixel reads as pale orange. Undo the
                   blend with white — c_true = white + (c - white)/coverage —
                   so edge pixels report their real colour instead of a tint.
                   Skip this and clustering invents colours that are not in the
                   drawing: pale blends like #7D9FA4 and #BEAF67 show up as if
                   they were strokes.
  3. cluster       k-means over the unmixed colours gives the true palette.
  4. rebuild       each colour's coverage map is resampled up with a cubic
                   filter, its edge slightly re-tightened, then composited in
                   its flat palette colour.

The result is larger, has exactly N flat colours, and has smooth edges — the
softness of the original is resampled rather than magnified.

Run:  python retrace.py                    # 3x, 12 colours
      python retrace.py --scale 4 -k 14
"""

import argparse
from pathlib import Path

import numpy as np
from PIL import Image
from scipy.cluster.vq import kmeans2
from scipy.ndimage import gaussian_filter, label as ndimage_label, zoom

HERE = Path(__file__).resolve().parent

# Pixels at least this covered are trusted to vote on the palette. Lower and the
# blends pollute the clusters; higher and thin strokes get no vote at all.
CORE_COVERAGE = 0.55
WHITE_LUM = 250.0     # luminance treated as bare paper
INK_LUM = 70.0        # luminance treated as fully covered


def coverage(rgb):
    """How much ink sits in each pixel, 0 (paper) to 1 (solid stroke)."""
    lum = rgb.mean(2)
    return np.clip((WHITE_LUM - lum) / (WHITE_LUM - INK_LUM), 0.0, 1.0)


def unmix(rgb, cov, eps=1e-3):
    """Recover each pixel's stroke colour by undoing its blend with white."""
    c = cov[..., None]
    out = 255.0 + (rgb - 255.0) / np.maximum(c, eps)
    return np.clip(out, 0, 255)


def palette(rgb, cov, k, seed=0):
    """The drawing's true colours, voted on by confidently-covered pixels only."""
    core = cov >= CORE_COVERAGE
    cent, _ = kmeans2(unmix(rgb, cov)[core], k, minit="++", seed=seed, iter=40)
    return cent


def clean_labels(lab, cov, k, sigma=1.8, rounds=2, min_blob=30):
    """
    Re-decide every inked pixel's colour from the colours around it.

    Nearest-colour matching judges each pixel alone, so along a stroke whose hue
    sits between two clusters the label flickers from pixel to pixel and the
    stroke comes out mottled. This is the fix, in two parts:

      * a coverage-weighted vote. Each colour blurs its own footprint into a
        score map and the winner takes the pixel, so a pixel surrounded by
        crimson becomes crimson even if, on its own, it measured slightly
        magenta. Weighting by coverage means solid stroke centres decide the
        vote and faint edge pixels mostly listen.
      * speck removal. Any island of a colour smaller than min_blob pixels is
        a mislabel — real strokes are long — so it is erased and re-voted from
        whatever surrounds it.
    """
    ink = cov > 0.02
    for _ in range(rounds):
        votes = np.stack(
            [gaussian_filter(np.where(lab == i, cov, 0.0), sigma) for i in range(k)],
            axis=-1)
        lab = np.where(ink, np.argmax(votes, axis=-1), -1)

    for i in range(k):
        comp, n = ndimage_label(lab == i)
        if not n:
            continue
        sizes = np.bincount(comp.ravel())
        sizes[0] = 0                       # background is not a component
        specks = np.isin(comp, np.flatnonzero((sizes > 0) & (sizes < min_blob)))
        lab[specks] = -1

    holes = ink & (lab < 0)
    if holes.any():
        votes = np.stack(
            [gaussian_filter(np.where(lab == i, cov, 0.0), sigma) for i in range(k)],
            axis=-1)
        lab = np.where(holes, np.argmax(votes, axis=-1), lab)
    return lab


def retrace(path, scale=3, k=12, pad=0, badge=None, sharpen=2.6,
            transparent=True):
    src = np.asarray(Image.open(path).convert("RGB")).astype(float)
    if badge:
        x0, y0, x1, y1 = badge
        src[y0:y1, x0:x1] = 255.0

    cov = coverage(src)
    cent = palette(src, cov, k)

    # Assign EVERY inked pixel — including soft edges — to the nearest true
    # colour, using the unmixed value so an edge pixel is judged on hue, not on
    # how pale it happens to be.
    #
    # Unmixing amplifies noise exactly where coverage is lowest, so a bare
    # nearest-colour test makes edge pixels flicker between two similar clusters
    # and leaves speckled outlines on the dots. Smoothing the colour field first
    # — weighted by coverage, so pale pixels borrow hue from their solid
    # neighbours rather than voting on it — removes the speckle without eroding
    # the strokes the way a median filter would.
    true = unmix(src, cov)
    num = gaussian_filter(true * cov[..., None], sigma=(1.0, 1.0, 0))
    den = gaussian_filter(cov, sigma=1.0)[..., None]
    true = np.where(den > 1e-3, num / np.maximum(den, 1e-3), true)
    ink = cov > 0.02
    d = np.linalg.norm(true[ink][:, None, :] - cent[None, :, :], axis=2)
    lab = np.full(cov.shape, -1, dtype=int)
    lab[ink] = np.argmin(d, axis=1)
    lab = clean_labels(lab, cov, len(cent))

    # Re-measure coverage AGAINST EACH PIXEL'S OWN COLOUR.
    #
    # The luminance ramp above treats every stroke as equally dark, which is
    # false: amber #BB7A00 sits at luminance ~125, so it tops out around 0.7
    # coverage and its interior never reaches solid. The stroke then renders
    # pale and mottled, tracking the source's noise, while dark purple strokes
    # come out clean — exactly the difference visible in the output.
    #
    # A pixel of colour c laid on white at coverage a is  px = white + a(c - white).
    # Solving for a by projection gives the true coverage for that colour, so
    # every stroke reaches 1.0 in its interior whatever its lightness.
    d_white = 255.0 - src
    cov_fixed = np.zeros_like(cov)
    for i, colour in enumerate(cent):
        vec = 255.0 - colour
        denom = float(vec @ vec)
        if denom < 1e-6:
            continue
        m = lab == i
        if m.any():
            cov_fixed[m] = np.clip((d_white[m] @ vec) / denom, 0.0, 1.0)
    cov = cov_fixed

    h, w = int(cov.shape[0] * scale), int(cov.shape[1] * scale)
    # Composite each colour "over" what is already there, tracking colour and
    # alpha separately. Building the alpha channel here rather than keying it
    # out of a white render afterwards is what keeps the edges clean: a keyed
    # edge pixel carries the background it was flattened against, which shows up
    # as a pale halo on any other background.
    acc_rgb = np.zeros((h, w, 3))
    acc_a = np.zeros((h, w))
    for i, colour in enumerate(cent):
        layer = np.where(lab == i, cov, 0.0)
        if layer.max() <= 0:
            continue
        big = np.clip(zoom(layer, scale, order=3), 0, 1)
        # Re-tighten the edge. Cubic resampling spreads the ramp over several
        # times as many pixels, which reads as blur; this pulls it back without
        # hard-thresholding it into jaggies.
        big = np.clip((big - 0.5) * sharpen + 0.5, 0, 1)
        a = big[..., None]
        acc_rgb = acc_rgb * (1 - a) + (colour / 255.0) * a
        acc_a = acc_a * (1 - big) + big

    if transparent:
        # Un-premultiply so the stored colour is the stroke's own, not a blend
        # with black. Viewers multiply by alpha again on display.
        safe = np.maximum(acc_a, 1e-4)[..., None]
        rgb = np.clip(acc_rgb / safe, 0, 1)
        rgba = np.dstack([rgb, acc_a])
        img = Image.fromarray((rgba * 255).astype(np.uint8), mode="RGBA")
        mask = acc_a > 0.02
    else:
        a = acc_a[..., None]
        flat = np.ones((h, w, 3)) * (1 - a) + acc_rgb
        img = Image.fromarray((np.clip(flat, 0, 1) * 255).astype(np.uint8))
        mask = acc_a > 0.02

    # Trim hard to the artwork. pad=0 means the ink touches all four edges.
    ys, xs = np.nonzero(mask)
    box = (max(0, xs.min() - pad), max(0, ys.min() - pad),
           min(img.width, xs.max() + 1 + pad), min(img.height, ys.max() + 1 + pad))
    return img.crop(box), cent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?", default="image_t091.png")
    ap.add_argument("-o", "--out", default="brain_traced.png")
    ap.add_argument("-k", type=int, default=12, help="number of stroke colours")
    ap.add_argument("--scale", type=float, default=3.0)
    ap.add_argument("--sharpen", type=float, default=2.6,
                    help="edge contrast; 1 = as resampled, higher = crisper")
    ap.add_argument("--pad", type=int, default=0, help="margin in px")
    ap.add_argument("--bg", choices=["none", "white"], default="none",
                    help="'none' writes RGBA with a transparent background")
    ap.add_argument("--badge", nargs=4, type=int, metavar=("X0", "Y0", "X1", "Y1"),
                    default=[35, 30, 70, 75], help="corner region to blank out")
    a = ap.parse_args()

    img, cent = retrace(HERE / a.src, scale=a.scale, k=a.k, badge=a.badge,
                        sharpen=a.sharpen, pad=a.pad,
                        transparent=(a.bg == "none"))
    img.save(HERE / a.out)
    print(f"wrote {a.out}  {img.width}x{img.height}  {img.mode}  ({a.k} colours)")
    for c in sorted(cent, key=lambda c: c.sum()):
        print(f"   #{int(c[0]):02X}{int(c[1]):02X}{int(c[2]):02X}")


if __name__ == "__main__":
    main()
