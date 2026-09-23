"""
The two scanning patterns a reader brings to a page: Z and F.

Each graphic is a stand-in page — a few grey blocks, no real words — with the
scan path drawn over it in the accent colour and its stops numbered. The page
content is deliberately unreadable: the point is the route the eye takes, and
legible text would pull attention into reading instead of looking.

The two pages differ on purpose, because the pattern follows the layout:

  Z   sparse page, a handful of elements. The eye sweeps the top, cuts back
      across the middle, sweeps the bottom. Landing pages, ads, posters — and
      slides.
  F   dense page, stacked text. The eye sweeps the first line, sweeps a shorter
      second, then runs down the left margin sampling openings. Articles,
      search results, documentation.

Both end where a designer should put the thing that matters: bottom-right for
Z, down the left edge for F.

Run:  python build_patterns.py              # images/z_pattern.png, f_pattern.png
      python build_patterns.py --bg white
"""

import argparse
import math
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, FancyArrowPatch, FancyBboxPatch, Rectangle

HERE = Path(__file__).resolve().parent

INK = "#262A33"           # captions
TEXT_2 = "#595959"        # the sub-caption
CONTEXT_LINE = "#8C8C8C"  # page border
ACCENT = "#0F5499"        # the scan path — the only saturated colour here
BLOCK = "#E4E6E9"         # placeholder content
SURFACE = "#FFFFFF"

# Page geometry, shared by both panels so the two graphics sit at the same
# scale when they are placed side by side on a slide.
PAGE_X0, PAGE_Y0, PAGE_W, PAGE_H = 0.0, 14.0, 100.0, 72.0
PAGE_X1, PAGE_Y1 = PAGE_X0 + PAGE_W, PAGE_Y0 + PAGE_H

PATH_LW = 4.6
HEAD = 22
BADGE_R = 3.6


def page(ax):
    """The blank device: white, hairline border, gently rounded."""
    ax.add_patch(FancyBboxPatch(
        (PAGE_X0, PAGE_Y0), PAGE_W, PAGE_H,
        boxstyle="round,pad=0,rounding_size=2.5",
        facecolor=SURFACE, edgecolor=CONTEXT_LINE, linewidth=1.0))


def block(ax, x, y, w, h, r=1.2):
    """One placeholder element — a heading, an image, a button, a line of text."""
    ax.add_patch(FancyBboxPatch((x, y), w, h,
                                boxstyle=f"round,pad=0,rounding_size={r}",
                                facecolor=BLOCK, edgecolor="none"))


def arrows(ax, pts, colour=ACCENT):
    """
    The scan path: one polyline, with an arrowhead at the MIDDLE of each leg.

    The turns are the whole point of both patterns, so every leg needs a head
    saying which way the eye is travelling. Putting those heads at the ends of
    the legs does not work here — the numbered stop sits on that same corner and
    buries the head, and the legs then read as plain lines. Mid-leg keeps both
    visible and leaves the corners clean for the numbers.
    """
    ax.plot([p[0] for p in pts], [p[1] for p in pts], color=colour, lw=PATH_LW,
            solid_capstyle="round", solid_joinstyle="round", zorder=4)
    for a, b in zip(pts[:-1], pts[1:]):
        dx, dy = b[0] - a[0], b[1] - a[1]
        n = math.hypot(dx, dy)
        if n < 1e-6:
            continue
        mx, my = (a[0] + b[0]) / 2, (a[1] + b[1]) / 2
        ux, uy = dx / n * 0.7, dy / n * 0.7
        ax.add_patch(FancyArrowPatch((mx - ux, my - uy), (mx + ux, my + uy),
                                     arrowstyle="-|>", mutation_scale=HEAD,
                                     linewidth=PATH_LW, color=colour,
                                     shrinkA=0, shrinkB=0, zorder=5))


def badge(ax, x, y, n, colour=ACCENT):
    """A numbered stop, so the order is stated and not just implied."""
    ax.add_patch(Circle((x, y), BADGE_R, facecolor=colour, edgecolor=SURFACE,
                        linewidth=1.4, zorder=6))
    ax.text(x, y - 0.15, str(n), ha="center", va="center", color=SURFACE,
            fontsize=9.5, fontweight="bold", zorder=7)


# --------------------------------------------------------------------------
# Z — a sparse page. Few elements, spread to the corners, which is exactly the
# layout that produces the Z sweep.
# --------------------------------------------------------------------------
def z_page(ax):
    page(ax)
    block(ax, 7, 74, 15, 7)                        # logo, top left
    for i in range(3):                             # nav, top right
        block(ax, 60 + i * 12, 76, 9, 3.5)
    block(ax, 7, 37, 86, 30)                       # hero
    block(ax, 7, 24, 34, 3.5)                      # strapline
    block(ax, 7, 18.5, 26, 3.5)
    block(ax, 71, 19, 22, 9, r=2)                  # the call to action

    stops = [(11, 77.5), (89, 77.5), (12, 23), (89, 23)]
    arrows(ax, stops)
    for i, (x, y) in enumerate(stops, 1):
        badge(ax, x, y, i)


# --------------------------------------------------------------------------
# F — a dense page. Stacked text lines, shortening down the page: the further
# down, the less of each line is actually read.
# --------------------------------------------------------------------------
def f_page(ax):
    page(ax)
    block(ax, 7, 76, 52, 6)                        # headline
    # Nine body lines, tapering: the further down the page, the less of each
    # line gets read. Step and top are chosen so the last line still clears the
    # page floor by about the same margin the headline leaves at the top.
    top, step, h = 67.0, 5.9, 3.6
    widths = [86, 82, 86, 78, 84, 80, 74, 66, 58]
    for i, w in enumerate(widths):
        block(ax, 7, top - i * step, w, h)

    stem_x, top_y = 11.0, 79.0            # 79 = centre of the headline block
    mid_y = top - step + h / 2            # centre of the second body line
    bot_y = 26.0
    arrows(ax, [(stem_x, top_y), (89, top_y)])     # full first sweep
    arrows(ax, [(stem_x, mid_y), (68, mid_y)])     # shorter second sweep
    arrows(ax, [(stem_x, top_y), (stem_x, bot_y)])  # down the left margin

    # Numbered at the END of each stroke, not the start: all three strokes
    # begin at the same corner, so start badges would land on top of each other.
    badge(ax, 89, top_y, 1)
    badge(ax, 68, mid_y, 2)
    badge(ax, stem_x, bot_y, 3)


# The two patterns do NOT rest on equal evidence, and the source lines say so.
# F came out of eyetracking 232 people; Z is received craft wisdom with no
# originating study, descended from Arnold's one-diagonal Gutenberg diagram.
# Both sources are kept to two lines so the panels crop to the same height.
PANELS = [
    ("z_pattern.png", z_page, "Z-PATTERN",
     "Sparse pages — landing pages, ads, slides",
     "Design convention; no originating study. After the\n"
     "Gutenberg diagram, E. C. Arnold (1950s)"),
    ("f_pattern.png", f_page, "F-PATTERN",
     "Dense pages — articles, search results, reports",
     "Nielsen, F-Shaped Pattern for Reading Web Content,\n"
     "Nielsen Norman Group (2006); revisited Pernice (2017)"),
]


def build(draw, title, sub, source, out_path, bg, show_source):
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Aptos", "Calibri", "Segoe UI",
                                       "DejaVu Sans", "Arial"]

    xlim, ylim = (-2.0, 102.0), (-2.0, 100.0)
    width = 5.4
    fig, ax = plt.subplots(
        figsize=(width, width * (ylim[1] - ylim[0]) / (xlim[1] - xlim[0])), dpi=300)
    ax.set_xlim(*xlim)
    ax.set_ylim(*ylim)
    ax.set_aspect("equal")
    ax.axis("off")

    draw(ax)
    ax.text(PAGE_X0, 8.5, title, ha="left", va="center",
            fontsize=15, color=INK, fontweight="bold")
    ax.text(PAGE_X0, 3.0, sub, ha="left", va="center",
            fontsize=10.5, color=TEXT_2)
    # Off by default: the deck slide carries the citation for both panels, and
    # printing it here as well states it twice on the same screen. Kept behind a
    # flag so these stay self-contained if they are ever used away from the deck.
    if show_source:
        # Anchored by its top so the block grows down; the tight bbox picks it up.
        ax.text(PAGE_X0, -1.0, source, ha="left", va="top", linespacing=1.5,
                fontsize=8.5, color=CONTEXT_LINE)

    if bg != "none":
        fig.patch.set_facecolor(bg)
    fig.subplots_adjust(0, 0, 1, 1)
    fig.savefig(out_path, dpi=300, transparent=(bg == "none"),
                bbox_inches="tight", pad_inches=0.04)
    plt.close(fig)
    print(f"wrote {out_path.name}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bg", default="none", help="'none' for transparent, or a colour")
    ap.add_argument("--outdir", default="images")
    ap.add_argument("--source", choices=["on", "off"], default="off",
                    help="draw the credit inside each panel")
    a = ap.parse_args()

    outdir = HERE / a.outdir
    outdir.mkdir(exist_ok=True)
    for name, draw, title, sub, source in PANELS:
        build(draw, title, sub, source, outdir / name, a.bg,
              show_source=(a.source == "on"))


if __name__ == "__main__":
    main()
