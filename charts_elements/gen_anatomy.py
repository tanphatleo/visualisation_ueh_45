#!/usr/bin/env python3
"""
gen_anatomy.py — Generate the two Part-06 "Chart Elements" images.

  anatomy_layout.png  — the 12 / 8 / 75 / 5 vertical layout guide
  anatomy_chart.png   — one chart with every element named

Saves PNGs to: D:/projects/datat_visualisation/day_1/charts_elements/
"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
from matplotlib.ticker import FuncFormatter
import numpy as np
import os

# ── Palette (mirrors the deck CSS variables) ─────────────────────────────────
NAVY     = '#0D1A63'
NAVY_1   = '#253A95'
NAVY_2   = '#5B73C4'
NAVY_3   = '#A8B4DE'
NAVY_4   = '#DDE2F2'
ORANGE   = '#F68048'
ORANGE_4 = '#FEEEE8'
WHITE    = '#FFFFFF'
GRAY     = '#9AA0A6'
LINE     = '#E8E8EE'

OUT = r'D:\projects\datat_visualisation\day_1\charts_elements'
os.makedirs(OUT, exist_ok=True)


# ═════════════════════════════════════════════════════════════════════════════
#  IMAGE 1 — Layout proportions: Title 12% / Subtitle 8% / Field 75% / Source 5%
# ═════════════════════════════════════════════════════════════════════════════
def img_layout():
    fig = plt.figure(figsize=(7.6, 7.2), dpi=150)
    fig.patch.set_facecolor(WHITE)
    ax = fig.add_axes([0.04, 0.04, 0.92, 0.92])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    ax.axis('off')

    blocks = [
        ('Title',       12, NAVY,   WHITE, 'States the insight — the one sentence you want remembered'),
        ('Subtitle',     8, NAVY_2, WHITE, 'Timeframe, units, scope'),
        ('Visual field', 75, NAVY_4, NAVY, 'Axes, gridlines, data series, legend, labels, callouts'),
        ('Source line',  5, ORANGE, WHITE, 'Where the data came from'),
    ]

    y = 1.0
    for name, pct, fill, txt, note in blocks:
        h = pct / 100
        y -= h
        ax.add_patch(mpatches.Rectangle((0, y), 1, h, facecolor=fill,
                                        edgecolor=WHITE, lw=2.5, zorder=2))
        # Name at the left edge, share right-aligned — never collide
        ty = y + h / 2 if pct <= 8 else y + h - 0.038
        ax.text(0.025, ty, name, fontsize=16, fontweight='800',
                color=txt, ha='left', va='center', zorder=3)
        ax.text(0.975, ty, f'{pct}%', fontsize=16, fontweight='400',
                color=txt, ha='right', va='center', zorder=3)
        if note and pct > 8:
            ax.text(0.025, ty - 0.048, note, fontsize=10.5, style='italic',
                    color=NAVY_1 if pct == 75 else '#B9C2E4',
                    ha='left', va='center', zorder=3)

    save(fig, 'anatomy_layout.png', tight=True)


# ═════════════════════════════════════════════════════════════════════════════
#  IMAGE 2 — One chart, every element named
# ═════════════════════════════════════════════════════════════════════════════
MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
          'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
D2019 = [42, 55, 48, 62, 58, 71, 75, 83, 78, 88, 95, 115]
D2020 = [38, 42, 35, 51, 60, 65, 72, 78, 85, 90, 88, 87]

W, H, DPI = 14.0, 7.8, 140          # → 1960 × 1092 px

# Plot rectangle in figure coords — deliberately narrow to leave label gutters
AX_L, AX_B, AX_W, AX_H = 0.300, 0.220, 0.460, 0.580
AX_R, AX_T = AX_L + AX_W, AX_B + AX_H       # 0.760 / 0.800


def label(fig, x, y, text, tx, ty, ha, rad=0.0, frm=None):
    """Draw a navy callout label at (x, y) with an arrow pointing to (tx, ty).

    frm overrides where the arrow starts (default: the label anchor point).
    """
    fig.add_artist(FancyArrowPatch(
        frm or (x, y), (tx, ty), transform=fig.transFigure, clip_on=False,
        arrowstyle='-|>', mutation_scale=14, lw=1.7, color=ORANGE,
        connectionstyle=f'arc3,rad={rad}', shrinkA=6, shrinkB=3, zorder=30))
    fig.text(x, y, text, transform=fig.transFigure, ha=ha, va='center',
             fontsize=12.5, fontweight='700', color=WHITE, zorder=31,
             clip_on=False,
             bbox=dict(boxstyle='round,pad=0.42', facecolor=NAVY,
                       edgecolor=NAVY, lw=1))


def img_chart():
    fig = plt.figure(figsize=(W, H), dpi=DPI)
    fig.patch.set_facecolor(WHITE)
    ax = fig.add_axes([AX_L, AX_B, AX_W, AX_H])

    x = np.arange(12)
    w = 0.36

    # ── 10 Data series ───────────────────────────────────────────────────────
    ax.bar(x - w / 2, D2019, w, color=NAVY_3, label='2019', zorder=3)
    b20 = ax.bar(x + w / 2, D2020, w, color=ORANGE, label='2020', zorder=3)

    # ── 4 Plot area ──────────────────────────────────────────────────────────
    ax.set_facecolor('#F0F3FA')

    # ── 9 Gridlines ──────────────────────────────────────────────────────────
    ax.yaxis.grid(True, color='#D5DBEC', lw=1.3, zorder=0)
    ax.set_axisbelow(True)

    # ── 6 / 13 Axes ──────────────────────────────────────────────────────────
    ax.set_xticks(x)
    ax.set_xticklabels(MONTHS, fontsize=11.5, color=NAVY_1)
    ax.set_xlim(-0.7, 11.7)          # fixed, so callout coords stay valid
    ax.set_ylim(0, 130)
    ax.set_yticks(range(0, 121, 20))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f'{v:.0f}M'))
    ax.tick_params(axis='y', labelsize=11.5, colors=NAVY_1, length=5)
    ax.tick_params(axis='x', length=5, colors=NAVY_1)
    for side in ('top', 'right'):
        ax.spines[side].set_visible(False)
    for side in ('left', 'bottom'):
        ax.spines[side].set_color(NAVY_1)
        ax.spines[side].set_linewidth(1.6)

    # ── 5 Legend (above the plot, right-aligned) ─────────────────────────────
    ax.legend(loc='lower right', bbox_to_anchor=(1.0, 1.035), ncol=2,
              frameon=False, fontsize=12, handlelength=1.5,
              columnspacing=1.6, labelcolor=NAVY_1)

    # ── 11 Data labels (selective — the two December bars only) ──────────────
    ax.text(11 - w / 2, 115 + 3.5, '115M', ha='center', va='bottom',
            fontsize=11.5, fontweight='700', color=NAVY, zorder=6)
    ax.text(11 + w / 2, 87 + 3.5, '87M', ha='center', va='bottom',
            fontsize=11.5, fontweight='700', color=ORANGE, zorder=6)

    # ── 12 Annotation / callout ──────────────────────────────────────────────
    ax.annotate('Dec 2019 peaked at 115M —\n2020 never caught up',
                xy=(11 - w / 2, 100), xytext=(4.5, 116),
                fontsize=11.5, color=NAVY, fontweight='600',
                ha='left', va='center', zorder=7,
                arrowprops=dict(arrowstyle='-|>', color=ORANGE, lw=1.7,
                                connectionstyle='arc3,rad=-0.12',
                                shrinkA=5, shrinkB=3),
                bbox=dict(boxstyle='round,pad=0.45', facecolor=ORANGE_4,
                          edgecolor=ORANGE, lw=1.4))

    # ── 1 Title · 2 Subtitle · 15 Source line ────────────────────────────────
    fig.text(AX_L, 0.952, 'December collapsed Sales',
             fontsize=19, fontweight='800', color=NAVY, ha='left', va='center')
    fig.text(AX_L, 0.898, 'Monthly sales, 2019 vs 2020 — EUR millions',
             fontsize=13.5, color=NAVY_2, ha='left', va='center')
    fig.text(AX_L, 0.048, 'Source: Kaggle — Online Retail II, 2009–2011',
             fontsize=11, color=GRAY, ha='left', va='center', style='italic')

    # ── 7 / 14 Axis titles ───────────────────────────────────────────────────
    fig.text(0.246, AX_B + AX_H / 2, 'EUR (M)', fontsize=13, fontweight='700',
             color=NAVY, ha='center', va='center', rotation=90)
    fig.text(AX_L + AX_W / 2, 0.122, 'Month', fontsize=13, fontweight='700',
             color=NAVY, ha='center', va='center')

    # ── 3 Chart area border ──────────────────────────────────────────────────
    fig.add_artist(FancyBboxPatch(
        (0.014, 0.014), 0.972, 0.972, boxstyle='square,pad=0',
        transform=fig.transFigure, clip_on=False, zorder=1,
        lw=1.6, edgecolor=NAVY_3, facecolor='none', linestyle=(0, (6, 4))))

    # ── Callout labels ───────────────────────────────────────────────────────
    LG, RG = 0.228, 0.782        # left / right gutter anchor x

    # Left gutter — arrows point right, threading between the y-axis furniture
    label(fig, LG, 0.952, '1  Chart title',   AX_L,  0.952, 'right')
    label(fig, LG, 0.898, '2  Subtitle',      AX_L,  0.898, 'right')
    label(fig, LG, 0.788, '4  Plot area',     0.318, 0.782, 'right')
    label(fig, LG, 0.690, '9  Gridlines',     0.400, 0.666, 'right', rad=-0.16)
    label(fig, LG, 0.590, '6  Y-axis',        AX_L,  0.620, 'right')
    label(fig, LG, 0.430, '7  Y-axis title',  0.254, 0.462, 'right')
    label(fig, LG, 0.320, '8  Axis units',    0.272, 0.309, 'right')
    label(fig, LG, 0.048, '15  Source line',  AX_L,  0.048, 'right')

    # Right gutter — arrows point left, all kept clear of the bars
    label(fig, 0.986, 0.912, '3  Chart area', 0.920, 0.980, 'right',
          frm=(0.920, 0.933))
    label(fig, RG, 0.845, '5  Legend',        0.762, 0.834, 'left')
    label(fig, RG, 0.780, '12  Annotation',   0.648, 0.748, 'left')
    label(fig, RG, 0.640, '11  Data labels / callouts', 0.752, 0.624, 'left')
    label(fig, RG, 0.460, '10  Data series',  0.733, 0.420, 'left')
    label(fig, RG, 0.205, '13  X-axis',       0.758, 0.212, 'left')
    label(fig, RG, 0.122, '14  X-axis title', 0.585, 0.122, 'left')

    save(fig, 'anatomy_chart.png', tight=False)


def save(fig, fname, tight=True):
    path = os.path.join(OUT, fname)
    kw = dict(dpi=fig.dpi, facecolor=WHITE)
    if tight:
        kw.update(bbox_inches='tight', pad_inches=0.12)
    fig.savefig(path, **kw)
    plt.close(fig)
    print(f'  saved: {fname}')


if __name__ == '__main__':
    print('Generating Part-06 anatomy images...')
    img_layout()
    img_chart()
    print('Done.')
