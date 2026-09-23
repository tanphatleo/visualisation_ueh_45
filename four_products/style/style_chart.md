# Chart Style Guide — Finance

**Version 1.0 · 15 August 2026**
Scope: explanatory charts for finance reporting — board decks, investor material, management reporting.
Primary medium: **PowerPoint 16:9**. Primary audience: **executive / board**.

> This document is an instruction set, not a menu. Every value below is decided. Follow it literally.
> Exploratory charts — the ones you make to find the story — are **out of scope**. Keep those ugly and fast.
> The `## Decisions & rationale` table at the end records who chose what, so this can be revised later.

---

## 1. The ten rules that outrank everything else

1. **Every chart carries a takeaway title** — a sentence stating the point ("Recurring revenue overtook licence in Q3"), not a label of the contents ("Revenue by segment"). The descriptive label goes in the subtitle.
2. **Grey is the default; colour is the exception.** Colour marks the one thing the takeaway title claims. Highlight ≤10% of the ink.
3. **No 3D. No pie or donut above 5 slices — prefer a bar or slope chart always. No secondary y-axis. No rainbow scales for ordered data. No diagonal axis labels.**
4. **Bar charts start at zero.** Line charts may be truncated, but the truncation must be visible and stated in the axis title.
5. **Direct labelling beats legends.** Label each series at its final point. A legend is a fallback, never the plan.
6. **Declutter, then direct.** Strip borders, heavy gridlines, redundant labels, trailing zeros, data markers — then add back contrast, one or two annotations, and a source line.
7. **Top-left carries the message.** Left-align the title, subtitle and axis titles. Readers land there first.
8. **One form of emphasis per element** — bold *or* colour *or* size. Never all three.
9. **Never encode meaning in colour alone.** Pair it with position, a label, weight, or a shape.
10. **In variance charts, colour by favourability, not by direction.** A cost that went up is *unfavourable* and must be claret even though the bar points up. State the convention in the subtitle.

---

## 2. Colour

Base: the Financial Times `o-colors` family, reduced and re-validated for a white background.

### 2.1 Roles — these names are what charts reference

| Role token | Hex | Contrast on white | L\* | Use for |
|---|---|---|---|---|
| `--accent` | `#0F5499` Oxford | 7.64 : 1 | 35.5 | The single highlighted series. The thing the title is about. |
| `--alert` | `#990F3D` Claret | 8.43 : 1 | 32.8 | Adverse variance, breaches, the one number that is a problem. |
| `--good` | `#0D7680` Teal | 5.36 : 1 | 45.1 | Favourable variance only. |
| `--bad` | `#990F3D` Claret | 8.43 : 1 | 32.8 | Unfavourable variance only (same value as `--alert`). |
| `--ink` | `#262A33` Slate | 14.37 : 1 | 17.0 | Titles, primary body text, the darkest series. |
| `--text-2` | `#595959` | 7.00 : 1 | 37.8 | Subtitles, axis titles, axis labels. |
| `--context-line` | `#8C8C8C` | 3.36 : 1 | 58.3 | De-emphasised **lines** and small marks. |
| `--context-fill` | `#BDBDBD` | 1.88 : 1 | 76.6 | De-emphasised **bars and areas** only — never a thin line. |
| `--grid` | `#E8E8E8` | 1.23 : 1 | 92.0 | Horizontal gridlines. |
| `--wash` | `#F7F7F7` | — | 97.2 | Table banding, forecast bands, panel fills. |
| `--surface` | `#FFFFFF` | — | 100 | Plot area and slide background. |

### 2.2 Multi-series ramp

Use **in this order**. The ordering is set by lightness so the chart survives greyscale printing and deuteranopia — do not reshuffle it.

| Slot | Hex | L\* | Δ L\* from previous | Contrast on white |
|---|---|---|---|---|
| `--series-1` | `#262A33` Slate | 17.0 | — | 14.37 |
| `--series-2` | `#0F5499` Oxford | 35.5 | 18.5 ✓ | 7.64 |
| `--series-3` | `#3D9199` Teal-light | 55.6 | 20.1 ✓ | 3.68 |
| `--series-4` | `#93B2D1` Steel | 71.3 | 15.7 ✓ | 2.20 ✗ |

**`--series-4` is fill-only.** It clears the lightness test but not the 3:1 mark-contrast test, so it may fill bars, areas and stack segments — never a line, marker or thin rule. Consequently:

- **Bar / stacked / area charts: maximum 4 series.**
- **Line charts: maximum 3 series** (`--series-1` to `--series-3`).
- Above those limits the palette is not the problem. Use small multiples, group the tail into "Other", or highlight one series against grey — in that order.

### 2.3 The emphasis pattern — your default

This is what most finance charts should be:

```
every series      →  --context-fill  #BDBDBD   (bars)   /  --context-line  #8C8C8C  (lines)
the one that matters →  --accent      #0F5499
if that one is bad   →  --alert       #990F3D
```

`--accent` is the highest-chroma colour in the working set (chroma 0.54 against 0.00 for every grey), so it always wins the eye. Protect that:

- The accent **never** colours a routine category, a gridline, a title, a footer, or a background.
- **One accent per chart.** Two highlights means you have two charts.

### 2.4 Retired FT colours — do not use

Wasabi `#96CC28`, Mandarin `#FF8833` and Candy `#FF7FAA` reach only **1.92 : 1, 2.38 : 1 and 2.37 : 1** on white. They fail the 3:1 floor for marks, and Mandarin's chroma (0.80) out-shouts the accent (0.54), which breaks the whole emphasis system. They are excluded from this guide entirely.

### 2.5 Background

**Pure white `#FFFFFF`** for both the slide and the plot area. The FT's warm paper `#FFF1E0` is deliberately **not** used: it must be applied to every surface or it reads as a printing error, and it costs contrast in a bright boardroom. If it is ever adopted, every value in §2.1 stays above its threshold (contrast drops ~10%: Oxford 6.88, Claret 7.59, Teal 4.82) and the gridline becomes `#E8E3DA`.

### 2.6 Sequential and diverging scales

- **Sequential** (heatmaps, ordered magnitude): `#E6EEF6` → `#B8CFE4` → `#7FA6C9` → `#3F73A6` → `#0F5499` → `#0F3762`. Light = low, dark = high, always.
- **Diverging** (variance around zero or a target): `#990F3D` → `#C97C93` → `#F2F2F2` → `#7FB3B8` → `#0D7680`. State the midpoint explicitly in the subtitle. Keep the arms symmetric — an asymmetric scale silently biases one side.

Teal and claret sit only 12.3 L\* apart, so in greyscale they are near-identical. **Every variance chart must therefore carry signed data labels (`+2,4` / `−1,8`) or ▲▼ marks.** Direction never depends on colour alone.

---

## 3. Typography

System-safe only — zero deployment risk.

**Font stack:** `Aptos, Calibri, "Segoe UI", Helvetica, Arial, sans-serif`
Enable **lining tabular figures** wherever numbers are stacked (data labels, tables, axis labels) so digits align in columns. In PowerPoint: Format Text → Number Forms → Tabular / Lining.

Seven levels. No more.

| Level | Size | Weight | Colour | Alignment |
|---|---|---|---|---|
| Takeaway title | 20 pt | Semibold | `--ink` | Left, top-left of the chart block |
| Subtitle (descriptive + units) | 14 pt | Regular | `--text-2` | Left, directly under the title |
| Axis title | 11 pt | Regular | `--text-2` | Left (y-axis title horizontal, above the axis — never rotated) |
| Axis label | 11 pt | Regular | `--text-2` | Per axis |
| Data label | 11 pt | Regular; **Semibold** on the highlighted series | Series colour | Outside end for bars, right of last point for lines |
| Annotation | 11 pt | Regular | Colour of the thing it annotates | Left, with a 0.75 pt leader line in the same colour |
| Source / as-of line | 9 pt | Regular | `--context-line` | Left, bottom of the chart block |

**Minimum sizes:** 11 pt for anything data-bearing; 9 pt for the source line only. Never rotate axis labels — if they do not fit, switch to a horizontal bar chart or shorten the labels.

---

## 4. Layout

| Property | Value |
|---|---|
| Slide canvas | 13.33" × 7.5" (16:9) |
| Chart block | 9.0" × 4.5", positioned at 0.75" left, 1.6" top |
| Title block | Above the chart, full width, 0.75" left margin |
| Plot area fill | `#FFFFFF` — no border, no shadow, no rounded frame |
| Gridlines | **Horizontal only**, `--grid` `#E8E8E8`, 0.75 pt, behind the data. None at all on bar charts that carry direct data labels. |
| Axis lines | y-axis line removed. x-axis (baseline) kept at 0.75 pt `--text-2` on column charts; removed on horizontal bars. |
| Tick marks | Removed on both axes. |
| Bar gap | 40% gap width for single series; 0% between bars within a cluster, 60% between clusters. |
| Bar corners | Square. No rounding. |
| Line weight | 2.5 pt highlighted, 1.5 pt context. |
| Data markers | Off by default. Turn on only for the first and last point of a highlighted line, or where points are irregularly spaced. |
| Zero line | When negatives are present, draw it at 1 pt `--text-2` and label it. |
| Animation | None, except a sequential reveal of an already-finished chart. |

---

## 5. Numbers, dates and units — vi-VN

| Convention | Rule | Example |
|---|---|---|
| Thousands separator | `.` (dot) | `1.234` |
| Decimal separator | `,` (comma) | `1.234,5` |
| Currency | `₫` **after** the number, non-breaking space | `1.234,5 ₫` |
| Scale in charts | Put the unit in the **axis title**, keep axis numbers bare | Axis title `Doanh thu (tỷ ₫)`, labels `0 · 50 · 100` |
| Large-number words | `tr` = triệu (10⁶) · `tỷ` = 10⁹ · `nghìn tỷ` = 10¹² | `1,2 tỷ ₫` |
| Percentages | 0 dp on charts, 1 dp in tables | `12%` on a chart, `12,4%` in a table |
| Percentage points | Always write `pp`, never `%`, for a difference of percentages | `+1,8 pp` |
| Negatives | True minus sign U+2212, never a hyphen, never parentheses | `−1,2` |
| Signed variance | Always show the sign on variance labels | `+2,4` / `−1,8` |
| Period labels | `2024` · `Q1'24` · `Jan'24` — locale-neutral, sorted left to right oldest to newest | `Q1'24 Q2'24 Q3'24 Q4'24` |
| Year basis | **Calendar year.** Jan–Dec. No fiscal offset note needed. |
| Trailing zeros | Removed unless a column of numbers needs them to align | `12` not `12,0` |
| Rounding | Round once, at presentation. Never let rounded components fail to sum — force the largest component to absorb the residual and note it if it exceeds 0,5. |

**Currency basis must be stated** on any chart spanning periods where FX moved: append `(constant currency)` or `(as reported)` to the subtitle.

---

## 6. Chart type selection

| The question the chart answers | Use | Notes |
|---|---|---|
| How did it change over time? | Line (≥7 periods) · Column (≤6 periods) | Never a line for unordered categories. |
| What drove the change? | **Waterfall / bridge** | The finance workhorse. See §6.1. |
| How did we do against budget / target? | Bar with a target marker, or a bullet chart | Target = 1.5 pt `--ink` vertical rule, labelled. |
| Who is biggest? | Horizontal bar, sorted descending | Sort by value, not alphabetically, unless the category has a natural order. |
| What is it made of? | Stacked bar, ≤3 segments · 100% stacked bar for share | Order segments largest-at-base. Above 3 segments, split into small multiples. |
| How did two points compare across many items? | Slope chart or dumbbell | Beats a clustered bar for before/after. |
| How is it distributed? | Histogram · box plot | Show n. |
| Is A related to B? | Scatter, outliers labelled directly | No trend line without stating R² and n. |
| Many entities, same measure? | **Small multiples**, shared axis, sorted | Always preferable to 6 lines on one chart. |

### 6.1 Waterfall / bridge — fixed colour rules

```
Opening and closing bars     --ink          #262A33
Subtotal bars                --accent       #0F5499
Favourable movements         --good         #0D7680   labelled  +x,x
Unfavourable movements       --bad          #990F3D   labelled  −x,x
Connector lines              --grid         #E8E8E8   0.75 pt, dashed
```

Order the movement bars **largest to smallest by absolute size**, not by chart-of-accounts order, unless the accounting sequence is itself the message.

### 6.2 Actual vs forecast

Solid line = actual. Dashed line (same colour, 2.5 pt, 6-2 dash) = forecast. Separate them with a 0.75 pt vertical `--context-line` rule at the boundary, labelled `Forecast →` in 11 pt `--text-2`. Optionally wash the forecast region with `--wash` `#F7F7F7`.

### 6.3 Indexing and log scales

- Rebased series: state the base explicitly in the axis title — `Chỉ số (Q1'22 = 100)`.
- A log scale is permitted for multi-year compounding but must be labelled `(log scale)` in the axis title, with gridlines at each decade.
- CAGR annotations take the form `CAGR 12,4% (2020–2024)`, placed on the plot area next to the series it describes, in that series' colour.

---

## 7. Worked example

Dataset — quarterly recurring revenue, tỷ ₫: `Q1'24 82` · `Q2'24 88` · `Q3'24 91` · `Q4'24 104`.

### As a column chart

```
Doanh thu định kỳ tăng tốc trong Q4                    ← 20 pt Semibold #262A33
Doanh thu định kỳ theo quý, tỷ ₫                       ← 14 pt Regular  #595959

        82        88        91       104               ← 11 pt data labels, outside end
                                     ▓▓▓                  Q4 label Semibold #0F5499
      ▓▓▓▓      ▓▓▓▓      ▓▓▓▓      ████                  others 11 pt Regular #595959
      ▓▓▓▓      ▓▓▓▓      ▓▓▓▓      ████
      ▓▓▓▓      ▓▓▓▓      ▓▓▓▓      ████               ← Q1–Q3 #BDBDBD · Q4 #0F5499
     ─────────────────────────────────────             ← baseline 0.75 pt #595959
      Q1'24     Q2'24     Q3'24     Q4'24              ← 11 pt #595959

              +14% so với Q3                           ← 11 pt annotation #0F5499
                                                          leader line 0.75 pt #0F5499

Nguồn: Báo cáo quản trị nội bộ · Số liệu đến 31/12/2024   ← 9 pt #8C8C8C
```

Axis removed entirely, gridlines removed (direct labels make both redundant), zero baseline kept, one colour, one annotation.

### As a line chart

Same data, same title. Line 2.5 pt `#0F5499`, markers on the first and last points only, final value `104` labelled to the right of the line in 11 pt Semibold `#0F5499`, y-axis starting at 0 with horizontal gridlines at 0/25/50/75/100 in `#E8E8E8`, y-axis line and tick marks removed, axis title `tỷ ₫` sitting horizontally above the top gridline. Same source line.

**Choose the columns here.** Four periods is a comparison, not a trend — reach for the line at seven periods or more.

---

## 8. Tool blocks

### 8.1 PowerPoint / Office theme

Set these in Design → Variants → Colours → Customise Colours, then save as a theme so every deck starts compliant.

| Theme slot | Hex | Maps to |
|---|---|---|
| Text/Background — Dark 1 | `#262A33` | `--ink` |
| Text/Background — Light 1 | `#FFFFFF` | `--surface` |
| Text/Background — Dark 2 | `#595959` | `--text-2` |
| Text/Background — Light 2 | `#F7F7F7` | `--wash` |
| Accent 1 | `#0F5499` | `--accent` / `--series-2` |
| Accent 2 | `#BDBDBD` | `--context-fill` |
| Accent 3 | `#262A33` | `--series-1` |
| Accent 4 | `#3D9199` | `--series-3` |
| Accent 5 | `#93B2D1` | `--series-4` (fills only) |
| Accent 6 | `#990F3D` | `--alert` / `--bad` |
| Hyperlink | `#0F5499` | — |

Theme fonts: heading and body both **Aptos** (fallback Calibri).

### 8.2 matplotlib — `finance.mplstyle`

```
figure.figsize      : 9.0, 4.5
figure.dpi          : 200
savefig.dpi         : 200
figure.facecolor    : FFFFFF
axes.facecolor      : FFFFFF
font.family         : sans-serif
font.sans-serif     : Aptos, Calibri, Segoe UI, DejaVu Sans, Arial
font.size           : 11
axes.titlesize      : 20
axes.titleweight    : semibold
axes.titlecolor     : 262A33
axes.titlelocation  : left
axes.titlepad       : 18
axes.labelsize      : 11
axes.labelcolor     : 595959
axes.edgecolor      : 595959
axes.linewidth      : 0.75
axes.spines.top     : False
axes.spines.right   : False
axes.spines.left    : False
axes.grid           : True
axes.grid.axis      : y
axes.prop_cycle     : cycler('color', ['262A33', '0F5499', '3D9199', '93B2D1'])
grid.color          : E8E8E8
grid.linewidth      : 0.75
xtick.color         : 595959
ytick.color         : 595959
xtick.labelsize     : 11
ytick.labelsize     : 11
xtick.major.size    : 0
ytick.major.size    : 0
lines.linewidth     : 2.5
lines.markersize    : 5
legend.frameon      : False
```

vi-VN number formatting — do not rely on `locale`, it is often unset on Windows:

```python
from matplotlib.ticker import FuncFormatter

def vn(x, pos=None, dp=0):
    s = f"{abs(x):,.{dp}f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return ("−" if x < 0 else "") + s        # U+2212, not a hyphen

ax.yaxis.set_major_formatter(FuncFormatter(vn))
```

### 8.3 CSS custom properties

```css
:root {
  --surface: #FFFFFF;  --wash: #F7F7F7;  --grid: #E8E8E8;
  --context-fill: #BDBDBD;  --context-line: #8C8C8C;
  --text-2: #595959;  --ink: #262A33;
  --accent: #0F5499;  --alert: #990F3D;
  --good: #0D7680;    --bad: #990F3D;
  --series-1: #262A33; --series-2: #0F5499;
  --series-3: #3D9199; --series-4: #93B2D1; /* fills only */
  --font: Aptos, Calibri, "Segoe UI", Helvetica, Arial, sans-serif;
}
```

---

## 9. Pre-flight checklist

Run this before declaring any chart done.

- [ ] The title is a **sentence stating the point**, not a label of the contents.
- [ ] The subtitle states what is plotted **and its units**.
- [ ] Exactly **one** thing is highlighted, and it is the thing the title claims.
- [ ] Everything else is `--context-fill` or `--context-line`.
- [ ] Series count is within limits: ≤4 for bars/areas, ≤3 for lines.
- [ ] `--series-4` `#93B2D1` is used as a fill only — never as a line or marker.
- [ ] Bars start at zero. Any truncated line axis says so in the axis title.
- [ ] Series are **directly labelled**; no legend, or a documented reason for one.
- [ ] No chart junk: no border, no shadow, no 3D, no tick marks, no y-axis line, no gradient.
- [ ] Gridlines are horizontal only, `#E8E8E8`, and absent where direct labels make them redundant.
- [ ] Axis labels are horizontal. Nothing is rotated.
- [ ] Numbers follow vi-VN: `1.234,5`, `₫` after the value, `−` for negatives, signed variance labels.
- [ ] Variance colour follows **favourability**, not direction — and carries a sign or ▲▼ so it works in greyscale.
- [ ] No text below 11 pt except the 9 pt source line.
- [ ] **Source line and data as-of date are present.**
- [ ] Printed in greyscale, the message still reads.

---

## 10. Decisions & rationale

| Decision | Value | Why | Chosen by |
|---|---|---|---|
| Medium | PowerPoint 16:9 | Board and investor material | User |
| Audience | Executive / board | Drives takeaway titles, ≤4 series, one message per chart | User |
| Palette family | Financial Times `o-colors` | Reads as serious financial press; Oxford/Claret/Teal all clear WCAG AA on white | User |
| Background | Pure white `#FFFFFF` | Maximum contrast for projection and print; FT paper tint rejected because a partial application looks like an error | User |
| Accent | Oxford `#0F5499` | Highest-chroma colour available once Mandarin was retired; 7.64:1 on white | Default — see below |
| Good/bad pair | Teal `#0D7680` / Claret `#990F3D` | Serious rather than alarming; suits margin and cost commentary where red overstates | User |
| Wasabi, Mandarin, Candy | **Retired** | 1.92:1, 2.38:1, 2.37:1 on white — all fail the 3:1 mark floor; Mandarin's chroma 0.80 also out-shouts the accent | Validation |
| `--series-4` restricted to fills | `#93B2D1` | Only 2.20:1 on white. Any lighter and lightness separation fails; any darker and it collides with `--series-3` | Validation |
| Series ordering by L\* | Slate → Oxford → Teal-light → Steel | Δ L\* of 18.5 / 20.1 / 15.7, all above the 15 threshold, so the set survives greyscale and deuteranopia | Validation |
| Variance needs signs or ▲▼ | Mandatory | Teal and claret are only 12.3 L\* apart and contrast 1.57:1 with each other — indistinguishable in greyscale | Validation |
| Typography | Aptos → Calibri → Segoe UI → Helvetica → Arial | System-safe, zero deployment risk, tabular figures available | User |
| Type scale | 7 levels, 20/14/11/11/11/11/9 pt | 11 pt floor keeps data-bearing text legible when projected | Default |
| Annotation density | Takeaway title + source + 1–2 annotations | Charts must work both presented and forwarded by email | User |
| Locale | vi-VN numbers, locale-neutral period labels | `1.234,5` · `₫` suffix · `Q1'24` | User |
| Year basis | Calendar year | No fiscal offset footnote required | User |
| Footer | Source line + data as-of date | An unsourced or undated number is unusable in finance | User |
| Logo / confidentiality mark | Not required | Not selected | User |

**Defaults you may want to revisit:** the accent choice (Oxford does double duty as `--accent` and `--series-2` — fine while charts are single-message, worth splitting if multi-series charts become common), the 7-level type scale, and the 9" × 4.5" chart footprint.

---

*Add to `CLAUDE.md`:* `Chart and slide work must follow style_chart.md in the project root.`
