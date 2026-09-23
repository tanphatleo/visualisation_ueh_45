# style_chart.md — "Warehouse Stand-Up"

**Version 1.0 · 2026-08-17 · scope: `mini_cases/inventory_turnover/` and any chart that asks for
this style by name.**

A light-hearted, funny house style for explanatory charts. Warm cream paper, one loud tangerine,
chunky rounded lines, and titles written like punchlines. It is deliberately **not** the project's
FT-ish serious style — ask for that one by its own name.

Read this as instructions, not suggestions. Every value here is resolved. Nothing in this file is
"it depends".

> **The one rule the rest of the file serves:** the joke lives in the *words*. The encoding stays
> boringly honest. A reader who does not find it funny must still be able to read the number.

---

## 1. Register — how "funny" is allowed to get

| Do | Don't |
| --- | --- |
| Takeaway title is a **punchline**: a sentence with a turn in it, ≤ 14 words. | A pun so dense the point disappears. |
| Subtitle plays the **straight man**: flat, factual, carries the units. | Two jokes in a row. Setup, then punchline. |
| Series get **nicknames** — `US`, `THE JONESES`, `THE PLAN` — instead of column names. | Nicknames that hide who the series is. Say the real thing in the subtitle. |
| **One** wisecrack annotation per chart. Every other annotation is informative. | A chart where each label is doing bits. |
| **One** emoji per chart, maximum, in the title or one annotation. Zero is also correct. | Emoji in axis titles, tick labels, legends, or table headers. Ever. |
| Sign the source line off with a dry half-sentence after the real source. | Faking or fudging the source for the gag. |
| Round caps, thick lines, friendly markers. | Cartoon fonts, gradients, drop shadows, 3D, clip-art bars. |

**No Comic Sans.** The chart is the comedian; the font is the straight man. Headlines are chunky
Verdana Bold, everything else is Segoe UI, and that is the whole typographic joke.

Non-negotiables that survive the comedy, unchanged: takeaway title on every chart, grey by default
with colour as the exception, bars start at zero, direct labelling over legends, declutter before
you decorate, no 3D, no secondary y-axis (one licensed exception in §7), no diagonal tick labels,
never encode meaning in colour alone.

---

## 2. Colour

Warm cream paper, warm greys, one high-chroma accent. Contrast ratios below are measured against
`--surface` `#FFF9F0`.

### 2.1 Roles

| Role | Hex | Contrast vs surface | Use for |
| --- | --- | --- | --- |
| `--surface` | `#FFF9F0` | — | chart area, plot area, slide background |
| `--wash` | `#FBEFDD` | 1.08 | table header band, callout box fills |
| `--grid` | `#E7D8C3` | 1.34 | horizontal gridlines, axis lines |
| `--context` | `#6E6259` | 5.64 | benchmark / de-emphasised series, dashed context lines |
| `--text-2` | `#5C5149` | 7.35 | tick labels, axis titles, source line |
| `--ink` | `#2F2A26` | 13.55 | takeaway title, subtitle, table body |
| `--accent` | `#EF5B23` | 3.24 | **the hero series — lines, bars, markers only** |
| `--accent-text` | `#C2400F` | 4.97 | the accent used as *text*: end labels, annotations |
| `--good` | `#0B7A72` | 4.96 | "this went the right way" |
| `--bad` | `#B8352C` | 5.60 | "this went the wrong way", event lines |

`--accent` at 3.24:1 clears the 3:1 mark threshold but **not** the 4.5:1 text threshold. That is why
there are two of them: paint with `--accent`, write with `--accent-text`. Never write body text in
`#EF5B23`.

### 2.2 Categorical set — use in this order, stop at four

| # | Hex | Text-safe variant | Name |
| --- | --- | --- | --- |
| 1 | `#EF5B23` | `#C2400F` | Tangerine (always the hero) |
| 2 | `#0E8F86` | `#0B7A72` | Teal |
| 3 | `#E0A22B` | `#8C6B12` | Mustard |
| 4 | `#52397F` | `#52397F` | Plum |

Ordered so neighbours differ in lightness, not only hue (≈ 54 → 31 → 52 → 35 on a 0–100 lightness
scale): the set survives greyscale printing and deuteranopia. **Five or more series is not a colour
problem, it is a chart-type problem** — switch to small multiples or a slope chart.

Green–red pairs are banned. `--good`/`--bad` are teal and claret, and they always carry a word or a
direction as well as the colour.

---

## 3. Type

| Element | Font | Size | Weight | Colour | Align |
| --- | --- | --- | --- | --- | --- |
| Takeaway title (the punchline) | Verdana | 16 pt | Bold | `--ink` | left |
| Subtitle (the straight man) | Segoe UI | 11 pt | Regular | `--text-2` | left |
| Chart title inside the chart object | Verdana | 13 pt | Bold | `--ink` | left |
| Axis title | Segoe UI | 8 pt | Regular, UPPERCASE | `--text-2` | left / rotated |
| Tick label | Segoe UI | 9 pt | Regular | `--text-2` | — |
| Series end label (direct label) | Segoe UI | 9 pt | Bold, UPPERCASE | series text-safe hex | left of nothing — it hugs the line |
| Annotation / wisecrack | Segoe UI | 9 pt | Italic | `--accent-text` or `--bad` | left |
| Source line | Segoe UI | 9 pt | Regular | `--text-2` | left |

Fallbacks, in order: Verdana → Tahoma → DejaVu Sans; Segoe UI → Calibri → DejaVu Sans. Both primaries
ship with Windows and are visible to Excel and matplotlib on this machine.

One form of emphasis per element: bold **or** colour **or** size. The end labels are bold *and*
coloured because they replace the legend — that is the single sanctioned exception, and it applies
to nothing else.

---

## 4. Layout

Four bands, top to bottom, on a 4:3 or 16:9 canvas:

| Band | Height | Contents |
| --- | --- | --- |
| Title | 14 % | the punchline, left-aligned, hard against the left margin |
| Subtitle | 8 % | one factual sentence; **units live here** |
| Visual | 72 % | the plot, plus direct labels and at most one annotation |
| Source | 6 % | `Source: … | ` + the dry sign-off |

Excel charts have no subtitle and no source line, so in `.xlsx` those bands are **worksheet cells**.
The punchline may live on the chart object instead — do that when the chart will be copied into a
deck, and then **never print the punchline twice**: cells carry the subtitle and the source, the
chart carries the title. In matplotlib all four bands are `fig.text` calls at fixed y fractions.

Chart canvas: **14.6 × 10.9 cm** in Excel (4:3), **8.0 × 6.0 in @ 150 dpi** in matplotlib. Leave
**≥ 24 % of the width free on the right** for end labels — Excel clips a data label that runs past
the chart edge, silently.

---

## 5. Marks

| Mark | Spec |
| --- | --- |
| Line — hero | 3 pt, `--accent`, round cap and round join, no shadow, `smooth = False` |
| Line — context / benchmark | 1.5 pt, `--context`, dashed (`prstDash="dash"` / `ls=(0,(5,3))`) |
| Marker | circle, 7 px, fill `--surface`, 2 pt ring in the series colour — the reader can count the periods |
| Bar / column | solid fill, no outline, gap width 40 %, `invertIfNegative = False` |
| Grey-out | any series that is not the point of the title gets `--context`; ≤ 10 % of the ink is accent |
| Area | series colour at 25 % alpha, 1.5 pt line at full strength on top |

Rounded joins and hollow markers are the whole "light-hearted" look. Do not substitute them with
picture fills unless the picture **is** the unit (a waffle of boxes for boxes of stock is fine; a
truck-shaped bar for revenue is not).

---

## 6. Numbers, dates, language

- Locale **en-US**: `1,234.5`. Decimal point, comma thousands. Write Excel format codes
  locale-neutrally: `#,##0.0`, never `#.##0,0`.
- Inventory turns: one decimal, with the multiplication sign — `8.9×`. Format code `0.0"×"`.
- Money: `$#,##0` — no decimals above $1,000, no scaling suffix without saying so in the subtitle.
- Percentages: `0.0%` for changes under 10 points, `0%` above.
- Fiscal years: label the axis `FISCAL YEAR ENDING 12/31` and the ticks with the year alone.
- Language: English. UPPERCASE for axis titles, table headers and end labels; sentence case for the
  punchline, the subtitle and annotations.
- Axis truncation: line charts may start above zero, but only with the baseline value written into
  the subtitle. Bars never.

---

## 7. Event lines — mandatory when the story has a cause

When a series bends and you know why, put a **vertical reference line at that period with the cause
written on it.** Build it out of data, never as a drawn shape:

- helper block on the sheet: two rows per event at the same x (the category's 1-based position,
  derived not typed), y = `0` and `1.00`, plus one cell holding the caption;
- add it as a **scatter with straight lines** on a hidden secondary axis pair, then `chart += ev`
  after the main chart is fully styled;
- secondary **value** axis pinned `0 … 1`, hidden;
- secondary **category** axis pinned to match the primary category placement — `1 … n` when the
  value axis uses `crossBetween="midCat"` (categories sit *on* the ticks), `0.5 … n+0.5` when it
  uses `between`. Getting this wrong parks the line between two years;
- style: 1 pt, dashed, `--bad`. It is context, not data;
- the caption is the **series name**, shown as the *top* point's data label only, so the words
  travel with the line;
- two events within three categories of each other: drop the second caption to `top = 0.88`. Past
  three events you have a timeline, not a chart — use a table.

This hidden secondary axis pair is the **one licensed secondary axis** in this guide. It encodes
nothing. Leave a note on the sheet saying so, or the next person will "fix" it.

---

## 8. Copy-paste: matplotlib

```python
import matplotlib as mpl

STYLE = {
    "figure.facecolor": "#FFF9F0", "axes.facecolor": "#FFF9F0",
    "figure.figsize": (8.0, 6.0), "figure.dpi": 150,
    "font.family": ["Segoe UI", "Calibri", "DejaVu Sans"], "font.size": 9,
    "text.color": "#2F2A26",
    "axes.edgecolor": "#E7D8C3", "axes.linewidth": 1.0,
    "axes.labelcolor": "#5C5149", "axes.labelsize": 8,
    "axes.titlelocation": "left", "axes.titlesize": 13, "axes.titleweight": "bold",
    "axes.titlecolor": "#2F2A26",
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "axes.grid.axis": "y",
    "grid.color": "#E7D8C3", "grid.linewidth": 0.9,
    "xtick.color": "#5C5149", "ytick.color": "#5C5149",
    "xtick.labelsize": 9, "ytick.labelsize": 9,
    "xtick.direction": "out", "ytick.direction": "out",
    "lines.linewidth": 3.0, "lines.solid_capstyle": "round",
    "lines.dash_capstyle": "round", "lines.markersize": 7,
    "legend.frameon": False,
    "savefig.facecolor": "#FFF9F0", "savefig.bbox": "standard",
}
mpl.rcParams.update(STYLE)

ACCENT, ACCENT_TEXT = "#EF5B23", "#C2400F"
CONTEXT, TEXT_2, INK, GRID = "#6E6259", "#5C5149", "#2F2A26", "#E7D8C3"
```

Keep `savefig.bbox` at `"standard"`: `"tight"` re-crops the canvas and the four fixed
bands in §4 stop landing where you put them.

Punchline title in Verdana Bold sits above the axes title slot:
`fig.text(0.02, 0.95, punchline, fontsize=16, fontweight="bold", fontfamily="Verdana", color=INK)`.

## 9. Copy-paste: openpyxl

```python
SURFACE, WASH, GRID = "FFF9F0", "FBEFDD", "E7D8C3"
CONTEXT, TEXT_2, INK = "6E6259", "5C5149", "2F2A26"
ACCENT, ACCENT_TEXT, GOOD, BAD = "EF5B23", "C2400F", "0B7A72", "B8352C"
FONT_HEAD, FONT_BODY = "Verdana", "Segoe UI"
PT = 12700
W_HERO, W_CONTEXT, W_AXIS = int(3 * PT), int(1.5 * PT), int(1 * PT)
TURNS_FMT = '0.0"×"'
CHART_W, CHART_H = 14.6, 10.9          # cm, 4:3
```

Per chart: `chart.roundedCorners = False`, `chart.style = None`, chart-area fill `SURFACE` with no
border, `chart.legend = None`, horizontal gridlines in `GRID` only, both axis lines `GRID` at 1 pt,
`majorTickMark = "out"`, `minorTickMark = "none"`, and `delete = False` **explicitly on both axes**
— openpyxl otherwise leaves Excel free to draw no tick labels at all.

Round line caps are not exposed by openpyxl's `LineProperties`; set `cap="rnd"` on the
`LineProperties` and check the render, or accept flat caps at 3 pt where the difference is one pixel.

---

## 10. Worked example

Data: inventory turns, `2020…2025`, `US = 8.9 8.4 7.0 7.2 6.5 8.1`,
`THE JONESES = 7.9 7.7 7.5 7.2 7.0 6.8`.

**As a line chart (the right answer — six periods, two series, a trend to show):**

- Punchline: *"Four years of losing to the Joneses — then someone found the warehouse keys."*
- Subtitle: *"Inventory turns per year, fiscal years ending 12/31. Higher is better; dashed line is
  the industry benchmark."*
- `US`: 3 pt `--accent`, hollow round markers. `THE JONESES`: 1.5 pt dashed `--context`.
- End labels at 2025: `US` in `--accent-text` bold, `THE JONESES` in `--context` bold. No legend.
- Value axis `0 … 10` step 2, `crossBetween="midCat"` so 2020 sits on the axis. Category axis
  labelled `FISCAL YEAR ENDING 12/31`.
- One event line at 2024, dashed `--bad`, caption *"New WMS goes live 🎉"* — the chart's one emoji
  and its one wisecrack, both spent here.
- Source: *"Source: data/inventory_turnover.csv | Turns = COGS ÷ average inventory. No inventory was
  harmed in the making of this chart."*

**As a bar chart (only if the question is "which year was worst", not "what happened"):** one column
per year, all `--context` except 2024 in `--accent`, value axis deleted, values printed above the
bars in `0.0"×"`, zero baseline, gap width 40 %. The benchmark becomes a single horizontal dashed
`--context` line labelled at its right end. Punchline names 2024 explicitly.

---

## 11. Pre-flight checklist

- [ ] Takeaway title is a **sentence with a point**, and it is funny once, not three times.
- [ ] Subtitle states the units and is completely straight.
- [ ] ≤ 1 emoji. None in axis titles, tick labels, or table headers.
- [ ] ≤ 1 wisecrack annotation.
- [ ] Accent is ≤ 10 % of the ink and marks exactly what the title claims.
- [ ] Text is `--accent-text` / `--ink` / `--text-2`, never `--accent`.
- [ ] Legend removed; every series direct-labelled at its end.
- [ ] Bars start at zero; any truncated line axis says so in the subtitle.
- [ ] Gridlines horizontal only, `--grid`, no vertical gridlines, no chart border.
- [ ] Both axes `delete = False`, tick marks `out`, tick labels horizontal (never diagonal).
- [ ] Number formats locale-neutral; turns carry `×`; no trailing-zero noise.
- [ ] Story has a known cause → event line present, built from data, caption on the line.
- [ ] Source line present, honest, with the dry sign-off.
- [ ] Excel: workbook **opened in Excel**, chart **exported to PNG and looked at**, no repair prompt,
      no label collisions.
- [ ] The Excel version and the notebook version of the same chart look like the same chart.

---

## 12. Decisions & rationale

| Decision | Value | Why | Chosen by |
| --- | --- | --- | --- |
| Register | punchline title, straight subtitle | Comedy needs a setup; analysis needs a fact. Splitting them gets both. | user (asked for light-hearted/funny) |
| Surface | warm cream `#FFF9F0` | Pure white reads corporate; cream reads friendly and still prints. | default |
| Accent | tangerine `#EF5B23` | Highest chroma in the set, so it wins every fight for attention; warm, so it belongs on cream. | default |
| Accent as text | `#C2400F` | `#EF5B23` is 3.24:1 on cream — legal for marks, illegal for text. | computed |
| Categorical order | tangerine → teal → mustard → plum | Neighbours differ ≈ 20 lightness points; survives greyscale and CVD. | computed |
| Good/bad pair | teal / claret | Green–red fails for ~8 % of men; teal/claret keeps the semantics and the eyes. | default |
| Fonts | Verdana Bold headlines, Segoe UI body | Chunky and friendly without Comic Sans; both are system fonts here, so Excel and matplotlib agree. | default |
| Emoji budget | 1 per chart | Zero is austere, two is a group chat. | default |
| Locale | en-US, `0.0"×"` for turns | Source material is US-style (`12/31`). | default |
| Legend | banned, direct labels only | The reader should never translate a colour into a name. | non-negotiable |
| Secondary axis | banned except the hidden 0–1 event-line pair | It encodes nothing and is invisible; it is scaffolding, not a second scale. | non-negotiable |
| Scope | this folder + charts that name this style | The project's serious style stays untouched. | default |

Paste into `CLAUDE.md` if this becomes the house style:
`Chart and slide work in mini_cases/inventory_turnover must follow style_chart.md in that folder.`
