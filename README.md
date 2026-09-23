# Data visualisation — 45-minute mini workshop

The deck, and everything that builds it. Sixty-two slides on why charts work, how to choose one, and
how to finish it. Every figure in the deck was rebuilt from its data by a script in the folder beside
it — nothing here is a screenshot of someone else's chart unless the slide says so.

```
45_min.md            the deck — Marp markdown; all styling lives in the frontmatter
45_min.pdf           the exported deck
references.md        the raw source URLs behind the REFERENCES slide
<topic>/             one folder per figure: the data, the script that draws it, images/
```

## Building the deck

```
npx @marp-team/marp-cli 45_min.md --pdf --allow-local-files -o 45_min.pdf
npx @marp-team/marp-cli -s .                      # live preview, rebuilds on save
```

`--allow-local-files` is not optional. Without it every `<img>` in the deck exports blank, silently.
It prints an "insecure local file accessing" warning; that is expected.

In VS Code, the **Marp for VS Code** extension gives the same preview (`Ctrl+Shift+V`) and an export
under `Ctrl+Shift+P` → *Marp: Export Slide Deck…*.

## The running order

| Slides | Section |
| --- | --- |
| 1 | Title |
| 2–6 | **Why visualisation** — the measles heatmap, then four products |
| 7 | **Data visualisation and storytelling** |
| 8–12 | **Science of visual processing** — the brain, reading patterns, magnitude, pre-attentive cues, perceived vs actual intensity |
| 13–23 | **Key questions** — what is the data type · what must the data say · who is the audience |
| 24–25 | **All the charts you'll ever need** — the 80% and the 95% |
| 26–54 | **Tips** — choose the right chart · substance over form · declutter · have a message · context matters · lay it out |
| 55–58 | **AI** — the Excel add-in, the terminal agent, and what each is for |
| 59 | **Know the chart elements** — the fifteen-part chart anatomy |
| 60–62 | References, thank you, questions |

## How a figure folder works

Most topic folders follow the same pattern:

```
build_xlsx.py        writes the .xlsx — data, chart, formatting, all in openpyxl
export_images.py     drives Excel to export the chart to images/ at 3x
style/xl_style.py    shared colours, fonts and axis defaults
style_chart.md       the written style guide, where the folder has one
images/              the PNGs the deck actually embeds
```

`export_images.py` scales every `sz=""` and `<a:ln w="">` in the chart XML before asking Excel to
enlarge the frame, because chart text is in **points** and does not scale with the frame. Skip that
step and you get three times the pixels with third-size labels.

Two folders — `inventory_turnover/` and `trade_balance/` — carry their own README and a notebook that
redraws the same chart in matplotlib. Read those first if you want the fullest worked example.

| Folder | What it makes |
| --- | --- |
| `00_us_diseases_chart/` | the WSJ measles heatmap, rebuilt |
| `ai_tips/` | the naive-prompt vs briefed-prompt pair, and `PROMPTS.md` |
| `anscombe_quartet/` | four identical summaries, four different pictures |
| `chart_chooser/` | the chart-choice diagram |
| `charts_elements/` | the annotated chart anatomy, and the title/subtitle/field/source layout figure |
| `closure_GV/` | the cluttered → decluttered sequence |
| `essential_visuals/` | the small catalogue of chart types |
| `four_products/` | the same four products, table against chart |
| `ft_visual_vocab/` | thumbnails for the FT families |
| `human_brain/` | the traced brain illustration |
| `identity_cues/` | pre-attentive attributes |
| `inventory_turnover/` | turns against the benchmark — fact, then explanation |
| `magnitude_effectiveness/` | the music-industry magnitude comparison |
| `makeup_use/` | six categories, 2019 against 2020 |
| `music_industry/` | US music revenue and technology adoption |
| `percieved_vs_intense/` | perceived versus actual intensity |
| `reading_patterns/` | the Z-pattern and the F-pattern |
| `rev_by_branch/` | revenue by branch |
| `sunburst/` | headcount by office, as a sunburst |
| `survey_feature/` | one survey table, asked six different questions |
| `trade_balance/` | Playfair's 1786 trade balance, redrawn |

`chart_layout/` builds an earlier version of the layout figure and is no longer used by any slide —
`charts_elements/images/anatomy_layout.png` replaced it.

## Third-party material

`ft_visual_vocab/` holds saved copies of the Financial Times *Visual Vocabulary* pages, including
their own `d3` and data files. `human_brain/` starts from a stock illustration. The datasets come
from RIAA, Kaggle and the World Bank. None of it is mine; it is here as teaching reference, and the
REFERENCES slide credits all of it in APA 7. Check the licence before reusing any of it.

## Presenter

Pham Tan Phat · tanphat.leo@gmail.com
