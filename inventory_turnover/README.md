# Mini case — annual inventory turnover

Six fiscal years of inventory turns, us against the industry benchmark. Rebuilt from `image.png`
twice: once as a faithful copy, once under a new light-hearted style guide written for this folder.

```
image.png                       the original, the target for the replica
style_chart.md                  "Warehouse Stand-Up" v1.0 — the style guide, written first
data/inventory_turnover.csv     the data
build_xlsx.py                   -> inventory_turnover.xlsx   (sheets: replica, standup)
build_notebook.py               -> inventory_turnover.ipynb   (--run also execs every cell)
inventory_turnover.ipynb        read the CSV, redraw both charts in matplotlib
images/                         the two matplotlib renders
```

```
C:\Python313\python.exe build_xlsx.py
C:\Python313\python.exe build_notebook.py --run
```

## The data

| Fiscal year | US | THE JONESES |
| --- | --- | --- |
| 2020 | 8.9× | 7.9× |
| 2021 | 8.4× | 7.7× |
| 2022 | 7.0× | 7.5× |
| 2023 | 7.2× | 7.2× |
| 2024 | 6.5× | 7.0× |
| 2025 | 8.1× | 6.8× |

Recovered from `image.png` by pixel-tracing, not by eye: the axes were calibrated from the tick
label centres (0 at y = 251.5 px, 2 at y = 218 px, …; 2020 at x = 58 px, 2025 at x = 308.5 px) and
each series' pixel column was averaged at each year. Values land within ±0.05 turns of the original
line, which is under one pixel at that scale. The palette was sampled the same way — the blue really
is `#357DBF`.

## The two renderings

**`replica`** — the original chart's geometry and palette, both measured out of the PNG. The
`# OF INVENTORY TURNS` axis title, the left-aligned `FISCAL YEAR ENDING 12/31`, the dashed benchmark
and the two end labels are all where the original put them.

**`standup`** — `style_chart.md` v1.0 applied literally: cream paper, one tangerine hero, hollow
round markers, Verdana Bold punchline, `US` / `THE JONESES` instead of column names, and a dashed
event line on 2024 carrying the one wisecrack and the one emoji the guide allows.

Both charts read the table on their own sheet, so editing a cell moves the line. The Excel and
matplotlib versions of each chart are two renderings of one piece of work and are kept identical.

## Notes for the next person

- Excel resolves a `Reference` to a string inside `add_data()`. Rename the sheet **before** building
  the chart or every series ends up `#REF!` — cost one silent-but-empty chart on the first build.
- A chart title that wraps needs `algn="l"` on its paragraph, or Excel centres it and indents the
  second line.
- `crossBetween="midCat"` (on the **value** axis) is what puts 2020 on the axis instead of half a
  slot to the right. The event line's hidden category axis must then be pinned `1 … n`, not
  `0.5 … n+0.5`, or the line lands between two years.
- Both workbooks were opened in Excel, checked for label collisions, and exported to PNG before this
  was called done — see `style_chart.md` §11.
