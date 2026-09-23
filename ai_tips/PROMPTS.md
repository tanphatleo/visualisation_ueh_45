# The prompts behind the AI slides

The two charts on the "same data, two prompts" slides come from `build_charts.py`,
which hard-codes both outcomes so the pair is reproducible. These are the prompts
they stand in for — what you would actually type to get each one.

## The data both prompts were given

| Channel | Revenue, FY2026 (billion VND) |
|---|---|
| Marketplace | 42.3 |
| Retail | 21.8 |
| Wholesale | 18.6 |
| Direct | 12.4 |
| Export | 7.1 |
| Affiliate | 3.2 |

## Prompt A — no style guide → `images/chart_no_style.png`

> Make a bar chart of revenue by channel from this table.

Nothing here is wrong. It is just unfinished: the tool has to guess at every
decision the brief did not make, and its guesses are defaults — a colour per
category, a legend for a single series, gridlines both ways, a title that
repeats the axis labels, rotated category names.

## Prompt B — with the style guide → `images/chart_styled.png`

> Make a bar chart of revenue by channel from this table. Follow `style_chart.md`:
> takeaway title stating the point, descriptive subtitle with the unit and the
> period, horizontal bars sorted descending, #BDBDBD on every bar except the one
> the title is about, which is #0F5499, value labels at the end of each bar to one
> decimal, no gridlines, no value axis, no box, category labels horizontal, source
> line at the foot. Aptos; title 20 pt semibold, subtitle 14 pt, labels 11 pt,
> source 9 pt.

In practice you do not type this every time. You write it **once**, as
`style_chart.md`, and then the prompt is:

> Make a bar chart of revenue by channel from this table, to `style_chart.md`.
> The point is that marketplace now beats retail and wholesale combined.

The second sentence is the part no style guide can supply: **which** bar gets the
accent, and what the title should say.

## Claude for Excel — prompts from Anthropic's own documentation

> Walk me through how the revenue number in cell C42 is calculated.

> What assumptions drive the gross margin forecast?

> Change the discount rate to 8% and update dependent calculations.

> Find the source of the #REF! error in the summary tab.

Source: <https://claude.com/docs/office-agents/excel>

## A terminal agent — the shape of the ask

> Read `feature_satisfaction.csv`. Build an .xlsx with a 100% stacked bar of the
> six answer options per feature, to `style/style_chart.md`, then export every
> chart to `images/` at 3x and show me the PNGs.

The difference is not the wording. It is that this one **runs**: it writes the
script, opens Excel, exports the pictures, and looks at what came out.
