# Building a sunburst — what the chartEx reference does not cover

`.claude/skills/excel-charts-openpyxl/references/chartex-charts.md` has the package-rewrite
recipe, and it holds. It documents **treemap** and **box & whisker** in detail; sunburst is
named in the `layoutId` list and nowhere else. These are the four things that cost a failed
build here, found by bisecting variants against a workbook Excel itself produced.

Reference workbook was made the way Rule 0 says to make one:

```powershell
$ws.Range("A1:D13").Select()
$ws.Shapes.AddChart2(-1, 120)      # 120 = sunburst
```

then the package was unzipped and read.

---

## 1. The hierarchy is ONE dimension over a MULTI-COLUMN range

This is the whole shape of a sunburst and it is not guessable from the treemap example,
which uses a single column.

```xml
<cx:strDim type="cat"><cx:f>'Sunburst'!$B$9:$D$20</cx:f></cx:strDim>
<cx:numDim type="size"><cx:f>'Sunburst'!$E$9:$E$20</cx:f></cx:numDim>
```

**One `cx:strDim`, three columns wide.** Office / Team / Employee are columns of a single
range, not one `strDim` per ring. Point it at a single column and you get a doughnut with one
ring — no error, no warning, just a chart that has quietly lost two levels.

Excel writes hidden `_xlchart.v1.N` defined names for these. A plain range works and is one
less moving part, as `chartex-charts.md` already says for treemap.

---

## 2. `cx:dataPt idx` walks EVERY wedge, depth first — not the data rows

The expensive one. For a treemap, `idx` enumerates the tiles, which are the data rows. A
sunburst has more wedges than rows, and the indexing follows the tree:

| idx | wedge | ring |
| --- | --- | --- |
| 0 | Boston | 1 |
| 1 | Boston / Sales | 2 |
| 2 | Lisa T. | 3 |
| 3 | Phil B. | 3 |
| 4 | Marc G. | 3 |
| 5 | Boston / IT | 2 |
| 6 | Andrea T. | 3 |
| … | | |
| 21 | Dennis M. | 3 |

**12 rows produce 22 wedges** (3 offices + 7 teams + 12 people). Colour them as 12 leaves and
the fills land on the wrong wedges — and the chart still renders, perfectly happily, with no
sign that anything is wrong except the picture. Everything past the last `cx:dataPt` falls
back to the theme palette, which is how a green office turned up in an all-blue chart.

`wedge_order()` in `build_xlsx.py` generates the sequence.

---

## 3. `cx:title/@align="l"` makes Excel refuse the workbook

Not the chart — the **workbook**, with the same "unable to get the Open property" that a
malformed package gives, so it looks like a packaging bug and it is not.

```xml
<cx:title pos="t" align="ctr" overlay="0">        <!-- "l" is rejected -->
  <cx:tx><cx:rich><a:bodyPr/><a:lstStyle/>
    <a:p><a:pPr algn="l">                          <!-- left-align HERE instead -->
```

Bisected across six title variants: `align="l"` was the only failure. `<a:pPr algn="l">`,
a full `<a:bodyPr rot="0" spcFirstLastPara="1" …/>`, `<a:latin typeface="Aptos"/>` and an
apostrophe in the title text all open fine.

A `cx:title` on an **axis** is still fatal, as the reference says. A sunburst has no axes, so
that trap does not arise here.

---

## 4. openpyxl's worksheet root has no `xmlns:r`

Generic to injecting a drawing into **any** openpyxl sheet that has no ordinary chart on it,
so it is not really a sunburst finding — but it is what broke the first build, and the
reference's recipe assumes the sheet already owns a drawing.

openpyxl writes:

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
```

with **no `xmlns:r`**, because a sheet it built has no relationship-bearing element that needs
one. Append `<drawing r:id="rIdDrawing"/>` to that and the part is not well-formed XML, so
Excel rejects the entire workbook. Declare the prefix on the element itself:

```python
'<drawing xmlns:r="%s" r:id="rIdDrawing"/></worksheet>' % R_NS
```

`<drawing>` sits near the end of `CT_Worksheet`'s sequence — after `pageMargins`, immediately
before `</worksheet>`, which is where Excel puts it too.

The alternative the reference assumes — anchor an ordinary chart first so openpyxl creates the
drawing — also works, and costs a chart you may not want on the sheet.

---

## Colour, briefly

One hue per office on a descending-lightness ramp, largest office darkest: `0F3762` /
`0F5499` / `3F73A6`. All three clear 4,5:1 against white, so **one** label colour serves the
whole chart and there is no per-label contrast switch to compute. One step lighter on the
house ramp (`93B2D1`) sits at 2,1:1 and would have forced ink labels on the pale office and
white everywhere else.

Rings are told apart by radius and the 1 pt white separators, not by shade, so an office reads
as one solid arc from centre to rim.

---

## Verified

- opens with no repair prompt (`verify_workbook.ps1`)
- `check_chart_overlap.ps1`: 23 text boxes, **0 overlaps** — but note it counts labels Excel
  has already auto-shrunk to fit its wedge, so it proves less here than on a bar chart
- 0 formula errors; the sheet total reads 12
- exported to PNG at 3× and looked at
