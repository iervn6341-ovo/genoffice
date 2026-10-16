# Table ops

> Edit an existing table element: cell text, merges, row/column structure, row heights, column widths, cell anchors, style presets and borders.

All ops take `target:{slide, el}` where `el` is the table's id (type `table` in
the outline). `row` and `col` are 0-based. Chart edits (`setChart`) live in the
`edit_chart` tool.

### setTableCell

`{row,col,paragraphs}`

Replaces one cell's text. Runs you do not restyle inherit the cell's current
formatting, so plain `{text}` runs keep size, color and bold.

| Field      | Type                                                                                            | Notes                   |
| ---------- | ----------------------------------------------------------------------------------------------- | ----------------------- |
| row, col   | integer                                                                                         | 0-based                 |
| paragraphs | array of `{runs:[{text, bold?, italic?, fontSize?, color?}], align?, bullet?, lineSpacingPct?}` | Same shape as `setText` |

```json
{
  "op": "setTableCell",
  "target": { "slide": 0, "el": "e_TABLE" },
  "row": 0,
  "col": 0,
  "paragraphs": [{ "runs": [{ "text": "Metric", "bold": true }], "align": "center" }]
}
```

Common mistakes

- Passing the cell text as a string: it must be a paragraph array.
- Row or column outside the grid: the op reports the failing (row, col).

### tableMerge

`{kind:"merge-right"|"merge-down"|"split",row,col}`

Merges the cell at (row, col) with its right or lower neighbor, or splits a
merged cell back into its grid cells.

```json
{
  "op": "tableMerge",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "merge-right",
  "row": 0,
  "col": 0
}
```

Common mistakes

- Merging across an existing merge boundary: the op refuses; split first.

### tableStructure

`{kind:"insert-row"|"delete-row"|"insert-col"|"delete-col",index,before?}`

Inserts or deletes a row or column. Insert goes after `index` unless
`before:true`.

```json
{
  "op": "tableStructure",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "insert-row",
  "index": 1
}
```

```json
{
  "op": "tableStructure",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "delete-col",
  "index": 0
}
```

Common mistakes

- Tables with merged cells refuse row/column surgery: split the merges first (`tableMerge` with `kind:"split"`).

### setTableRowHeight

`{row,hEmu}`

Sets one row's height in EMU (the table frame grows or shrinks accordingly).

```json
{ "op": "setTableRowHeight", "target": { "slide": 0, "el": "e_TABLE" }, "row": 0, "hEmu": 457200 }
```

### setTableCellAnchor

`{row,col,anchor:"top"|"middle"|"bottom"}`

Vertical alignment of the text inside one cell.

```json
{
  "op": "setTableCellAnchor",
  "target": { "slide": 0, "el": "e_TABLE" },
  "row": 0,
  "col": 1,
  "anchor": "middle"
}
```

### setTableColWidth

`{col,wEmu}`

Sets one column's width in EMU.

```json
{ "op": "setTableColWidth", "target": { "slide": 0, "el": "e_TABLE" }, "col": 0, "wEmu": 2743200 }
```

### setTableStyle

`{styleName} | {styleId?,firstRow?,lastRow?,firstCol?,lastCol?,bandRow?,bandCol?,keepFormatting?,shadingColor?,borderColor?,borderWidthPt?,borderPreset?}`

Restyles the whole table: apply one fixed-color preset by `styleName`, pick one
of PowerPoint's 74 built-in styles by `styleId` (gallery name such as
`"Medium Style 2 - Accent 1"` or its GUID; they follow the deck's theme
colors), or change individual region flags, cell shading and border lines. A
preset wins over the other fields; a preset or `styleId` clears direct cell
fills and borders like PowerPoint's style gallery (`keepFormatting: true`
keeps them). `slides read` shows a table's current `style`.

| Field                               | Type                  | Notes                                                                                                                                                                                                                                          |
| ----------------------------------- | --------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| styleName                           | string                | `none`, `lightGrid`, `zebraBlue`, `zebraGray`, `headerDarkBlue`, `headerOrange`, `noBorder`, `fullBorder`                                                                                                                                      |
| styleId                             | string                | Built-in gallery name or `{GUID}`: `No Style, No Grid`, `No Style, Table Grid`, `Themed Style 1/2 - Accent N`, `Light Style 1/2/3 [- Accent N]`, `Medium Style 1/2/3/4 [- Accent N]`, `Dark Style 1 [- Accent N]`, `Dark Style 2 [- Accent N]` |
| firstRow                            | boolean               | Header-row emphasis                                                                                                                                                                                                                            |
| lastRow, firstCol, lastCol, bandCol | boolean               | Total row, first/last column emphasis, banded columns                                                                                                                                                                                          |
| bandRow                             | boolean               | Banded rows                                                                                                                                                                                                                                    |
| keepFormatting                      | boolean               | With `styleId`: keep direct cell fills/borders instead of clearing them                                                                                                                                                                        |
| shadingColor                        | `#RRGGBB` or `"none"` | Cell fill for every cell                                                                                                                                                                                                                       |
| borderColor                         | `#RRGGBB`             | Border line color                                                                                                                                                                                                                              |
| borderWidthPt                       | number (pt)           | Border line width; > 0                                                                                                                                                                                                                         |
| borderPreset                        | `"all"` or `"none"`   | Draw all border lines, or clear them                                                                                                                                                                                                           |

```json
{ "op": "setTableStyle", "target": { "slide": 0, "el": "e_TABLE" }, "styleName": "zebraBlue" }
```

```json
{
  "op": "setTableStyle",
  "target": { "slide": 0, "el": "e_TABLE" },
  "styleId": "Medium Style 2 - Accent 1",
  "firstRow": true,
  "bandRow": true
}
```

```json
{
  "op": "setTableStyle",
  "target": { "slide": 0, "el": "e_TABLE" },
  "firstRow": true,
  "borderPreset": "all",
  "borderColor": "#BFBFBF",
  "borderWidthPt": 1
}
```

Common mistakes

- Passing a color name (`"blue"`): colors are `#RRGGBB`.
- Mixing `styleName` with the other fields: the preset is applied and the rest is ignored.
- A built-in name in `styleName`: gallery names and GUIDs go in `styleId`.
- Restyling one cell's text: that is `setTableCell` with styled runs, not this op.

### setChartLegendPosition

`{position:"b"|"t"|"r"|"l"|"none"}`

Repositions or removes a chart legend through a focused native chartSpace patch. Preserves imported series, linked data, unrelated formatting, and extensions. `none` removes the legend subtree, including its own formatting, as in a native chart without a legend.

```json
{ "op": "setChartLegendPosition", "target": { "slide": 0, "el": "e_CHART" }, "position": "b" }
```

### setChartQuickLayout

`{layout:"bottom-values"|"right-values"|"values-only"|"mac-1"|"mac-2"|"mac-3"|"mac-4"|"mac-5"|"mac-6"|"mac-7"|"mac-8"|"mac-9"|"mac-10"|"mac-11"}`

Applies a narrow native Quick Layout. The three named layouts show value labels with a bottom, right, or absent legend. PowerPoint Mac clustered-column Layout 2 adds native value labels to all three series, hides the value axis with c:delete=1, removes known major gridlines, and places the legend above the chart at gap 150/overlap -25. It preserves imported hidden-axis formatting rather than removing spPr/txPr. Layouts 1/3 hide value labels, move the legend right/bottom, and set gap/overlap to 150/omitted or 75/-25. Layout 4 adds native value labels to all three series, removes a safe chart title and known Mac major gridlines, and sets gap 75 with a bottom legend. Layout 7 removes a safe chart title, adds automatic axis titles and styled value-axis minor gridlines, and sets gap 300. Layout 8 adds automatic axis titles, removes the known Mac major-gridline and legend subtrees, and sets gap 0. Layout 9 adds native automatic category/value axis titles and turns out value-axis major ticks. Layout 10 adds value labels to the third series, with gap 75 and overlap 40. Layout 11 removes the chart title, sets autoTitleDeleted=1 and turns out both axes' major ticks. Layout 5 adds the value-axis automatic title and a native data table (with legend keys) and removes the legend. Layout 6 adds the value-axis title, removes the legend, and labels the last category of every series with value and series name; both require no existing axis titles, data table or series-local labels. The Mac layouts require no data table. Existing series-local labels are unsupported except compatible labels on all three series for Layouts 2/4 or the third series for Layout 10; Layout 7 accepts only absent or its own exact minor-gridline subtree. Layouts 2/4/8 accept only absent or the known Mac major-gridline structure, so unknown imported children are never discarded. Edits only the required native chart elements, preserving imported series, linked workbooks, chart relationships, unrelated styling, and unknown extensions.

```json
{
  "op": "setChartQuickLayout",
  "target": { "slide": 0, "el": "e_CHART" },
  "layout": "bottom-values"
}
```

### setChartAreaFormat

`{format:{area?:"chart"|"plot",fill?:"#RRGGBB"|null,stroke?:{color:"#RRGGBB",widthEmu:number}|null}}`

Changes chart-area or plot-area fill and outline through a focused native chartSpace/plotArea shape-property patch. Omit `area` for the whole chart. Other chart content and unselected effects remain intact. `null` sets a native no-fill marker.

```json
{
  "op": "setChartAreaFormat",
  "target": { "slide": 0, "el": "e_CHART" },
  "format": { "area": "plot", "fill": "#F2F2F2" }
}
```

### setChart (not-ai-callable)

`{patch:ChartEdit} — use the edit_chart tool instead`

Changes a chart's type, data, colors or elements. The `edit_chart` tool exposes
this with a validated schema.
