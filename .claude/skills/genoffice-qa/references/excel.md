# GenOffice Excel Development & Test Workflow

Benchmark against Microsoft Excel interaction semantics.

Always distinguish displayed values from stored cell values.

## Conceptual Execution Model

```text
pointer/keyboard
→ active cell/range
→ edit mode/command
→ worksheet model
→ formula/dependency engine
→ history
→ grid renderer
→ XLSX serializer
```

## Cell Input

Test:

- text;
- integer;
- decimal;
- zero;
- negative values;
- dates;
- time;
- percentages;
- booleans;
- formulas;
- empty values;
- overwrite;
- edit existing;
- cancel edit.

## Selection

Test:

- single cell;
- range;
- row;
- column;
- arrows;
- Enter;
- Tab;
- Shift combinations;
- selection after scrolling.

## Formula Workflow

For formula-related changes:

```text
enter formula
→ verify formula
→ verify displayed result
→ modify precedent cell
→ verify recalculation
→ undo
→ verify recalculation
→ redo
→ save/reopen
```

Where supported, test:

- relative references;
- absolute references;
- mixed references;
- ranges;
- invalid references;
- error propagation;
- dependency chains;
- circular references.

## Structural Changes

For row/column insert/delete:

Verify:

- values;
- formulas;
- references;
- formatting;
- merged ranges;
- selection;
- undo/redo;
- persistence.

## Merge / Unmerge

```text
select range
→ merge
→ verify region
→ edit
→ undo
→ redo
→ save/reopen
→ interact with adjacent cells
```

Verify content retention policy for non-anchor cells.

## Formatting

Test:

- number format;
- font;
- fill;
- border;
- alignment;
- wrap;
- row height;
- column width.

Include multi-cell selections.

## Clipboard

Validate:

- values;
- formulas;
- relative-reference adjustment;
- formatting;
- merged cells;
- destination overlap;
- undo/redo.

## Performance

For grid-related changes, test a larger worksheet.

Watch for:

- full-grid rerenders;
- quadratic operations;
- excessive recalculation;
- selection lag;
- large-paste instability.

## XLSX Round Trip

```text
XLSX import
→ inspect
→ modify
→ export
→ reopen
```

Validate formulas and formatting where supported.
