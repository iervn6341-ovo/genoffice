# Excel-like (Sheets) — E2E Checklist

Read `README.md` first (how to run, harness, false failures). Statuses,
severity and lifecycle follow `../references/completion-gate.md`.

## Excel Test Suite

### X-01 Workbook Lifecycle

Test:

- launch;
- new workbook;
- open workbook;
- save;
- save as;
- reopen;
- close;
- unsaved-change warning.

---

### X-02 Worksheet Lifecycle

Test:

- add sheet;
- rename sheet;
- duplicate sheet;
- reorder sheet;
- delete sheet;
- switch between sheets.

Verify active sheet and persistence.

---

### X-03 Cell Input

Test:

- text;
- integer;
- decimal;
- zero;
- negative;
- date;
- time;
- percentage;
- boolean;
- long text;
- Unicode/CJK;
- formula;
- empty value.

Verify stored value versus displayed value.

---

### X-04 Edit Mode

Test:

- click cell;
- double-click edit;
- keyboard entry;
- Enter commit;
- Escape cancel;
- overwrite;
- edit existing value;
- click outside to commit where expected.

---

### X-05 Cell and Range Selection

Test:

- one cell;
- drag range;
- Shift range;
- row selection;
- column selection;
- select all;
- large range;
- selection after scrolling;
- selection after editing.

Verify active-cell indicator.

---

### X-06 Keyboard Navigation

Test:

- Arrow keys;
- Tab;
- Shift+Tab;
- Enter;
- Shift+Enter;
- Home/End where supported;
- Page Up/Down where supported.

---

### X-07 Formulas

Minimum cases:

```text
=A1+B1
=A1-B1
=A1*B1
=A1/B1
=SUM(A1:A10)
```

Where supported also test:

- AVERAGE;
- MIN;
- MAX;
- IF;
- COUNT;
- nested formulas.

For each major formula workflow:

```text
enter formula
→ verify formula
→ verify displayed result
→ change precedent
→ verify recalculation
→ undo
→ verify recalculation
→ redo
→ save/reopen
```

---

### X-08 References

Where supported test:

- relative references;
- absolute references;
- mixed references;
- ranges;
- cross-sheet references.

Then verify:

- copy/paste;
- fill;
- structural changes.

---

### X-09 Formula Errors

Where supported:

- divide by zero;
- invalid reference;
- unknown function;
- invalid syntax;
- circular reference.

Verify error display and recovery.

---

### X-10 Rows and Columns

Test:

- insert row;
- delete row;
- insert column;
- delete column;
- resize row;
- resize column;
- hide/show where supported.

Verify:

- values;
- formulas;
- references;
- formatting;
- selection;
- undo/redo;
- persistence.

---

### X-11 Merge / Unmerge

Test:

- horizontal merge;
- rectangular merge;
- unmerge;
- edit merged cell;
- select merged region;
- resize adjacent row/column;
- copy/paste merged region;
- undo/redo;
- save/reopen.

Verify non-anchor cell content policy.

---

### X-12 Formatting

Test implemented formats:

- font;
- font size;
- bold;
- italic;
- underline;
- text color;
- fill;
- borders;
- horizontal alignment;
- vertical alignment;
- wrap;
- number format;
- percentage;
- currency;
- decimal places;
- date formats.

Test both one cell and multi-cell ranges.

---

### X-13 Clipboard

Test:

- value copy/paste;
- formula copy/paste;
- formatted-cell copy;
- range copy;
- cut;
- repeated paste;
- overlapping destination;
- cross-sheet paste;
- cross-workbook paste if supported.

Verify relative reference adjustment.

---

### X-14 Fill and Series

Where supported:

- drag fill;
- number series;
- date series;
- formula propagation;
- repeated values.

---

### X-15 Sort / Filter

Where supported:

- ascending;
- descending;
- single-column sort;
- multi-column sort if supported;
- filter;
- clear filter.

Verify row integrity.

---

### X-16 Find / Replace

Where supported:

- find text;
- find number;
- replace once;
- replace all;
- no-match case.

---

### X-17 Freeze / View / Zoom

Where supported:

- freeze row;
- freeze column;
- unfreeze;
- zoom;
- large-sheet scrolling.

Verify selection and hit-testing after zoom/scroll.

---

### X-18 Undo / Redo Stress

Sequence:

```text
enter values
→ formula
→ format
→ insert row
→ resize column
→ paste
→ undo repeatedly
→ redo repeatedly
```

Verify exact restoration.

---

### X-19 Save / Reload

Create workbook containing:

- multiple sheets;
- formulas;
- merged cells;
- formatting;
- resized columns;
- dates;
- percentages.

Save/reopen.

Verify all relevant state.

---

### X-20 XLSX Interoperability

#### Import

Use representative XLSX containing:

- multiple sheets;
- formulas;
- formatting;
- merged cells;
- row/column sizes;
- styles.

#### Export

Create equivalent content in GenOffice.

Export XLSX.

Reopen in GenOffice.

If Microsoft Excel is available, open there too.

---

### X-21 Scale / Performance

Test realistic larger content:

- hundreds/thousands of populated cells;
- large paste;
- many formulas;
- long scrolling;
- large range selection.

Observe:

- latency;
- freezes;
- rerender defects;
- recalculation cost;
- memory growth.

---

### X-22 Invalid Operations and Recovery

Test safely:

- invalid formula;
- empty range action;
- merge impossible/invalid ranges where applicable;
- repeated delete;
- unsupported pasted content;
- malformed test import if available.

Verify no crash or silent corruption.

---

### X-23 Cross-Feature Workflows

#### Workflow A — Expense Sheet

```text
create headers
→ enter data
→ currency format
→ SUM
→ percentage
→ style header
→ resize columns
→ save
→ reopen
```

#### Workflow B — Formula Regression

```text
dependency chain
→ copy formulas
→ insert row
→ delete column
→ undo
→ redo
→ save/reopen
```

---

## Excel Final Regression Sweep

Verify at minimum:

```text
[ ] Launch
[ ] Workbook lifecycle
[ ] Worksheet lifecycle
[ ] Cell editing
[ ] Selection/navigation
[ ] Formula recalculation
[ ] Formatting
[ ] Clipboard
[ ] Undo/redo
[ ] Save/reopen
[ ] XLSX round-trip if supported
[ ] Runtime errors checked
```

---

---

## GenOffice notes — Sheets

- **State:** `window.__genofficeDebug.univerAPI` (`getActiveWorkbook()`,
  `getActiveSheet()`, `getRange(r,c).getValue()`).
- **Existing specs:** `e2e/sheets-*.spec.ts`.
  - Shortcuts, AutoSum (⇧⌘T / ⌥=), number formats: `sheets-excel-shortcuts`.
  - Insert/Delete Cells: `sheets-insert-cells-dialog`.
  - CSV: `sheets-csv-*`.
- **Traps:**
  - ⌘⇧S and ⌘⇧Z are native menu accelerators; click the Save As / Redo
    buttons instead.
  - Large workbooks stream: Find/Replace and undo must cover cells outside
    the loaded window.
- **Interop checks** (each has a regression test; recheck when touching
  save):
  - CSV keeps its delimiter, the `sep=` line and UTF-16.
  - Conditional formats: unsupported rules (timePeriod, aboveAverage) are
    kept verbatim, and dxf borders and number formats survive an edit.
  - Untouched notes keep their rich text and shape.
  - Range hyperlinks apply to every cell.
  - Unicode defined names work in shared formulas.
  - `date1904` workbooks get correct validation dates.
  - Formula cells are saved with cached `<v>` values.
- **Open the saved file in Excel** whenever the XML changes: there must be
  no repair prompt, and the values and formats must match.
- **Rust sidecar:** run `cargo test` in `apps/sheets/native/xlsx-engine`
  after changing reader code.
- **Real Excel is scriptable** where it is installed (Windows: `New-Object -ComObject Excel.Application`,
  headless). Read the behavior through COM first (`Range.DataSeries / FillUp / PasteSpecial / ClearHyperlinks`,
  `Workbook.TableStyles`, `ListObject.Unlist`), pin it in a unit test, then open the xlsx the app saved
  (`Workbooks.Open(path, 0, $true)`) and read links / notes / values / `ListObjects` back. Always wrap a
  probe in try/catch and `Quit()` in `finally`.
- **Windows runs:** `apps/sheets` `npm run build` needs `cargo`; without it use `npx electron-vite build`
  (the sidecar is prebuilt) and `npm run build -w @genoffice/shell` after touching `apps/*/src/main`.
  A new blank workbook is already file-backed ("Streaming … 1 rows available"): bulk edits must go through
  `ensureRangeResident` (`lazy-resident.ts`). Read raw numbers with `getRawValues()` — `getValues()` returns
  the display text of formatted cells (`#######` in a narrow column). Univer's `deleteNote()` pushes no undo
  entry; use `note-undo.ts`. Pointer-location pitfalls: `[aria-label="Themes"]` also matches the ribbon
  *group* — target `.cover-select[aria-label=…]`; a selection draws a blue wash over cells, so deselect
  before sampling screenshot pixels.
- **Specs for the Home / table / theme work:** `sheets-home-gaps` (Fill, Series, Clear, Cells menus),
  `sheets-home-gaps-save` (save → read xlsx → reopen; `QA_KEEP_DIR=<dir>` keeps the saved files for Excel),
  `sheets-paste-special`, `sheets-paste-extras`, `sheets-table-gallery`, `sheets-table-design`,
  `sheets-table-toggles` (Header / Total Row, Filter Button, Resize), `sheets-table-new-style`,
  `sheets-file-table-design`, `sheets-fill-extras`, `sheets-undo-after-save`, `sheets-theme-gallery`.
- **Tables across a Save:** a save reopens the workbook, so a table made in the session becomes the *file's*
  table. To test Table Design on a file table, make the table, Save As, wait for `tableAdds()` to be empty,
  then continue (see `sheets-file-table-design`). Undo after the save works through the *effective* look
  (journal over file); assert it by saving again and reading `xl/tables/tableN.xml`, then open the files
  in Excel (`QA_KEEP_DIR`, `ListObjects(1).TableStyle.Name`, `ShowTableStyle…`, `ShowTotals`, `ShowHeaders`).
- **Table looks are pixels:** compare the *set of colorful colors* (≥250 px each) before / after a restyle
  instead of absolute hex — chrome colors are the same in both; move the selection off the table first.
  Univer draws only the first range-theme rule that hits a cell, so a table's own plain theme must be
  removed before a laid theme is visible; overlapping rules fail silently.
- **Open → first undo:** the load-time `addTableTheme` entry once made the first ⌘Z after opening a
  workbook with a table a no-op; `undo-carry.ts` `DECORATION_MUTATIONS` now lists `add-range-theme` /
  `remove-range-theme`. If an undo "does nothing" right after open, look at the stack first.
- **Don't rewrite repo files with Python text mode on Windows** (LF → CRLF); `node tools/format-changed.mjs --write`
  repairs it, `--check` verifies.
- **Ribbon sweep** (`e2e/sheets-ribbon-audit.spec.ts`, ~14 min for all seven tabs; `AUDIT_TABS`, `AUDIT_CONTROLS`,
  `AUDIT_SKIP`, `AUDIT_OUT`): a "failed" click that stops at *scrolling into view* is usually something on top of
  the item — read `document.elementFromPoint` at its centre before blaming the item (the AI panel's 6px resize
  handle once covered the Size menu). The sweep restarts the window for each dropdown and each tab; state left by
  the previous tab's last button (an open note editor) makes the next tab's first dropdown fail spuriously.
  A dialog that Esc "does not close" must be checked from the ribbon (focus on the grid), not with the focus
  already inside it; `use-dialog-escape.ts` is the shared rule for modal dialogs.
- **`QA_KEEP_DIR` must exist** before the run: the specs that keep their saved workbooks `copyFile` into it in a
  `finally` and fail with `ENOENT` otherwise, which reads like a product failure.
