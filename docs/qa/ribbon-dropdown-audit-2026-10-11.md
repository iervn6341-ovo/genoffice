# Ribbon and dropdown audit — Slides (2026-10-11)

Goal: every ribbon control and dropdown should behave and look like Microsoft 365's — nothing
scrolls sideways, nothing runs off the window, and a command that exists in PowerPoint should be
reachable. Docs was audited in a second pass the same day (see the Docs section below). Sheets is being audited in its own session.

## Method

`e2e/_probe/audit.ts` (probe tooling, not committed) drives the built app through Playwright:

1. For each ribbon tab it lists every control, clicks it and records the popup that appears
   (including panels that are merely un-hidden, not newly mounted).
2. Per popup it checks: off-screen rectangle, horizontal scrollbar, clipped text, taller than the window.
3. It clicks the popup's items (all of a short menu, a spread sample of a long gallery) and, for items
   that run an action, compares the **document model** before the click, after it, and after Undo.
4. `compare-ref.ts` matches the collected labels against `office-ref.ts` (the Microsoft 365 command
   list per tab) to list missing commands.

Slides was audited with a text box selected at 1280 px (all 9 tabs, 70+ popups, ~190 actions) and
again at 1280 / 1100 / 1000 / 900 px for ribbon overflow.

## Findings and fixes

| Finding                                                                                                                            | Fix                                                                                                                                                                                                                                                      |
| ---------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Table "More Table Styles" panel had a horizontal scrollbar (440 px box, 7 × 63 px columns); light styles looked black in dark mode | Panel sizes from its content (`max-content`, 7 columns, vertical scroll only); previews sit on white paper (`--slides-paper`)                                                                                                                            |
| Home → Paragraph opened one 363 × 673 px panel, 13 px below the window                                                             | Replaced by PowerPoint's own controls: **Bullets ▾**, **Numbering ▾** (compact libraries, size / indent / color behind "Bullet Options"), Line Spacing ▾, **Columns ▾**, alignment row, Text Direction ▾, Align Text ▾                                   |
| Insert tab scrolled sideways at 1280 px (content 1341 px); Slide Show tab at ≤ 1000 px                                             | Groups fold into dropdown buttons on a narrow window (`Group fold`), lowest priority first; every tab now fits from 1280 down to 900 px                                                                                                                  |
| Every `.rb-drop` could outgrow the window                                                                                          | `max-height: calc(100vh - 150px)` with vertical scroll as a safety net                                                                                                                                                                                   |
| Slide layout dropdown was one 868 px strip                                                                                         | Three-column grid; added PowerPoint's Comparison, Content with Caption, Picture with Caption layouts                                                                                                                                                     |
| No Shape Effects on the ribbon although the engine has them                                                                        | **Shape Effects ▾** (shadow 3 × 3, reflection, glow, soft edges, link to the Format pane) on Home → Drawing and the Shape Tools tab                                                                                                                      |
| No text columns                                                                                                                    | `setTextBodyProps columns {num, spacing}` (engine + op + IPC + ribbon); resize-to-fit boxes balance their columns like PowerPoint                                                                                                                        |
| Transitions tab had only the effect strip                                                                                          | **Effect Options ▾** (direction / orientation / through-black), **Duration**, **Advance Slide** (On Mouse Click, After n s), Apply To All copies the whole setting; the slide show honors duration, push/cover direction, click-advance and auto-advance |
| Home → Section was only "Add Section"                                                                                              | **Section ▾**: Add, Rename, Remove, Remove All                                                                                                                                                                                                           |

## Tests

- Unit: `transition-detail.test.ts` (8), `text-body-props.test.ts` columns, `builtin-layouts.test.ts` (new layouts),
  `text-layout.test.ts` balanced columns, op-docs sync for the new op.
- E2E: `e2e/slides-ribbon-parity.spec.ts` — ribbon fits at 1280 / 1000 / 900 px, folded group flyout, Bullets /
  Numbering / Columns, Shape Effects, table style panel, Transitions tab.

## Known gaps (Slides, vs PowerPoint)

- Transitions: 11 effects (PowerPoint has 40+), no Sound; direction only for push / wipe / cover / pull / split.
- Animations: effect list is a subset (no Float In, Wheel, Random Bars, Color Pulse …), no Animation Painter,
  no Move Earlier / Later.
- Home: no Paste options menu, Convert to SmartArt, Replace Fonts, Select ▾, text Shadow, list-level commands;
  vertical-text layouts missing.
- Draw: no Lasso, Ink to Shape / Math, Ink Replay.
- Insert: no Photo Album, Action, Object, Symbol, Screen Recording as ribbon commands.
- Increase Indent on a plain (non-bulleted) paragraph changes nothing in the render model; PowerPoint indents it
  to level 2. Still to be verified against PowerPoint.

## Docs (second pass)

Audited with the cursor in a selected phrase, all 8 tabs, 107 popups on Insert alone.

| Finding                                                                                                                                                 | Fix                                                                                                                                                                     |
| ------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Insert tab scrolled sideways at 1280 px (1394 px of content); References 1092, Review 1203, View 1225 px at 900 px                                      | Groups fold into dropdown buttons (`RibbonFoldGroup`, now with an optional caption row) on Insert, References, Review and View; every tab fits from 1280 down to 900 px |
| The selection mini toolbar covered the last item of the Change Case menu (z-index 61 over the ribbon's 60), so the item could not be clicked            | The mini toolbar steps aside while any ribbon dropdown is open                                                                                                          |
| Change Case lacked Word's **tOGGLE cASE**                                                                                                               | Added (all 20 locales)                                                                                                                                                  |
| Design changes (page color, watermark, theme fonts and colors) and header / footer edits (page numbers) were not on the ⌘Z stack — only page layout was | They join the layout snapshot on the undo stack (`LayoutSnapshot.design`); ⌘Z / ⌘⇧Z restore them                                                                        |

Tests: `e2e/docs-ribbon-parity.spec.ts` (fit at 1280 / 1000 / 900 px, folded flyout, no folding at 1600 px,
page color + watermark undo/redo, tOGGLE cASE), `apps/docs/tests/case-transform.test.ts`.
All 47 existing Docs e2e specs still pass (6 skipped as before).

### Docs (third pass): the gaps above, filled

Layout, Home and Design now carry what Word has there. Every item below was driven through the real menus
(`e2e/docs-ribbon-gaps.spec.ts`) and the saved `document.xml` / `settings.xml` / `styles.xml` read back.

| Area                    | Added                                                                                                                                                                                                                                            |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Layout → Columns        | Left / Right (unequal widths), **More Columns…** (count, spacing, equal width, per-column width, line between)                                                                                                                                   |
| Layout → Breaks         | Page / Column / Text Wrapping next to the four section breaks                                                                                                                                                                                    |
| Layout → Line Numbers   | None / Continuous / Restart Each Page / Restart Each Section / Line Numbering Options (start, count by, distance) — written to `w:lnNumType`                                                                                                     |
| Layout → Hyphenation    | None / Automatic (`settings.xml w:autoHyphenation`, applied live)                                                                                                                                                                                |
| Layout → Margins / Size | **Mirrored** preset, gutter and "Mirror margins" in Custom Margins; **More Paper Sizes…**                                                                                                                                                        |
| Layout → Arrange        | Bring Forward / Send Backward (in front of / behind text), Align (page-relative)                                                                                                                                                                 |
| Home → Line Spacing     | Add / Remove Space Before / After Paragraph                                                                                                                                                                                                      |
| Home → Borders          | each side toggles on its own, All / Outside / Inside / Inside Horizontal, Horizontal Line, View Gridlines, **Borders and Shading…** (line style, weight, color, the rule between paragraphs, shading); `w:between` and the line style round-trip |
| Home → Styles           | No Spacing, Title, Subtitle, Quote, Intense Quote, List Paragraph (defined on first use, saved as real styles), **Styles pane**                                                                                                                  |
| Home → Paste            | Paste ▾: Keep Source Formatting / Merge Formatting / Keep Text Only for that paste                                                                                                                                                               |
| Design                  | **Paragraph Spacing** presets (rewrite Normal, undoable), **Page Borders…** (sides, style, weight, color, which pages, measure from, distance)                                                                                                   |

Engine changes behind them: `SectionSettings.colSep`, `lineNumbers: null` (remove), page borders written from their details (a custom box
is kept byte for byte on a margin edit — it used to collapse to the default box), unequal → equal columns, `ParaBorderLine.val` and
`w:between`, `applySettingsFlag` for `autoHyphenation` / `mirrorMargins`.

Also fixed on the way: Paste Options chip could not be opened (React 19 re-set every Craft icon's `innerHTML` on each render, so a
click whose mousedown landed on the glyph lost its target — `stableHtml`); margins + a document flag in one handler overwrote each
other (two `commitLayout` calls in one event).

### Docs gaps vs Word (still open)

- **Mirrored margins are saved and shown in the dialogs but the page view does not swap inside / outside on even pages**
  (the page is one continuous flow, so a per-page horizontal shift needs a layout change).
- Borders: no Diagonal borders, no Draw Table (no table-drawing mode); Inside Vertical is listed but disabled for paragraphs, as in Word.
- Line Numbers: no "Suppress for Current Paragraph" (the model reads `w:suppressLineNumbers` but has no editable paragraph attribute).
- Arrange: no Group, Selection Pane or Rotate in Layout (Rotate / Flip stay on the Picture Format tab); no separate Bring to Front vs
  Bring in Front of Text.
- Design: no Effects (theme effects have nothing to draw), no Set as Default, no Fill Effects; Page Borders has no Art or Shadow / 3-D.
- Styles: no Create a Style / Apply Styles; the Styles pane lists and applies, but has no Manage Styles or per-style Modify.
- No Paste Special dialog; Paste ▾ has no "Set Default Paste".
- Header / footer typed directly on the page is still outside the ⌘Z stack for the Design-tab items.

### Docs (fourth pass): compared with real Word

Word 16 (zh-TW) is installed on the QA machine, so the defaults of the third pass were read from Word instead of assumed: every fixture
is what `Word.Application` (COM, `SaveAs2` to .docx) wrote after the same ribbon choice. Word's style names follow its UI language — drive
it with the numeric `wdBuiltinStyle` constants, and pass `[object]` references to `SaveAs2`.

| Item                                                                  | Word wrote                                                                                            | GenOffice                                                                                                                                                                                    |
| --------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Home → Borders (Bottom, Outside)                                      | `w:space` 1 pt on the top, bottom and the rule between paragraphs, 4 pt on the left and right         | **Differed**: the borders hugged the text (space 0). `linesWithWordSpacing` (`paragraph-borders.ts`) gives each side that is new Word's gap; the ribbon and Borders and Shading… both use it |
| Layout → Line Numbers (Continuous / Restart Each Section / Each Page) | `w:lnNumType w:countBy="1" w:restart="…"`; Restart Each Page writes no `restart` (the schema default) | Same — the engine reads what Word wrote and writes it back unchanged                                                                                                                         |
| Design → Page Borders ▸ Box                                           | `w:pgBorders w:offsetFrom="page"`, four sides single / ½ pt / auto / 24 pt from the page edge         | Same                                                                                                                                                                                         |
| Layout → Columns ▸ Two, line between                                  | `<w:cols w:num="2" w:sep="1" w:space="425"/>`                                                         | Same                                                                                                                                                                                         |

Tests: `packages/docx-engine/tests/word-parity-sectpr.test.ts` (the `w:sectPr` Word wrote, read and written back),
`apps/docs/tests/paragraph-borders.test.ts`, `e2e/docs-word-parity.spec.ts` (2: the ribbon choices, the saved `document.xml` compared with Word's).
Only the four items above were compared. Styles (Title / Quote …), Paragraph Spacing presets, hyphenation and the rest of the third pass were
not put next to Word.

## Sheets

Audited with a 4×3 range selected (values, formulas, styles), every tab, every dropdown option pulled once, the model compared
before / after ⌘Z (cell values, formulas and resolved styles over A1:H14, merges, row heights, column widths).
`Pivot` was skipped by the matrix (it opens a dialog per cell and exceeded the one-hour budget); it is covered by its own specs.

| Finding                                                                                                                                                                                                                 | Fix                                                                                                                                                                                             |
| ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Ribbon Paste did nothing** — `univer.command.paste` was called without its params object and threw a `TypeError`; Copy / Cut worked, so ⌘V was the only way                                                           | Called with `{}`; `e2e/sheets-ribbon-fixes.spec.ts` copies and cuts a 3×3 block through the buttons and pastes it                                                                               |
| AutoSum over several columns and **Cell Styles** took one ⌘Z per column / per property (fill, font, border … each a separate undo entry)                                                                                | `undo-group.ts` folds the entries one ribbon command pushed into one; AutoSum writes all totals with one `setFormulas`                                                                          |
| Every non-Home tab scrolled sideways at 1280 px and below (Insert 1500+ px of content)                                                                                                                                  | 32 groups fold into dropdown buttons by priority (`RibbonFoldGroup`) — every tab fits at 1280 / 1000 / 900 px, nothing folds at 2400 px                                                         |
| Shapes gallery scrolled sideways; the six equation presets were clipped                                                                                                                                                 | Gallery sized to its grid; presets laid out as 3 + 2 + 1 cards so the long series formulas are not cut off                                                                                      |
| Conditional Formatting was a single button that opened the manage panel                                                                                                                                                 | Menu with Excel's ten entries (Highlight Cells, Top/Bottom, Data Bars, Color Scales, Icon Sets, New Rule, Clear Rules ×2, Manage Rules); each opens the rule panel preset to that kind          |
| Borders had only the single-edge presets                                                                                                                                                                                | Bottom Double, Thick Bottom, Top and Bottom, Top and Thick Bottom, Top and Double Bottom, Inside Borders, Inside Horizontal, Inside Vertical — each one undo step                               |
| Font size list stopped short of Excel's (8 … 72)                                                                                                                                                                        | `8 9 10 11 12 14 16 18 20 22 24 26 28 36 48 72`                                                                                                                                                 |
| Number format list had no way to the full dialog                                                                                                                                                                        | **More Number Formats…** opens Format Cells                                                                                                                                                     |
| **Format as Table ignored the picked style** — every style drew Univer's lavender default (the pick only reached the saved `tableStyleInfo`)                                                                            | `table-styles.ts` maps Light / Medium / Dark 1–14 to a Univer table theme in the Office accents (header band, stripes); one `addTableTheme` per table, folded into the same undo step           |
| **⌘Z after Format as Table left a ghost table in the saved file**, and the next Format as Table failed with "range overlaps table" (undo reached Univer's table, not the journal entry the save writes)                 | `table-sync.ts` follows Univer's delete-table / add-table mutations (undo and redo) in the journal                                                                                              |
| Format as Table over an existing table errored instead of restyling it                                                                                                                                                  | Inside an existing table it replaces it with the new style (same name, one ⌘Z step); across its edge it still refuses, as Excel does                                                            |
| "Medium 4 (orange)" saved `TableStyleMedium4`, which is gray in Excel                                                                                                                                                   | The entry is now **Medium 3 (orange)** → `TableStyleMedium3` (all 20 locales)                                                                                                                   |
| **Page Layout edits were not on ⌘Z** (Orientation, Size, Margins, Scale / Fit to width, Print Area / Titles, Breaks, Header & Footer, Themes): Undo silently took back the _previous cell edit_ and left the page setup | `journal-undo.ts` pushes one undo entry per page-setup edit (and per document Theme / Colors / Fonts change, `theme-undo.ts`) whose mutation restores the journal state; the ribbon re-reads it |
| Hiding the Formula Bar (`display:none`) made Univer log `The column width is less than 0…` on every selection afterwards                                                                                                | The row collapses to zero height but keeps its width                                                                                                                                            |

Tests: `e2e/sheets-ribbon-fit.spec.ts` (5), `e2e/sheets-ribbon-fixes.spec.ts` (7), `apps/sheets/tests/table-styles.test.ts`,
`table-sync.test.ts`, 20 new i18n keys (plus the Medium 3 rename) in all 20 locales.

Audit findings that are not defects: **Format Painter** only arms the painter (nothing to undo, so Undo reaches the previous edit);
**Insert row** was flagged once in the matrix but is clean in isolation (insert, ⌘Z, model identical, twice) — matrix residue.
The `column width is less than 0` error of the View tab was the hidden Formula Bar (above), not Page Break Preview.

### Sheets (fourth pass): the gaps above, filled, against real Excel

Excel 16 is installed on the QA machine, so every new behavior below was first **read from Excel through COM**
(`Range.DataSeries / FillUp / PasteSpecial / ClearHyperlinks`, `Workbook.TableStyles`, `ListObject.Unlist`), pinned in unit
tests, then driven through the real menus (`e2e/sheets-home-gaps*.spec.ts`, `sheets-paste-special.spec.ts`,
`sheets-table-gallery.spec.ts`, `sheets-table-design.spec.ts`, `sheets-theme-gallery.spec.ts`) and, for saves, read back from the xlsx and
opened in Excel.

| Area                                  | Added                                                                                                                                                                                                                                         | Observed in Excel (and kept)                                                                                                                                                                                                                                                                                                                                                                                                                            |
| ------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Home → Fill                           | **Up**, **Left**, **Series…** dialog (Rows / Columns; Linear / Growth / Date — day, weekday, month, year — / AutoFill; Trend; Step; Stop)                                                                                                     | Fill Up copies the bottom row (or the cell below a single cell) with relative references moved. A Series line starts from its first cell; a Stop value ends it inclusively and leaves later cells alone; one cell + Stop extends past the selection; Trend ignores Step / Stop and fits a least-squares line (Growth: log-linear); months clamp from the START (Jan 31 → Feb 29 → Mar 31); weekdays skip the weekend; the first cell's format is copied |
| Home → Clear                          | **Clear Comments and Notes**, **Clear Hyperlinks**, **Remove Hyperlinks**; Clear Contents now also takes the hyperlink, Clear All also takes notes and hyperlinks                                                                             | Clear Hyperlinks keeps text, note and the blue underline; Remove Hyperlinks drops the link formatting too; Clear Contents drops the link but keeps format and note                                                                                                                                                                                                                                                                                      |
| Home → Paste ▾                        | **Transpose**, **Values & Number Formatting**, **Formulas & Number Formatting**, **Paste Special…** dialog (All / Formulas / Values / Formats / All except borders / Column widths / … and number formats; Operation; Skip blanks; Transpose) | Transposed formulas: relative references swap their offsets, absolute references inside the block follow their cells. Transpose is never tiled and refuses a block that overlaps its source. Operations are destination ∘ source with blank = 0; a text / blank source leaves the destination; a formula destination is wrapped (`=(2+3)+1`); ÷ 0 shows #DIV/0!. Skip blanks leaves the destination under a blank source                                |
| Home → Cells                          | Insert ▾ (Insert Cells…, Insert Sheet Rows / Columns, Insert Sheet) and Delete ▾ (…, Delete Sheet) replace the four icon buttons; Sheet Rows / Columns act on every selected row / column                                                     | —                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| Home → Format as Table                | A gallery of all **60** built-in styles (Light 1–21, Medium 1–28, Dark 1–11) drawn as thumbnails; the table draws the style's real elements (outline and inside lines, stripes, first / last column) in the **workbook's own theme colors**   | Definitions read from `TableStyles` as theme-slot + tint references; exact resolved colors stored for the Office, Office 2013–2022 and Office 2007–2010 schemes (Medium 2 stripe under 2013–2022 = D9E1F2); Medium 3 is orange, Medium 4 gray                                                                                                                                                                                                           |
| Table Design (contextual tab)         | Table Name (Excel's naming rules), Banded Rows / Banded Columns / First Column / Last Column (saved in `tableStyleInfo`), the style gallery (keeps the options), **Convert to Range** (removes the table, bakes the style into the cells)     | Convert to Range keeps the style as cell formatting (header fill, bold, white text, stripes, lines)                                                                                                                                                                                                                                                                                                                                                     |
| Page Layout → Themes / Colors / Fonts | Galleries of cards with a **live preview** on hover that restores on leave / Esc / close and journals nothing; a click is one undo step                                                                                                       | —                                                                                                                                                                                                                                                                                                                                                                                                                                                       |

Behind them: every command is one ⌘Z step; `note-undo.ts` (Univer's `deleteNote()` pushes no undo entry — the **existing** Review → Delete
note in a file-backed workbook and an AI note removal were not undoable either; fixed at the root in `op-executor.ts`); hyperlink clears
undo through per-cell _effective targets_ (a save reopens the workbook from the file, so rolling the journal back would have left the saved
file without the links); `lazy-resident.ts` (a fresh workbook's file has one row, `ensureLazyRangeLoaded` refuses anything beyond it). The
save path gained optional `bandedColumns / firstColumn / lastColumn` on a table addition (schema → preload → main → gateway →
`xlsx-table-add.ts`): **rebuild the shell** after pulling that.

Tests: unit — `fill-series` (38), `paste-special` (29: transposed formulas, operations, skip blanks, number formats), `table-styles` (13: all
60 styles × 3 option sets — the Univer theme composed by Univer's own `RangeThemeStyle` equals the Excel-precedence cell look —,
Excel-resolved colors, Convert to Range baking), `table-design`, `table-add-options`; e2e — `sheets-home-gaps` (6), `sheets-home-gaps-save`
(3: save → read the xlsx → reopen), `sheets-paste-special` (7), `sheets-table-gallery` (3), `sheets-table-design` (5),
`sheets-theme-gallery` (2). The workbooks the app saved in these runs were opened in Excel (COM): links, legacy notes, series values, table
name / style / options read back as written.

An independent read-only review of this pass (13 findings, each re-checked against the code before acting) led to these fixes, all pinned by
unit or e2e tests (**not covered by a test**: the rollback path of a refused re-create, rename addressed by name, the Medium 2 default for a
style-less table, Clear All's undo fold, the large-selection load cap, the Step / Stop Trend lock, and the theme preview meeting ⌘Z or an
unmount): Paste Special resolves shared-formula followers (`si` without `f`, what a fill or tiled paste leaves) on the source and the
destination; Table Name refuses what Univer would refuse or rename (backslash, a sheet or defined name) and a failed re-create restores the
table instead of leaving a ghost in the journal; Escape in the Table Name box no longer renames, and a rename addresses the table by name;
the Table Design gallery restyles that table (it used to go through Format as Table's region guess and refuse next to other data); Convert to
Range keeps a cell's own fill / font / lines above the style and bakes a style-less table as Medium 2; Clear Contents / Clear All are one ⌘Z
step and do not stream a whole file first; Series replaces the old format down the line, no longer locks Step / Stop after a hidden Trend, and
a long weekday series is linear; transposed formulas leave structured references and absolute references to other sheets alone; a theme
preview that meets ⌘Z, or a gallery that unmounts while open, no longer leaves the preview behind.

### Sheets (fifth pass): the gaps of the fourth pass, filled, against real Excel

The fourth pass listed what was not there. Each item was read from Excel 16 through COM first, pinned in a unit test, driven through the
real UI, and — when it reaches the file — the saved workbook was opened in Excel and read back.

| Area                                                        | Added                                                                                                                                                                                                                                                                                                                                | Observed in Excel (and kept)                                                                                                                                                                                                                                                                                                                                                           |
| ----------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Paste Special                                               | **Comments**, **Validation**, **All using source theme**, **All merging conditional formats**; merged cells ride with All / Formats / All except borders; refused after Cut with a message                                                                                                                                           | A Comments paste replaces the destination's note with the source's and removes it where the source has none; Validation does the same with rules; "merging" keeps the destination's own conditional formats next to the source's; `PasteSpecial` after Cut fails in Excel too                                                                                                          |
| Page Layout → Themes / Colors / Fonts                       | The gallery lists Excel's own **48 document themes**, plus their color and font schemes (read from the Office install: `Document Themes 16`, `Theme Colors`, `Theme Fonts`), not only the 8 presets; a theme pick uses the preset's own colors                                                                                       | —                                                                                                                                                                                                                                                                                                                                                                                      |
| Table Design → Header Row / Total Row / Filter Button       | Header Row off starts the table one row lower and empties the names; Total Row adds "Total" and a `SUBTOTAL`; Filter Button only hides the buttons — each one ⌘Z step, drawn (header band, totals row), saved (`headerRowCount="0"`, `totalsRowCount`, `totalsRowLabel` / `totalsRowFunction`, an autoFilter without the totals row) | Opened in Excel: `ShowHeaders` / `ShowTotals` / `ShowAutoFilterDropDown` read back as set, the table range as Excel moves it                                                                                                                                                                                                                                                           |
| Table Design → Resize Table                                 | Dialog with the range typed as in Excel; refuses a range that drops the header row, cuts off the totals row or leaves the sheet                                                                                                                                                                                                      | `ListObject.Resize` refuses the same ranges                                                                                                                                                                                                                                                                                                                                            |
| Table Styles → **New Table Style…**                         | A style of the user's own (fill / font / borders per table element, with a live preview), listed under **Custom** in both galleries, drawn on the table, kept in the workbook and saved as a `<tableStyle>` with its `<dxf>`s in `styles.xml`                                                                                        | Opened in Excel: the table wears the style, and its header fill reads back as the colour picked                                                                                                                                                                                                                                                                                        |
| Table Design on **a table the file already had**            | The style gallery, Format as Table over it, and Banded Rows / Banded Columns / First Column / Last Column; the save patches the table part's `tableStyleInfo` (and writes the custom style a pick names). Display: the baked look is taken off the loaded cells and range themes draw the new one                                    | Opened in Excel: `TableStyle.Name`, the four `ShowTableStyle…` flags and the header colour read back (orange `#ED7D31` for Medium 3, `#1F3864` for the custom style); totals-row and header-less tables keep `ShowTotals` / `ShowHeaders`. Name, Resize, Convert to Range, Header Row, Total Row and Filter Button stay **session-table only**; the tab shows them disabled with a tip |
| Home → Fill → **Justify**                                   | Joins the text of a column of cells and breaks it again into lines no wider than the selection, in its first column; the lines that do not fit go below (the room must be empty); left-over cells are emptied; a number or formula in the selection is refused                                                                       | `Range.Justify`: same joining and breaking, measured in the selection's column widths (here measured with a canvas, so a line can break one word apart from Excel's)                                                                                                                                                                                                                   |
| Home → Fill → **Across Worksheets…**                        | Dialog picks the sheets (the tabs are not grouped here) and what to copy — All, Contents or Formats; formulas are copied as they are                                                                                                                                                                                                 | `Range.FillAcrossSheets`: formulas keep their relative references, so the copy reads the other sheet's cells                                                                                                                                                                                                                                                                           |
| Keyboard Delete on a hyperlink cell                         | Takes the link out of the journal like Clear Contents does, so the saved file has no link                                                                                                                                                                                                                                            | Clear Contents drops the link but keeps format and note                                                                                                                                                                                                                                                                                                                                |
| ⌘Z across a Save (page setup, document theme, Table Design) | The undo entries carry the _effective_ value (journal over file), so after a Save has emptied the journal an undo writes the value the saved file lacks back into the journal and the next save puts it in the file                                                                                                                  | e2e saves again after the undo and reads the xlsx back (Table Design: opened in Excel, the stripes option is the earlier one)                                                                                                                                                                                                                                                          |

Fixes that came out of it:

- **The intermittent Paste Special failure** of the fourth pass was a test artefact, not a bug: Excel COM `Copy()` run next to the suite replaced
  the OS clipboard, and Paste Special reads the copy id from it. 12 of 12 undisturbed repeats passed; the e2e spec writes the clipboard itself
  (`clipboard.writeText` in the main process) so it no longer depends on the machine.
- **The first undo after opening a workbook that has a table did nothing** (also after a Save): the plain theme registered over each file table
  (`addTableTheme`) pushes an undo entry made of `add-range-theme` + `set-sheet-table`, which `undo-carry` did not know as a load artifact, so it
  stayed on the stack as if it were an edit. Both ids are now in its list.
- A header typed after the table was made left the table part's column name stale; Excel refuses such a file. The gateway reconciles the column
  names with the header cells at save time (`reconcileTableHeaders`), making clashes unique the way Excel does.
- `note-undo.ts`: the facade's `createOrUpdateNote()` pushes no undo entry either; notes set through Paste Special Comments are one undo step.

Tests: unit — `themes` (8), `paste-special` (40), `table-design` (8), `table-design-toggles` (15), `table-add-options` (8), `table-styles` (20),
`xlsx-table-add` (13), `custom-table-styles` (11), `xlsx-table-style-add` (5), `xlsx-table-edit` (11), `file-table-design` (9), `fill-justify` (9),
`page-setup-undo` (8), `undo-carry` (+1); Sheets 3009 pass (1 skipped), gateway 60; e2e — `sheets-paste-extras` (6), `sheets-table-toggles` (7), `sheets-table-new-style` (2),
`sheets-file-table-design` (5), `sheets-fill-extras` (4), `sheets-undo-after-save` (2), `sheets-theme-gallery` (3), `sheets-home-gaps-save` (4),
`sheets-dialog-escape` (6), `sheets-ribbon-fit` (+1) — the whole Sheets e2e suite: 109 pass.
The save path gained `tableStyleAdditions`, `tableEdits` (and `readTableInfo` over IPC): **rebuild the shell** after pulling that.

### Sheets (fifth pass): the full ribbon sweep, re-run

`e2e/sheets-ribbon-audit.spec.ts` (now in the repo; the probe of the earlier passes never was) presses every dropdown option and button of
all seven tabs once, each dropdown on a fresh window, and records page errors, console errors, whether one ⌘Z brings the model back and
what opens a dialog. Last run (final code): **284 choices, 0 page errors, 0 console errors, 0 not undoable, 0 incomplete undos, 0 failed**;
17 controls are on the skip list (PivotTable / PivotChart — covered by their own specs — plus Pictures, Screenshot, Comment / Note editors,
Print setup toggles, Protect Workbook, the AI buttons, New Rule…, New Table Style…). `AUDIT_TABS`, `AUDIT_CONTROLS`, `AUDIT_SKIP` and
`AUDIT_OUT` narrow a run.

| Finding of the sweep                                                                                                                                                                                                                                                                                                                                                                                                              | Fix                                                                                                                                                                                                                 |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Size ▸ Letter … A5 could not be clicked** at a wide window: the dropdown reaches over the AI panel's right edge, and the panel's 6px resize handle (`.ai-panel-resizer`, `z-index: 20`) painted above the ribbon (`.excel-header`, `z-index: 10`) — a transparent strip that swallowed the clicks on the dropdown's items                                                                                                       | `.excel-header` is `z-index: 25`; `sheets-ribbon-fit` has a test that no option is covered at the handle's columns                                                                                                  |
| **18 dialogs ignored Esc** (Sort, Series, Fill Across, Go To, Icons, Recommended Charts, Header & Footer, Equation, Symbol, Name Manager, Insert Function, Goal Seek, Advanced Filter, Remove Duplicates, Consolidate, Subtotal, Allow Edit Ranges, Pivot, Screenshot, Select Data): most never handled the key, the others only heard it with the focus inside, and a dialog opened from the ribbon leaves the focus on the grid | `use-dialog-escape.ts`: a window listener, one stack of open dialogs, one Esc closes the top one only (Goal Seek not while it solves); `sheets-dialog-escape` opens six of them from the ribbon and presses Esc     |
| Zoom ▸ 75 / 100 / 125 % failed in a full run only                                                                                                                                                                                                                                                                                                                                                                                 | A harness flaw, not the app: the first dropdown of a tab ran in the window the previous tab's last button left (Review's note editor was open over the Zoom menu). The sweep now starts every tab on a fresh window |

What stays in the sweep's report and is expected: **Watch Window** is a non-modal panel (Esc does not close it, as in Excel); the recovery and
full-load prompts need a choice; one run in four reports a dialog (Consolidate) as not closing that closes in isolation (timing under load).

### Sheets gaps vs Excel (still open)

- **Table Design on a table the file had changes its style and the four banding / column options only.** Name, Resize Table, Convert to
  Range, Header Row, Total Row and Filter Button belong to tables made this session (the table part's columns, filter, ref and totals row are
  left as the file has them). A table that rows or columns edits moved in this session is left alone until the next save. A restyled table is
  drawn with display-only range themes: the save does not bake them into cells (Excel draws the saved style itself), and the print / PDF
  path reads cell formats, so until the workbook is saved and reopened a restyled table is likely to print without its new look — not
  checked in this pass. A table with more than 500 000 loaded cells is refused. A table saved without a style cannot be
  given "none" back after a Save (the style stays the file's). A table that wears a style the **file defines itself** (a corporate
  template's `<tableStyle>`) cannot be drawn here: the gallery does not list that style, an option toggle on such a table is refused with a
  message, and picking a gallery style first replaces it; the file's own style names are kept out of New Table Style (the save would keep
  the file's definition under a clashing name). Renaming a session table does not rewrite structured references to it elsewhere.
- The same display-only layers draw the totals row and the header-less look of a session table; the left / right outline of a header-less style
  can differ from Excel's by a line.
- Paste Special works from what the system clipboard names (the copy id Univer writes into it), so a paste of text copied elsewhere is refused
  with a message; a copy made in this window is needed, and a Save or AutoSave drops the cache — copy again after one; a rich-text source cell
  is pasted as its plain text.
- Fill: Across Worksheets picks the sheets in a dialog because the tabs are not grouped; Justify measures with a canvas, so a break can differ
  from Excel's by a word.
- ⌘Z across a Save for **page setup** writes back what the journal can say: orientation, paper, scale, fit, the three margin presets, a
  single print area, titles, header / footer and the manual page breaks. A file's own **custom margins** and a print area of **several
  ranges** cannot be written back after a Save (the journal has no wire for them), so an undo that would restore one leaves the setup the
  save wrote.
- The preview of a theme recolors cells that carry theme provenance (charts and conditional formats follow after a save, as for a real pick).
- The in-app look of a **file's own** tables comes from the Rust sidecar's palette, not from `table-style-defs.ts`; the two agree on the cases
  compared (Medium 3 header colour, custom style fill) but were not compared cell by cell for all 60 styles.
- Formulas recalculate slowly after a row insert in a large sheet (seconds); observed, not changed.
- The 17 non-English locales of the new strings are machine translations (Excel terms reused where known), not native-reviewed.

## Not done

- Docs: only the four items in "Docs (fourth pass)" were compared with real Word; the rest of the Docs work was checked against the OOXML
  schema and Word's documented defaults only.
- Native review of the machine-translated strings.
