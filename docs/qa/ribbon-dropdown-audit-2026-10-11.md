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

### Sheets gaps vs Excel (still open)

Checked against `ExcelShell.tsx` / `ribbon-actions.ts`:

- Home → Fill has Down / Right only (no Up / Left / Series…); Clear has All / Formats / Contents (no Clear Comments / Clear Hyperlinks).
- Paste Special has Values, Formulas, Formats, Column Widths, All Except Borders — no **Transpose**, no Values & Number Formatting, no Paste Special dialog.
- Cells group: four separate Insert / Delete Row / Column buttons. Excel's Insert Cells / Delete Cells (shift right / down / left / up) dialogs exist but only through ⌃⇧= / ⌘-, with no ribbon dropdown; Insert / Delete Sheet are not in these menus.
- Format as Table offers 6 of Excel's 60 built-in styles and no New Table Style; the in-app look is an approximation of Excel's (Excel draws its own from the saved style name).
- No Page Layout → Themes gallery with live preview.

## Not done

- No real Word / Excel side by side for the new Docs and Sheets items; the behaviour was written from Office's documented menus.
- The 15 non-English locales of the new strings are machine translations, not native-reviewed.
