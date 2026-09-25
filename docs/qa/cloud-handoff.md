# GenOffice QA — Cloud Handoff (2026-09-25)

Branch: `qa/local-fixes`. Written for the cloud session that picks up the
remaining fixes. Follow the `genoffice-qa` skill
(`.claude/skills/genoffice-qa/`, pushed temporarily on this branch) —
especially `references/completion-gate.md` and `templates/test-report.md`.

## 0. Ground rules for the cloud session

- **Environment limits.** The cloud box is Linux. It has no macOS Electron
  GUI, no ⌘ keys, no second display, no Microsoft Office and no local test
  PDF. So:
  - Do **not** mark UI, keyboard, focus, drag or rendering behavior as PASS.
    Mark it `NOT TESTED (needs local e2e)`.
  - Verification here = unit tests, typecheck, lint, build.
  - The local Mac machine will run Playwright e2e and compare against real
    Word/Excel afterwards.
- **Repo rules** (see `CLAUDE.md`):
  - Theme tokens only in renderer CSS.
  - New UI strings go into all 20 locale shards.
  - After touching `apps/*/src/main`, rebuild the shell.
- **Commits.** One commit per task, and each commit message carries its own
  test report. Do not rewrite or squash the existing commits.
- **Do not touch** the Playwright specs' expectations in order to make them
  pass. Add new e2e specs for new behavior. They will be run locally.

### Commands

```bash
npm ci
npm run typecheck
npm run lint
npm test -w @genoffice/docs      # likewise sheets / slides / pdf / markdown / html
cargo test --manifest-path apps/sheets/native/xlsx-engine/Cargo.toml
npm run build:all                # must succeed before handing back
```

Known environment noise:

- `apps/sheets/tests/xlsx-sidecar-cancel.test.ts` fails under Node 26. It
  should pass on Node 22 (CI).
- `packages/html2docx` tests need `CHROME_PATH` pointing at a Chromium
  build.

## 1. Already fixed on this branch (verified locally)

These were verified with unit tests plus Playwright e2e on macOS. The last
full e2e run was 159 passed / 6 skipped / 1 failed; that failure was fixed
afterwards and passed 4/4 runs.

### PDF

- Redaction undo/redo.
- Rotate pages now also rotates pending text/images.
- The Crop dialog fits the window.
- N-up merge honors CropBox and /Rotate.
- Annotation appearances are flattened on merge.
- Page import lands at the correct index after a deleted page.

### Docs

- Home/End on Mac.
- List AutoFormat (`* `, `- `, `1. `, `a) `).
- Clear Formatting (works at a caret; keeps comments, tracked changes and
  links).
- Heading/Normal restyle survives reopen.
- Continue Numbering keeps the list kind.
- Table handle selection.
- Header/footer mixed run formatting is preserved.
- Dragged picture position survives save.
- Double-click past the line end no longer merges paragraphs.

### Markdown / HTML

- Home/End.
- Double-click past the line end (Markdown).
- Race when editing right after an insert (HTML).

### Sheets

- AutoSum on a single cell (⇧⌘T / ⌥=).
- Ctrl+Shift number-format shortcuts.
- Normal font echoes Calibri.
- CSV save keeps the delimiter, `sep=` and UTF-16.
- Unsupported CF rules are kept, and CF borders/number formats survive a CF
  edit.
- Untouched notes keep their formatting and shape.
- Streamed Replace All is a single undo.
- Unicode names in shared formulas.
- Range hyperlinks apply to every cell.

### Slides

- Show controls/keys and the presenter display.
- Notes editor.
- Font extras and text case.
- Ribbon fold.

### E2E infrastructure

- Background mode is the default (`GENOFFICE_E2E_BACKGROUND=1`). Use
  `foreground: true` per spec when a spec needs it.

## 2. Open issues — cloud task list

Each task follows the skill's flow: Classify → Define Expected Behavior →
Reproduce/Characterize → Trace → Implement → Narrow Verification →
Lifecycle. Priority order is top to bottom.

### C1 · Docs — style-inherited bold/italic cannot be turned off (Medium→High)

- **Scenario**
  - Put the caret in Heading 1 text; its bold comes from the paragraph
    style.
  - The Bold button shows inactive.
  - Clicking Bold adds direct bold; clicking again goes back to inherited
    bold. Non-bold text is never reachable.
  - Format Painter from Normal text onto a heading fails the same way.
- **Word behavior**
  - The Bold button reflects the *effective* bold, including bold inherited
    from the style.
  - Toggling it off writes `<w:b w:val="0"/>` as a direct run property.
- **Likely layers**
  - Tri-state bold/italic mark attrs in `apps/docs/src/renderer/editor/convert.ts`
    (import `w:b w:val=0` → an explicit-off mark).
  - The generator/serializer that writes the run properties.
  - Merging with rawRPr.
  - An effective-format resolver used by `components/ribbon-format-state.ts`
    (style chain → effective bold).
  - The Bold command toggle.
  - Format Painter.
- **Unit tests required**
  - DOCX round-trip of `w:b w:val="0"` under a bold style.
  - The ribbon state reports bold for inherited bold.
  - Toggling produces explicit off.
  - Format Painter onto a bold style.
- **Local e2e to add** (spec file only; do not run it in the cloud)
  - Heading 1, turn off Bold, type, save/reopen: the text is not bold.
  - Then undo/redo.

### C2 · Sheets — date validation shifted 1462 days in `date1904` workbooks (Medium)

- **Scenario**
  - Open an xlsx that has `<workbookPr date1904="1">`.
  - Add a Date validation between 2026-01-01 and 2026-12-31.
  - Save.
  - Excel then reads the bounds as 2030/2031.
- **Fix**
  - When `date1904` is set, serialize the validation date serials through
    the workbook's epoch.
  - Check the import direction too, and CF date literals if they are
    affected.
- **Tests**
  - Gateway unit tests for both epochs.
  - A round-trip test.

### C3 · Sheets — formula cells saved without cached `<v>` values (Medium)

- **Symptom**: other readers (Quick Look, pandas, mobile viewers) show empty
  cells until Excel recalculates.
- **Fix**
  - Write the last computed value as `<v>`, with the matching `t=` for
    strings, booleans and errors.
  - Keep `fullCalcOnLoad` behavior sane.
- **Tests**
  - Unit tests that check the XML for number, string, boolean and error
    results.

### C4 · Sheets — Insert Cells dialog (⌃⇧= / ⌘⇧+) missing (Medium, feature)

Excel's dialog offers four options:

- Shift cells right
- Shift cells down
- Entire row
- Entire column

Scope:

- The dialog component, using existing dialog patterns (`role`, focus,
  Esc).
- Wiring to Univer's insert commands.
- i18n in all 20 locales.
- Undo as a single step.
- Unit-test the command mapping.
- Add an e2e spec for local runs.

Also check "Delete Cells" (⌘-) parity.

### C5 · Docs — Find is not Unicode case-insensitive (Low→Medium)

- **Scenario**: searching `straße` / `STRASSE`, Greek final sigma, or
  Turkish İ does not match the way Word does.
- **Fix**: case-fold with `toLocaleLowerCase` plus a normalization (NFC)
  mapping that keeps the offsets correct.
- **Tests**: unit tests on the search helper, including offset mapping for
  highlights.

### C6 · Docs — inserting a section break auto-saves and wipes undo (High, larger)

- **Scenario**
  - Open a DOCX with AutoSave off.
  - Type something.
  - Insert a section break.
  - The file is written immediately, and after the reparse Undo is gone.
- **Direction**
  - Model sections locally in the editor (a node or doc attribute) instead
    of going through save + reparse.
  - Serialize them at save time.
- **If too large**: produce a design note plus the minimal change that at
  least avoids the silent write and keeps undo.
- **Tests**: unit tests for the section model round-trip.

### C7 · Docs — page layout (margins, orientation, header/footer, page numbers) outside undo (High, larger)

- **Scenario**
  - Type a word, then change the margins.
  - ⌘Z undoes the word, not the margin change.
- **Direction**: route layout changes through a ProseMirror transaction,
  either as doc attributes or a history-aware step, so they share the undo
  stack.
- **Tests**: unit tests on the transaction/undo stack.

### C8 · PDF — Crop / page size not undoable; watermark text not searchable (Low)

- Add crop and page size to the `EditSnapshot` buckets, the same way
  redactions were added (`apps/pdf/src/renderer/edit-state.ts`,
  `edit-ops/registry.ts`).
- Include `apps/pdf/tests/edit-ops.test.ts` coverage.

## 3. Test plan after handoff (run locally on the Mac)

Per `references/completion-gate.md`. The cloud fills the *Cloud* column; the
local session fills the rest.

| Task | Cloud: unit/type/lint/build | Local e2e | Office comparison | Lifecycle (undo → redo → save → reopen → continue editing) |
|---|---|---|---|---|
| C1 | required | new spec | Word: Heading 1 bold toggle | required |
| C2 | required | — | Excel opens the file; validation accepts 2026 dates | save/reopen |
| C3 | required | — | Excel opens without repair; Quick Look shows values | save/reopen |
| C4 | required | new spec | Excel Insert Cells dialog | required |
| C5 | required | docs find spec | Word Find | N/A |
| C6 | required | new spec | Word section break | required |
| C7 | required | new spec | Word layout undo | required |
| C8 | required | pdf-edit spec | Acrobat/Preview (optional) | required |

After merging back locally:

1. Rebuild everything.

   ```bash
   npm run build:all
   ```

2. Run the full e2e suite in background mode.

   ```bash
   GENOFFICE_E2E_NO_VIDEO=1 npx playwright test --config e2e/playwright.config.ts --reporter=line
   ```

   The baseline is 160 pass and 6 skip.

3. Do the Office comparison for each task in the table.
4. Write the QA report using `templates/test-report.md`.

## 4. Status legend

`PASS` / `FAIL` / `NOT TESTED` / `BLOCKED` / `N/A`. Never mark anything
PASS without executing it.
