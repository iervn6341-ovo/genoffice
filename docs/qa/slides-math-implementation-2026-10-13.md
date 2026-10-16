# Slides Equation and Symbol implementation — 2026-10-13

## Result

Insert → Equation now opens the shared math workbench with 11 structure categories, a contextual math-symbol gallery, searchable symbol subsets, and professional/linear preview. It submits LaTeX source through `slides:apply-txn` as the existing `insertEquation` operation. The operation writes an OOXML `mc:AlternateContent` run whose PowerPoint choice contains `a14:m`/OMML and whose fallback contains linear text. It no longer inserts a Cambria Math Unicode approximation. The transaction owns validation, undo, dirty state, serialization, and saved-slide rebuild.

Selecting a native equation and choosing Equation or Edit Equation, or double-clicking it, opens the same workbench for in-place editing. `replaceEquation` changes the one supported OMML run inside the existing shape. It preserves the shape's identity, geometry, animation references, other runs, and neighboring elements. The editor derives its initial source from the current OMML with a strict reverse converter. If an imported equation uses unsupported OMML, the UI leaves it read-only and the operation refuses replacement; it does not flatten those original bytes. This also makes external PowerPoint edits authoritative rather than trusting potentially stale metadata.

Insert → Symbol opens the shared picker with font, subset, search, Unicode code point, and recent symbols. When a text editor is active, the ribbon saves its caret/range and the chosen glyph is inserted into that selection with the chosen font. Without an active caret, it creates a normal text box containing the glyph. The new labels and 11 category/6 subset names are localized in all 20 Slides locales.

The render tree now exposes MathML and reversible source derived from native OMML. A standalone, unrotated equation gets a browser MathML preview over the Konva shape with `pointer-events: none`, so selection and hit testing stay on the underlying shape. Mixed text and rotated equations retain the linear fallback on the slide canvas. The dialog's professional preview renders structured math for all gallery structures. The saved PPTX always contains native OMML; the screen overlay is not serialized into the presentation.

## Validation

- `packages/pptx-ops/tests/equation-edit.test.ts`: insert → save → reopen → replace → save → reopen; reversible gallery templates; mixed rich text and neighboring shape preservation; unsupported OMML rejection.
- `packages/pptx-render/tests/equation-preview.test.ts`: MathML/reverse source metadata and mixed-text overlay guard.
- Existing `apps/slides/tests/animation-equation-ops.test.ts` and `packages/pptx-engine/tests/equation-table.test.ts`: native OMML and imported table-equation preservation.
- `apps/slides/tests/equation-symbol-actions.test.ts` and `edit-selection.test.ts`: IPC operation payload, symbol fallback text box, saved-caret insertion, and selection restoration.
- Focused suite: 6 files, 22 tests passed. TypeScript checks passed for `pptx-ops`, `pptx-render`, and `apps/slides`. ESLint had no errors; three existing `App.tsx` React-hook warnings remain. Prettier applied to changed math files and Slides locale shards.
- The complete `pptx-ops` package suite passed after documenting `replaceEquation`: 4 files, 36 tests.

## Remaining verification and limits

- Native PowerPoint for Mac and Electron UI were not launched in this implementation pass; a separate reviewer will verify the screen flow and saved file in those environments.
- Slide-image/PDF export currently uses the linear equation fallback because the HTML MathML overlay belongs to the editing canvas. Mixed text and rotated equation shapes also use linear fallback on screen. PPTX save/reopen/re-edit is native OMML.
- Imported OMML outside the strict reverse converter remains preserved and readable, but not editable through the current workbench. The user can create a new equation without changing the imported bytes.
- No claim is made about exact visual parity with PowerPoint's native equation renderer or its entire 188-item Equation Symbols gallery. The shared workbench supplies the 11 observed structure categories and its own searchable Unicode galleries.
