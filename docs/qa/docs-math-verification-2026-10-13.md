# Docs equation and symbol verification — 2026-10-13

## Result

The Docs equation and symbol workflow passed the opt-in Electron UI run against a scratch DOCX. The run authored a nested inline equation with a Greek symbol from the math gallery, inserted U+1D11E from the advanced picker at a bold red 14 pt cursor, saved, relaunched, re-edited the equation from its recovered OMML source, saved again, and relaunched a second time. The final DOCX contains native fraction, radical, n-ary operator, and script nodes; the symbol run retains Segoe UI Symbol, bold, color `C00000`, and size 28 half-points.

The final narrow 1280 px screenshots show the equation dialog, recovered-source edit dialog, and final reopened equation. The equation dialog and preview fit the window, and the inline-equation checkbox now renders at the intended compact size beside its label. The final editor status says `Missing document fonts: Segoe UI Symbol (substitutes shown)`, so the package font assignment is verified but rendering with that actual font face is not. The final screenshot also shows a `0 words` status for a document containing the anchors and equation; this math run does not validate word-count behavior, so that status remains unverified rather than an accepted result.

The final UI run was executed serially by the parent agent because this repo shares a single desktop UI slot. Its result is `passed` with both relaunches in [evidence.json](office-mac-math-evidence-2026-10-13/docs/2026-10-13T03-47-12-634Z-BUmx9A/evidence.json). The screenshots and saved packages are in the same evidence folder: [equation dialog](office-mac-math-evidence-2026-10-13/docs/2026-10-13T03-47-12-634Z-BUmx9A/equation-dialog-narrow-1280.png), [re-edit dialog](office-mac-math-evidence-2026-10-13/docs/2026-10-13T03-47-12-634Z-BUmx9A/equation-reopened-for-edit.png), [first reopen](office-mac-math-evidence-2026-10-13/docs/2026-10-13T03-47-12-634Z-BUmx9A/first-reopen-narrow-1280.png), [final reopen](office-mac-math-evidence-2026-10-13/docs/2026-10-13T03-47-12-634Z-BUmx9A/final-reopen-narrow-1280.png), [initial DOCX](office-mac-math-evidence-2026-10-13/docs/2026-10-13T03-47-12-634Z-BUmx9A/initial-saved.docx), and [edited DOCX](office-mac-math-evidence-2026-10-13/docs/2026-10-13T03-47-12-634Z-BUmx9A/edited-saved.docx).

## Source and regression review

- `MathWorkbench` and `SymbolPicker` are shared UI components. Math insertion uses the LaTeX-to-OMML subset and converter-generated MathML preview. The 42-template catalog remains reversible through the corresponding OMML-to-LaTeX subset; the symbol gallery inserts Unicode at the current LaTeX selection.
- Docs re-edit reads the current equation OMML and converts it to LaTeX. Unsupported OMML returns `null`, so full source editing is unavailable instead of regenerating a lossy approximation. The strict reverse-conversion tests cover unsupported structures and formatting properties.
- `SymbolDialog` uses stored marks or marks at the captured insertion position, then applies the selected font. A Docs unit regression verifies the inserted scalar retains bold, size, and color when the chosen font is applied, both in the ProseMirror node and after DOCX save/reparse.
- The independent reverse-conversion review found a hidden radical degree with content was previously flattened to `\sqrt{x}`. A focused regression was added; the core now rejects that case, and all 49 office-math authoring tests pass.
- A final source-only audit found a property-fidelity edge case for imported native function names and `lim` bases. The reverse converter now requires plain-style runs before recognizing `m:func` / `m:limLow` as named operators; unsupported italic/default function-name styles return `null`, while non-plain `lim` bases take the generic `\\underset` path so their math style remains italic. Regression cases cover omitted and italic `m:sty`; the focused office-math authoring suite passes all 51 tests.
- A check of the 20 Docs app locale shards found all 20 required math and symbol keys in every shard.

## Verification

- `npm run test -w @genoffice/docs -- equation-insert.test.ts`: 8 passed.
- `npm run test -w @genoffice/ui -- math-workbench.test.ts`: 3 passed.
- `npm run test -w @genoffice/office-math -- authoring.test.ts`: 51 passed, including all 42 generated templates and the strict-property/style fixtures.
- `npm run typecheck -w @genoffice/docs` and `npm run typecheck -w @genoffice/office-math`: passed.
- Focused ESLint and Prettier checks for the new E2E and regression tests: passed.
- `e2e/docs-math-verification.spec.ts` was discovered by Playwright and passed in the final parent-run desktop session (`2026-10-13T03:47:12Z`). It is opt-in with `GENOFFICE_DOCS_MATH_VERIFICATION=1` and uses a temporary scratch DOCX and user-data directory.
- Independent ZIP/XML inspection of both final captured DOCX files confirmed one native equation with fraction, radical, n-ary operator, script nodes, and the expected Greek token; the supplementary-plane symbol run retains Segoe UI Symbol, bold, red `C00000`, and 28 half-points in both packages.

## Limits

This verifies the supported authoring subset and one real Docs save/reopen/edit/reopen path. It does not establish full Word equation-gallery parity or preserve unsupported OMML through a full source edit; unsupported imports remain outside the authorable subset. Root additionally opened a scratch copy of the final edited DOCX in Microsoft Word for Mac: no repair dialog appeared, the structured equation rendered, and selecting it exposed Word's native Equation ribbon. This was a recognition/display check only; Word editing and re-saving were not tested. The system font substitution noted above limits the visual font-face claim to package metadata and the app's substituted glyph rendering.
