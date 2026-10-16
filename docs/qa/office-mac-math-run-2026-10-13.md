# Equation and symbol integration — 2026-10-13

Status: complete for this authoring/persistence wave. Docs, Sheets, and Slides passed built-app save/reopen/re-edit acceptance; independent source and evidence reviews passed.

## Scope and team

This wave implements equation authoring, equation-internal symbols, ordinary symbol insertion, and native DOCX/XLSX/PPTX save/reopen/re-edit. The reference is Microsoft 365 for Mac desktop.

Three existing GPT-6 Sol / High implementers own Docs, Sheets, and Slides. Three existing GPT-6 Luna / Extra reviewers independently inspect them. Root owns shared conversion/catalog code, package integration, builds, and serial desktop verification. Four agents may run concurrently, including root; reviewers reuse their existing roles as slots become available.

## Delivered behavior

`@genoffice/office-math` provides the independent LaTeX/OMML/MathML converter, 11 structure groups and 42 selection-aware templates, and 774 named Unicode symbols in six subsets. The groups cover fraction, script, radical, integral, large operator, delimiter, function, accent, limit/log, operator, and matrix. Docs retains its prior math exports through a compatibility re-export, avoiding a Word/PowerPoint package cycle.

The shared workbench inserts templates or internal symbols at the source selection and offers professional MathML and linear source previews. The ordinary symbol picker provides font, subset, search, Unicode code point, and recent symbols. Control labels are integrated into all 20 application locales; canonical symbol search names remain English. Docs places Symbol immediately to the right of Equation in the Insert ribbon.

Equations are stored as native Office math. Sheets and Slides use DrawingML `a14:m` with native OMML and an AlternateContent fallback, following [Microsoft's DrawingML math specification](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/853b19c7-68a9-4f9a-a2ae-5e6cb0d02e62). The current native OMML is authoritative when reopening. Unsupported structures or properties remain preserved and read-only instead of being reduced to a partial editable source.

- Docs inserts inline/block equations, recovers supported source from native OMML, and retains cursor formatting when inserting a symbol with a different font.
- Sheets edits existing native equation drawings in place, preserving their identity, anchors, transforms, shape styling, descriptions, and adjacent text. Source metadata must match both the OMML digest and regenerated OMML. Picture or ambiguous text fallbacks remain read-only. Symbol insertion retains existing rich text and inherited styling across reopens.
- Slides updates a supported equation in its existing shape and retains identity, geometry, animation, and adjacent text. Selected equations reopen through the ribbon or canvas double-click. Unsupported math is preserved after unrelated edits.

## Desktop acceptance evidence

Each probe uses scratch documents and fresh Electron processes. It inserts a structured equation and symbols, saves, closes/relaunches, edits the reopened equation, saves again, closes/relaunches again, and confirms that source remains editable. Saved OOXML is inspected in addition to the visible editor.

| Editor | Desktop result                                                                                                             | Evidence and independent review                                                                                                                                                                                                                                      |
| ------ | -------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Docs   | 1/1 passed; two process reopens, equation re-edit, styled supplementary-plane symbol                                       | [Evidence](office-mac-math-evidence-2026-10-13/docs/2026-10-13T03-47-12-634Z-BUmx9A/evidence.json), [independent review](docs-math-verification-2026-10-13.md)                                                                                                       |
| Sheets | 2/2 passed on the final preservation build (22.5 s); two process reopens, double-click/F2 re-edit, inherited symbol styles | [Equation drawing after edit](office-mac-math-evidence-2026-10-13/sheets/sheets-math-edited-drawing.xml), [symbol summary](office-mac-math-evidence-2026-10-13/sheets/sheets-math-symbol-summary.json), [independent review](sheets-math-verification-2026-10-13.md) |
| Slides | 1/1 passed; two process reopens, actual canvas double-click, stable shape ID, unsupported import retained                  | [Evidence](office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/slides-math-verification.json), [independent review](slides-math-verification-2026-10-13.md)                                                                                   |

Root additionally opened a scratch copy of the final edited DOCX in native Word for Mac. Word opened it without a repair dialog, rendered the structured equation, and exposed its native Equation ribbon after selection. This checks recognition/display; native Word editing/re-saving and native Excel/PowerPoint math workflows were not tested.

## Regression checks

Accepted test/build logs are retained in [checks](office-mac-math-evidence-2026-10-13/checks/), including the [final Sheets desktop run](office-mac-math-evidence-2026-10-13/checks/genoffice-math-sheets-e2e-final4-20261013.log). Native Electron launch required execution outside the filesystem sandbox; the sandboxed attempt aborted before opening the UI. The authorized rerun uses scratch workbooks and passed without weakening assertions.

| Suite                                       | Result                   |
| ------------------------------------------- | ------------------------ |
| Shared authoring, Word math, shared math UI | 91 passed across 3 files |
| Full DOCX engine                            | 1,420 passed, 1 skipped  |
| Full PPTX engine                            | 1,065 passed, 1 skipped  |
| Full PPTX render                            | 279 passed               |
| Full PPTX operations                        | 36 passed                |
| XLSX gateway existing full suite            | 60 passed                |

Affected app/package typechecks and focused lint checks passed. The shared core also passes strict TypeScript with `noUncheckedIndexedAccess` and `exactOptionalPropertyTypes`. Docs, Sheets, Slides, PDF, and the unified shell were built successfully, including the final Sheets preservation change and its shell integration. Theme-color enforcement and `git diff --check` pass. Three existing Slides App hook warnings remain. A repository-wide format check includes unrelated existing warnings; it is not represented as a clean global formatting pass. The English-comments checker cannot run because its referenced local QA skill is absent.

Implementation details and focused test cases are recorded separately: [Docs](docs-math-implementation-2026-10-13.md), [Sheets](sheets-math-implementation-2026-10-13.md), [Slides](slides-math-implementation-2026-10-13.md).

## Remaining differences from Microsoft 365

- The 42 templates cover the 11 groups but not every Office variant. Authoring uses LaTeX and a shared modal; the full native contextual ribbon, UnicodeMath input, and all Office conversion modes are not implemented. Professional/linear switches change the preview.
- Imported OMML outside the conservative reversible subset remains read-only. Sheets imports with picture or ambiguous text fallbacks also remain read-only so their original content can be retained.
- Resolved in the [follow-up rendering wave](office-mac-math-followup-run-2026-10-13.md): Slides mixed/rotated equations, foreground stacking, PNG and PDF now use structured math inside the shared canvas renderer. PDF remains a raster-page export.
- Symbol font metadata persists, but an unavailable Mac font may render with substitution. Segoe UI Symbol was unavailable on this machine, so its exact glyph appearance was not verified.
- Resolved in the [follow-up formula-symbol wave](office-mac-math-followup-run-2026-10-13.md): Sheets inserts symbols into editable formula source and preserves references through two saves/reopens. The later [formula-cache correction](sheets-formula-cache-run-2026-10-13.md) also refreshes caches for unedited dependent formulas.

These results establish the implemented GenOffice authoring and persistence paths, not complete Microsoft 365 interface or rendering parity. Existing uncommitted work and user files were retained; no commit, push, or pull request was created.
