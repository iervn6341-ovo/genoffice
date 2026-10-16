# Equation rendering and formula-symbol follow-up — 2026-10-13

Status: requested fixes implemented and accepted in rebuilt Electron apps. Independent source, test, and saved-artifact reviews completed. The separate formula-cache finding below was subsequently resolved in the [formula-cache correction](sheets-formula-cache-run-2026-10-13.md).

## Scope

This follows the accepted [equation/symbol authoring wave](office-mac-math-run-2026-10-13.md). The user requests structured Slides equations in rotated shapes, mixed text, and PDF exports, plus Sheets symbol insertion that retains formula source instead of replacing a calculated value.

The existing Slides and Sheets GPT-6 Sol / High implementers and GPT-6 Luna / Extra reviewers are reused. Root supervises package integration, builds, serial Electron UI runs, and final artifact inspection. Existing dirty work and user files remain in place.

## Delivered behavior

Slides derives an atomic math run from native OMML, reserves inline layout bounds, and paints a transparent MathML raster inside the shared Konva node. That path is shared by editor, thumbnails, PNGs, and the existing rasterized PDF pipeline. Source OMML remains native and is not rewritten by visual rendering. Rotation and stacking use the shape's normal transforms and draw order.

Sheets captures the actual formula source and live caret/selection before opening Symbol. A formula source field receives the chosen symbols, and explicit Apply writes the formula to the original captured cell and clears stale calculated/rich/shared-formula data; Cancel retains its saved formula. An unsaved formula-editor draft is discarded on Cancel, like Esc. Ordinary active text commits before symbol insertion. The captured workbook/sheet/cell prevents Univer's Enter movement to the next row from redirecting insertion. Folded Symbols launchers preserve editor focus only where enabled.

## Built-app acceptance

Each persistence scenario saves a scratch Office file and fully closes/relaunches Electron twice. Root and the independent reviewers inspected the saved artifacts as well as source tests.

| Scenario | Accepted result | Evidence |
| --- | --- | --- |
| Slides rotated/mixed equations and exports | Case passed in 5.3 s; native save, reopen, canvas double-click/edit, second save/reopen, actual GUI PNG/PDF export | [Evidence JSON](office-mac-math-followup-evidence-2026-10-13/slides/2026-10-13T04-59-58-150Z-3eGom3/evidence.json), [PDF page image](office-mac-math-followup-evidence-2026-10-13/slides/2026-10-13T04-59-58-150Z-3eGom3/pdf-render.png), [review](slides-math-render-verification-2026-10-13.md) |
| Sheets formula symbols | Final10: 1/1 passed in 12.1 s; live selection capture, Cancel, plain-text commit, two saves/reopens, edited formula recalculation and renewed editability | [Summary](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-summary.json), [third-open screenshot](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-third-reopen.png), [review](sheets-formula-symbol-verification-2026-10-13.md) |
| Slides prior native math/symbol authoring | Case passed in 8.4 s after overlay removal; supported math editing, shape identity, unsupported import preservation | [Regression log](office-mac-math-followup-evidence-2026-10-13/checks/genoffice-math-followup-e2e-final3-20261013.log) |
| Sheets ordinary rich symbols | Case passed in 8.2 s; inherited runs, fonts/styles, save/reopen preserved | [Regression log](office-mac-math-followup-evidence-2026-10-13/checks/genoffice-math-followup-e2e-final3-20261013.log) |

The Slides final6 combined log contains the accepted Slides case and an earlier failed Sheets case; it is not an all-pass run. The final3 combined regression log contains two accepted regression cases and two earlier failing cases. Sheets final10 is its final standalone all-pass acceptance. Earlier failures and traces were retained. Electron required approved execution outside the filesystem sandbox, using scratch profiles and documents.

Root and the Slides reviewer visually inspected the final PNG and Poppler-rendered PDF: the 25° rotated nested fraction/root/script, styled normal text on both sides, compact parallel fractions, and foreground shape overlap are correct in this fixture. The foreground pixel assertion is `[211, 47, 47]`. The PPTX retains shape 40's rotation and byte-identical native equation payload. Edited shape 41 retains its numeric identity and surrounding text while its equation changes from `z` to `q`; the save adds a persisted creation ID.

Sheets first saves `="Ω "&SUM(B1:B2)` with display `Ω 5`. After reopening and changing B1 from 2 to 5, the edited formula displays `Ω 8`. Inserting γ then saves `="Ωγ "&SUM(B1:B2)` with display `Ωγ 8`; a second full-process reopen retains source and display and reopens the editable source dialog. Ordinary live text becomes `DraftΩ` in C1; C2 remains empty. Both XLSX ZIPs are valid and retain native formula elements.

## Source/build checks

Accepted logs are preserved in [checks](office-mac-math-followup-evidence-2026-10-13/checks/).

- Final full PPTX-render suite: 14 files / 288 passed.
- Current Sheets/shared-UI focused tests: 5 files / 27 passed, covering formula source/selection, target identity, rich text, and folded-launcher focus. The independent reviewer separately passed 4 files / 23 tests; overlapping cases are not added to the root count.
- Independent Slides focused tests: 3 files / 14 passed, including equation-array height, structural width, atomic RTL layout, raster preparation and failure propagation.
- Full Slides suite before the final layout refinements: 91 files / 927 passed / 1 skipped. It is not represented as a full-suite run after those refinements.
- Final Sheets and Slides typechecks passed. Final Sheets, Slides, and unified Shell builds passed. Targeted formatting, theme-color enforcement, and diff whitespace checks passed. Existing unrelated formatting/hook warnings are not represented as resolved.

Implementation notes: [Slides](slides-math-render-implementation-2026-10-13.md), [Sheets](sheets-formula-symbol-implementation-2026-10-13.md). Existing uncommitted work and user files were retained; no commit, push, or pull request was created.

## Limits and separate remaining finding

- PDF retains the existing rasterized-page format; equations are visually structured but are not searchable/vector math. Chromium measures the rendered math, while native layout reserves structural estimates. Complex imports, unusual fonts/WordArt styling, and approximate mixed-text baselines may differ from PowerPoint.
- GUI acceptance covers this four-shape fixture. Group/table/vertical/RTL and imported equation-array paths have source-level coverage, not exhaustive native Office visual parity. Unsupported imported OMML remains preserved and read-only as in the prior wave.
- Sheets validates formula length, leading `=`, and balanced quotes/parentheses; this is not full formula parsing. Formula characters do not retain the Symbol picker's font choice.
- The live formula selection was established through the public `FDocument.setSelection` API in an actual F2 editing session. Capture was checked before opening Symbol. Native keyboard selection and native Excel/PowerPoint editing were not verified.
- **Historical finding, now resolved:** in this wave's saved artifact, untouched D1 retains native `SUM(B1:B2)` but its cached result remains 5 in the second saved XLSX after B1 becomes 5 and B2 remains 3. The edited A1 correctly recalculates and persists `Ωγ 8`. The file requests recalculation via `fullCalcOnLoad="1"`; native Excel behavior was not observed. This acceptance establishes formula-source preservation/editability and the edited formula's recalculation, not correctness of all unrelated formula caches. The later [formula-cache correction](sheets-formula-cache-run-2026-10-13.md) refreshes these imported formula caches and verifies both supported calculation modes; the earlier artifact is retained unchanged as historical evidence.
