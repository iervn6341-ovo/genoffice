# Slides math render verification — 2026-10-13

## Independent source review — 2026-10-13

This is an independent source, focused-test, and saved-artifact review of mixed-text, transformed, and exported equations. No product source was changed and I did not launch Electron or run a build; the coordinating agent ran the UI scenario and provided its saved artifacts for review.

The implementation now carries each native math run as an atomic `GlyphRun` with its MathML, estimated dimensions, and paragraph/run identity. `NodeBody` draws a cached SVG raster through Konva at the run's position, so normal shape stacking and shape/group transforms apply. `renderSlidesToPngBase64` waits for raster images in shapes, groups, and table cells before canvas capture. This closes the previous App-overlay/export mismatch at the source level. The authoritative `TextRun.rawXml` remains the save representation.

## Constraints for the structured renderer

- The render tree now uses one atomic math run in ordinary mixed text, including RTL reordering; it tracks paragraph plus run identity. Shape, group, and table-cell render paths all reach the same in-node raster renderer.
- Equation dimensions remain heuristic in the main-process layout. The latest source adds `mathWidthPx()`, which walks the MathML tree and estimates widths for fractions, radicals, scripts, and table columns rather than measuring the linear fallback string. `mathHeightEm()` now accumulates rows for both matrix (`m:m`) and equation-array (`m:eqArr`) structures. Chromium measures actual MathML for the SVG, then `equationSvg()` scales content down when the measured DOM exceeds the reserved width or height. The structural estimates reduce the known over-wide slots and account for multiline equation arrays, but they do not guarantee native PowerPoint metrics. Mixed-text baseline alignment also remains an approximation: the equation is centered in the raster box while layout reserves a fixed 78/22 ascent/descent split.
- The shared Konva path applies normal shape stacking and transforms, and export awaits cached raster images before capture. The rebuilt UI/PNG/PDF run is recorded below. The PDF contains raster page images, so this review does not establish searchable or selectable math text.

## Validation and accepted artifact review

`npx vitest run packages/pptx-render/tests/equation-preview.test.ts packages/pptx-render/tests/equation-inline-review.test.ts apps/slides/tests/equation-raster.test.ts` passed: 3 files, 14 tests. Coverage includes structurally sized nested/parallel fractions, equation-array row height, RTL math atomicity, paragraph/run identity, SVG construction, group/table raster preload, and export failure propagation. These tests do not measure actual rendered pixels or baseline alignment.

The rebuilt Electron scenario passed **1/1** in the coordinating agent's run. It opened the deck, saved it, closed and reopened the app, edited the mixed equation, saved again, closed and reopened a second time, re-opened the equation for verification, then exported PNG and PDF. I inspected the saved evidence independently.

The final editor/export PNG and the Poppler-rendered PDF page visibly contain the 25° rotated nested fraction/root/script, the mixed prefix/equation/suffix runs, the parallel fractions, and the red foreground rectangle covering part of the equation. Equation spacing is compact in the updated render; the mixed runs remain legible. The foreground pixel check in the run records `[211, 47, 47]` at the rectangle center. The PDF is one raster page.

The PPTX packages confirm that shape `cNvPr id="40"` retains `rot="1500000"` (25°), and its complete `mc:AlternateContent` equation payload is byte-identical between the initial-saved and edited-saved decks. Shape `cNvPr id="41"` remains present after the edit; its equation changes from `z` to `q`, while the `Before ` and ` after` fallback text is retained. The edit-save adds an `a16:creationId` to shape 41; the cNvPr id stays 41 and the new creation ID is persisted for later opens. The test/evidence reports four math-bearing shapes through the final reopen.

Evidence files:

- [Editor/export PNG](office-mac-math-followup-evidence-2026-10-13/slides/2026-10-13T04-59-58-150Z-3eGom3/rotated-mixed-equations-01.png)
- [PDF](office-mac-math-followup-evidence-2026-10-13/slides/2026-10-13T04-59-58-150Z-3eGom3/structured-equations.pdf) and [Poppler-rendered page](office-mac-math-followup-evidence-2026-10-13/slides/2026-10-13T04-59-58-150Z-3eGom3/pdf-render.png)
- [Evidence JSON](office-mac-math-followup-evidence-2026-10-13/slides/2026-10-13T04-59-58-150Z-3eGom3/evidence.json), [initial-saved PPTX](office-mac-math-followup-evidence-2026-10-13/slides/2026-10-13T04-59-58-150Z-3eGom3/initial-saved.pptx), and [edited-saved PPTX](office-mac-math-followup-evidence-2026-10-13/slides/2026-10-13T04-59-58-150Z-3eGom3/edited-saved.pptx)
- [UI verification spec](../../e2e/slides-math-render-verification.spec.ts)

This acceptance covers the four shapes in this fixture and this Chromium/MathML rendering path. It does not establish PowerPoint parity across imported equations, all fonts, vertical text, groups, tables, or other math structures; the focused equation-array test is source-level only. The PDF is a raster export and is not searchable/selectable math text. No native PowerPoint app was used for this review.
