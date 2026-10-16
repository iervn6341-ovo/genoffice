# Slides math and symbol verification — 2026-10-13

## Result

The final built-app Electron probe passed 1/1 tests in 8.6 seconds. It authored a nested equation from the Insert ribbon, added an internal Greek symbol, saved a native PPTX, restarted and reopened it, edited the equation both through selection and canvas double-click, saved and restarted again, and confirmed the edited equation and its shape ID survived. It also inserted a standalone `∑` with `Segoe UI Symbol` and confirmed the font persisted after reopening. An unsupported imported OMML equation remained read-only and its math block was preserved byte-for-byte after an unrelated symbol edit and save.

I independently inspected the run JSON, screenshots, and both saved PPTX archives. Both ZIP packages pass integrity checks. The edited equation package contains an `a14:m` PowerPoint math branch with OMML plus a linear fallback; the unsupported-import package retains its `m:phant` structure. The equation remains shape ID `4` through the second save/reopen. The symbol shape contains `∑` and the `Segoe UI Symbol` typeface in its saved DrawingML runs.

The run artifacts are in [the final evidence directory](office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/): [run JSON](office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/slides-math-verification.json), [edited PPTX](office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/math-symbol-roundtrip.pptx), [unsupported-import PPTX](office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/unsupported-import.pptx), [equation on canvas](office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/equation-canvas.png), [before double-click](office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/before-double-click.png), and [equation editor reopened by double-click](office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/after-double-click.png).

## Source review

- `packages/pptx-ops/src/ops/equation-ops.ts` writes one paragraph-level `mc:AlternateContent` run. The PowerPoint branch carries `<a14:m>` with `<m:oMathPara>` and OMML; the fallback contains linear text. Replacement requires exactly one equation and successful strict OMML-to-LaTeX conversion, then updates the existing math run.
- `packages/pptx-engine/src/parse.ts` models paragraph-level math alternate content as a raw `TextRun`, taking display text from its fallback. It also normalizes direct `<a14:m>` blocks without the alternate-content wrapper. `packages/pptx-engine/src/generate.ts` recognizes both wrapper and direct math blocks as raw run spans. The UI probe exercises wrapped output; it does not include a separate direct-`a14:m` import fixture.
- `packages/pptx-render/src/build-slide.ts` derives MathML and editable source from OMML. `apps/slides/src/renderer/App.tsx` overlays actual MathML for standalone, unrotated equations and uses the underlying canvas shape for selection and hit testing. Mixed-text and rotated equations keep the linear canvas display described in the implementation note.
- `apps/slides/src/renderer/insert-actions.ts` uses the equation transaction for insert/replace. Symbol insertion restores a saved text selection when one exists, restricts inline font names to the picker allow-list, and otherwise creates an ordinary text box with the chosen font family.
- The strict reverse converter refuses unsupported OMML such as `m:phant` for editing. The operation tests cover refusal without changing imported raw XML; the UI probe confirms the unsupported block remains byte-for-byte intact after a separate edit and save.

## Validation

- Focused authoring, equation operation, preview, and symbol-action tests passed: 4 files, 57 tests.
- The later equation-hit and operation regression tests passed: 7 tests total. TypeScript and lint checks reported no errors in the parent run.
- Locale-key audit checked all 20 `apps/slides/src/renderer/i18n/ribbon/*.ts` shards; all 29 new math and symbol keys were present.
- The final opt-in Electron UI run passed 1/1: `GENOFFICE_SLIDES_MATH_VERIFICATION=1 npx playwright test --config=e2e/playwright.config.ts e2e/slides-math-verification.spec.ts`.
- Independent archive inspection found valid ZIP integrity for both saved PPTX files and verified native OMML, linear fallback, unsupported `m:phant`, stable shape ID, and the saved symbol font.

## Limits

The browser's MathML preview is a DOM overlay positioned over the Konva stage. Its stacking is not derived from the presentation's full shape draw order, so overlaps between a math preview and other slide shapes remain a visual limitation to check. The final probe confirms the preview's placement and editing behavior for its authored standalone equation; it does not establish all overlap or stacking cases.

Slide-image and PDF export still use the linear equation fallback because the HTML MathML preview belongs to the editing canvas. Mixed-text and rotated equation shapes also use the linear canvas fallback. Imported OMML outside the strict reverse converter remains preserved and visible but is read-only in this editor. The probe ran in the built Electron application and inspected generated PPTX packages; it did not open the file in native PowerPoint for Mac. No claim is made about exact visual parity with PowerPoint's native renderer or full coverage of its Equation Symbols gallery.
