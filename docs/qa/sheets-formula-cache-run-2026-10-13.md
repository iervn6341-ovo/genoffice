# Sheets unedited-formula cache correction — 2026-10-13

Status: implemented and independently accepted in rebuilt Electron apps. Both full formula mode and active dependency-closure mode passed three saves and three full-process reopens, including immediate saves after precedent edits. The prior D1 stale-cache finding is resolved for these supported paths.

The existing GPT-6 Sol / High Sheets implementer and GPT-6 Luna / Extra reviewer were reused. Root integrated builds, serial Electron tests, raw-value checks, native XLSX inspection, and screenshot review. Existing uncommitted work and user files were retained; no commit, push, or pull request was created.

## Problem and resulting behavior

The pre-fix baseline4 app calculated all 10 imported formula results correctly after changing Sheet1 B1 and Inputs A1, but every formula's native cached `<v>` remained at its original value after save. The save collected authored/verified formulas and fallback overlays, omitting unedited imported formulas in the live engine. `fullCalcOnLoad` did not refresh the file itself.

Normal automatic-calculation saves now force a calculation pass and wait until results are applied. They harvest indexed, resident imported formula results when the live source still matches the saved formula. The gateway preserves `<f>` and writes fresh values with the correct numeric, text, boolean, or error type. Content-edited formulas retain their existing save path; style-only edits remain eligible. A calculation timeout rejects normal Save/Save As before writing the file and retains pending edits, with a retry message in all 20 locales.

Cross-sheet error display follows direct references, ranges, and indexed formula chains to identify edited precedents. Genuine errors caused by changed inputs replace old numeric caches; unrelated cross-sheet edits retain existing capability fallbacks. Iterative traversal, point lookup, cycle tracking, and work budgets avoid unbounded painting work.

Actual UI testing additionally exposed exact-case sheet lookup in Univer. A narrow per-runtime/per-workbook compatibility wrapper now resolves names case-insensitively at its two lookup boundaries while retaining exact-match priority. Neither original formula text nor enumerable name-map contents is changed. The real lowercase `inputs!A1` reference now works against the `Inputs` tab through all saves/reopens.

## Desktop acceptance

[Verification spec](../../e2e/sheets-formula-cache-verification.spec.ts) imports a native cached workbook and only changes precedent constants. Its 11 formula cells cover direct SUM, text concatenation, numeric chains, boolean results, transitions into and out of `#DIV/0!`, mixed-case cross-sheet references, absolute/range references, a cross-sheet chain, and an offscreen formula. Constants and formula sources are also checked.

Each run:

1. Opens the imported baseline; changes precedents on both sheets; confirms live results and saves.
2. Fully closes/relaunches; changes both precedents; saves immediately without a harness calculation wait.
3. Fully closes/relaunches; changes only Inputs A1; saves immediately, including the dependent K1 error and cross-sheet chains.
4. Fully closes/relaunches; checks every formula source and result again.

Every saved XLSX is checked directly for `<f>`, `<v>`, cell type, unchanged constants, edited precedent values, and ZIP integrity. Root and Luna reviewed native files and screenshots independently.

| Final case | Result | Evidence |
| --- | --- | --- |
| Active closure; M5000, 65k-cell used rectangle | 1/1 passed, 32.6 s | [Summary](sheets-formula-cache-evidence-2026-10-13/closure-case-final/summary.json), [third saved XLSX](sheets-formula-cache-evidence-2026-10-13/closure-case-final/third-save.xlsx), [final reopen](sheets-formula-cache-evidence-2026-10-13/closure-case-final/final.png) |
| Full formula mode; M2000, 26k-cell used rectangle | 1/1 passed, 26.0 s | [Summary](sheets-formula-cache-evidence-2026-10-13/preloaded-case-final/summary.json), [third saved XLSX](sheets-formula-cache-evidence-2026-10-13/preloaded-case-final/third-save.xlsx), [final reopen](sheets-formula-cache-evidence-2026-10-13/preloaded-case-final/final.png) |

All 11 formulas retain their original source at each save. D1's caches change 5 → 8 → 10 → 10; its text remains `SUM(B1:B2)`. K1 remains exactly `1/(inputs!A1-40)` and its saved values/types change from numeric to error, back to numeric, and back to error. The remote formula changes 15 → 24 → 30 → 30. Inputs B1's cross-sheet chain changes 55 → 100 → 120 → 110. Final visible results and all native caches agree.

Earlier exact-case-only probes also passed in both modes (32.5 s and 25.2 s). The first mixed-case built-app attempt failed with an initial engine `#VALUE!`, and prompted the lookup correction. These intermediate runs are not represented as acceptance of the final source. Earlier harness failures involved a wrong facade name and reading formatted rather than raw values; baseline3 was interrupted. Baseline4 is the actual stale-cache reproduction. A loader-test mock initially omitted the real `getWorkbook()` facade API; it was updated without weakening the existing view/freeze assertions.

## Validation and records

Accepted/check logs and the meaningful pre-fix failures are retained in [checks](sheets-formula-cache-evidence-2026-10-13/checks/). Native Electron runs required approved execution outside the filesystem sandbox and used scratch documents/profiles.

- Final related source tests: 9 files / 80 passed.
- Loader and lookup boundary regressions: 2 files / 30 passed; the lookup tests overlap the preceding suite and are not added to its count.
- Independent reviewer focused tests: 6 files / 70 passed; overlapping tests are not added to the root count.
- Sheets typecheck and final Sheets/unified Shell builds passed.
- Targeted ESLint and Prettier checks passed. Theme-color enforcement and diff whitespace checks passed.
- Prior formula-symbol built-app regression: 1/1 passed in 11.8 s, including live selection, ordinary text, two saves and two process reopens. D1 now retains its formula and saves the refreshed cache 8 instead of the old 5.

[Implementation](sheets-formula-cache-implementation-2026-10-13.md) and [independent final verification](sheets-formula-cache-verification-2026-10-13.md) describe the individual source guards and tests.

## Practical limits

This acceptance covers automatic calculation with unchanged coordinates in the existing supported full/closure modes. Full mode and an active closure each have a 50,000-cell engine budget, matching the result-harvest limit. Oversized/unavailable closures, unsupported formulas, structural-coordinate changes, manual calculation, and best-effort crash-recovery snapshots retain existing separate behavior. Native Excel was not tested.

The fallback dependency parser follows explicit A1/range references and indexed formula chains; this is not a complete dynamic/defined-name/structured-reference dependency analyzer. It visits at most 1,000 formulas or scans 10,000 range entries. If content edits exist and that bound is exceeded, it conservatively displays the engine result, which can expose an unrelated capability error. Unsupported `#NAME?` and masked engine failures remain cache-preserved. These limits are not represented as complete Microsoft 365 calculation parity.
