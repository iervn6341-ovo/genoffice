# Sheets formula-symbol verification — 2026-10-13

## Scope and status

This independent review covers the Insert → Symbol formula path, focused unit coverage, and root's accepted Electron E2E. It did not run the build or launch Electron itself. The accepted UI test passed; no native Excel parity or native keyboard-selection parity is claimed.

## Findings

- Formula capture records the active workbook, sheet, row, and column before the ribbon click can move focus. For an active formula editor it reads the live document text and Docs selection range; otherwise it tries the facade formula, harvested lazy formula index, then raw cell formula. A live non-formula editor takes precedence over an older formula in that cell and falls through to the commit-before-picker path.
- The installed Docs selection service sets its current document selection to `{ unitId, subUnitId: unitId }` when the normal editor becomes current, and `getTextRanges()` defaults to that current selection. Capture prefers that current selection when it belongs to the active editor, then falls back to editor-owned and default subunits. Independent tests now model this lookup rather than returning an empty selection for every query.
- Formula Apply uses the captured workbook/sheet/cell identity and refuses a different active workbook, missing captured sheet, source without `=`, source longer than 8,192 characters, and unmatched parentheses or quotes. Its `setValue` mutation reaches the existing range-values journal listener, where formula text is recorded for save. Cancel closes the dialog without calling Apply.
- The normal text-edit path waits for `endEditingAsync(true)` before opening the ordinary symbol picker, avoiding a stale read of the cell while that edit is still open. Formula editing is captured and then aborted with `endEditingAsync(false)` before showing the formula-source dialog.
- **Validation is now guarded:** the symbol path rejects non-`=` sources, overlong sources, unmatched parentheses, and unmatched double quotes, while ignoring parentheses inside quoted strings and escaped doubled quotes. The writer checks the same helper independently before writing. This closes the earlier inconsistency with `handleApplyFormula`. The helper is still a lightweight delimiter check rather than full formula parsing.
- **Editor-close handling is now awaited:** formula drafts open only after `endEditingAsync(false)` succeeds; ordinary in-progress text edits open the picker only after `endEditingAsync(true)` succeeds. False results and rejected promises leave the picker closed and show the existing symbol-insert failure message.
- **Cancel semantics were observed:** a saved numeric-result formula in D1 remains unchanged after opening Symbol and pressing Cancel. A live unsaved `+1` formula draft in A1 is shown in the modal, then discarded on Cancel while the previously saved formula remains. A plain-text C1 draft commits before the symbol is appended.

## Accepted Electron E2E

Root's final10 run passed the opt-in [formula-symbol E2E](../../e2e/sheets-formula-symbol-verification.spec.ts): 1 test passed in 12.1 seconds (11.6 seconds for the case). The probe established an actual F2 cell-edit session, used the public `FDocument.setSelection(2, 7)` API to select the formula literal, and checked the read-only live-draft hook returned the same range before opening Symbol. After inserting Ω, the formula became `="Ω "&SUM(B1:B2)` and saved as an XLSX formula with the cached display `Ω 5`.

After the first reopen, the probe changed B1 from 2 to 5 and observed `Ω 8`, reopened the formula-symbol dialog, inserted γ, and saved again. The second saved workbook contains `="Ωγ "&SUM(B1:B2)` with cached result `Ωγ 8`; D1 retains `SUM(B1:B2)` with cached result 5, C1 contains `DraftΩ`, and C2 is absent. A third reopen restored the final formula and display and reopened the formula-source dialog with the saved expression.

The caret-selection range in this E2E is set through Univer's public selection API because formula arrow keys can enter reference selection; it verifies the product's captured selection path but does not establish native keyboard-selection parity.

Durable evidence: [run log](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-final10.log), [summary JSON](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-summary.json), [first saved workbook](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-first-save.xlsx), [second saved workbook](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-second-save.xlsx), [first-apply screenshot](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-first-apply.png), [second-apply screenshot](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-second-apply.png), and [third-reopen screenshot](office-mac-math-followup-evidence-2026-10-13/sheets/formula-symbol-third-reopen.png).

## Independent tests

Added [symbol-formula-verification.test.ts](../../apps/sheets/tests/symbol-formula-verification.test.ts), covering harvested formula-index fallback, rejection of stale indexed formulas after structural or cell edits, fixed target identity after the active range moves, UTF-16 offsets for supplementary-plane symbols, live formula source precedence, editor-location mismatch behavior, and Docs selection lookup with both own and default subunits.

Command run with the already-installed local Vitest binary:

```text
node_modules/.bin/vitest run apps/sheets/tests/symbol-formula.test.ts apps/sheets/tests/symbol-formula-verification.test.ts apps/sheets/tests/symbol-dialog.test.ts packages/ui/tests/ribbon-fold-group.test.ts
```

Result: 4 test files passed, 23 tests passed. Root's current broader focused run also reported 5 test files and 27 tests passed. An earlier run exposed an exact-object matcher in the implementation test that did not account for Univer normalizing `setValue({ f })` to include `p: null`, `si: null`, and `v: null`; Sol corrected that assertion/normalization, after which the focused pair passed.
