# Docs equation and symbol implementation — 2026-10-13

## Delivered

- The Insert Equation dialog uses the shared math workbench with 11 structure categories and 42 selection-aware templates. Its Symbols category has searchable mathematical, Greek, and arrow galleries; clicking a glyph inserts the actual Unicode character at the LaTeX cursor. The professional preview renders converter-produced MathML, and the linear preview shows the editable source.
- The Insert ribbon still offers the eight existing complete equation presets. Its new-equation action opens the workbench. If a reversible inline or display equation is selected, the Equation ribbon button opens the same workbench in re-edit mode. Existing inline double-click and display edit controls also reopen it.
- Re-edit starts from the equation's current OMML conversion, rather than trusting stored LaTeX metadata. Unsupported imported OMML has no full-re-edit action and remains unchanged on save. A supported edit generates updated OMML and persists through save, reopen, and further editing.
- The advanced Symbols dialog supports a font choice, six subsets, search by Unicode name or code point, direct code point entry, and recent symbols. It inserts the chosen scalar at the captured document selection with its selected font. The existing quick palette remains available.
- All new Docs interface strings are present in the 20 app locale shards. Styles use semantic theme variables.

## Source and checks

Shared workbench and symbol picker: `packages/ui/src/MathWorkbench.tsx`, `SymbolPicker.tsx`, and `math-tools.css`. Docs integration: `apps/docs/src/renderer/components/EquationModal.tsx`, `SymbolDialog.tsx`, `ribbon-insert-tab.tsx`, `main.tsx`, `styles.css`, and the app locale shards. Focused tests: `packages/ui/tests/math-workbench.test.ts` and `apps/docs/tests/equation-insert.test.ts`.

- UI catalog test: all 42 templates render MathML, reverse to editable source, and preview again; all mathematical, Greek, and arrow gallery characters render after insertion. Passed (3 tests).
- Docs insertion test: display and inline equations save and reopen for re-edit; selected Equation ribbon target resolves to the existing node; supplementary-plane symbols insert with the chosen font while keeping inherited bold; unsupported imported OMML resists full re-edit and is preserved. Passed (8 tests).
- `tsc --noEmit -p packages/ui/tsconfig.json` and `tsc --noEmit -p apps/docs/tsconfig.json`: passed.
- Focused ESLint and Prettier checks for changed math files and all 20 locale shards: passed.

The unit runner prints jsdom's existing canvas `getContext` notice during the Docs tests; it does not fail the suite. The narrow Electron screenshot revealed that the generic modal input rule enlarged the inline-equation checkbox and pushed its label to the next line. A scoped CSS rule now gives it a native 16 px box beside the label; root QA will verify the rebuilt UI. Native Office and remaining Electron visual checks are assigned to root QA.

## Remaining priority tickets

1. Verify the compact ribbon, modal, template gallery, and font rendering visually in Electron at narrow and wide widths, including dark mode and RTL locales.
2. Extend the first wave's 42 structures toward the larger Word template gallery only after the supported converter subset and accessibility behavior are confirmed.
3. Preserve richer unsupported OMML when providing future full structure editing. Current full re-edit intentionally stays unavailable for unsupported originals; leaf-token editing follows the pre-existing protected formula path.
4. Search uses English Unicode canonical names across all locales. Localized symbol names would require a translated data catalog, separate from the existing 20 interface translations.
