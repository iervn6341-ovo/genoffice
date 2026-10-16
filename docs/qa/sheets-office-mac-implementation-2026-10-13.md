# Sheets Office Mac implementation — 2026-10-13

Scope: `apps/sheets` and `packages/xlsx-gateway`. This is the first implementation wave from the 2026-10-12 Sheets details audit. The native Excel and GenOffice desktop UI verification is assigned to the supervising QA pass.

## Implemented

- **P1 Data Validation list mapping.** The XLSX writer recognizes Univer's UI-authored JSON string-array formula, emits an Excel inline-list formula (`["Open","Closed"]` → `"Open,Closed"`), and retains existing CSV literals and `=` range/formula sources. It rejects options that an Excel inline list cannot represent without loss: commas inside a label, quotes, newlines, empty items, and lists longer than 255 characters. This is an explicit save error that directs the user to a cell-range source. The read mapping now returns Univer's current JSON string-array representation, preserving whitespace in labels, and repairs the quoted JSON literal produced by the prior exporter when it parses unambiguously. Existing validation sqref, dropdown suppression, input prompts, and error alerts continue through the same writer.
- **UI list color metadata.** Univer's dropdown editor writes per-item colors into list `formula2` (`,` for two default-color options). XLSX list validation uses only `formula1` as its source, so the writer now omits list `formula2` while preserving `formula2` for non-list rules such as between bounds. This fixes the first real UI save failure that occurred after the JSON-list correction.
- **Header/footer first and even pages.** The dialog exposes odd, first, and even page sections with enable switches. The edit journal, save request schema, preload validation, XLSX gateway, and effective print setup carry the new variants. The gateway edits requested sections and flags while keeping untouched sections and attributes. Existing `&G` picture relationships remain in the workbook; the UI has no image-authoring control because the current save path does not create those media relationships.
- **Disabled variant retention.** The editor reads stored first/even section text independently from the print enable flags. After a variant is disabled, saved, and reopened, its text remains prefilled when the user re-enables it; PDF printing continues to use odd-page sections while the variant flag is off.
- **Fields and text format.** The dialog inserts page number, total pages, date, time, source workbook filename, source directory path, and sheet name. It also inserts font family, size, bold, italic, and underline codes. The read decoder preserves imported formatting codes for edit/save. PDF header/footer rendering applies common font size, family, color, bold, italic, underline, strike, super/subscript codes, resolves `&Z` from the source workbook path when available, and suppresses unsupported control codes instead of printing them as text. `&F` now receives the source workbook name while the PDF export request retains a `.pdf` output name.
- **Snapshot display name.** The sidecar reads a UUID-named temporary workbook snapshot. The main-process open response now replaces that internal name with the user's workbook basename (or the original CSV/recovery name, or the suggested XLSX name for legacy imports). The header/footer `&F` field therefore uses the visible workbook name in PDF export, and `&Z` uses the original CSV path for CSV sessions.

## Verification

- Sheets and xlsx-gateway TypeScript typechecks: passed.
- Focused Vitest suites: `xlsx-dv`, `univer-data-validation`, `xlsx-page-setup`, `print-settings`, and `workbook-display-name`: 5 files / 124 tests passed. Cases cover UI JSON export/import and color metadata, whitespace, prior defective JSON repair, unsupported inline labels, sqref, first/even gateway flags and preview state, font/path print rendering, and user-visible names across normal, recovery, CSV, and legacy imports.
- Focused ESLint and Prettier: passed.
- No dependency installation, native Office operation, Electron run, full build, or commit was performed in this implementation pass.

## Remaining tickets

1. **P1 QA acceptance:** GenOffice UI create/list save/reopen and Excel for Mac open/dropdown/valid-invalid input checks. Verify first/even header/footer UI save/reopen and PDF pages in the desktop app. The current unit results do not establish native Excel behavior.
2. **P1 complex dropdown labels:** create a managed helper range and formula reference for labels that cannot fit an Excel inline list. Until then, save rejects them explicitly.
3. **P1 header/footer pictures:** author new `&G` media and VML relationships, with image resizing and roundtrip, before exposing an image insert control.
4. **P2 text fidelity:** PDF font-family rendering currently accepts ASCII family names only, so imported Chinese and other non-ASCII font names do not preview in that family. Excel double underline `&E` currently prints as single underline; theme/tint color `&K` specifications and page-number offsets such as `&P+/-n` are not fully rendered. Verify imported first/even formatted headers in native Excel and PDF.
5. **P2 further Office parity:** Format Cells controls, Text to Columns wizard, Notes versus threaded Comments, equation data semantics, and unverified contextual tool details remain as prioritized in the 2026-10-12 audit.
