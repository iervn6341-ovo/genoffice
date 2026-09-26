# GenOffice E2E Testing — How to Run It

Use this when the task is to **test** a GenOffice service end to end, not just
to change code.

Status rules, severity, lifecycle and the final gate live in
`../references/completion-gate.md`. Use `../templates/test-report.md` for the
report. This folder adds the per-service checklists and the GenOffice-specific
mechanics that the generic guides lack.

| Service | Checklist |
|---|---|
| Word-like (Docs) | `word.md` |
| Excel-like (Sheets) | `excel.md` |
| PowerPoint-like (Slides) | `powerpoint.md` |
| PDF editor, Markdown, HTML | `pdf-markdown-html.md` |

Each checklist has two parts. The first is the service test suite: which
features to cover. The second is the "GenOffice notes": known traps,
Office-parity details already probed, and where state can be read.

---

## 1. Run order

1. **Classify the run.**
   - A *smoke run* covers the service suite once.
   - A *regression run* covers the checklist plus existing specs.
   - A *bug hunt* explores the checklist interactively, then writes a spec
     for every defect it finds.
2. **Build first.** Specs run against the built `out/` folders, so rebuild
   before testing. Never build while a suite is running.

   ```bash
   npm run build -w @genoffice/<app> && npm run build -w @genoffice/shell
   ```

   Or rebuild everything:

   ```bash
   npm run build:all
   ```

   Main-process code (`apps/*/src/main`) is compiled into the **shell**. If
   you change it and don't rebuild the shell, nothing changes.
3. **Run the existing specs for the service**, as the baseline.

   ```bash
   GENOFFICE_E2E_NO_VIDEO=1 npx playwright test --config e2e/playwright.config.ts --reporter=line e2e/<service>-*.spec.ts
   ```

   The full suite takes about 18 minutes. The last known baseline is 166
   passed / 6 skipped.
4. **Walk the checklist.** For each item: if an existing spec covers it, cite
   the spec; otherwise test it interactively or write a new spec.
5. **Compare with Microsoft Office** (Word / Excel / PowerPoint are
   installed on the dev Mac) for any behavior you are about to call a defect.
   Record whether the Office behavior was *directly observed* or *inferred*.
6. **For every defect you fix:** write a failing spec, fix the code, make the
   spec pass, then run the neighbouring specs.

## 2. Harness facts

- **Launching the app.** `e2e/helpers.ts` → `launchShell(options)`:

  | Option | Effect |
  |---|---|
  | `lang` | UI language, default `'en'`. Use `'zh-TW'` for CJK behavior. |
  | `openFile` | Opens a file at launch. |
  | `onboardingSeen: true` | Skips onboarding. |
  | `settings` | Preset settings. |
  | `foreground` | Brings the window forward (see below). |
  | `env`, `userDataDir` | Extra environment variables / data folder. |

  - Test windows are 1280×800.
  - To open an app, call `launched.page.locator('.quick-card').first().click()`
    to create a new document, then `waitForPageWithUrl(app, '://docs/')`
    (or `://sheets/`, `://slides/`, `://pdf/`, …).
- **Background mode** is the default (`GENOFFICE_E2E_BACKGROUND=1`): windows
  never take focus from the user. A spec that needs real OS keyboard focus,
  such as a slide show, passes `foreground: true`. To force foreground for a
  whole run, set `GENOFFICE_E2E_FOREGROUND=1`.
- **State hooks.** Assert on model state, not only on pixels:

  | Service | Hook |
  |---|---|
  | Docs | `window.__aidocs.editor` (TipTap/ProseMirror): `state.doc`, `state.selection`, `storage.*`. |
  | Sheets | `window.__genofficeDebug.univerAPI` (Univer facade). Only present when `GENOFFICE_DEBUG_HOOKS=1` (set by the helpers). |
  | Slides / PDF / Markdown / HTML | DOM plus the saved file. |

- **Save dialogs.** Stub them from the main process:

  ```js
  app.evaluate(({dialog}, f) => {
    dialog.showSaveDialog = async () => ({canceled: false, filePath: f})
  }, out)
  ```

  Then check the file with `unzip -p file word/document.xml` (or `xl/…`,
  `ppt/…`).
- **Save As.** A new Docs document saves silently to the default folder on
  the first ⌘S; use ⌘⇧S (handled in the renderer). In Sheets, ⌘⇧S is a native
  menu accelerator, so click `button[aria-label^="Save As"]` instead.

## 3. Known false failures

Check these before filing a bug.

| Symptom | Real cause | What to do |
|---|---|---|
| A key after a click, Home or End acts on the old caret | A synthetic key burst outran `selectionchange`; a person can't type that fast | `expect.poll` on `editor.state.selection` before the next key |
| ⌘⇧Z, ⌘⇧S or another menu shortcut does nothing in Playwright | It is a native menu accelerator, which Playwright keyboard events never reach | Click the equivalent toolbar/QAT button (same code path) or call the menu item |
| Focus/typing specs fail only in some runs | The screen is locked, or the user is active in another app (macOS won't let a background app take key focus) | Check the lock state (`CGSessionCopyCurrentDictionary`) and the frontmost app, then rerun while idle |
| One unit test times out only in the full `npm test` | Load (e.g. the password-hash dialog test) | Rerun the file alone; report it only if it fails alone |
| `apps/sheets/tests/xlsx-sidecar-cancel.test.ts` fails | Node 26 locally; CI uses Node 22 | Pre-existing, ignore |
| html2docx tests fail | They need Chromium | Set `CHROME_PATH="/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"` |

## 4. Mac and Office parity rules already verified

Treat these as the expected behavior.

- **Home / End.** They move to the start / end of the line on Mac.
  Chromium's default is to scroll, and that is a bug in GenOffice.
- **Redo.** ⌘⇧Z in GenOffice (Word and PowerPoint for Mac use it too).
  Check Excel's own binding in Excel before asserting it.
- **CJK documents.** Probed in zh-TW Word:
  - The default tab stop is 480 twips (24pt, i.e. 2 characters at 12pt).
  - Indents are shown in 字元 (characters). The customary first-line indent
    is 2 字元.
  - English documents use 0.5" (720 twips).

## 5. Test data

- **Test PDF:** `1_Stock_Funds_Rose_______in_____.pdf` at the repo root. It
  is intentionally not committed.
- **Generated fixtures.** Specs build their own files in `mkdtemp`. Prefer
  that over committing binaries.
- **Office-authored files.** When a behavior depends on how Office writes a
  file (notes, CF rules, sections), create the file in the real Office app,
  save it to a temp path, and note that in the report.
