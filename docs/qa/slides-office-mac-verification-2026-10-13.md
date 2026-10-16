# Slides Header/Footer verification — 2026-10-13

Scope: independent review of the Insert → Header and Footer slide controls and their PPTX persistence. This is a verification record for the implementation in [slides-office-mac-implementation-2026-10-13.md](slides-office-mac-implementation-2026-10-13.md), not a claim of full PowerPoint for Mac parity.

## Source review

- Effective state is read from a local placeholder first, then from the currently rendered layout/master decoration. A blank local placeholder therefore remains the effective “off” value after save and reopen.
- Clearing a value backed only by an inherited placeholder materializes an empty local placeholder. Enabling or changing one clones the nearest inherited layout/master placeholder, preserving its transform and run style. Relationship attributes are remapped into the destination slide relationship part; unresolved or missing package relationships reject the edit rather than writing a dangling reference.
- Current-only title exclusion uses local blank masks and does not write a forbidden slide-level `<p:hf>`. Apply to All stores deck-level exclusion on title layouts. Consequently, after a current-only title mask, the slide reopens with the inherited fields off while the deck-level exclusion checkbox remains unchecked; the implementation note records this distinction.
- The implementation note reports 10 Header/Footer engine cases, one operation-payload case, two date-format cases, three typechecks, and lint/diff checks passing. This independent review did not rerun those checks.

## Independent UI probe

The opt-in Playwright probe is [slides-office-mac-verification.spec.ts](../../e2e/slides-office-mac-verification.spec.ts). It creates an isolated scratch copy of the existing fixture with active master date/footer/slide-number placeholders, an uncommon cached dynamic date field, a styled footer, and a master hyperlink. It exercises effective UI echo, current-only title masking, current-slide editing, save/close/reopen of a fixed date, automatic date-format preview/cached text, Apply to All, title-layout exclusion, and a second reopen. It checks saved slide XML and the remapped slide relationship. The source fixture and historical audit evidence are not modified.

Status: **passed in the built Electron UI**. The verifier did not build the app; it used the prebuilt shell provided for the coordinated E2E slot. The passing command was:

```sh
GENOFFICE_SLIDES_HF_VERIFICATION=1 \
GENOFFICE_E2E_NO_VIDEO=1 \
GENOFFICE_SLIDES_VERIFICATION_DIR=/private/tmp/genoffice-slides-verification-20261013/slides \
npm run test:e2e -- e2e/slides-office-mac-verification.spec.ts \
  --output /private/tmp/genoffice-slides-verification-20261013/test-results --reporter=list
```

Result: **1 passed in 6.1 seconds**. The probe found the two title-layout slides at indexes 0 and 4, and content slides at indexes 1–3. On the imported master-only fixture, the dialog echoed the master’s dynamic `datetimeFigureOut` field, cached date `1/27/13`, slide number, and styled footer. Apply current blanked inherited decorations on the selected title slide and changed only the selected content slide. After save and reopen, the fixed-date radio remained selected with `January 2, 2001`; the new footer, inherited geometry and style, and exact external hyperlink relationship were present, while the other content slides did not receive the current-only footer.

For Apply to All, the UI preview and cached date text agreed (`00:48` at run time) for `datetime10`. The saved content slides all carried the dynamic field and new footer. Both title slides had blank local masks and their title layouts had footer-family suppression flags. After a second reopen, the title-exclusion checkbox remained checked, inherited master footer content remained hidden, and the dialog showed the saved automatic date format and footer. The fixture presents its local date/time example in the current UI locale; no comparison with native PowerPoint wording or locale output is claimed.

The first probe attempt timed out because `launchShell()` returns the shell home window before the Slides editor WebContents exists; the test now waits for the editor URL. A following attempt exposed a probe assumption that all slides after slide one were content slides. The fixture also has a second title-layout slide, so the assertion now classifies slides through their slide-layout relationships. These harness corrections preceded the passing run and did not change application code. Both failed setup runs are retained separately in [test-setup-issues](office-mac-evidence-2026-10-13/slides/test-setup-issues/README.md).

The final run’s JSON record, screenshot, imported/saved scratch PPTX files, and concise log are archived under [the passing run folder](office-mac-evidence-2026-10-13/slides/runs/2026-10-12T16-48-22-377Z-Kyl2sC/). Their original run directory is `/private/tmp/genoffice-slides-verification-20261013/slides/2026-10-12T16-48-22-377Z-Kyl2sC/`; Playwright output is under `/private/tmp/genoffice-slides-verification-20261013/test-results/`. The JSON preserves the exact paths written during the run. `GENOFFICE_SLIDES_VERIFICATION_DIR` is resolved to an absolute path so an accidental relative override cannot write evidence into the repository.

## Final visual follow-up after radio-size CSS fix

After the final Slides rebuild, the same probe was rerun with a new assertion that measures the `Update automatically` and `Fixed` radio inputs. Both returned `getBoundingClientRect()` dimensions of **14 × 14 px**, within the 18 px limit. I inspected the final screenshot and confirmed the radio controls are compact and aligned with their labels. The full Header/Footer acceptance path still passes: inherited master echo, current-only masks, fixed-date save/reopen, inherited style/geometry and hyperlink remap, automatic `datetime10` preview/cache agreement, Apply to All, both title slides excluded, and final reopen state.

The post-CSS run passed **1/1 in 6.9 seconds** against the rebuilt shell. Its JSON, screenshot, scratch PPTX files, and Playwright status are archived in [the final visual run folder](office-mac-evidence-2026-10-13/slides/runs/2026-10-12T17-06-37-426Z-JLZ5pQ/). The JSON records both radio bounds as `{width: 14, height: 14}`. The concise `playwright.log` notes that the initial `tee` path did not exist when the test started; Playwright itself reported the pass, and the accompanying `playwright-last-run.json` confirms `status: passed` with no failed tests. The earlier functional pass and setup-issue artifacts remain in their separate folders.

## Remaining limits

- The app build was supplied by the lead agent and was not independently reproduced by this verification run.
- Native PowerPoint save/reopen has not been run for this implementation. Exact PowerPoint menu wording and locale-specific date examples remain unverified on the native app.
- Notes and Handout Header/Footer controls and their persistence are outside this implementation. No claim is made for them or full PowerPoint parity.
