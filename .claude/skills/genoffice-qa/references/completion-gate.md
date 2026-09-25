# GenOffice Completion Gate

Before completing a development task, evaluate:

```text
[ ] Target service identified
[ ] Expected user workflow documented
[ ] Microsoft Office analogue considered
[ ] Baseline reproduced or characterized
[ ] Root cause identified
[ ] Correct implementation layer changed
[ ] Primary automated tests executed
[ ] Interactive workflow tested where possible
[ ] Selection/focus tested where applicable
[ ] Keyboard/pointer tested where applicable
[ ] Undo tested
[ ] Redo tested
[ ] Save/reload tested
[ ] Repeat-operation behavior tested
[ ] Boundary/invalid case tested
[ ] Adjacent regression tested
[ ] DOCX/XLSX/PPTX round-trip tested where applicable
[ ] Runtime errors inspected where possible
[ ] Known limitations documented
```

Use explicit statuses:

- PASS
- FAIL
- NOT TESTED
- BLOCKED
- N/A

## Severity

Critical:
- data loss;
- document corruption;
- save/open failure;
- widespread crash.

High:
- common editing workflow broken;
- incorrect formulas;
- broken undo/redo;
- corrupt clipboard;
- serious file compatibility issue.

Medium:
- uncommon workflow;
- state synchronization problem;
- significant UX inconsistency.

Low:
- cosmetic;
- minor non-blocking inconsistency.

## Final Status

Use:

- PASS
- PASS WITH KNOWN LIMITATIONS
- FAIL
- BLOCKED
- PARTIALLY TESTED
