---
name: genoffice-qa
description: >
  Development and QA workflow for GenOffice Word, Excel, and PowerPoint-like services.
  Use whenever implementing, modifying, debugging, reviewing, or testing GenOffice editor
  functionality. Benchmark user-visible behavior against Microsoft Office and verify
  realistic workflows, undo/redo, persistence, edge cases, interoperability, and regressions.
---

# GenOffice Development & QA Skill

## Goal

Treat GenOffice as an Office-class productivity suite.

Microsoft Office is the primary behavioral benchmark:

- Word-like service → Microsoft Word
- Excel-like service → Microsoft Excel
- PowerPoint-like service → Microsoft PowerPoint

The goal is behavioral consistency, not pixel-perfect cloning.

A feature is NOT complete simply because:

- code compiles;
- UI renders;
- no exception occurs;
- one happy-path test succeeds.

A feature is complete only when the relevant user workflow, editor state,
undo/redo, save/reload, edge cases, and regressions have been verified.

---

# Skill Routing

Load service-specific guidance depending on the task.

## Word

Read:

`references/word.md`

Use for:

- text editing;
- selection/caret;
- character formatting;
- paragraph formatting;
- lists;
- tables;
- images;
- DOCX behavior.

## Excel

Read:

`references/excel.md`

Use for:

- cells;
- ranges;
- formulas;
- recalculation;
- rows/columns;
- merge/unmerge;
- formatting;
- XLSX behavior.

## PowerPoint

Read:

`references/powerpoint.md`

Use for:

- slides;
- objects;
- text boxes;
- shapes;
- images;
- geometry;
- layering;
- PPTX behavior.

## Shared Office Infrastructure

Read:

`references/shared-office.md`

when modifying:

- undo/redo;
- selection;
- focus;
- keyboard shortcuts;
- clipboard;
- persistence;
- serialization;
- import/export;
- rendering;
- shared commands.

Before declaring a task complete, read:

`references/completion-gate.md`

Use:

`templates/test-report.md`

for the final QA report.

---

# Required Development Flow

## Phase 1 — Classify

Identify:

- affected service;
- feature or bug;
- expected user scenario;
- Office equivalent;
- modules likely involved;
- shared infrastructure touched;
- risk level.

Risk levels:

- Critical
- High
- Medium
- Low

---

## Phase 2 — Define Expected Behavior

Before changing code, document:

```text
Feature:
User scenario:
Precondition:
Action:
Expected visible result:
Expected state change:
Expected persistence:
Undo expectation:
Redo expectation:
Office-equivalent behavior:
Related features:
```

---

## Phase 3 — Reproduce or Characterize

For bugs:

1. reproduce the issue when possible;
2. record exact actions;
3. record actual result;
4. identify expected result.

For new features:

1. describe current baseline;
2. define acceptance criteria.

---

## Phase 4 — Trace Execution

Trace the relevant flow:

```text
UI Event
→ Command / Action
→ Editor / Controller
→ Document Model
→ History
→ Renderer
→ Serializer / Persistence
```

Find the earliest point where actual behavior diverges from expected behavior.

Prefer root-cause fixes.

---

## Phase 5 — Implement

Prefer:

- existing architecture;
- reusable logic;
- history-aware mutations;
- deterministic serialization;
- general solutions.

Avoid:

- test-specific hardcoded branches;
- hiding invalid state only in the renderer;
- bypassing history;
- changing expected tests only to silence failures.

---

## Phase 6 — Narrow Verification

Run the smallest relevant automated tests first.

Do not continue to broad regression testing while the primary scenario still fails.

---

## Phase 7 — Interactive Verification

When tools allow, verify through real user interaction.

Observe:

- visible result;
- toolbar state;
- selection;
- focus;
- active document/sheet/slide;
- runtime errors.

Do not treat direct internal API calls as sufficient verification for UI behavior.

---

## Phase 8 — Lifecycle Verification

For persistent editor mutations:

```text
perform action
→ verify state
→ undo
→ verify original state
→ redo
→ verify changed state
→ save
→ reload/reopen
→ verify persisted result
→ continue editing
```

---

## Phase 9 — Regression

Test nearby features sharing:

- document model;
- history;
- renderer;
- commands;
- serializer;
- clipboard;
- selection;
- file-format adapters.

---

# Status Rules

Allowed test statuses:

- PASS
- FAIL
- NOT TESTED
- BLOCKED
- N/A

Never mark a test PASS if it was not actually executed or deterministically verified.

Final status must be one of:

- PASS
- PASS WITH KNOWN LIMITATIONS
- FAIL
- BLOCKED
- PARTIALLY TESTED

---

# Truthfulness

Clearly distinguish:

- code inspected;
- automated test executed;
- interactive test executed;
- save/reload verified;
- file round-trip verified;
- Microsoft Office behavior directly observed;
- Microsoft Office behavior inferred.

Do not claim:

- "same as Microsoft Office";
- "fully compatible";
- "all tests pass";

without sufficient evidence.
