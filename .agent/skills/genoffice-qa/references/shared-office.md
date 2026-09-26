# Shared Office QA Workflow

Apply this file to Word, Excel, and PowerPoint shared infrastructure.

## High-Risk Shared Systems

- undo/redo;
- selection;
- focus;
- keyboard shortcuts;
- clipboard;
- save/load;
- autosave;
- import/export;
- rendering;
- zoom/scroll;
- shared command state.

## Required Mutation Lifecycle

```text
open/create
→ perform operation
→ verify UI
→ verify model
→ undo
→ verify previous state
→ redo
→ verify changed state
→ save
→ reopen
→ verify
→ continue editing
```

## Undo / Redo

Verify:

- one logical operation creates the expected history unit;
- undo restores exact previous state;
- redo reapplies the state;
- a new edit after undo clears the correct redo branch;
- grouped typing/dragging behaves sensibly;
- opening a document does not create invalid history.

## Clipboard

Test where applicable:

```text
copy → paste
cut → paste
paste → undo
undo → redo
copy → paste twice
cross-document paste
```

Validate both visible output and underlying structure.

## Persistence

Verify:

```text
model
→ serialization
→ storage
→ deserialization
→ model
→ rendering
```

A successful save request alone is not sufficient.

## File Compatibility

For supported DOCX/XLSX/PPTX workflows:

```text
Office file
→ GenOffice import
→ inspect
→ modify
→ export
→ reopen in GenOffice
```

When Microsoft Office is available:

```text
GenOffice export
→ Microsoft Office open
→ inspect
```

Report unsupported constructs explicitly.

## Runtime Checks

Inspect where possible:

- JavaScript exceptions;
- rejected promises;
- failed network/file operations;
- renderer warnings;
- invalid model state.

A visually correct result with runtime errors should be investigated.
