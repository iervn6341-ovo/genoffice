# GenOffice PowerPoint Development & Test Workflow

Benchmark against Microsoft PowerPoint interaction semantics.

Preserve:

- slide order;
- object geometry;
- layering;
- editable structure.

## Conceptual Execution Model

```text
slide navigator
→ active slide
→ object selection
→ command/manipulation
→ slide/object model
→ history
→ canvas renderer
→ PPTX serializer
```

## Slide Operations

Test:

- create;
- duplicate;
- delete;
- reorder;
- select;
- undo/redo;
- save/reopen.

## Object Operations

For supported objects:

- text boxes;
- shapes;
- images;
- tables;
- media.

Test:

```text
insert
→ select
→ move
→ resize
→ edit
→ duplicate
→ copy/paste
→ layer change
→ undo/redo
→ save/reopen
```

Verify geometry numerically when possible.

## Selection Modes

Test:

- single object;
- multi-select;
- empty canvas click;
- text editing inside object;
- selection after slide change;
- selection after toolbar action.

Distinguish:

```text
object selection mode
!=
text editing mode
```

## Geometry

For move/resize/rotate where supported:

Verify:

- x/y;
- width/height;
- minimum size;
- drag direction;
- aspect ratio;
- zoomed canvas behavior;
- persistence.

## Text Inside Objects

Test:

```text
enter text mode
→ select text
→ format text
→ exit text mode
→ move shape
→ undo
→ redo
→ save/reopen
```

Text commands must not accidentally transform the shape itself.

## Layering / Grouping / Alignment

Where supported, test:

- bring forward;
- send backward;
- front/back;
- group;
- ungroup;
- align;
- distribute.

Include undo/redo and persistence.

## PPTX Round Trip

```text
PPTX import
→ inspect
→ modify
→ export
→ reopen
```

Check:

- slides;
- text;
- object positions;
- sizes;
- supported formatting;
- layering.
