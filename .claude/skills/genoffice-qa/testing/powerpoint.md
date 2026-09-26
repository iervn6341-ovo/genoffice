# PowerPoint-like (Slides) — E2E Checklist

Read `README.md` first (how to run, harness, false failures). Statuses,
severity and lifecycle follow `../references/completion-gate.md`.

## PowerPoint Test Suite

### P-01 Presentation Lifecycle

Test:

- launch;
- new presentation;
- open;
- save;
- save as;
- reopen;
- close;
- unsaved-change warning.

---

### P-02 Slide Lifecycle

Test:

- add;
- duplicate;
- delete;
- reorder;
- select;
- navigate;
- multi-select slides where supported.

Verify thumbnail and canvas synchronization.

---

### P-03 Slide Layouts

Where supported:

- title slide;
- title/content;
- blank;
- other templates/layouts.

Verify placeholders remain editable.

---

### P-04 Text Boxes

Test:

- insert;
- select;
- enter text mode;
- type;
- select text;
- format;
- exit text mode;
- move object;
- resize;
- delete;
- undo/redo;
- save/reopen.

Verify object-selection mode and text-editing mode remain distinct.

---

### P-05 Shapes

Test:

- insert;
- move;
- resize;
- rotate where supported;
- fill;
- border;
- duplicate;
- copy/paste;
- delete;
- save/reopen.

---

### P-06 Images

Test:

- insert;
- select;
- move;
- resize;
- aspect ratio behavior;
- crop where supported;
- duplicate;
- delete;
- persistence.

---

### P-07 Tables

Where supported:

- insert;
- edit cells;
- add/delete rows;
- add/delete columns;
- formatting;
- resize;
- copy/paste;
- undo/redo;
- save/reopen.

---

### P-08 Object Selection

Test:

- click object;
- click empty canvas;
- Shift multi-select;
- drag marquee where supported;
- overlapping objects;
- selection after zoom;
- selection after slide switch;
- selection after toolbar interaction.

---

### P-09 Move / Resize / Rotate

For multiple object types:

```text
select
→ drag move
→ verify position
→ resize
→ verify dimensions
→ rotate
→ undo
→ redo
→ save/reopen
```

Inspect numeric x/y/width/height values where possible.

---

### P-10 Alignment / Distribution

Where supported:

- align left;
- align right;
- align top;
- align bottom;
- center;
- distribute horizontally;
- distribute vertically.

Use multi-selection.

---

### P-11 Layering

Where supported:

- bring forward;
- send backward;
- bring to front;
- send to back.

Use overlapping objects.

Verify persistence.

---

### P-12 Group / Ungroup

Where supported:

```text
select multiple objects
→ group
→ move group
→ resize if supported
→ ungroup
→ undo
→ redo
→ save/reopen
```

---

### P-13 Text Formatting

Inside text boxes/shapes test:

- bold;
- italic;
- underline;
- font;
- size;
- color;
- alignment;
- bullets;
- numbering;
- line spacing where supported.

Test mixed formatting.

---

### P-14 Clipboard

Test:

- object copy;
- object cut;
- object paste;
- text copy/paste;
- multi-object copy;
- paste on another slide;
- duplicate;
- repeated paste.

Verify geometry offsets and layering.

---

### P-15 Keyboard Shortcuts

Test supported shortcuts:

- copy;
- cut;
- paste;
- undo;
- redo;
- select all;
- delete;
- duplicate;
- save;
- arrow movement where supported.

---

### P-16 Notes / Speaker Content

Where supported:

- add notes;
- edit notes;
- switch slide;
- save;
- reopen.

---

### P-17 Present / Slideshow Mode

Where supported:

- start presentation;
- next slide;
- previous slide;
- exit;
- verify slide order;
- verify rendering.

---

### P-18 Themes / Backgrounds

Where supported:

- theme;
- background;
- slide-level application;
- presentation-level application.

Verify persistence.

---

### P-19 Undo / Redo Stress

Sequence:

```text
add slide
→ add text box
→ type
→ add shape
→ move
→ resize
→ group
→ reorder slide
→ undo repeatedly
→ redo repeatedly
```

Verify exact restoration.

---

### P-20 Save / Reload

Create presentation containing:

- multiple slides;
- different layouts;
- text;
- shapes;
- images;
- tables where supported;
- layered objects.

Save/reopen.

Verify:

- slide order;
- object geometry;
- formatting;
- layering;
- content.

---

### P-21 PPTX Interoperability

#### Import

Use representative PPTX containing:

- multiple slides;
- text;
- shapes;
- images;
- tables;
- layering;
- formatting.

#### Export

Create equivalent content in GenOffice.

Export PPTX.

Reopen in GenOffice.

If Microsoft PowerPoint is available, open there too.

---

### P-22 Scale / Performance

Test:

- many slides;
- many objects on one slide;
- multiple images;
- rapid slide switching;
- repeated drag/resize.

Observe:

- responsiveness;
- canvas artifacts;
- memory;
- delayed selection;
- hit-testing defects.

---

### P-23 Invalid Operations and Recovery

Test safely:

- delete final object;
- delete slides repeatedly;
- move/resize near canvas boundaries;
- unsupported pasted content;
- malformed test import where available.

Verify no crash or silent corruption.

---

### P-24 Cross-Feature Workflows

#### Workflow A — Business Deck

```text
create title slide
→ add content slides
→ add text
→ add shapes
→ add image
→ align objects
→ reorder slides
→ save
→ reopen
→ present
```

#### Workflow B — Editing Stress

```text
duplicate slide
→ multi-select objects
→ group
→ move
→ resize
→ copy to another slide
→ undo
→ redo
→ save/reopen
```

---

## PowerPoint Final Regression Sweep

Verify at minimum:

```text
[ ] Launch
[ ] Presentation lifecycle
[ ] Slide lifecycle
[ ] Text editing
[ ] Object selection
[ ] Move/resize
[ ] Layering/grouping where supported
[ ] Clipboard
[ ] Undo/redo
[ ] Save/reopen
[ ] PPTX round-trip if supported
[ ] Runtime errors checked
```

---

---

## GenOffice notes — Slides

- **Existing specs:** `e2e/slides-*.spec.ts` (text editing, font styles,
  notes, ribbon, show, show controls).
- **Slide show:** needs real OS focus. Its specs use `foreground: true`. A
  second display is attached to the dev Mac for presenter-view testing. Esc
  must end the show, and arrows / PageUp / PageDown / B / W must work.
- **Text:**
  - Home/End.
  - ⌘⇧> / ⌘⇧< font-size steps.
  - Change Case.
  - Check the saved `a:rPr` in `ppt/slides/slideN.xml`.
- **Save checks:** `unzip -p f.pptx ppt/slides/slide1.xml`. Open the file in
  PowerPoint for any geometry or XML change.
- **Theme:** chrome follows the light/dark theme, but slide content never
  changes with it (CLAUDE.md rule 4). Compare the exported output in both
  themes.
