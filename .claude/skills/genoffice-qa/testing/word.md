# Word-like (Docs) — E2E Checklist

Read `README.md` first (how to run, harness, false failures). Statuses,
severity and lifecycle follow `../references/completion-gate.md`.

## Word Test Suite

### W-01 Application and Document Lifecycle

Test:

- launch Word service;
- create blank document;
- create multiple documents if supported;
- open existing document;
- open recent document;
- save;
- save as;
- rename;
- close;
- reopen;
- unsaved-change warning.

Verify:

- active document state;
- document title;
- persistence;
- no unexpected content loss.

---

### W-02 Basic Text Editing

Test:

- type one character;
- type a sentence;
- type multiple paragraphs;
- insert at beginning;
- insert in middle;
- append at end;
- Backspace;
- Delete;
- Enter;
- Shift+Enter if supported;
- multiline paste;
- Unicode;
- CJK;
- emoji if supported.

Boundary cases:

- empty document;
- empty paragraph;
- long paragraph;
- many paragraphs.

---

### W-03 Caret and Selection

Test:

- click to position caret;
- Arrow navigation;
- Home/End where supported;
- select character;
- select word;
- select paragraph;
- drag selection;
- Shift+Arrow;
- select all;
- forward selection;
- reverse selection;
- cross-paragraph selection.

Verify:

- visible selection;
- insertion point;
- selection after toolbar interaction;
- selection after scrolling.

---

### W-04 Character Formatting

Test every implemented formatting feature, such as:

- bold;
- italic;
- underline;
- strike-through;
- font family;
- font size;
- font color;
- highlight;
- superscript/subscript where supported.

For each feature:

```text
create text
→ select partial text
→ apply formatting
→ verify rendering
→ verify toolbar state
→ move caret into formatted text
→ type
→ verify inherited formatting
→ undo
→ redo
→ save
→ reopen
```

Also test:

- no selection;
- mixed-format selection;
- toggling;
- combining multiple styles.

---

### W-05 Paragraph Formatting

Test implemented features such as:

- left alignment;
- center;
- right;
- justify;
- indentation;
- first-line indent;
- hanging indent;
- line spacing;
- paragraph spacing;
- bullets;
- numbered lists;
- checklist where supported.

Cases:

- one paragraph;
- multiple paragraphs;
- partial multi-paragraph selection;
- empty paragraph;
- Enter inside list;
- Enter at list end;
- Backspace at list boundary;
- nested list where supported.

---

### W-06 Styles and Headings

Where supported test:

- Normal;
- Heading 1;
- Heading 2;
- Heading 3;
- Title;
- Subtitle;
- custom styles.

Verify:

- visual style;
- structural role;
- behavior after Enter;
- style persistence;
- interaction with manual formatting.

---

### W-07 Tables

Where supported:

- insert table;
- type in cells;
- Tab navigation;
- select cell;
- select row/column;
- insert/delete row;
- insert/delete column;
- merge/split cells;
- resize;
- alignment;
- borders;
- copy/paste;
- delete table;
- undo/redo;
- save/reopen.

---

### W-08 Images and Embedded Objects

Where supported:

- insert image;
- select;
- move;
- resize;
- delete;
- copy/paste;
- text wrapping;
- alignment;
- alt text/caption where supported;
- save/reopen.

---

### W-09 Links

Where supported:

- create;
- edit;
- remove;
- click;
- copy linked text;
- undo/redo;
- save/reopen.

---

### W-10 Page and Document Structure

Where supported:

- page break;
- section break;
- header/footer;
- page number;
- margins;
- orientation;
- page size.

Verify pagination and persistence.

---

### W-11 Clipboard

Test:

- copy text;
- cut text;
- paste;
- copy formatted text;
- paste formatted text;
- paste into different paragraph;
- paste across documents;
- copy table;
- copy image;
- repeated paste;
- undo/redo.

---

### W-12 Keyboard Shortcuts

Test supported shortcuts:

- Ctrl/Cmd+C;
- Ctrl/Cmd+X;
- Ctrl/Cmd+V;
- Ctrl/Cmd+Z;
- Redo shortcut;
- Ctrl/Cmd+A;
- Ctrl/Cmd+B;
- Ctrl/Cmd+I;
- Ctrl/Cmd+U;
- Ctrl/Cmd+S.

Verify the application captures the shortcut rather than triggering unintended browser behavior.

---

### W-13 Undo / Redo Stress

Sequence:

```text
type
→ bold
→ Enter
→ type
→ paste
→ delete
→ undo repeatedly
→ redo repeatedly
```

Also test:

```text
action A
→ action B
→ undo
→ new action C
```

Verify redo history is handled correctly.

---

### W-14 Save / Reload

Create a document containing:

- multiple paragraphs;
- mixed formatting;
- list;
- table;
- image where supported;
- link where supported.

Then:

```text
save
→ close/reload
→ reopen
→ compare
```

Check:

- content;
- formatting;
- structure;
- object state;
- metadata where relevant.

---

### W-15 DOCX Interoperability

If DOCX is supported:

#### Import

Open representative DOCX files containing:

- headings;
- paragraphs;
- mixed fonts;
- lists;
- tables;
- images.

Record preserved and unsupported constructs.

#### Export

Create equivalent content in GenOffice.

Export DOCX.

Reopen in GenOffice.

If Microsoft Word is available, open the exported DOCX there.

---

### W-16 Invalid Operations and Recovery

Test safely:

- Backspace at document start;
- Delete at document end;
- delete all content;
- unsupported pasted content;
- malformed test import if fixture exists.

Verify no crash or silent corruption.

---

### W-17 Large Document Behavior

Test:

- many paragraphs;
- long pasted content;
- repeated formatting;
- long undo history;
- scrolling large document.

Observe:

- typing latency;
- selection lag;
- rendering defects;
- excessive memory growth.

---

### W-18 Cross-Feature Workflows

#### Workflow A — Report Creation

```text
new document
→ title
→ headings
→ paragraphs
→ bold/italic
→ bullet list
→ table
→ image
→ save
→ reopen
→ edit
→ export
```

#### Workflow B — Heavy Editing

```text
paste long text
→ multi-paragraph selection
→ format
→ cut section
→ paste elsewhere
→ undo
→ redo
→ save
```

---

## Word Final Regression Sweep

Before completing the run, verify at minimum:

```text
[ ] Launch
[ ] Create/open document
[ ] Type/edit
[ ] Selection/caret
[ ] Formatting
[ ] Clipboard
[ ] Undo/redo
[ ] Save/reopen
[ ] DOCX round-trip if supported
[ ] Runtime errors checked
```

---

---

## GenOffice notes — Docs

- **State:** `window.__aidocs.editor`.
  - Paragraph nodes are `docParagraph` (attrs `indentLeft`,
    `indentFirstLine`; negative first-line = hanging; all in twips).
  - List items are `docListItem`.
  - The tab grid is in `storage.tabStops.defaultTabStopTwips`.
- **Existing specs:** `e2e/docs-*.spec.ts`.
  - Typing, Home/End, list AutoFormat, Clear Formatting, double-click past
    line end, Tab/Backspace at a paragraph start: `docs-word-typing`.
  - Special indent (first line / hanging), ⌘T: `docs-paragraph-indent`.
  - Lists, tables, pictures: `docs-lists-tables-pictures`.
  - Section-break undo: `docs-section-break-undo`.
  - Layout undo: `docs-layout-undo`.
  - Style-inherited bold: `docs-style-bold-toggle`.
- **Word behavior already probed** (expected behavior; directly observed in
  Word for Mac, zh-TW):
  - **Tab at the start of a non-empty paragraph.**
    - 1st Tab → first-line indent of one tab stop.
    - 2nd Tab → left indent += one stop.
    - Backspace reverses it: first-line → 0, then left -= one stop.
    - Mid-text Tab types a tab character.
  - **⌘T** → hanging indent of one tab stop (left 24pt, first line −24pt in
    zh-TW).
  - **Paragraph dialog:**
    - Left / Right in 字元.
    - Special = (無)/首行/凸排 with a 位移點數 value.
    - For a hanging paragraph, Left shows `w:left − hang`.
    - OK on untouched fields must not change them.
  - **List AutoFormat:** `* `, `- `, `1. `, `a) ` start a list; Enter on an
    empty item leaves the list; ⌘Z right after restores the typed marker.
  - **Double-click** in the blank area right of a line → caret at the line
    end (Click-and-Type); it never merges paragraphs.
- **CJK checklist additions** (always run these in `lang: 'zh-TW'`):
  - Full-width punctuation and IME composition, then undo.
  - 2-字元 first-line indent.
  - Hanging indent.
  - Line wrap near the right margin compared with Word.
  - PMingLiU / 新細明體 defaults on a new document.
- **Save checks:** `unzip -p f.docx word/document.xml`. Verify `w:ind`,
  `w:pStyle`, `w:numPr` and `w:sectPr`, and open the saved file in Word when
  the change touches serialization.
- **Known open items:** see `docs/qa/cloud-handoff.md` (character-unit
  indents are saved as twips, not as `w:firstLineChars`).
