# GenOffice Word Development & Test Workflow

Benchmark against Microsoft Word interaction semantics.

## Conceptual Execution Model

```text
document
→ paragraph
→ text/run
→ formatting
→ selection/caret
→ command/history
→ renderer
→ DOCX serializer
```

Use the repository's actual architecture.

## Text Editing

Test:

- insert at start/middle/end;
- Backspace;
- Delete;
- Enter/new paragraph;
- multiline paste;
- empty document;
- long paragraph;
- Unicode/CJK/emoji where supported.

## Selection and Caret

Test:

- caret only;
- partial text selection;
- cross-run selection;
- cross-paragraph selection;
- select all;
- forward/reverse selection;
- selection after toolbar interaction.

## Character Formatting

For:

- bold;
- italic;
- underline;
- strike;
- font family;
- font size;
- color;
- highlight.

Verify:

```text
select text
→ apply format
→ verify rendering
→ verify toolbar state
→ move caret inside formatted text
→ type
→ verify inherited formatting
→ undo
→ redo
→ save/reopen
```

Test mixed-format selections.

## Paragraph Formatting

Test:

- alignment;
- indentation;
- line spacing;
- bullets;
- numbering.

Include:

- one paragraph;
- multiple paragraphs;
- empty paragraph;
- partial cross-paragraph selection;
- list boundary behavior.

## Structured Content

Where supported, test:

- tables;
- images;
- links;
- page breaks;
- headers;
- footers.

For each:

```text
insert
→ select
→ edit
→ delete
→ undo
→ redo
→ save/reopen
```

## DOCX Round Trip

When applicable:

```text
DOCX import
→ inspect
→ modify
→ export
→ reopen
```

Check:

- text;
- formatting;
- paragraphs;
- tables;
- images;
- supported metadata.

## Regression Areas

Formatting changes can regress:

- typing attributes;
- copy/paste;
- style inheritance;
- paragraph splitting;
- DOCX serialization.

List changes can regress:

- Enter;
- Backspace;
- indentation;
- numbering;
- selection;
- undo grouping.
