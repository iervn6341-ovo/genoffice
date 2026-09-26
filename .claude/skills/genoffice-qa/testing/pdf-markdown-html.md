# PDF Editor, Markdown, HTML — E2E Checklist

Read `README.md` first. Statuses, severity and lifecycle follow
`../references/completion-gate.md`.

Office benchmark:

- **PDF:** Acrobat or Preview where one exists. Otherwise the saved file must
  render identically in Preview and Chromium.
- **Markdown and HTML:** Word's editing conventions (caret, selection,
  shortcuts), plus round-trip fidelity of the source text.

---

## PDF Test Suite

Test file: `1_Stock_Funds_Rose_______in_____.pdf` (repo root, not committed).
Existing specs: `e2e/pdf-*.spec.ts` (`pdf-edit`, `pdf-fit-zoom`,
`pdf-outline-layout`, `pdf-shift-arrow-selection`).

### F-01 Open / View

- Open the file and check that every page renders.
- Zoom: fit width, fit page, and custom.
- Outline / thumbnails navigation.
- Rotated and cropped pages display correctly.

### F-02 Text Editing

- Edit a line of existing text. Check font matching, size and colour.
- Insert a new text box and type CJK text.
- Change the font, size, bold and colour.
- Check line re-measure after the edit (see `pdf-edit`).
- Lifecycle: undo → redo → save → reopen, then extract the text to check it
  (`pdftotext` or the app's read).

### F-03 Images

Insert, move, resize, crop, cutout, replace and delete an image.

- The Crop / Cutout dialog must fit the window at 1280×800.
- Lifecycle: undo → redo → save → reopen.

### F-04 Pages

- Rotate, delete, reorder and import pages.
- Import after a deleted page must land at the correct index.
- N-up / merge must honour CropBox and /Rotate.
- Pending edits must rotate with their page.

### F-05 Annotations and Redaction

- Highlight, draw and stamp.
- Area redaction: the content under it must be gone from the saved file
  (text extraction returns nothing there).
- Redaction undo/redo.
- Annotation appearances must be flattened when the file is merged.

### F-06 Layout Tools

- Crop box, page size and watermark.
- Crop and page size must be undoable.
- Watermark text must be searchable.

### F-07 Save / Interop

- Save and Save As a copy.
- Reopen the result in GenOffice and in Preview.
- The file size must stay reasonable.
- No other pages may change.

## Markdown Test Suite

Existing specs: `e2e/markdown-tab.spec.ts`, plus the Markdown cases in
`docs-*`.

### M-01 Editing

- Typing and Enter.
- Home/End on Mac.
- Double-click past the line end must not merge paragraphs.
- Lists, headings and code blocks.

### M-02 Source Fidelity

- Open a `.md` file, edit one paragraph, and save.
- Diff against the original: untouched lines must stay byte-identical,
  including front matter, tables and reference links.

### M-03 Export

Export to DOCX / HTML / PDF, then open the result. Headings, lists and code
must survive.

### M-04 Undo / Redo and Search Highlight

- Undo/redo across formatting.
- Find highlights must be cleared after an edit.

## HTML Test Suite

Existing specs: `e2e/html-tab.spec.ts`, `e2e/html-insert-drag.spec.ts`.

### H-01 Preview Editing

- Double-click text to edit it in the preview.
- Home/End on Mac.
- Commit the edit with Esc or by clicking outside.

### H-02 Insert and Drag

- Insert a paragraph, heading or image.
- A newly inserted text element must become editable right after the
  preview reload. This used to be a race; it is now version-gated.
- Drag to reorder.

### H-03 Source Round-Trip

After an edit, the source view shows minimal changes: attributes, scripts
and comments are preserved.

### H-04 Export

- HTML → DOCX (`packages/html2docx`; tests need `CHROME_PATH`).
- Open the exported file in Word.

---

## GenOffice notes

- The PDF editor's state is undo-bucketed in
  `apps/pdf/src/renderer/edit-state.ts`. A new editable feature needs a
  bucket, or it will not undo.
- PDF save logic lives in `apps/pdf/src/main/save-pdf.ts`. That is
  main-process code, so rebuild the **shell** after changing it.
- To verify redaction and text edits, extract the text from the saved file.
  Don't trust the canvas alone.
