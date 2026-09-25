# C6 — Section breaks without save + reparse (design note)

## Problem

Layout ▸ Breaks ▸ (Next Page | Continuous | Even | Odd) inserted a generated
section-break paragraph and then forced `save(false, true)` so the reparse
would produce the new `SectionInfo`. With AutoSave off that was a silent
write to disk, and the reparse reset the editor history (Undo gone).

## What changed (minimal fix, this branch)

1. **No implicit save.** The insert is a single ProseMirror transaction; the
   status bar says the new section takes effect on the next save.
2. **Start type lives on the paragraph.** The generated `docProtected` break
   gets `breakStartType`. Word stores a section's start type in the sectPr
   that _ends_ it, so the chosen type belongs to the _next_ sectPr.
   `insertedBreakTypes()` (`apps/docs/src/renderer/ai/pending-sections.ts`)
   derives, at save time and from the paragraphs still in the document:
   - the next inserted break in the same section → its genXml is rewritten;
   - otherwise the original section's sectPr → `sectionStartType` (final
     section) or a sectPr rewrite (non-final section).
     Because nothing is held in app state, undo / redo of the insert and the
     saved file cannot disagree.
3. The AI `insert_section_break` path keeps its own pending-section model
   (unchanged); it no longer triggers the implicit save either.

## Known limitations (until the full model lands)

- Pagination/layout of the new section (different margins, orientation,
  odd/even start) is applied after the next save's reparse, not live.
- Page setup changed _inside_ a UI-inserted section before saving applies
  to the original section's sectPr.

## Full model (follow-up)

Represent sections in the editor document instead of `SectionInfo[]` state:
a `sectPr` attribute (or dedicated node) on the paragraph that ends each
section, plus a doc-level attr for the body sectPr. Page setup, headers /
footers and page numbering then become ProseMirror steps (shared undo stack
with C7), pagination reads sections from the doc, and the serializer writes
each paragraph's sectPr verbatim. `readSections()` becomes the importer into
that model; `sectionsDirty` / `trailingStartType` go away.
