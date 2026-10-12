import { IUndoRedoService } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'

/// the fields of a Univer note this needs; the rest travels along untouched for the undo
export interface RemovableNote {
  readonly id: string
  readonly row: number
  readonly col: number
}

/**
 * Removes notes as ONE undoable step. The facade's `deleteNote()` runs the remove mutation without
 * pushing an undo entry, so ⌘Z after Review → Delete (or an AI-proposed note removal, or Clear
 * Comments and Notes) silently took back the previous edit and left the note gone. Pushing the
 * mutation pair ourselves keeps the note's journaling (the mutation listener sees it) and makes it
 * restorable; inside an undo batch (an AI plan) it merges into the batch's single entry.
 */
export async function deleteNotesUndoable(
  runtime: UniverRuntime,
  sheetId: string,
  notes: readonly RemovableNote[],
): Promise<number> {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId || notes.length === 0) return 0
  const redoMutations = notes.map((note) => ({
    id: 'sheet.mutation.remove-note',
    params: { unitId, sheetId, noteId: note.id },
  }))
  const undoMutations = notes.map((note) => ({
    id: 'sheet.mutation.update-note',
    params: { unitId, sheetId, row: note.row, col: note.col, note: { ...note } },
  }))
  for (const mutation of redoMutations) {
    await runtime.univerAPI.executeCommand(mutation.id, mutation.params)
  }
  runtime.univer
    .__getInjector()
    .get(IUndoRedoService)
    .pushUndoRedo({ unitID: unitId, undoMutations, redoMutations })
  return notes.length
}

/** a note as Univer's model holds it */
export interface NoteData extends RemovableNote {
  readonly width: number
  readonly height: number
  readonly note: string
}

/**
 * Creates or replaces a note as ONE undoable step. The facade's `createOrUpdateNote()` runs the
 * update mutation without pushing an undo entry either, so ⌘Z after an AI-proposed note or a
 * Paste Special → Comments took back the previous edit and left the note in place. `previous` is
 * the note the cell had (undo restores it; without one undo removes the new note).
 */
export async function setNoteUndoable(
  runtime: UniverRuntime,
  sheetId: string,
  next: NoteData,
  previous: NoteData | null | undefined,
): Promise<void> {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  const redo = {
    id: 'sheet.mutation.update-note',
    params: { unitId, sheetId, row: next.row, col: next.col, note: { ...next } },
  }
  const undo = previous
    ? {
        id: 'sheet.mutation.update-note',
        params: { unitId, sheetId, row: next.row, col: next.col, note: { ...previous } },
      }
    : {
        id: 'sheet.mutation.remove-note',
        params: { unitId, sheetId, noteId: next.id, row: next.row, col: next.col },
      }
  await runtime.univerAPI.executeCommand(redo.id, redo.params)
  runtime.univer
    .__getInjector()
    .get(IUndoRedoService)
    .pushUndoRedo({ unitID: unitId, undoMutations: [undo], redoMutations: [redo] })
}
