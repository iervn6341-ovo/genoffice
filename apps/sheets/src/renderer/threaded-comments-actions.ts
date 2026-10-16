import { newGuid } from '@genoffice/xlsx-gateway/gateway/xlsx-threaded-comments'

import type { WorkbookSheetThread, WorkbookThreadPerson } from '../shared/desktop-api'
import { isSheetRemoved, journalSize, recordNoteChange } from './edit-journal'
import { withJournalUndo, type JournalUndoContext, type JournalUndoKind } from './journal-undo'
import {
  addReply,
  addThread,
  deleteComment,
  editComment,
  ensurePerson,
  setThreadDone,
  threadTimestamp,
} from './threaded-comments'
import type { LazyWorkbookState } from './univer-state'

/**
 * Threaded comments live only in the edit journal (Univer has no model for them), so every change
 * goes through `withJournalUndo`: one ⌘Z entry per edit whose undo puts the sheet's thread list and
 * the person list back, and marks the sheet's comment set dirty so the next save writes it.
 */
interface ThreadSnapshot {
  threads: WorkbookSheetThread[]
  persons: WorkbookThreadPerson[]
}

const listeners = new Set<() => void>()

/** The panel and the cell markers repaint through this. */
export function subscribeThreads(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

export function notifyThreadsChanged(): void {
  for (const listener of [...listeners]) listener()
}

/** True while Univer is replaying an Undo, so an inverse row insert may restore stashed threads. */
export const threadUndoState = { undoing: false }

export const threadsKind: JournalUndoKind<ThreadSnapshot> = {
  id: 'genoffice.mutation.set-threaded-comments',
  capture: (state, sheetId) =>
    structuredClone({
      threads: state.editJournal.threads.get(sheetId) ?? [],
      persons: state.editJournal.threadPersons.list,
    }),
  restore: (state, sheetId, snapshot, ctx) => {
    const { editJournal } = state
    if (snapshot.threads.length > 0 || editJournal.threads.has(sheetId)) {
      editJournal.threads.set(sheetId, snapshot.threads)
    }
    editJournal.threadPersons.list = snapshot.persons
    recordNoteChange(editJournal, sheetId)
    ctx.setPendingEdits(journalSize(editJournal))
    notifyThreadsChanged()
  },
}

const AUTHOR_STORAGE_KEY = 'genoffice.sheets.commentAuthor'

export function commentAuthor(): string {
  try {
    return localStorage.getItem(AUTHOR_STORAGE_KEY) || 'GenOffice'
  } catch {
    return 'GenOffice'
  }
}

export function setCommentAuthor(name: string): void {
  try {
    localStorage.setItem(AUTHOR_STORAGE_KEY, name)
  } catch {
    // the name then only lasts for this session
  }
}

export type ThreadEditError = 'nofile' | 'occupied' | 'empty' | 'removed'

function editThreads(
  ctx: JournalUndoContext,
  sheetId: string,
  author: string,
  change: (
    state: LazyWorkbookState,
    threads: readonly WorkbookSheetThread[],
    person: { id: string; persons: WorkbookThreadPerson[] },
  ) => { threads: WorkbookSheetThread[]; error?: ThreadEditError },
): Promise<ThreadEditError | null> {
  const state = ctx.lazyWorkbookRef.current
  if (!state) return Promise.resolve('nofile')
  if (isSheetRemoved(state.editJournal, sheetId)) return Promise.resolve('removed')
  let failure: ThreadEditError | null = null
  return withJournalUndo(ctx, threadsKind, sheetId, () => {
    const current = state.editJournal.threads.get(sheetId) ?? []
    const who = ensurePerson(state.editJournal.threadPersons.list, author, () => newGuid())
    const result = change(state, current, { id: who.personId, persons: who.persons })
    if (result.error) {
      failure = result.error
      return false
    }
    state.editJournal.threads.set(sheetId, result.threads)
    state.editJournal.threadPersons.list = who.persons
    recordNoteChange(state.editJournal, sheetId)
    ctx.setPendingEdits(journalSize(state.editJournal))
    notifyThreadsChanged()
    return true
  }).then(() => failure)
}

export function createThread(
  ctx: JournalUndoContext,
  sheetId: string,
  cell: { row: number; column: number },
  text: string,
  author: string,
): Promise<ThreadEditError | null> {
  return editThreads(ctx, sheetId, author, (_state, threads, person) =>
    addThread(threads, {
      ...cell,
      text,
      personId: person.id,
      time: threadTimestamp(),
      id: newGuid(),
    }),
  )
}

export function replyToThread(
  ctx: JournalUndoContext,
  sheetId: string,
  threadId: string,
  text: string,
  author: string,
): Promise<ThreadEditError | null> {
  return editThreads(ctx, sheetId, author, (_state, threads, person) => ({
    threads: addReply(threads, threadId, {
      id: newGuid(),
      text,
      personId: person.id,
      time: threadTimestamp(),
    }),
  }))
}

export function resolveThread(
  ctx: JournalUndoContext,
  sheetId: string,
  threadId: string,
  done: boolean,
  author: string,
): Promise<ThreadEditError | null> {
  return editThreads(ctx, sheetId, author, (_state, threads) => ({
    threads: setThreadDone(threads, threadId, done),
  }))
}

export function removeThreadComment(
  ctx: JournalUndoContext,
  sheetId: string,
  threadId: string,
  commentId: string,
  author: string,
): Promise<ThreadEditError | null> {
  return editThreads(ctx, sheetId, author, (_state, threads) => ({
    threads: deleteComment(threads, threadId, commentId),
  }))
}

export function changeThreadComment(
  ctx: JournalUndoContext,
  sheetId: string,
  threadId: string,
  commentId: string,
  text: string,
  author: string,
): Promise<ThreadEditError | null> {
  return editThreads(ctx, sheetId, author, (_state, threads) => ({
    threads: editComment(threads, threadId, commentId, text),
  }))
}
