/// Editor-side model of threaded comments (Excel 365 "Comments"): immutable operations over the
/// sheet's thread list, the person list, and the position bookkeeping that row/column edits need.
/// Univer has no model for them, so the list lives in the edit journal and every change is one
/// `withJournalUndo` step (see threaded-comments-actions.ts).

import type { WorkbookSheetThread, WorkbookThreadPerson } from '../shared/desktop-api'

export type SheetThreads = readonly WorkbookSheetThread[]
export type ThreadComment = WorkbookSheetThread['comments'][number]

export const COMMENT_TEXT_LIMIT = 32_767

export interface ThreadStashEntry {
  readonly sheetId: string
  readonly axis: 'row' | 'column'
  /// the remove op that took the thread out (a matching insert during Undo brings it back)
  readonly index: number
  readonly count: number
  readonly thread: WorkbookSheetThread
}

/// Excel stores dT as local time with centiseconds: 2026-10-15T08:00:00.00
export function threadTimestamp(now: Date = new Date()): string {
  const pad = (value: number, size = 2): string => String(value).padStart(size, '0')
  return (
    `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}` +
    `T${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}` +
    `.${pad(Math.floor(now.getMilliseconds() / 10))}`
  )
}

export function threadAt(
  threads: SheetThreads,
  row: number,
  column: number,
): WorkbookSheetThread | undefined {
  return threads.find((thread) => thread.row === row && thread.column === column)
}

export function normalizeCommentText(text: string): string {
  return text.replace(/\r\n?/g, '\n').slice(0, COMMENT_TEXT_LIMIT)
}

/// Finds or appends the person for a display name; the same name keeps the same id.
export function ensurePerson(
  persons: readonly WorkbookThreadPerson[],
  displayName: string,
  makeId: () => string,
): { persons: WorkbookThreadPerson[]; personId: string } {
  const name = displayName.trim() || 'GenOffice'
  const found = persons.find((person) => person.displayName === name)
  if (found) return { persons: [...persons], personId: found.id }
  const personId = makeId()
  return {
    persons: [
      ...persons,
      { id: personId, displayName: name, userId: name.toLowerCase(), providerId: 'None' },
    ],
    personId,
  }
}

export function addThread(
  threads: SheetThreads,
  input: {
    row: number
    column: number
    text: string
    personId: string
    time: string
    id: string
  },
): { threads: WorkbookSheetThread[]; error?: 'occupied' | 'empty' } {
  const text = normalizeCommentText(input.text)
  if (text.trim() === '') return { threads: [...threads], error: 'empty' }
  if (threadAt(threads, input.row, input.column))
    return { threads: [...threads], error: 'occupied' }
  return {
    threads: [
      ...threads,
      {
        id: input.id,
        row: input.row,
        column: input.column,
        done: false,
        comments: [{ id: input.id, personId: input.personId, time: input.time, text }],
      },
    ],
  }
}

function mapThread(
  threads: SheetThreads,
  threadId: string,
  change: (thread: WorkbookSheetThread) => WorkbookSheetThread | null,
): WorkbookSheetThread[] {
  const next: WorkbookSheetThread[] = []
  for (const thread of threads) {
    if (thread.id !== threadId) next.push(thread)
    else {
      const changed = change(thread)
      if (changed) next.push(changed)
    }
  }
  return next
}

export function addReply(
  threads: SheetThreads,
  threadId: string,
  reply: { id: string; text: string; personId: string; time: string },
): WorkbookSheetThread[] {
  const text = normalizeCommentText(reply.text)
  if (text.trim() === '') return [...threads]
  return mapThread(threads, threadId, (thread) => ({
    ...thread,
    // replying to a resolved thread reopens it, like Excel
    done: false,
    comments: [
      ...thread.comments,
      { id: reply.id, personId: reply.personId, time: reply.time, text, parentId: thread.id },
    ],
  }))
}

export function setThreadDone(
  threads: SheetThreads,
  threadId: string,
  done: boolean,
): WorkbookSheetThread[] {
  return mapThread(threads, threadId, (thread) => {
    if (thread.done === done) return thread
    const [root, ...replies] = thread.comments
    // the root element carries the done flag: drop its raw xml so the writer regenerates it
    const { raw: _raw, ...rootWithoutRaw } = root!
    return { ...thread, done, comments: [rootWithoutRaw, ...replies] }
  })
}

export function editComment(
  threads: SheetThreads,
  threadId: string,
  commentId: string,
  text: string,
): WorkbookSheetThread[] {
  const next = normalizeCommentText(text)
  if (next.trim() === '') return [...threads]
  return mapThread(threads, threadId, (thread) => ({
    ...thread,
    comments: thread.comments.map((comment) => {
      if (comment.id !== commentId || comment.text === next) return comment
      const { raw: _raw, ...rest } = comment
      return { ...rest, text: next }
    }),
  }))
}

/// Deleting the root removes the whole thread (Excel's "Delete thread"); a reply goes alone.
export function deleteComment(
  threads: SheetThreads,
  threadId: string,
  commentId: string,
): WorkbookSheetThread[] {
  return mapThread(threads, threadId, (thread) => {
    if (thread.comments[0]?.id === commentId) return null
    const comments = thread.comments.filter((comment) => comment.id !== commentId)
    return comments.length === thread.comments.length ? thread : { ...thread, comments }
  })
}

export type ThreadStructuralOp =
  | {
      kind: 'insert-rows' | 'remove-rows' | 'insert-cols' | 'remove-cols'
      index: number
      count: number
    }
  | { kind: 'move-rows'; index: number; count: number; before: number }

function moveRow(row: number, index: number, count: number, before: number): number {
  if (row >= index && row < index + count) {
    return before > index ? before - count + (row - index) : before + (row - index)
  }
  if (before > index && row >= index + count && row < before) return row - count
  if (before < index && row >= before && row < index) return row + count
  return row
}

/**
 * Re-places threads after a row/column edit. Removed lines take their threads out of the list
 * into the stash; the inverse insert that Undo runs brings them back at their old cells.
 */
export function shiftThreads(
  sheetId: string,
  threads: SheetThreads,
  stash: readonly ThreadStashEntry[],
  op: ThreadStructuralOp,
  duringUndo: boolean,
): { threads: WorkbookSheetThread[]; stash: ThreadStashEntry[] } {
  if (op.kind === 'move-rows') {
    return {
      threads: threads.map((thread) => ({
        ...thread,
        row: moveRow(thread.row, op.index, op.count, op.before),
      })),
      stash: [...stash],
    }
  }
  const axis = op.kind.endsWith('rows') ? 'row' : 'column'
  const at = (thread: WorkbookSheetThread): number => (axis === 'row' ? thread.row : thread.column)
  const place = (thread: WorkbookSheetThread, value: number): WorkbookSheetThread =>
    axis === 'row' ? { ...thread, row: value } : { ...thread, column: value }
  if (op.kind.startsWith('insert')) {
    const moved = threads.map((thread) =>
      at(thread) >= op.index ? place(thread, at(thread) + op.count) : thread,
    )
    if (!duringUndo) return { threads: moved, stash: [...stash] }
    const restorable = stash.filter(
      (entry) =>
        entry.sheetId === sheetId &&
        entry.axis === axis &&
        entry.index === op.index &&
        entry.count === op.count,
    )
    return {
      threads: [...moved, ...restorable.map((entry) => entry.thread)],
      stash: stash.filter((entry) => !restorable.includes(entry)),
    }
  }
  const end = op.index + op.count
  const kept: WorkbookSheetThread[] = []
  const removed: ThreadStashEntry[] = []
  for (const thread of threads) {
    const position = at(thread)
    if (position >= op.index && position < end)
      removed.push({ sheetId, axis, index: op.index, count: op.count, thread })
    else kept.push(position >= end ? place(thread, position - op.count) : thread)
  }
  return { threads: kept, stash: [...stash, ...removed] }
}
