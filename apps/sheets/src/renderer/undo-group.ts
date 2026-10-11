import { IUndoRedoService } from '@univerjs/core'

import type { UniverRuntime } from './univer-state'

/**
 * One ribbon command that runs several Univer commands (a cell style is a fill, a font, borders...)
 * leaves one undo entry per command, so a single ⌘Z took back only the last one. This runs `fn`
 * and folds every undo entry it pushed into one: undo replays the entries' undo mutations from
 * the newest to the oldest, redo replays the redo mutations in the original order.
 */
interface UndoItem {
  unitID: string
  undoMutations: readonly unknown[]
  redoMutations: readonly unknown[]
}

interface UndoRedoServiceInternals {
  _undoStacks?: Map<string, UndoItem[]>
  _updateStatus?: () => void
  pitchTopUndoElement(): unknown
}

interface UndoMark {
  service: UndoRedoServiceInternals
  unitId: string
  base: unknown
}

function markUndoTop(runtime: UniverRuntime): UndoMark | null {
  try {
    const service = runtime.univer.__getInjector().get<UndoRedoServiceInternals>(IUndoRedoService)
    const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
    if (!unitId) return null
    return { service, unitId, base: service.pitchTopUndoElement() }
  } catch {
    return null
  }
}

function foldSince(mark: UndoMark | null): void {
  if (!mark) return
  const { service, unitId, base } = mark
  const stack = service._undoStacks?.get(unitId)
  if (!stack) return
  // the entries pushed by fn sit after the entry that was on top before it ran
  const start = base ? stack.indexOf(base as UndoItem) + 1 : 0
  if (base && start === 0) return // the old top fell off the stack limit: leave it alone
  const added = stack.slice(start)
  if (added.length < 2) return
  const merged: UndoItem = {
    unitID: unitId,
    undoMutations: [...added].reverse().flatMap((item) => item.undoMutations),
    redoMutations: added.flatMap((item) => item.redoMutations),
  }
  stack.splice(start, added.length, merged)
  service._updateStatus?.()
}

export function runAsOneUndoStep(runtime: UniverRuntime, fn: () => void): void {
  const mark = markUndoTop(runtime)
  fn()
  foldSince(mark)
}

/** The same for a command whose Univer commands settle asynchronously (the Table facade). */
export async function runAsOneUndoStepAsync(
  runtime: UniverRuntime,
  fn: () => Promise<unknown>,
): Promise<void> {
  const mark = markUndoTop(runtime)
  try {
    await fn()
  } finally {
    foldSince(mark)
  }
}
