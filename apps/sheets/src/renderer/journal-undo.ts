import { CommandType, ICommandService, IUndoRedoService } from '@univerjs/core'

import type { LazyWorkbookState, UniverRuntime } from './univer-state'

/**
 * Some edits only change the journal the save writes — Univer has no model for them (page setup,
 * the document theme) — so they never reached the ⌘Z stack: Undo after "Landscape" silently took
 * back the previous cell edit. Such an edit runs through `withJournalUndo`, which snapshots the
 * part of the journal it touches before and after and, when it changed, pushes one undo entry
 * whose mutations put the snapshots back.
 */
export interface JournalUndoContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setPendingEdits: (count: number) => void
  refreshPageBreakPreview?: () => void
  /// re-renders the ribbon, which reads these values from the journal
  requestRender?: () => void
}

export interface JournalUndoKind<Snapshot> {
  /** the Univer mutation id this kind registers */
  readonly id: string
  /** a deep copy of what the edit may change (must be structured-cloneable) */
  capture(state: LazyWorkbookState, key: string): Snapshot
  /** puts a captured snapshot back, then repaints whatever shows it */
  restore(state: LazyWorkbookState, key: string, snapshot: Snapshot, ctx: JournalUndoContext): void
}

interface Params {
  unitId: string
  key: string
  snapshot: unknown
}

/** the context the registered mutations apply through (refreshed by every edit) */
let applyContext: JournalUndoContext | null = null
const registered = new WeakMap<object, Set<string>>()

function register<S>(runtime: UniverRuntime, kind: JournalUndoKind<S>): void {
  const commands = runtime.univer.__getInjector().get(ICommandService)
  let ids = registered.get(commands)
  if (!ids) {
    ids = new Set()
    registered.set(commands, ids)
  }
  if (ids.has(kind.id)) return
  ids.add(kind.id)
  commands.registerCommand({
    id: kind.id,
    type: CommandType.MUTATION,
    handler: (_accessor, params) => {
      const { key, snapshot } = params as Params
      const ctx = applyContext
      const state = ctx?.lazyWorkbookRef.current
      if (!ctx || !state) return false
      kind.restore(state, key, structuredClone(snapshot) as S, ctx)
      ctx.requestRender?.()
      return true
    },
  })
}

/**
 * Runs `edit` and, when it changed what `kind` captures for `key`, records that as one undo step.
 * `edit` may be async; returning false means nothing was applied.
 */
export async function withJournalUndo<S>(
  ctx: JournalUndoContext,
  kind: JournalUndoKind<S>,
  key: string,
  edit: () => Promise<boolean | void> | boolean | void,
): Promise<void> {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  const unitId = runtime?.univerAPI.getActiveWorkbook()?.getId()
  if (!runtime || !state || !unitId) {
    await edit()
    return
  }
  const before = kind.capture(state, key)
  const applied = await edit()
  if (applied === false || ctx.lazyWorkbookRef.current !== state) return
  const after = kind.capture(state, key)
  if (JSON.stringify(before) === JSON.stringify(after)) return
  register(runtime, kind)
  applyContext = ctx
  runtime.univer
    .__getInjector()
    .get(IUndoRedoService)
    .pushUndoRedo({
      unitID: unitId,
      undoMutations: [{ id: kind.id, params: { unitId, key, snapshot: before } }],
      redoMutations: [{ id: kind.id, params: { unitId, key, snapshot: after } }],
    })
}
