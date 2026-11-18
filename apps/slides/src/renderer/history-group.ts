/**
 * Undo granularity helpers. PowerPoint treats one user action over a selection (a nudge, a fill,
 * a delete, ...) as ONE undo step; our main-process handlers each push their own snapshot, so
 * an action that has to call several of them wraps the calls in a history group (they collapse
 * into the snapshot taken before the first one) or, for geometry, goes through the batcher
 * below (one `batchEditTransform` IPC for everything requested in the same tick).
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import { FIT_WIDTH } from './app-constants'

/**
 * Run `fn` so every edit it makes is a single undo step. Nestable; a failed `fn` still closes
 * the group, so the edits made so far collapse into one step instead of leaving a stuck batch.
 */
export async function withHistoryGroup<T>(fn: () => Promise<T>): Promise<T> {
  const api = window.slidesApi
  const began = (await api.beginHistoryBatch?.()) === true
  try {
    return await fn()
  } finally {
    if (began) await api.endHistoryBatch({ snapshot: false })
  }
}

export interface TransformReq {
  sourceId: string
  box: {
    x: number
    y: number
    w: number
    h: number
    rotationDeg: number
    tableRowHeightsPx?: number[]
  }
  groupId?: string
  /** Tables: pure move (the box size is the drawn size, not a resize) */
  keepTableSize?: boolean
}

type FlushFn = (slideIndex: number, reqs: TransformReq[]) => Promise<void>
const pendingGeometry = new Set<Promise<void>>()
export async function flushPendingTransforms(): Promise<void> {
  while (pendingGeometry.size) await Promise.all([...pendingGeometry])
}

/**
 * Collects the geometry commits requested within one tick on one slide — a multi-selection
 * drag/resize/rotate commit fires one per node, an arrow nudge loops over the selection — and
 * hands them over together so they can take one undo step. A later request for the same element
 * in the same tick replaces the earlier one.
 */
export function createTransformBatcher(
  flush: FlushFn,
  schedule: (cb: () => void) => void = (cb) => void setTimeout(cb, 0),
) {
  let slideIndex = -1
  let pending: TransformReq[] = []
  let waiters: Array<() => void> = []
  let inFlight = false
  let needsRerun = false

  const run = () => {
    if (inFlight) {
      needsRerun = true
      return
    }
    if (!pending.length) return
    inFlight = true
    const reqs = pending
    const done = waiters
    const slide = slideIndex
    pending = []
    waiters = []
    slideIndex = -1
    void flush(slide, reqs).finally(() => {
      inFlight = false
      done.forEach((d) => d())
      if (needsRerun || pending.length) {
        needsRerun = false
        schedule(run)
      }
    })
  }
  return {
    push(slide: number, req: TransformReq): Promise<void> {
      // a request for another slide never rides along with a pending one
      if (pending.length && slide !== slideIndex) run()
      slideIndex = slide
      const at = pending.findIndex((p) => p.sourceId === req.sourceId && p.groupId === req.groupId)
      if (at >= 0) pending[at] = req
      else pending.push(req)
      const done = new Promise<void>((resolve) => {
        waiters.push(resolve)
        if (!inFlight && waiters.length === 1) schedule(run)
      })
      pendingGeometry.add(done)
      void done.finally(() => pendingGeometry.delete(done))
      return done
    },
  }
}

/**
 * Flush used by the app: one element goes through editTransform (table grids, in-group boxes),
 * several go through batchEditTransform (one atomic transaction = one undo step). When the batch
 * is refused (one element cannot take a transform) the elements are applied one by one inside a
 * history group — still a single undo step, and the movable ones still move.
 */
export async function flushTransforms(
  slideIndex: number,
  reqs: TransformReq[],
  apply: (slide: RenderSlide) => void,
): Promise<void> {
  const api = window.slidesApi
  const single = async (r: TransformReq) =>
    api.editTransform({
      slideIndex,
      sourceId: r.sourceId,
      xPx: r.box.x,
      yPx: r.box.y,
      wPx: r.box.w,
      hPx: r.box.h,
      rotationDeg: r.box.rotationDeg,
      fitWidthPx: FIT_WIDTH,
      ...(r.groupId ? { groupId: r.groupId } : {}),
      ...(r.keepTableSize ? { keepTableSize: true } : {}),
      ...(r.box.tableRowHeightsPx ? { tableRowHeightsPx: r.box.tableRowHeightsPx } : {}),
    })
  if (reqs.length === 1) {
    const updated = await single(reqs[0]!)
    if (updated) apply(updated)
    return
  }
  const batch = await api.batchEditTransform({
    slideIndex,
    fitWidthPx: FIT_WIDTH,
    items: reqs.map((r) => ({
      sourceId: r.sourceId,
      ...(r.groupId ? { groupId: r.groupId } : {}),
      xPx: r.box.x,
      yPx: r.box.y,
      wPx: r.box.w,
      hPx: r.box.h,
      rotationDeg: r.box.rotationDeg,
      // same as editTransform: a resized table redistributes its grid (a no-op for plain moves)
      resizeTableGrid: true,
      ...(r.keepTableSize ? { keepTableSize: true } : {}),
      ...(r.box.tableRowHeightsPx ? { tableRowHeightsPx: r.box.tableRowHeightsPx } : {}),
    })),
  })
  if (batch) {
    apply(batch)
    return
  }
  await withHistoryGroup(async () => {
    for (const r of reqs) {
      const updated = await single(r)
      if (updated) apply(updated)
    }
  })
}
