import { UniverInstanceType } from '@univerjs/core'
import type { IRenderManagerService } from '@univerjs/engine-render'

/** Limit backing-store reallocations and full grid redraws during native window drags. */
export function coalesceCanvasResize(engine: { resize(): void }): () => void {
  const original = engine.resize
  let initialized = false
  let timer: ReturnType<typeof setTimeout> | undefined
  engine.resize = function () {
    // The first mount must size the canvas before the first grid paint.
    if (!initialized) {
      initialized = true
      original.call(engine)
      return
    }
    if (timer !== undefined) return
    timer = setTimeout(() => {
      timer = undefined
      // Read the current container size, rather than an intermediate resize entry.
      original.call(engine)
    }, 80)
  }
  return () => {
    clearTimeout(timer)
    engine.resize = original
  }
}

/** Workbook replacements create a new engine; follow its lifetime without delaying cell editors. */
export function installCanvasResizeCoalescing(manager: IRenderManagerService): () => void {
  const pending = new Map<string, () => void>()
  const install = (render: ReturnType<IRenderManagerService['getRenderById']>) => {
    if (!render || render.type !== UniverInstanceType.UNIVER_SHEET || pending.has(render.unitId))
      return
    pending.set(render.unitId, coalesceCanvasResize(render.engine))
  }
  for (const render of manager.getRenderAll().values()) install(render)
  const created = manager.created$.subscribe(install)
  const disposed = manager.disposed$.subscribe((id) => {
    pending.get(id)?.()
    pending.delete(id)
  })
  return () => {
    created.unsubscribe()
    disposed.unsubscribe()
    for (const dispose of pending.values()) dispose()
    pending.clear()
  }
}
