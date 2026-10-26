/**
 * Element arrangement actions extracted from App.tsx: grouping,
 * ungrouping, align/distribute, z-order, and freehand ink. Functions read the
 * latest App state through ActionCtx.
 */
import type { GroupRenderNode, RenderNode } from '@genoffice/pptx-render'
import type { ReorderDirection } from '../shared/ipc'
import type { ActionCtx } from './action-context'
import { FIT_WIDTH } from './app-constants'
import { withHistoryGroup } from './history-group'
import { alignPositions, type AlignOp, type AlignTo } from './align-geometry'
import { inkNodesOf, rasterizeStroke, type InkStroke } from './ink'
import { t } from './i18n/locale'

export async function groupSelected(ctx: ActionCtx): Promise<void> {
  const { slide, selectedIds, current } = ctx
  if (!slide || selectedIds.length < 2) return
  const GROUPABLE = new Set(['text', 'shape', 'picture'])
  const nodes = selectedIds
    .map((id) => slide.nodes.find((n) => n.sourceId === id))
    .filter(Boolean) as RenderNode[]
  if (nodes.some((n) => !GROUPABLE.has(n.type))) return
  const result = await window.slidesApi.groupElements({
    slideIndex: current,
    sourceIds: selectedIds,
  })
  if (result) {
    ctx.applySlide(current, result.slide)
    ctx.setSelectedIds([result.groupId])
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusGrouped'))
  }
}

export async function ungroupSelected(ctx: ActionCtx): Promise<void> {
  const { slide, selectedIds, current } = ctx
  if (!slide || selectedIds.length !== 1) return
  const id = selectedIds[0]!
  const node = slide.nodes.find((n) => n.sourceId === id)
  if (!node || node.type !== 'group') return
  const groupNode = node as GroupRenderNode
  const childIds = groupNode.children.map((c) => c.sourceId)
  const updated = await window.slidesApi.ungroupElement({
    slideIndex: current,
    sourceId: id,
  })
  if (updated) {
    ctx.applySlide(current, updated)
    const newIds = childIds.filter((cid) => updated.nodes.some((n) => n.sourceId === cid))
    ctx.setSelectedIds(newIds.length > 0 ? newIds : [])
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusUngrouped'))
  }
}

export type { AlignOp, AlignTo } from './align-geometry'

/**
 * Align/distribute selected elements.
 * Coordinates are computed in the renderer (pure geometry, align-geometry.ts) and updated in bulk via
 * batchEditTransform (one undo step).
 * - align: 6 alignments, relative to the selection's bounding box or to the slide (`to`);
 *   a single selected element always aligns to the slide
 * - distribute: horizontal / vertical, 3+ elements (selection) or 2+ (slide edges)
 */
export async function alignSelected(
  ctx: ActionCtx,
  op: AlignOp,
  to: AlignTo = 'selection',
): Promise<void> {
  const { slide, selectedIds, current } = ctx
  if (!slide || selectedIds.length === 0) return
  const nodes = selectedIds
    .map((id) => slide.nodes.find((n) => n.sourceId === id))
    .filter(Boolean) as RenderNode[]
  if (nodes.length === 0) return

  const rects = nodes.map((n) => ({ x: n.box.x, y: n.box.y, w: n.box.w, h: n.box.h }))
  const newPositions = alignPositions(rects, op, to, { w: slide.widthPx, h: slide.heightPx })
  if (!newPositions) return

  const items = nodes.map((n, i) => ({
    sourceId: n.sourceId,
    xPx: newPositions[i]!.x,
    yPx: newPositions[i]!.y,
    wPx: n.box.w,
    hPx: n.box.h,
    rotationDeg: n.box.rotationDeg,
  }))

  const updated = await window.slidesApi.batchEditTransform({
    slideIndex: current,
    fitWidthPx: FIT_WIDTH,
    items,
  })
  if (updated) {
    ctx.applySlide(current, updated)
    ctx.setDirty(true)
  }
}

export async function reorderSelected(
  ctx: ActionCtx,
  sourceId: string,
  dir: ReorderDirection,
): Promise<void> {
  // Right-clicking inside a multi-selection reorders the whole selection (PowerPoint), keeping
  // its internal stacking (see stackSorted for the move order)
  const sel = ctx.selectedIds.includes(sourceId) ? ctx.selectedIds : [sourceId]
  const ids = sel.length > 1 ? stackSorted(ctx.slide?.nodes ?? [], sel, dir) : [sourceId]
  await withHistoryGroup(async () => {
    for (const id of ids) {
      const updated = await window.slidesApi.reorderElement({
        slideIndex: ctx.current,
        sourceId: id,
        dir,
      })
      if (updated) ctx.applySlide(ctx.current, updated)
    }
  })
}

/**
 * Ids in the order the moves must run so the group's relative stacking survives:
 *  - front: lowest first (each lands on top of the previous, the highest ends up topmost)
 *  - back: highest first (the lowest ends up bottommost)
 *  - forward: topmost first (so it is not blocked by a selected neighbour)
 *  - backward: bottommost first
 */
export function stackSorted(
  nodes: Array<{ sourceId: string }>,
  ids: string[],
  dir: ReorderDirection,
): string[] {
  const order = new Map(nodes.map((n, i) => [n.sourceId, i]))
  const z = (id: string) => order.get(id) ?? 0
  const highestFirst = dir === 'back' || dir === 'forward'
  return [...ids].sort((a, b) => (highestFirst ? z(b) - z(a) : z(a) - z(b)))
}

// ── Draw (freehand ink): one stroke = one transparent PNG picture element ────

export async function commitInk(ctx: ActionCtx, stroke: InkStroke): Promise<void> {
  if (!ctx.slide) return
  const raster = rasterizeStroke(stroke)
  const r = await window.slidesApi.addInk({
    slideIndex: ctx.current,
    base64: raster.base64,
    xPx: raster.xPx,
    yPx: raster.yPx,
    wPx: raster.wPx,
    hPx: raster.hPx,
    fitWidthPx: FIT_WIDTH,
    payload: raster.payload,
  })
  if (r) ctx.applySlide(ctx.current, r.slide)
}

/**
 * Mirror the selection across its own axis. Unlike align/distribute this is
 * not renderer geometry: flipH/flipV are a:xfrm attributes, so the main process toggles
 * them on the model.
 */
export async function flipSelected(ctx: ActionCtx, axis: 'h' | 'v'): Promise<void> {
  if (ctx.selectedIds.length === 0) return
  const groupId = ctx.groupIdOf(ctx.selectedIds[0]!)
  const updated = await window.slidesApi.flipElements({
    slideIndex: ctx.current,
    sourceIds: ctx.selectedIds,
    axis,
    ...(groupId ? { groupId } : {}),
  })
  if (updated) {
    ctx.applySlide(ctx.current, updated)
    ctx.setDirty(true)
  }
}

/**
 * Rotate each selected element by ±90° around its own visual center (PowerPoint
 * context-menu semantics), through the same transform-commit path as the rotate handle.
 * The model/render convention pivots rotation on the box origin (flip-adjusted corner),
 * so keeping the center fixed means moving the origin to its new orbit position — the
 * same compensation the Konva Transformer bakes into x/y during a handle rotation.
 * Connectors are skipped — their geometry is endpoint-based (no Transformer either).
 */
export async function rotateSelected(ctx: ActionCtx, deltaDeg: number): Promise<void> {
  const items: Array<{
    sourceId: string
    box: { x: number; y: number; w: number; h: number; rotationDeg: number }
    groupId?: string
  }> = []
  for (const id of ctx.selectedIds) {
    const found = ctx.findNodeCtx(id)
    if (!found) continue
    const { node, groupId } = found
    const isConnector =
      (node.type === 'shape' || node.type === 'text') && !!(node as { line?: unknown }).line
    if (isConnector) continue
    const b = node.box
    const deg = b.rotationDeg ?? 0
    // Konva transform order is translate(origin) → rotate → scale(flip): the local center
    // lands at origin + R(θ)·S·(w/2, h/2). Track where it goes under the old and new
    // angles and shift the origin by the difference (the flip part of the origin cancels).
    const vx = (b.flipH ? -1 : 1) * (b.w / 2)
    const vy = (b.flipV ? -1 : 1) * (b.h / 2)
    const orbit = (angleDeg: number) => {
      const t = (angleDeg * Math.PI) / 180
      return { x: vx * Math.cos(t) - vy * Math.sin(t), y: vx * Math.sin(t) + vy * Math.cos(t) }
    }
    const p0 = orbit(deg)
    const p1 = orbit(deg + deltaDeg)
    items.push({
      sourceId: id,
      box: {
        x: b.x + p0.x - p1.x,
        y: b.y + p0.y - p1.y,
        w: b.w,
        h: b.h,
        rotationDeg: (((deg + deltaDeg) % 360) + 360) % 360,
      },
      ...(groupId ? { groupId } : {}),
    })
  }
  if (!items.length) return
  // Top-level elements batch into one IPC so a multi-rotate is a single undo step;
  // in-group children keep the per-element path (the batch op has no group support)
  if (items.every((it) => !it.groupId)) {
    const updated = await window.slidesApi.batchEditTransform({
      slideIndex: ctx.current,
      fitWidthPx: FIT_WIDTH,
      items: items.map((it) => ({
        sourceId: it.sourceId,
        xPx: it.box.x,
        yPx: it.box.y,
        wPx: it.box.w,
        hPx: it.box.h,
        rotationDeg: it.box.rotationDeg,
      })),
    })
    if (updated) {
      ctx.applySlide(ctx.current, updated)
      ctx.setDirty(true)
    }
    return
  }
  // In-group children: onTransform batches whatever is requested in the same tick (group-aware)
  await Promise.all(items.map((it) => ctx.onTransform(it.sourceId, it.box, undefined, it.groupId)))
}

export async function eraseInk(ctx: ActionCtx, sourceIds: string[]): Promise<void> {
  // a clear wipes many strokes: still one undo step
  await withHistoryGroup(async () => {
    for (const id of sourceIds) {
      const updated = await window.slidesApi.deleteElement({
        slideIndex: ctx.current,
        sourceId: id,
      })
      if (updated) ctx.applySlide(ctx.current, updated)
    }
  })
}

export async function clearInk(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide) return
  await eraseInk(
    ctx,
    inkNodesOf(ctx.slide).map((e) => e.node.sourceId),
  )
  ctx.setStatus(t('appStatusInkCleared'))
}
