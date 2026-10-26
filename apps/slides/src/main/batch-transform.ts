/**
 * Geometry transforms for one or many elements, shared by the single-element
 * `slides:edit-transform` handler and the multi-element `slides:batch-edit-transform`
 * one. Kept out of slides-main so the history semantics (one gesture / one keypress over a
 * whole selection = ONE undo step) can be unit-tested without Electron.
 */
import { resolveGroupChildId, runTxn, type Op } from '@genoffice/pptx-ops'
import { findGroupChild } from '@genoffice/pptx-engine'
import { EMU_PER_PX_96 } from '@genoffice/pptx-render'
import type { BatchEditTransformOp } from '../shared/ipc'
import { journalOps, pushHistory, scheduleDeckBroadcast, type Session } from './session-state'

/** One element's target box in viewport px (the shape both IPC handlers receive). */
export interface TransformItem {
  sourceId: string
  /** In-group editing: sourceId is a direct child of that group; the box is in group-local px */
  groupId?: string
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  rotationDeg: number
  /** Tables redistribute gridCol widths / tr heights so the file matches the new frame */
  resizeTableGrid?: boolean
  /**
   * Tables only: the request is a pure move. The size that came with it is what was DRAWN (rows grow to
   * fit their text, so it is larger than the stored grid); re-applying it as a resize would inflate the
   * stored row heights a little more on every move.
   */
    keepTableSize?: boolean
  tableRowHeightsPx?: number[]
}

/**
 * px → EMU translation into a canonical `setTransform` op (including the group-local scale
 * inversion for children of an entered group). Null when a named group child cannot be found.
 */
export function transformPayload(
  session: Session,
  slideIndex: number,
  fitWidthPx: number,
  item: TransformItem,
): Op | null {
  const slide = session.opened.deck.slides[slideIndex]
  if (!slide) return null
  const childId = item.groupId
    ? resolveGroupChildId(slide, item.groupId, item.sourceId)
    : item.sourceId
  const grpChild = item.groupId ? findGroupChild(slide, item.groupId, childId) : null
  if (item.groupId && !grpChild) return null
  const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
  const scale = fitWidthPx / baseWidthPx
  const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
  let box: { x: number; y: number; cx: number; cy: number }
  if (grpChild) {
    // In-group editing: the pixel box is in group-local coords (with ext/chExt scaling baked in);
    // divide out the group scale first, then convert back to the child EMU coordinate system
    const ch = grpChild.grp.childOffset
    const chX = ch?.x ?? grpChild.grp.transform.offset.x
    const chY = ch?.y ?? grpChild.grp.transform.offset.y
    const gExt = grpChild.grp.transform.offset
    const gsx = ch?.cx ? gExt.cx / ch.cx : 1
    const gsy = ch?.cy ? gExt.cy / ch.cy : 1
    box = {
      x: toEmu(item.xPx / gsx) + chX,
      y: toEmu(item.yPx / gsy) + chY,
      cx: toEmu(item.wPx / gsx),
      cy: toEmu(item.hPx / gsy),
    }
  } else {
    box = { x: toEmu(item.xPx), y: toEmu(item.yPx), cx: toEmu(item.wPx), cy: toEmu(item.hPx) }
  }
  const stored =
    item.keepTableSize && !item.groupId ? slide.elements.find((e) => e.id === item.sourceId) : null
  if (stored?.type === 'table') {
    box = { ...box, cx: stored.transform.offset.cx, cy: stored.transform.offset.cy }
  }
  const rowHeightsEmu = item.tableRowHeightsPx?.map((px) => toEmu(px))
  return {
    op: 'setTransform',
    target: { slide: slideIndex, el: item.sourceId },
    box,
    rotDeg: item.rotationDeg,
    ...(item.groupId
      ? { group: item.groupId }
      : item.resizeTableGrid && !(item.keepTableSize && stored?.type === 'table')
        ? {
            resizeTableGrid: true,
            ...(rowHeightsEmu ? { tableRowHeightsEmu: rowHeightsEmu } : {}),
          }
        : {}),
  }
}

/**
 * Apply every item as ONE atomic transaction and ONE undo step. Returns false (history and
 * model untouched) when any item is invalid.
 *
 * Gesture interplay: a Transformer resize of several elements streams live `preview` frames
 * through edit-transform (the first pushes the pre-gesture snapshot and sets
 * `transformPreview`); the final commit then arrives as this batch and must NOT push a second
 * snapshot — it only clears the flag, exactly like the single-element commit.
 */
export function applyTransformBatch(session: Session, op: BatchEditTransformOp): boolean {
  const ops: Op[] = []
  for (const item of op.items) {
    const payload = transformPayload(session, op.slideIndex, op.fitWidthPx, item)
    if (!payload) return false
    ops.push(payload)
  }
  if (!ops.length) return false
  const req = { ops }
  // Plan first: pushHistory clears the redo stack, so an invalid request must not touch it
  if (runTxn(session.opened, { ...req, dryRun: true }).failures?.length) return false
  let pushed = false
  if (session.transformPreview) {
    session.transformPreview = false
  } else {
    pushHistory(session)
    pushed = true
  }
  const r = runTxn(session.opened, req)
  if (!r.applied) {
    if (pushed) session.undoStack.pop()
    return false
  }
  journalOps(session, 'edit', r.records ?? [])
  scheduleDeckBroadcast(session)
  return true
}
