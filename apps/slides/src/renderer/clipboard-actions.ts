/**
 * Clipboard and format-painter actions extracted from App.tsx:
 * element copy/cut/paste/duplicate, slide copy/paste, external clipboard
 * ingestion, and the format painter. Functions take the ActionCtx built fresh per call.
 */
import type { ActionCtx } from './action-context'
import type { PasteSlideMode } from '../shared/ipc'
import { FIT_WIDTH } from './app-constants'
import { withHistoryGroup } from './history-group'
import { renderElementsToPngBase64, renderSlidesToPngBase64 } from './export-render'
import { copySelectedCells, pasteIntoCells, selectedTable } from './table-select-actions'
import {
  extractBrushFormat,
  computeBrushApply,
  buildEditParagraphsWithFormat,
} from './format-brush'
import { t } from './i18n/locale'

export async function deleteSelected(ctx: ActionCtx): Promise<void> {
  const ids = [...ctx.selectedIds]
  // deleteElement is per element: group the calls so Delete over a selection is one undo step
  const run = async () => {
    for (const id of ids) {
      const updated = await window.slidesApi.deleteElement({
        slideIndex: ctx.current,
        sourceId: id,
      })
      if (updated) ctx.applySlide(ctx.current, updated)
    }
  }
  if (ids.length > 1) await withHistoryGroup(run)
  else await run()
  ctx.setSelectedIds([])
}

/** A picture of the selection for apps that cannot take the editable shapes (copy still succeeds without it) */
async function selectionPng(ctx: ActionCtx): Promise<string | undefined> {
  try {
    if (!ctx.slide) return undefined
    return (await renderElementsToPngBase64(ctx.slide, ctx.selectedIds, ctx.images)) ?? undefined
  } catch {
    return undefined
  }
}

/** Cell-range selection inside a table: Copy / Cut / Paste act on the cells, not on the whole table */
function cellMode(ctx: ActionCtx): boolean {
  return !!ctx.cellSel && !!selectedTable(ctx) && !ctx.editingCell
}

export async function copySelected(ctx: ActionCtx): Promise<void> {
  if (cellMode(ctx)) {
    await copySelectedCells(ctx)
    return
  }
  if (ctx.selectedIds.length === 0) return
  const png = await selectionPng(ctx)
  const n = await window.slidesApi.copyElements({
    slideIndex: ctx.current,
    sourceIds: ctx.selectedIds,
    ...(png ? { png } : {}),
  })
  if (n > 0) {
    ctx.setHasClipboard(true)
    ctx.setStatus(t('appStatusCopied', { count: n }))
  }
}

export async function cutSelected(ctx: ActionCtx): Promise<void> {
  if (cellMode(ctx)) {
    await copySelectedCells(ctx, true)
    return
  }
  if (ctx.selectedIds.length === 0) return
  const png = await selectionPng(ctx)
  const n = await window.slidesApi.copyElements({
    slideIndex: ctx.current,
    sourceIds: ctx.selectedIds,
    cut: true,
    ...(png ? { png } : {}),
  })
  if (n > 0) {
    ctx.setHasClipboard(true)
    await deleteSelected(ctx)
    ctx.setStatus(t('appStatusCut', { count: n }))
  }
}

/** Insert an external image at page center at natural size (clamped to half the page). */
export async function insertExternalImage(
  ctx: ActionCtx,
  base64: string,
  ext: string,
  atPx?: { x: number; y: number },
): Promise<void> {
  const dims = await new Promise<{ w: number; h: number }>((resolve) => {
    const img = new Image()
    img.onload = () => resolve({ w: img.naturalWidth || 320, h: img.naturalHeight || 240 })
    img.onerror = () => resolve({ w: 320, h: 240 })
    img.src = `data:image/${ext};base64,${base64}`
  })
  const maxW = (ctx.slide?.widthPx ?? FIT_WIDTH) / 2
  const maxH = (ctx.slide?.heightPx ?? FIT_WIDTH * 0.5625) / 2
  const k = Math.min(1, maxW / dims.w, maxH / dims.h)
  const w = Math.max(24, dims.w * k)
  const h = Math.max(24, dims.h * k)
  const x = atPx ? atPx.x - w / 2 : ((ctx.slide?.widthPx ?? FIT_WIDTH) - w) / 2
  const y = atPx ? atPx.y - h / 2 : ((ctx.slide?.heightPx ?? FIT_WIDTH * 0.5625) - h) / 2
  const r = await window.slidesApi.addImageBytes({
    slideIndex: ctx.current,
    base64,
    ext,
    xPx: x,
    yPx: y,
    wPx: w,
    hPx: h,
    fitWidthPx: FIT_WIDTH,
  })
  if (!r) return
  if ('error' in r) {
    ctx.setSelectedIds([])
    ctx.setStatus(t('appStatusImageUnsupported', { ext: r.ext.toUpperCase() }))
    return
  }
  ctx.applySlide(ctx.current, r.slide)
  ctx.setSelectedIds([r.sourceId])
}

export async function copySlideAt(ctx: ActionCtx, index: number): Promise<void> {
  // A rendering of the page rides along for 'picture'-mode pastes (the source
  // deck may be gone by paste time). Copy still succeeds if the capture fails.
  let png: string | undefined
  try {
    const slide = ctx.slides[index]
    if (slide) [png] = await renderSlidesToPngBase64([slide], ctx.images)
  } catch {
    png = undefined
  }
  const ok = await window.slidesApi.copySlide(index, png)
  ctx.setCanPasteSlide(ok)
  ctx.setStatus(ok ? t('appStatusSlideCopied') : t('appStatusSlideCopyFailed'))
}

export async function pasteSlideAfter(
  ctx: ActionCtx,
  index: number,
  mode: PasteSlideMode = 'theme',
): Promise<void> {
  const r = await window.slidesApi.pasteSlide({ afterIndex: index, fitWidthPx: FIT_WIDTH, mode })
  if (!r) {
    ctx.setStatus(t('appStatusSlidePasteFailed'))
    return
  }
  ctx.setSlides(r.slides)
  ctx.setCurrent(r.index)
  ctx.setSelectedIds(r.sourceId ? [r.sourceId] : [])
  ctx.setEditing(null)
  ctx.setDirty(true)
  ctx.setPasteFloater({ index: r.index, mode })
}

/** Paste-options floater: redo the just-completed paste with another mode. */
export async function repasteSlideAs(ctx: ActionCtx, mode: PasteSlideMode): Promise<void> {
  if (ctx.pasteFloater?.mode === mode) return
  const r = await window.slidesApi.repasteSlide({ mode, fitWidthPx: FIT_WIDTH })
  if (!r) {
    ctx.setPasteFloater(null)
    ctx.setStatus(t('appStatusPasteOptionsExpired'))
    return
  }
  ctx.setSlides(r.slides)
  ctx.setCurrent(r.index)
  ctx.setSelectedIds(r.sourceId ? [r.sourceId] : [])
  ctx.setEditing(null)
  ctx.setDirty(true)
  ctx.setPasteFloater({ index: r.index, mode })
}

/**
 * Paste: whole slide (slide marker newest) → app elements (element marker not
 * overwritten externally) → external images → external text into a text box
 */
export async function pasteClipboard(ctx: ActionCtx): Promise<void> {
  if (cellMode(ctx)) {
    await pasteIntoCells(ctx)
    return
  }
  let external = await window.slidesApi.clipboardExternal()
  if (external.kind === 'native') {
    // PowerPoint shapes / a table / an HTML table from another app: editable elements, centred like PowerPoint
    const r = await window.slidesApi.pasteSystem({ slideIndex: ctx.current, fitWidthPx: FIT_WIDTH })
    if (r) {
      ctx.applySlide(ctx.current, r.slide)
      ctx.setSelectedIds(r.sourceIds)
      ctx.setEditing(null)
      return
    }
    // unreadable native data: fall back to its picture / text rendition
    external = await window.slidesApi.clipboardExternal(true)
  }
  if (external.kind === 'slide') {
    await pasteSlideAfter(ctx, ctx.current)
    return
  }
  if (external.kind === 'image') {
    await insertExternalImage(ctx, external.base64, external.ext)
    return
  }
  if (external.kind === 'text') {
    const w = 400
    const r = await window.slidesApi.addElement({
      slideIndex: ctx.current,
      kind: 'textbox',
      xPx: ((ctx.slide?.widthPx ?? FIT_WIDTH) - w) / 2,
      yPx: (ctx.slide?.heightPx ?? FIT_WIDTH * 0.5625) / 2 - 40,
      wPx: w,
      hPx: 80,
      fitWidthPx: FIT_WIDTH,
      paragraphs: external.text.split(/\r?\n/).map((line) => ({ runs: [{ text: line }] })),
    })
    if (r) {
      ctx.applySlide(ctx.current, r.slide)
      ctx.setSelectedIds([r.sourceId])
    }
    return
  }
  const r = await window.slidesApi.pasteElements({ slideIndex: ctx.current, fitWidthPx: FIT_WIDTH })
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds(r.sourceIds)
    ctx.setEditing(null)
  }
}

/** Duplicate in place (⌘D / Option+drag copy): doesn't touch the app clipboard, the copy becomes selected. */
export async function duplicateSelected(
  ctx: ActionCtx,
  ids: string[],
  dxPx = 16,
  dyPx = 16,
): Promise<void> {
  if (!ids.length) return
  const r = await window.slidesApi.duplicateElements({
    slideIndex: ctx.current,
    sourceIds: ids,
    dxPx,
    dyPx,
    fitWidthPx: FIT_WIDTH,
  })
  if (r) {
    ctx.applySlide(ctx.current, r.slide)
    ctx.setSelectedIds(r.sourceIds)
  }
}

// ── Format painter operations ───────────────────────────────────────────────

/** ⌘⇧C: copy the selected single element's format to the clipboard (without entering paste mode) */
export function copyFormat(ctx: ActionCtx): void {
  if (ctx.selectedIds.length !== 1 || !ctx.slide) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  if (!node) return
  const fmt = extractBrushFormat(node)
  if (fmt) {
    ctx.setBrushFormat(fmt)
    ctx.setStatus(t('appStatusFormatCopied'))
  }
}

/** ⌘⇧V: paste format onto selected elements (format must have been copied first) */
export async function pasteFormat(ctx: ActionCtx): Promise<void> {
  if (!ctx.brushFormat || ctx.selectedIds.length === 0 || !ctx.slide) return
  // fill + stroke + text are separate calls per element: one paste = one undo step
  await withHistoryGroup(() => pasteFormatEach(ctx))
  ctx.setStatus(t('appStatusFormatPasted'))
}

async function pasteFormatEach(ctx: ActionCtx): Promise<void> {
  if (!ctx.brushFormat || !ctx.slide) return
  for (const id of ctx.selectedIds) {
    const node = ctx.slide.nodes.find((n) => n.sourceId === id)
    if (!node) continue
    await applyBrushFormatTo(ctx, ctx.brushFormat, id, node)
  }
}

/**
 * Apply a copied format to one element. Only the parts the plan carries are written: an outline
 * (or text) only paste must leave the target's existing fill untouched.
 */
async function applyBrushFormatTo(
  ctx: ActionCtx,
  brush: NonNullable<ActionCtx['brushFormat']>,
  targetId: string,
  node: NonNullable<ActionCtx['slide']>['nodes'][number],
): Promise<void> {
  const plan = computeBrushApply(brush, node)
  if (plan.fillColor !== undefined) {
    const updated = await window.slidesApi.editFill({
      slideIndex: ctx.current,
      sourceId: targetId,
      fill: plan.fillColor,
    })
    if (updated) ctx.applySlide(ctx.current, updated)
  }
  if (plan.stroke !== undefined) {
    const updated = await window.slidesApi.editStroke({
      slideIndex: ctx.current,
      sourceId: targetId,
      stroke: plan.stroke,
    })
    if (updated) ctx.applySlide(ctx.current, updated)
  }
  if ((plan.runFormat || plan.align) && (node.type === 'shape' || node.type === 'text')) {
    const paras = buildEditParagraphsWithFormat(node, plan.runFormat, plan.align)
    if (paras) {
      const updated = await window.slidesApi.editText({
        slideIndex: ctx.current,
        sourceId: targetId,
        paragraphs: paras,
      })
      if (updated) ctx.applySlide(ctx.current, updated)
    }
  }
}

/** Format painter button single click: copy current format + enter single-paste mode */
export function onFormatBrushClick(ctx: ActionCtx): void {
  if (ctx.selectedIds.length !== 1 || !ctx.slide) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  if (!node) return
  const fmt = extractBrushFormat(node)
  if (!fmt) return
  ctx.setBrushFormat(fmt)
  ctx.setBrushMode('once')
  ctx.setStatus(t('appStatusBrushOnce'))
}

/** Format painter button double click: enter continuous mode */
export function onFormatBrushDoubleClick(ctx: ActionCtx): void {
  if (ctx.selectedIds.length !== 1 || !ctx.slide) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  if (!node) return
  const fmt = extractBrushFormat(node)
  if (!fmt) return
  ctx.setBrushFormat(fmt)
  ctx.setBrushMode('continuous')
  ctx.setStatus(t('appStatusBrushContinuous'))
}

/** Apply format when an element is clicked in format painter mode */
export async function applyBrushToElement(ctx: ActionCtx, targetId: string): Promise<void> {
  const brush = ctx.brushFormat
  if (!brush || !ctx.slide) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === targetId)
  if (!node) return
  // fill / outline / text are separate calls: one brush click = one undo step
  await withHistoryGroup(() => applyBrushFormatTo(ctx, brush, targetId, node))

  if (ctx.brushMode === 'once') {
    ctx.setBrushMode(null)
    ctx.setStatus(t('appStatusFormatApplied'))
  }
}
