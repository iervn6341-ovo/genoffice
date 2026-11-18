/**
 * Picture crop and cutout (background removal) actions. Extracted from
 * App.tsx; functions read the latest App state through ActionCtx.
 */
import type { PictureRenderNode } from '@genoffice/pptx-render'
import { FIT_WIDTH } from './app-constants'
import type { ActionCtx } from './action-context'
import { t } from './i18n/locale'

/** Enter picture crop mode: find the selected picture node and read its box and srcRect */
export function startCrop(ctx: ActionCtx): void {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const id = ctx.selectedIds[0]!
  const node = ctx.slide.nodes.find((n) => n.sourceId === id)
  if (!node || node.type !== 'picture') return
  const picNode = node as PictureRenderNode
  // Reconstruct the full original-image extent from the displayed (cropped) box:
  // the crop is non-destructive (srcRect), so the frame may expand back out to it.
  const sr = picNode.srcRect
  const kw = sr ? Math.max(1 - sr.l - sr.r, 0.01) : 1
  const kh = sr ? Math.max(1 - sr.t - sr.b, 0.01) : 1
  const fullW = picNode.box.w / kw
  const fullH = picNode.box.h / kh
  ctx.setCropTarget({
    sourceId: id,
    box: { x: picNode.box.x, y: picNode.box.y, w: picNode.box.w, h: picNode.box.h },
    srcRect: sr,
    fullBox: {
      x: picNode.box.x - (sr?.l ?? 0) * fullW,
      y: picNode.box.y - (sr?.t ?? 0) * fullH,
      w: fullW,
      h: fullH,
    },
    dataUrl: picNode.dataUrl,
  })
}

export async function commitCrop(
  ctx: ActionCtx,
  rect: { l: number; t: number; r: number; b: number } | null,
): Promise<void> {
  if (!ctx.cropTarget) return
  const { sourceId, srcRect: prev, fullBox } = ctx.cropTarget
  ctx.setCropTarget(null)
  // `rect` is relative to the FULL original image (the frame may have been dragged
  // outward past the previous crop); null = frame covers the whole original.
  if (!rect && !prev) return // never cropped and frame left at full — nothing to do
  // The element frame moves/resizes to the on-screen crop frame in the same atomic
  // edit: the kept region stays exactly where it was framed (PowerPoint semantics),
  // dragging outward restores original content, and one undo reverts everything.
  const updated = await window.slidesApi.editPictureSrcRect({
    slideIndex: ctx.current,
    sourceId,
    srcRect: rect,
    boxPx: rect
      ? {
          x: fullBox.x + rect.l * fullBox.w,
          y: fullBox.y + rect.t * fullBox.h,
          w: Math.max(1, fullBox.w * (1 - rect.l - rect.r)),
          h: Math.max(1, fullBox.h * (1 - rect.t - rect.b)),
        }
      : { x: fullBox.x, y: fullBox.y, w: fullBox.w, h: fullBox.h }, // full frame = original restored
    fitWidthPx: FIT_WIDTH,
  })
  if (updated) {
    ctx.applySlide(ctx.current, updated)
    ctx.setDirty(true)
    ctx.setStatus(rect ? t('appStatusCropApplied') : t('appStatusCropRemoved'))
  }
}

export interface CropRect {
  l: number
  t: number
  r: number
  b: number
}

/** Reconstruct the full original image bounding box from displayed box and srcRect */
export function reconstructFullBox(
  box: { x: number; y: number; w: number; h: number },
  sr?: CropRect | null,
): { x: number; y: number; w: number; h: number } {
  const kw = sr ? Math.max(1 - sr.l - sr.r, 0.01) : 1
  const kh = sr ? Math.max(1 - sr.t - sr.b, 0.01) : 1
  const fullW = box.w / kw
  const fullH = box.h / kh
  return {
    x: box.x - (sr?.l ?? 0) * fullW,
    y: box.y - (sr?.t ?? 0) * fullH,
    w: fullW,
    h: fullH,
  }
}

/** Compute center-crop fractions and fitted box for an aspect ratio preset */
export function computeAspectRatioCrop(
  box: { x: number; y: number; w: number; h: number },
  fullBox: { x: number; y: number; w: number; h: number },
  aspectW: number,
  aspectH: number,
): { srcRect: CropRect; boxPx: { x: number; y: number; w: number; h: number } } {
  const targetAspect = aspectW / aspectH
  const fullAspect = fullBox.w / fullBox.h

  let srcRect: CropRect
  if (fullAspect > targetAspect) {
    const fraction = targetAspect / fullAspect
    const crop = (1 - fraction) / 2
    srcRect = { l: crop, t: 0, r: crop, b: 0 }
  } else {
    const fraction = fullAspect / targetAspect
    const crop = (1 - fraction) / 2
    srcRect = { l: 0, t: crop, r: 0, b: crop }
  }

  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  let nw = box.w
  let nh = box.w / targetAspect
  if (nh > box.h) {
    nh = box.h
    nw = box.h * targetAspect
  }
  const boxPx = {
    x: cx - nw / 2,
    y: cy - nh / 2,
    w: Math.max(1, nw),
    h: Math.max(1, nh),
  }
  return { srcRect, boxPx }
}

/** Compute Fill crop: centers the image so it covers the existing frame without distortion */
export function computeFillCrop(
  box: { x: number; y: number; w: number; h: number },
  fullBox: { x: number; y: number; w: number; h: number },
): { srcRect: CropRect | null; boxPx: { x: number; y: number; w: number; h: number } } {
  const frameAspect = box.w / box.h
  const fullAspect = fullBox.w / fullBox.h

  let srcRect: CropRect | null = null
  if (Math.abs(frameAspect - fullAspect) > 0.001) {
    if (fullAspect > frameAspect) {
      const fraction = frameAspect / fullAspect
      const crop = (1 - fraction) / 2
      srcRect = { l: crop, t: 0, r: crop, b: crop }
    } else {
      const fraction = fullAspect / frameAspect
      const crop = (1 - fraction) / 2
      srcRect = { l: 0, t: crop, r: 0, b: crop }
    }
  }
  return { srcRect, boxPx: { ...box } }
}

/** Compute Fit: fits the uncropped image entirely inside the current frame with original proportions */
export function computeFitCrop(
  box: { x: number; y: number; w: number; h: number },
  fullBox: { x: number; y: number; w: number; h: number },
): { srcRect: null; boxPx: { x: number; y: number; w: number; h: number } } {
  const fullAspect = fullBox.w / fullBox.h
  let nw = box.w
  let nh = box.w / fullAspect
  if (nh > box.h) {
    nh = box.h
    nw = box.h * fullAspect
  }
  const cx = box.x + box.w / 2
  const cy = box.y + box.h / 2
  return {
    srcRect: null,
    boxPx: {
      x: cx - nw / 2,
      y: cy - nh / 2,
      w: Math.max(1, nw),
      h: Math.max(1, nh),
    },
  }
}

/** Cancel crop mode: dismisses the overlay without applying pending adjustments */
export function cancelCrop(ctx: ActionCtx): void {
  ctx.setCropTarget(null)
}

/** Crop the selected picture to a specific aspect ratio */
export async function cropToAspectRatio(
  ctx: ActionCtx,
  aspectW: number,
  aspectH: number,
): Promise<void> {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const id = ctx.selectedIds[0]!
  const node = ctx.slide.nodes.find((n) => n.sourceId === id)
  if (!node || node.type !== 'picture') return
  const pic = node as PictureRenderNode

  if (ctx.cropTarget) ctx.setCropTarget(null)

  const fullBox = reconstructFullBox(pic.box, pic.srcRect)
  const { srcRect, boxPx } = computeAspectRatioCrop(pic.box, fullBox, aspectW, aspectH)

  const updated = await window.slidesApi.editPictureSrcRect({
    slideIndex: ctx.current,
    sourceId: id,
    srcRect,
    boxPx,
    fitWidthPx: FIT_WIDTH,
  })
  if (updated) {
    ctx.applySlide(ctx.current, updated)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusCropApplied'))
  }
}

/** Fill the picture frame (object-fit: cover) */
export async function cropFill(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const id = ctx.selectedIds[0]!
  const node = ctx.slide.nodes.find((n) => n.sourceId === id)
  if (!node || node.type !== 'picture') return
  const pic = node as PictureRenderNode

  if (ctx.cropTarget) ctx.setCropTarget(null)

  const fullBox = reconstructFullBox(pic.box, pic.srcRect)
  const { srcRect, boxPx } = computeFillCrop(pic.box, fullBox)

  const updated = await window.slidesApi.editPictureSrcRect({
    slideIndex: ctx.current,
    sourceId: id,
    srcRect,
    boxPx,
    fitWidthPx: FIT_WIDTH,
  })
  if (updated) {
    ctx.applySlide(ctx.current, updated)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusCropApplied'))
  }
}

/** Fit the picture inside its frame (object-fit: contain) */
export async function cropFit(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const id = ctx.selectedIds[0]!
  const node = ctx.slide.nodes.find((n) => n.sourceId === id)
  if (!node || node.type !== 'picture') return
  const pic = node as PictureRenderNode

  if (ctx.cropTarget) ctx.setCropTarget(null)

  const fullBox = reconstructFullBox(pic.box, pic.srcRect)
  const { srcRect, boxPx } = computeFitCrop(pic.box, fullBox)

  const updated = await window.slidesApi.editPictureSrcRect({
    slideIndex: ctx.current,
    sourceId: id,
    srcRect,
    boxPx,
    fitWidthPx: FIT_WIDTH,
  })
  if (updated) {
    ctx.applySlide(ctx.current, updated)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusCropApplied'))
  }
}

/** Reset crop: remove srcRect and restore full original picture extent */
export async function resetCrop(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const id = ctx.selectedIds[0]!
  const node = ctx.slide.nodes.find((n) => n.sourceId === id)
  if (!node || node.type !== 'picture') return
  const pic = node as PictureRenderNode

  if (ctx.cropTarget) ctx.setCropTarget(null)

  const fullBox = reconstructFullBox(pic.box, pic.srcRect)

  const updated = await window.slidesApi.editPictureSrcRect({
    slideIndex: ctx.current,
    sourceId: id,
    srcRect: null,
    boxPx: fullBox,
    fitWidthPx: FIT_WIDTH,
  })
  if (updated) {
    ctx.applySlide(ctx.current, updated)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusCropRemoved'))
  }
}

// ── Picture cutout (background removal) ───────────────────────────────────────

/** Enter cutout mode: a single selected picture (except audio/video poster frames) → open the tolerance preview dialog */
export function startCutout(ctx: ActionCtx): void {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  if (!node || node.type !== 'picture') return
  const pic = node as PictureRenderNode
  if (pic.media) {
    ctx.setStatus(t('appStatusCutoutMediaUnsupported'))
    return
  }
  if (!pic.dataUrl) {
    ctx.setStatus(t('appStatusCutoutNoData'))
    return
  }
  ctx.setCutoutTarget({ sourceId: pic.sourceId, dataUrl: pic.dataUrl })
}

/**
 * Apply the cutout result: swap the picture's backing image for the
 * background-removed PNG in place. One atomic IPC — frame, rotation, z-order,
 * border and effects all survive; the crop window is kept because the result
 * PNG shares the source image's pixel geometry.
 */
export async function applyCutout(ctx: ActionCtx, pngDataUrl: string): Promise<void> {
  if (!ctx.cutoutTarget || !ctx.slide) return
  const targetId = ctx.cutoutTarget.sourceId
  const node = ctx.slide.nodes.find((n) => n.sourceId === targetId)
  ctx.setCutoutTarget(null)
  if (!node || node.type !== 'picture') return
  const base64 = pngDataUrl.split(',')[1]
  if (!base64) {
    ctx.setStatus(t('appStatusCutoutEncodeFailed'))
    return
  }
  const updated = await window.slidesApi.replacePictureBytes({
    slideIndex: ctx.current,
    sourceId: targetId,
    base64,
    ext: 'png',
    keepSrcRect: true,
  })
  if (!updated || 'error' in updated) {
    ctx.setStatus(t('appStatusCutoutInsertFailed'))
    return
  }
  ctx.applySlide(ctx.current, updated)
  ctx.setSelectedIds([targetId])
  ctx.setDirty(true)
  ctx.setStatus(t('appStatusCutoutDone'))
}

/** Replace the selected picture's image with a file from disk; frame, z-order and effects survive */
export async function replacePicture(ctx: ActionCtx): Promise<void> {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return
  const targetId = ctx.selectedIds[0]!
  const node = ctx.slide.nodes.find((n) => n.sourceId === targetId)
  if (!node || node.type !== 'picture') return
  const picked = await window.slidesApi.pickPictureFile()
  if (!picked) return
  const updated = await window.slidesApi.replacePictureBytes({
    slideIndex: ctx.current,
    sourceId: targetId,
    base64: picked.base64,
    ext: picked.ext,
  })
  if (!updated) return
  if ('error' in updated) {
    ctx.setStatus(t('appStatusImageUnsupported', { ext: updated.ext }))
    return
  }
  ctx.applySlide(ctx.current, updated)
  ctx.setSelectedIds([targetId])
  ctx.setDirty(true)
}
