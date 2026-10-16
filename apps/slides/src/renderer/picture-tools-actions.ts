/**
 * Picture Format extras: Picture Styles, Alt Text and Picture Quality (compress).
 * Styles and alt text are native ops through the journaled transaction pipeline;
 * compression resamples in the renderer and swaps the media bytes in place.
 */
import type { PictureRenderNode } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import { t } from './i18n/locale'
import { PICTURE_STYLES, pictureStyleOps } from './picture-styles'

function selectedPicture(ctx: ActionCtx): PictureRenderNode | null {
  if (!ctx.slide || ctx.selectedIds.length !== 1) return null
  const node = ctx.slide.nodes.find((n) => n.sourceId === ctx.selectedIds[0])
  return node && node.type === 'picture' ? (node as PictureRenderNode) : null
}

export async function applyPictureStyle(ctx: ActionCtx, styleId: string): Promise<void> {
  const pic = selectedPicture(ctx)
  const style = PICTURE_STYLES.find((s) => s.id === styleId)
  if (!pic || !style || pic.media) return
  const r = await window.slidesApi.applyTxn({
    ops: pictureStyleOps(ctx.current, pic.sourceId, style),
  })
  if (r?.applied && r.slides) {
    ctx.setSlides(r.slides)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusPictureStyle'))
  } else {
    ctx.setStatus(r?.failures?.[0]?.error ?? t('appStatusPictureStyleFailed'))
  }
}

export async function savePictureAltText(
  ctx: ActionCtx,
  sourceId: string,
  text: string,
  decorative: boolean,
): Promise<boolean> {
  const r = await window.slidesApi.applyTxn({
    ops: [
      {
        op: 'setAltText',
        target: { slide: ctx.current, el: sourceId },
        text: decorative ? null : text,
        decorative,
      },
    ],
  })
  if (r?.applied && r.slides) {
    ctx.setSlides(r.slides)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusAltTextSaved'))
    return true
  }
  ctx.setStatus(r?.failures?.[0]?.error ?? t('appStatusAltTextFailed'))
  return false
}

export async function setPictureAspectLock(
  ctx: ActionCtx,
  sourceId: string,
  locked: boolean,
): Promise<void> {
  const r = await window.slidesApi.applyTxn({
    ops: [{ op: 'setLockAspect', target: { slide: ctx.current, el: sourceId }, locked }],
  })
  if (r?.applied && r.slides) {
    ctx.setSlides(r.slides)
    ctx.setDirty(true)
  }
}

export interface CompressPlan {
  /** source rectangle to draw (natural pixels) */
  sx: number
  sy: number
  sw: number
  sh: number
  /** output size */
  width: number
  height: number
}

/**
 * Work out the resample for a target resolution. `displayInches` is the size of
 * the picture's frame; with a crop and `deleteCropped` off, the full image spans
 * the frame scaled up by the crop. Returns null when the image is not larger
 * than the target (compressing would only lose quality).
 */
export function planCompression(
  natural: { w: number; h: number },
  srcRect: { l: number; t: number; r: number; b: number } | undefined,
  displayInches: { w: number; h: number },
  ppi: number,
  deleteCropped: boolean,
): CompressPlan | null {
  const crop = srcRect ?? { l: 0, t: 0, r: 0, b: 0 }
  const keepW = Math.max(0.001, 1 - crop.l - crop.r)
  const keepH = Math.max(0.001, 1 - crop.t - crop.b)
  const cropped = deleteCropped && (crop.l || crop.t || crop.r || crop.b)
  const region = cropped
    ? {
        sx: Math.round(crop.l * natural.w),
        sy: Math.round(crop.t * natural.h),
        sw: Math.max(1, Math.round(keepW * natural.w)),
        sh: Math.max(1, Math.round(keepH * natural.h)),
      }
    : { sx: 0, sy: 0, sw: natural.w, sh: natural.h }
  // inches spanned by the drawn region
  const regionInchesW = cropped ? displayInches.w : displayInches.w / keepW
  const regionInchesH = cropped ? displayInches.h : displayInches.h / keepH
  const width = Math.max(1, Math.round(regionInchesW * ppi))
  const height = Math.max(1, Math.round(regionInchesH * ppi))
  const shrinks = width < region.sw || height < region.sh
  if (!shrinks && !cropped) return null
  return {
    ...region,
    width: Math.min(width, region.sw),
    height: Math.min(height, region.sh),
  }
}

function loadImage(dataUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image()
    img.onload = () => resolve(img)
    img.onerror = () => reject(new Error('image load failed'))
    img.src = dataUrl
  })
}

export async function compressPicture(
  ctx: ActionCtx,
  opts: { ppi: number; deleteCropped: boolean },
): Promise<void> {
  const pic = selectedPicture(ctx)
  const slide = ctx.slide
  if (!pic || !slide || !pic.dataUrl || pic.media) return
  const mime = /^data:(image\/(?:png|jpeg));base64,/.exec(pic.dataUrl)?.[1]
  if (!mime) {
    ctx.setStatus(t('appStatusPictureCompressUnsupported'))
    return
  }
  try {
    const img = await loadImage(pic.dataUrl)
    const plan = planCompression(
      { w: img.naturalWidth, h: img.naturalHeight },
      pic.srcRect,
      { w: pic.box.w / slide.scale / 96, h: pic.box.h / slide.scale / 96 },
      opts.ppi,
      opts.deleteCropped,
    )
    if (!plan) {
      ctx.setStatus(t('appStatusPictureAlreadySmall'))
      return
    }
    const canvas = document.createElement('canvas')
    canvas.width = plan.width
    canvas.height = plan.height
    const g = canvas.getContext('2d')
    if (!g) return
    g.imageSmoothingQuality = 'high'
    g.drawImage(img, plan.sx, plan.sy, plan.sw, plan.sh, 0, 0, plan.width, plan.height)
    const out = canvas.toDataURL(mime, 0.88)
    const base64 = out.split(',')[1]
    const before = pic.dataUrl.length
    if (!base64 || out.length >= before) {
      ctx.setStatus(t('appStatusPictureAlreadySmall'))
      return
    }
    const updated = await window.slidesApi.replacePictureBytes({
      slideIndex: ctx.current,
      sourceId: pic.sourceId,
      base64,
      ext: mime === 'image/png' ? 'png' : 'jpg',
      // the crop survives only when the full image was resampled (same geometry)
      keepSrcRect: !(opts.deleteCropped && pic.srcRect),
    })
    if (!updated || 'error' in updated) {
      ctx.setStatus(t('appStatusPictureCompressUnsupported'))
      return
    }
    ctx.applySlide(ctx.current, updated)
    ctx.setSelectedIds([pic.sourceId])
    ctx.setDirty(true)
    ctx.setStatus(
      t('appStatusPictureCompressed', {
        before: Math.round(((before * 3) / 4 / 1024) * 10) / 10,
        after: Math.round(((out.length * 3) / 4 / 1024) * 10) / 10,
      }),
    )
  } catch {
    ctx.setStatus(t('appStatusPictureCompressUnsupported'))
  }
}
