/** Record / clear slide narration (see narration.ts for the storage model). */
import type { PictureRenderNode, RenderSlide } from '@genoffice/pptx-render'
import type { ActionCtx } from './action-context'
import { t } from './i18n/locale'
import {
  NARRATION_NAME,
  isNarrationRecording,
  startNarrationRecording,
  stopNarrationRecording,
} from './narration'

export function narrationNodes(slide: RenderSlide | undefined): PictureRenderNode[] {
  return (slide?.nodes ?? []).filter(
    (n): n is PictureRenderNode =>
      n.type === 'picture' &&
      n.media === 'audio' &&
      (n as PictureRenderNode).name === NARRATION_NAME,
  )
}

/** Start recording, or stop and store the narration on the current slide (replacing the old one). */
export async function toggleNarration(
  ctx: ActionCtx,
  setRecording: (on: boolean) => void,
): Promise<void> {
  if (!isNarrationRecording()) {
    if (await startNarrationRecording()) {
      setRecording(true)
      ctx.setStatus(t('appStatusNarrationRecording'))
    } else {
      ctx.setStatus(t('appStatusNarrationUnavailable'))
    }
    return
  }
  setRecording(false)
  const wav = await stopNarrationRecording()
  const slide = ctx.slide
  if (!wav || !slide) {
    ctx.setStatus(t('appStatusNarrationEmpty'))
    return
  }
  const emu = (px: number) => Math.round((px * 9525) / slide.scale)
  const size = 457200
  const target = { slide: ctx.current }
  const r = await window.slidesApi.applyTxn({
    ops: [
      ...narrationNodes(slide).map((n) => ({
        op: 'deleteElement',
        target: { slide: ctx.current, el: n.sourceId },
      })),
      {
        op: 'addMedia',
        target,
        kind: 'audio',
        bytes: wav,
        ext: 'wav',
        name: NARRATION_NAME,
        offset: {
          x: emu(slide.widthPx) - size - 182880,
          y: emu(slide.heightPx) - size - 182880,
          cx: size,
          cy: size,
        },
      },
    ],
  })
  if (r?.applied && r.slides) {
    ctx.setSlides(r.slides)
    ctx.setDirty(true)
    ctx.setStatus(t('appStatusNarrationSaved'))
  } else {
    ctx.setStatus(r?.failures?.[0]?.error ?? t('appStatusNarrationFailed'))
  }
}

export async function clearNarration(ctx: ActionCtx): Promise<void> {
  const nodes = narrationNodes(ctx.slide)
  if (!nodes.length) return
  const r = await window.slidesApi.applyTxn({
    ops: nodes.map((n) => ({
      op: 'deleteElement',
      target: { slide: ctx.current, el: n.sourceId },
    })),
  })
  if (r?.applied && r.slides) {
    ctx.setSlides(r.slides)
    ctx.setDirty(true)
    ctx.setSelectedIds([])
    ctx.setStatus(t('appStatusNarrationCleared'))
  }
}
