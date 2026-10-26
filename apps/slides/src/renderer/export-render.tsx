/**
 * Offscreen rendering for export — draws each RenderSlide page into an offscreen Konva Stage
 * and exports high-resolution PNGs. Reuses SlideThumb (same rendering as the main canvas and
 * thumbnails), guaranteeing the export matches what the editor shows.
 */
import React from 'react'
import { createRoot } from 'react-dom/client'
import type Konva from 'konva'
import type { RenderSlide } from '@genoffice/pptx-render'
import { SlideThumb } from './SlideThumb'
import { prepareEquationRasters } from './equation-raster'

/** Pixel ratio of the exported bitmap (2x hi-res, 1280 viewport width → 2560px PNG) */
const EXPORT_PIXEL_RATIO = 2

/**
 * Render each page to PNG base64 (without the data: prefix).
 * Reuse a single offscreen root page by page, grabbing each page as it's drawn, so the whole
 * deck never sits in memory at once.
 * pixelRatio 1 is enough for AI-vision screenshots (half the tokens of the 2x export default).
 */
export async function renderSlidesToPngBase64(
  slides: RenderSlide[],
  images: Map<string, HTMLImageElement>,
  pixelRatio: number = EXPORT_PIXEL_RATIO,
): Promise<string[]> {
  return renderPages(slides, images, pixelRatio)
}

async function renderPages(
  slides: RenderSlide[],
  images: Map<string, HTMLImageElement>,
  pixelRatio: number,
  crop?: { x: number; y: number; width: number; height: number },
): Promise<string[]> {
  // Offscreen container: mounted outside the body viewport (display:none would give the Konva canvas zero size, unusable)
  const container = document.createElement('div')
  container.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none;'
  document.body.appendChild(container)
  const root = createRoot(container)
  const out: string[] = []
  try {
    for (const slide of slides) {
      await prepareEquationRasters(slide)
      const stage = await new Promise<Konva.Stage>((resolve) => {
        root.render(
          <SlideThumb
            slide={slide}
            images={images}
            width={slide.widthPx}
            stageRef={(s) => s && resolve(s)}
          />,
        )
      })
      // Wait one frame for Konva to finish batchDraw, then capture
      await new Promise((r) => requestAnimationFrame(r))
      const dataUrl = stage.toDataURL({ mimeType: 'image/png', pixelRatio, ...(crop ?? {}) })
      out.push(dataUrl.replace(/^data:image\/png;base64,/, ''))
    }
  } finally {
    root.unmount()
    container.remove()
  }
  return out
}

/**
 * Render just the given elements (transparent background, cropped to their union with a small
 * margin for shadows) to PNG base64 — the picture rendition of a copy, for apps that cannot
 * take the editable shapes. Null when nothing renderable is selected.
 */
export async function renderElementsToPngBase64(
  slide: RenderSlide,
  sourceIds: readonly string[],
  images: Map<string, HTMLImageElement>,
  pixelRatio: number = EXPORT_PIXEL_RATIO,
): Promise<string | null> {
  const ids = new Set(sourceIds)
  const nodes = slide.nodes.filter((n) => ids.has(n.sourceId) && !n.decoration)
  if (!nodes.length) return null
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -Infinity
  let y1 = -Infinity
  for (const n of nodes) {
    const b = n.box
    // a rotated box's reach: its circumscribed square around the centre
    const r = b.rotationDeg ? Math.hypot(b.w, b.h) / 2 : 0
    const cx = b.x + b.w / 2
    const cy = b.y + b.h / 2
    x0 = Math.min(x0, r ? cx - r : b.x)
    y0 = Math.min(y0, r ? cy - r : b.y)
    x1 = Math.max(x1, r ? cx + r : b.x + b.w)
    y1 = Math.max(y1, r ? cy + r : b.y + b.h)
  }
  const margin = 10
  const x = Math.max(0, Math.floor(x0 - margin))
  const y = Math.max(0, Math.floor(y0 - margin))
  const w = Math.max(1, Math.min(slide.widthPx, Math.ceil(x1 + margin)) - x)
  const h = Math.max(1, Math.min(slide.heightPx, Math.ceil(y1 + margin)) - y)
  const only: RenderSlide = { ...slide, background: { kind: 'solid', color: '#FFFFFF00' }, nodes }
  const [png] = await renderPages([only], images, pixelRatio, { x, y, width: w, height: h })
  return png ?? null
}
