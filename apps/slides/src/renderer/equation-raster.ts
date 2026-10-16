import type { RenderNode, RenderSlide, ShapeRenderNode } from '@genoffice/pptx-render'
import { normalizeColor, type GlyphDraw } from './konva-adapter'

export interface EquationRasterSpec {
  mathML: string
  width: number
  height: number
  fontSize: number
  color: string
}

type CachedRaster = { image: HTMLImageElement | null; ready: Promise<void> }
const cache = new Map<string, CachedRaster>()
const listeners = new Set<() => void>()
let epoch = 0
let fontsObserved = false

export function subscribeEquationRasters(listener: () => void): () => void {
  listeners.add(listener)
  return () => listeners.delete(listener)
}

export function getEquationRasterEpoch(): number {
  return epoch
}

function notify(): void {
  epoch += 1
  for (const listener of listeners) listener()
}

function escapeAttr(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
}

/** One SVG image, with no external resources, can be drawn by Konva and exported untainted. */
export function equationSvg(
  spec: EquationRasterSpec,
  measured?: { width: number; height: number },
): string {
  const width = Math.max(1, spec.width)
  const height = Math.max(1, spec.height)
  const naturalW = Math.max(1, measured?.width || width)
  const naturalH = Math.max(1, measured?.height || height)
  const scale = Math.min(1, width / naturalW, height / naturalH)
  const left = (width - naturalW * scale) / 2
  const top = (height - naturalH * scale) / 2
  const math = spec.mathML
    .replace(/<math\b/g, '<math xmlns="http://www.w3.org/1998/Math/MathML"')
    .replace(/\sdisplay="block"/g, ' display="inline"')
  const style = `position:absolute;left:${left}px;top:${top}px;transform:scale(${scale});transform-origin:top left;font-family:'Cambria Math',serif;font-size:${Math.max(1, spec.fontSize)}px;color:${spec.color};white-space:nowrap;line-height:normal;`
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${Math.ceil(width * 2)}" height="${Math.ceil(height * 2)}" viewBox="0 0 ${width} ${height}"><foreignObject width="${width}" height="${height}"><div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:${width}px;height:${height}px;overflow:hidden;"><div style="${escapeAttr(style)}">${math}</div></div></foreignObject></svg>`
}

function measureEquation(spec: EquationRasterSpec): { width: number; height: number } {
  const element = document.createElement('div')
  element.style.cssText = `position:fixed;left:-100000px;top:0;display:inline-block;width:max-content;font-family:'Cambria Math',serif;font-size:${Math.max(1, spec.fontSize)}px;line-height:normal;white-space:nowrap;`
  element.innerHTML = spec.mathML.replace(/\sdisplay="block"/g, ' display="inline"')
  document.body.appendChild(element)
  const rect = element.getBoundingClientRect()
  element.remove()
  return { width: rect.width || spec.width, height: rect.height || spec.height }
}

function keyOf(spec: EquationRasterSpec): string {
  return JSON.stringify(spec)
}

export function ensureEquationRaster(spec: EquationRasterSpec): Promise<void> {
  if (!fontsObserved && document.fonts?.addEventListener) {
    fontsObserved = true
    document.fonts.addEventListener('loadingdone', () => {
      cache.clear()
      notify()
    })
  }
  const key = keyOf(spec)
  const existing = cache.get(key)
  if (existing) return existing.ready
  const entry: CachedRaster = { image: null, ready: Promise.resolve() }
  cache.set(key, entry)
  entry.ready = new Promise<void>((resolve, reject) => {
    const svg = equationSvg(spec, measureEquation(spec))
    const image = new Image()
    image.onload = () => {
      entry.image = image
      notify()
      resolve()
    }
    image.onerror = () => {
      cache.delete(key)
      reject(new Error('Unable to rasterize native equation MathML'))
    }
    image.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`
  })
  return entry.ready
}

export function equationRaster(spec: EquationRasterSpec): HTMLImageElement | null {
  const entry = cache.get(keyOf(spec))
  if (!entry) void ensureEquationRaster(spec).catch(() => {})
  return entry?.image ?? null
}

export function equationRasterForGlyph(glyph: GlyphDraw): HTMLImageElement | null {
  if (!glyph.mathML) return null
  return equationRaster({
    mathML: glyph.mathML,
    width: glyph.mathWidthPx ?? glyph.fontSize,
    height: glyph.mathHeightPx ?? glyph.fontSize,
    fontSize: glyph.fontSize,
    color: glyph.fill,
  })
}

function specsInText(text: ShapeRenderNode['text']): EquationRasterSpec[] {
  return (text?.lines ?? []).flatMap((line) =>
    line.runs.flatMap((run) =>
      run.mathML
        ? [
            {
              mathML: run.mathML,
              width: run.widthPx,
              height:
                (run.mathAscentPx ?? run.fontSizePx * 0.8) +
                (run.mathDescentPx ?? run.fontSizePx * 0.2),
              fontSize: run.fontSizePx,
              color: normalizeColor(run.color),
            },
          ]
        : [],
    ),
  )
}

function specsInNode(node: RenderNode): EquationRasterSpec[] {
  if (node.type === 'group') return node.children.flatMap(specsInNode)
  if (node.type === 'table') return node.cells.flatMap((cell) => specsInText(cell.text))
  if (node.type !== 'shape' && node.type !== 'text') return []
  return specsInText((node as ShapeRenderNode).text)
}

/** Called before Konva stage.toDataURL so PNG/PDF never capture fallback math. */
export async function prepareEquationRasters(slide: RenderSlide): Promise<void> {
  await document.fonts?.ready
  const specs = slide.nodes.flatMap(specsInNode)
  for (let attempt = 0; attempt < 3; attempt++) {
    await Promise.all(specs.map(ensureEquationRaster))
    if (specs.every((spec) => cache.get(keyOf(spec))?.image)) return
  }
  throw new Error('Equation raster was invalidated before export capture')
}
