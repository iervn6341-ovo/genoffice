/**
 * Dark slide content preview — display-time color remap only.
 *
 * The slide model, the saved PPTX, exports (PNG/PDF/headless), printing and the slide show never
 * see this: callers hand `mapSlideColors(slide)` to the on-screen editing surfaces (canvas,
 * sidebar / sorter thumbnails) and keep the authored RenderSlide for everything else. The result
 * is a fresh tree (the input is never mutated) with identical ids, boxes and structure, so hit
 * testing and selection behave the same. Pictures, videos and other bitmap content keep their
 * pixels; only vector colors (fills, gradients, patterns, lines, text, tables, charts) are mapped.
 *
 * Mapping = Univer's hue-preserving luminance inversion, compressed onto the dark paper tone so
 * authored white lands on DARK_PAPER_HEX (same algorithm as Docs' editor/dark-page.ts).
 */
import type {
  ChartRenderNode,
  GlyphRun,
  GroupRenderNode,
  PictureRenderNode,
  RenderFill,
  RenderNode,
  RenderSlide,
  RenderStroke,
  RenderTextLayout,
  ShapeRenderNode,
  TableCellRender,
  TableRenderNode,
  TextLine,
} from '@genoffice/pptx-render'

/** dark paper tone: authored white maps exactly here */
export const DARK_PAPER_HEX = '#1e1e1e'
const PAPER_FLOOR = 0x1e / 255

const M = [
  [0.333, -0.667, -0.667],
  [-0.667, 0.333, -0.667],
  [-0.667, -0.667, 0.333],
] as const

const cache = new Map<string, string>()

function clamp01(v: number): number {
  return v > 1 ? 1 : v < 0 ? 0 : v
}

function hex2(v: number): string {
  return Math.round(clamp01(v) * 255)
    .toString(16)
    .padStart(2, '0')
}

const HEX_RE = /^#?([0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i
const RGB_RE =
  /^rgba?\(\s*([\d.]+)\s*[, ]\s*([\d.]+)\s*[, ]\s*([\d.]+)\s*(?:[,/]\s*([\d.]+%?)\s*)?\)$/i

/** Remap one authored CSS color for the dark preview; keywords/unparseable input pass through. */
export function darkContentColor(css: string): string {
  const hit = cache.get(css)
  if (hit !== undefined) return hit
  const s = css.trim()
  let out = css
  let r = NaN
  let g = 0
  let b = 0
  let alpha: string | undefined
  const hex = HEX_RE.exec(s)
  const rgb = hex ? null : RGB_RE.exec(s)
  if (hex) {
    let h = hex[1]!
    if (h.length <= 4) h = [...h].map((c) => c + c).join('')
    const n = parseInt(h.slice(0, 6), 16)
    r = ((n >> 16) & 255) / 255
    g = ((n >> 8) & 255) / 255
    b = (n & 255) / 255
    if (h.length === 8) alpha = h.slice(6, 8)
  } else if (rgb) {
    r = Number(rgb[1]) / 255
    g = Number(rgb[2]) / 255
    b = Number(rgb[3]) / 255
    if (rgb[4] !== undefined) alpha = rgb[4]
  }
  if (!Number.isNaN(r)) {
    const [mr, mg, mb] = M.map((row) => {
      const inv = clamp01(row[0] * r + row[1] * g + row[2] * b + 1)
      return PAPER_FLOOR + inv * (1 - PAPER_FLOOR)
    }) as [number, number, number]
    if (alpha === undefined) out = `#${hex2(mr)}${hex2(mg)}${hex2(mb)}`
    else if (hex) out = `#${hex2(mr)}${hex2(mg)}${hex2(mb)}${alpha}`
    else
      out = `rgb(${Math.round(mr * 255)} ${Math.round(mg * 255)} ${Math.round(mb * 255)} / ${alpha})`
  }
  cache.set(css, out)
  return out
}

const c = darkContentColor

function mapStops<T extends { color: string }>(stops: T[]): T[] {
  return stops.map((s) => ({ ...s, color: c(s.color) }))
}

export function mapFill(fill: RenderFill): RenderFill {
  switch (fill.kind) {
    case 'solid':
      return { ...fill, color: c(fill.color) }
    case 'gradient':
      return { ...fill, stops: mapStops(fill.stops) }
    case 'pattern':
      return { ...fill, fg: c(fill.fg), bg: c(fill.bg) }
    default:
      return fill // none / image: bitmap content keeps its pixels
  }
}

function mapStroke(s: RenderStroke): RenderStroke {
  return {
    ...s,
    color: c(s.color),
    ...(s.gradient ? { gradient: { ...s.gradient, stops: mapStops(s.gradient.stops) } } : {}),
  }
}

function mapRun(r: GlyphRun): GlyphRun {
  return {
    ...r,
    color: c(r.color),
    ...(r.highlight ? { highlight: c(r.highlight) } : {}),
    ...(r.outline ? { outline: { ...r.outline, color: c(r.outline.color) } } : {}),
    ...(r.gradient ? { gradient: { ...r.gradient, stops: mapStops(r.gradient.stops) } } : {}),
  }
}

function mapText(t: RenderTextLayout): RenderTextLayout {
  return {
    ...t,
    lines: t.lines.map((l): TextLine => ({ ...l, runs: l.runs.map(mapRun) })),
    ...(t.extrusion ? { extrusion: { ...t.extrusion, color: c(t.extrusion.color) } } : {}),
  }
}

function mapShape(n: ShapeRenderNode): ShapeRenderNode {
  return {
    ...n,
    fill: mapFill(n.fill),
    ...(n.fillOverlay ? { fillOverlay: mapFill(n.fillOverlay) } : {}),
    ...(n.stroke ? { stroke: mapStroke(n.stroke) } : {}),
    ...(n.extrusion
      ? {
          extrusion: {
            ...n.extrusion,
            faces: n.extrusion.faces.map((f) => ({
              ...f,
              color: c(f.color),
              ...(f.stroke ? { stroke: c(f.stroke) } : {}),
            })),
          },
        }
      : {}),
    ...(n.text ? { text: mapText(n.text) } : {}),
  }
}

function mapPicture(n: PictureRenderNode): PictureRenderNode {
  return {
    ...n,
    ...(n.fill ? { fill: mapFill(n.fill) } : {}),
    ...(n.stroke ? { stroke: mapStroke(n.stroke) } : {}),
  }
}

function mapCell(cell: TableCellRender): TableCellRender {
  const b = cell.borders
  return {
    ...cell,
    fill: mapFill(cell.fill),
    ...(b
      ? {
          borders: {
            ...(b.l ? { l: mapStroke(b.l) } : {}),
            ...(b.r ? { r: mapStroke(b.r) } : {}),
            ...(b.t ? { t: mapStroke(b.t) } : {}),
            ...(b.b ? { b: mapStroke(b.b) } : {}),
          },
        }
      : {}),
    ...(cell.text ? { text: mapText(cell.text) } : {}),
  }
}

function mapTable(n: TableRenderNode): TableRenderNode {
  return {
    ...n,
    cells: n.cells.map(mapCell),
    ...(n.bgFill ? { bgFill: mapFill(n.bgFill) } : {}),
  }
}

function mapChart(n: ChartRenderNode): ChartRenderNode {
  const pr = n.plotRect
  return {
    ...n,
    ...(n.bgFill ? { bgFill: mapFill(n.bgFill) } : {}),
    ...(n.border ? { border: { ...n.border, color: c(n.border.color) } } : {}),
    ...(pr
      ? {
          plotRect: {
            ...pr,
            ...(pr.fill ? { fill: mapFill(pr.fill) } : {}),
            ...(pr.borderColor ? { borderColor: c(pr.borderColor) } : {}),
          },
        }
      : {}),
    gridLines: n.gridLines.map((x) => ({ ...x, color: c(x.color) })),
    axisLines: n.axisLines.map((x) => ({ ...x, color: c(x.color) })),
    labels: n.labels.map((x) => ({ ...x, color: c(x.color) })),
    bars: n.bars.map((x) => ({ ...x, color: c(x.color) })),
    polylines: n.polylines.map((x) => ({
      ...x,
      color: c(x.color),
      ...(x.fill ? { fill: c(x.fill) } : {}),
    })),
    markers: n.markers.map((x) => ({ ...x, color: c(x.color) })),
    swatches: n.swatches.map((x) => ({ ...x, color: c(x.color) })),
    ...(n.paths
      ? {
          paths: n.paths.map((x) => ({
            ...x,
            fill: c(x.fill),
            ...(x.stroke ? { stroke: c(x.stroke) } : {}),
          })),
        }
      : {}),
    ...(n.wedges
      ? {
          wedges: n.wedges.map((x) => ({
            ...x,
            color: c(x.color),
            ...(x.stroke ? { stroke: c(x.stroke) } : {}),
          })),
        }
      : {}),
  }
}

export function mapNodeColors(n: RenderNode): RenderNode {
  switch (n.type) {
    case 'shape':
    case 'text':
      return mapShape(n)
    case 'picture':
      return mapPicture(n)
    case 'group':
      return {
        ...(n as GroupRenderNode),
        children: (n as GroupRenderNode).children.map(mapNodeColors),
      }
    case 'table':
      return mapTable(n)
    case 'chart':
      return mapChart(n)
    default:
      return n
  }
}

// Identity-stable per source slide: memoized thumbnails keep skipping re-render while the slide
// object is unchanged (the authored RenderSlide is replaced on every model edit).
const slideCache = new WeakMap<RenderSlide, RenderSlide>()

/** Authored RenderSlide → dark preview copy (never mutates the input; cached per input object). */
export function mapSlideColors(slide: RenderSlide): RenderSlide {
  const hit = slideCache.get(slide)
  if (hit) return hit
  const out: RenderSlide = {
    ...slide,
    // an unset background paints white (SlideCanvas/SlideThumb fallback) → the dark paper tone
    background:
      slide.background.kind === 'none'
        ? { kind: 'solid', color: DARK_PAPER_HEX }
        : mapFill(slide.background),
    nodes: slide.nodes.map(mapNodeColors),
  }
  slideCache.set(slide, out)
  return out
}

/** Display slide for a content-darkness flag: dark → remapped preview, light → authored as-is. */
export function displaySlide<T extends RenderSlide | null | undefined>(slide: T, dark: boolean): T {
  return (dark && slide ? mapSlideColors(slide) : slide) as T
}
