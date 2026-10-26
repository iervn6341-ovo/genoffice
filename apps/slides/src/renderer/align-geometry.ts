/**
 * Align / distribute geometry (pure). PowerPoint's Align menu has a "relative to" switch:
 *  - Align Selected Objects: against the selection's bounding box (distribute needs 3+)
 *  - Align to Slide: against the slide edges (distribute spreads first..last across the slide, 2+)
 * A single selected object always aligns to the slide.
 */
export type AlignOp =
  'left' | 'center-h' | 'right' | 'top' | 'center-v' | 'bottom' | 'distribute-h' | 'distribute-v'
export type AlignTo = 'selection' | 'slide'

export interface Rect {
  x: number
  y: number
  w: number
  h: number
}

export function boundingBox(rs: Rect[]): Rect {
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const r of rs) {
    minX = Math.min(minX, r.x)
    minY = Math.min(minY, r.y)
    maxX = Math.max(maxX, r.x + r.w)
    maxY = Math.max(maxY, r.y + r.h)
  }
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY }
}

/** Fewest elements a distribute needs in the given mode */
export function minForDistribute(to: AlignTo): number {
  return to === 'slide' ? 2 : 3
}

/** New top-left positions (same order as `rects`), or null when the operation does not apply. */
export function alignPositions(
  rects: Rect[],
  op: AlignOp,
  to: AlignTo,
  page: { w: number; h: number },
): Array<{ x: number; y: number }> | null {
  if (rects.length === 0) return null
  const pageRect: Rect = { x: 0, y: 0, w: page.w, h: page.h }
  const toSlide = to === 'slide' || rects.length === 1
  const container = toSlide ? pageRect : boundingBox(rects)

  if (op === 'distribute-h' || op === 'distribute-v') {
    if (rects.length < minForDistribute(to)) return null
    const horizontal = op === 'distribute-h'
    const pos = (r: Rect) => (horizontal ? r.x : r.y)
    const size = (r: Rect) => (horizontal ? r.w : r.h)
    const order = rects.map((r, i) => ({ r, i })).sort((a, b) => pos(a.r) - pos(b.r))
    const start = to === 'slide' ? 0 : pos(order[0]!.r)
    const span =
      to === 'slide'
        ? horizontal
          ? page.w
          : page.h
        : pos(order[order.length - 1]!.r) + size(order[order.length - 1]!.r) - start
    const total = order.reduce((s, { r }) => s + size(r), 0)
    const gap = (span - total) / (rects.length - 1)
    const out = rects.map((r) => ({ x: r.x, y: r.y }))
    let cursor = start
    for (const { r, i } of order) {
      out[i] = horizontal ? { x: cursor, y: r.y } : { x: r.x, y: cursor }
      cursor += size(r) + gap
    }
    return out
  }

  return rects.map((r) => {
    let x = r.x
    let y = r.y
    if (op === 'left') x = container.x
    else if (op === 'center-h') x = container.x + (container.w - r.w) / 2
    else if (op === 'right') x = container.x + container.w - r.w
    else if (op === 'top') y = container.y
    else if (op === 'center-v') y = container.y + (container.h - r.h) / 2
    else if (op === 'bottom') y = container.y + container.h - r.h
    return { x, y }
  })
}
