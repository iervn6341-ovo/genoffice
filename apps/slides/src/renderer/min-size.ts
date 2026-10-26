/**
 * Minimum on-slide size for a resize gesture. Without one a handle dragged past
 * the opposite edge collapses the element to nothing or (with Konva's default
 * flip) mirrors it — text turned upside down, a table flipped over. PowerPoint
 * stops the box at a small minimum instead; lines/connectors keep their legitimate
 * zero extent on one axis.
 */
import { tableMinScale, tableColFloors } from '@genoffice/pptx-render'

export const MIN_ELEMENT_PX = 24
/** Text boxes keep room for a few characters per line instead of a one-glyph-per-line column */
export const MIN_TEXT_W_PX = 64
/** A table cannot go below a readable cell grid */
export const MIN_TABLE_ROW_PX = 24
/** Column floor of a table whose content is unknown (no measured column needs) */
export const MIN_TABLE_COL_PX = 64

export interface MinSize {
  w: number
  h: number
}

/** The element geometry minSizeFor needs (a render node reduced to what decides its floor) */
export interface MinNode {
  type: string
  rows?: number
  cols?: number
  /** Current column widths / row heights of a table (px): a column already narrower than the floor keeps its own width as floor */
  colWidths?: number[]
  rowHeights?: number[]
  /** Height each row's wrapped text needs: the table cannot be dragged shorter than the sum */
  rowNeeds?: number[]
  /** Width each column needs so no cell cuts a word apart (px, 0 = none): the table cannot be narrowed
      until a column drops under it */
  colNeeds?: number[]
  /** Narrowest width keeping the table's bottom inside the slide (see TableRenderNode.fitMinWPx) */
  fitMinW?: number
  isConnector?: boolean
  /** Current size and children of a group: no child may be squeezed below its own floor */
  w?: number
  h?: number
  children?: MinNode[]
}

/** PowerPoint's element size limit (56 in at 96 dpi) */
export const MAX_ELEMENT_PX = 5376

/** Floor of a child that is already smaller than its floor: it may shrink at most to half */
function childFloor(min: number, size: number): number {
  return size <= min ? size * 0.5 : min
}

/**
 * Smallest box that still keeps every member at or above its own floor when the whole set is
 * scaled uniformly per axis (group children, multi-selection): the tightest member decides.
 */
export function minSizeForMembers(members: MinNode[], w: number, h: number): MinSize {
  let rw = 0
  let rh = 0
  for (const m of members) {
    if (m.isConnector) continue
    const mw = m.w ?? 0
    const mh = m.h ?? 0
    const mm = minSizeFor(m)
    if (mw > 0.5) rw = Math.max(rw, childFloor(mm.w, mw) / mw)
    if (mh > 0.5) rh = Math.max(rh, childFloor(mm.h, mh) / mh)
  }
  return { w: Math.max(MIN_ELEMENT_PX, rw * w), h: Math.max(MIN_ELEMENT_PX, rh * h) }
}

/**
 * Narrowest a table may be made. The engine narrows a table by scaling all its columns by one
 * factor, so the table stops where the first column reaches its content floor (the widest
 * unbreakable word of its cells + insets, never under TABLE_COL_FLOOR_PX, never above the width
 * the column already has): shrinking further would cut a word into pieces. Without measured
 * column needs (hand-built nodes) the old fixed per-column floor applies.
 */
function tableMinWidth(node: MinNode): number {
  const { colWidths: ws, cols = 0, colNeeds } = node
  if (!ws || ws.length !== cols) return cols * MIN_TABLE_COL_PX
  if (!colNeeds || colNeeds.length !== cols)
    return ws.reduce((a, v) => a + Math.min(Math.max(v, 1), MIN_TABLE_COL_PX), 0)
  const total = ws.reduce((a, v) => a + v, 0)
  return total * tableMinScale(ws, colNeeds)
}

export function minSizeFor(node: MinNode): MinSize {
  if (node.isConnector) return { w: 0, h: 0 }
  if (node.type === 'table' && node.rows && node.cols) {
    // a row that is already thinner than the floor may not be forced taller than it is
    const rowFloor = (): number => {
      const { rowHeights: hs, rows = 0, rowNeeds: needs } = node
      if (!hs || hs.length !== rows) return rows * MIN_TABLE_ROW_PX
      return hs.reduce(
        (a, v, i) =>
          a + Math.max(Math.min(Math.max(v, 1), MIN_TABLE_ROW_PX), Math.min(needs?.[i] ?? 0, v)),
        0,
      )
    }
    return {
      w: Math.max(MIN_ELEMENT_PX, tableMinWidth(node), node.fitMinW ?? 0),
      h: Math.max(MIN_ELEMENT_PX, rowFloor()),
    }
  }
  if (node.type === 'group' && node.children?.length && node.w && node.h) {
    return minSizeForMembers(node.children, node.w, node.h)
  }
  if (node.type === 'text') return { w: MIN_TEXT_W_PX, h: MIN_ELEMENT_PX }
  return { w: MIN_ELEMENT_PX, h: MIN_ELEMENT_PX }
}

/**
 * A typed/dragged numeric size → a legal one: non-finite values are rejected (null), values under the
 * floor or over PowerPoint's limit are clamped. `min` is a per-axis floor from minSizeFor.
 */
export function sanitizeSizePx(v: number, min: number, max = MAX_ELEMENT_PX): number | null {
  if (!Number.isFinite(v)) return null
  return Math.min(Math.max(v, min), max)
}

/** A typed position (px): finite and within ±MAX_ELEMENT_PX, else rejected */
export function sanitizePosPx(v: number, max = MAX_ELEMENT_PX): number | null {
  if (!Number.isFinite(v)) return null
  return Math.min(Math.max(v, -max), max)
}

export interface TBox {
  x: number
  y: number
  width: number
  height: number
  rotation: number
}

/**
 * Clamp the proposed transformer box to `min`. For an unrotated box the edge being
 * dragged moves and the opposite edge stays put; a rotated box falls back to the
 * previous box when it would go under the minimum (its corner math is not axis-aligned).
 */
export function clampTransformBox(old: TBox, next: TBox, min: MinSize): TBox {
  const w = Math.abs(next.width)
  const h = Math.abs(next.height)
  const flipped = next.width < 0 || next.height < 0
  if (!flipped && w >= min.w - 1e-6 && h >= min.h - 1e-6) return next
  if (next.rotation !== 0 || old.rotation !== 0) return old
  const out = { ...next }
  if (flipped || w < min.w) {
    out.width = Math.max(min.w, old.width < min.w ? old.width : min.w)
    // left-side handle: keep the right edge where it was
    if (Math.abs(next.x - old.x) > 1e-6) out.x = old.x + old.width - out.width
    else out.x = old.x
  }
  if (flipped || h < min.h) {
    out.height = Math.max(min.h, old.height < min.h ? old.height : min.h)
    if (Math.abs(next.y - old.y) > 1e-6) out.y = old.y + old.height - out.height
    else out.y = old.y
  }
  return out
}

export interface AnchorRect {
  x: number
  y: number
  width: number
  height: number
  rotation: number
}

/**
 * Constrain the dragged transformer anchor so the box can never collapse below `min` or
 * be dragged through the opposite edge (Konva would otherwise swap handles and re-form
 * the box on the far side, which is what made elements "break"). `name` is the Konva
 * anchor ('top-left', 'middle-right', …); `pos` and `rect` are absolute stage coordinates,
 * `rect.rotation` in radians. The opposite edge stays fixed.
 */
export function clampAnchorPos(
  name: string,
  pos: { x: number; y: number },
  rect: AnchorRect,
  min: MinSize,
): { x: number; y: number } {
  const left = name.includes('left')
  const right = name.includes('right')
  const top = name.includes('top')
  const bottom = name.includes('bottom')
  if (!left && !right && !top && !bottom) return pos
  const cos = Math.cos(rect.rotation)
  const sin = Math.sin(rect.rotation)
  // into the box's own frame (origin = its top-left corner)
  const dx = pos.x - rect.x
  const dy = pos.y - rect.y
  let lx = dx * cos + dy * sin
  let ly = -dx * sin + dy * cos
  if (left) lx = Math.min(lx, rect.width - min.w)
  if (right) lx = Math.max(lx, min.w)
  if (top) ly = Math.min(ly, rect.height - min.h)
  if (bottom) ly = Math.max(ly, min.h)
  return {
    x: rect.x + lx * cos - ly * sin,
    y: rect.y + lx * sin + ly * cos,
  }
}

export const ALL_ANCHORS = [
  'top-left',
  'top-center',
  'top-right',
  'middle-left',
  'middle-right',
  'bottom-left',
  'bottom-center',
  'bottom-right',
] as const

/**
 * Transformer handles that fit a box of the given on-screen size. When an edge is shorter than three
 * handle widths its middle handle would sit on top of the two corner handles (and win the hit test
 * from them), so — like PowerPoint on a tiny object — only the corners stay. `anchorPx` is the
 * on-screen handle size.
 */
export function anchorsForScreenSize(screenW: number, screenH: number, anchorPx: number): string[] {
  const needed = 3 * anchorPx
  const midX = screenW >= needed
  const midY = screenH >= needed
  return ALL_ANCHORS.filter((a) => {
    if (a === 'top-center' || a === 'bottom-center') return midX
    if (a === 'middle-left' || a === 'middle-right') return midY
    return true
  })
}

/**
 * Legal size for a typed/stepped value pair: floors and PowerPoint's limit applied, and with the
 * aspect ratio locked a clamp on one side carries over to the other so the proportions survive.
 * Null when either value is not a finite number.
 */
export function fitSizeToLimits(
  w: number,
  h: number,
  min: MinSize,
  opts: { lockRatio?: boolean; max?: number } = {},
): MinSize | null {
  const max = opts.max ?? MAX_ELEMENT_PX
  if (!Number.isFinite(w) || !Number.isFinite(h)) return null
  const minW = Math.max(min.w, 1)
  const minH = Math.max(min.h, 1)
  if (opts.lockRatio && w > 0 && h > 0) {
    // one uniform factor so neither side leaves its range
    let f = 1
    if (w < minW || h < minH) f = Math.max(minW / w, minH / h)
    else if (w > max || h > max) f = Math.min(max / w, max / h)
    return { w: Math.min(Math.max(w * f, minW), max), h: Math.min(Math.max(h * f, minH), max) }
  }
  return { w: Math.min(Math.max(w, minW), max), h: Math.min(Math.max(h, minH), max) }
}

/** Narrowest column / row a border drag or a typed cell size may produce */
export const MIN_CELL_COL_PX = 32
export const MIN_CELL_ROW_PX = 12

/**
 * Where a dragged inner column border ends up. Like PowerPoint the border moves between the two
 * columns it separates — the table keeps its width — and stops when either would drop under its
 * floor: MIN_CELL_COL_PX, or with measured `needs` (width each column needs to keep its words whole)
 * the content floor, never forcing a column wider than it already is. The last column's outer edge
 * instead changes the table width (no neighbour). `widths` are the current column widths (px),
 * `col` the column whose right/outer edge was dragged, `want` its requested width. Non-finite input
 * leaves the column as it is.
 */
export function resolveColDrag(
  widths: number[],
  col: number,
  want: number,
  needs?: number[],
): { wPx: number; neighbor?: { col: number; wPx: number } } {
  const cur = widths[col] ?? MIN_CELL_COL_PX
  if (!Number.isFinite(want)) return { wPx: cur }
  const { lo, hi } = colDragRange(widths, col, needs)
  const w = Math.min(Math.max(want, lo), hi)
  const next = widths[col + 1]
  return next == null ? { wPx: w } : { wPx: w, neighbor: { col: col + 1, wPx: cur + next - w } }
}

/** Legal width range of the column whose right/outer border is dragged (see resolveColDrag) */
export function colDragRange(
  widths: number[],
  col: number,
  needs?: number[],
): { lo: number; hi: number } {
  const cur = widths[col] ?? MIN_CELL_COL_PX
  const floors = needs ? tableColFloors(widths, needs, MIN_CELL_COL_PX) : undefined
  const next = widths[col + 1]
  if (next == null) return { lo: floors ? floors[col]! : MIN_CELL_COL_PX, hi: MAX_ELEMENT_PX }
  const pair = cur + next
  const loA = floors ? floors[col]! : Math.min(MIN_CELL_COL_PX, pair / 2)
  const loB = floors ? floors[col + 1]! : Math.min(MIN_CELL_COL_PX, pair / 2)
  return { lo: loA, hi: pair - loB }
}

/** Narrowest width a typed value may give the listed columns (all get the same width): the neediest decides */
export function colWidthFloor(widths: number[], cols: number[], needs?: number[]): number {
  const floors = tableColFloors(widths, needs, MIN_CELL_COL_PX)
  return Math.max(MIN_CELL_COL_PX / 2, ...cols.map((c) => floors[c] ?? MIN_CELL_COL_PX))
}

/**
 * A dragged/typed row height: finite, at least MIN_CELL_ROW_PX, at most PowerPoint's limit. With the
 * height the row's wrapped text `need`s the row stops there (the engine would grow it back at once,
 * which shows as a snap when the drag ends) — never forcing a row taller than it already is.
 */
export function resolveRowDrag(want: number, cur: number, need = 0): number {
  if (!Number.isFinite(want)) return cur
  return Math.min(Math.max(want, rowDragFloor(cur, need)), MAX_ELEMENT_PX)
}

/** Shortest a row may be dragged: its text need (never more than its own height), never under MIN_CELL_ROW_PX */
export function rowDragFloor(cur: number, need = 0): number {
  return Math.max(MIN_CELL_ROW_PX, Math.min(need, Math.max(cur, 1)))
}
