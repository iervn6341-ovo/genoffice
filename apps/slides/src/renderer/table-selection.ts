/**
 * Table cell-range selection model (pure — no React, no IPC).
 *
 * A selection is an anchor cell and a focus cell; the highlighted block is the rectangle they span,
 * widened until no merged cell is cut in two (PowerPoint does the same). Coordinates are the render
 * tree's model coordinates: `row` is the a:tr index, `col` the a:tc index — which equals the grid
 * column, because every grid column has a tc in each row (covered cells of a merge included).
 */
import type { TableCellRender, TableRenderNode } from '@genoffice/pptx-render'

export interface CellAddr {
  row: number
  col: number
}

export interface TableCellSel {
  sourceId: string
  anchor: CellAddr
  focus: CellAddr
}

/** Inclusive grid rectangle */
export interface CellRect {
  r0: number
  c0: number
  r1: number
  c1: number
}

type CellsOf = Pick<TableRenderNode, 'cells' | 'gridX' | 'gridY'>

export const gridRows = (t: CellsOf): number => Math.max(0, t.gridY.length - 1)
export const gridCols = (t: CellsOf): number => Math.max(0, t.gridX.length - 1)

/** Grid rectangle a (possibly merged) cell covers */
export function cellRect(
  cell: Pick<TableCellRender, 'row' | 'col' | 'rowSpan' | 'gridSpan'>,
): CellRect {
  return {
    r0: cell.row,
    c0: cell.col,
    r1: cell.row + (cell.rowSpan ?? 1) - 1,
    c1: cell.col + (cell.gridSpan ?? 1) - 1,
  }
}

/** The rendered (anchor) cell covering grid position (r, c) */
export function cellCovering(t: CellsOf, r: number, c: number): TableCellRender | undefined {
  return t.cells.find((cell) => {
    const q = cellRect(cell)
    return r >= q.r0 && r <= q.r1 && c >= q.c0 && c <= q.c1
  })
}

const union = (a: CellRect, b: CellRect): CellRect => ({
  r0: Math.min(a.r0, b.r0),
  c0: Math.min(a.c0, b.c0),
  r1: Math.max(a.r1, b.r1),
  c1: Math.max(a.c1, b.c1),
})

const intersects = (a: CellRect, b: CellRect): boolean =>
  a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0

/** Widen a rectangle until every merged cell it touches lies completely inside */
export function expandToMerges(t: CellsOf, rect: CellRect): CellRect {
  let cur = rect
  for (let guard = 0; guard < 64; guard++) {
    let next = cur
    for (const cell of t.cells) {
      const q = cellRect(cell)
      if (intersects(cur, q)) next = union(next, q)
    }
    if (next.r0 === cur.r0 && next.c0 === cur.c0 && next.r1 === cur.r1 && next.c1 === cur.c1) {
      return cur
    }
    cur = next
  }
  return cur
}

/** The block a selection highlights (anchor/focus merged cells included, merges never cut) */
export function selectionRect(t: CellsOf, sel: Pick<TableCellSel, 'anchor' | 'focus'>): CellRect {
  const a = cellCovering(t, sel.anchor.row, sel.anchor.col)
  const f = cellCovering(t, sel.focus.row, sel.focus.col)
  const ra = a
    ? cellRect(a)
    : { r0: sel.anchor.row, c0: sel.anchor.col, r1: sel.anchor.row, c1: sel.anchor.col }
  const rf = f
    ? cellRect(f)
    : { r0: sel.focus.row, c0: sel.focus.col, r1: sel.focus.row, c1: sel.focus.col }
  return expandToMerges(t, union(ra, rf))
}

/** Rendered cells inside a rectangle (after expandToMerges every one lies fully inside) */
export function cellsInRect(t: CellsOf, rect: CellRect): TableCellRender[] {
  return t.cells.filter((cell) => intersects(rect, cellRect(cell)))
}

/** Addresses (row, tc index) of the rendered cells in a rectangle, row-major */
export function cellAddrsInRect(t: CellsOf, rect: CellRect): CellAddr[] {
  return cellsInRect(t, rect)
    .map((c) => ({ row: c.row, col: c.col }))
    .sort((a, b) => a.row - b.row || a.col - b.col)
}

/** Table-local pixel bounds of a rectangle (union of the rendered cells in it) */
export function rectBounds(
  t: CellsOf,
  rect: CellRect,
): { x: number; y: number; w: number; h: number } | null {
  const cells = cellsInRect(t, rect)
  if (!cells.length) return null
  const x0 = Math.min(...cells.map((c) => c.x))
  const y0 = Math.min(...cells.map((c) => c.y))
  const x1 = Math.max(...cells.map((c) => c.x + c.w))
  const y1 = Math.max(...cells.map((c) => c.y + c.h))
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 }
}

export const singleCellSel = (sourceId: string, at: CellAddr): TableCellSel => ({
  sourceId,
  anchor: at,
  focus: at,
})

/** Whole row(s) between two row indices */
export function rowSel(t: CellsOf, sourceId: string, from: number, to = from): TableCellSel {
  const last = gridCols(t) - 1
  return {
    sourceId,
    anchor: { row: from, col: 0 },
    focus: { row: to, col: Math.max(0, last) },
  }
}

/** Whole column(s) between two column indices */
export function colSel(t: CellsOf, sourceId: string, from: number, to = from): TableCellSel {
  const last = gridRows(t) - 1
  return {
    sourceId,
    anchor: { row: 0, col: from },
    focus: { row: Math.max(0, last), col: to },
  }
}

export function allSel(t: CellsOf, sourceId: string): TableCellSel {
  return {
    sourceId,
    anchor: { row: 0, col: 0 },
    focus: { row: Math.max(0, gridRows(t) - 1), col: Math.max(0, gridCols(t) - 1) },
  }
}

export type Dir = 'up' | 'down' | 'left' | 'right'

/**
 * One step from a cell in a direction: the rendered cell beyond its edge (a merged cell is crossed
 * in one step); stays put at the table edge. `via` picks the row/column the step travels along.
 */
export function stepFrom(t: CellsOf, from: CellAddr, dir: Dir): CellAddr {
  const here = cellCovering(t, from.row, from.col)
  const q = here ? cellRect(here) : { r0: from.row, c0: from.col, r1: from.row, c1: from.col }
  let r = from.row
  let c = from.col
  if (dir === 'up') r = q.r0 - 1
  else if (dir === 'down') r = q.r1 + 1
  else if (dir === 'left') c = q.c0 - 1
  else c = q.c1 + 1
  if (r < 0 || c < 0 || r >= gridRows(t) || c >= gridCols(t)) return from
  const next = cellCovering(t, r, c)
  return next ? { row: next.row, col: next.col } : from
}

/** Arrow key without Shift: collapse to one cell next to the focus */
export function moveSel(t: CellsOf, sel: TableCellSel, dir: Dir): TableCellSel {
  const to = stepFrom(t, sel.focus, dir)
  return singleCellSel(sel.sourceId, to)
}

/** Shift+arrow: move only the focus corner, the anchor stays */
export function extendSel(t: CellsOf, sel: TableCellSel, dir: Dir): TableCellSel {
  return { ...sel, focus: stepFrom(t, sel.focus, dir) }
}

/** Tab / Shift+Tab: next / previous rendered cell in reading order, null at the ends */
export function tabTarget(t: CellsOf, from: CellAddr, dir: 1 | -1): CellAddr | null {
  const sorted = [...t.cells].sort((a, b) => a.row - b.row || a.col - b.col)
  const i = sorted.findIndex((c) => c.row === from.row && c.col === from.col)
  const j = i < 0 ? 0 : i + dir
  const next = sorted[j]
  return next ? { row: next.row, col: next.col } : null
}

export const sameSel = (a: TableCellSel | null, b: TableCellSel | null): boolean =>
  a === b ||
  (!!a &&
    !!b &&
    a.sourceId === b.sourceId &&
    a.anchor.row === b.anchor.row &&
    a.anchor.col === b.anchor.col &&
    a.focus.row === b.focus.row &&
    a.focus.col === b.focus.col)

/** Keep a stored selection valid after the grid changed (rows/columns inserted or deleted) */
export function clampSel(t: CellsOf, sel: TableCellSel): TableCellSel | null {
  const rows = gridRows(t)
  const cols = gridCols(t)
  if (!rows || !cols) return null
  const cl = (a: CellAddr): CellAddr => ({
    row: Math.min(Math.max(a.row, 0), rows - 1),
    col: Math.min(Math.max(a.col, 0), cols - 1),
  })
  return { ...sel, anchor: cl(sel.anchor), focus: cl(sel.focus) }
}

/** Plain text of one cell (paragraphs joined by \n, soft breaks kept as \n, bullets dropped) */
export function cellPlainText(cell: Pick<TableCellRender, 'text'>): string {
  const lines = cell.text?.lines ?? []
  const paras: string[] = []
  for (const line of lines) {
    let s = ''
    for (const r of line.runs) if (!r.isBullet) s += r.text
    if (line.trailingSpace) s += line.trailingText ?? ' '
    if (line.paraStart === false && paras.length) paras[paras.length - 1] += s
    else paras.push(s)
    if (line.softBreakAfter != null) paras[paras.length - 1] += '\n'
  }
  return paras.join('\n')
}

/** Selected block as tab/newline separated text (what Excel and Word accept as a range) */
export function rangeToTsv(t: CellsOf, rect: CellRect): string {
  const rows: string[] = []
  for (let r = rect.r0; r <= rect.r1; r++) {
    const parts: string[] = []
    for (let c = rect.c0; c <= rect.c1; c++) {
      const cell = t.cells.find((x) => x.row === r && x.col === c)
      if (!cell) {
        // covered by a merge anchored elsewhere: an empty field keeps the columns aligned
        parts.push('')
        continue
      }
      const text = cellPlainText(cell)
      parts.push(/[\t\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text)
    }
    rows.push(parts.join('\t'))
  }
  return rows.join('\n')
}

const escHtml = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/** Selected block as an HTML table (what Excel / Numbers / Pages / Word paste as a grid): text, fills and merges */
export function rangeToHtml(t: CellsOf, rect: CellRect): string {
  const rows: string[] = []
  for (let r = rect.r0; r <= rect.r1; r++) {
    const tds: string[] = []
    for (let c = rect.c0; c <= rect.c1; c++) {
      const cell = t.cells.find((x) => x.row === r && x.col === c)
      if (!cell) continue
      const span =
        (cell.gridSpan && cell.gridSpan > 1 ? ` colspan="${Math.min(cell.gridSpan, rect.c1 - c + 1)}"` : '') +
        (cell.rowSpan && cell.rowSpan > 1 ? ` rowspan="${Math.min(cell.rowSpan, rect.r1 - r + 1)}"` : '')
      const fill =
        cell.fill.kind === 'solid' && /^#[0-9a-f]{6}/i.test(cell.fill.color)
          ? ` style="background:${cell.fill.color.slice(0, 7)}"`
          : ''
      const body = cellPlainText(cell).split('\n').map(escHtml).join('<br>')
      tds.push(`<td${span}${fill}>${body}</td>`)
    }
    rows.push(`<tr>${tds.join('')}</tr>`)
  }
  return `<table border="0" cellspacing="0" cellpadding="3">${rows.join('')}</table>`
}

/** Parse tab/newline separated clipboard text (RFC 4180 style quoting, like Excel) into rows of fields */
export function parseTsv(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  let i = 0
  const src = text.replace(/\r\n?/g, '\n')
  while (i < src.length) {
    const ch = src[i]!
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') {
          field += '"'
          i += 2
          continue
        }
        quoted = false
        i++
        continue
      }
      field += ch
      i++
      continue
    }
    if (ch === '"' && field === '') {
      quoted = true
      i++
    } else if (ch === '\t') {
      row.push(field)
      field = ''
      i++
    } else if (ch === '\n') {
      row.push(field)
      rows.push(row)
      row = []
      field = ''
      i++
    } else {
      field += ch
      i++
    }
  }
  if (field !== '' || row.length) {
    row.push(field)
    rows.push(row)
  }
  return rows
}
