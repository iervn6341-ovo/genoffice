import type { TableRenderNode } from '@genoffice/pptx-render'
import { describe, expect, it } from 'vitest'
import {
  allSel,
  cellAddrsInRect,
  cellPlainText,
  clampSel,
  colSel,
  extendSel,
  moveSel,
  parseTsv,
  rangeToHtml,
  rangeToTsv,
  rectBounds,
  rowSel,
  sameSel,
  selectionRect,
  singleCellSel,
  stepFrom,
  tabTarget,
} from '../src/renderer/table-selection'

/** 4 x 4 grid of 50 x 20 px cells; `merges` = [row, col, rowSpan, gridSpan] anchors (covered cells dropped) */
function makeTable(merges: Array<[number, number, number, number]> = []): TableRenderNode {
  const covered = new Set<string>()
  for (const [r, c, rs, gs] of merges)
    for (let i = 0; i < rs; i++)
      for (let j = 0; j < gs; j++) if (i || j) covered.add(`${r + i},${c + j}`)
  const cells = []
  for (let r = 0; r < 4; r++)
    for (let c = 0; c < 4; c++) {
      if (covered.has(`${r},${c}`)) continue
      const m = merges.find(([mr, mc]) => mr === r && mc === c)
      cells.push({
        x: c * 50,
        y: r * 20,
        w: 50 * (m?.[3] ?? 1),
        h: 20 * (m?.[2] ?? 1),
        row: r,
        col: c,
        ...(m && m[3] > 1 ? { gridSpan: m[3] } : {}),
        ...(m && m[2] > 1 ? { rowSpan: m[2] } : {}),
        fill: { kind: 'none' },
      })
    }
  return {
    type: 'table',
    sourceId: 't1',
    id: 'r_t1',
    box: { x: 0, y: 0, w: 200, h: 80 },
    cells,
    gridX: [0, 50, 100, 150, 200],
    gridY: [0, 20, 40, 60, 80],
  } as unknown as TableRenderNode
}

const sel = (ar: number, ac: number, fr: number, fc: number) => ({
  sourceId: 't1',
  anchor: { row: ar, col: ac },
  focus: { row: fr, col: fc },
})

describe('selectionRect', () => {
  it('spans anchor and focus whichever way they were dragged', () => {
    const t = makeTable()
    expect(selectionRect(t, sel(1, 1, 2, 3))).toEqual({ r0: 1, c0: 1, r1: 2, c1: 3 })
    expect(selectionRect(t, sel(2, 3, 1, 1))).toEqual({ r0: 1, c0: 1, r1: 2, c1: 3 })
    expect(selectionRect(t, sel(0, 0, 0, 0))).toEqual({ r0: 0, c0: 0, r1: 0, c1: 0 })
  })

  it('widens until no merged cell is cut', () => {
    // merge covers rows 1-2, cols 1-2; dragging (0,0)..(1,1) must swallow it whole
    const t = makeTable([[1, 1, 2, 2]])
    expect(selectionRect(t, sel(0, 0, 1, 1))).toEqual({ r0: 0, c0: 0, r1: 2, c1: 2 })
    // a focus inside the merge's covered area resolves to the whole merge
    expect(selectionRect(t, sel(3, 3, 2, 2))).toEqual({ r0: 1, c0: 1, r1: 3, c1: 3 })
  })

  it('chained merges widen repeatedly', () => {
    const t = makeTable([
      [0, 0, 1, 2],
      [0, 1, 1, 1],
    ])
    expect(selectionRect(t, sel(0, 0, 0, 0)).c1).toBeGreaterThanOrEqual(1)
  })
})

describe('cells and bounds in a rectangle', () => {
  it('lists rendered cells row-major and measures the union', () => {
    const t = makeTable([[1, 1, 2, 2]])
    const rect = selectionRect(t, sel(0, 0, 1, 1))
    const addrs = cellAddrsInRect(t, rect)
    // (0,0)(0,1)(0,2)(1,0)(1,1 merged anchor)(2,0)
    expect(addrs).toEqual([
      { row: 0, col: 0 },
      { row: 0, col: 1 },
      { row: 0, col: 2 },
      { row: 1, col: 0 },
      { row: 1, col: 1 },
      { row: 2, col: 0 },
    ])
    expect(rectBounds(t, rect)).toEqual({ x: 0, y: 0, w: 150, h: 60 })
  })
})

describe('whole row / column / table selection', () => {
  it('row, column and all', () => {
    const t = makeTable()
    expect(selectionRect(t, rowSel(t, 't1', 2))).toEqual({ r0: 2, c0: 0, r1: 2, c1: 3 })
    expect(selectionRect(t, rowSel(t, 't1', 1, 3))).toEqual({ r0: 1, c0: 0, r1: 3, c1: 3 })
    expect(selectionRect(t, colSel(t, 't1', 1))).toEqual({ r0: 0, c0: 1, r1: 3, c1: 1 })
    expect(selectionRect(t, allSel(t, 't1'))).toEqual({ r0: 0, c0: 0, r1: 3, c1: 3 })
  })
})

describe('keyboard movement', () => {
  it('arrows collapse to the neighbouring cell and stop at the edge', () => {
    const t = makeTable()
    const s = singleCellSel('t1', { row: 0, col: 0 })
    expect(moveSel(t, s, 'right').focus).toEqual({ row: 0, col: 1 })
    expect(moveSel(t, s, 'up').focus).toEqual({ row: 0, col: 0 })
    expect(moveSel(t, s, 'left').focus).toEqual({ row: 0, col: 0 })
    expect(moveSel(t, singleCellSel('t1', { row: 3, col: 3 }), 'down').focus).toEqual({
      row: 3,
      col: 3,
    })
  })

  it('a step crosses a merged cell in one move', () => {
    const t = makeTable([[1, 1, 1, 2]])
    expect(stepFrom(t, { row: 1, col: 0 }, 'right')).toEqual({ row: 1, col: 1 })
    expect(stepFrom(t, { row: 1, col: 1 }, 'right')).toEqual({ row: 1, col: 3 })
    expect(stepFrom(t, { row: 1, col: 3 }, 'left')).toEqual({ row: 1, col: 1 })
  })

  it('shift+arrow moves the focus corner only', () => {
    const t = makeTable()
    let s = singleCellSel('t1', { row: 1, col: 1 })
    s = extendSel(t, s, 'right')
    s = extendSel(t, s, 'down')
    expect(s.anchor).toEqual({ row: 1, col: 1 })
    expect(s.focus).toEqual({ row: 2, col: 2 })
    expect(selectionRect(t, s)).toEqual({ r0: 1, c0: 1, r1: 2, c1: 2 })
    s = extendSel(t, s, 'left')
    s = extendSel(t, s, 'up')
    expect(selectionRect(t, s)).toEqual({ r0: 1, c0: 1, r1: 1, c1: 1 })
  })

  it('tab walks reading order and returns null past either end', () => {
    const t = makeTable([[0, 0, 1, 2]])
    expect(tabTarget(t, { row: 0, col: 0 }, 1)).toEqual({ row: 0, col: 2 })
    expect(tabTarget(t, { row: 0, col: 3 }, 1)).toEqual({ row: 1, col: 0 })
    expect(tabTarget(t, { row: 1, col: 0 }, -1)).toEqual({ row: 0, col: 3 })
    expect(tabTarget(t, { row: 0, col: 0 }, -1)).toBeNull()
    expect(tabTarget(t, { row: 3, col: 3 }, 1)).toBeNull()
  })
})

describe('selection bookkeeping', () => {
  it('sameSel compares sourceId and both corners', () => {
    expect(sameSel(sel(0, 0, 1, 1), sel(0, 0, 1, 1))).toBe(true)
    expect(sameSel(sel(0, 0, 1, 1), sel(0, 0, 1, 2))).toBe(false)
    expect(sameSel(null, null)).toBe(true)
    expect(sameSel(sel(0, 0, 0, 0), null)).toBe(false)
  })

  it('clampSel pulls a stale selection back inside a smaller grid', () => {
    const t = makeTable()
    expect(clampSel(t, sel(0, 0, 9, 9))!.focus).toEqual({ row: 3, col: 3 })
  })
})

describe('text interchange', () => {
  const textCell = (paras: string[][]) => ({
    text: {
      lines: paras.flatMap((lines, pi) =>
        lines.map((s, li) => ({
          runs: [{ text: s }],
          top: 0,
          height: 10,
          ...(li > 0 ? { paraStart: false } : pi >= 0 ? { paraStart: true } : {}),
        })),
      ),
    },
  })

  it('cellPlainText joins wrapped lines and separates paragraphs', () => {
    const c = textCell([['Hello ', 'world'], ['Second']]) as never
    expect(cellPlainText(c)).toBe('Hello world\nSecond')
  })

  it('rangeToTsv quotes fields that hold tabs, newlines or quotes', () => {
    const t = makeTable()
    t.cells[0]!.text = textCell([['a']]).text as never
    t.cells[1]!.text = textCell([['b"c']]).text as never
    t.cells[4]!.text = textCell([['x'], ['y']]).text as never
    expect(rangeToTsv(t, { r0: 0, c0: 0, r1: 1, c1: 1 })).toBe('a\t"b""c"\n"x\ny"\t')
  })

  it('rangeToHtml writes an HTML table with text, merges and fills for spreadsheets', () => {
    const t = makeTable([[0, 0, 1, 2]])
    t.cells[0]!.text = textCell([['a<b']]).text as never
    t.cells[0]!.fill = { kind: 'solid', color: '#FFC000' } as never
    const html = rangeToHtml(t, { r0: 0, c0: 0, r1: 1, c1: 1 })
    expect(html).toContain('<table')
    expect(html).toContain('colspan="2"')
    expect(html).toContain('background:#FFC000')
    expect(html).toContain('a&lt;b')
    expect(html.match(/<tr>/g)).toHaveLength(2)
  })

  it('parseTsv reads Excel style ranges back, including quotes and CRLF', () => {
    expect(parseTsv('a\tb\r\nc\td')).toEqual([
      ['a', 'b'],
      ['c', 'd'],
    ])
    expect(parseTsv('"x\ny"\t"q""r"\nz')).toEqual([['x\ny', 'q"r'], ['z']])
    expect(parseTsv('')).toEqual([])
    expect(parseTsv('only')).toEqual([['only']])
    expect(parseTsv('a\t\tb\n')).toEqual([['a', '', 'b']])
  })
})
