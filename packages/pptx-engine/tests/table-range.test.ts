/** Table cell-range edits: range merge/split, multi-row/column structure, scoped formatting, clearing. */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  addTable,
  clearTableCellText,
  editTableCellText,
  editTableStructure,
  mergeTableCells,
  openPptx,
  savePptx,
  setElementFont,
  setElementParagraphFormat,
} from '../src/index'
import type { TableElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

async function makeTable(rows = 4, cols = 4) {
  const opened = await openPptx(fx('01_standard_business.pptx'))
  const r = addTable(opened, 0, {
    rows,
    cols,
    offset: { x: 914400, y: 914400, cx: 4572000, cy: 1828800 },
  })!
  for (let i = 0; i < rows; i++)
    for (let j = 0; j < cols; j++)
      editTableCellText(r.slide, r.elementId, i, j, [{ runs: [{ text: `${i}${j}` }] }])
  return { opened, tblId: r.elementId }
}

const tableOf = (opened: Awaited<ReturnType<typeof makeTable>>['opened'], id: string) =>
  opened.deck.slides[0]!.elements.find((e) => e.id === id) as TableElement
const textOf = (t: TableElement, r: number, c: number) =>
  (t.rows[r]![c]!.text?.paragraphs ?? []).map((p) => p.runs.map((x) => x.text).join('')).join('|')

describe('mergeTableCells: merge-range / split-range', () => {
  it('merges a 2x2 block: spans on the anchor, covered cells flagged, text folded row-major', async () => {
    const { opened, tblId } = await makeTable()
    const r = mergeTableCells(opened, 0, tblId, {
      kind: 'merge-range',
      row: 1,
      col: 1,
      rows: 2,
      cols: 2,
    })!
    expect(r).not.toBeNull()
    const t = tableOf(opened, r.elementId)
    expect(t.rows[1]![1]!.gridSpan).toBe(2)
    expect(t.rows[1]![1]!.rowSpan).toBe(2)
    expect(t.rows[1]![2]!.merged).toBe(true)
    expect(t.rows[2]![1]!.merged).toBe(true)
    expect(t.rows[2]![2]!.merged).toBe(true)
    expect(textOf(t, 1, 1)).toBe('11|12|21|22')
    // outside the block nothing changed
    expect(t.rows[0]![0]!.merged ?? false).toBe(false)
    expect(textOf(t, 3, 3)).toBe('33')
  })

  it('survives save and reopen', async () => {
    const { opened, tblId } = await makeTable()
    mergeTableCells(opened, 0, tblId, { kind: 'merge-range', row: 0, col: 0, rows: 1, cols: 3 })
    mergeTableCells(
      opened,
      0,
      opened.deck.slides[0]!.elements.find((e) => e.type === 'table')!.id,
      {
        kind: 'merge-range',
        row: 2,
        col: 0,
        rows: 2,
        cols: 2,
      },
    )
    const reopened = await openPptx(await savePptx(opened))
    const t = reopened.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(t.rows[0]![0]!.gridSpan).toBe(3)
    expect(t.rows[2]![0]!.gridSpan).toBe(2)
    expect(t.rows[2]![0]!.rowSpan).toBe(2)
    expect(t.rows[3]![1]!.merged).toBe(true)
  })

  it('refuses a rectangle that cuts an existing merge, and a 1x1 rectangle', async () => {
    const { opened, tblId } = await makeTable()
    const r1 = mergeTableCells(opened, 0, tblId, {
      kind: 'merge-range',
      row: 0,
      col: 0,
      rows: 2,
      cols: 2,
    })!
    // cols 1..2 cut the merge that covers cols 0..1
    expect(
      mergeTableCells(opened, 0, r1.elementId, {
        kind: 'merge-range',
        row: 0,
        col: 1,
        rows: 2,
        cols: 2,
      }),
    ).toBeNull()
    expect(
      mergeTableCells(opened, 0, r1.elementId, {
        kind: 'merge-range',
        row: 3,
        col: 3,
        rows: 1,
        cols: 1,
      }),
    ).toBeNull()
  })

  it('merging over an existing merge that is fully inside extends it', async () => {
    const { opened, tblId } = await makeTable()
    const r1 = mergeTableCells(opened, 0, tblId, {
      kind: 'merge-range',
      row: 0,
      col: 0,
      rows: 1,
      cols: 2,
    })!
    const r2 = mergeTableCells(opened, 0, r1.elementId, {
      kind: 'merge-range',
      row: 0,
      col: 0,
      rows: 2,
      cols: 3,
    })!
    const t = tableOf(opened, r2.elementId)
    expect(t.rows[0]![0]!.gridSpan).toBe(3)
    expect(t.rows[0]![0]!.rowSpan).toBe(2)
    expect(textOf(t, 0, 0)).toBe('00|01|02|10|11|12')
  })

  it('split-range clears every merge inside the rectangle', async () => {
    const { opened, tblId } = await makeTable()
    const r1 = mergeTableCells(opened, 0, tblId, {
      kind: 'merge-range',
      row: 1,
      col: 1,
      rows: 2,
      cols: 2,
    })!
    const r2 = mergeTableCells(opened, 0, r1.elementId, {
      kind: 'split-range',
      row: 0,
      col: 0,
      rows: 4,
      cols: 4,
    })!
    const t = tableOf(opened, r2.elementId)
    for (const row of t.rows)
      for (const cell of row) {
        expect(cell.merged ?? false).toBe(false)
        expect(cell.gridSpan ?? 1).toBe(1)
        expect(cell.rowSpan ?? 1).toBe(1)
      }
    // nothing merged any more: splitting again is refused
    expect(
      mergeTableCells(opened, 0, r2.elementId, {
        kind: 'split-range',
        row: 0,
        col: 0,
        rows: 4,
        cols: 4,
      }),
    ).toBeNull()
  })
})

describe('editTableStructure count', () => {
  it('inserts N rows before / after and deletes N rows', async () => {
    const { opened, tblId } = await makeTable(3, 2)
    const before = tableOf(opened, tblId)
    const h = before.rowHeights.reduce((a, b) => a + b, 0)
    const r1 = editTableStructure(opened, 0, tblId, {
      kind: 'insert-row',
      index: 1,
      before: true,
      count: 2,
    })!
    const t1 = tableOf(opened, r1.elementId)
    expect(t1.rows.length).toBe(5)
    expect(textOf(t1, 3, 0)).toBe('10')
    expect(t1.rowHeights.reduce((a, b) => a + b, 0)).toBeGreaterThan(h)
    const r2 = editTableStructure(opened, 0, r1.elementId, {
      kind: 'delete-row',
      index: 1,
      count: 2,
    })!
    expect(tableOf(opened, r2.elementId).rows.length).toBe(3)
    expect(textOf(tableOf(opened, r2.elementId), 1, 0)).toBe('10')
    // never deletes the last row
    expect(
      editTableStructure(opened, 0, r2.elementId, { kind: 'delete-row', index: 0, count: 3 }),
    ).toBeNull()
  })

  it('inserts and deletes N columns, keeping the frame width in sync', async () => {
    const { opened, tblId } = await makeTable(2, 3)
    const w0 = tableOf(opened, tblId).colWidths.reduce((a, b) => a + b, 0)
    const r1 = editTableStructure(opened, 0, tblId, {
      kind: 'insert-col',
      index: 2,
      count: 2,
    })!
    const t1 = tableOf(opened, r1.elementId)
    expect(t1.colWidths.length).toBe(5)
    expect(t1.transform.offset.cx).toBe(t1.colWidths.reduce((a, b) => a + b, 0))
    expect(t1.transform.offset.cx).toBeGreaterThan(w0)
    const r2 = editTableStructure(opened, 0, r1.elementId, {
      kind: 'delete-col',
      index: 0,
      count: 2,
    })!
    const t2 = tableOf(opened, r2.elementId)
    expect(t2.colWidths.length).toBe(3)
    expect(textOf(t2, 0, 0)).toBe('02')
  })
})

describe('cell-scoped formatting and clearing', () => {
  it('setElementFont with cells touches only those cells', async () => {
    const { opened, tblId } = await makeTable(2, 3)
    const slide = opened.deck.slides[0]!
    expect(
      setElementFont(slide, tblId, { bold: true, fontSizePt: 30 }, [
        { row: 0, col: 1 },
        { row: 1, col: 2 },
      ]),
    ).toBe(true)
    const t = slide.elements.find((e) => e.id === tblId) as TableElement
    const run = (r: number, c: number) => t.rows[r]![c]!.text!.paragraphs[0]!.runs[0]!
    expect(run(0, 1).bold).toBe(true)
    expect(run(0, 1).fontSize).toBe(30)
    expect(run(1, 2).bold).toBe(true)
    expect(run(0, 0).bold).not.toBe(true)
    expect(run(1, 1).fontSize).not.toBe(30)
    // and the XML carries it
    const reopened = await openPptx(await savePptx(opened))
    const t2 = reopened.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(t2.rows[0]![1]!.text!.paragraphs[0]!.runs[0]!.fontSize).toBe(30)
    expect(t2.rows[0]![0]!.text!.paragraphs[0]!.runs[0]!.fontSize).not.toBe(30)
  })

  it('setElementParagraphFormat with cells aligns only those cells', async () => {
    const { opened, tblId } = await makeTable(2, 2)
    const slide = opened.deck.slides[0]!
    expect(
      setElementParagraphFormat(slide, tblId, { align: 'center' }, undefined, [{ row: 1, col: 0 }]),
    ).toBe(true)
    const t = slide.elements.find((e) => e.id === tblId) as TableElement
    expect(t.rows[1]![0]!.text!.paragraphs[0]!.align).toBe('center')
    expect(t.rows[0]![0]!.text!.paragraphs[0]!.align).not.toBe('center')
  })

  it('clearTableCellText empties text but keeps the run formatting', async () => {
    const { opened, tblId } = await makeTable(2, 2)
    const slide = opened.deck.slides[0]!
    setElementFont(slide, tblId, { bold: true, color: '#FF0000' }, [{ row: 0, col: 0 }])
    expect(
      clearTableCellText(slide, tblId, [
        { row: 0, col: 0 },
        { row: 1, col: 1 },
      ]),
    ).toBe(true)
    const t = slide.elements.find((e) => e.id === tblId) as TableElement
    expect(textOf(t, 0, 0)).toBe('')
    expect(textOf(t, 1, 1)).toBe('')
    expect(textOf(t, 0, 1)).toBe('01')
    const run = t.rows[0]![0]!.text!.paragraphs[0]!.runs[0]!
    expect(run.bold).toBe(true)
    expect(run.color).toBe('#FF0000')
    const reopened = await openPptx(await savePptx(opened))
    const t2 = reopened.deck.slides[0]!.elements.find((e) => e.type === 'table') as TableElement
    expect(textOf(t2, 0, 0)).toBe('')
    expect(textOf(t2, 1, 0)).toBe('10')
  })
})
