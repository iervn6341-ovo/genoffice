import { HorizontalAlign, VerticalAlign, WrapStrategy, type ICellData } from '@univerjs/core'
import { Font, FontCache } from '@univerjs/engine-render'
import { describe, expect, it, vi } from 'vitest'

import { createEditJournal, recordNeutralStyleEdit } from '../src/renderer/edit-journal'
import { selectionProtection } from '../src/renderer/selection-protection'
import {
  applyJournalOverlay,
  inheritedStyleIndexLookup,
  inheritedVerticalLookup,
  inheritedWrapLookup,
  patchWorksheetRangeInner,
  reapplyLiveMultilineVisual,
} from '../src/renderer/univer-sync'
import {
  VERTICAL_DISTRIBUTION_KEY,
  installVerticalDistributionRender,
  verticalDistributionOffsets,
  type VerticalDistribution,
} from '../src/renderer/vertical-distribution'
import type { WorkbookCellStyle, WorkbookFile } from '../src/shared/desktop-api'

const base: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: true,
  diagonalUp: false,
  diagonalDown: false,
  verticalAlignment: 'justify',
}

function renderedCell(style: WorkbookCellStyle, pending?: 'distributed' | 'top'): ICellData {
  let matrix: ICellData[][] = []
  const journal = createEditJournal()
  if (pending) recordNeutralStyleEdit(journal, 'sheet-1', 0, 0, { verticalAlignment: pending })
  const worksheet = {
    getSheetId: () => 'sheet-1',
    getRange: () => ({ setValues: (values: ICellData[][]) => (matrix = values) }),
    getSheet: () => ({ getMergedCell: () => null, getColumnWidth: () => 90 }),
  }
  patchWorksheetRangeInner(
    worksheet as never,
    undefined,
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    [{ row: 0, column: 0, value: 'One\nTwo\nThree', styleIndex: 0 }],
    [style],
    [],
    [],
    null,
    false,
    undefined,
    undefined,
    undefined,
    undefined,
    undefined,
    journal,
  )
  return matrix[0]?.[0] ?? {}
}

describe('vertical multiline distribution geometry', () => {
  const lines = [
    { top: 0, lineHeight: 12 },
    { top: 12, lineHeight: 12 },
    { top: 24, lineHeight: 12 },
  ]

  it('pins the first and last line to the inner bounds for Justify', () => {
    const offsets = verticalDistributionOffsets(lines, 76, 'justify')
    expect(offsets).toEqual([0, 20, 40])
    expect(lines[2]!.top + lines[2]!.lineHeight + offsets![2]!).toBe(76)
  })

  it('uses half-gaps at the top and bottom for Distributed', () => {
    const offsets = verticalDistributionOffsets(lines, 76, 'distributed')!
    expect(offsets[0]).toBeCloseTo(40 / 6)
    expect(offsets[1]).toBeCloseTo(20)
    expect(offsets[2]).toBeCloseTo((40 * 5) / 6)
    expect(offsets[0]).toBeCloseTo(76 - (lines[2]!.top + lines[2]!.lineHeight + offsets[2]!))
  })

  it('never squeezes overflowing text, and skips single or huge line groups', () => {
    expect(verticalDistributionOffsets(lines, 24, 'justify')).toEqual([0, 0, 0])
    expect(verticalDistributionOffsets(lines.slice(0, 1), 76, 'justify')).toBeNull()
    expect(
      verticalDistributionOffsets(
        Array.from({ length: 257 }, () => lines[0]!),
        76,
        'justify',
      ),
    ).toBeNull()
  })

  it('marks imported and pending native modes while keeping the original multiline content', () => {
    const imported = renderedCell(base)
    expect(imported.custom?.[VERTICAL_DISTRIBUTION_KEY]).toBe('justify')
    expect(imported.p?.body?.dataStream).toContain('One\rTwo\rThree')
    const edited = renderedCell(base, 'distributed')
    expect(edited.custom?.[VERTICAL_DISTRIBUTION_KEY]).toBe('distributed')
    expect(edited.p?.body?.dataStream).toBe(imported.p?.body?.dataStream)
    expect(renderedCell(base, 'top').custom?.[VERTICAL_DISTRIBUTION_KEY]).toBeUndefined()
  })

  it('replays newly typed multiline text into an inherited blank with wrap and distribution', () => {
    const journal = createEditJournal()
    journal.cells.set(
      'sheet-1',
      new Map([['7:6', { row: 7, column: 6, hasValue: true, value: 'ONE\nTWO\nTHREE' }]]),
    )
    const cell: ICellData = {}
    const writes: ICellData[] = []
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getSheet: () => ({ getComposedCellStyle: () => ({}), getCell: () => cell }),
      getRange: () => ({
        setValues: (values: ICellData[][]) => {
          const patch = values[0]?.[0] ?? {}
          writes.push(patch)
          Object.assign(cell, patch)
        },
      }),
    }
    applyJournalOverlay(
      worksheet as never,
      journal,
      { startRow: 7, endRow: 7, startColumn: 6, endColumn: 6 },
      () => ({ style: base, inherited: true }),
    )
    expect(cell.p?.body?.dataStream).toContain('ONE\rTWO\rTHREE')
    expect(writes).toContainEqual({ s: { tb: WrapStrategy.WRAP } })
    expect(cell.custom?.[VERTICAL_DISTRIBUTION_KEY]).toBe('justify')
  })

  it('restores inherited rendering after undo then redo restores only the live text', () => {
    const cell: ICellData = {}
    const worksheet = {
      getSheet: () => ({ getCell: () => cell }),
      getRange: () => ({
        setValues: (values: ICellData[][]) => {
          const patch = values[0]?.[0] ?? {}
          if (patch.s) cell.s = patch.s
          if (patch.custom) cell.custom = { ...cell.custom, ...patch.custom }
        },
      }),
    }
    const visual = { style: base, inherited: true }
    reapplyLiveMultilineVisual(worksheet as never, 7, 6, visual)
    expect(cell.custom).toBeUndefined()

    // Univer's redo command replays the user's original value mutation.
    // The canvas marker was installed under journal suppression and is not
    // part of that undo/redo entry.
    cell.v = 'ONE\nTWO\nTHREE'
    reapplyLiveMultilineVisual(worksheet as never, 7, 6, visual)
    expect(cell.s).toMatchObject({ tb: WrapStrategy.WRAP })
    expect(cell.custom?.[VERTICAL_DISTRIBUTION_KEY]).toBe('justify')
  })

  it('uses cell XF before row, row before column, and the last styled column span', () => {
    const styles: WorkbookCellStyle[] = [
      { ...base, wrapText: false, verticalAlignment: 'bottom' },
      { ...base, verticalAlignment: 'justify' },
      { ...base, verticalAlignment: 'distributed' },
      { ...base, verticalAlignment: 'top' },
    ]
    const inherited = inheritedVerticalLookup(
      styles,
      [
        { row: 0, styleIndex: 1 },
        { row: 1, styleIndex: 3 },
      ] as never,
      [
        { startColumn: 0, endColumn: 1, styleIndex: 2 },
        { startColumn: 1, endColumn: 2, styleIndex: 1 },
      ] as never,
    )
    expect(inherited(0, 0)).toBe('justify')
    expect(inherited(1, 0)).toBe('top')
    expect(inherited(2, 0)).toBe('distributed')
    expect(inherited(2, 1)).toBe('justify')
    expect(inherited(2, 3)).toBe('bottom')

    const sheet = {
      columnWidths: [
        { startColumn: 0, endColumn: 1, styleIndex: 2 },
        { startColumn: 1, endColumn: 2, styleIndex: 1 },
      ],
    } as WorkbookFile['sheets'][number]
    const selected = selectionProtection(
      styles,
      sheet,
      { cells: [], rows: [] },
      createEditJournal(),
      'sheet-1',
      { startRow: 2, endRow: 2, startColumn: 1, endColumn: 1 },
    )
    expect(selected.alignment.vertical).toBe('justify')

    const rowSelection = selectionProtection(
      styles,
      sheet,
      {
        cells: [
          { row: 0, column: 1, value: 'implicit', styleIndex: 0, explicitStyle: false },
          { row: 0, column: 2, value: 'explicit', styleIndex: 0, explicitStyle: true },
        ],
        rows: [{ row: 0, hidden: false, styleIndex: 1 }],
      },
      createEditJournal(),
      'sheet-1',
      { startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 },
    )
    expect(rowSelection.alignment.vertical).toBeNull()
    expect(rowSelection.cells.map((cell) => cell.verticalAlignment)).toEqual(['justify', 'bottom'])

    let matrix: ICellData[][] = []
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getRange: () => ({ setValues: (values: ICellData[][]) => (matrix = values) }),
      getSheet: () => ({ getMergedCell: () => null, getColumnWidth: () => 90 }),
    }
    patchWorksheetRangeInner(
      worksheet as never,
      undefined,
      { startRow: 0, endRow: 2, startColumn: 0, endColumn: 3 },
      [
        { row: 0, column: 0, value: 'ONE\nTWO', styleIndex: 3 },
        { row: 0, column: 1, value: 'ONE\nTWO', styleIndex: 0, explicitStyle: false },
        { row: 0, column: 2, value: 'ONE\nTWO', styleIndex: 0, explicitStyle: true },
        { row: 1, column: 0, value: 'ONE\nTWO', styleIndex: 0, explicitStyle: false },
        { row: 2, column: 0, value: 'ONE\nTWO', styleIndex: 0, explicitStyle: false },
        { row: 2, column: 1, value: 'ONE\nTWO', styleIndex: 0, explicitStyle: false },
        { row: 2, column: 3, value: 'ONE\nTWO', styleIndex: 0, explicitStyle: false },
      ],
      styles,
      [],
      [],
      null,
      false,
      undefined,
      undefined,
      inheritedWrapLookup(
        styles,
        [
          { row: 0, styleIndex: 1 },
          { row: 1, styleIndex: 3 },
        ] as never,
        sheet.columnWidths,
      ),
      undefined,
      undefined,
      undefined,
      inherited,
      inheritedStyleIndexLookup(
        [
          { row: 0, styleIndex: 1 },
          { row: 1, styleIndex: 3 },
        ] as never,
        sheet.columnWidths,
      ),
    )
    const mode = (row: number, column: number) =>
      matrix[row]?.[column]?.custom?.[VERTICAL_DISTRIBUTION_KEY]
    expect(mode(0, 0)).toBeUndefined() // Explicit top XF beats row justify.
    expect(mode(0, 1)).toBe('justify')
    expect(mode(0, 2)).toBeUndefined() // Explicit s="0" blocks the row xf.
    expect(mode(1, 0)).toBeUndefined() // Explicit row top blocks column distributed.
    expect(mode(2, 0)).toBe('distributed')
    expect(mode(2, 1)).toBe('justify')
    expect(mode(2, 3)).toBeUndefined() // Base XF bottom.
    expect(matrix[2]?.[0]?.p?.body?.dataStream).toContain('ONE\rTWO')
    expect(matrix[0]?.[1]?.s).toMatchObject({ tb: WrapStrategy.WRAP })
    expect(matrix[2]?.[0]?.s).toMatchObject({ tb: WrapStrategy.WRAP })
    expect(matrix[0]?.[1]?.p?.body?.dataStream).toContain('ONE\rTWO')
    expect(matrix[0]?.[2]?.p?.body?.dataStream).toBe('ONETWO\r\n')
  })

  it('passes real plain-text draw baselines through the Univer render hook', () => {
    const measurement = vi.spyOn(FontCache, 'getMeasureText').mockImplementation((text) => ({
      width: text.length * 6,
      fontBoundingBoxAscent: 9,
      fontBoundingBoxDescent: 3,
      actualBoundingBoxAscent: 9,
      actualBoundingBoxDescent: 3,
    }))
    const baselines: number[] = []
    const ctx = {
      save() {},
      restore() {},
      fillText(_text: string, _x: number, y: number) {
        baselines.push(y)
      },
    }
    const render = (mode: VerticalDistribution | null) => {
      baselines.length = 0
      const font = Font.prototype as unknown as {
        _renderText(ctx: unknown, row: number, col: number, context: unknown, cache: unknown): void
      }
      font._renderText.call(
        {},
        ctx,
        0,
        0,
        {
          startX: 0,
          startY: 0,
          endX: 90,
          endY: 76,
          fontCache: {
            cellData: { v: 'ONE\nTWO\nTHREE' },
            wrapStrategy: WrapStrategy.WRAP,
            horizontalAlign: HorizontalAlign.LEFT,
            verticalAlign: VerticalAlign.MIDDLE,
            fontString: '12px Arial',
            style: {},
          },
          spreadsheetSkeleton: {
            worksheet: {
              getCell: () => ({ custom: { [VERTICAL_DISTRIBUTION_KEY]: mode } }),
            },
          },
        },
        {},
      )
      return [...baselines]
    }
    try {
      installVerticalDistributionRender()
      const regular = render(null)
      const justified = render('justify')
      const distributed = render('distributed')
      expect(regular).toHaveLength(3)
      expect(justified[0]).toBeLessThan(regular[0]!)
      expect(justified[2]).toBeGreaterThan(regular[2]!)
      expect(distributed[0]).toBeGreaterThan(justified[0]!)
      expect(distributed[2]).toBeLessThan(justified[2]!)
    } finally {
      measurement.mockRestore()
    }
  })
})
