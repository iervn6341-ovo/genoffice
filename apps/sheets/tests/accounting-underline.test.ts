import { TextDecoration, VerticalAlign, type ICellData } from '@univerjs/core'
import { Font, Text } from '@univerjs/engine-render'
import { describe, expect, it, vi } from 'vitest'

import {
  ACCOUNTING_UNDERLINE_KEY,
  accountingUnderlineLines,
  accountingUnderlineMode,
  installAccountingUnderlineRender,
} from '../src/renderer/accounting-underline'
import { createEditJournal, recordNeutralStyleEdit } from '../src/renderer/edit-journal'
import { patchWorksheetRangeInner } from '../src/renderer/univer-sync'
import type { WorkbookCellStyle } from '../src/shared/desktop-api'

const style: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: true,
  underlineStyle: 'singleAccounting',
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
}

function renderedCell(
  styleInFile: WorkbookCellStyle,
  edit?: { underline: boolean; underlineStyle: WorkbookCellStyle['underlineStyle'] },
): ICellData {
  let matrix: ICellData[][] = []
  const journal = createEditJournal()
  if (edit) recordNeutralStyleEdit(journal, 'sheet-1', 0, 0, edit)
  const sheet = {
    getSheetId: () => 'sheet-1',
    getRange: () => ({ setValues: (values: ICellData[][]) => (matrix = values) }),
    getSheet: () => ({ getMergedCell: () => null, getColumnWidth: () => 140 }),
  }
  patchWorksheetRangeInner(
    sheet as never,
    undefined,
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    [{ row: 0, column: 0, value: 'Accounting', styleIndex: 0 }],
    [styleInFile],
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

describe('accounting underline canvas geometry and import marker', () => {
  it('distinguishes exact native accounting font styles from ordinary single/double', () => {
    expect(accountingUnderlineMode(style)).toBe('singleAccounting')
    expect(accountingUnderlineMode({ ...style, underlineStyle: 'doubleAccounting' })).toBe(
      'doubleAccounting',
    )
    expect(accountingUnderlineMode({ ...style, underlineStyle: 'single' })).toBeNull()
    expect(accountingUnderlineMode({ ...style, underline: false })).toBeNull()
  })

  it('spans the text area below descenders for both accounting variants', () => {
    const lines = [{ height: 12 }]
    expect(
      accountingUnderlineLines(
        lines,
        { left: 2, top: 2, width: 116, height: 18 },
        12,
        VerticalAlign.BOTTOM,
        'singleAccounting',
      ),
    ).toEqual([{ x: 2, y: 19, width: 116, lineType: TextDecoration.SINGLE }])
    expect(
      accountingUnderlineLines(
        lines,
        { left: 2, top: 2, width: 116, height: 18 },
        12,
        VerticalAlign.BOTTOM,
        'doubleAccounting',
      ),
    ).toEqual([{ x: 2, y: 17, width: 116, lineType: TextDecoration.DOUBLE }])
    expect(
      accountingUnderlineLines(
        [],
        { left: 0, top: 0, width: 100, height: 20 },
        0,
        VerticalAlign.BOTTOM,
        'singleAccounting',
      ),
    ).toEqual([])
  })

  it('tags imported and edited accounting cells without replacing their source text', () => {
    const imported = renderedCell(style)
    expect(imported.custom?.[ACCOUNTING_UNDERLINE_KEY]).toBe('singleAccounting')
    expect(imported.v).toBe('Accounting')
    expect(
      renderedCell(style, { underline: true, underlineStyle: 'doubleAccounting' }).custom?.[
        ACCOUNTING_UNDERLINE_KEY
      ],
    ).toBe('doubleAccounting')
    expect(
      renderedCell(style, { underline: true, underlineStyle: 'single' }).custom?.[
        ACCOUNTING_UNDERLINE_KEY
      ],
    ).toBeUndefined()
  })

  it('replaces only the tagged cell underline with a full-width canvas stroke', () => {
    const originalRender = (Font.prototype as any)._renderText
    const originalDraw = Text.drawWith
    const originalDecoration = (Text as any)._drawTextDecoration
    const skeleton = { calculate: () => [{ height: 12 }], getTotalHeight: () => 12 }
    const draw = vi.fn(() => 12)
    const decoration = vi.fn()
    ;(Font.prototype as any)._renderText = function () {
      Text.drawWith(
        {} as never,
        {
          text: 'A',
          fontStyle: '12px Arial',
          width: 120,
          height: 18,
          left: 2,
          top: 2,
          vAlign: VerticalAlign.BOTTOM,
          underline: true,
        } as never,
        skeleton as never,
      )
    }
    Text.drawWith = draw as never
    ;(Text as any)._drawTextDecoration = decoration
    try {
      installAccountingUnderlineRender()
      ;(Font.prototype as any)._renderText(
        {},
        0,
        0,
        {
          fontCache: {},
          spreadsheetSkeleton: {
            worksheet: {
              getCell: () => ({ custom: { [ACCOUNTING_UNDERLINE_KEY]: 'doubleAccounting' } }),
            },
          },
        },
        {},
      )
      expect(draw).toHaveBeenCalledWith(
        expect.anything(),
        expect.objectContaining({ text: 'A', underline: false }),
        skeleton,
      )
      expect(decoration).toHaveBeenCalledWith(
        expect.anything(),
        expect.objectContaining({ x: 2, width: 120, lineType: TextDecoration.DOUBLE }),
      )
    } finally {
      ;(Font.prototype as any)._renderText = originalRender
      Text.drawWith = originalDraw
      ;(Text as any)._drawTextDecoration = originalDecoration
    }
  })
})
