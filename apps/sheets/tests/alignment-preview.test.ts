import { HorizontalAlign, WrapStrategy, type ICellData, type IStyleData } from '@univerjs/core'
import { FontCache } from '@univerjs/engine-render'
import { beforeAll, describe, expect, it } from 'vitest'

import { CENTER_ACROSS_END_KEY } from '../src/renderer/center-continuous'
import {
  createEditJournal,
  fromNeutralStyle,
  recordNeutralStyleEdit,
} from '../src/renderer/edit-journal'
import {
  patchWorksheetRangeInner,
  repeatFillPreview,
  toUniverStyle,
} from '../src/renderer/univer-sync'
import type { WorkbookCellStyle, WorkbookRangeResult } from '../src/shared/desktop-api'

const base: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
  horizontalAlignment: 'left',
  verticalAlignment: 'bottom',
  fontSize: 10,
}

function patched(
  cells: WorkbookRangeResult['cells'],
  journal = createEditJournal(),
  styles: readonly WorkbookCellStyle[] = [base],
): ICellData[][] {
  let result: ICellData[][] = []
  const worksheet = {
    getSheetId: () => 'sheet-1',
    getRange: () => ({ setValues: (matrix: ICellData[][]) => (result = matrix) }),
    getSheet: () => ({ getMergedCell: () => null, getColumnWidth: () => 80 }),
  }
  patchWorksheetRangeInner(
    worksheet as never,
    undefined,
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
    cells,
    styles,
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
  return result
}

describe('advanced alignment preview', () => {
  beforeAll(() => {
    ;(FontCache as unknown as { _context: unknown })._context = {
      font: '',
      textBaseline: 'alphabetic',
      measureText(text: string) {
        return {
          width: text.length * 7,
          fontBoundingBoxAscent: 10,
          fontBoundingBoxDescent: 3,
        }
      },
    }
  })

  it('repeats Fill text only in the display document and keeps the raw value', () => {
    expect(repeatFillPreview('AB', 50, (value) => value.length * 7)).toBe('ABABABAB')
    const matrix = patched(
      [{ row: 0, column: 0, value: 'AB', styleIndex: 0 }],
      createEditJournal(),
      [{ ...base, horizontalAlignment: 'fill' }],
    )
    const cell = matrix[0]?.[0]
    expect(cell?.v).toBe('AB')
    expect(cell?.p?.body?.dataStream).toContain('ABAB')
    expect((cell?.s as IStyleData).tb).toBe(WrapStrategy.CLIP)
  })

  it('previews a pending Center Across range without merging or changing text', () => {
    const journal = createEditJournal()
    recordNeutralStyleEdit(journal, 'sheet-1', 0, 0, {
      horizontalAlignment: 'centerContinuous',
    })
    recordNeutralStyleEdit(journal, 'sheet-1', 0, 1, {
      horizontalAlignment: 'centerContinuous',
    })
    const matrix = patched([{ row: 0, column: 0, value: 'Title', styleIndex: 0 }], journal)
    expect(matrix[0]?.[0]?.v).toBe('Title')
    expect(matrix[0]?.[0]?.custom).toEqual({ [CENTER_ACROSS_END_KEY]: 1 })
    expect((matrix[0]?.[0]?.s as IStyleData).ht).toBe(HorizontalAlign.CENTER)
    expect((matrix[0]?.[0]?.s as IStyleData).tb).toBe(WrapStrategy.OVERFLOW)
    expect(matrix[0]?.[1]?.v).toBeUndefined()
  })

  it('previews pending rotation, shrink, and vertical alignment while retaining text', () => {
    const journal = createEditJournal()
    recordNeutralStyleEdit(journal, 'sheet-1', 0, 0, {
      verticalAlignment: 'distributed',
      textRotation: 127,
      shrinkToFit: true,
    })
    const cell = patched(
      [{ row: 0, column: 0, value: 'A very long label', styleIndex: 0 }],
      journal,
    )[0]?.[0]
    expect(cell?.v).toBe('A very long label')
    expect((cell?.s as IStyleData).tr).toEqual({ a: -37 })
    expect((cell?.s as IStyleData).fs).toBeLessThan(10)
  })

  it('uses explicit Univer General (0) to clear an earlier Left alignment', () => {
    expect(fromNeutralStyle({ horizontalAlignment: 'general' })).toEqual({ ht: 0 })
    expect(toUniverStyle({ ...base, horizontalAlignment: 'general' }).ht).toBe(0)
    const journal = createEditJournal()
    recordNeutralStyleEdit(journal, 'sheet-1', 0, 0, { horizontalAlignment: 'general' })
    const cell = patched([{ row: 0, column: 0, value: 'Text', styleIndex: 0 }], journal)[0]?.[0]
    expect(cell?.v).toBe('Text')
    expect((cell?.s as IStyleData).ht).toBe(0)
    recordNeutralStyleEdit(journal, 'sheet-1', 0, 0, { horizontalAlignment: 'center' })
    expect(
      (
        patched([{ row: 0, column: 0, value: 'Text', styleIndex: 0 }], journal)[0]?.[0]
          ?.s as IStyleData
      ).ht,
    ).toBe(2)
  })
})
