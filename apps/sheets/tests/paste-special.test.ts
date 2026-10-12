import { shiftRuleFormulas } from '../src/renderer/paste-extras'
import { describe, expect, it } from 'vitest'
import type { IStyleData } from '@univerjs/core'

import { columnIndex, columnName, transposeFormula } from '../src/renderer/formula-transpose'
import {
  DEFAULT_PASTE_SPECIAL,
  mergedRangesAfterPaste,
  needsOwnPaste,
  pasteExtras,
  pastedCells,
  planPaste,
  type CopiedSource,
  type CopiedSourceCell,
  type DestinationCell,
  type PasteSpecialOptions,
  type PlannedWrite,
} from '../src/renderer/paste-special'

// Every expectation was read from Excel 16 (Range.PasteSpecial via COM), see
// docs/qa/handoff-ribbon-audit-2026-10-11.md.

const NUM = 2
const STR = 1
const num = (v: number, extra: Partial<CopiedSourceCell> = {}): CopiedSourceCell => ({
  v,
  t: NUM,
  style: null,
  ...extra,
})
const text = (v: string): CopiedSourceCell => ({ v, t: STR, style: null })
const formula = (f: string, v: number): CopiedSourceCell => ({ v, t: NUM, f, style: null })

/** a copied block whose top-left is at (row, column) */
function block(cells: (CopiedSourceCell | null)[][], row = 0, column = 0): CopiedSource {
  return {
    rows: cells.map((_, i) => row + i),
    columns: (cells[0] ?? []).map((_, j) => column + j),
    cells,
  }
}

interface Run {
  readonly writes: readonly PlannedWrite[]
  /** value written at a cell: the constant, or the formula string */
  at(row: number, column: number): string | number | boolean | null | undefined
  write(row: number, column: number): PlannedWrite | undefined
}

function run(
  source: CopiedSource,
  options: Partial<PasteSpecialOptions>,
  target: { row: number; column: number; rows?: number; columns?: number },
  destination: Record<string, DestinationCell> = {},
  destinationStyle: Record<string, IStyleData> = {},
  sameSheet = true,
): Run {
  const plan = planPaste({
    source,
    options: { ...DEFAULT_PASTE_SPECIAL, ...options },
    target: {
      anchorRow: target.row,
      anchorColumn: target.column,
      selectionRows: target.rows ?? 1,
      selectionColumns: target.columns ?? 1,
    },
    sameSheet,
    readDestination: (row, column) => destination[`${row}:${column}`] ?? {},
    readDestinationStyle: (row, column) => destinationStyle[`${row}:${column}`] ?? null,
    // a stand-in for Univer's lexer: shifts every plain A1 reference
    shiftFormula: (f, columns, rows) =>
      f.replace(
        /\b([A-Z]+)(\d+)\b/g,
        (_m, c: string, r: string) => `${columnName(columnIndex(c) + columns)}${Number(r) + rows}`,
      ),
  })
  if ('error' in plan) throw new Error(`plan error: ${plan.error}`)
  const write = (row: number, column: number) =>
    plan.writes.find((w) => w.row === row && w.column === column)
  return {
    writes: plan.writes,
    write,
    at: (row, column) => {
      const value = write(row, column)?.value
      return value === undefined ? undefined : (value.f ?? value.v)
    },
  }
}

describe('transposeFormula (Excel: relative offsets swap, absolute refs inside the block follow)', () => {
  const block2x3 = { rows: [0, 1], columns: [0, 1, 2], anchorRow: 0, anchorColumn: 4 }

  it('=A2+B2 in C2, block A1:C2 pasted at E1 → =F1+F2 in F3', () => {
    expect(transposeFormula('=A2+B2', { row: 1, column: 2 }, { row: 2, column: 5 }, block2x3)).toBe(
      '=F1+F2',
    )
  })

  it('=$A$1*2 with A1:B1 pasted at D1 → =$D$1*2', () => {
    expect(
      transposeFormula(
        '=$A$1*2',
        { row: 0, column: 1 },
        { row: 1, column: 3 },
        { rows: [0], columns: [0, 1], anchorRow: 0, anchorColumn: 3 },
      ),
    ).toBe('=$D$1*2')
  })

  it('2×2: B2 =A1+A2 pasted at D1 lands in E2 as =D1+E1', () => {
    expect(
      transposeFormula(
        '=A1+A2',
        { row: 1, column: 1 },
        { row: 1, column: 4 },
        { rows: [0, 1], columns: [0, 1], anchorRow: 0, anchorColumn: 3 },
      ),
    ).toBe('=D1+E1')
  })

  it('leaves absolute references outside the block, mixed references, strings and function names alone', () => {
    const where = { rows: [0], columns: [0, 1], anchorRow: 0, anchorColumn: 3 }
    expect(transposeFormula('=$Z$9+1', { row: 0, column: 1 }, { row: 1, column: 3 }, where)).toBe(
      '=$Z$9+1',
    )
    expect(transposeFormula('=$A1+B$1', { row: 0, column: 1 }, { row: 1, column: 3 }, where)).toBe(
      '=$A1+B$1',
    )
    expect(
      transposeFormula('="A1"&LOG10(A1)', { row: 0, column: 1 }, { row: 1, column: 3 }, where),
    ).toContain('"A1"&LOG10(')
  })

  it('transposes a range reference and keeps its corners in order', () => {
    expect(
      transposeFormula(
        '=SUM(A1:A2)',
        { row: 2, column: 0 },
        { row: 0, column: 2 },
        { rows: [0, 1, 2], columns: [0], anchorRow: 0, anchorColumn: 0 },
      ),
    ).toBe('=SUM(A1:B1)')
  })

  it('a reference pushed off the sheet becomes #REF!', () => {
    expect(
      transposeFormula(
        '=A1',
        { row: 5, column: 5 },
        { row: 0, column: 0 },
        { rows: [5], columns: [5], anchorRow: 0, anchorColumn: 0 },
      ),
    ).toBe('=#REF!')
  })

  it('column letters round-trip', () => {
    for (const [name, index] of [
      ['A', 0],
      ['Z', 25],
      ['AA', 26],
      ['AZ', 51],
      ['XFD', 16383],
    ] as const) {
      expect(columnIndex(name)).toBe(index)
      expect(columnName(index)).toBe(name)
    }
  })
})

describe('Paste Special → Transpose', () => {
  const sheet = block([
    [num(1), num(2), num(3)],
    [num(4), num(5), formula('=A2+B2', 9)],
  ])

  it('turns a 2×3 block into 3×2 (Excel: 1 2 3 / 4 5 9 → E1:F3)', () => {
    const r = run(sheet, { transpose: true }, { row: 0, column: 4 })
    expect([r.at(0, 4), r.at(1, 4), r.at(2, 4)]).toEqual([1, 2, 3])
    expect([r.at(0, 5), r.at(1, 5)]).toEqual([4, 5])
    // the formula follows its cells: C2 =A2+B2 → F3 =F1+F2
    expect(r.at(2, 5)).toBe('=F1+F2')
  })

  it('is never tiled over a bigger selection (Excel: one block at the top-left)', () => {
    const one = block([[num(1), num(2)]])
    const r = run(one, { transpose: true }, { row: 0, column: 3, rows: 5, columns: 2 })
    expect(r.writes).toHaveLength(2)
    expect([r.at(0, 3), r.at(1, 3)]).toEqual([1, 2])
  })

  it('refuses a transposed block that overlaps its source, unless same size and shape', () => {
    const row = block([[num(1), num(2), num(3)]])
    expect(() => run(row, { transpose: true }, { row: 0, column: 0 })).toThrow('overlap')
    const square = block([
      [num(1), num(2)],
      [num(3), num(4)],
    ])
    const r = run(square, { transpose: true }, { row: 0, column: 0 })
    expect([r.at(0, 1), r.at(1, 0)]).toEqual([3, 2])
    // another sheet never overlaps
    expect(() => run(row, { transpose: true }, { row: 0, column: 0 }, {}, {}, false)).not.toThrow()
  })

  it('transposes formats with their cells', () => {
    const bold: IStyleData = { bl: 1 }
    const r = run(
      block([[num(1, { style: bold }), num(2)]]),
      { transpose: true },
      { row: 0, column: 3 },
    )
    expect(r.write(0, 3)?.style).toEqual({ mode: 'replace', style: bold })
    expect(r.write(1, 3)?.style).toEqual({ mode: 'replace', style: null })
  })
})

describe('Paste Special → Operation', () => {
  const dest = (...values: (number | string | null)[]): Record<string, DestinationCell> =>
    Object.fromEntries(values.map((v, i) => [`${i}:2`, v === null ? {} : { v }]))
  const col = (cells: number[]) => block(cells.map((v) => [num(v)]))

  it('Add / Subtract / Multiply / Divide act as destination ∘ source', () => {
    const cases: [PasteSpecialOptions['operation'], number[], number[]][] = [
      ['add', [1, 2], [11, 22]],
      ['subtract', [1, 2], [9, 18]],
      ['multiply', [3, 4], [30, 80]],
      ['divide', [2, 5], [5, 4]],
    ]
    for (const [operation, source, expected] of cases) {
      const r = run(
        col(source),
        { what: 'values', operation },
        { row: 0, column: 2, rows: 2 },
        dest(10, 20),
      )
      expect([r.at(0, 2), r.at(1, 2)]).toEqual(expected)
    }
  })

  it('a blank destination counts as 0: add 5, subtract −5, divide 0', () => {
    const five = col([5])
    expect(run(five, { what: 'values', operation: 'add' }, { row: 0, column: 2 }).at(0, 2)).toBe(5)
    expect(
      run(five, { what: 'values', operation: 'subtract' }, { row: 0, column: 2 }).at(0, 2),
    ).toBe(-5)
    expect(run(five, { what: 'values', operation: 'divide' }, { row: 0, column: 2 }).at(0, 2)).toBe(
      0,
    )
  })

  it('a text or blank source leaves the destination alone', () => {
    const r = run(
      block([[text('abc')], [null]]),
      { what: 'values', operation: 'add' },
      { row: 0, column: 2, rows: 2 },
      dest(5, 5),
    )
    expect(r.at(0, 2)).toBeUndefined()
    expect(r.at(1, 2)).toBeUndefined()
    expect(r.writes).toHaveLength(0)
  })

  it('a text destination stays; a formula destination is wrapped', () => {
    const one = col([1])
    expect(
      run(
        one,
        { what: 'values', operation: 'add' },
        { row: 0, column: 2 },
        { '0:2': { v: 'abc' } },
      ).at(0, 2),
    ).toBeUndefined()
    expect(
      run(
        one,
        { what: 'values', operation: 'add' },
        { row: 0, column: 2 },
        { '0:2': { v: 5, f: '=2+3' } },
      ).at(0, 2),
    ).toBe('=(2+3)+1')
  })

  it('dividing by zero yields a formula that shows #DIV/0!', () => {
    expect(
      run(col([0]), { what: 'values', operation: 'divide' }, { row: 0, column: 2 }, dest(10)).at(
        0,
        2,
      ),
    ).toBe('=10/0')
  })

  it('"All" with an operation still pastes the format; "Formats" ignores the operation', () => {
    const bold: IStyleData = { bl: 1 }
    const source = block([[num(1, { style: bold })]])
    const all = run(source, { what: 'all', operation: 'add' }, { row: 0, column: 2 }, dest(10))
    expect(all.at(0, 2)).toBe(11)
    expect(all.write(0, 2)?.style).toEqual({ mode: 'replace', style: bold })
    const formats = run(
      source,
      { what: 'formats', operation: 'add' },
      { row: 0, column: 2 },
      dest(10),
    )
    expect(formats.write(0, 2)?.value).toBeUndefined()
  })

  it('decimals do not drift: 0.1 + 0.2 → 0.3', () => {
    expect(
      run(col([0.2]), { what: 'values', operation: 'add' }, { row: 0, column: 2 }, dest(0.1)).at(
        0,
        2,
      ),
    ).toBe(0.3)
  })
})

describe('Paste Special → Skip blanks', () => {
  const source = block([[num(1)], [null]])
  const dest: Record<string, DestinationCell> = { '0:2': { v: 9 }, '1:2': { v: 8 } }

  it('skip: a blank source leaves the destination (Excel: 1, 8)', () => {
    const r = run(source, { skipBlanks: true }, { row: 0, column: 2, rows: 2 }, dest)
    expect(r.at(0, 2)).toBe(1)
    expect(r.write(1, 2)).toBeUndefined()
  })

  it('no skip: a blank source clears the destination (Excel: 1, blank)', () => {
    const r = run(source, {}, { row: 0, column: 2, rows: 2 }, dest)
    expect(r.at(0, 2)).toBe(1)
    expect(r.write(1, 2)?.value).toEqual({ v: null, t: null, f: null })
  })
})

describe('Paste Special → number formats and the rest of the menu', () => {
  const fmt: IStyleData = { n: { pattern: '0.00' }, bl: 1 }
  const source = block([[formula('=1/3', 1 / 3)].map((c) => ({ ...c, style: fmt }))])

  it('Values & Number Formats: the value as a constant, the number format only (Excel: 0.33, not bold)', () => {
    const r = run(source, { what: 'valuesAndNumberFormats' }, { row: 0, column: 1 })
    expect(r.write(0, 1)?.value).toEqual({ v: 1 / 3, t: NUM, f: null })
    expect(r.write(0, 1)?.style).toEqual({ mode: 'numfmt', pattern: '0.00' })
  })

  it('Formulas & Number Formats: the formula stays a formula', () => {
    const r = run(source, { what: 'formulasAndNumberFormats' }, { row: 0, column: 1 })
    expect(r.at(0, 1)).toBe('=1/3')
    expect(r.write(0, 1)?.style).toEqual({ mode: 'numfmt', pattern: '0.00' })
  })

  it('a source without a number format resets the destination to General', () => {
    const r = run(block([[num(1)]]), { what: 'valuesAndNumberFormats' }, { row: 0, column: 1 })
    expect(r.write(0, 1)?.style).toEqual({ mode: 'numfmt', pattern: null })
  })

  it('Values drops the formula and every format; Formulas keeps the formula and no format', () => {
    const values = run(source, { what: 'values' }, { row: 0, column: 1 })
    expect(values.write(0, 1)?.value).toEqual({ v: 1 / 3, t: NUM, f: null })
    expect(values.write(0, 1)?.style).toBeUndefined()
    const formulas = run(source, { what: 'formulas' }, { row: 0, column: 1 })
    expect(formulas.at(0, 1)).toBe('=1/3')
    expect(formulas.write(0, 1)?.style).toBeUndefined()
  })

  it('All except borders keeps the destination borders', () => {
    const border = { t: { s: 1 } } as unknown as IStyleData['bd']
    const r = run(
      block([[num(1, { style: { bl: 1, bd: { b: { s: 1 } } as IStyleData['bd'] } })]]),
      { what: 'allExceptBorders', operation: 'add' },
      { row: 0, column: 2 },
      { '0:2': { v: 1 } },
      { '0:2': { bd: border } as IStyleData },
    )
    expect(r.write(0, 2)?.style).toEqual({ mode: 'replace', style: { bl: 1, bd: border } })
  })

  it('formulas move with the paste when not transposed (relative refs shift)', () => {
    const r = run(
      block([[formula('=A1*2', 2)]], 0, 1),
      { what: 'formulasAndNumberFormats' },
      { row: 3, column: 5 },
    )
    expect(r.at(3, 5)).toBe('=E4*2')
  })

  it('tiles a small source over a selection that is a whole multiple (Excel: 7 → 2×2 of 7)', () => {
    const r = run(
      block([[num(7)]]),
      { what: 'values', skipBlanks: true },
      { row: 0, column: 2, rows: 2, columns: 2 },
    )
    expect(r.writes.map((w) => [w.row, w.column])).toEqual([
      [0, 2],
      [0, 3],
      [1, 2],
      [1, 3],
    ])
  })

  it('a selection that is not a multiple pastes once at the anchor', () => {
    const r = run(
      block([[num(1), num(2)]]),
      { what: 'values', skipBlanks: true },
      { row: 0, column: 0, rows: 1, columns: 3 },
    )
    expect(r.writes).toHaveLength(2)
  })

  it('knows which choices Univer cannot do itself', () => {
    expect(needsOwnPaste(DEFAULT_PASTE_SPECIAL)).toBe(false)
    expect(needsOwnPaste({ ...DEFAULT_PASTE_SPECIAL, what: 'values' })).toBe(false)
    expect(needsOwnPaste({ ...DEFAULT_PASTE_SPECIAL, transpose: true })).toBe(true)
    expect(needsOwnPaste({ ...DEFAULT_PASTE_SPECIAL, skipBlanks: true })).toBe(true)
    expect(needsOwnPaste({ ...DEFAULT_PASTE_SPECIAL, operation: 'add' })).toBe(true)
    expect(needsOwnPaste({ ...DEFAULT_PASTE_SPECIAL, what: 'valuesAndNumberFormats' })).toBe(true)
    expect(needsOwnPaste({ ...DEFAULT_PASTE_SPECIAL, what: 'columnWidths' })).toBe(false)
  })
})

describe('transposeFormula leaves what is not a reference into the block', () => {
  const where = {
    rows: [0, 1],
    columns: [0, 1],
    anchorRow: 0,
    anchorColumn: 3,
    sheetName: 'Sheet1',
  }
  const source = { row: 1, column: 1 }
  const dest = { row: 1, column: 4 }

  it('does not touch column names inside structured references', () => {
    expect(transposeFormula('=SUM(Sales[Q1])+A1', source, dest, where)).toBe('=SUM(Sales[Q1])+D1')
    expect(transposeFormula('=Table1[@FY24]*2', source, dest, where)).toBe('=Table1[@FY24]*2')
    expect(transposeFormula('=Sales[[#This Row],[Q1]]+B2', source, dest, where)).toBe(
      '=Sales[[#This Row],[Q1]]+E2',
    )
  })

  it('an absolute reference to another sheet is not a reference into the block', () => {
    expect(transposeFormula('=Sheet2!$A$1+$A$1', source, dest, where)).toBe('=Sheet2!$A$1+$D$1')
    expect(transposeFormula("='My Sheet'!$B$2", source, dest, where)).toBe("='My Sheet'!$B$2")
    // the block's own sheet, named, still follows its cells
    expect(transposeFormula('=Sheet1!$A$1', source, dest, where)).toBe('=Sheet1!$D$1')
  })
})

describe('what a choice carries besides the cells', () => {
  const none = DEFAULT_PASTE_SPECIAL
  it('lists merges for All, Formats and All except borders only (read from Excel)', () => {
    const withMerges = ['all', 'formats', 'allExceptBorders', 'allUsingSourceTheme'] as const
    for (const what of withMerges) expect(pasteExtras(what).merges).toBe(true)
    for (const what of [
      'values',
      'formulas',
      'valuesAndNumberFormats',
      'formulasAndNumberFormats',
      'columnWidths',
      'comments',
      'validation',
    ] as const) {
      expect(pasteExtras(what).merges).toBe(false)
    }
  })

  it('Comments and Validation are their own paste; All carries both', () => {
    expect(pasteExtras('comments')).toEqual({ notes: true, validation: false, merges: false })
    expect(pasteExtras('validation')).toEqual({ notes: false, validation: true, merges: false })
    expect(pasteExtras('all')).toEqual({ notes: true, validation: true, merges: true })
    expect(pasteExtras('formats')).toEqual({ notes: false, validation: false, merges: true })
    expect(needsOwnPaste({ ...none, what: 'comments' })).toBe(true)
    expect(needsOwnPaste({ ...none, what: 'validation' })).toBe(true)
    expect(needsOwnPaste({ ...none, what: 'allMergingConditionalFormats' })).toBe(true)
    // within one workbook there is one theme: this is plain All
    expect(needsOwnPaste({ ...none, what: 'allUsingSourceTheme' })).toBe(false)
  })

  it('a Comments or Validation paste writes no cell value or format', () => {
    for (const what of ['comments', 'validation'] as const) {
      const result = run(block([[num(1), num(2)]]), { what }, { row: 5, column: 0 })
      expect(result.writes).toEqual([])
    }
  })

  it('maps every pasted cell back to its source cell, through Transpose and tiling', () => {
    const source = { rows: [2, 3], columns: [1, 2, 3] }
    const straight = pastedCells(source, none, {
      anchorRow: 10,
      anchorColumn: 5,
      selectionRows: 1,
      selectionColumns: 1,
    })
    expect(straight).toHaveLength(6)
    expect(straight.find((cell) => cell.sourceRow === 3 && cell.sourceColumn === 2)).toEqual({
      sourceRow: 3,
      sourceColumn: 2,
      row: 11,
      column: 6,
    })
    const turned = pastedCells(
      source,
      { ...none, transpose: true },
      {
        anchorRow: 10,
        anchorColumn: 5,
        selectionRows: 1,
        selectionColumns: 1,
      },
    )
    expect(turned.find((cell) => cell.sourceRow === 3 && cell.sourceColumn === 2)).toEqual({
      sourceRow: 3,
      sourceColumn: 2,
      row: 11,
      column: 6,
    })
    expect(turned.find((cell) => cell.sourceRow === 2 && cell.sourceColumn === 3)).toEqual({
      sourceRow: 2,
      sourceColumn: 3,
      row: 12,
      column: 5,
    })
    // a 2 × 3 block over a 4 × 6 selection is laid out four times
    const tiled = pastedCells(source, none, {
      anchorRow: 0,
      anchorColumn: 0,
      selectionRows: 4,
      selectionColumns: 6,
    })
    expect(tiled).toHaveLength(24)
  })

  it('moves merged ranges to every tile and turns them with Transpose (Excel: A1:B1 → D13:D14)', () => {
    const source = { rows: [0], columns: [0, 1] }
    const target = { anchorRow: 12, anchorColumn: 3, selectionRows: 1, selectionColumns: 1 }
    const merge = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }
    expect(mergedRangesAfterPaste([merge], source, none, target)).toEqual([
      { startRow: 12, endRow: 12, startColumn: 3, endColumn: 4 },
    ])
    expect(mergedRangesAfterPaste([merge], source, { ...none, transpose: true }, target)).toEqual([
      { startRow: 12, endRow: 13, startColumn: 3, endColumn: 3 },
    ])
    // repeated over a selection that is a whole multiple of the block
    expect(
      mergedRangesAfterPaste([merge], source, none, {
        ...target,
        selectionRows: 3,
        selectionColumns: 2,
      }),
    ).toHaveLength(3)
  })

  it('does not carry a merge that is only partly inside the block, or a 1 × 1 one', () => {
    const source = { rows: [0, 1], columns: [0, 1] }
    const target = { anchorRow: 5, anchorColumn: 5, selectionRows: 1, selectionColumns: 1 }
    expect(
      mergedRangesAfterPaste(
        [
          { startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 },
          { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
        ],
        source,
        none,
        target,
      ),
    ).toEqual([])
  })

  it('does not carry a merge across rows a filter hid', () => {
    // rows 0 and 2 were copied (row 1 is filtered out): a merge over rows 0–2 is not contiguous in the block
    const source = { rows: [0, 2], columns: [0] }
    const target = { anchorRow: 9, anchorColumn: 0, selectionRows: 1, selectionColumns: 1 }
    expect(
      mergedRangesAfterPaste(
        [{ startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }],
        source,
        none,
        target,
      ),
    ).toEqual([])
  })
})

describe('a conditional format or validation rule pasted elsewhere moves its formulas with it', () => {
  it('shifts relative references by the distance the rule moved, and leaves absolute ones', () => {
    const cf = {
      rule: { type: 'highlightCell', subType: 'formula', value: '=A1>$B$1', style: { bg: 'x' } },
    }
    expect(shiftRuleFormulas(cf, 0, 3).rule.value).toBe('=D1>$B$1')
    expect(shiftRuleFormulas(cf, 2, 0).rule.value).toBe('=A3>$B$1')
    const dv = { type: 'custom', formula1: '=COUNTIF(A$1:A$9,A1)=1', formula2: '' }
    expect(shiftRuleFormulas(dv, 1, 1).formula1).toBe('=COUNTIF(B$1:B$9,B2)=1')
  })

  it('leaves plain values and rules that did not move as they are', () => {
    const list = { type: 'list', formula1: 'a,b,c' }
    expect(shiftRuleFormulas(list, 5, 5)).toEqual(list)
    const same = { value: '=A1>1' }
    expect(shiftRuleFormulas(same, 0, 0)).toBe(same)
  })
})
