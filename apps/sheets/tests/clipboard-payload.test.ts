import { describe, expect, it } from 'vitest'

import {
  PAYLOAD_ATTR,
  PAYLOAD_VERSION,
  createStyleTable,
  decodePayload,
  encodePayload,
  extractPayload,
  type ClipboardPayload,
} from '../src/renderer/clipboard-payload'
import {
  foreignFromPayload,
  foreignFromText,
  foreignFormula,
  hasSheetReference,
  virtualSourceSheet,
} from '../src/renderer/clipboard-foreign'

const payload = (): ClipboardPayload => {
  const styles = createStyleTable()
  const bold = styles.indexOf({ bl: 1, bg: { rgb: '#FFC000' } })
  return {
    version: PAYLOAD_VERSION,
    bookId: 'book-1',
    sheetId: 'sheet-1',
    sheetName: 'Data 中文',
    rows: [2, 3],
    columns: [1, 2, 3],
    styles: styles.styles,
    cells: [
      [{ v: 'Total ✓', t: 1, s: bold! }, { v: 3, t: 2, f: '=B3+C3' }, 0],
      [0, { v: 'x', t: 1 }, { v: 1.5, t: 2, s: styles.indexOf({ n: { pattern: '0.00' } })! }],
    ],
    merges: [{ startRow: 2, endRow: 2, startColumn: 1, endColumn: 2 }],
    columnWidths: [100, 80, 60],
    rowHeights: [20, 30],
    notes: [{ row: 2, column: 1, text: 'note' }],
    validations: [
      {
        uid: 'dv1',
        type: 'list',
        ranges: [{ startRow: 3, endRow: 3, startColumn: 2, endColumn: 2 }],
      },
    ],
    conditionalFormats: [],
    links: [{ row: 2, column: 1, target: 'https://example.com' }],
  }
}

describe('clipboard payload', () => {
  it('round-trips through the HTML attribute value, including non-ASCII text', () => {
    const encoded = encodePayload(payload())!
    expect(encoded).toMatch(/^[A-Za-z0-9+/=]+$/)
    expect(decodePayload(encoded)).toEqual(payload())
    const html = `<table data-copy-id="abc" ${PAYLOAD_ATTR}="${encoded}"><tr><td>x</td></tr></table>`
    expect(extractPayload(html)).toEqual(payload())
  })

  it('shares one style entry between cells that look alike', () => {
    const table = createStyleTable()
    const a = table.indexOf({ bl: 1 })
    const b = table.indexOf({ bl: 1 })
    expect(a).toBe(b)
    expect(table.styles).toHaveLength(1)
    expect(table.indexOf({})).toBeUndefined()
    expect(table.indexOf(null)).toBeUndefined()
  })

  it('rejects text that is not a payload of this version or is malformed', () => {
    expect(decodePayload('not base64 !!')).toBeNull()
    expect(decodePayload(btoa('{"version":99}'))).toBeNull()
    const bad = { ...payload(), cells: [[0]] }
    expect(decodePayload(Buffer.from(JSON.stringify(bad)).toString('base64'))).toBeNull()
    expect(extractPayload('<table><tr><td>plain</td></tr></table>')).toBeNull()
  })

  it('is left out when the block is too big to carry', () => {
    const big = payload()
    const cells = Array.from({ length: 2 }, () =>
      Array.from({ length: 3 }, () => ({ v: 'x'.repeat(8_000_000), t: 1 })),
    )
    expect(encodePayload({ ...big, cells })).toBeNull()
  })
})

describe('foreign copy from a payload', () => {
  it('rebuilds cells with their styles, formulas and the features of the block', () => {
    const foreign = foreignFromPayload(payload())
    expect(foreign.origin).toBe('genoffice')
    expect(foreign.copied.rows).toEqual([2, 3])
    expect(foreign.copied.cells[0]![0]).toMatchObject({
      v: 'Total ✓',
      s: { bl: 1, bg: { rgb: '#FFC000' } },
    })
    expect(foreign.copied.cells[0]![1]).toMatchObject({ f: '=B3+C3', v: 3 })
    expect(foreign.copied.cells[0]![2]).toBeNull()
    expect(foreign.copied.cells[1]![2]).toMatchObject({ s: { n: { pattern: '0.00' } } })
    expect(foreign.merges).toHaveLength(1)
    expect(foreign.links[0]!.target).toBe('https://example.com')
    expect(foreign.literalFormulas).toBe(false)
  })

  it('stands in for a source sheet when the extras are pasted', () => {
    const sheet = virtualSourceSheet(foreignFromPayload(payload()))
    const [merge] = sheet.getMergedRanges()
    expect([merge!.getRow(), merge!.getColumn(), merge!.getHeight(), merge!.getWidth()]).toEqual([
      2, 1, 1, 2,
    ])
    expect(sheet.getNotes()).toEqual([{ row: 2, col: 1, note: 'note' }])
    expect(sheet.getDataValidations()[0]!.rule.uid).toBe('dv1')
  })
})

describe('formulas that cannot follow a block into another workbook', () => {
  it('detects references to other sheets', () => {
    expect(hasSheetReference('=Sheet2!A1+1')).toBe(true)
    expect(hasSheetReference("=SUM('My Sheet'!A1:B2)")).toBe(true)
    expect(hasSheetReference("='It''s'!A1")).toBe(true)
    expect(hasSheetReference('=A1+B2')).toBe(false)
    expect(hasSheetReference('=IF(A1="a!b",1,2)')).toBe(false)
    expect(hasSheetReference('=#REF!+1')).toBe(false)
  })

  it('keeps plain formulas and drops the formula (keeping the value) of cross-sheet ones', () => {
    expect(foreignFormula('=A1*2')).toBe('=A1*2')
    expect(foreignFormula('=Sheet2!A1*2')).toBeNull()
    expect(foreignFormula(undefined)).toBeNull()
    // a defined name only means something where the pasted-into workbook has it too
    expect(foreignFormula('=Rate*B3')).toBeNull()
    expect(foreignFormula('=Rate*B3', new Set(['RATE']))).toBe('=Rate*B3')
    expect(foreignFormula('=SUM(A1:B2)+IF(TRUE,1,2)')).toBe('=SUM(A1:B2)+IF(TRUE,1,2)')
  })
})

describe('foreign copy from text', () => {
  it('types each field and reads quoted fields', () => {
    const foreign = foreignFromText('Name\tQty\tShip\nA "x"\t12\t"two\nlines"\n\t=B2*2\t')
    expect(foreign.origin).toBe('text')
    const cells = foreign.copied.cells
    expect(foreign.copied.rows).toEqual([0, 1, 2])
    expect(foreign.copied.columns).toEqual([0, 1, 2])
    expect(cells[1]![1]).toMatchObject({ v: 12, t: 2 })
    expect(cells[1]![2]).toMatchObject({ v: 'two\nlines', t: 1 })
    expect(cells[2]![0]).toBeNull()
    expect(cells[2]![1]).toMatchObject({ f: '=B2*2' })
    expect(foreign.literalFormulas).toBe(true)
  })
})

import { usedExtent } from '../src/renderer/clipboard-payload'
import { linkEditsForLocalPaste } from '../src/renderer/clipboard-links'

describe('usedExtent (whole-row / whole-column copies carry the used part only)', () => {
  it('stops at the last row and column that hold something', () => {
    expect(
      usedExtent([
        [0, 0],
        [2, 3],
        [1, 1],
      ]),
    ).toEqual({ rows: 3, columns: 4 })
    expect(usedExtent([])).toEqual({ rows: 1, columns: 1 })
  })
})

describe('hyperlinks of a block pasted inside one window', () => {
  const base = {
    sourceRows: [0, 1],
    sourceColumns: [0, 1],
    sourceSheetId: 's1',
    targetSheetId: 's1',
    cut: false,
    targetLink: () => null,
  }
  it('gives every pasted cell the link of the source cell it came from', () => {
    const edits = linkEditsForLocalPaste({
      ...base,
      targetRows: [5, 6],
      targetColumns: [3, 4],
      sourceTarget: (r, c) => (r === 0 && c === 1 ? 'https://a.example' : null),
    })
    expect(edits.get('s1')).toEqual([{ row: 5, column: 4, target: 'https://a.example' }])
  })
  it('repeats the block over a bigger selection and clears links the source does not have', () => {
    const edits = linkEditsForLocalPaste({
      ...base,
      targetRows: [4, 5, 6, 7],
      targetColumns: [0, 1],
      sourceTarget: (r, c) => (r === 0 && c === 0 ? 'https://a.example' : null),
      targetLink: (r, c) => (r === 5 && c === 1 ? 'https://old.example' : null),
    })
    const list = edits.get('s1')!
    expect(list).toContainEqual({ row: 4, column: 0, target: 'https://a.example' })
    expect(list).toContainEqual({ row: 6, column: 0, target: 'https://a.example' })
    expect(list).toContainEqual({ row: 5, column: 1, target: null })
    expect(list).toHaveLength(3)
  })
  it('takes the links off the source after a Cut, on its own sheet', () => {
    const edits = linkEditsForLocalPaste({
      ...base,
      cut: true,
      targetSheetId: 's2',
      targetRows: [0, 1],
      targetColumns: [0, 1],
      sourceTarget: (r, c) => (r === 1 && c === 1 ? 'https://a.example' : null),
    })
    expect(edits.get('s2')).toEqual([{ row: 1, column: 1, target: 'https://a.example' }])
    expect(edits.get('s1')).toEqual([{ row: 1, column: 1, target: null }])
  })
})
