import { describe, expect, it } from 'vitest'

import { buildTableXml, type TableAddition } from '@genoffice/xlsx-gateway/gateway/xlsx-table-add'

const base: TableAddition = {
  worksheetPath: 'xl/worksheets/sheet1.xml',
  area: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
  name: 'Table1',
  columnNames: ['a', 'b', 'c'],
  bandedRows: true,
}

const styleInfo = (xml: string): string => /<tableStyleInfo [^>]*\/>/.exec(xml)![0]

describe('a session table writes Table Design → Table Style Options into tableStyleInfo', () => {
  it('keeps the old output for a table without options: banded rows only', () => {
    expect(styleInfo(buildTableXml(1, base))).toBe(
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
    )
  })

  it('writes first column, last column and banded columns, and banded rows off', () => {
    const xml = buildTableXml(2, {
      ...base,
      style: 'TableStyleMedium14',
      bandedRows: false,
      bandedColumns: true,
      firstColumn: true,
      lastColumn: true,
    })
    expect(styleInfo(xml)).toBe(
      '<tableStyleInfo name="TableStyleMedium14" showFirstColumn="1" showLastColumn="1" showRowStripes="0" showColumnStripes="1"/>',
    )
  })

  it('treats an explicit false like a missing option', () => {
    expect(
      styleInfo(
        buildTableXml(3, { ...base, bandedColumns: false, firstColumn: false, lastColumn: false }),
      ),
    ).toContain('showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"')
  })
})

describe('Header Row, Total Row and Filter Button in the table part (Excel 16 writes the same)', () => {
  it('writes the old part for a table without them: header row, filter on, no totals', () => {
    const xml = buildTableXml(1, base)
    expect(xml).toContain('ref="A1:C5" totalsRowShown="0">')
    expect(xml).toContain('<autoFilter ref="A1:C5"/>')
    expect(xml).not.toContain('headerRowCount')
    expect(xml).not.toContain('totalsRowCount')
  })

  it('a totals row is counted, labelled and given its function; the filter stops above it', () => {
    const xml = buildTableXml(1, {
      ...base,
      area: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 2 },
      totalsRow: true,
      totals: [{ label: 'Total' }, null, { fn: 'sum' }],
    })
    // Excel (COM, saved): totalsRowCount="1", autoFilter over header + data only, totalsRowLabel on the
    // first column and totalsRowFunction on the last
    expect(xml).toContain('ref="A1:C6" totalsRowCount="1">')
    expect(xml).not.toContain('totalsRowShown')
    expect(xml).toContain('<autoFilter ref="A1:C5"/>')
    expect(xml).toContain('<tableColumn id="1" name="a" totalsRowLabel="Total"/>')
    expect(xml).toContain('<tableColumn id="2" name="b"/>')
    expect(xml).toContain('<tableColumn id="3" name="c" totalsRowFunction="sum"/>')
  })

  it('a table without a header row has headerRowCount="0" and no filter', () => {
    const xml = buildTableXml(1, {
      ...base,
      area: { startRow: 1, startColumn: 0, endRow: 4, endColumn: 2 },
      headerRow: false,
    })
    expect(xml).toContain('ref="A2:C5" headerRowCount="0" totalsRowShown="0">')
    expect(xml).not.toContain('<autoFilter')
    // the columns keep their names
    expect(xml).toContain('<tableColumn id="1" name="a"/>')
  })

  it('Filter Button off drops the autoFilter only', () => {
    const xml = buildTableXml(1, { ...base, filterButton: false })
    expect(xml).not.toContain('<autoFilter')
    expect(xml).not.toContain('headerRowCount')
  })

  it('escapes a label and leaves a "none" function out', () => {
    const xml = buildTableXml(1, {
      ...base,
      totalsRow: true,
      area: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 2 },
      totals: [{ label: 'A & "B"', fn: 'none' }],
    })
    expect(xml).toContain('totalsRowLabel="A &amp; &quot;B&quot;"')
    expect(xml).not.toContain('totalsRowFunction')
  })
})
