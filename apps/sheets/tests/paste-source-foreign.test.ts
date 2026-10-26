// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'

import { PAYLOAD_ATTR, encodePayload, createStyleTable } from '../src/renderer/clipboard-payload'
import { buildClipboardHtml } from '../src/renderer/clipboard-html'
import { foreignFromClipboardData } from '../src/renderer/paste-source'

const payload = () => {
  const styles = createStyleTable()
  const bold = styles.indexOf({ bl: 1 })!
  return {
    version: 1,
    bookId: 'b',
    sheetId: 's',
    sheetName: 'Sheet1',
    rows: [0],
    columns: [0, 1],
    styles: styles.styles,
    cells: [
      [
        { v: 'a', t: 1, s: bold },
        { v: 2, t: 2 },
      ],
    ],
    merges: [],
    columnWidths: [null, null],
    rowHeights: [null],
    notes: [],
    validations: [],
    conditionalFormats: [],
    links: [],
  }
}

const withPayload = (copyId: string): string =>
  buildClipboardHtml({
    copyId,
    payloadValue: encodePayload(payload()),
    cells: [
      [
        { text: 'a', v: 'a', t: 1, style: { bl: 1 } },
        { text: '2', v: 2, t: 2, style: null },
      ],
    ],
    columnWidths: [null, null],
    rowHeights: [null],
  })

describe('which copies are foreign to a window', () => {
  const never = () => false

  it('reads the GenOffice payload of another window', () => {
    const foreign = foreignFromClipboardData(withPayload('xyz'), 'a\t2', never)!
    expect(foreign.origin).toBe('genoffice')
    expect(foreign.copied.cells[0]![0]).toMatchObject({ v: 'a', s: { bl: 1 } })
  })

  it('leaves a copy of this window to Univer (its cache has the id)', () => {
    expect(foreignFromClipboardData(withPayload('mine'), 'a\t2', (id) => id === 'mine')).toBeNull()
  })

  it('reads Excel-style markup without a payload', () => {
    const html = withPayload('x').replace(new RegExp(`${PAYLOAD_ATTR}="[^"]*"`), '')
    const foreign = foreignFromClipboardData(html, 'a\t2', never)!
    expect(foreign.origin).toBe('excel')
    expect(foreign.copied.cells[0]![1]).toMatchObject({ v: 2, t: 2 })
  })

  it('leaves other table markup (web pages) to Univer', () => {
    const html = '<table><tr><td>a</td><td>b</td></tr></table>'
    expect(foreignFromClipboardData(html, 'a\tb', never)).toBeNull()
  })

  it('splits plain text at tabs and line breaks, even when other markup came with it', () => {
    const foreign = foreignFromClipboardData('<div>x</div>', 'a\tb\n1\t2', never)!
    expect(foreign.origin).toBe('text')
    expect(foreign.copied.rows).toHaveLength(2)
    expect(foreignFromClipboardData(null, 'plain', never)).toBeNull()
    expect(foreignFromClipboardData('', '', never)).toBeNull()
  })
})
