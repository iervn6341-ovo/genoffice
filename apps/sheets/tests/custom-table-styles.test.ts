import { RangeThemeStyle } from '@univerjs/sheets'
import { afterEach, describe, expect, it } from 'vitest'

import {
  addCustomTableStyle,
  customStyleDef,
  customStylesUsedBy,
  listCustomTableStyles,
  nextStyleName,
  openCustomTableStylesFor,
  saveCustomTableStylesTo,
  styleNameProblem,
  toWireStyle,
  type CustomTableStyle,
} from '../src/renderer/custom-table-styles'
import { tableCellLook, tableThemeFor, tableStyleDef } from '../src/renderer/table-styles'
import {
  buildTableStyleDxf,
  type TableStyleElementAdd,
} from '@genoffice/xlsx-gateway/gateway/xlsx-table-style-add'

const STYLE: CustomTableStyle = {
  name: 'Ledger',
  elements: {
    headerRow: { fill: '#1F3864', font: '#FFFFFF', bold: true },
    firstRowStripe: { fill: '#E2EFDA' },
    wholeTable: {
      line: 'thin',
      lineColor: '#808080',
      edges: ['top', 'bottom', 'left', 'right', 'insideH'],
    },
    totalRow: { bold: true, line: 'double', lineColor: '#000000', edges: ['top'] },
  },
}

afterEach(() => openCustomTableStylesFor('reset'))

describe('a custom table style is drawn like the built-in ones', () => {
  it('turns into a style definition with its own RGB colors, whatever the workbook theme', () => {
    const def = customStyleDef(STYLE)
    expect(def.header?.fill).toEqual({ rgb: '#1F3864' })
    expect(def.header?.bold).toBe(true)
    expect(def.stripe1?.fill).toEqual({ rgb: '#E2EFDA' })
    expect(def.whole?.insideH).toEqual({ l: 'thin', c: { rgb: '#808080' } })
    expect(def.total?.top).toEqual({ l: 'double', c: { rgb: '#000000' } })
    expect(def.stripe2).toBeUndefined()
  })

  it('is known by name once added, to the theme the grid gets and the cell looks the gallery draws', () => {
    addCustomTableStyle(STYLE)
    expect(tableStyleDef('Ledger')).not.toBeNull()
    const theme = tableThemeFor('Ledger', 'x')!
    expect(theme.headerRowStyle?.bg?.rgb).toBe('#1F3864')
    expect(theme.firstRowStyle?.bg?.rgb).toBe('#E2EFDA')
    const range = new RangeThemeStyle('x', theme)
    // header, then the stripe, then the plain row — the same as the cell looks
    for (const row of [0, 1, 2, 3]) {
      const univer = range.getStyle(row, 1, row === 3, false, false)
      const look = tableCellLook('Ledger', { rows: 4, columns: 3, row, col: 1 })
      expect(univer?.bg?.rgb, `fill r${row}`).toBe(look.bg)
      expect(univer?.cl?.rgb, `color r${row}`).toBe(look.color)
    }
    expect(tableCellLook('Ledger', { rows: 4, columns: 3, row: 0, col: 0 }).bold).toBe(true)
  })

  it('is forgotten when another file is opened, and kept when a save reopens the same one', () => {
    openCustomTableStylesFor('a.xlsx')
    addCustomTableStyle(STYLE)
    saveCustomTableStylesTo('b.xlsx') // Save As: the file is b.xlsx now
    openCustomTableStylesFor('b.xlsx')
    expect(listCustomTableStyles().map((style) => style.name)).toEqual(['Ledger'])
    openCustomTableStylesFor('c.xlsx')
    expect(listCustomTableStyles()).toEqual([])
    expect(tableStyleDef('Ledger')).toBeNull()
  })

  it('lists the styles a save must write: those the session’s tables use', () => {
    addCustomTableStyle(STYLE)
    addCustomTableStyle({ name: 'Unused', elements: { headerRow: { fill: '#FF0000' } } })
    expect(
      customStylesUsedBy(['TableStyleMedium2', 'Ledger', undefined]).map((style) => style.name),
    ).toEqual(['Ledger'])
    expect(customStylesUsedBy([])).toEqual([])
  })
})

describe('the name of a new table style', () => {
  it('is refused when it is empty, a built-in name, or in use (any case)', () => {
    expect(styleNameProblem('  ', [])).toBe('empty')
    expect(styleNameProblem('TableStyleMedium2', [])).toBe('builtin')
    expect(styleNameProblem('tablestylelight7', [])).toBe('builtin')
    expect(styleNameProblem('Ledger', ['ledger'])).toBe('taken')
    expect(styleNameProblem('a<b', [])).toBe('invalid')
    expect(styleNameProblem('Ledger', [])).toBeNull()
    expect(styleNameProblem('My TableStyle', [])).toBeNull()
  })

  it('defaults to the first free "Table Style n"', () => {
    expect(nextStyleName('Table Style', [])).toBe('Table Style 1')
    expect(nextStyleName('Table Style', ['Table Style 1', 'table style 2'])).toBe('Table Style 3')
  })
})

describe('the style as the save carries it', () => {
  it('lists each formatted element with its fill, font, bold and the edges it draws', () => {
    const wire = toWireStyle(STYLE)
    expect(wire.name).toBe('Ledger')
    expect(wire.elements).toEqual([
      {
        type: 'wholeTable',
        line: 'thin',
        lineColor: '#808080',
        edges: ['top', 'bottom', 'left', 'right', 'insideH'],
      },
      { type: 'headerRow', fill: '#1F3864', font: '#FFFFFF', bold: true },
      {
        type: 'totalRow',
        bold: true,
        line: 'double',
        lineColor: '#000000',
        edges: ['top'],
      },
      { type: 'firstRowStripe', fill: '#E2EFDA' },
    ])
  })

  it('leaves out an element with nothing set, an edge the element has no side for, and a line with no edge', () => {
    const wire = toWireStyle({
      name: 'X',
      elements: {
        headerRow: {},
        lastColumn: { line: 'thin', lineColor: '#000000', edges: ['top', 'right'] },
        firstRowStripe: { line: 'thin', edges: [] },
        secondRowStripe: { fill: 'not a color' },
      },
    })
    expect(wire.elements).toEqual([
      { type: 'lastColumn', line: 'thin', lineColor: '#000000', edges: ['right'] },
    ])
  })
})

describe('the dxf the save writes for an element', () => {
  const dxf = (element: Partial<TableStyleElementAdd>) =>
    buildTableStyleDxf({ type: 'headerRow', ...element })

  it('is font, fill, border in the order the schema asks for', () => {
    expect(
      dxf({
        fill: '#1f3864',
        font: '#FFFFFF',
        bold: true,
        line: 'thin',
        lineColor: '#808080',
        edges: ['top', 'bottom'],
      }),
    ).toBe(
      '<dxf><font><b/><color rgb="FFFFFFFF"/></font>' +
        '<fill><patternFill><bgColor rgb="FF1F3864"/></patternFill></fill>' +
        '<border><top style="thin"><color rgb="FF808080"/></top>' +
        '<bottom style="thin"><color rgb="FF808080"/></bottom></border></dxf>',
    )
  })

  it('writes the borders in CT_Border order and names the inside lines vertical / horizontal', () => {
    const xml = dxf({
      line: 'double',
      lineColor: '#000000',
      edges: ['insideH', 'insideV', 'right', 'left', 'bottom', 'top'],
    })
    const order = [...xml.matchAll(/<(left|right|top|bottom|vertical|horizontal) /g)].map(
      (match) => match[1],
    )
    expect(order).toEqual(['left', 'right', 'top', 'bottom', 'vertical', 'horizontal'])
  })

  it('is empty for an element with nothing to say', () => {
    expect(dxf({})).toBe('<dxf></dxf>')
  })
})
