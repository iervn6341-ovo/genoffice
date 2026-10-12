import { RangeThemeStyle } from '@univerjs/sheets'
import { describe, expect, it } from 'vitest'

import { TABLE_STYLE_DEFS } from '../src/renderer/table-style-defs'
import {
  bakeTableStyle,
  DEFAULT_THEME_COLORS,
  resolveColor,
  tableCellLook,
  tableStyleLabel,
  tableThemeFor,
  TABLE_STYLE_GROUPS,
  totalRowTheme,
  type TableStyleOptions,
  type ThemeColors,
} from '../src/renderer/table-styles'
import { COLOR_SCHEMES } from '../src/renderer/themes'

const bgOf = (style: string, part: 'headerRowStyle' | 'firstRowStyle' | 'secondRowStyle') =>
  tableThemeFor(style, 'unique')?.[part]?.bg?.rgb

// Excel 16 (TableStyles via COM) resolved these under the three schemes below
const OFFICE_2013 = COLOR_SCHEMES.find((scheme) => scheme.id === 'office')!.values
const OFFICE_2007 = [
  '#FFFFFF',
  '#000000',
  '#EEECE1',
  '#1F497D',
  '#4F81BD',
  '#C0504D',
  '#9BBB59',
  '#8064A2',
  '#4BACC6',
  '#F79646',
  '#0000FF',
  '#800080',
]
const OFFICE_NEW = [
  '#FFFFFF',
  '#000000',
  '#E8E8E8',
  '#0E2841',
  '#156082',
  '#E97132',
  '#196B24',
  '#0F9ED5',
  '#A02B93',
  '#4EA72E',
  '#467886',
  '#96607D',
]

describe('Format as Table: all 60 built-in styles', () => {
  it("lists Excel's gallery: Light 1–21, Medium 1–28, Dark 1–11, every one defined", () => {
    expect(TABLE_STYLE_GROUPS.map((g) => [g.family, g.names.length])).toEqual([
      ['Light', 21],
      ['Medium', 28],
      ['Dark', 11],
    ])
    for (const { names } of TABLE_STYLE_GROUPS) {
      for (const name of names) expect(TABLE_STYLE_DEFS[name], name).toBeDefined()
    }
    expect(Object.keys(TABLE_STYLE_DEFS)).toHaveLength(60)
    expect(tableStyleLabel('TableStyleMedium2')).toBe('Medium 2')
  })

  it("draws the picked Medium style with the workbook's accent: header, stripe, outline (Office 2013–2022)", () => {
    expect(bgOf('TableStyleMedium2', 'headerRowStyle')).toBe('#4472C4')
    // Excel: Accent 1 lighter 80% is D9E1F2 (the old approximation gave DAE3F3)
    expect(bgOf('TableStyleMedium2', 'firstRowStyle')).toBe('#D9E1F2')
    expect(tableThemeFor('TableStyleMedium2', 'x')?.headerRowStyle?.cl?.rgb).toBe('#FFFFFF')
    // Medium 3 is orange, Medium 7 green (the gallery's numbering, read from Excel)
    expect(bgOf('TableStyleMedium3', 'headerRowStyle')).toBe('#ED7D31')
    expect(bgOf('TableStyleMedium3', 'firstRowStyle')).toBe('#FCE4D6')
    expect(bgOf('TableStyleMedium7', 'headerRowStyle')).toBe('#70AD47')
    expect(bgOf('TableStyleMedium7', 'firstRowStyle')).toBe('#E2EFDA')
    expect(tableThemeFor('TableStyleMedium2', 'x')?.lastRowStyle?.bd?.b?.cl.rgb).toBe('#8EA9DB')
  })

  it("follows the workbook's theme colors, as Excel does (Office 2007–2010 and the new Office theme)", () => {
    const under = (colors: ThemeColors) => ({
      header: tableThemeFor('TableStyleMedium2', 'x', colors)?.headerRowStyle?.bg?.rgb,
      stripe: tableThemeFor('TableStyleMedium2', 'x', colors)?.firstRowStyle?.bg?.rgb,
      outline: tableThemeFor('TableStyleMedium2', 'x', colors)?.lastRowStyle?.bd?.b?.cl.rgb,
    })
    expect(under(OFFICE_2007)).toEqual({ header: '#4F81BD', stripe: '#DCE6F1', outline: '#95B3D7' })
    expect(under(OFFICE_NEW)).toEqual({ header: '#156082', stripe: '#C0E6F5', outline: '#44B3E1' })
    // other accents: Medium 3 is Accent 2 in every scheme
    expect(tableThemeFor('TableStyleMedium3', 'x', OFFICE_2007)?.headerRowStyle?.bg?.rgb).toBe(
      '#C0504D',
    )
    expect(tableThemeFor('TableStyleMedium3', 'x', OFFICE_NEW)?.firstRowStyle?.bg?.rgb).toBe(
      '#FBE2D5',
    )
  })

  it('resolves a tint it has no exact value for with the HSL formula', () => {
    const custom = [...OFFICE_2013]
    custom[4] = '#336699'
    const stripe = tableThemeFor('TableStyleMedium2', 'x', custom)?.firstRowStyle?.bg?.rgb
    expect(stripe).toMatch(/^#[0-9A-F]{6}$/)
    expect(stripe).not.toBe('#D9E1F2')
    expect(resolveColor({ s: 'a1', t: 0.8 }, custom)).toBe(stripe)
    expect(resolveColor({ s: 'a1' }, custom)).toBe('#336699')
  })

  it('keeps Light 1 plain, Light 9 solid, and the Dark family two-toned under a black header', () => {
    const light1 = tableThemeFor('TableStyleLight1', 'x')!
    expect(light1.headerRowStyle?.bg).toBeUndefined()
    expect(light1.headerRowStyle?.bd?.b).toBeDefined()
    expect(light1.firstRowStyle?.bg?.rgb).toBe('#D9D9D9')
    expect(bgOf('TableStyleLight9', 'headerRowStyle')).toBe('#4472C4')
    const dark2 = tableThemeFor('TableStyleDark2', 'x')!
    // Excel: Dark 2 = black header over Accent 1 rows alternating with Accent 1 darker 25%
    expect(dark2.headerRowStyle?.bg?.rgb).toBe('#000000')
    expect(dark2.wholeStyle?.bg?.rgb).toBe('#4472C4')
    expect(dark2.firstRowStyle?.bg?.rgb).toBe('#305496')
    expect(dark2.wholeStyle?.cl?.rgb).toBe('#FFFFFF')
    expect(tableThemeFor('TableStyleDark1', 'x')?.firstRowStyle?.bg?.rgb).toBe('#404040')
  })

  it('draws every style, with an Excel-exact color for the known schemes', () => {
    for (const { names } of TABLE_STYLE_GROUPS) {
      for (const name of names) {
        for (const colors of [DEFAULT_THEME_COLORS, OFFICE_2007, OFFICE_NEW]) {
          const theme = tableThemeFor(name, `t-${name}`, colors)
          expect(theme, name).not.toBeNull()
          expect(theme?.name).toBe(`t-${name}`)
        }
      }
    }
  })

  it('turns banded rows off, banded columns and first / last column on', () => {
    const plain = tableThemeFor('TableStyleMedium2', 'x', DEFAULT_THEME_COLORS, {
      bandedRows: false,
    })!
    expect(plain.firstRowStyle?.bg).toBeUndefined()
    const options: TableStyleOptions = { bandedColumns: true, firstColumn: true, lastColumn: true }
    const full = tableThemeFor('TableStyleMedium2', 'x', DEFAULT_THEME_COLORS, options)!
    expect(full.headerColumnStyle?.bl).toBe(1)
    expect(full.lastColumnStyle?.bl).toBe(1)
    expect(full.secondColumnStyle?.bg?.rgb).toBe('#D9E1F2')
    expect(tableThemeFor('TableStyleMedium2', 'x')?.headerColumnStyle?.bl).toBeUndefined()
  })

  it('names every theme as asked, because Univer keys themes by name', () => {
    expect(tableThemeFor('TableStyleMedium2', 'genoffice-ai-table-1')?.name).toBe(
      'genoffice-ai-table-1',
    )
  })

  it('leaves styles it draws no look for to Univer', () => {
    expect(tableThemeFor(undefined, 'x')).toBeNull()
    expect(tableThemeFor('PivotStyleLight16', 'x')).toBeNull()
    expect(tableThemeFor('TableStyleMedium29', 'x')).toBeNull()
    expect(tableThemeFor('TableStyleLight99', 'x')).toBeNull()
  })
})

describe('the Univer theme and the gallery thumbnails draw the same table', () => {
  const rows = 6
  const columns = 4
  const variants: { name: string; options: TableStyleOptions }[] = [
    { name: 'default', options: {} },
    { name: 'no stripes', options: { bandedRows: false } },
    { name: 'columns', options: { bandedColumns: true, firstColumn: true, lastColumn: true } },
  ]

  it('agrees on fill, text color and bold for every cell of every style', () => {
    for (const { family, names } of TABLE_STYLE_GROUPS) {
      for (const name of names) {
        for (const { options } of variants) {
          const theme = tableThemeFor(name, 'x', DEFAULT_THEME_COLORS, options)!
          const range = new RangeThemeStyle('x', theme)
          for (let row = 0; row < rows; row += 1) {
            for (let col = 0; col < columns; col += 1) {
              const univer = range.getStyle(row, col, row === rows - 1, col === columns - 1, false)
              const look = tableCellLook(
                name,
                { rows, columns, row, col },
                DEFAULT_THEME_COLORS,
                options,
              )
              const where = `${family} ${name} r${row}c${col} ${JSON.stringify(options)}`
              expect(univer?.bg?.rgb, `${where} fill`).toBe(look.bg)
              expect(univer?.cl?.rgb, `${where} color`).toBe(look.color)
              expect(univer?.bl === 1 ? true : undefined, `${where} bold`).toBe(look.bold)
            }
          }
        }
      }
    }
  })

  it('agrees on the outline and the lines inside', () => {
    // a NONE edge takes an inherited line off again
    const lineOf = (value: unknown): string | undefined => {
      const edge = value as { s?: number; cl?: { rgb?: string } } | null | undefined
      return edge && edge.s !== 0 ? edge.cl?.rgb : undefined
    }
    for (const { names } of TABLE_STYLE_GROUPS) {
      for (const name of names) {
        const theme = tableThemeFor(name, 'x')!
        const range = new RangeThemeStyle('x', theme)
        for (let row = 0; row < rows; row += 1) {
          for (let col = 0; col < columns; col += 1) {
            const univer = range.getStyle(row, col, row === rows - 1, col === columns - 1, false)
            const look = tableCellLook(name, { rows, columns, row, col })
            const where = `${name} r${row}c${col}`
            expect(lineOf(univer?.bd?.t), `${where} top`).toBe(look.top?.color)
            expect(lineOf(univer?.bd?.b), `${where} bottom`).toBe(look.bottom?.color)
            expect(lineOf(univer?.bd?.l), `${where} left`).toBe(look.left?.color)
            expect(lineOf(univer?.bd?.r), `${where} right`).toBe(look.right?.color)
          }
        }
      }
    }
  })
})

describe('Convert to Range bakes the style into the cells (Excel keeps the look)', () => {
  const area = { startRow: 2, startColumn: 1, endRow: 5, endColumn: 2 }
  const OFFICE_NEW_COLORS = [
    '#FFFFFF',
    '#000000',
    '#E8E8E8',
    '#0E2841',
    '#156082',
    '#E97132',
    '#196B24',
    '#0F9ED5',
    '#A02B93',
    '#4EA72E',
    '#467886',
    '#96607D',
  ]

  it('Medium 2: header fill, bold white text, stripe fill and the lines (read from Excel after Unlist)', async () => {
    const { bakeTableStyle } = await import('../src/renderer/table-styles')
    const baked = bakeTableStyle('TableStyleMedium2', area, OFFICE_NEW_COLORS)!
    // Excel: A1 fill 156082, bold, white; A2 fill C0E6F5 with a bottom line; A3 no fill
    expect(baked[2]![1]!.s.bg?.rgb).toBe('#156082')
    expect(baked[2]![1]!.s.bl).toBe(1)
    expect(baked[2]![1]!.s.cl?.rgb).toBe('#FFFFFF')
    expect(baked[3]![1]!.s.bg?.rgb).toBe('#C0E6F5')
    expect(baked[3]![1]!.s.bd?.b?.cl.rgb).toBe('#44B3E1')
    expect(baked[4]![1]!.s.bg).toBeUndefined()
    // the table's outline closes on the last row and the last column
    expect(baked[5]![2]!.s.bd?.b?.cl.rgb).toBe('#44B3E1')
    expect(baked[3]![2]!.s.bd?.r?.cl.rgb).toBe('#44B3E1')
  })

  it('a name that is not one of the 60 styles bakes nothing', async () => {
    const { bakeTableStyle } = await import('../src/renderer/table-styles')
    expect(bakeTableStyle('TableStyleMedium29', area)).toBeNull()
  })
})

describe("Convert to Range keeps the cells' own formatting above the style", () => {
  it('drops each baked attribute the cell already carries', async () => {
    const { withoutOwnFormatting } = await import('../src/renderer/table-styles')
    const item = {
      bg: { rgb: '#156082' },
      cl: { rgb: '#FFFFFF' },
      bl: 1,
      bd: { b: { s: 1, cl: { rgb: '#000000' } }, r: { s: 1, cl: { rgb: '#000000' } } },
    } as never
    // a yellow cell keeps its yellow; the style still bolds it and draws its lines
    const yellow = withoutOwnFormatting(item, { bg: { rgb: '#FFFF00' } })!
    expect(yellow.bg).toBeUndefined()
    expect(yellow.cl?.rgb).toBe('#FFFFFF')
    expect(yellow.bl).toBe(1)
    expect(Object.keys(yellow.bd ?? {})).toEqual(['b', 'r'])
    // its own bottom line stays, the right one is the style's
    const lined = withoutOwnFormatting(item, { bd: { b: { s: 1, cl: { rgb: '#FF0000' } } } })!
    expect(Object.keys(lined.bd ?? {})).toEqual(['r'])
    // nothing of its own: the whole style; everything of its own: nothing is baked
    expect(withoutOwnFormatting(item, null)).toBe(item)
    expect(withoutOwnFormatting({ bg: { rgb: '#156082' } }, { bg: { rgb: '#FFFF00' } })).toBeNull()
  })
})

describe('a table without a header row, and one with a totals row, are drawn like Excel draws them', () => {
  const lineOf = (value: unknown): string | undefined => {
    const edge = value as { s?: number; cl?: { rgb?: string } } | null | undefined
    return edge && edge.s !== 0 ? edge.cl?.rgb : undefined
  }
  const names = TABLE_STYLE_GROUPS.flatMap((group) => group.names)

  it('lays the theme over the data rows: the first data row takes the first stripe, the rest swap', () => {
    const rows = 5
    const columns = 3
    for (const name of names) {
      for (const options of [{ headerRow: false }, { headerRow: false, bandedRows: false }]) {
        const theme = tableThemeFor(name, 'x', DEFAULT_THEME_COLORS, options)!
        const range = new RangeThemeStyle('x', theme)
        for (let row = 0; row < rows; row += 1) {
          for (let col = 0; col < columns; col += 1) {
            const univer = range.getStyle(row, col, row === rows - 1, col === columns - 1, false)
            const look = tableCellLook(
              name,
              { rows, columns, row, col },
              DEFAULT_THEME_COLORS,
              options,
            )
            const where = `${name} r${row}c${col} ${JSON.stringify(options)}`
            expect(univer?.bg?.rgb, `${where} fill`).toBe(look.bg)
            expect(univer?.cl?.rgb, `${where} color`).toBe(look.color)
            expect(lineOf(univer?.bd?.t), `${where} top`).toBe(look.top?.color)
            expect(lineOf(univer?.bd?.b), `${where} bottom`).toBe(look.bottom?.color)
          }
        }
      }
    }
  })

  it('stripes a headerless table from its first row, as Excel does (Medium 2: stripe, plain, stripe)', () => {
    const look = (row: number) =>
      tableCellLook(
        'TableStyleMedium2',
        { rows: 4, columns: 2, row, col: 0 },
        DEFAULT_THEME_COLORS,
        {
          headerRow: false,
        },
      )
    expect(look(0).bg).toBe('#D9E1F2')
    expect(look(1).bg).toBeUndefined()
    expect(look(2).bg).toBe('#D9E1F2')
    expect(look(0).bold).toBeUndefined()
  })

  it('leaves the totals row out of the table theme: the last data row ends on its top rule', () => {
    const rows = 6
    const columns = 3
    for (const name of names) {
      const options = { totalRow: true }
      const theme = tableThemeFor(name, 'x', DEFAULT_THEME_COLORS, options)!
      const range = new RangeThemeStyle('x', theme)
      for (let row = 0; row < rows; row += 1) {
        for (let col = 0; col < columns; col += 1) {
          const univer = range.getStyle(row, col, row === rows - 1, col === columns - 1, false)
          // the grid the look sees has the totals row under the six rows the theme covers
          const look = tableCellLook(
            name,
            { rows: rows + 1, columns, row, col },
            DEFAULT_THEME_COLORS,
            options,
          )
          const where = `${name} r${row}c${col}`
          expect(univer?.bg?.rgb, `${where} fill`).toBe(look.bg)
          expect(univer?.cl?.rgb, `${where} color`).toBe(look.color)
        }
      }
    }
  })

  it('draws the totals row apart with the style’s total element and the outline closing the table', () => {
    const columns = 4
    for (const name of names) {
      const theme = totalRowTheme(name, 'x')!
      const range = new RangeThemeStyle('x', theme)
      for (let col = 0; col < columns; col += 1) {
        const univer = range.getStyle(0, col, true, col === columns - 1, false)
        const look = tableCellLook(name, { rows: 7, columns, row: 6, col }, DEFAULT_THEME_COLORS, {
          totalRow: true,
        })
        const where = `${name} c${col}`
        expect(univer?.bg?.rgb, `${where} fill`).toBe(look.bg)
        expect(univer?.cl?.rgb, `${where} color`).toBe(look.color)
        expect(univer?.bl === 1 ? true : undefined, `${where} bold`).toBe(look.bold)
        expect(lineOf(univer?.bd?.t), `${where} top`).toBe(look.top?.color)
        expect(lineOf(univer?.bd?.b), `${where} bottom`).toBe(look.bottom?.color)
        expect(lineOf(univer?.bd?.l), `${where} left`).toBe(look.left?.color)
        expect(lineOf(univer?.bd?.r), `${where} right`).toBe(look.right?.color)
      }
    }
  })

  it('Medium 2 totals: white-less bold row over a line, as Excel resolves the total element', () => {
    const look = tableCellLook(
      'TableStyleMedium2',
      { rows: 5, columns: 2, row: 4, col: 0 },
      DEFAULT_THEME_COLORS,
      {
        totalRow: true,
      },
    )
    expect(look.bold).toBe(true)
    expect(look.top?.line).toBe('double')
  })

  it('bakes a totals row and a headerless table into the cells on Convert to Range', () => {
    const area = { startRow: 1, startColumn: 0, endRow: 4, endColumn: 1 }
    const withTotals = bakeTableStyle('TableStyleMedium2', area, DEFAULT_THEME_COLORS, {
      totalRow: true,
    })!
    // the totals row (the last one) is bold with the style's rule on top; the header row keeps its fill
    expect(withTotals[4]?.[0]?.s.bl).toBe(1)
    expect(withTotals[1]?.[0]?.s.bg?.rgb).toBe('#4472C4')
    const headerless = bakeTableStyle('TableStyleMedium2', area, DEFAULT_THEME_COLORS, {
      headerRow: false,
    })!
    // no header band: the first row is a stripe row
    expect(headerless[1]?.[0]?.s.bg?.rgb).toBe('#D9E1F2')
    expect(headerless[1]?.[0]?.s.bl).toBeUndefined()
  })
})
