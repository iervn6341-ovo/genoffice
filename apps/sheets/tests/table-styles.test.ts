import { describe, expect, it } from 'vitest'

import { tableThemeFor } from '../src/renderer/table-styles'

const bgOf = (style: string, part: 'headerRowStyle' | 'firstRowStyle' | 'secondRowStyle') =>
  tableThemeFor(style, 'unique')?.[part]?.bg?.rgb

describe('Format as Table: built-in style → Univer table theme', () => {
  it('draws the accent of the picked Medium style on the header and a light stripe', () => {
    expect(bgOf('TableStyleMedium2', 'headerRowStyle')).toBe('#4472C4')
    expect(bgOf('TableStyleMedium3', 'headerRowStyle')).toBe('#ED7D31')
    expect(bgOf('TableStyleMedium7', 'headerRowStyle')).toBe('#70AD47')
    // "Lighter 80%" of the accent
    expect(bgOf('TableStyleMedium2', 'firstRowStyle')).toBe('#DAE3F3')
    expect(tableThemeFor('TableStyleMedium2', 'x')?.headerRowStyle?.cl?.rgb).toBe('#FFFFFF')
  })

  it('keeps the Light family unfilled in the header and the Dark family two-toned', () => {
    const light1 = tableThemeFor('TableStyleLight1', 'x')!
    expect(light1.headerRowStyle?.bg).toBeUndefined()
    expect(light1.headerRowStyle?.bd?.b).toBeDefined()
    expect(light1.firstRowStyle?.bg?.rgb).toBe('#D9D9D9')
    const light9 = tableThemeFor('TableStyleLight9', 'x')!
    expect(light9.headerRowStyle?.bg?.rgb).toBe('#4472C4')
    const dark2 = tableThemeFor('TableStyleDark2', 'x')!
    expect(dark2.firstRowStyle?.bg?.rgb).not.toBe(dark2.secondRowStyle?.bg?.rgb)
    // a header deeper than the accent it is built on
    expect(dark2.headerRowStyle?.bg?.rgb).toBe('#223962')
  })

  it('names every theme as asked, because Univer keys themes by name', () => {
    expect(tableThemeFor('TableStyleMedium2', 'genoffice-ai-table-1')?.name).toBe(
      'genoffice-ai-table-1',
    )
  })

  it('leaves styles it draws no look for to Univer', () => {
    expect(tableThemeFor(undefined, 'x')).toBeNull()
    expect(tableThemeFor('PivotStyleLight16', 'x')).toBeNull()
    expect(tableThemeFor('TableStyleMedium15', 'x')).toBeNull()
    expect(tableThemeFor('TableStyleLight99', 'x')).toBeNull()
  })
})
