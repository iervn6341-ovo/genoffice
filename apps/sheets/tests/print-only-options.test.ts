import { describe, expect, it } from 'vitest'

import { buildSheetPrintPayload, renderHeaderFooterHtml } from '../src/renderer/print-html'
import type { PrintWorksheet } from '../src/renderer/print-html'
import {
  firstPageOffset,
  isDefaultPrintOptions,
  printedErrorText,
  verticalCenterPadding,
} from '../src/renderer/print-options'
import { printScopeOf } from '../src/renderer/print-entry'
import type { EffectivePageSetup } from '../src/renderer/print-settings'

const GRID = [
  ['1', '#DIV/0!'],
  ['#N/A', 'ok'],
]

function sheet(): PrintWorksheet {
  return {
    getLastRow: () => 1,
    getLastColumn: () => 1,
    getRowHeight: () => 20,
    getColumnWidth: () => 100,
    getMergedRanges: () => [],
    getRange: ((row: number, column: number, numRows?: number, numColumns?: number) => ({
      getDisplayValues: () =>
        GRID.slice(row, row + (numRows ?? 1)).map((r) =>
          r.slice(column, column + (numColumns ?? 1)),
        ),
      getValues: () =>
        GRID.slice(row, row + (numRows ?? 1)).map((r) =>
          r.slice(column, column + (numColumns ?? 1)),
        ),
      getCellStyleData: () => ({ cl: { rgb: '#ff0000' }, bg: { rgb: '#00ff00' } }),
    })) as PrintWorksheet['getRange'],
  }
}

function setup(overrides: Partial<EffectivePageSetup> = {}): EffectivePageSetup {
  return {
    orientation: 'portrait',
    paperSize: 9,
    scale: 100,
    fitToWidth: 0,
    fitToHeight: 0,
    fitToPage: false,
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    printGridlines: true,
    printHeadings: false,
    printAreas: [],
    printTitles: null,
    header: null,
    footer: { center: 'p &P' },
    firstPage: null,
    evenPages: null,
    headerFooterScaleWithDoc: true,
    headerFooterPictures: [],
    ...overrides,
  }
}

const build = (options: Parameters<typeof buildSheetPrintPayload>[9]) =>
  buildSheetPrintPayload(
    sheet(),
    setup(),
    'B.pdf',
    'S1',
    new Map(),
    { visuals: [], css: '' },
    'B.xlsx',
    undefined,
    undefined,
    options,
  )

describe('print-only options', () => {
  it('maps error values per "cell errors as"', () => {
    expect(printedErrorText('#DIV/0!', 'blank')).toBe('')
    expect(printedErrorText('#DIV/0!', 'dash')).toBe('--')
    expect(printedErrorText('#DIV/0!', 'na')).toBe('#N/A')
    expect(printedErrorText('#DIV/0!', 'displayed')).toBe('#DIV/0!')
    expect(printedErrorText('#DIV/0!', undefined)).toBe('#DIV/0!')
    expect(printedErrorText('plain', 'blank')).toBe('plain')
  })

  it('changes the laid-out cells', () => {
    expect(build({}).html).toContain('>#DIV/0!</td>')
    const blank = build({ errorsAs: 'blank' }).html
    expect(blank).not.toContain('#DIV/0!')
    expect(blank).not.toContain('>#N/A<')
    expect(build({ errorsAs: 'dash' }).html).toContain('>--</td>')
  })

  it('centers horizontally with a shrink-wrapped area', () => {
    expect(build({}).html).not.toContain('margin-left: auto')
    expect(build({ centerHorizontally: true }).html).toContain('margin-left: auto')
  })

  it('centers a one-page area vertically and leaves tall ones alone', () => {
    expect(verticalCenterPadding(100, 700)).toBe(298)
    expect(verticalCenterPadding(700, 700)).toBe(0)
    expect(verticalCenterPadding(900, 700)).toBe(0)
    expect(build({}).html).not.toContain('padding-top')
    expect(build({ centerVertically: true }).html).toMatch(/class="area" style="padding-top:\d+/)
  })

  it('prints black and white', () => {
    const html = build({ blackAndWhite: true }).html
    expect(html).toContain('color: #000 !important')
    expect(html).toContain('background: transparent !important')
    expect(build({}).html).not.toContain('!important')
  })

  it('draft quality drops the gridlines', () => {
    expect(build({}).html).toContain('0.5pt solid #c0c0c0')
    expect(build({ draft: true }).html).not.toContain('#c0c0c0')
  })

  it('numbers pages from the first page number', () => {
    expect(firstPageOffset({ firstPageNumber: 5 })).toBe(4)
    expect(firstPageOffset(undefined)).toBe(0)
    expect(
      renderHeaderFooterHtml('&P', 'b', 's', new Date(), undefined, 1, undefined, undefined, 4),
    ).toContain('data-genoffice-page-offset="+4"')
    expect(
      renderHeaderFooterHtml('&P+1', 'b', 's', new Date(), undefined, 1, undefined, undefined, 4),
    ).toContain('data-genoffice-page-offset="+5"')
    expect(renderHeaderFooterHtml('&P', 'b', 's', new Date())).toContain('class="pageNumber"')
    expect(build({ firstPageNumber: 3 }).footerTemplate).toContain(
      'data-genoffice-page-offset="+2"',
    )
  })

  it('knows when nothing departs from the defaults', () => {
    expect(isDefaultPrintOptions(undefined)).toBe(true)
    expect(isDefaultPrintOptions({ errorsAs: 'displayed', firstPageNumber: 1 })).toBe(true)
    expect(isDefaultPrintOptions({ draft: true })).toBe(false)
  })
})

describe('Mac print scope entries', () => {
  it('maps each File-menu entry to a print scope', () => {
    expect(printScopeOf('active')).toEqual({ scope: 'active' })
    expect(printScopeOf('workbook')).toEqual({ scope: 'workbook' })
    expect(printScopeOf('selection')).toEqual({ scope: 'selection' })
    expect(printScopeOf('ignore-area')).toEqual({ scope: 'active', ignorePrintArea: true })
  })
})
