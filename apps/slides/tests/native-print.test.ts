import { describe, expect, it } from 'vitest'
import {
  DEFAULT_PRINT_PREFS,
  NATIVE_HANDOUT_LAYOUTS,
  nativePrintParams,
  normalizePrintPrefs,
  selectPrintSlides,
  usesMacPrintPanel,
} from '../src/shared/print-prefs'
import { handoutSpec, printPageCount } from '../src/shared/print-html'
import { printEntryFor } from '../src/renderer/print-entry'

describe('platform split', () => {
  it('only macOS print jobs flagged native open the system panel', () => {
    expect(usesMacPrintPanel('darwin', { nativePanel: true })).toBe(true)
    expect(usesMacPrintPanel('darwin', {})).toBe(false)
    expect(usesMacPrintPanel('win32', { nativePanel: true })).toBe(false)
    expect(usesMacPrintPanel('linux', { nativePanel: true })).toBe(false)
  })
  it('renderer entry follows navigator.platform', () => {
    expect(printEntryFor('MacIntel')).toBe('native-panel')
    expect(printEntryFor('Win32')).toBe('in-app')
    expect(printEntryFor('Linux x86_64')).toBe('in-app')
  })
})

describe('menu layouts', () => {
  it('every handout entry is a real handout (N per page, vertical or horizontal)', () => {
    expect(NATIVE_HANDOUT_LAYOUTS.map((l) => handoutSpec(l))).toEqual([
      { per: 1, vertical: false },
      { per: 2, vertical: false },
      { per: 3, vertical: false },
      { per: 4, vertical: false },
      { per: 4, vertical: true },
      { per: 6, vertical: false },
      { per: 6, vertical: true },
      { per: 9, vertical: false },
      { per: 9, vertical: true },
    ])
    expect(printPageCount(10, 'handout4v')).toBe(3)
  })
})

describe('normalizePrintPrefs', () => {
  it('fills defaults and drops invalid values', () => {
    expect(normalizePrintPrefs(null)).toEqual(DEFAULT_PRINT_PREFS)
    expect(
      normalizePrintPrefs({ layout: 'bogus', output: 'sepia', paper: 'huge', frame: 'yes' }),
    ).toEqual(DEFAULT_PRINT_PREFS)
    expect(
      normalizePrintPrefs({ layout: 'handout6v', output: 'bw', paper: 'letter', frame: true }),
    ).toMatchObject({ layout: 'handout6v', output: 'bw', paper: 'letter', frame: true })
  })
})

describe('selectPrintSlides', () => {
  const deck = [{}, { hidden: true }, {}, { hidden: true }]
  it('skips hidden slides unless asked', () => {
    expect(selectPrintSlides(deck, false)).toEqual([0, 2])
    expect(selectPrintSlides(deck, true)).toEqual([0, 1, 2, 3])
  })
})

describe('nativePrintParams', () => {
  const prefs = {
    ...DEFAULT_PRINT_PREFS,
    frame: true,
    scaleFit: true,
    output: 'grayscale' as const,
  }
  it('full layout on slide-sized paper: no orientation, no scale-to-fit', () => {
    expect(nativePrintParams('full', prefs)).toEqual({
      layout: 'full',
      paper: 'slide',
      frame: true,
      output: 'grayscale',
      scaleFit: false,
      orientation: null,
    })
  })
  it('full layout on real paper scales up and keeps the master orientation', () => {
    expect(nativePrintParams('full', { ...prefs, paper: 'a4' })).toMatchObject({
      paper: 'a4',
      scaleFit: true,
      orientation: 'portrait',
    })
  })
  it('handouts / notes / outline always use real paper', () => {
    expect(nativePrintParams('handout4h', prefs, 'landscape')).toMatchObject({
      paper: 'a4',
      scaleFit: false,
      orientation: 'landscape',
      frame: true,
    })
    expect(nativePrintParams('notes', prefs).frame).toBe(false)
    expect(nativePrintParams('outline', { ...prefs, paper: 'letter' }).paper).toBe('letter')
  })
})
