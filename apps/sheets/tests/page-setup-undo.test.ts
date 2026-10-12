import { describe, expect, it } from 'vitest'

import { pageSetupPatch } from '../src/renderer/page-setup-undo'
import { MARGIN_PRESETS } from '../src/renderer/print-settings'

const portrait = {
  orientation: 'portrait' as const,
  paperSize: 9,
  scale: 100,
  fitToWidth: 0,
  fitToHeight: 0,
  fitToPage: false,
  margins: MARGIN_PRESETS.normal,
  printGridlines: false,
  printHeadings: false,
  printAreas: [] as readonly string[],
  printTitles: null,
  header: null,
  footer: null,
}

describe('⌘Z after a save writes the difference back into the journal', () => {
  it('has nothing to write when the setup in force is the one the snapshot remembers', () => {
    expect(pageSetupPatch(portrait, portrait)).toEqual({})
  })

  it('writes the orientation, paper size and scale the file now lacks', () => {
    const saved = { ...portrait, orientation: 'landscape' as const, paperSize: 1, scale: 80 }
    expect(pageSetupPatch(saved, portrait)).toEqual({
      orientation: 'portrait',
      paperSize: 9,
      scale: 100,
    })
  })

  it('maps margins back to the preset they were, and leaves margins that are no preset alone', () => {
    const wide = { ...portrait, margins: MARGIN_PRESETS.wide }
    expect(pageSetupPatch(wide, portrait)).toEqual({ margins: 'normal' })
    const custom = { ...portrait, margins: { ...MARGIN_PRESETS.normal, left: 0.9 } }
    // the setup to restore has margins no preset names: nothing the journal can write
    expect(pageSetupPatch(portrait, custom)).toEqual({})
  })

  it('writes a print area, print titles and header or footer back, or clears them', () => {
    const edited = {
      ...portrait,
      printAreas: ['A1:C9'],
      printTitles: '1:1',
      header: { center: 'Report' },
    }
    expect(pageSetupPatch(edited, portrait)).toEqual({
      printArea: null,
      printTitles: null,
      header: null,
    })
    expect(pageSetupPatch(portrait, edited)).toEqual({
      printArea: 'A1:C9',
      printTitles: '1:1',
      header: { center: 'Report' },
    })
  })

  it('writes fit-to-page and gridline switches', () => {
    const fit = { ...portrait, fitToPage: true, fitToWidth: 1, printGridlines: true }
    expect(pageSetupPatch(fit, portrait)).toEqual({
      fitToPage: false,
      fitToWidth: 0,
      printGridlines: false,
    })
  })
})

describe('manual page breaks come back after a save too', () => {
  const breaks = (rowBreaks: number[], colBreaks: number[] = []) => ({ rowBreaks, colBreaks })

  it('writes the break set the snapshot knew when the file now has another', () => {
    const before = { ...portrait, breaks: breaks([10]) }
    const saved = { ...portrait, breaks: breaks([10, 20], [4]) }
    expect(pageSetupPatch(saved, before)).toEqual({ rowBreaks: [10], colBreaks: [] })
  })

  it('writes only the axis that differs, and nothing when the sets agree', () => {
    const before = { ...portrait, breaks: breaks([10], [4]) }
    expect(pageSetupPatch({ ...portrait, breaks: breaks([10], [4]) }, before)).toEqual({})
    expect(pageSetupPatch({ ...portrait, breaks: breaks([10], [4, 8]) }, before)).toEqual({
      colBreaks: [4],
    })
  })

  it('writes the remembered set when the file has not been indexed yet, and leaves an unknown one alone', () => {
    expect(
      pageSetupPatch({ ...portrait, breaks: null }, { ...portrait, breaks: breaks([3]) }),
    ).toEqual({ rowBreaks: [3], colBreaks: [] })
    expect(
      pageSetupPatch({ ...portrait, breaks: breaks([3]) }, { ...portrait, breaks: null }),
    ).toEqual({})
  })
})
