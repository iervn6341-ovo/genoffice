import { describe, expect, it } from 'vitest'

import {
  blendedDisplayOf,
  blendPatternColors,
  FILL_PATTERN_KEY,
  fillPatternOf,
  markerFromFile,
} from '../src/renderer/fill-pattern'
import { createEditJournal, recordSetRangeValues, toSaveEdits } from '../src/renderer/edit-journal'
import { draftFromSelection, formatCellsCommands } from '../src/renderer/format-cells'
import { toSelectionFormat } from '../src/renderer/selection-format'

const marker = { pattern: 'darkHorizontal', fg: '#FF0000', bg: '#0000FF' } as const

describe('pattern fill model', () => {
  it('blends like the file importer (per-channel floor average)', () => {
    expect(blendPatternColors('#FF0000', '#0000FF')).toBe('#7F007F')
    expect(blendedDisplayOf({ pattern: 'lightGray', fg: '#000000', bg: '#FFFFFF' })).toBe('#7F7F7F')
    expect(blendedDisplayOf({ pattern: 'solid', fg: '#000000', bg: '#FFFFFF' })).toBeNull()
  })

  it('trusts a marker only while the cell still shows the blend', () => {
    const custom = { [FILL_PATTERN_KEY]: marker }
    expect(fillPatternOf(custom, '#7f007f')).toEqual(marker)
    expect(fillPatternOf(custom, '#FFFF00')).toBeNull()
    expect(fillPatternOf(custom, null)).toBeNull()
    expect(
      fillPatternOf({ [FILL_PATTERN_KEY]: { pattern: 'bogus', fg: '#000', bg: '#fff' } }, null),
    ).toBeNull()
  })

  it('imports two-color patterns only', () => {
    expect(
      markerFromFile({ patternType: 'darkGrid', foreground: '#112233', background: '#FFFFFF' }),
    ).toEqual({ pattern: 'darkGrid', fg: '#112233', bg: '#FFFFFF' })
    expect(markerFromFile({ patternType: 'darkGrid', foreground: '#112233' })).toBeNull()
    expect(markerFromFile(undefined)).toBeNull()
  })
})

describe('pattern fill journal and dialog', () => {
  it('journals the marker as a native pattern spec and a later plain fill replaces it', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 's', {
      0: { 0: { s: { bg: { rgb: '#7F007F' } }, custom: { [FILL_PATTERN_KEY]: marker } } },
    })
    expect(toSaveEdits(journal)[0]?.style?.fill).toEqual({
      pattern: 'darkHorizontal',
      fg: '#FF0000',
      bg: '#0000FF',
    })
    recordSetRangeValues(journal, 's', { 0: { 0: { s: { bg: { rgb: '#00FF00' } } } } })
    const style = toSaveEdits(journal)[0]?.style
    expect(style?.fill).toBeUndefined()
    expect(style?.fillColor).toBe('#00FF00')
  })

  it('undoing back to a cell without the marker also drops the journaled pattern', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 's', {
      0: { 0: { s: { bg: { rgb: '#7F007F' } }, custom: { [FILL_PATTERN_KEY]: marker } } },
    })
    recordSetRangeValues(journal, 's', {
      0: { 0: { s: { bg: { rgb: '' } }, custom: null } },
    })
    expect(toSaveEdits(journal)[0]?.style?.fill).toBeUndefined()
  })

  it('emits a fill-pattern command only for a changed pattern and rewrites on removal', () => {
    const base = toSelectionFormat({ bg: { rgb: '#7F007F' } }, 'General', null, marker)
    const initial = draftFromSelection(base)
    expect(initial).toMatchObject({
      fill: '#0000ff',
      fillPattern: 'darkHorizontal',
      patternColor: '#ff0000',
    })
    expect(formatCellsCommands(initial, initial)).toEqual([])
    expect(formatCellsCommands(initial, { ...initial, fillPattern: 'darkGrid' })).toEqual([
      'fill-pattern:darkGrid:#ff0000:#0000ff',
    ])
    expect(formatCellsCommands(initial, { ...initial, fillPattern: '' })).toEqual(['fill:#0000ff'])
    const plain = draftFromSelection(toSelectionFormat({}, 'General'))
    expect(
      formatCellsCommands(plain, { ...plain, fillPattern: 'lightUp', patternColor: '#000000' }),
    ).toEqual(['fill-pattern:lightUp:#000000:#ffffff'])
  })
})
