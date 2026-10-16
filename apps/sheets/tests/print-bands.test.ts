import { describe, expect, it } from 'vitest'

import {
  BAND_MAX_CELLS,
  PRINT_JOB_MAX_CELLS,
  areaCells,
  bandProgress,
  bandsOf,
  exceedsJobLimit,
  unitsForAreas,
} from '../src/renderer/print-bands'

const area = (rows: number, columns: number) => ({
  startRow: 0,
  startColumn: 0,
  endRow: rows - 1,
  endColumn: columns - 1,
})

describe('print bands', () => {
  it('keeps a small area as one unit', () => {
    const { banded, units } = unitsForAreas([area(100, 10)])
    expect(banded).toBe(false)
    expect(units).toHaveLength(1)
  })

  it('splits a big area into contiguous row bands that cover it exactly', () => {
    const big = area(100_001, 12)
    const bands = bandsOf(big)
    expect(bands[0]!.startRow).toBe(0)
    expect(bands[bands.length - 1]!.endRow).toBe(100_000)
    for (let i = 1; i < bands.length; i++) expect(bands[i]!.startRow).toBe(bands[i - 1]!.endRow + 1)
    for (const b of bands) expect(areaCells(b)).toBeLessThanOrEqual(BAND_MAX_CELLS)
    expect(bands.reduce((s, b) => s + areaCells(b), 0)).toBe(areaCells(big))
  })

  it('a band holds at least one row even for very wide areas', () => {
    expect(bandsOf(area(3, 50_000))).toHaveLength(3)
  })

  it('bands several areas in order, one unit per band', () => {
    const { banded, units } = unitsForAreas([area(30_000, 4), area(30_000, 4)])
    expect(banded).toBe(true)
    expect(units.every((u) => u.length === 1)).toBe(true)
    expect(units.length).toBe(bandsOf(area(30_000, 4)).length * 2)
  })

  it('flags jobs beyond the hard limit', () => {
    expect(exceedsJobLimit([area(1000, 1000)])).toBe(false)
    expect(exceedsJobLimit([area(PRINT_JOB_MAX_CELLS, 2)])).toBe(true)
  })

  it('clamps progress', () => {
    expect(bandProgress(5, 4)).toEqual({ done: 4, total: 4 })
    expect(bandProgress(-1, 4)).toEqual({ done: 0, total: 4 })
  })
})
