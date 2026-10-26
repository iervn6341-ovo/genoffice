import { describe, expect, it } from 'vitest'

import { multiSelectionLayout } from '../src/renderer/clipboard-multi'

const r = (startRow: number, endRow: number, startColumn: number, endColumn: number) => ({
  startRow,
  endRow,
  startColumn,
  endColumn,
})

describe('copying several selected areas', () => {
  it('lays areas with the same rows side by side', () => {
    const layout = multiSelectionLayout([r(9, 10, 0, 1), r(9, 10, 3, 4)])!
    expect(layout.bounds).toEqual(r(9, 10, 0, 4))
    expect(layout.rows([9, 10])).toEqual([9, 10])
    expect(layout.columns([0, 1, 2, 3, 4])).toEqual([0, 1, 3, 4])
  })
  it('stacks areas with the same columns', () => {
    const layout = multiSelectionLayout([r(0, 1, 2, 3), r(5, 6, 2, 3)])!
    expect(layout.rows([0, 1, 2, 3, 4, 5, 6])).toEqual([0, 1, 5, 6])
    expect(layout.columns([2, 3])).toEqual([2, 3])
  })
  it('keeps the rows a filter left in the bounding copy only', () => {
    const layout = multiSelectionLayout([r(0, 3, 0, 0), r(0, 3, 2, 2)])!
    expect(layout.rows([0, 2, 3])).toEqual([0, 2, 3])
  })
  it('refuses areas that do not line up, and a single area', () => {
    expect(multiSelectionLayout([r(0, 1, 0, 1), r(3, 5, 3, 4)])).toBeNull()
    expect(multiSelectionLayout([r(0, 1, 0, 1)])).toBeNull()
  })
})
