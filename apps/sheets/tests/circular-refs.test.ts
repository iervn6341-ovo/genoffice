import { describe, expect, it } from 'vitest'
import {
  circularCells,
  formulaDependencyTrees,
  formulaRefs,
  type DependencyTreeLike,
} from '../src/renderer/circular-refs'

const tree = (
  treeId: number,
  row: number,
  column: number,
  parents: number[],
  subUnitId = 's1',
): DependencyTreeLike => ({
  treeId,
  row,
  column,
  parents,
  subUnitId,
})

describe('circularCells', () => {
  it('finds a cell that reads itself (=C3 in C3)', () => {
    expect(circularCells([tree(0, 0, 1, []), tree(1, 2, 2, [1])])).toEqual([
      { subUnitId: 's1', row: 2, column: 2 },
    ])
  })

  it('finds every cell of a longer loop, top-left first, and ignores cells that only depend on it', () => {
    // D4 = E4 + 1, E4 = D4 + 1, F4 = D4 (reads the loop but is not on it)
    const out = circularCells([tree(0, 3, 4, [1]), tree(1, 3, 3, [0]), tree(2, 3, 5, [1])])
    expect(out).toEqual([
      { subUnitId: 's1', row: 3, column: 3 },
      { subUnitId: 's1', row: 3, column: 4 },
    ])
  })

  it('returns nothing for an acyclic graph, including shared precedents', () => {
    expect(circularCells([tree(0, 0, 1, []), tree(1, 1, 1, [0]), tree(2, 2, 1, [0, 1])])).toEqual(
      [],
    )
  })

  it('handles a 10,000-cell chain without recursion', () => {
    const chain = Array.from({ length: 10_000 }, (_, i) =>
      tree(i, i, 0, i === 0 ? [9_999] : [i - 1]),
    )
    expect(circularCells(chain)).toHaveLength(10_000)
  })
})

describe('formulaRefs', () => {
  it('reads cells, ranges, anchors, whole columns/rows and sheet-qualified refs', () => {
    expect(formulaRefs("=A1+$B$2+SUM(C3:D4)+SUM(E:E)+SUM(2:3)+'Data 2024'!F6+Sheet2!G7")).toEqual([
      { sheet: null, r1: 0, c1: 0, r2: 0, c2: 0 },
      { sheet: null, r1: 1, c1: 1, r2: 1, c2: 1 },
      { sheet: null, r1: 2, c1: 2, r2: 3, c2: 3 },
      { sheet: null, r1: 0, c1: 4, r2: 1_048_575, c2: 4 },
      { sheet: null, r1: 1, c1: 0, r2: 2, c2: 16_383 },
      { sheet: 'Data 2024', r1: 5, c1: 5, r2: 5, c2: 5 },
      { sheet: 'Sheet2', r1: 6, c1: 6, r2: 6, c2: 6 },
    ])
  })

  it('ignores function names, string literals and names that look like cells', () => {
    expect(formulaRefs('=LOG10(A1)&"B2"&ATAN2(1,2)')).toEqual([
      { sheet: null, r1: 0, c1: 0, r2: 0, c2: 0 },
    ])
  })
})

describe('formulaDependencyTrees + circularCells', () => {
  const names = new Map([
    ['Sheet1', 's1'],
    ['Data 2024', 's2'],
  ])
  it('finds a cross-sheet loop and a range that contains its own cell', () => {
    const trees = formulaDependencyTrees(
      [
        { sheetId: 's1', row: 0, column: 0, formula: "='Data 2024'!A1+1" },
        { sheetId: 's2', row: 0, column: 0, formula: '=Sheet1!A1*2' },
        { sheetId: 's1', row: 9, column: 1, formula: '=SUM(B1:B10)' },
        { sheetId: 's1', row: 4, column: 4, formula: '=A1' },
      ],
      names,
    )
    expect(circularCells(trees)).toEqual([
      { subUnitId: 's1', row: 0, column: 0 },
      { subUnitId: 's1', row: 9, column: 1 },
      { subUnitId: 's2', row: 0, column: 0 },
    ])
  })
})
