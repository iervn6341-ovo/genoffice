import { describe, expect, it } from 'vitest'
import type { IStyleData } from '@univerjs/core'

import { fillMissingMergedBorderSegments } from '../src/renderer/merge-border-fix'

const MERGE = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }

function cache() {
  const values = new Map<string, Record<string, unknown>>()
  return {
    getValue: (row: number, col: number) => values.get(`${row}:${col}`),
    setValue: (row: number, col: number, value: Record<string, unknown>) => {
      values.set(`${row}:${col}`, value)
    },
  }
}

describe('sparse merged-cell border rendering', () => {
  it('spans an anchor-only outline across every perimeter segment without writing a cell', () => {
    const border = { s: 1, cl: { rgb: '#112233' } }
    const style = { bd: { t: border, b: border, l: border, r: border } } as IStyleData
    const render = cache()
    fillMissingMergedBorderSegments(render, MERGE, style)
    expect(render.getValue(0, 0)).toMatchObject({ t: { style: 1 }, l: { style: 1 } })
    expect(render.getValue(0, 1)).toMatchObject({ t: { style: 1 }, r: { style: 1 } })
    expect(render.getValue(1, 0)).toMatchObject({ b: { style: 1 }, l: { style: 1 } })
    expect(render.getValue(1, 1)).toMatchObject({ b: { style: 1 }, r: { style: 1 } })
  })

  it('does not overwrite a native covered-cell segment with a different edge style', () => {
    const render = cache()
    render.setValue(0, 1, { t: { type: 't', style: 7, color: 'rgb(1,2,3)' } })
    fillMissingMergedBorderSegments(render, MERGE, {
      bd: { t: { s: 1, cl: { rgb: '#112233' } }, l: { s: 1, cl: { rgb: '#112233' } } },
    } as IStyleData)
    expect(render.getValue(0, 1)).toEqual({
      t: { type: 't', style: 7, color: 'rgb(1,2,3)' },
    })
    expect(render.getValue(1, 0)).toMatchObject({ l: { style: 1 } })
    expect(render.getValue(1, 1)).toBeUndefined()
  })
})
