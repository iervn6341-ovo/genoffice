import { CellValueType } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { installCachedValueFallbackInterceptor } from '../src/renderer/formula-cached-fallback'

interface Interceptor {
  handler: (
    cell: Record<string, unknown> | undefined,
    location: { subUnitId: string; row: number; col: number; rawData?: { f?: string } },
    next: (cell: Record<string, unknown> | undefined) => unknown,
  ) => unknown
}

function captureInterceptor(state: unknown): Interceptor {
  let captured: Interceptor | undefined
  const runtime = {
    univer: {
      __getInjector: () => ({
        get: () => ({
          intercept: (_point: unknown, interceptor: Interceptor) => {
            captured = interceptor
            return { dispose: () => undefined }
          },
        }),
      }),
    },
  }
  installCachedValueFallbackInterceptor(runtime as never, { current: state } as never)
  if (!captured) throw new Error('interceptor not registered')
  return captured
}

function makeState() {
  return {
    file: {
      sheets: [
        { id: 'data', name: 'Data' },
        { id: 'inputs', name: 'Inputs' },
        { id: 'calc', name: 'Calc Sheet' },
      ],
    },
    editJournal: { structuralOps: new Map(), cells: new Map() },
    cachedFormulaValues: new Map([['data', new Map([['0:0', 1]])]]),
    formulaText: new Map(),
  }
}

const passthrough = (cell: Record<string, unknown> | undefined) => cell
const dataCell = { subUnitId: 'data', row: 0, col: 0 }

describe('cached formula fallback dependency verification', () => {
  it('exposes a recalculated error for an edited cell inside a cross-sheet range', () => {
    const state = makeState()
    state.editJournal.cells.set('inputs', new Map([['1:0', { hasValue: true }]]))
    const { handler } = captureInterceptor(state)
    const error = { v: '#DIV/0!' }

    expect(
      handler(error, { ...dataCell, rawData: { f: '=1/(SUM(Inputs!A1:A2)-40)' } }, passthrough),
    ).toBe(error)
  })

  it('matches a cross-sheet qualifier without depending on sheet-name casing', () => {
    const state = makeState()
    state.editJournal.cells.set('inputs', new Map([['0:0', { hasValue: true }]]))
    const { handler } = captureInterceptor(state)
    const error = { v: '#DIV/0!' }

    expect(handler(error, { ...dataCell, rawData: { f: '=1/(inputs!A1-40)' } }, passthrough)).toBe(
      error,
    )
  })

  it('follows a cross-sheet range through an intermediate formula on a quoted sheet', () => {
    const state = makeState()
    state.editJournal.cells.set('inputs', new Map([['1:0', { hasValue: true }]]))
    state.formulaText.set('calc', new Map([['0:0', '=SUM(Inputs!$A$1:$A$2)']]))
    const { handler } = captureInterceptor(state)
    const error = { v: '#DIV/0!' }

    expect(
      handler(error, { ...dataCell, rawData: { f: "=1/('Calc Sheet'!A1-40)" } }, passthrough),
    ).toBe(error)
  })

  it('terminates a cyclic formula chain and keeps the cached result for an unrelated edit', () => {
    const state = makeState()
    state.editJournal.cells.set('inputs', new Map([['9:9', { hasValue: true }]]))
    state.formulaText.set(
      'calc',
      new Map([
        ['0:0', '=B1'],
        ['0:1', '=A1'],
      ]),
    )
    const { handler } = captureInterceptor(state)
    const error = { v: '#DIV/0!' }

    expect(
      handler(error, { ...dataCell, rawData: { f: "=1/('Calc Sheet'!A1-40)" } }, passthrough),
    ).toEqual({ v: 1, t: CellValueType.NUMBER })
  })

  it('keeps the cached result for a wide range when there are no content edits', () => {
    const state = makeState()
    const formulas = new Map<string, string>()
    for (let row = 0; row < 11_000; row += 1) formulas.set(`${row}:1`, '=B1+1')
    state.formulaText.set('calc', formulas)
    const { handler } = captureInterceptor(state)

    expect(
      handler(
        { v: '#DIV/0!' },
        { ...dataCell, rawData: { f: "=SUM('Calc Sheet'!A1:A11000)" } },
        passthrough,
      ),
    ).toEqual({ v: 1, t: CellValueType.NUMBER })
  })

  it('bounds a long dependency chain and conservatively shows the engine error', () => {
    const state = makeState()
    state.editJournal.cells.set('inputs', new Map([['0:0', { hasValue: true }]]))
    const formulas = new Map<string, string>()
    for (let row = 0; row < 1_100; row += 1) {
      formulas.set(`${row}:0`, row === 1_099 ? '=Inputs!A1' : `=A${row + 2}`)
    }
    state.formulaText.set('calc', formulas)
    const { handler } = captureInterceptor(state)
    const error = { v: '#DIV/0!' }

    expect(
      handler(error, { ...dataCell, rawData: { f: "=1/('Calc Sheet'!A1-40)" } }, passthrough),
    ).toBe(error)
  })
})
