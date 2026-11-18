import { describe, expect, it } from 'vitest'

import { liveFormulaCacheValues } from '../src/renderer/save-actions'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'

/// The state slice liveFormulaCacheValues reads, no Univer needed: the
/// journal's cells, the formula-text harvest and the file cache it pairs with.
function state(parts: {
  journal?: Map<
    string,
    Map<string, { formula?: string; value?: string | number | boolean | null; hasValue?: boolean }>
  >
  removed?: Set<string>
  formulas?: Map<string, Map<string, string>>
  cached?: Map<string, Map<string, string | number | boolean>>
  overlay?: Map<string, Map<string, unknown>>
  external?: Map<string, Map<string, unknown>>
  structural?: Map<string, { kind: string; [key: string]: unknown }[]>
}): LazyWorkbookState {
  return {
    editJournal: {
      cells: parts.journal ?? new Map(),
      sheets: { removed: parts.removed ?? new Set() },
      structuralOps: parts.structural ?? new Map(),
    },
    recalc: {
      overlay: parts.overlay ?? new Map(),
      externalOverlay: parts.external ?? new Map(),
      incompleteRuns: new Set(),
    },
    formulaText: parts.formulas ?? new Map(),
    cachedFormulaValues: parts.cached ?? new Map(),
    sharedFormulaGroups: new Map(),
  } as unknown as LazyWorkbookState
}

/// A model answer for one address; `formula` omitted mimics a withheld cell,
/// which answers with a placeholder and no formula text.
function cellOf(value: string | number | boolean | null, formula?: string): CellState {
  return {
    value,
    rawValue: value,
    ...(formula === undefined ? {} : { formula }),
  }
}

/// Fake `ctx.readCells` that records every requested address and returns
/// what was staged. Addresses arrive in A1 notation ('B2', 'AA10').
function reader(staged: Map<string, Record<string, CellState>>): {
  (addresses: string[], sheetId: string): Record<string, CellState>
  asked: string[]
} {
  const asked: string[] = []
  const fn = (addresses: string[], sheetId: string) => {
    for (const a of addresses) asked.push(`${sheetId}!${a}`)
    const sheet = staged.get(sheetId) ?? {}
    return Object.fromEntries(addresses.map((a) => [a, sheet[a] ?? cellOf(null)]))
  }
  fn.asked = asked
  return fn
}

describe('liveFormulaCacheValues', () => {
  it('reads every journal-owned formula cell unconditionally', () => {
    const journal = new Map([
      [
        'sheet-1',
        new Map([
          ['1:1', { formula: '=A2*2', value: null }],
          ['2:3', { formula: '=SUM(A1:C1)', value: null }],
        ]),
      ],
    ])
    const read = reader(
      new Map([
        [
          'sheet-1',
          {
            B2: cellOf(10, '=A2*2'),
            D3: cellOf(6, '=SUM(A1:C1)'),
          },
        ],
      ]),
    )
    expect(liveFormulaCacheValues(state({ journal }), read)).toEqual([
      { sheetId: 'sheet-1', row: 1, column: 1, value: 10 },
      { sheetId: 'sheet-1', row: 2, column: 3, value: 6 },
    ])
  })

  it('skips a journal formula cell whose model answer carries no formula', () => {
    const journal = new Map([['sheet-1', new Map([['1:1', { formula: '=A2*2', value: null }]])]])
    const read = reader(new Map([['sheet-1', { B2: cellOf(10) }]]))
    expect(liveFormulaCacheValues(state({ journal }), read)).toEqual([])
  })

  it('reads non-journal formula cells whose live result disagrees with the file cache', () => {
    const read = reader(new Map([['sheet-1', { B2: cellOf(10, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 6]])]]),
        }),
        read,
      ),
    ).toEqual([{ sheetId: 'sheet-1', row: 1, column: 1, value: 10 }])
  })

  it('skips non-journal formula cells that agree with the file cache', () => {
    const read = reader(new Map([['sheet-1', { B2: cellOf(6, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 6]])]]),
        }),
        read,
      ),
    ).toEqual([])
  })

  it('treats floating-point numbers within 1e-9 relative as agreeing', () => {
    // 0.1 + 0.2 vs 0.3-class precision differences must not re-dirty the file
    const read = reader(new Map([['sheet-1', { B2: cellOf(1.0000000000000002, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 1]])]]),
        }),
        read,
      ),
    ).toEqual([])
  })

  it('skips cells the engine or external-link overlay already contributed', () => {
    const read = reader(
      new Map([['sheet-1', { B2: cellOf(10, '=A2*2'), C3: cellOf(20, '=A3*2') }]]),
    )
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([
            [
              'sheet-1',
              new Map([
                ['1:1', '=A2*2'],
                ['2:2', '=A3*2'],
              ]),
            ],
          ]),
          cached: new Map([
            [
              'sheet-1',
              new Map([
                ['1:1', 6],
                ['2:2', 6],
              ]),
            ],
          ]),
          overlay: new Map([['sheet-1', new Map([['1:1', { v: 10 }]])]]),
          external: new Map([['sheet-1', new Map([['2:2', { v: 20 }]])]]),
        }),
        read,
      ),
    ).toEqual([])
    expect(read.asked).toEqual([])
  })

  it('skips non-journal cells the journal already owns as a literal edit', () => {
    // Cell B2 was typed over as a constant in the journal: its formulaText
    // entry is the document before the edit, and comparing it would read the
    // new constant as if it were a formula result.
    const journal = new Map([['sheet-1', new Map([['1:1', { value: 5, hasValue: true }]])]])
    const read = reader(new Map([['sheet-1', { B2: cellOf(10, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          journal,
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 6]])]]),
        }),
        read,
      ),
    ).toEqual([])
  })

  it('caches an error literal as one, skips the engine-internal failure', () => {
    const journal = new Map([
      ['sheet-1', new Map([['0:0', { formula: '=1/0', value: null }]])],
      ['sheet-2', new Map([['0:0', { formula: '=BROKEN()', value: null }]])],
    ])
    const read = reader(
      new Map([
        ['sheet-1', { A1: cellOf('#DIV/0!', '=1/0') }],
        ['sheet-2', { A1: cellOf('#ERROR!', '=BROKEN()') }],
      ]),
    )
    expect(liveFormulaCacheValues(state({ journal }), read)).toEqual([
      { sheetId: 'sheet-1', row: 0, column: 0, value: { error: '#DIV/0!' } },
    ])
  })

  it('skips comparison cells whose saved form the session never saw', () => {
    const read = reader(new Map([['sheet-1', { D4: cellOf(10, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({ formulas: new Map([['sheet-1', new Map([['3:3', '=A2*2']])]]) }),
        read,
      ),
    ).toEqual([])
  })

  it('skips a removed sheet entirely', () => {
    const journal = new Map([['sheet-1', new Map([['0:0', { formula: '=A2', value: null }]])]])
    const read = reader(new Map([['sheet-1', { A1: cellOf(1, '=A2') }]]))
    expect(liveFormulaCacheValues(state({ journal, removed: new Set(['sheet-1']) }), read)).toEqual(
      [],
    )
    expect(read.asked).toEqual([])
  })

  it('reads the live grid with A1 addresses, never with a row:col key', () => {
    const journal = new Map([['sheet-1', new Map([['2:2', { formula: '=A2+1', value: null }]])]])
    const read = reader(new Map([['sheet-1', { C3: cellOf(6, '=A2+1') }]]))
    liveFormulaCacheValues(
      state({
        journal,
        formulas: new Map([['sheet-1', new Map([['3:3', '=A2+1']])]]),
        cached: new Map([['sheet-1', new Map([['3:3', 6]])]]),
      }),
      read,
    )
    expect(read.asked).toEqual(['sheet-1!C3', 'sheet-1!D4'])
  })

  it('keeps the file cache when the engine cannot evaluate the formula (#NAME?)', () => {
    const read = reader(new Map([['sheet-1', { B2: cellOf('#NAME?', '=FOO(A2)') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=FOO(A2)']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 42]])]]),
        }),
        read,
      ),
    ).toEqual([])
  })

  it('keeps the file cache over #NAME? even once the sheet has been edited', () => {
    // The engine cannot evaluate that formula under any inputs, so the file's
    // own result remains the best value — an edit elsewhere does not change
    // that (the display-only fallback switches off, the save must not).
    const journal = new Map([['sheet-1', new Map([['0:0', { value: 5, hasValue: true }]])]])
    const read = reader(new Map([['sheet-1', { B2: cellOf('#NAME?', '=FOO(A2)') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          journal,
          formulas: new Map([['sheet-1', new Map([['1:1', '=FOO(A2)']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 42]])]]),
        }),
        read,
      ),
    ).toEqual([])
  })

  it('keeps the file cache when a computed error lands on a usable cache before any edit', () => {
    const read = reader(new Map([['sheet-1', { B2: cellOf('#DIV/0!', '=A2/A3') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2/A3']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 42]])]]),
        }),
        read,
      ),
    ).toEqual([])
  })

  it('writes the computed error once the sheet has been edited', () => {
    // After an edit the error may be the true result of the new inputs, so it
    // replaces the file's cache.
    const journal = new Map([['sheet-1', new Map([['0:0', { value: 0, hasValue: true }]])]])
    const read = reader(new Map([['sheet-1', { B2: cellOf('#DIV/0!', '=A2/A3') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          journal,
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2/A3']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 42]])]]),
        }),
        read,
      ),
    ).toEqual([{ sheetId: 'sheet-1', row: 1, column: 1, value: { error: '#DIV/0!' } }])
  })

  it('an edit on another sheet lets a dependent write its computed blank or error', () => {
    // Sheet 2 sums a cell the user emptied on sheet 1: the blank is the true
    // result, not a capability gap, so the flag is workbook-wide.
    const journal = new Map([['sheet-1', new Map([['0:0', { value: null, hasValue: true }]])]])
    const read = reader(new Map([['sheet-2', { A1: cellOf('#DIV/0!', '=1/Sheet1!A1') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          journal,
          formulas: new Map([['sheet-2', new Map([['0:0', '=1/Sheet1!A1']])]]),
          cached: new Map([['sheet-2', new Map([['0:0', 0.5]])]]),
        }),
        read,
      ),
    ).toEqual([{ sheetId: 'sheet-2', row: 0, column: 0, value: { error: '#DIV/0!' } }])
  })

  it('a formula-only journal entry counts as an edit', () => {
    const journal = new Map([
      ['sheet-1', new Map([['0:0', { formula: '=1', value: null, hasValue: false }]])],
    ])
    const read = reader(
      new Map([['sheet-1', { A1: cellOf(1, '=1'), B2: cellOf('#DIV/0!', '=A2/A3') }]]),
    )
    expect(
      liveFormulaCacheValues(
        state({
          journal,
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2/A3']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 42]])]]),
        }),
        read,
      ),
    ).toContainEqual({ sheetId: 'sheet-1', row: 1, column: 1, value: { error: '#DIV/0!' } })
  })

  it('skips the comparison on a sheet whose rows shifted', () => {
    // The formula text is in file coordinates after an insert, while the cache
    // and the live grid are in screen coordinates: comparing them would
    // rewrite the wrong cells.
    const read = reader(new Map([['sheet-1', { B2: cellOf(10, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 6]])]]),
          structural: new Map([['sheet-1', [{ kind: 'insert-rows', index: 0, count: 1 }]]]),
        }),
        read,
      ),
    ).toEqual([])
    expect(read.asked).toEqual([])
  })

  it("still reads the journal's own formula cells on a shifted sheet", () => {
    const journal = new Map([['sheet-1', new Map([['2:2', { formula: '=A2+1', value: null }]])]])
    const read = reader(new Map([['sheet-1', { C3: cellOf(6, '=A2+1') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          journal,
          structural: new Map([['sheet-1', [{ kind: 'remove-rows', index: 0, count: 1 }]]]),
        }),
        read,
      ),
    ).toEqual([{ sheetId: 'sheet-1', row: 2, column: 2, value: 6 }])
  })

  it('a layout-only structural op leaves the comparison alone', () => {
    const read = reader(new Map([['sheet-1', { B2: cellOf(10, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 6]])]]),
          structural: new Map([
            ['sheet-1', [{ kind: 'set-row-size', start: 0, end: 0, size: 20 }]],
          ]),
        }),
        read,
      ),
    ).toEqual([{ sheetId: 'sheet-1', row: 1, column: 1, value: 10 }])
  })
})
