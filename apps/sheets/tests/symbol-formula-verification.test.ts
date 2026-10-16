import { describe, expect, it, vi } from 'vitest'
import { IUniverInstanceService } from '@univerjs/core'
import { DocSelectionManagerService } from '@univerjs/docs'
import { IEditorBridgeService } from '@univerjs/sheets-ui'

import {
  applyFormulaSymbolDraft,
  captureFormulaSymbolDraft,
  insertIntoFormulaSource,
} from '../src/renderer/symbol-formula'

function makeRuntime(
  options: {
    readonly formula?: string
    readonly rawFormula?: string
    readonly row?: number
    readonly column?: number
    readonly live?: string
    readonly editRow?: number
    readonly editColumn?: number
    readonly start?: number
    readonly end?: number
    readonly selectionSubUnit?: 'self' | 'default'
    readonly currentSelection?: 'self' | 'default' | 'other'
    readonly staleOwn?: boolean
  } = {},
) {
  let activeRow = options.row ?? 4
  let activeColumn = options.column ?? 2
  const setValue = vi.fn()
  const getRange = vi.fn(() => ({
    getFormula: () => options.formula ?? '',
    getCellDatas: () => [[options.rawFormula ? { f: options.rawFormula } : {}]],
    setValue,
  }))
  const sheet = {
    getSheetId: () => 'sheet-1',
    getRange,
  }
  const workbook = {
    getId: () => 'book-1',
    getActiveSheet: () => sheet,
    getSheetBySheetId: (id: string) => (id === 'sheet-1' ? sheet : null),
    getActiveRange: () => ({
      getRow: () => activeRow,
      getColumn: () => activeColumn,
    }),
    moveActiveRange: (row: number, column: number) => {
      activeRow = row
      activeColumn = column
    },
  }
  const bridge = {
    isVisible: () => ({ visible: options.live !== undefined }),
    getEditCellState: () => ({
      sheetId: 'sheet-1',
      row: options.editRow ?? options.row ?? 4,
      column: options.editColumn ?? options.column ?? 2,
    }),
    getCurrentEditorId: () => 'editor-1',
  }
  const currentSelection = {
    unitId:
      options.currentSelection === 'other' || !options.currentSelection ? 'other' : 'editor-1',
    subUnitId:
      options.currentSelection === 'self'
        ? 'editor-1'
        : options.currentSelection === 'default'
          ? ''
          : 'other',
  }
  const rangesForSubUnit = (subUnitId: string) => {
    if (subUnitId === (options.selectionSubUnit === 'self' ? 'editor-1' : '')) {
      return [
        {
          startOffset: options.start ?? 0,
          endOffset: options.end ?? 0,
          isActive: true,
        },
      ]
    }
    if (subUnitId === 'editor-1' && options.staleOwn) {
      return [{ startOffset: 0, endOffset: 0, isActive: true }]
    }
    return undefined
  }
  const services = new Map<unknown, unknown>([
    [IEditorBridgeService, bridge],
    [
      IUniverInstanceService,
      { getUnit: () => ({ getBody: () => ({ dataStream: `${options.live ?? ''}\r\n` }) }) },
    ],
    [
      DocSelectionManagerService,
      {
        __getCurrentSelection: () => currentSelection,
        getTextRanges: (params: { unitId: string; subUnitId: string } = currentSelection) => {
          if (params.unitId !== 'editor-1') return undefined
          return rangesForSubUnit(params.subUnitId)
        },
      },
    ],
  ])
  const runtime = {
    univerAPI: { getActiveWorkbook: () => workbook },
    univer: { __getInjector: () => ({ get: (token: unknown) => services.get(token) }) },
  }
  return { runtime: runtime as never, workbook, sheet, getRange, setValue }
}

function lazyWithFormula(
  formula: string,
  options: { structural?: boolean; journaled?: boolean } = {},
) {
  return {
    editJournal: {
      structuralOps: new Map(options.structural ? [['sheet-1', [{}]]] : []),
      cells: new Map(options.journaled ? [['sheet-1', new Map([['4:2', {}]])]] : []),
    },
    formulaText: new Map([['sheet-1', new Map([['4:2', formula]])]]),
  } as never
}

describe('formula-symbol independent verification', () => {
  it('uses harvested formula text when the resident cell has no formula', () => {
    const { runtime } = makeRuntime()
    const captured = captureFormulaSymbolDraft(runtime, lazyWithFormula('=SUM(C1:C8)'))
    expect(captured?.draft.source).toBe('=SUM(C1:C8)')
    expect(captured?.draft).toMatchObject({
      workbookId: 'book-1',
      sheetId: 'sheet-1',
      row: 4,
      column: 2,
    })
  })

  it.each([
    ['a structurally shifted sheet', { structural: true }],
    ['a cell overwritten in this session', { journaled: true }],
  ])('does not reuse the harvested formula for %s', (_label, guards) => {
    const { runtime } = makeRuntime()
    expect(captureFormulaSymbolDraft(runtime, lazyWithFormula('=STALE(A1)', guards))).toBeNull()
  })

  it('writes to the original target even if the active range changes while the picker is open', () => {
    const { runtime, workbook, getRange, setValue } = makeRuntime({ formula: '=A1' })
    const draft = captureFormulaSymbolDraft(runtime, null)!.draft
    ;(
      workbook as unknown as { moveActiveRange: (row: number, column: number) => void }
    ).moveActiveRange(9, 6)
    expect(applyFormulaSymbolDraft(runtime, draft, '=A1+Ω')).toBe(true)
    expect(getRange).toHaveBeenLastCalledWith(4, 2, 1, 1)
    expect(setValue).toHaveBeenCalledWith(expect.objectContaining({ f: '=A1+Ω' }))
  })

  it('inserts supplementary-plane symbols using the source field UTF-16 selection offsets', () => {
    const source = '=A1+"ab"'
    const from = source.indexOf('ab')
    expect(insertIntoFormulaSource(source, from, from + 2, '🧪')).toEqual({
      source: '=A1+"🧪"',
      caret: from + 2,
    })
  })

  it('prefers the live editor source over a different persisted formula when editing the same cell', () => {
    const { runtime } = makeRuntime({ formula: '=OLD(A1)', live: '=NEW(Ω)', row: 4, column: 2 })
    const captured = captureFormulaSymbolDraft(runtime, null)
    expect(captured?.wasEditing).toBe(true)
    expect(captured?.draft.source).toBe('=NEW(Ω)')
  })

  it('reads the current Docs selection from the normal editor unit and its own subunit', () => {
    const source = '="Total "&C1'
    const { runtime } = makeRuntime({
      formula: source,
      live: source,
      start: 2,
      end: 7,
      selectionSubUnit: 'self',
      currentSelection: 'self',
    })
    const captured = captureFormulaSymbolDraft(runtime, null)
    expect(captured?.draft.selectionStart).toBe(2)
    expect(captured?.draft.selectionEnd).toBe(7)
  })

  it('prefers the current default-subunit selection over a stale own-subunit caret', () => {
    const source = '="Total "&C1'
    const { runtime } = makeRuntime({
      formula: source,
      live: source,
      start: 2,
      end: 7,
      selectionSubUnit: 'default',
      currentSelection: 'default',
      staleOwn: true,
    })
    const captured = captureFormulaSymbolDraft(runtime, null)
    expect(captured?.draft.selectionStart).toBe(2)
    expect(captured?.draft.selectionEnd).toBe(7)
  })

  it('does not attach a live editor in another cell to the active formula target', () => {
    const { runtime } = makeRuntime({
      formula: '=SUM(A1:A2)',
      live: 'draft',
      editRow: 8,
      editColumn: 8,
    })
    const captured = captureFormulaSymbolDraft(runtime, null)
    expect(captured?.wasEditing).toBe(false)
    expect(captured?.draft.source).toBe('=SUM(A1:A2)')
  })
})
