import { describe, expect, it, vi } from 'vitest'
import { DocSelectionManagerService } from '@univerjs/docs'
import { IEditorBridgeService } from '@univerjs/sheets-ui'
import { IUniverInstanceService } from '@univerjs/core'
import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

import { buildEditFixture } from './fixture-builder'
import { createEditJournal, recordSetRangeValues, toSaveEdits } from '../src/renderer/edit-journal'
import {
  applyFormulaSymbolDraft,
  captureFormulaSymbolDraft,
  formulaSymbolSourceProblem,
  insertIntoFormulaSource,
} from '../src/renderer/symbol-formula'

function runtimeWithFormula(
  formula: string,
  options: {
    live?: string
    start?: number
    end?: number
    selectionSubUnit?: 'self' | 'default'
    currentSelection?: 'self' | 'default' | 'other'
    staleOwn?: boolean
  } = {},
) {
  const setValue = vi.fn()
  const cell = { getFormula: () => formula, getCellDatas: () => [[{ f: formula }]], setValue }
  const sheet = {
    getSheetId: () => 'sheet-1',
    getRange: () => cell,
  }
  const workbook = {
    getId: () => 'book-1',
    getActiveSheet: () => sheet,
    getSheetBySheetId: (id: string) => (id === 'sheet-1' ? sheet : null),
    getActiveRange: () => ({ getRow: () => 2, getColumn: () => 1 }),
  }
  const bridge = {
    isVisible: () => ({ visible: options.live !== undefined }),
    getEditCellState: () => ({ sheetId: 'sheet-1', row: 2, column: 1 }),
    getCurrentEditorId: () => 'editor',
  }
  const currentSelection = {
    unitId: options.currentSelection === 'other' || !options.currentSelection ? 'other' : 'editor',
    subUnitId:
      options.currentSelection === 'self'
        ? 'editor'
        : options.currentSelection === 'default'
          ? ''
          : 'other',
  }
  const services = new Map<unknown, unknown>([
    [IEditorBridgeService, bridge],
    [
      IUniverInstanceService,
      {
        getUnit: () => ({ getBody: () => ({ dataStream: `${options.live ?? ''}\r\n` }) }),
      },
    ],
    [
      DocSelectionManagerService,
      {
        __getCurrentSelection: () => currentSelection,
        getTextRanges: (params: { unitId: string; subUnitId: string } = currentSelection) => {
          if (params.unitId !== 'editor') return undefined
          if (params.subUnitId === (options.selectionSubUnit === 'self' ? 'editor' : '')) {
            return [
              { startOffset: options.start ?? 0, endOffset: options.end ?? 0, isActive: true },
            ]
          }
          return params.subUnitId === 'editor' && options.staleOwn
            ? [{ startOffset: 0, endOffset: 0, isActive: true }]
            : undefined
        },
      },
    ],
  ])
  const runtime = {
    univerAPI: { getActiveWorkbook: () => workbook },
    univer: { __getInjector: () => ({ get: (token: unknown) => services.get(token) }) },
  }
  return { runtime: runtime as never, workbook, cell, setValue }
}

describe('symbol insertion into formula source', () => {
  it('captures live source and caret before dialog focus loss, then replaces only the selected literal text', () => {
    const source = '="Total "&SUM(B1:B2)'
    const start = source.indexOf('Total')
    const { runtime } = runtimeWithFormula('=0', { live: source, start, end: start + 5 })
    const captured = captureFormulaSymbolDraft(runtime, null)
    expect(captured?.wasEditing).toBe(true)
    expect(captured?.draft.source).toBe(source)
    const inserted = insertIntoFormulaSource(
      captured!.draft.source,
      captured!.draft.selectionStart,
      captured!.draft.selectionEnd,
      'Ω',
    )
    expect(inserted).toEqual({ source: '="Ω "&SUM(B1:B2)', caret: start + 1 })
  })

  it('also captures an editor-unit sub-selection when Univer stores it under its own ID', () => {
    const source = '="A"&C1'
    const { runtime } = runtimeWithFormula(source, {
      live: source,
      start: 2,
      end: 3,
      selectionSubUnit: 'self',
    })
    const captured = captureFormulaSymbolDraft(runtime, null)
    expect(captured?.draft.selectionStart).toBe(2)
    expect(captured?.draft.selectionEnd).toBe(3)
  })

  it('prefers the focused default selection over a stale editor-owned zero caret', () => {
    const source = '="Total "&C1'
    const { runtime } = runtimeWithFormula(source, {
      live: source,
      start: 2,
      end: 7,
      currentSelection: 'default',
      staleOwn: true,
    })
    const captured = captureFormulaSymbolDraft(runtime, null)
    expect(captured?.draft.selectionStart).toBe(2)
    expect(captured?.draft.selectionEnd).toBe(7)
  })

  it('offers the saved formula when no editor is active and never uses its calculated display', () => {
    const { runtime, setValue } = runtimeWithFormula('=SUM(C1:C2)')
    const captured = captureFormulaSymbolDraft(runtime, null)
    expect(captured?.draft.source).toBe('=SUM(C1:C2)')
    expect(captured?.draft.selectionStart).toBe('=SUM(C1:C2)'.length)
    expect(setValue).not.toHaveBeenCalled()
  })

  it('does not discard a live plain-text edit on top of a formerly formula cell', () => {
    const { runtime } = runtimeWithFormula('=C1*2', { live: 'typed draft' })
    expect(captureFormulaSymbolDraft(runtime, null)).toBeNull()
  })

  it('commits to the captured sheet/cell as formula and refuses another workbook', () => {
    const { runtime, setValue } = runtimeWithFormula('="A"')
    const draft = captureFormulaSymbolDraft(runtime, null)!.draft
    expect(applyFormulaSymbolDraft(runtime, draft, '="AΩ"')).toBe(true)
    expect(setValue).toHaveBeenCalledWith({ f: '="AΩ"', v: null, p: null, si: null })
    expect(setValue).toHaveBeenCalledTimes(1)
    expect(applyFormulaSymbolDraft(runtime, { ...draft, workbookId: 'other' }, '="bad"')).toBe(
      false,
    )
    expect(applyFormulaSymbolDraft(runtime, draft, 'AΩ')).toBe(false)
    expect(setValue).toHaveBeenCalledTimes(1)
  })

  it('rejects an incomplete formula without writing, but ignores parentheses inside quoted text', () => {
    const { runtime, setValue } = runtimeWithFormula('=SUM(C1:C2)')
    const draft = captureFormulaSymbolDraft(runtime, null)!.draft
    expect(formulaSymbolSourceProblem('=SUM(')).toBe('parentheses')
    expect(formulaSymbolSourceProblem('=("(")')).toBeNull()
    expect(formulaSymbolSourceProblem('="escaped "" quote ("&C1')).toBeNull()
    expect(applyFormulaSymbolDraft(runtime, draft, '=SUM(')).toBe(false)
    expect(setValue).not.toHaveBeenCalled()
  })

  it('saves a source edit as an XLSX formula and preserves other source references', async () => {
    const formula = '="Ω "&SUM(C1:C2)'
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 2: { 1: { f: formula, v: null } } })
    const edit = toSaveEdits(journal)[0]!
    expect(edit.formula).toBe(formula)
    const source = await createBufferEntrySource(await buildEditFixture())
    const plan = await planCellEditsToXlsx(source, [
      {
        sheetName: 'Data',
        row: edit.row,
        column: edit.column,
        writeValue: true,
        cell: { value: edit.value, formula: edit.formula },
      },
    ])
    const xml = plan.replaced.get('xl/worksheets/sheet1.xml') ?? ''
    expect(xml).toContain('<c r="B3"><f>"Ω "&amp;SUM(C1:C2)</f></c>')
    expect(xml).not.toContain('<c r="B3"><f>"Ω "&amp;SUM(C1:C2)</f><v>5</v></c>')
  })
})
