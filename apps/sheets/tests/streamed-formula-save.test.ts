import { describe, expect, it, vi } from 'vitest'
import type { DesktopApi } from '../src/shared/desktop-api'
import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'
import { streamedFormulaValuesForSave } from '../src/renderer/streamed-formula-save'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

function fixture() {
  const journal = createEditJournal()
  recordSetRangeValues(journal, 's1', { 0: { 0: { v: 12345 } } })
  const state = {
    file: { sessionId: 'session', sheets: [{ id: 's1' }, { id: 's2' }] },
    formulaMode: false,
    closure: { status: 'unavailable' },
    recalc: { engineOverBudget: false },
    editJournal: journal,
  } as unknown as LazyWorkbookState
  const readWorkbookFormulas = vi.fn(async ({ sheetId }) => ({
    indexingComplete: true,
    truncated: false,
    cells: [0, 20_000, 49_999].map((row) => ({
      row,
      column: 15,
      formula: sheetId === 's1' ? `=A${row + 1}*2` : `=Large!P${row + 1}`,
    })),
  }))
  const recalcWorkbook = vi.fn(async ({ reads, edits }) => ({
    cells: reads.map(({ sheetId, range }: any) => ({
      sheetId,
      row: range.startRow,
      column: 15,
      isFormula: true,
      formatted: String(
        range.startRow === 0 ? Number(edits[0].input) * 2 : (range.startRow + 1) * 2,
      ),
      number: range.startRow === 0 ? Number(edits[0].input) * 2 : (range.startRow + 1) * 2,
    })),
  }))
  return {
    state,
    readWorkbookFormulas,
    recalcWorkbook,
    api: { readWorkbookFormulas, recalcWorkbook } as unknown as DesktopApi,
  }
}

describe('streamed formula values at explicit save', () => {
  it('collects offscreen and cross-sheet caches from bounded worker reads without changing formulas or journal', async () => {
    const f = fixture()
    const values = await streamedFormulaValuesForSave(f.state, f.api, () => true)
    expect(values).toHaveLength(6)
    expect(values).toContainEqual({ sheetId: 's1', row: 0, column: 15, value: 24690 })
    expect(values).toContainEqual({ sheetId: 's2', row: 49_999, column: 15, value: 100000 })
    for (const [{ reads }] of f.recalcWorkbook.mock.calls)
      for (const { range } of reads) {
        expect(
          (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1),
        ).toBeLessThanOrEqual(20_000)
      }
    expect(f.state.editJournal.cells.get('s1')?.size).toBe(1)
  })

  it('keeps existing calculation budgets and avoids worker work on formatting-only saves', async () => {
    const f = fixture()
    f.state.recalc.engineOverBudget = true
    expect(await streamedFormulaValuesForSave(f.state, f.api, () => true)).toEqual([])
    f.state.recalc.engineOverBudget = false
    f.state.editJournal.cells.clear()
    expect(await streamedFormulaValuesForSave(f.state, f.api, () => true)).toEqual([])
    expect(f.readWorkbookFormulas).not.toHaveBeenCalled()
  })

  it('does not return stale results after the workbook is replaced', async () => {
    const f = fixture()
    let current = true
    f.recalcWorkbook.mockImplementationOnce(async () => {
      current = false
      return { cells: [] }
    })
    await expect(streamedFormulaValuesForSave(f.state, f.api, () => current)).rejects.toThrow(
      'no longer current',
    )
  })

  it('preserves unsupported formula caches while accepting real calculation errors', async () => {
    const f = fixture()
    f.recalcWorkbook.mockImplementation(async ({ reads }) => ({
      cells: reads.map(({ sheetId, range }: any) => ({
        sheetId,
        row: range.startRow,
        column: 15,
        isFormula: true,
        isError: true,
        formatted: range.startRow === 0 ? '#NAME?' : '#DIV/0!',
      })),
    }))
    const values = await streamedFormulaValuesForSave(f.state, f.api, () => true)
    expect(values).toHaveLength(4)
    expect(values[0]?.value).toEqual({ error: '#DIV/0!' })
  })

  it('keeps cache-only saves available for files the fallback engine cannot import', async () => {
    const f = fixture()
    f.recalcWorkbook.mockRejectedValue(new Error('Formula engine import failed: missing styles'))
    expect(await streamedFormulaValuesForSave(f.state, f.api, () => true)).toEqual([])
  })

  it('does not wait for a truncated oversized formula index to finish', async () => {
    const f = fixture()
    f.readWorkbookFormulas.mockResolvedValue({
      indexingComplete: false,
      truncated: true,
      cells: [],
    })
    expect(await streamedFormulaValuesForSave(f.state, f.api, () => true)).toEqual([])
    expect(f.readWorkbookFormulas).toHaveBeenCalledTimes(1)
    expect(f.recalcWorkbook).not.toHaveBeenCalled()
  })
})
