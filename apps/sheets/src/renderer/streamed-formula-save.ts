import type { DesktopApi, WorkbookRecalcRequest } from '../shared/desktop-api'
import { isSheetRemoved, toRecalcUserInput } from './edit-journal'
import { cellKey, closureFetchRanges } from './formula-closure'
import type { CachedFormulaValue } from './formula-values'
import type { LazyWorkbookState } from './univer-state'
import { engineResultKeepsCache } from './univer-sync'

type FormulaApi = Pick<DesktopApi, 'readWorkbookFormulas' | 'recalcWorkbook'>

/** An explicit save needs all formula caches, beyond the fallback's viewport-sized overlay.
 * Evaluate in the existing sidecar worker and read bounded batches; never preload the grid.
 * Keep the existing size/structural capability gates and manual-calculation policy at the caller.
 */
export async function streamedFormulaValuesForSave(
  state: LazyWorkbookState,
  api: FormulaApi,
  isCurrent: () => boolean,
): Promise<CachedFormulaValue[]> {
  if (
    state.formulaMode ||
    state.closure?.status !== 'unavailable' ||
    !state.recalc ||
    state.recalc.engineOverBudget
  )
    return []
  if (
    state.editJournal.sheets.added.size > 0 ||
    [...state.editJournal.structuralOps.values()].some((ops) => ops.length > 0)
  )
    return []
  const edits: WorkbookRecalcRequest['edits'] = []
  for (const [sheetId, cells] of state.editJournal.cells) {
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    for (const entry of cells.values()) {
      if (!entry.hasValue && !entry.formula) continue
      edits.push({ sheetId, row: entry.row, column: entry.column, input: toRecalcUserInput(entry) })
    }
  }
  if (edits.length === 0 || edits.length > 10_000) return []
  const deadline = Date.now() + 30_000
  const check = () => {
    if (!isCurrent() || Date.now() > deadline)
      throw new Error('Formula save snapshot is no longer current')
  }
  const wait = async () => {
    await new Promise((resolve) => setTimeout(resolve, 100))
    check()
  }
  const values: CachedFormulaValue[] = []
  for (const sheet of state.file.sheets) {
    if (isSheetRemoved(state.editJournal, sheet.id)) continue
    check()
    let index = await api.readWorkbookFormulas({
      sessionId: state.file.sessionId,
      sheetId: sheet.id,
    })
    while (!index.indexingComplete && !index.truncated) {
      await wait()
      index = await api.readWorkbookFormulas({ sessionId: state.file.sessionId, sheetId: sheet.id })
    }
    // Match the existing fallback's incomplete-index gate; never cold-import an unbounded file.
    if (index.truncated) return []
    const formulas = new Map(
      index.cells
        .filter((cell) => cell.formula)
        .map((cell) => [`${cell.row}:${cell.column}`, cell.formula!]),
    )
    for (const entry of state.editJournal.cells.get(sheet.id)?.values() ?? []) {
      if (entry.formula) formulas.set(`${entry.row}:${entry.column}`, entry.formula)
      else if (entry.hasValue) formulas.delete(`${entry.row}:${entry.column}`)
    }
    const keys = new Set(
      [...formulas.keys()].map((key) => {
        const [row, column] = key.split(':').map(Number)
        return cellKey(row!, column!)
      }),
    )
    for (const range of closureFetchRanges(keys)) {
      check()
      let result
      for (;;) {
        try {
          result = await api.recalcWorkbook({
            sessionId: state.file.sessionId,
            edits,
            reads: [{ sheetId: sheet.id, range }],
          })
          break
        } catch (error) {
          // Some non-Excel producers cannot be imported at all. Preserve the existing
          // cache-only save behavior for that capability gap, like unsupported formulas.
          if (
            error instanceof Error &&
            /Formula engine import failed|formula engine could not process this workbook/.test(
              error.message,
            )
          )
            return []
          if (
            !(error instanceof Error) ||
            !error.message.includes('busy with another recalculation')
          )
            throw error
          await wait()
        }
      }
      check()
      for (const cell of result.cells) {
        if (cell.sheetId !== sheet.id || !cell.isFormula) continue
        const formula = formulas.get(`${cell.row}:${cell.column}`)
        if (!formula || engineResultKeepsCache(cell.formatted, formula)) continue
        values.push({
          sheetId: sheet.id,
          row: cell.row,
          column: cell.column,
          value: cell.isError ? { error: cell.formatted } : (cell.number ?? cell.formatted),
        })
      }
    }
  }
  return values
}
