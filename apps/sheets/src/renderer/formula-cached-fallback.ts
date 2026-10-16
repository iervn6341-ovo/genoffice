/**
 * Display fallback to the file's cached formula results.
 *
 * When the engine's recalculation of a file formula errors — a function it
 * lacks or a name it can't resolve — Excel's cached result from the file is
 * the better display. Display-only: the model keeps the engine's result and
 * the chain continues, so number formats still apply. A cell the user
 * overwrote this session shows the live engine truth instead.
 */
import { CellValueType } from '@univerjs/core'
import { ERROR_TYPE_SET, ErrorType } from '@univerjs/engine-formula'
import { INTERCEPTOR_POINT, SheetInterceptorService } from '@univerjs/sheets'
import { qualifierMatches } from '@genoffice/xlsx-gateway/gateway/xlsx-structure'

import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import { parseFormulaReferences } from './formula-closure'

/// Journal ops that renumber rows/columns; the recorded cache coordinates are
/// stale after these. Layout-only ops (sizing, hiding, outlining, merging)
/// leave addresses intact.
const COORDINATE_SHIFTING_OPS = new Set([
  'insert-rows',
  'remove-rows',
  'insert-cols',
  'remove-cols',
  'move-rows',
])

// A formula can reference another formula which in turn reads an edited
// sheet. Bound that walk so an unusual whole-column chain cannot stall paint.
const MAX_FALLBACK_DEPENDENCY_VISITS = 1000
const MAX_FALLBACK_RANGE_SCAN = 10_000

function readsEditedPrecedent(
  state: LazyWorkbookState,
  hostSheetId: string,
  formula: string,
): boolean {
  if (state.editJournal.cells.size === 0) return false
  let anyContentEdit = false
  for (const sheetEdits of state.editJournal.cells.values()) {
    for (const edit of sheetEdits.values()) {
      if (edit.hasValue || edit.formula !== undefined) {
        anyContentEdit = true
        break
      }
    }
    if (anyContentEdit) break
  }
  if (!anyContentEdit) return false
  const seen = new Set<string>()
  let visits = 0
  let scanned = 0
  const pending = [{ sheetId: hostSheetId, source: formula }]
  while (pending.length > 0) {
    if (++visits > MAX_FALLBACK_DEPENDENCY_VISITS) return true
    const current = pending.pop()!
    const sheetId = current.sheetId
    const source = current.source
    for (const reference of parseFormulaReferences(source)) {
      const qualifier = reference.qualifier
      const unquoted =
        qualifier?.startsWith("'") && qualifier.endsWith("'")
          ? qualifier.slice(1, -1).replaceAll("''", "'")
          : qualifier
      const targetId =
        qualifier === undefined
          ? sheetId
          : (
              state.file.sheets.find((sheet) => qualifierMatches(qualifier, sheet.name)) ??
              state.file.sheets.find(
                (sheet) => sheet.name.toLowerCase() === unquoted?.toLowerCase(),
              )
            )?.id
      if (!targetId) continue
      const { token } = reference
      const edits = state.editJournal.cells.get(targetId)
      const formulas = state.formulaText.get(targetId)
      const pointKey =
        token.startRow !== null &&
        token.startRow === token.endRow &&
        token.startColumn !== null &&
        token.startColumn === token.endColumn
          ? `${token.startRow}:${token.startColumn}`
          : null
      const queueFormula = (key: string, nested: string | undefined): void => {
        if (!nested) return
        const identity = `${targetId}:${key}`
        if (seen.has(identity)) return
        seen.add(identity)
        pending.push({ sheetId: targetId, source: nested })
      }
      if (pointKey !== null) {
        const edit = edits?.get(pointKey)
        if (edit?.hasValue || edit?.formula !== undefined) return true
        queueFormula(pointKey, formulas?.get(pointKey))
        continue
      }
      const contains = (key: string): boolean => {
        const [row, column] = key.split(':').map(Number)
        return (
          row !== undefined &&
          column !== undefined &&
          (token.startRow === null || row >= token.startRow) &&
          (token.endRow === null || row <= token.endRow) &&
          (token.startColumn === null || column >= token.startColumn) &&
          (token.endColumn === null || column <= token.endColumn)
        )
      }
      for (const [key, edit] of edits ?? []) {
        if (++scanned > MAX_FALLBACK_RANGE_SCAN) return true
        if ((edit.hasValue || edit.formula !== undefined) && contains(key)) return true
      }
      for (const [key, nested] of formulas ?? []) {
        if (++scanned > MAX_FALLBACK_RANGE_SCAN) return true
        if (!contains(key)) continue
        queueFormula(key, nested)
      }
    }
  }
  return false
}

/**
 * #VALUE! from a formula built only of literals, references and operators is
 * the engine's own arithmetic verdict (text the coercion rules reject, like
 * Excel recalculating "76,92"+…), not a capability gap — the file cache must
 * not mask it. Functions, defined names and structured/external refs keep
 * the fallback: those are where the engine errors where Excel would not.
 */
export function isPlainArithmeticFormula(formula: string): boolean {
  let body = formula.startsWith('=') ? formula.slice(1) : formula
  body = body.replace(/"(?:[^"]|"")*"/g, '""').replace(/\$/g, '')
  // Structured/external refs, array literals and spill refs are not "plain".
  if (/[[{#]/.test(body)) return false
  body = body.replace(/'(?:[^']|'')*'!/g, '').replace(/[A-Za-z_][\w.]*!/g, '')
  // Scientific literals (1E-3) would otherwise shed a bare E token.
  body = body.replace(/(?<![\w.])\d+(?:\.\d+)?E[+-]?\d+/gi, '0')
  for (const match of body.matchAll(/[A-Za-z_][\w.]*/g)) {
    const token = match[0]
    // A ref-shaped name followed by "(" is a function call (LOG10, ...).
    if (/^\s*\(/.test(body.slice((match.index ?? 0) + token.length))) return false
    if (!/^[A-Za-z]{1,3}\d+$/.test(token) && !/^(?:TRUE|FALSE)$/i.test(token)) return false
  }
  return true
}

export function installCachedValueFallbackInterceptor(
  runtime: UniverRuntime,
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null },
): { dispose(): void } {
  const interceptorService = runtime.univer.__getInjector().get(SheetInterceptorService)
  return interceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
    // Below the formula-view interceptors (formula view must keep showing
    // the formula text), above NUMFMT (10).
    priority: 9997,
    handler: (cell, location, next) => {
      const value = cell?.v
      if (typeof value !== 'string' || !ERROR_TYPE_SET.has(value as ErrorType)) {
        return next(cell)
      }
      if (value === ErrorType.VALUE) {
        const formula = location.rawData?.f
        if (typeof formula === 'string' && isPlainArithmeticFormula(formula)) return next(cell)
      }
      const state = lazyWorkbookRef.current
      if (!state) return next(cell)
      const sheetId = location.subUnitId
      const ops = state.editJournal.structuralOps.get(sheetId)
      // Row/column shifts invalidate the recorded coordinates; user
      // overwrites own the cell.
      if (ops?.some((op) => COORDINATE_SHIFTING_OPS.has(op.kind))) return next(cell)
      const key = `${location.row}:${location.col}`
      const edits = state.editJournal.cells.get(sheetId)
      // Style-only journal entries (hasValue false) don't change any input.
      if (edits?.get(key)?.hasValue) return next(cell)
      // An edited precedent on this or another sheet can make a computed
      // error (#DIV/0!, #N/A, ...) the true result. Unrelated edits must not
      // expose an engine capability error previously hidden by Excel's cache.
      // #NAME? still falls back: the engine cannot evaluate that formula
      // under any inputs, so the file's value remains the best display.
      if (value !== ErrorType.NAME) {
        for (const edit of edits?.values() ?? []) {
          if (edit.hasValue || edit.formula !== undefined) return next(cell)
        }
        const formula = location.rawData?.f ?? state.formulaText.get(sheetId)?.get(key)
        if (formula && readsEditedPrecedent(state, sheetId, formula)) return next(cell)
      }
      const cached = state.cachedFormulaValues.get(sheetId)?.get(key)
      if (cached === undefined || cached === value) return next(cell)
      return next({
        ...cell,
        v: cached,
        t:
          typeof cached === 'number'
            ? CellValueType.NUMBER
            : typeof cached === 'boolean'
              ? CellValueType.BOOLEAN
              : CellValueType.STRING,
      })
    },
  })
}
