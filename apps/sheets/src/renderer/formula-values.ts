/**
 * Read back the values of formulas a batch just wrote.
 *
 * A formula's *text* lands synchronously, but the engine computes its result
 * asynchronously. An external caller that reads immediately therefore sees
 * `value: null` for a formula it just wrote correctly, and `apply_sheet_ops`
 * returning "ok" while `read_sheet` shows nothing is indistinguishable from a
 * silent failure — an agent can conclude the edit did not take.
 *
 * In-app callers never meet this: the user watches the cell fill in. So the
 * MCP path closes the gap here, polling the addressed cells until the engine has
 * written their values and reporting them alongside the apply result. Polling
 * (rather than waiting on the engine's calculationEnd event) keeps it precise:
 * a batch of plain values triggers no recalculation at all, and must not hold
 * the response for a timeout that will never be met.
 */
import type { CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'
import { EXCEL_ERROR_LITERALS, modelCellValue } from './univer-sync'

/** one `set_formula` target extracted from a batch */
export interface FormulaTarget {
  readonly sheetId: string
  readonly address: string
}

/** a resolved formula cell, as reported back to the caller */
export interface FormulaValue {
  readonly sheetId: string
  readonly address: string
  readonly value: unknown
  readonly formula?: string
}

/**
 * The cells a batch is expected to leave holding a computed value: the
 * `set_formula` ops. Plain `set_cell` writes are readable immediately and need
 * no wait; conditional-format and data-validation formulas are rules, not cell
 * values, and are skipped.
 */
export function formulaTargetsFromOps(ops: readonly unknown[]): FormulaTarget[] {
  const targets: FormulaTarget[] = []
  const seen = new Set<string>()
  for (const raw of ops) {
    if (!raw || typeof raw !== 'object') continue
    const op = raw as { op?: unknown; sheetId?: unknown; address?: unknown; formula?: unknown }
    if (op.op !== 'set_formula') continue
    if (typeof op.sheetId !== 'string' || typeof op.address !== 'string') continue
    if (typeof op.formula !== 'string' || !op.formula.startsWith('=')) continue
    const key = `${op.sheetId}!${op.address}`
    if (seen.has(key)) continue
    seen.add(key)
    targets.push({ sheetId: op.sheetId, address: op.address })
  }
  return targets
}

/**
 * How long to keep polling. A formula over a few hundred cells settles in
 * milliseconds; the budget exists for a large workbook's first full pass. On
 * expiry the targets that never resolved are simply omitted — a partially
 * computed batch is still a successful write.
 */
const VALUE_TIMEOUT_MS = 4_000
const VALUE_POLL_MS = 25

/**
 * Poll `read` until every target reports a value (or the budget runs out), then
 * return what resolved. `read` takes one sheet's addresses because the readers
 * are per-worksheet; targets are grouped accordingly.
 */
export async function awaitFormulaValues(
  targets: readonly FormulaTarget[],
  read: (addresses: readonly string[], sheetId: string) => Record<string, CellState>,
): Promise<FormulaValue[]> {
  if (targets.length === 0) return []
  const bySheet = new Map<string, string[]>()
  for (const t of targets) {
    const list = bySheet.get(t.sheetId)
    if (list) list.push(t.address)
    else bySheet.set(t.sheetId, [t.address])
  }

  const resolved = new Map<string, FormulaValue>()
  const deadline = Date.now() + VALUE_TIMEOUT_MS
  for (;;) {
    for (const [sheetId, addresses] of bySheet) {
      let cells: Record<string, CellState>
      try {
        cells = read(addresses, sheetId)
      } catch {
        // a sheet removed by this very batch is not an error to report here
        continue
      }
      for (const address of addresses) {
        const key = `${sheetId}!${address}`
        if (resolved.has(key)) continue
        const cell = cells[address]
        // null means the engine has not written a result yet; '' or an error
        // string are real outcomes and count as settled.
        if (!cell || cell.value === null) continue
        const value = modelCellValue(cell)
        if (value === null) continue
        resolved.set(key, {
          sheetId,
          address,
          value,
          ...(cell.formula === undefined ? {} : { formula: cell.formula }),
        })
      }
    }
    if (resolved.size === targets.length || Date.now() >= deadline) break
    await new Promise((resolve) => setTimeout(resolve, VALUE_POLL_MS))
  }
  return [...resolved.values()]
}

/**
 * Formula cells this session has seen the engine settle on, keyed
 * `sheetId!address`, with the formula text that was settled.
 *
 * The save path refuses to cache a formula it has not seen a value for (the
 * overlay it used to read deliberately skips journaled cells, and trusting it
 * could cache the *previous* formula's result right after a replacement). That
 * protection left every MCP-written formula without a cached `<v>` on disk:
 * Excel recalculated on open so nothing looked broken, but read-only consumers
 * — the CLI, a preview, a diff — showed those cells as empty.
 *
 * Only the *address* is trusted from here. The value written at save time is
 * read live from the grid, so an edit to a precedent cell after the batch (by
 * the user or a later batch) cannot leave the file carrying the earlier result.
 */
export interface VerifiedFormulaValue {
  readonly formula: string
}

/**
 * One renderer process hosts exactly one open workbook session, so the verified
 * cells live in a module-level map rather than keyed by the workbook facade:
 * `univerAPI.getActiveWorkbook()` hands back a fresh facade on every call, and a
 * WeakMap keyed on it would never match the instance the save path asks with.
 * `clearVerifiedFormulaValues` runs when a different workbook is loaded.
 */
const verified = new Map<string, VerifiedFormulaValue>()

/** forget every verified cell — a save replaced the session, so they are stale */
export function clearVerifiedFormulaValues(): void {
  verified.clear()
}

/** remember that a formula was observed to settle on a value */
export function rememberFormulaValue(
  sheetId: string,
  address: string,
  entry: VerifiedFormulaValue,
): void {
  verified.set(`${sheetId}!${address}`, entry)
}

/** the cached-value shape the save pipeline accepts for one formula cell */
export interface CachedFormulaValue {
  readonly sheetId: string
  readonly row: number
  readonly column: number
  readonly value: string | number | boolean | { error: string } | null
}

/**
 * Live values of the verified cells, in the shape the save path's
 * `formulaValues` takes. A cell whose formula text no longer matches what was
 * settled, or whose result the engine has not written yet, is left out rather
 * than cached against the wrong result; a sheet the reader cannot resolve
 * (removed since) is skipped.
 */
export function verifiedFormulaValues(
  read: (addresses: string[], sheetId: string) => Record<string, CellState>,
): CachedFormulaValue[] {
  if (verified.size === 0) return []
  const bySheet = new Map<string, Map<string, VerifiedFormulaValue>>()
  for (const [key, entry] of verified) {
    const at = key.lastIndexOf('!')
    if (at <= 0) continue
    const sheetId = key.slice(0, at)
    const cells = bySheet.get(sheetId) ?? new Map<string, VerifiedFormulaValue>()
    cells.set(key.slice(at + 1), entry)
    bySheet.set(sheetId, cells)
  }
  const out: CachedFormulaValue[] = []
  for (const [sheetId, entries] of bySheet) {
    let cells: Record<string, CellState>
    try {
      cells = read([...entries.keys()], sheetId)
    } catch {
      continue
    }
    for (const [address, entry] of entries) {
      const parsed = parseAddressParts(address)
      const cell = cells[address]
      if (!parsed || !cell || cell.formula !== entry.formula) continue
      // The model value, not the rendered text: General shrinks a number to fit
      // the column (numfmt-fix.ts), so saving the display would cache "0.333333"
      // as a string where Excel stores a full-precision number.
      const live = modelCellValue(cell)
      // #ERROR! is IronCalc's own failure, never a value Excel would cache
      if (live === null || live === '#ERROR!') continue
      out.push({
        sheetId,
        row: parsed.row,
        column: parsed.column,
        value: isErrorResult(live) ? { error: live } : live,
      })
    }
  }
  return out
}

/** `A1` -> zero-based row/column, matching the save path's coordinates */
function parseAddressParts(address: string): { row: number; column: number } | null {
  const match = /^([A-Z]+)(\d+)$/.exec(address.toUpperCase())
  if (!match) return null
  const letters = match[1]!
  const digits = match[2]!
  let column = 0
  for (const ch of letters) column = column * 26 + (ch.charCodeAt(0) - 64)
  return { row: Number(digits) - 1, column: column - 1 }
}

/** true when the value a formula produced is an Excel error, not a literal */
export function isErrorResult(value: unknown): value is string {
  return typeof value === 'string' && EXCEL_ERROR_LITERALS.has(value)
}

/** the journal slice {@link journaledFormulaValues} reads */
export interface JournaledFormulaCell {
  readonly row: number
  readonly column: number
  readonly formula?: string | undefined
}

/**
 * Cells per save whose results are read back for caching. Above this the file
 * still saves (its <f> is intact and fullCalcOnLoad makes Excel recompute); only
 * the cached values of the overflow are left out.
 */
export const JOURNALED_VALUE_LIMIT = 50_000

/**
 * Live results of formulas the USER typed this session (journaled formula
 * cells). The overlay path skips them because it may hold the previous
 * formula's result, and only MCP batches were ever verified, so a formula
 * entered in the grid was saved as a bare <f> — Quick Look, pandas and mobile
 * viewers showed an empty cell until Excel recalculated. The caller must have
 * awaited the engine settling first; a cell whose live formula differs from
 * the journaled one, or has no result yet, is left out rather than cached
 * against the wrong result.
 */
export function journaledFormulaValues(
  journal: ReadonlyMap<string, ReadonlyMap<string, JournaledFormulaCell>>,
  read: (addresses: string[], sheetId: string) => Record<string, CellState>,
  skipSheet: (sheetId: string) => boolean = () => false,
  limit = JOURNALED_VALUE_LIMIT,
): CachedFormulaValue[] {
  const out: CachedFormulaValue[] = []
  let budget = limit
  for (const [sheetId, cells] of journal) {
    if (budget <= 0) break
    if (skipSheet(sheetId)) continue
    const wanted = new Map<string, JournaledFormulaCell>()
    for (const cell of cells.values()) {
      if (cell.formula === undefined || budget <= 0) continue
      wanted.set(a1Address(cell.row, cell.column), cell)
      budget -= 1
    }
    if (wanted.size === 0) continue
    let live: Record<string, CellState>
    try {
      live = read([...wanted.keys()], sheetId)
    } catch {
      continue
    }
    for (const [address, entry] of wanted) {
      const cell = live[address]
      if (!cell || cell.formula === undefined) continue
      if (normalizeFormula(cell.formula) !== normalizeFormula(entry.formula!)) continue
      const value = modelCellValue(cell)
      if (value === null || value === '#ERROR!') continue
      out.push({
        sheetId,
        row: entry.row,
        column: entry.column,
        value: isErrorResult(value) ? { error: value } : value,
      })
    }
  }
  return out
}

function normalizeFormula(formula: string): string {
  return formula.trim().replace(/^=/, '')
}

/** zero-based row/column → `A1` */
function a1Address(row: number, column: number): string {
  let letters = ''
  for (let n = column + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters
  }
  return `${letters}${row + 1}`
}
