/**
 * What Paste Special carries beside cell values and formats: merged cells, notes (Comments),
 * data validation and — for "All merging conditional formats" — the source's conditional formats.
 * Each works from the sheet the block was copied from, so it needs that sheet to still exist; the
 * caller runs them inside the paste's one undo step. Behavior was read from Excel 16 (COM):
 * a Comments paste replaces the destination's note with the source's and removes it where the
 * source has none; Validation does the same with rules; merges ride with All, Formats and All
 * except borders; "merging" keeps the destination's own conditional formats next to the source's.
 */
import type { IRange } from '@univerjs/core'
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { offsetFormulaRefs } from '@genoffice/xlsx-gateway/domain/formula-shift'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

import {
  mergedRangesAfterPaste,
  movedRange,
  tileOrigins,
  type PastedCell,
  type PasteSpecialOptions,
  type PasteTarget,
} from './paste-special'
import type { RibbonCommandContext } from './ribbon-actions'
import type { UniverRuntime, UniverWorksheet } from './univer-state'

/** the `row:column` key the maps below use */
const keyOf = (row: number, column: number): string => `${row}:${column}`

const contains = (range: IRange, row: number, column: number): boolean =>
  row >= range.startRow &&
  row <= range.endRow &&
  column >= range.startColumn &&
  column <= range.endColumn

const intersects = (a: IRange, b: IRange): boolean =>
  a.startRow <= b.endRow &&
  b.startRow <= a.endRow &&
  a.startColumn <= b.endColumn &&
  b.startColumn <= a.endColumn

/**
 * A rule's formulas anchor at the top-left cell of its first range, so a rule that lands elsewhere
 * has to have its relative references moved by the same distance (Excel: copy A1:A10 with
 * "=A1>5" to D1 and the rule reads "=D1>5"). Formula texts sit in `value` (conditional formats)
 * and `formula1` / `formula2` (validation).
 */
export function shiftRuleFormulas<T>(rule: T, rowDelta: number, columnDelta: number): T {
  if (rowDelta === 0 && columnDelta === 0) return rule
  const walk = (node: unknown): unknown => {
    if (Array.isArray(node)) return node.map(walk)
    if (node !== null && typeof node === 'object') {
      const out: Record<string, unknown> = {}
      for (const [key, value] of Object.entries(node)) {
        out[key] =
          (key === 'value' || key === 'formula1' || key === 'formula2') &&
          typeof value === 'string' &&
          value.startsWith('=')
            ? offsetFormulaRefs(value, rowDelta, columnDelta)
            : walk(value)
      }
      return out
    }
    return node
  }
  return walk(rule) as T
}

/** cells as ranges: consecutive cells of a row become one range */
export function runsOf(cells: readonly { row: number; column: number }[]): IRange[] {
  const byRow = new Map<number, number[]>()
  for (const cell of cells) {
    const columns = byRow.get(cell.row) ?? []
    columns.push(cell.column)
    byRow.set(cell.row, columns)
  }
  const runs: IRange[] = []
  for (const [row, columns] of [...byRow].sort((a, b) => a[0] - b[0])) {
    columns.sort((a, b) => a - b)
    let start = columns[0]!
    let previous = start
    for (const column of columns.slice(1)) {
      if (column === previous + 1) {
        previous = column
        continue
      }
      runs.push({ startRow: row, endRow: row, startColumn: start, endColumn: previous })
      start = column
      previous = column
    }
    runs.push({ startRow: row, endRow: row, startColumn: start, endColumn: previous })
  }
  return runs
}

const rangeOfFacade = (range: {
  getRow(): number
  getColumn(): number
  getHeight(): number
  getWidth(): number
}): IRange => ({
  startRow: range.getRow(),
  startColumn: range.getColumn(),
  endRow: range.getRow() + range.getHeight() - 1,
  endColumn: range.getColumn() + range.getWidth() - 1,
})

let serial = 0
const freshId = (prefix: string): string => {
  serial += 1
  return `${prefix}-${Date.now().toString(36)}-${serial}`
}

/** Merged cells of the destination that the paste's area touches are unmerged first (the pasted
 * formats decide what is merged there), then the source's merges are made at every tile. */
export async function pasteMerges(
  sheet: UniverWorksheet,
  sourceSheet: UniverWorksheet,
  source: { readonly rows: readonly number[]; readonly columns: readonly number[] },
  options: PasteSpecialOptions,
  target: PasteTarget,
  extent: IRange,
): Promise<void> {
  const wanted = mergedRangesAfterPaste(
    sourceSheet.getMergedRanges().map(rangeOfFacade),
    source,
    options,
    target,
  )
  for (const merged of sheet.getMergedRanges()) {
    if (intersects(rangeOfFacade(merged), extent)) await merged.breakApart()
  }
  for (const range of wanted) {
    await sheet
      .getRange(
        range.startRow,
        range.startColumn,
        range.endRow - range.startRow + 1,
        range.endColumn - range.startColumn + 1,
      )
      .merge()
  }
}

/** Comments: each pasted cell gets the source cell's note; where the source has none the note goes. */
export async function pasteNotes(
  ctx: RibbonCommandContext,
  sheet: UniverWorksheet,
  sourceSheet: UniverWorksheet,
  cells: readonly PastedCell[],
): Promise<void> {
  const sourceNotes = new Map<string, string>()
  for (const note of sourceSheet.getNotes()) sourceNotes.set(keyOf(note.row, note.col), note.note)
  const destinationNotes = new Map<string, string>()
  for (const note of sheet.getNotes()) destinationNotes.set(keyOf(note.row, note.col), note.note)
  const sheetId = sheet.getSheetId()
  const ops: WorkbookOperation[] = []
  for (const cell of cells) {
    const wanted = sourceNotes.get(keyOf(cell.sourceRow, cell.sourceColumn)) ?? null
    const have = destinationNotes.get(keyOf(cell.row, cell.column)) ?? null
    if (wanted === have) continue
    ops.push({
      op: 'set_note',
      sheetId,
      address: formatAddress(cell.row, cell.column),
      text: wanted,
    })
  }
  if (ops.length > 0) await ctx.runOps(ops, null)
}

/** Validation: each pasted cell gets the source cell's rule (one new rule per source rule, over all
 * the cells it lands on); a cell whose source has no rule loses the one it had. */
export async function pasteValidation(
  runtime: UniverRuntime,
  sheet: UniverWorksheet,
  sourceSheet: UniverWorksheet,
  cells: readonly PastedCell[],
): Promise<void> {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  const subUnitId = sheet.getSheetId()
  const sourceRules = sourceSheet.getDataValidations().map((validation) => validation.rule)
  const destinationRules = sheet.getDataValidations().map((validation) => validation.rule)
  const groups = new Map<string, { rule: (typeof sourceRules)[number]; cells: PastedCell[] }>()
  for (const cell of cells) {
    const rule = sourceRules.find((candidate) =>
      (candidate.ranges ?? []).some((range: IRange) =>
        contains(range, cell.sourceRow, cell.sourceColumn),
      ),
    )
    if (!rule) continue
    const group = groups.get(rule.uid) ?? { rule, cells: [] }
    group.cells.push(cell)
    groups.set(rule.uid, group)
  }
  // clear only what could hold a rule: the destination's own rules inside the pasted area
  const hadRule = (cell: PastedCell): boolean =>
    destinationRules.some((rule) =>
      (rule.ranges ?? []).some((range: IRange) => contains(range, cell.row, cell.column)),
    )
  const toClear = cells.filter(hadRule)
  if (toClear.length > 0) {
    await runtime.univerAPI.executeCommand('sheets.command.clear-range-data-validation', {
      unitId,
      subUnitId,
      ranges: runsOf(toClear),
    })
  }
  for (const { rule, cells: landed } of groups.values()) {
    const ranges = runsOf(landed)
    const from = rule.ranges?.[0]
    const moved =
      from && ranges[0]
        ? shiftRuleFormulas(
            rule,
            ranges[0].startRow - from.startRow,
            ranges[0].startColumn - from.startColumn,
          )
        : rule
    await runtime.univerAPI.executeCommand('sheet.command.addDataValidation', {
      unitId,
      subUnitId,
      rule: { ...moved, uid: freshId('dv'), ranges },
    })
  }
}

/** All merging conditional formats: the source's rules are added over the pasted cells and the
 * destination's own rules stay. Returns how many rules were added. */
export function mergeConditionalFormats(
  sheet: UniverWorksheet,
  sourceSheet: UniverWorksheet,
  source: { readonly rows: readonly number[]; readonly columns: readonly number[] },
  options: PasteSpecialOptions,
  target: PasteTarget,
): number {
  const block = {
    startRow: source.rows[0]!,
    endRow: source.rows[source.rows.length - 1]!,
    startColumn: source.columns[0]!,
    endColumn: source.columns[source.columns.length - 1]!,
  }
  const origins = tileOrigins(source, options, target)
  let added = 0
  for (const rule of sourceSheet.getConditionalFormattingRules()) {
    // the part of each of the rule's ranges that lies in the copied block, moved to every tile
    const ranges: IRange[] = []
    for (const range of rule.ranges) {
      if (!intersects(range, block)) continue
      const clipped = {
        startRow: Math.max(range.startRow, block.startRow),
        endRow: Math.min(range.endRow, block.endRow),
        startColumn: Math.max(range.startColumn, block.startColumn),
        endColumn: Math.min(range.endColumn, block.endColumn),
      }
      for (const origin of origins) {
        const moved = movedRange(clipped, source, options, origin)
        if (moved) ranges.push(moved)
      }
    }
    if (ranges.length === 0) continue
    const from = rule.ranges[0]
    const shifted =
      from && ranges[0]
        ? shiftRuleFormulas(
            rule,
            ranges[0].startRow - from.startRow,
            ranges[0].startColumn - from.startColumn,
          )
        : rule
    sheet.addConditionalFormattingRule({ ...shifted, cfId: freshId('cf'), ranges })
    added += 1
  }
  return added
}
