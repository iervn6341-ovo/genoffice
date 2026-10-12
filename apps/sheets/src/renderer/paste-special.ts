/// Paste Special beyond what Univer's own paste variants do: Transpose, Operation (add / subtract /
/// multiply / divide), Skip blanks, and the two number-format variants. Pure planning from a
/// copied block to a list of cell writes; the action layer turns the plan into Univer mutations.
/// Behavior was read from Excel 16 through COM (`Range.PasteSpecial`):
///   - an operation combines the destination with the source (dest − source, dest ÷ source …);
///     a blank destination counts as 0; a text or blank source leaves the destination alone; a
///     formula destination is wrapped: `=(2+3)` + 1 → `=(2+3)+1`; a ÷ 0 shows #DIV/0!;
///   - Skip blanks leaves the destination under a blank source untouched, otherwise a blank
///     source clears the destination;
///   - Transpose pastes the block once at the top-left of the selection (it is never tiled) and
///     refuses a block that overlaps its source unless it has the same size and shape;
///   - Values & Number Formats / Formulas & Number Formats carry the number format only;
///   - a source smaller than the selection tiles when the selection is a whole multiple of it;
///   - merged cells travel with the choices that carry formats (All, Formats, All except borders)
///     and turn with Transpose; Values, Formulas and the number-format variants leave merges alone;
///   - Comments and Validation replace the destination's note / rule with the source's (a source
///     without one clears the destination's); All and All except borders carry both;
///   - after Cut, Paste Special is refused (Range.PasteSpecial throws).
import type { IRange, IStyleData } from '@univerjs/core'

import { transposeFormula } from './formula-transpose'

export type PasteWhat =
  | 'all'
  | 'formulas'
  | 'values'
  | 'formats'
  | 'allExceptBorders'
  | 'columnWidths'
  | 'formulasAndNumberFormats'
  | 'valuesAndNumberFormats'
  | 'comments'
  | 'validation'
  /// within one workbook there is one theme, so this is All
  | 'allUsingSourceTheme'
  | 'allMergingConditionalFormats'

export type PasteOperation = 'none' | 'add' | 'subtract' | 'multiply' | 'divide'

export interface PasteSpecialOptions {
  readonly what: PasteWhat
  readonly operation: PasteOperation
  readonly skipBlanks: boolean
  readonly transpose: boolean
}

export const DEFAULT_PASTE_SPECIAL: PasteSpecialOptions = {
  what: 'all',
  operation: 'none',
  skipBlanks: false,
  transpose: false,
}

/// Univer's own paste commands, for the choices it already implements exactly
export const NATIVE_PASTE_COMMANDS: Partial<Record<PasteWhat, string>> = {
  all: 'univer.command.paste',
  formulas: 'sheet.command.paste-formula',
  values: 'sheet.command.paste-value',
  formats: 'sheet.command.paste-format',
  allExceptBorders: 'sheet.command.paste-besides-border',
  columnWidths: 'sheet.command.paste-col-width',
  allUsingSourceTheme: 'univer.command.paste',
}

/// what a choice carries besides cell values and formats
export interface PasteExtras {
  readonly notes: boolean
  readonly validation: boolean
  readonly merges: boolean
}

export function pasteExtras(what: PasteWhat): PasteExtras {
  switch (what) {
    case 'all':
    case 'allExceptBorders':
    case 'allUsingSourceTheme':
    case 'allMergingConditionalFormats':
      return { notes: true, validation: true, merges: true }
    case 'formats':
      return { notes: false, validation: false, merges: true }
    case 'comments':
      return { notes: true, validation: false, merges: false }
    case 'validation':
      return { notes: false, validation: true, merges: false }
    default:
      return { notes: false, validation: false, merges: false }
  }
}

/// true when Univer's paste variants cannot do it and the plan below has to
export function needsOwnPaste(options: PasteSpecialOptions): boolean {
  return (
    options.operation !== 'none' ||
    options.skipBlanks ||
    options.transpose ||
    options.what === 'formulasAndNumberFormats' ||
    options.what === 'valuesAndNumberFormats' ||
    options.what === 'comments' ||
    options.what === 'validation' ||
    options.what === 'allMergingConditionalFormats'
  )
}

export interface CopiedSourceCell {
  readonly v?: string | number | boolean | null
  readonly t?: number | null
  readonly f?: string | null
  readonly style: IStyleData | null
}

export interface CopiedSource {
  readonly rows: readonly number[]
  readonly columns: readonly number[]
  readonly cells: readonly (readonly (CopiedSourceCell | null)[])[]
}

export interface DestinationCell {
  readonly v?: string | number | boolean | null
  readonly f?: string | null
}

export interface PlannedWrite {
  readonly row: number
  readonly column: number
  /// undefined = the cell's value stays as it is
  readonly value?: {
    readonly v: string | number | boolean | null
    readonly t: number | null
    readonly f: string | null
  }
  /// undefined = the cell's format stays as it is
  readonly style?:
    | { readonly mode: 'replace'; readonly style: IStyleData | null }
    | { readonly mode: 'numfmt'; readonly pattern: string | null }
}

export type PastePlan =
  | { readonly writes: readonly PlannedWrite[]; readonly extent: IRange }
  | { readonly error: 'overlap' | 'empty' }

export interface PasteTarget {
  readonly anchorRow: number
  readonly anchorColumn: number
  /// size of the destination selection (a single cell is 1 × 1)
  readonly selectionRows: number
  readonly selectionColumns: number
}

export interface PasteInput {
  readonly source: CopiedSource
  readonly options: PasteSpecialOptions
  readonly target: PasteTarget
  /// the destination is on the sheet the block was copied from
  readonly sameSheet: boolean
  readonly readDestination: (row: number, column: number) => DestinationCell
  readonly readDestinationStyle: (row: number, column: number) => IStyleData | null
  /// the name of the sheet the block was copied from (for references to other sheets)
  readonly sourceSheetName?: string
  /// moves the relative references of a formula by whole columns / rows (Univer's lexer)
  readonly shiftFormula: (formula: string, columns: number, rows: number) => string
}

const NUMBER_TYPE = 2
const STRING_TYPE = 1
const BOOLEAN_TYPE = 3

const isBlank = (cell: CopiedSourceCell | null): boolean =>
  cell === null || ((cell.v === null || cell.v === undefined || cell.v === '') && !cell.f)

const asNumber = (value: unknown): number | null =>
  typeof value === 'number' && Number.isFinite(value) ? value : null

const OPERATORS: Record<Exclude<PasteOperation, 'none'>, string> = {
  add: '+',
  subtract: '-',
  multiply: '*',
  divide: '/',
}

function combine(
  operation: Exclude<PasteOperation, 'none'>,
  destination: number,
  source: number,
): number | null {
  switch (operation) {
    case 'add':
      return destination + source
    case 'subtract':
      return destination - source
    case 'multiply':
      return destination * source
    case 'divide':
      return source === 0 ? null : destination / source
  }
}

/// 15 significant digits, like the cell Excel stores
const tidy = (value: number): number => Number(value.toPrecision(15))

const withoutBorders = (style: IStyleData | null): IStyleData | null => {
  if (!style) return null
  const { bd: _bd, ...rest } = style
  return rest
}

/// the value side of one pasted cell, or undefined when the cell's value must not change
function planValue(
  source: CopiedSourceCell | null,
  destination: DestinationCell,
  options: PasteSpecialOptions,
  formula: string | null,
): PlannedWrite['value'] | undefined {
  const { what, operation } = options
  if (
    what === 'formats' ||
    what === 'columnWidths' ||
    what === 'comments' ||
    what === 'validation'
  ) {
    return undefined
  }
  const keepsFormulas =
    what === 'all' ||
    what === 'allUsingSourceTheme' ||
    what === 'allMergingConditionalFormats' ||
    what === 'formulas' ||
    what === 'allExceptBorders' ||
    what === 'formulasAndNumberFormats'
  if (operation !== 'none') {
    // a text / blank / error source has nothing to combine: the destination stays as it is
    const sourceNumber = asNumber(source?.v)
    if (sourceNumber === null || source?.t === STRING_TYPE || source?.t === BOOLEAN_TYPE) {
      return undefined
    }
    const op = OPERATORS[operation]
    if (destination.f) {
      return {
        v: null,
        t: null,
        f: `=(${destination.f.replace(/^=/, '')})${op}${sourceNumber}`,
      }
    }
    const destinationValue = destination.v
    if (
      destinationValue !== null &&
      destinationValue !== undefined &&
      destinationValue !== '' &&
      asNumber(destinationValue) === null
    ) {
      return undefined // text destination
    }
    const result = combine(operation, asNumber(destinationValue) ?? 0, sourceNumber)
    // ÷ 0 has no number: a formula shows Excel's #DIV/0!
    if (result === null) {
      return { v: null, t: null, f: `=${asNumber(destinationValue) ?? 0}/0` }
    }
    return { v: tidy(result), t: NUMBER_TYPE, f: null }
  }
  if (!source || isBlank(source)) return { v: null, t: null, f: null }
  if (keepsFormulas && formula) return { v: null, t: null, f: formula }
  return { v: source.v ?? null, t: source.t ?? null, f: null }
}

/// the rectangle a paste covers: the block (turned for Transpose), repeated over a selection that is a
/// whole multiple of it
export function pasteExtent(
  rowCount: number,
  columnCount: number,
  options: PasteSpecialOptions,
  target: PasteTarget,
): IRange {
  const blockRows = options.transpose ? columnCount : rowCount
  const blockColumns = options.transpose ? rowCount : columnCount
  const tileRows =
    !options.transpose && target.selectionRows % blockRows === 0
      ? target.selectionRows / blockRows
      : 1
  const tileColumns =
    !options.transpose && target.selectionColumns % blockColumns === 0
      ? target.selectionColumns / blockColumns
      : 1
  return {
    startRow: target.anchorRow,
    startColumn: target.anchorColumn,
    endRow: target.anchorRow + blockRows * tileRows - 1,
    endColumn: target.anchorColumn + blockColumns * tileColumns - 1,
  }
}

export function planPaste(input: PasteInput): PastePlan {
  const { source, options, target } = input
  const rowCount = source.rows.length
  const columnCount = source.columns.length
  if (rowCount === 0 || columnCount === 0) return { error: 'empty' }

  const blockRows = options.transpose ? columnCount : rowCount
  const blockColumns = options.transpose ? rowCount : columnCount

  if (options.transpose && input.sameSheet) {
    const top = source.rows[0]!
    const bottom = source.rows[rowCount - 1]!
    const left = source.columns[0]!
    const right = source.columns[columnCount - 1]!
    const overlaps =
      target.anchorRow <= bottom &&
      target.anchorRow + blockRows - 1 >= top &&
      target.anchorColumn <= right &&
      target.anchorColumn + blockColumns - 1 >= left
    const sameShape =
      blockRows === bottom - top + 1 &&
      blockColumns === right - left + 1 &&
      target.anchorRow === top &&
      target.anchorColumn === left
    if (overlaps && !sameShape) return { error: 'overlap' }
  }

  // tiling: a selection that is a whole multiple of the block repeats it (never for Transpose)
  const tileRows =
    !options.transpose && target.selectionRows % blockRows === 0
      ? target.selectionRows / blockRows
      : 1
  const tileColumns =
    !options.transpose && target.selectionColumns % blockColumns === 0
      ? target.selectionColumns / blockColumns
      : 1

  const transposeBlock = {
    rows: source.rows,
    columns: source.columns,
    anchorRow: target.anchorRow,
    anchorColumn: target.anchorColumn,
    ...(input.sourceSheetName === undefined ? {} : { sheetName: input.sourceSheetName }),
  }
  const writes: PlannedWrite[] = []
  for (let tr = 0; tr < tileRows; tr += 1) {
    for (let tc = 0; tc < tileColumns; tc += 1) {
      const baseRow = target.anchorRow + tr * blockRows
      const baseColumn = target.anchorColumn + tc * blockColumns
      for (let i = 0; i < rowCount; i += 1) {
        for (let j = 0; j < columnCount; j += 1) {
          const cell = source.cells[i]?.[j] ?? null
          const row = baseRow + (options.transpose ? j : i)
          const column = baseColumn + (options.transpose ? i : j)
          if (options.skipBlanks && isBlank(cell)) continue

          const sourceRow = source.rows[i]!
          const sourceColumn = source.columns[j]!
          let formula: string | null = null
          if (cell?.f) {
            formula = options.transpose
              ? transposeFormula(
                  cell.f,
                  { row: sourceRow, column: sourceColumn },
                  { row, column },
                  { ...transposeBlock, anchorRow: baseRow, anchorColumn: baseColumn },
                )
              : input.shiftFormula(cell.f, column - sourceColumn, row - sourceRow)
          }
          const destination = input.readDestination(row, column)
          const value = planValue(cell, destination, options, formula)

          let style: PlannedWrite['style']
          switch (options.what) {
            case 'all':
            case 'allUsingSourceTheme':
            case 'allMergingConditionalFormats':
            case 'formats':
              style = { mode: 'replace', style: cell?.style ?? null }
              break
            case 'allExceptBorders': {
              const keep = input.readDestinationStyle(row, column)?.bd
              const next = withoutBorders(cell?.style ?? null)
              style = { mode: 'replace', style: keep ? { ...(next ?? {}), bd: keep } : next }
              break
            }
            case 'formulasAndNumberFormats':
            case 'valuesAndNumberFormats':
              style = { mode: 'numfmt', pattern: cell?.style?.n?.pattern ?? null }
              break
            default:
              style = undefined
          }
          if (value === undefined && style === undefined) continue
          writes.push({
            row,
            column,
            ...(value === undefined ? {} : { value }),
            ...(style === undefined ? {} : { style }),
          })
        }
      }
    }
  }
  return {
    writes,
    extent: {
      startRow: target.anchorRow,
      startColumn: target.anchorColumn,
      endRow: target.anchorRow + blockRows * tileRows - 1,
      endColumn: target.anchorColumn + blockColumns * tileColumns - 1,
    },
  }
}

/** one source cell and where the paste puts it (every tile of a repeated block) */
export interface PastedCell {
  readonly sourceRow: number
  readonly sourceColumn: number
  readonly row: number
  readonly column: number
}

function tiling(
  rowCount: number,
  columnCount: number,
  options: PasteSpecialOptions,
  target: PasteTarget,
): { blockRows: number; blockColumns: number; tileRows: number; tileColumns: number } {
  const blockRows = options.transpose ? columnCount : rowCount
  const blockColumns = options.transpose ? rowCount : columnCount
  return {
    blockRows,
    blockColumns,
    tileRows:
      !options.transpose && target.selectionRows % blockRows === 0
        ? target.selectionRows / blockRows
        : 1,
    tileColumns:
      !options.transpose && target.selectionColumns % blockColumns === 0
        ? target.selectionColumns / blockColumns
        : 1,
  }
}

/** where each copy of the block starts: one per tile (a selection that is a whole multiple of the block repeats it) */
export function tileOrigins(
  source: { readonly rows: readonly number[]; readonly columns: readonly number[] },
  options: PasteSpecialOptions,
  target: PasteTarget,
): { baseRow: number; baseColumn: number }[] {
  const { blockRows, blockColumns, tileRows, tileColumns } = tiling(
    source.rows.length,
    source.columns.length,
    options,
    target,
  )
  const origins: { baseRow: number; baseColumn: number }[] = []
  for (let tr = 0; tr < tileRows; tr += 1) {
    for (let tc = 0; tc < tileColumns; tc += 1) {
      origins.push({
        baseRow: target.anchorRow + tr * blockRows,
        baseColumn: target.anchorColumn + tc * blockColumns,
      })
    }
  }
  return origins
}

/** a range of the copied block (by source row / column) as the paste lays it out at one tile */
export function movedRange(
  range: IRange,
  source: { readonly rows: readonly number[]; readonly columns: readonly number[] },
  options: PasteSpecialOptions,
  origin: { baseRow: number; baseColumn: number },
): IRange | null {
  const i1 = source.rows.indexOf(range.startRow)
  const i2 = source.rows.indexOf(range.endRow)
  const j1 = source.columns.indexOf(range.startColumn)
  const j2 = source.columns.indexOf(range.endColumn)
  if (i1 < 0 || i2 < 0 || j1 < 0 || j2 < 0) return null
  if (i2 - i1 !== range.endRow - range.startRow) return null
  if (j2 - j1 !== range.endColumn - range.startColumn) return null
  return options.transpose
    ? {
        startRow: origin.baseRow + j1,
        endRow: origin.baseRow + j2,
        startColumn: origin.baseColumn + i1,
        endColumn: origin.baseColumn + i2,
      }
    : {
        startRow: origin.baseRow + i1,
        endRow: origin.baseRow + i2,
        startColumn: origin.baseColumn + j1,
        endColumn: origin.baseColumn + j2,
      }
}

/** the source cell behind each destination cell of the paste, in source order within each tile */
export function pastedCells(
  source: { readonly rows: readonly number[]; readonly columns: readonly number[] },
  options: PasteSpecialOptions,
  target: PasteTarget,
): PastedCell[] {
  const { blockRows, blockColumns, tileRows, tileColumns } = tiling(
    source.rows.length,
    source.columns.length,
    options,
    target,
  )
  const cells: PastedCell[] = []
  for (let tr = 0; tr < tileRows; tr += 1) {
    for (let tc = 0; tc < tileColumns; tc += 1) {
      const baseRow = target.anchorRow + tr * blockRows
      const baseColumn = target.anchorColumn + tc * blockColumns
      source.rows.forEach((sourceRow, i) => {
        source.columns.forEach((sourceColumn, j) => {
          cells.push({
            sourceRow,
            sourceColumn,
            row: baseRow + (options.transpose ? j : i),
            column: baseColumn + (options.transpose ? i : j),
          })
        })
      })
    }
  }
  return cells
}

/**
 * The merged ranges a paste makes: each merge inside the copied block, moved to every tile (and
 * turned by Transpose). A merge that is only partly inside the block, or that spans a gap of
 * filtered rows, is not carried; a 1 × 1 range is no merge.
 */
export function mergedRangesAfterPaste(
  merges: readonly IRange[],
  source: { readonly rows: readonly number[]; readonly columns: readonly number[] },
  options: PasteSpecialOptions,
  target: PasteTarget,
): IRange[] {
  const result: IRange[] = []
  const origins = tileOrigins(source, options, target)
  for (const merge of merges) {
    if (merge.startRow === merge.endRow && merge.startColumn === merge.endColumn) continue
    for (const origin of origins) {
      const moved = movedRange(merge, source, options, origin)
      if (moved) result.push(moved)
    }
  }
  return result
}
