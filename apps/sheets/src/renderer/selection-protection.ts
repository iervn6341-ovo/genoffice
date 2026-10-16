import type { WorkbookCellStyle, WorkbookFile, WorkbookRangeResult } from '../shared/desktop-api'
import type { EditJournal } from './edit-journal'
import type { StructuralOp } from '@genoffice/xlsx-gateway/gateway/xlsx-structure'
import { screenToFile } from './view-transform'
import { visibleBorderSeam } from './merged-border-geometry'

export interface CellProtectionFlags {
  readonly row: number
  readonly column: number
  /// The cell's original explicit XF, or null when it inherits a row/column XF.
  readonly styleReference?: number | null
  readonly locked: boolean
  readonly hidden: boolean
  readonly horizontalAlignment: string
  readonly verticalAlignment: string
  readonly textRotation: number
  /// A rich or multiline document uses Univer's document renderer, which
  /// does not support the full-cell accounting underline stroke.
  readonly richDocument?: boolean
  readonly shrinkToFit: boolean
  readonly underlineStyle?: 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting'
  readonly fontVerticalAlign?: 'baseline' | 'superscript' | 'subscript'
  readonly borderTop?: boolean
  readonly borderBottom?: boolean
  readonly borderLeft?: boolean
  readonly borderRight?: boolean
  readonly diagonalUp?: boolean
  readonly diagonalDown?: boolean
}

export type SelectionBorderControl =
  | 'left'
  | 'top'
  | 'bottom'
  | 'right'
  | 'insideHorizontal'
  | 'insideVertical'
  | 'diagonalUp'
  | 'diagonalDown'

export type SelectionBorderState = Readonly<Record<SelectionBorderControl, boolean | null>>
export interface SelectionBorderInfo extends SelectionBorderState {
  readonly lineStyle: string | null
  readonly lineColor: string | null
}

export interface SelectionAlignment {
  readonly horizontal: string | null
  readonly vertical: string | null
  readonly rotation: number | null
  readonly shrinkToFit: boolean | null
  readonly merged: boolean | null
}

export interface SelectionProtection {
  readonly sheetId: string
  readonly sessionId?: string
  readonly range: ProtectionRange
  readonly merges?: readonly ProtectionRange[]
  /// Merge definitions are after sheetData in OOXML; an indexed row alone does not prove completeness.
  readonly mergesComplete?: boolean
  /// null means the selected cells disagree; the dialog shows Mixed.
  readonly locked: boolean | null
  readonly hidden: boolean | null
  readonly alignment: SelectionAlignment
  readonly font?: {
    readonly underlineStyle:
      'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | null
    readonly verticalAlign: 'baseline' | 'superscript' | 'subscript' | null
  }
  readonly border?: SelectionBorderInfo
  readonly cells: readonly CellProtectionFlags[]
}

export interface ProtectionRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/**
 * The native range reader is bounded by the file's used extent, while Excel
 * still lets the user format blank grid cells beyond that extent. Read the
 * closest in-file row/column for row defaults, then resolve the full selected
 * screen range from the returned sparse cells, sheet columns, and XF 0.
 */
export function protectionReadRange(
  selection: ProtectionRange,
  rowCount: number,
  columnCount: number,
  hasStructuralOps: boolean,
): ProtectionRange | null {
  if (hasStructuralOps) return selection
  if (rowCount < 1 || columnCount < 1) return null
  const lastRow = rowCount - 1
  const lastColumn = columnCount - 1
  return {
    startRow: Math.min(selection.startRow, lastRow),
    endRow: Math.min(selection.endRow, lastRow),
    startColumn: Math.min(selection.startColumn, lastColumn),
    endColumn: Math.min(selection.endColumn, lastColumn),
  }
}

/** Resolves OOXML cell > row > column > normal xf, then pending per-cell edits. */
export function selectionProtection(
  styles: readonly WorkbookCellStyle[],
  sheet: WorkbookFile['sheets'][number],
  read:
    | (Pick<WorkbookRangeResult, 'cells' | 'rows'> & Partial<Pick<WorkbookRangeResult, 'merges'>>)
    | null,
  journal: EditJournal,
  sheetId: string,
  range: ProtectionRange,
  ops: readonly StructuralOp[] = [],
): SelectionProtection {
  const fileCells = new Map(read?.cells.map((cell) => [`${cell.row}:${cell.column}`, cell]) ?? [])
  const fileRows = new Map(read?.rows.map((row) => [row.row, row.styleIndex]) ?? [])
  const journalCells = journal.cells.get(sheetId)
  const cells: CellProtectionFlags[] = []
  let locked: boolean | null | undefined
  let hidden: boolean | null | undefined
  let horizontal: string | null | undefined
  let vertical: string | null | undefined
  let rotation: number | null | undefined
  let shrinkToFit: boolean | null | undefined
  let underlineStyle: CellProtectionFlags['underlineStyle'] | null | undefined
  let fontVerticalAlign: CellProtectionFlags['fontVerticalAlign'] | null | undefined
  const borderLines: { style: string; color: string }[] = []
  const overlaps = (a: ProtectionRange, b: ProtectionRange): boolean =>
    a.startRow <= b.endRow &&
    a.endRow >= b.startRow &&
    a.startColumn <= b.endColumn &&
    a.endColumn >= b.startColumn
  const intersects = (merge: ProtectionRange): boolean => overlaps(merge, range)
  let activeMerges = read?.merges?.filter(intersects) ?? []
  for (const operation of ops) {
    if (operation.kind === 'merge-cells' && intersects(operation.range)) {
      activeMerges = [
        ...activeMerges.filter((item) => !overlaps(item, operation.range)),
        operation.range,
      ]
    }
    if (operation.kind === 'unmerge-cells' && intersects(operation.range)) {
      activeMerges = activeMerges.filter((item) => !overlaps(item, operation.range))
    }
  }
  const onlyMerge = activeMerges.length === 1 ? activeMerges[0] : undefined
  const merged: boolean | null =
    activeMerges.length === 0
      ? false
      : onlyMerge &&
          onlyMerge.startRow <= range.startRow &&
          onlyMerge.endRow >= range.endRow &&
          onlyMerge.startColumn <= range.startColumn &&
          onlyMerge.endColumn >= range.endColumn
        ? true
        : null
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const key = `${row}:${column}`
      const edit = journalCells?.get(key)
      const fileColumn = screenToFile(ops, 'column', column)
      let columnStyle: number | undefined
      if (fileColumn !== null) {
        for (let index = sheet.columnWidths.length - 1; index >= 0; index -= 1) {
          const span = sheet.columnWidths[index]
          if (
            span?.styleIndex !== undefined &&
            fileColumn >= span.startColumn &&
            fileColumn <= span.endColumn
          ) {
            columnStyle = span.styleIndex
            break
          }
        }
      }
      const fileCell = fileCells.get(key)
      const styleIndex =
        (edit?.styleReset
          ? 0
          : edit?.styleReference === null
            ? undefined
            : (edit?.styleReference ??
              (fileCell?.explicitStyle === false ? undefined : fileCell?.styleIndex))) ??
        fileRows.get(row) ??
        columnStyle ??
        0
      const style = styles[styleIndex]
      const cellLocked = edit?.style?.protectionLocked ?? style?.protectionLocked !== false
      const cellHidden = edit?.style?.protectionHidden ?? style?.protectionHidden === true
      const cellHorizontal =
        edit?.style?.horizontalAlignment ?? style?.horizontalAlignment ?? 'general'
      const cellVertical = edit?.style?.verticalAlignment ?? style?.verticalAlignment ?? 'bottom'
      const cellRotation = edit?.style?.textRotation ?? style?.textRotation ?? 0
      const cellValue = edit?.hasValue ? edit.value : fileCell?.value
      const richDocument =
        Boolean((edit?.hasValue ? edit.rich : fileCell?.rich)?.length) ||
        (typeof cellValue === 'string' && /[\r\n]/.test(cellValue)) ||
        (cellHorizontal === 'fill' && typeof cellValue === 'string' && cellValue !== '')
      const cellShrink = edit?.style?.shrinkToFit ?? style?.shrinkToFit === true
      const cellUnderline: NonNullable<CellProtectionFlags['underlineStyle']> =
        edit?.style?.underline === false
          ? 'none'
          : (edit?.style?.underlineStyle ??
            (edit?.style?.underline === true
              ? 'single'
              : style?.underline
                ? (style.underlineStyle ?? 'single')
                : 'none'))
      const cellFontVertical: NonNullable<CellProtectionFlags['fontVerticalAlign']> =
        edit?.style?.fontVerticalAlign ?? style?.fontVerticalAlign ?? 'baseline'
      const borderOn = (
        key: 'borderTop' | 'borderBottom' | 'borderLeft' | 'borderRight',
      ): boolean => (edit?.style?.[key] === undefined ? style?.[key] : edit.style[key]) != null
      for (const key of [
        'borderTop',
        'borderBottom',
        'borderLeft',
        'borderRight',
        'borderDiagonal',
      ] as const) {
        const edge = edit?.style?.[key] === undefined ? style?.[key] : edit.style[key]
        if (edge)
          borderLines.push({
            style: edge.style,
            color: typeof edge.color === 'string' ? edge.color.toUpperCase() : '#000000',
          })
      }
      const diagonalEdge =
        edit?.style?.borderDiagonal === undefined
          ? style?.borderDiagonal
          : edit.style.borderDiagonal
      cells.push({
        row,
        column,
        styleReference: fileCell?.explicitStyle === false ? null : (fileCell?.styleIndex ?? null),
        locked: cellLocked,
        hidden: cellHidden,
        horizontalAlignment: cellHorizontal,
        verticalAlignment: cellVertical,
        textRotation: cellRotation,
        richDocument,
        shrinkToFit: cellShrink,
        underlineStyle: cellUnderline,
        fontVerticalAlign: cellFontVertical,
        borderTop: borderOn('borderTop'),
        borderBottom: borderOn('borderBottom'),
        borderLeft: borderOn('borderLeft'),
        borderRight: borderOn('borderRight'),
        diagonalUp: diagonalEdge != null && (edit?.style?.diagonalUp ?? style?.diagonalUp) === true,
        diagonalDown:
          diagonalEdge != null && (edit?.style?.diagonalDown ?? style?.diagonalDown) === true,
      })
      locked = locked === undefined ? cellLocked : locked === cellLocked ? locked : null
      hidden = hidden === undefined ? cellHidden : hidden === cellHidden ? hidden : null
      horizontal =
        horizontal === undefined
          ? cellHorizontal
          : horizontal === cellHorizontal
            ? horizontal
            : null
      vertical = vertical === undefined ? cellVertical : vertical === cellVertical ? vertical : null
      rotation = rotation === undefined ? cellRotation : rotation === cellRotation ? rotation : null
      shrinkToFit =
        shrinkToFit === undefined ? cellShrink : shrinkToFit === cellShrink ? shrinkToFit : null
      underlineStyle =
        underlineStyle === undefined
          ? cellUnderline
          : underlineStyle === cellUnderline
            ? underlineStyle
            : null
      fontVerticalAlign =
        fontVerticalAlign === undefined
          ? cellFontVertical
          : fontVerticalAlign === cellFontVertical
            ? fontVerticalAlign
            : null
    }
  }
  return {
    sheetId,
    range,
    merges: activeMerges,
    locked: locked ?? null,
    hidden: hidden ?? null,
    alignment: {
      horizontal: horizontal ?? null,
      vertical: vertical ?? null,
      rotation: rotation ?? null,
      shrinkToFit: shrinkToFit ?? null,
      merged,
    },
    font: {
      underlineStyle: underlineStyle ?? null,
      verticalAlign: fontVerticalAlign ?? null,
    },
    border: {
      ...selectionBorderState(cells, range, activeMerges),
      lineStyle:
        borderLines.length === 0
          ? 'thin'
          : borderLines.every((line) => line.style === borderLines[0]?.style)
            ? borderLines[0]!.style
            : null,
      lineColor:
        borderLines.length === 0
          ? '#000000'
          : borderLines.every((line) => line.color === borderLines[0]?.color)
            ? borderLines[0]!.color
            : null,
    },
    cells,
  }
}

export function selectionBorderState(
  cells: readonly CellProtectionFlags[],
  range: ProtectionRange,
  merges: readonly ProtectionRange[] = [],
): SelectionBorderState {
  const byCell = new Map(cells.map((cell) => [`${cell.row}:${cell.column}`, cell]))
  const values: Record<SelectionBorderControl, boolean[]> = {
    left: [],
    top: [],
    bottom: [],
    right: [],
    insideHorizontal: [],
    insideVertical: [],
    diagonalUp: [],
    diagonalDown: [],
  }
  for (const cell of cells) {
    if (cell.column === range.startColumn) values.left.push(cell.borderLeft === true)
    if (cell.column === range.endColumn) values.right.push(cell.borderRight === true)
    if (cell.row === range.startRow) values.top.push(cell.borderTop === true)
    if (cell.row === range.endRow) values.bottom.push(cell.borderBottom === true)
    values.diagonalUp.push(cell.diagonalUp === true)
    values.diagonalDown.push(cell.diagonalDown === true)
    if (cell.row < range.endRow && visibleBorderSeam(merges, cell.row, cell.column, 'down')) {
      const below = byCell.get(`${cell.row + 1}:${cell.column}`)
      values.insideHorizontal.push(cell.borderBottom === true || below?.borderTop === true)
    }
    if (
      cell.column < range.endColumn &&
      visibleBorderSeam(merges, cell.row, cell.column, 'right')
    ) {
      const right = byCell.get(`${cell.row}:${cell.column + 1}`)
      values.insideVertical.push(cell.borderRight === true || right?.borderLeft === true)
    }
  }
  const mixed = (items: readonly boolean[]): boolean | null =>
    items.length === 0 ? null : items.every((value) => value === items[0]) ? items[0]! : null
  return {
    left: mixed(values.left),
    top: mixed(values.top),
    bottom: mixed(values.bottom),
    right: mixed(values.right),
    insideHorizontal: mixed(values.insideHorizontal),
    insideVertical: mixed(values.insideVertical),
    diagonalUp: mixed(values.diagonalUp),
    diagonalDown: mixed(values.diagonalDown),
  }
}
