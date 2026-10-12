import {
  tableCellLook,
  type CellEdge,
  type CellLook,
  type TableStyleOptions,
  type ThemeColors,
} from './table-styles'

/// the thumbnail's grid: a header row and four body rows, four columns
export const PREVIEW_ROWS = 5
export const PREVIEW_COLUMNS = 4

const BORDER_WIDTH: Record<CellEdge['line'], string> = {
  thin: '1px solid',
  medium: '1.5px solid',
  thick: '2px solid',
  double: '3px double',
}

const edgeCss = (edge: CellEdge | undefined): string | undefined =>
  edge ? `${BORDER_WIDTH[edge.line]} ${edge.color}` : undefined

/**
 * The inline style of one thumbnail cell. These colors are the table style's own — document data
 * that is not re-themed — so they are written inline on the cells, on a white paper.
 */
export interface PreviewCellStyle {
  readonly background?: string
  readonly borderTop?: string
  readonly borderBottom?: string
  readonly borderLeft?: string
  readonly borderRight?: string
}

export function previewCellStyle(look: CellLook): PreviewCellStyle {
  const style: { -readonly [K in keyof PreviewCellStyle]: PreviewCellStyle[K] } = {}
  if (look.bg) style.background = look.bg
  const top = edgeCss(look.top)
  const bottom = edgeCss(look.bottom)
  const left = edgeCss(look.left)
  const right = edgeCss(look.right)
  if (top) style.borderTop = top
  if (bottom) style.borderBottom = bottom
  if (left) style.borderLeft = left
  if (right) style.borderRight = right
  return style
}

/** the thumbnail of `style` under `colors`, row by row */
export function tableStylePreview(
  style: string,
  colors: ThemeColors,
  options: TableStyleOptions = {},
): PreviewCellStyle[] {
  const cells: PreviewCellStyle[] = []
  for (let row = 0; row < PREVIEW_ROWS; row += 1) {
    for (let col = 0; col < PREVIEW_COLUMNS; col += 1) {
      cells.push(
        previewCellStyle(
          tableCellLook(
            style,
            { rows: PREVIEW_ROWS, columns: PREVIEW_COLUMNS, row, col },
            colors,
            options,
          ),
        ),
      )
    }
  }
  return cells
}
