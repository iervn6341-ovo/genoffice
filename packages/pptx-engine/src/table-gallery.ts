import {
  BUILTIN_TABLE_STYLES,
  cellPartStyle,
  cellStyleBorders,
  resolveTableStyle,
  type BuiltinTableStyle,
  type TableStyleFlags,
} from './table-style'
import type { Theme } from './theme'
import type { Fill, Stroke } from './types'

/** One painted cell of a gallery thumbnail: resolved colors only, no theme lookups left to do. */
export interface StylePreviewCell {
  fill?: string
  text?: string
  bold?: boolean
  /** edge line colors */
  l?: string
  r?: string
  t?: string
  b?: string
}

export interface StyleGalleryEntry extends BuiltinTableStyle {
  /** rows × cols thumbnail under the requested region toggles */
  cells: StylePreviewCell[][]
}

const solidHex = (fill: Fill | undefined): string | undefined =>
  fill?.type === 'solid' ? fill.color.slice(0, 7).toUpperCase() : undefined

const lineHex = (stroke: Stroke | undefined): string | undefined => solidHex(stroke?.fill)

/**
 * The built-in table styles as small painted grids for the ribbon's style gallery. Colors come from
 * the same resolution the renderer uses, so a thumbnail matches what applying the style produces
 * in this deck (its theme accents, the current Header Row / Banded Rows / … toggles).
 */
export function tableStyleGallery(
  theme: Theme | undefined,
  flags: TableStyleFlags,
  rows = 4,
  cols = 5,
): StyleGalleryEntry[] {
  return BUILTIN_TABLE_STYLES.map((style) => {
    const def = resolveTableStyle(style.id, undefined, theme) ?? {}
    const cells: StylePreviewCell[][] = []
    for (let r = 0; r < rows; r++) {
      const row: StylePreviewCell[] = []
      for (let c = 0; c < cols; c++) {
        const part = cellPartStyle(def, flags, r, c, rows, cols)
        const edges = cellStyleBorders(def, flags, r, c, rows, cols)
        const cell: StylePreviewCell = {}
        const fill = solidHex(part.fill) ?? solidHex(def.tblBg)
        if (fill) cell.fill = fill
        if (part.textColor) cell.text = part.textColor.slice(0, 7).toUpperCase()
        if (part.bold) cell.bold = true
        for (const k of ['l', 'r', 't', 'b'] as const) {
          const hex = lineHex(edges[k])
          if (hex) cell[k] = hex
        }
        row.push(cell)
      }
      cells.push(row)
    }
    return { ...style, cells }
  })
}
