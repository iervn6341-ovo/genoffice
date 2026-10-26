/**
 * The GenOffice sheet payload that rides inside the copied HTML.
 *
 * Univer's own clipboard only keeps a cache id in the HTML (`data-copy-id`) and the cells in
 * a per-window cache, so a paste into another workbook window sees nothing but the visible
 * text and the few styles that survive as CSS: formulas turn into values, number formats,
 * borders and fills drift, merged cells, notes, validation and conditional formats vanish.
 * The payload carries the copied block itself (cell values, formulas, styles, merges, sizes
 * and the sheet features that live on the block) so another GenOffice window rebuilds the
 * same block. Other programs ignore the attribute and read the ordinary table markup.
 */
import type { IRange, IStyleData } from '@univerjs/core'

export const PAYLOAD_ATTR = 'data-genoffice-sheet'
export const PAYLOAD_VERSION = 1
/// A block whose payload is bigger than this travels as ordinary table markup only.
export const PAYLOAD_MAX_JSON_CHARS = 12_000_000

export interface PayloadCell {
  readonly v?: string | number | boolean | null
  readonly t?: number
  readonly f?: string
  /// index into `styles`
  readonly s?: number
  /// rich text document of the cell
  readonly p?: unknown
}

export interface PayloadNote {
  readonly row: number
  readonly column: number
  readonly text: string
}

export interface PayloadLink {
  readonly row: number
  readonly column: number
  readonly target: string
}

export interface ClipboardPayload {
  readonly version: number
  readonly bookId: string
  readonly sheetId: string
  readonly sheetName: string
  /// absolute rows / columns of the copied block (filtered-out rows are not listed)
  readonly rows: readonly number[]
  readonly columns: readonly number[]
  readonly styles: readonly IStyleData[]
  /// cells[i][j] belongs to rows[i] × columns[j]; 0 = empty
  readonly cells: readonly (readonly (PayloadCell | 0)[])[]
  readonly merges: readonly IRange[]
  /// pixel sizes per copied column / row (null = the sheet default)
  readonly columnWidths: readonly (number | null)[]
  readonly rowHeights: readonly (number | null)[]
  readonly notes: readonly PayloadNote[]
  /// data validation rules touching the block (ranges are absolute source coordinates)
  readonly validations: readonly unknown[]
  /// conditional format rules touching the block
  readonly conditionalFormats: readonly unknown[]
  readonly links: readonly PayloadLink[]
  /// the block was cut: the window it came from clears it once another window pastes it
  readonly cut?: boolean
  readonly copyId?: string
}

/// A style table that hands out one index per distinct style.
export function createStyleTable(): {
  indexOf(style: IStyleData | null | undefined): number | undefined
  readonly styles: IStyleData[]
} {
  const styles: IStyleData[] = []
  const known = new Map<string, number>()
  return {
    styles,
    indexOf(style) {
      if (!style || Object.keys(style).length === 0) return undefined
      const key = JSON.stringify(style)
      let index = known.get(key)
      if (index === undefined) {
        index = styles.length
        styles.push(style)
        known.set(key, index)
      }
      return index
    },
  }
}

function toBase64(text: string): string {
  const bytes = new TextEncoder().encode(text)
  let binary = ''
  const chunk = 0x8000
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk))
  }
  return btoa(binary)
}

function fromBase64(encoded: string): string {
  const binary = atob(encoded)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
  return new TextDecoder().decode(bytes)
}

/// The attribute value for a payload, or null when the block is too big to carry.
export function encodePayload(payload: ClipboardPayload): string | null {
  const json = JSON.stringify(payload)
  if (json.length > PAYLOAD_MAX_JSON_CHARS) return null
  return toBase64(json)
}

const isNumberArray = (value: unknown): value is number[] =>
  Array.isArray(value) && value.every((item) => typeof item === 'number' && Number.isInteger(item))

/// Reads an attribute value back; anything that is not a payload of this version is null.
export function decodePayload(encoded: string): ClipboardPayload | null {
  try {
    const raw = JSON.parse(fromBase64(encoded)) as Partial<ClipboardPayload> | null
    if (!raw || raw.version !== PAYLOAD_VERSION) return null
    if (!isNumberArray(raw.rows) || !isNumberArray(raw.columns)) return null
    if (raw.rows.length === 0 || raw.columns.length === 0) return null
    if (!Array.isArray(raw.cells) || raw.cells.length !== raw.rows.length) return null
    for (const line of raw.cells) {
      if (!Array.isArray(line) || line.length !== raw.columns.length) return null
    }
    return {
      version: PAYLOAD_VERSION,
      bookId: typeof raw.bookId === 'string' ? raw.bookId : '',
      sheetId: typeof raw.sheetId === 'string' ? raw.sheetId : '',
      sheetName: typeof raw.sheetName === 'string' ? raw.sheetName : '',
      rows: raw.rows,
      columns: raw.columns,
      styles: Array.isArray(raw.styles) ? raw.styles : [],
      cells: raw.cells as ClipboardPayload['cells'],
      merges: Array.isArray(raw.merges) ? raw.merges : [],
      columnWidths: Array.isArray(raw.columnWidths) ? raw.columnWidths : [],
      rowHeights: Array.isArray(raw.rowHeights) ? raw.rowHeights : [],
      notes: Array.isArray(raw.notes) ? raw.notes : [],
      validations: Array.isArray(raw.validations) ? raw.validations : [],
      conditionalFormats: Array.isArray(raw.conditionalFormats) ? raw.conditionalFormats : [],
      links: Array.isArray(raw.links) ? raw.links : [],
      ...(raw.cut === true && typeof raw.copyId === 'string'
        ? { cut: true, copyId: raw.copyId }
        : {}),
    }
  } catch {
    return null
  }
}

/// The payload carried by clipboard HTML, if any.
export function extractPayload(html: string): ClipboardPayload | null {
  const match = new RegExp(`${PAYLOAD_ATTR}="([A-Za-z0-9+/=]+)"`).exec(html)
  return match ? decodePayload(match[1]!) : null
}

/// The used part of a block: trailing rows / columns with no value and no look are dropped when
/// the selection was a whole-row / whole-column one. `used` lists the (row, column) indexes inside
/// the block of the cells that hold something.
export function usedExtent(used: Iterable<readonly [number, number]>): {
  rows: number
  columns: number
} {
  let rows = 1
  let columns = 1
  for (const [i, j] of used) {
    rows = Math.max(rows, i + 1)
    columns = Math.max(columns, j + 1)
  }
  return { rows, columns }
}
