import { BooleanNumber, BorderStyleTypes, type IStyleData } from '@univerjs/core'

import {
  EXACT_TINTS,
  TABLE_STYLE_DEFS,
  type BorderDef,
  type BorderLine,
  type ColorRef,
  type ElementDef,
  type TableStyleDef,
  type ThemeSlot,
} from './table-style-defs'
import { applyTint, COLOR_SCHEMES } from './themes'

/**
 * Home → Format as Table. Univer paints every table with its own lavender theme, so the style the
 * user picked (which the save writes as `tableStyleInfo`) was never drawn. This turns Excel's
 * built-in style definition (table-style-defs.ts — all 60, with their theme-color references) into
 * a Univer table theme, resolved against the workbook's own theme colors, and into the cell looks
 * the style gallery draws. Excel draws its own exact style from the saved name on open.
 */
export interface TableThemeItem {
  bg?: { rgb: string }
  cl?: { rgb: string }
  bl?: BooleanNumber
  bd?: { t?: BorderEdge; b?: BorderEdge; l?: BorderEdge; r?: BorderEdge }
}

interface BorderEdge {
  s: BorderStyleTypes
  cl: { rgb: string }
}

export interface TableThemeJson {
  name: string
  wholeStyle?: TableThemeItem
  headerRowStyle?: TableThemeItem
  headerColumnStyle?: TableThemeItem
  firstRowStyle?: TableThemeItem
  secondRowStyle?: TableThemeItem
  lastRowStyle?: TableThemeItem
  firstColumnStyle?: TableThemeItem
  secondColumnStyle?: TableThemeItem
  lastColumnStyle?: TableThemeItem
}

/** Excel's Table Design → Table Style Options that change how the style is drawn */
export interface TableStyleOptions {
  /** default true: the style's row stripes */
  readonly bandedRows?: boolean
  readonly bandedColumns?: boolean
  readonly firstColumn?: boolean
  readonly lastColumn?: boolean
  /** default true; without a header row the first row of the area is data */
  readonly headerRow?: boolean
  /** the last row of the area is the totals row */
  readonly totalRow?: boolean
}

/** theme colors in theme index order [lt1, dk1, lt2, dk2, accent1–6, hlink, folHlink], #RRGGBB */
export type ThemeColors = readonly string[]

/** a workbook without its own theme part draws with Office's 2013–2022 colors, like Excel does */
export const DEFAULT_THEME_COLORS: ThemeColors = COLOR_SCHEMES[0]!.values

const SLOT_INDEX: Record<ThemeSlot, number> = {
  lt1: 0,
  dk1: 1,
  lt2: 2,
  dk2: 3,
  a1: 4,
  a2: 5,
  a3: 6,
  a4: 7,
  a5: 8,
  a6: 9,
}

const BLACK = '#000000'

/** the color a reference stands for under `colors` (the exact color Excel gives for known schemes) */
export function resolveColor(ref: ColorRef, colors: ThemeColors = DEFAULT_THEME_COLORS): string {
  if ('rgb' in ref) return ref.rgb.toUpperCase()
  const base = (colors[SLOT_INDEX[ref.s]] ?? BLACK).toUpperCase()
  if (!ref.t) return base
  const exact = EXACT_TINTS[`${base.slice(1)}:${ref.t}`]
  return exact ? `#${exact}` : applyTint(base, ref.t)
}

const LINE_STYLE: Record<BorderLine, BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  medium: BorderStyleTypes.MEDIUM,
  thick: BorderStyleTypes.THICK,
  double: BorderStyleTypes.DOUBLE,
}

const rgb = (hex: string) => ({ rgb: hex })

/// the table styles made in this app (Table Design → New Table Style), by name: a custom style is
/// known to the workbook by the name alone, and drawn from this definition
const customDefs = new Map<string, TableStyleDef>()

/** replaces the custom styles the workbook holds (a different workbook has different ones) */
export function setCustomTableStyleDefs(defs: ReadonlyMap<string, TableStyleDef>): void {
  customDefs.clear()
  for (const [name, def] of defs) customDefs.set(name, def)
}

export function setCustomTableStyleDef(name: string, def: TableStyleDef | null): void {
  if (def) customDefs.set(name, def)
  else customDefs.delete(name)
}

/** the table style behind a menu value: one of Excel's 60, a custom one, or null for anything else */
export function tableStyleDef(style: string | undefined): TableStyleDef | null {
  return style === undefined ? null : (TABLE_STYLE_DEFS[style] ?? customDefs.get(style) ?? null)
}

/** the names of Excel's 60 built-in styles */
export const TABLE_STYLE_NAMES: ReadonlySet<string> = new Set(Object.keys(TABLE_STYLE_DEFS))

export type TableStyleFamily = 'Light' | 'Medium' | 'Dark'

const FAMILY_SIZE: Record<TableStyleFamily, number> = { Light: 21, Medium: 28, Dark: 11 }

/** Excel's gallery order: Light 1–21, Medium 1–28, Dark 1–11 */
export const TABLE_STYLE_GROUPS: readonly {
  readonly family: TableStyleFamily
  readonly names: readonly string[]
}[] = (['Light', 'Medium', 'Dark'] as const).map((family) => ({
  family,
  names: Array.from({ length: FAMILY_SIZE[family] }, (_, i) => `TableStyle${family}${i + 1}`),
}))

/** "TableStyleMedium2" → "Medium 2" */
export function tableStyleLabel(style: string): string {
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(style)
  return match ? `${match[1]} ${match[2]}` : style
}

// ── Univer theme ─────────────────────────────────────────────────────────────────────────────

/** 'none' takes an edge another style of the merge drew off again (the outline's last row / column) */
type EdgeSpec = BorderDef | 'none'

interface ItemParts {
  bg?: string
  cl?: string
  bold?: boolean
  t?: EdgeSpec
  b?: EdgeSpec
  l?: EdgeSpec
  r?: EdgeSpec
}

function toItem(parts: ItemParts, colors: ThemeColors): TableThemeItem | undefined {
  const item: TableThemeItem = {}
  if (parts.bg) item.bg = rgb(parts.bg)
  if (parts.cl) item.cl = rgb(parts.cl)
  if (parts.bold) item.bl = BooleanNumber.TRUE
  const bd: NonNullable<TableThemeItem['bd']> = {}
  for (const edge of ['t', 'b', 'l', 'r'] as const) {
    const def = parts[edge]
    if (def === 'none') bd[edge] = { s: BorderStyleTypes.NONE, cl: rgb(BLACK) }
    else if (def) bd[edge] = { s: LINE_STYLE[def.l], cl: rgb(resolveColor(def.c, colors)) }
  }
  if (Object.keys(bd).length > 0) item.bd = bd
  return Object.keys(item).length > 0 ? item : undefined
}

/** an element's own fill / font / bold plus its own edge lines, as item parts */
function partsOf(element: ElementDef | undefined, colors: ThemeColors): ItemParts {
  if (!element) return {}
  return {
    ...(element.fill ? { bg: resolveColor(element.fill, colors) } : {}),
    ...(element.font ? { cl: resolveColor(element.font, colors) } : {}),
    ...(element.bold ? { bold: true } : {}),
    ...(element.top ? { t: element.top } : {}),
    ...(element.bottom ? { b: element.bottom } : {}),
    ...(element.left ? { l: element.left } : {}),
    ...(element.right ? { r: element.right } : {}),
  }
}

/**
 * The theme for a built-in style name, or null for a name that is none of Excel's 60 (the table
 * then keeps Univer's default). `uniqueName` must differ per application — Univer keys themes by
 * name. Univer draws edges per cell, so the whole-table outline rides on the edge rows / columns
 * (top on the header row, bottom on the last row, left on the first column, right on the last) and
 * the inside lines on every row / column; Excel's precedence (whole < column stripes < row stripes
 * < first / last column < header < total) is Univer's merge order too.
 */
export function tableThemeFor(
  style: string | undefined,
  uniqueName: string,
  colors: ThemeColors = DEFAULT_THEME_COLORS,
  options: TableStyleOptions = {},
): TableThemeJson | null {
  const def = tableStyleDef(style)
  if (!def) return null
  const whole = def.whole
  const theme: TableThemeJson = { name: uniqueName }
  const set = <K extends keyof TableThemeJson>(key: K, parts: ItemParts): void => {
    const item = toItem(parts, colors)
    if (item) (theme as unknown as Record<string, unknown>)[key] = item
  }

  set('wholeStyle', {
    ...(whole?.fill ? { bg: resolveColor(whole.fill, colors) } : {}),
    ...(whole?.font ? { cl: resolveColor(whole.font, colors) } : {}),
  })

  const insideH = whole?.insideH
  const insideV = whole?.insideV
  const header = partsOf(def.header, colors)
  const banded = options.bandedRows !== false
  const headerless = options.headerRow === false
  if (headerless) {
    // The theme is laid over the data rows (nothing is a header): the first data row stands where
    // Univer expects the header, so it takes the first stripe and the stripes after it swap.
    set('headerRowStyle', {
      ...(banded ? partsOf(def.stripe1, colors) : {}),
      ...(whole?.top ? { t: whole.top } : {}),
      ...(insideH ? { b: insideH } : {}),
    })
  } else {
    set('headerRowStyle', {
      ...header,
      ...(whole?.top && !header.t ? { t: whole.top } : {}),
      ...(insideH && !header.b ? { b: insideH } : {}),
    })
  }
  set('firstRowStyle', {
    ...(banded ? partsOf(headerless ? def.stripe2 : def.stripe1, colors) : {}),
    ...(insideH ? { b: insideH } : {}),
  })
  set('secondRowStyle', {
    ...(banded ? partsOf(headerless ? def.stripe1 : def.stripe2, colors) : {}),
    ...(insideH ? { b: insideH } : {}),
  })
  if (options.totalRow) {
    // the totals row is drawn apart (totalRowTheme); the last data row meets its top rule
    const rule = def.total?.top ?? insideH
    if (rule) set('lastRowStyle', { b: rule })
  } else if (whole?.bottom || insideH) {
    // the table's last row closes the outline (and has no line to the row below it)
    set('lastRowStyle', { b: whole?.bottom ?? 'none' })
  }

  // Univer counts column 0 as its "second" column style, so Excel's first column stripe lands there
  if (options.bandedColumns) {
    set('secondColumnStyle', {
      ...partsOf(def.colStripe1, colors),
      ...(insideV ? { r: insideV } : {}),
    })
    set('firstColumnStyle', {
      ...partsOf(def.colStripe2, colors),
      ...(insideV ? { r: insideV } : {}),
    })
  } else if (insideV) {
    set('secondColumnStyle', { r: insideV })
    set('firstColumnStyle', { r: insideV })
  }
  // the first / last column carry the outline's left / right edge, and their emphasis when on
  set('headerColumnStyle', {
    ...(options.firstColumn ? partsOf(def.firstCol, colors) : {}),
    ...(whole?.left ? { l: whole.left } : {}),
    ...(insideV ? { r: insideV } : {}),
  })
  set('lastColumnStyle', {
    ...(options.lastColumn ? partsOf(def.lastCol, colors) : {}),
    ...(whole?.right || insideV ? { r: whole?.right ?? 'none' } : {}),
  })
  return theme
}

/** a theme that draws nothing: what a table gets when its look is laid over it apart */
export const plainTheme = (name: string): TableThemeJson => ({ name })

/**
 * The totals row's look, as a theme for the one-row range it covers: Univer's table has no totals
 * row, so this is laid over the row on its own (display only — Excel draws the saved totals row
 * from the table style when it opens the file). Fill, font and the rule across its top come from
 * the style's total element; the outline's bottom, left and right edges close the table.
 */
export function totalRowTheme(
  style: string | undefined,
  uniqueName: string,
  colors: ThemeColors = DEFAULT_THEME_COLORS,
  options: TableStyleOptions = {},
): TableThemeJson | null {
  const def = tableStyleDef(style)
  if (!def) return null
  const whole = def.whole
  const total = partsOf(def.total, colors)
  const theme: TableThemeJson = { name: uniqueName }
  const set = <K extends keyof TableThemeJson>(key: K, parts: ItemParts): void => {
    const item = toItem(parts, colors)
    if (item) (theme as unknown as Record<string, unknown>)[key] = item
  }
  const insideV = whole?.insideV
  set('wholeStyle', {
    ...(whole?.fill ? { bg: resolveColor(whole.fill, colors) } : {}),
    ...(whole?.font ? { cl: resolveColor(whole.font, colors) } : {}),
    ...total,
    ...(whole?.bottom ? { b: whole.bottom } : {}),
  })
  // a one-row range counts as both the first and the last row: the edges ride on the columns
  if (insideV) {
    set('secondColumnStyle', { r: insideV })
    set('firstColumnStyle', { r: insideV })
  }
  set('headerColumnStyle', {
    ...(options.firstColumn ? partsOf(def.firstCol, colors) : {}),
    ...(whole?.left ? { l: whole.left } : {}),
    ...(insideV ? { r: insideV } : {}),
  })
  set('lastColumnStyle', {
    ...(options.lastColumn ? partsOf(def.lastCol, colors) : {}),
    ...(whole?.right || insideV ? { r: whole?.right ?? 'none' } : {}),
  })
  return theme
}

/** Univer's facade call that puts a theme on an existing table */
export interface TableThemeSheet {
  addTableTheme(tableId: string, theme: TableThemeJson): Promise<boolean> | boolean
}

// ── cell looks, for the gallery thumbnails (and the tests that pin the theme to them) ───────────

export interface CellEdge {
  readonly line: BorderLine
  readonly color: string
}

export interface CellLook {
  bg?: string
  color?: string
  bold?: boolean
  top?: CellEdge
  bottom?: CellEdge
  left?: CellEdge
  right?: CellEdge
}

/**
 * How one cell of a table draws under `style`: `row` 0 is the header, `col` 0 the first column.
 * Excel's precedence: whole < column stripes < row stripes < first / last column < header.
 */
export function tableCellLook(
  style: string,
  grid: { rows: number; columns: number; row: number; col: number },
  colors: ThemeColors = DEFAULT_THEME_COLORS,
  options: TableStyleOptions = {},
): CellLook {
  const def = tableStyleDef(style)
  if (!def) return {}
  const look: CellLook = {}
  const edge = (d: BorderDef): CellEdge => ({ line: d.l, color: resolveColor(d.c, colors) })
  const apply = (element: ElementDef | undefined): void => {
    if (!element) return
    if (element.fill) look.bg = resolveColor(element.fill, colors)
    if (element.font) look.color = resolveColor(element.font, colors)
    if (element.bold) look.bold = true
    if (element.top) look.top = edge(element.top)
    if (element.bottom) look.bottom = edge(element.bottom)
    if (element.left) look.left = edge(element.left)
    if (element.right) look.right = edge(element.right)
  }
  const { rows, columns, row, col } = grid
  const lastRow = row === rows - 1
  const lastCol = col === columns - 1
  // the whole-table outline and inside lines
  const whole = def.whole
  if (whole) {
    if (whole.fill) look.bg = resolveColor(whole.fill, colors)
    if (whole.font) look.color = resolveColor(whole.font, colors)
    if (row === 0 && whole.top) look.top = edge(whole.top)
    if (lastRow && whole.bottom) look.bottom = edge(whole.bottom)
    if (col === 0 && whole.left) look.left = edge(whole.left)
    if (lastCol && whole.right) look.right = edge(whole.right)
    if (whole.insideH && !lastRow) look.bottom = edge(whole.insideH)
    if (whole.insideV && !lastCol) look.right = edge(whole.insideV)
  }
  // a header row above the data, a totals row under it, and the stripes count the rows between
  const headerRows = options.headerRow === false ? 0 : 1
  const totalsRow = options.totalRow === true && lastRow
  if (options.bandedColumns) apply(col % 2 === 0 ? def.colStripe1 : def.colStripe2)
  if (options.bandedRows !== false && row >= headerRows && !totalsRow) {
    apply((row - headerRows) % 2 === 0 ? def.stripe1 : def.stripe2)
  }
  if (options.firstColumn && col === 0) apply(def.firstCol)
  if (options.lastColumn && lastCol) apply(def.lastCol)
  if (row === 0 && headerRows === 1) apply(def.header)
  if (totalsRow) apply(def.total)
  return look
}

// ── Convert to Range ─────────────────────────────────────────────────────────────────────────

/**
 * What Excel's Convert to Range leaves behind: the table style's look baked into the cells' own
 * formats (Excel 16, probed through COM: header fill, bold, white text, stripe fills and borders
 * all stay on the cells after `ListObject.Unlist()`). A `set-range-values` matrix keyed by
 * absolute row and column; null for a name that is not one of the 60 styles.
 */
export function bakeTableStyle(
  style: string,
  area: { startRow: number; startColumn: number; endRow: number; endColumn: number },
  colors: ThemeColors = DEFAULT_THEME_COLORS,
  options: TableStyleOptions = {},
): Record<number, Record<number, { s: TableThemeItem }>> | null {
  if (!tableStyleDef(style)) return null
  const rows = area.endRow - area.startRow + 1
  const columns = area.endColumn - area.startColumn + 1
  const matrix: Record<number, Record<number, { s: TableThemeItem }>> = {}
  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < columns; col += 1) {
      const look = tableCellLook(style, { rows, columns, row, col }, colors, options)
      const item: TableThemeItem = {}
      if (look.bg) item.bg = rgb(look.bg)
      if (look.color) item.cl = rgb(look.color)
      if (look.bold) item.bl = BooleanNumber.TRUE
      const bd: NonNullable<TableThemeItem['bd']> = {}
      for (const [key, edge] of [
        ['t', look.top],
        ['b', look.bottom],
        ['l', look.left],
        ['r', look.right],
      ] as const) {
        if (edge) bd[key] = { s: LINE_STYLE[edge.line], cl: rgb(edge.color) }
      }
      if (Object.keys(bd).length > 0) item.bd = bd
      if (Object.keys(item).length === 0) continue
      ;(matrix[area.startRow + row] ??= {})[area.startColumn + col] = { s: item }
    }
  }
  return matrix
}

/**
 * Excel puts a cell's own formatting above the table style: a yellow cell in a banded table stays
 * yellow after Convert to Range. `own` is what the cell already has; each attribute it sets is
 * dropped from the baked item.
 */
export function withoutOwnFormatting(
  item: TableThemeItem,
  own: Pick<IStyleData, 'bg' | 'cl' | 'bl' | 'bd'> | null,
): TableThemeItem | null {
  if (!own) return item
  const next: TableThemeItem = { ...item }
  if (own.bg) delete next.bg
  if (own.cl) delete next.cl
  if (own.bl !== undefined && own.bl !== null) delete next.bl
  if (next.bd) {
    const bd = { ...next.bd }
    for (const edge of ['t', 'b', 'l', 'r'] as const) {
      if (own.bd?.[edge]) delete bd[edge]
    }
    if (Object.keys(bd).length > 0) next.bd = bd
    else delete next.bd
  }
  return Object.keys(next).length > 0 ? next : null
}
