/**
 * Excel's own clipboard data → a `ForeignCopy`.
 *
 * Excel for Mac's HTML flavor carries only the displayed text, the style classes and the number
 * formats. Beside it Excel puts an embedded copy of the range (`com.microsoft.Embed-Source`: an
 * OLE file whose `Package` stream is a small .xlsx) and the source address
 * (`com.microsoft.Link`: `Excel\0<path>\0[Book.xlsx]Sheet1\0R1C1:R4C6`). Reading those gives what
 * the HTML cannot: formulas (moved by the distance between source and destination, like Excel's
 * own paste), exact values, real styles, merged cells and links.
 */
import * as CFB from 'cfb'
import { XMLParser } from 'fast-xml-parser'
import JSZip from 'jszip'

import { offsetFormulaRefs } from '@genoffice/xlsx-gateway/domain/formula-shift'
import { DEFAULT_THEME_PALETTE, applyTint } from '@genoffice/xlsx-gateway/domain/style-color'
import type { IRange } from '@univerjs/core'

import type { WorkbookCellStyle } from '../shared/desktop-api'
import type { ForeignCopy } from './clipboard-foreign'
import type { PayloadLink } from './clipboard-payload'
import type { CopiedCell } from './paste-source'

export const EXCEL_EMBED_FORMAT = 'com.microsoft.Embed-Source'
export const EXCEL_LINK_FORMAT = 'com.microsoft.Link'

/// The copied block's address from `com.microsoft.Link`.
export interface ExcelLink {
  readonly sheetName: string
  readonly range: IRange
}

export function parseExcelLink(link: string): ExcelLink | null {
  const match = /\[[^\]]*\]([^\0]+)\0R(\d+)C(\d+)(?::R(\d+)C(\d+))?/.exec(link)
  if (!match) return null
  const startRow = Number(match[2]) - 1
  const startColumn = Number(match[3]) - 1
  return {
    sheetName: match[1]!,
    range: {
      startRow,
      startColumn,
      endRow: match[4] === undefined ? startRow : Number(match[4]) - 1,
      endColumn: match[5] === undefined ? startColumn : Number(match[5]) - 1,
    },
  }
}

const BUILTIN_NUMBER_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
}

/// Excel's legacy indexed palette (indexed="8"… are the ones files use in practice)
const INDEXED_COLORS = [
  '#000000',
  '#FFFFFF',
  '#FF0000',
  '#00FF00',
  '#0000FF',
  '#FFFF00',
  '#FF00FF',
  '#00FFFF',
  '#000000',
  '#FFFFFF',
  '#FF0000',
  '#00FF00',
  '#0000FF',
  '#FFFF00',
  '#FF00FF',
  '#00FFFF',
  '#800000',
  '#008000',
  '#000080',
  '#808000',
  '#800080',
  '#008080',
  '#C0C0C0',
  '#808080',
  '#9999FF',
  '#993366',
  '#FFFFCC',
  '#CCFFFF',
  '#660066',
  '#FF8080',
  '#0066CC',
  '#CCCCFF',
  '#000080',
  '#FF00FF',
  '#FFFF00',
  '#00FFFF',
  '#800080',
  '#800000',
  '#008080',
  '#0000FF',
  '#00CCFF',
  '#CCFFFF',
  '#CCFFCC',
  '#FFFF99',
  '#99CCFF',
  '#FF99CC',
  '#CC99FF',
  '#FFCC99',
  '#3366FF',
  '#33CCCC',
  '#99CC00',
  '#FFCC00',
  '#FF9900',
  '#FF6600',
  '#666699',
  '#969696',
  '#003366',
  '#339966',
  '#003300',
  '#333300',
  '#993300',
  '#993366',
  '#333399',
  '#333333',
]

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  removeNSPrefix: true,
  isArray: (name) =>
    [
      'sheet',
      'Relationship',
      'si',
      'r',
      't',
      'font',
      'fill',
      'border',
      'xf',
      'numFmt',
      'row',
      'c',
      'col',
      'mergeCell',
      'hyperlink',
      'definedName',
    ].includes(name),
})

type Node = Record<string, unknown>
const asNode = (value: unknown): Node => (value && typeof value === 'object' ? (value as Node) : {})
const list = (value: unknown): Node[] => (Array.isArray(value) ? (value as Node[]) : [])
const attr = (node: Node, name: string): string | undefined => {
  const value = node[`@_${name}`]
  return typeof value === 'string' ? value : undefined
}
const flag = (node: Node | undefined, name: string): boolean => {
  if (!node || !(name in node)) return false
  const value = attr(asNode(node[name]), 'val')
  return value === undefined || (value !== '0' && value !== 'false')
}

/// The text of a shared string item (rich runs joined, phonetic runs left out).
function stringItemText(item: Node): string {
  const direct = item.t
  if (Array.isArray(direct)) {
    return direct.map((t) => textOf(t)).join('')
  }
  return list(item.r)
    .map((run) => list(run.t).map(textOf).join(''))
    .join('')
}
function textOf(value: unknown): string {
  if (typeof value === 'string') return value
  const node = asNode(value)
  const text = node['#text']
  return typeof text === 'string' ? text : ''
}

interface Palette {
  readonly theme: readonly string[]
}

function themePalette(themeXml: string | undefined): Palette {
  if (!themeXml) return { theme: DEFAULT_THEME_PALETTE }
  const scheme = asNode(
    asNode(asNode(asNode(parser.parse(themeXml)).theme).themeElements).clrScheme,
  )
  const pick = (name: string): string | undefined => {
    const node = asNode(scheme[name])
    const srgb = attr(asNode(node.srgbClr), 'val')
    if (srgb) return `#${srgb.toUpperCase()}`
    const sys = attr(asNode(node.sysClr), 'lastClr')
    return sys ? `#${sys.toUpperCase()}` : undefined
  }
  // Excel's theme index order swaps the dark / light pairs
  const order = [
    'lt1',
    'dk1',
    'lt2',
    'dk2',
    'accent1',
    'accent2',
    'accent3',
    'accent4',
    'accent5',
    'accent6',
    'hlink',
    'folHlink',
  ]
  return { theme: order.map((name, i) => pick(name) ?? DEFAULT_THEME_PALETTE[i]!) }
}

function colorOf(node: Node | undefined, palette: Palette): string | undefined {
  if (!node) return undefined
  const rgb = attr(node, 'rgb')
  if (rgb) return `#${rgb.slice(-6).toUpperCase()}`
  const theme = attr(node, 'theme')
  if (theme !== undefined) {
    const base = palette.theme[Number(theme)]
    if (!base) return undefined
    const tint = Number(attr(node, 'tint') ?? 0)
    return applyTint(base, Number.isFinite(tint) ? tint : 0)
  }
  const indexed = attr(node, 'indexed')
  if (indexed !== undefined) return INDEXED_COLORS[Number(indexed)]
  return undefined
}

interface Stylesheet {
  readonly xfs: readonly WorkbookCellStyle[]
}

function readStyles(xml: string, palette: Palette): Stylesheet {
  const root = asNode(asNode(parser.parse(xml)).styleSheet)
  const formats = new Map<number, string>()
  for (const fmt of list(asNode(root.numFmts).numFmt)) {
    const id = Number(attr(fmt, 'numFmtId'))
    const code = attr(fmt, 'formatCode')
    if (Number.isFinite(id) && code !== undefined) formats.set(id, code)
  }
  const fonts = list(asNode(root.fonts).font)
  const fills = list(asNode(root.fills).fill)
  const borders = list(asNode(root.borders).border)
  const xfs = list(asNode(root.cellXfs).xf).map((xf): WorkbookCellStyle => {
    const font = fonts[Number(attr(xf, 'fontId') ?? 0)] ?? {}
    const fill = fills[Number(attr(xf, 'fillId') ?? 0)] ?? {}
    const border = borders[Number(attr(xf, 'borderId') ?? 0)] ?? {}
    const alignment = asNode(xf.alignment)
    const numFmtId = Number(attr(xf, 'numFmtId') ?? 0)
    const format = formats.get(numFmtId) ?? BUILTIN_NUMBER_FORMATS[numFmtId]
    const underline = font.u === undefined ? undefined : (attr(asNode(font.u), 'val') ?? 'single')
    const pattern = asNode(fill.patternFill)
    const solid = attr(pattern, 'patternType') === 'solid'
    const fillColor = solid ? colorOf(asNode(pattern.fgColor), palette) : undefined
    const edge = (name: string): { style: string; color?: string } | undefined => {
      const side = asNode(border[name])
      const style = attr(side, 'style')
      if (!style) return undefined
      const color = colorOf(asNode(side.color), palette)
      return { style, ...(color ? { color } : {}) }
    }
    const top = edge('top')
    const bottom = edge('bottom')
    const left = edge('left')
    const right = edge('right')
    const fontFamily = attr(asNode(font.name), 'val')
    const fontSize = Number(attr(asNode(font.sz), 'val'))
    const fontColor = colorOf(asNode(font.color), palette)
    const horizontal = attr(alignment, 'horizontal')
    const vertical = attr(alignment, 'vertical')
    const rotation = Number(attr(alignment, 'textRotation') ?? 0)
    const indent = Number(attr(alignment, 'indent') ?? 0)
    return {
      ...(fontFamily ? { fontFamily } : {}),
      ...(Number.isFinite(fontSize) && fontSize > 0 ? { fontSize } : {}),
      bold: flag(font, 'b'),
      italic: flag(font, 'i'),
      underline: underline !== undefined && underline !== 'none',
      ...(underline === 'double' || underline === 'doubleAccounting'
        ? { underlineStyle: underline }
        : {}),
      strikethrough: flag(font, 'strike'),
      wrapText: attr(alignment, 'wrapText') === '1' || attr(alignment, 'wrapText') === 'true',
      ...(fontColor && fontColor !== '#000000' ? { fontColor } : {}),
      ...(fillColor ? { fillColor } : {}),
      ...(horizontal ? { horizontalAlignment: horizontal } : {}),
      ...(vertical ? { verticalAlignment: vertical } : {}),
      ...(indent > 0 ? { indent } : {}),
      ...(rotation > 0 ? { textRotation: rotation } : {}),
      ...(format && format !== 'General' ? { numberFormat: format } : {}),
      ...(top ? { borderTop: top } : {}),
      ...(bottom ? { borderBottom: bottom } : {}),
      ...(left ? { borderLeft: left } : {}),
      ...(right ? { borderRight: right } : {}),
      diagonalUp: false,
      diagonalDown: false,
    }
  })
  return { xfs }
}

const columnIndexOf = (label: string): number => {
  let n = 0
  for (const ch of label) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

function parseRef(ref: string): { row: number; column: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)
  return m ? { column: columnIndexOf(m[1]!), row: Number(m[2]) - 1 } : null
}

function parseArea(ref: string): IRange | null {
  const [a, b] = ref.split(':')
  const first = parseRef(a ?? '')
  const last = parseRef(b ?? a ?? '')
  if (!first || !last) return null
  return {
    startRow: first.row,
    endRow: last.row,
    startColumn: first.column,
    endColumn: last.column,
  }
}

export interface ExcelEmbedOptions {
  /// converts a style to Univer's style data (injected so this module stays free of the grid)
  readonly toStyle: (style: WorkbookCellStyle) => unknown
}

/// Reads the embedded workbook; null when the data is not what Excel writes.
export async function foreignFromExcelEmbed(
  embed: Uint8Array,
  link: ExcelLink,
  options: ExcelEmbedOptions,
): Promise<ForeignCopy | null> {
  let packageBytes: Uint8Array
  try {
    const container = CFB.read(embed, { type: 'array' })
    const stream = CFB.find(container, 'Package') ?? CFB.find(container, '/Package')
    if (!stream?.content) return null
    packageBytes = Uint8Array.from(stream.content as ArrayLike<number>)
  } catch {
    return null
  }
  const zip = await JSZip.loadAsync(packageBytes).catch(() => null)
  if (!zip) return null
  const read = async (path: string): Promise<string | undefined> => zip.file(path)?.async('string')

  const workbookXml = await read('xl/workbook.xml')
  const relsXml = await read('xl/_rels/workbook.xml.rels')
  const stylesXml = await read('xl/styles.xml')
  if (!workbookXml || !relsXml || !stylesXml) return null
  const workbook = asNode(asNode(parser.parse(workbookXml)).workbook)
  const sheets = list(asNode(workbook.sheets).sheet)
  const sheet = sheets.find((s) => attr(s, 'name') === link.sheetName) ?? sheets[0]
  if (!sheet) return null
  const rels = list(asNode(asNode(parser.parse(relsXml)).Relationships).Relationship)
  const targetOf = (id: string | undefined): string | undefined => {
    const rel = rels.find((r) => attr(r, 'Id') === id)
    const target = rel ? attr(rel, 'Target') : undefined
    if (!target) return undefined
    return target.startsWith('/') ? target.slice(1) : `xl/${target}`
  }
  const sheetPath = targetOf(attr(sheet, 'id'))
  const sheetXml = sheetPath ? await read(sheetPath) : undefined
  if (!sheetXml) return null

  const palette = themePalette(await read('xl/theme/theme1.xml'))
  const styles = readStyles(stylesXml, palette)
  const sharedXml = await read('xl/sharedStrings.xml')
  const strings = sharedXml
    ? list(asNode(asNode(parser.parse(sharedXml)).sst).si).map(stringItemText)
    : []
  const sheetRoot = asNode(asNode(parser.parse(sheetXml)).worksheet)
  // a whole-row / whole-column copy names a million cells; only the used part is carried
  const dimension = parseArea(attr(asNode(sheetRoot.dimension), 'ref') ?? '')
  const range: IRange = {
    startRow: link.range.startRow,
    startColumn: link.range.startColumn,
    endRow: dimension
      ? Math.max(link.range.startRow, Math.min(link.range.endRow, dimension.endRow))
      : link.range.endRow,
    endColumn: dimension
      ? Math.max(link.range.startColumn, Math.min(link.range.endColumn, dimension.endColumn))
      : link.range.endColumn,
  }
  const rows = Array.from(
    { length: range.endRow - range.startRow + 1 },
    (_, i) => range.startRow + i,
  )
  const columns = Array.from(
    { length: range.endColumn - range.startColumn + 1 },
    (_, j) => range.startColumn + j,
  )
  if (rows.length * columns.length > 1_000_000) return null

  const cells: (CopiedCell | null)[][] = rows.map(() => columns.map(() => null))
  const rowHeights: (number | null)[] = rows.map(() => null)
  const shared = new Map<string, { row: number; column: number; formula: string }>()
  const pending: { i: number; j: number; si: string; row: number; column: number }[] = []
  for (const row of list(asNode(sheetRoot.sheetData).row)) {
    const rowNumber = Number(attr(row, 'r')) - 1
    const i = rowNumber - range.startRow
    if (i < 0 || i >= rows.length) continue
    const height = Number(attr(row, 'ht'))
    if (Number.isFinite(height) && height > 0) rowHeights[i] = (height * 4) / 3
    for (const c of list(row.c)) {
      const at = parseRef(attr(c, 'r') ?? '')
      if (!at) continue
      const j = at.column - range.startColumn
      if (j < 0 || j >= columns.length) continue
      const type = attr(c, 't')
      const rawValue = textOf(c.v)
      const formulaNode = c.f === undefined ? undefined : asNode(c.f)
      let formula: string | undefined
      if (formulaNode) {
        const text = textOf(c.f)
        const si = attr(formulaNode, 'si')
        if (attr(formulaNode, 't') === 'shared' && si !== undefined) {
          if (text) {
            shared.set(si, { row: at.row, column: at.column, formula: `=${text}` })
            formula = `=${text}`
          } else pending.push({ i, j, si, row: at.row, column: at.column })
        } else if (text) formula = `=${text}`
      }
      let value: string | number | boolean | null = null
      let valueType: number | null = null
      if (type === 's') {
        value = strings[Number(rawValue)] ?? ''
        valueType = 1
      } else if (type === 'inlineStr') {
        value = stringItemText(asNode(c.is))
        valueType = 1
      } else if (type === 'str' || type === 'e') {
        value = rawValue
        valueType = 1
      } else if (type === 'b') {
        value = rawValue === '1' ? 1 : 0
        valueType = 3
      } else if (rawValue !== '') {
        const n = Number(rawValue)
        if (Number.isFinite(n)) {
          value = n
          valueType = 2
        }
      }
      const xf = styles.xfs[Number(attr(c, 's') ?? 0)]
      const style = xf && Number(attr(c, 's') ?? 0) !== 0 ? options.toStyle(xf) : null
      if (value === null && !formula && !style) continue
      cells[i]![j] = {
        ...(value === null ? {} : { v: value }),
        ...(valueType === null ? {} : { t: valueType }),
        ...(formula ? { f: formula } : {}),
        s: (style ?? null) as Exclude<CopiedCell['s'], undefined>,
      }
    }
  }
  // followers of a shared formula carry no text: move the master's by their distance
  for (const follower of pending) {
    const master = shared.get(follower.si)
    const cell = cells[follower.i]![follower.j]
    if (!master || !cell) continue
    cells[follower.i]![follower.j] = {
      ...cell,
      f: offsetFormulaRefs(
        master.formula,
        follower.row - master.row,
        follower.column - master.column,
      ),
    }
  }

  const merges: IRange[] = []
  for (const merge of list(asNode(sheetRoot.mergeCells).mergeCell)) {
    const area = parseArea(attr(merge, 'ref') ?? '')
    if (
      area &&
      area.startRow >= range.startRow &&
      area.endRow <= range.endRow &&
      area.startColumn >= range.startColumn &&
      area.endColumn <= range.endColumn
    ) {
      merges.push(area)
    }
  }
  const columnWidths: (number | null)[] = columns.map(() => null)
  for (const col of list(asNode(sheetRoot.cols).col)) {
    const min = Number(attr(col, 'min')) - 1
    const max = Number(attr(col, 'max')) - 1
    const width = Number(attr(col, 'width'))
    if (!Number.isFinite(width) || width <= 0) continue
    for (
      let column = Math.max(min, range.startColumn);
      column <= Math.min(max, range.endColumn);
      column += 1
    ) {
      columnWidths[column - range.startColumn] = Math.trunc((width * 7 + 5) * 1.0)
    }
  }
  const links: PayloadLink[] = []
  const sheetRelsPath = sheetPath?.replace(/worksheets\/([^/]+)$/, 'worksheets/_rels/$1.rels')
  const sheetRelsXml = sheetRelsPath ? await read(sheetRelsPath) : undefined
  const sheetRels = sheetRelsXml
    ? list(asNode(asNode(parser.parse(sheetRelsXml)).Relationships).Relationship)
    : []
  for (const hyperlink of list(asNode(sheetRoot.hyperlinks).hyperlink)) {
    const area = parseArea(attr(hyperlink, 'ref') ?? '')
    if (!area) continue
    const rel = sheetRels.find((r) => attr(r, 'Id') === attr(hyperlink, 'id'))
    const location = attr(hyperlink, 'location')
    const target = rel ? attr(rel, 'Target') : location ? `#${location}` : undefined
    if (target) links.push({ row: area.startRow, column: area.startColumn, target })
  }
  return {
    origin: 'excel',
    copied: { unitId: '', sheetId: '', cut: false, rows, columns, cells },
    sheetName: link.sheetName,
    merges,
    columnWidths,
    rowHeights,
    notes: [],
    validations: [],
    conditionalFormats: [],
    links,
    literalFormulas: false,
  }
}
