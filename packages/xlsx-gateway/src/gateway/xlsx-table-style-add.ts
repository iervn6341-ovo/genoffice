import { StylesheetEditor } from './xlsx-styles'
import type { MutablePackage } from './xlsx-drawing-add'

/// Persists the table styles made in the editor (Table Design → New Table Style): one <dxf> per
/// formatted element in xl/styles.xml's <dxfs>, and a <tableStyle> naming them under <tableStyles>
/// (CT_TableStyle: wholeTable, headerRow, … each pointing at its dxf). A table refers to the style by
/// name alone. A style the file already has under the name is left as it is.

export class TableStyleAddError extends Error {}

export type TableStyleElementType =
  | 'wholeTable'
  | 'headerRow'
  | 'totalRow'
  | 'firstColumn'
  | 'lastColumn'
  | 'firstRowStripe'
  | 'secondRowStripe'
  | 'firstColumnStripe'
  | 'secondColumnStripe'

export type TableStyleLine = 'thin' | 'medium' | 'thick' | 'double'
export type TableStyleEdge = 'top' | 'bottom' | 'left' | 'right' | 'insideH' | 'insideV'

export interface TableStyleElementAdd {
  readonly type: TableStyleElementType
  /// #RRGGBB
  readonly fill?: string | undefined
  readonly font?: string | undefined
  readonly bold?: boolean | undefined
  readonly line?: TableStyleLine | undefined
  readonly lineColor?: string | undefined
  readonly edges?: readonly TableStyleEdge[] | undefined
}

export interface TableStyleAddition {
  readonly name: string
  readonly elements: readonly TableStyleElementAdd[]
}

const HEX = /^#[0-9A-Fa-f]{6}$/
const ELEMENT_TYPES: ReadonlySet<string> = new Set<TableStyleElementType>([
  'wholeTable',
  'headerRow',
  'totalRow',
  'firstColumn',
  'lastColumn',
  'firstRowStripe',
  'secondRowStripe',
  'firstColumnStripe',
  'secondColumnStripe',
])

const argb = (hex: string): string => `FF${hex.slice(1).toUpperCase()}`

const escapeAttribute = (input: string): string =>
  input.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

/// the <dxf> of one element: its font (bold, color), fill and borders (CT_Dxf order: font, fill, border)
export function buildTableStyleDxf(element: TableStyleElementAdd): string {
  const fontParts: string[] = []
  if (element.bold) fontParts.push('<b/>')
  if (element.font) fontParts.push(`<color rgb="${argb(element.font)}"/>`)
  const font = fontParts.length > 0 ? `<font>${fontParts.join('')}</font>` : ''
  const fill = element.fill
    ? `<fill><patternFill><bgColor rgb="${argb(element.fill)}"/></patternFill></fill>`
    : ''
  let border = ''
  if (element.line && element.edges && element.edges.length > 0) {
    const color = argb(element.lineColor ?? '#000000')
    const side = (tag: string, edge: TableStyleEdge): string =>
      element.edges!.includes(edge)
        ? `<${tag} style="${element.line}"><color rgb="${color}"/></${tag}>`
        : ''
    // CT_Border order: left, right, top, bottom, … vertical, horizontal
    const lines =
      side('left', 'left') +
      side('right', 'right') +
      side('top', 'top') +
      side('bottom', 'bottom') +
      side('vertical', 'insideV') +
      side('horizontal', 'insideH')
    if (lines !== '') border = `<border>${lines}</border>`
  }
  return `<dxf>${font}${fill}${border}</dxf>`
}

function validate(addition: TableStyleAddition): void {
  if (
    addition.name.trim() === '' ||
    addition.name.length > 255 ||
    /[\u0000-\u001f]/.test(addition.name)
  ) {
    throw new TableStyleAddError(`"${addition.name}" is not a name a table style can have.`)
  }
  if (/^TableStyle(?:Light|Medium|Dark)\d+$/i.test(addition.name)) {
    throw new TableStyleAddError(
      `"${addition.name}" is the name of one of Excel's own table styles.`,
    )
  }
  for (const element of addition.elements) {
    if (!ELEMENT_TYPES.has(element.type)) {
      throw new TableStyleAddError(`"${element.type}" is not a table style element.`)
    }
    for (const color of [element.fill, element.font, element.lineColor]) {
      if (color !== undefined && !HEX.test(color)) {
        throw new TableStyleAddError(`"${color}" is not a #RRGGBB color.`)
      }
    }
  }
}

export async function applyTableStyleAdditions(
  pkg: MutablePackage,
  additions: readonly TableStyleAddition[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (additions.length === 0) return
  const stylesPath = 'xl/styles.xml'
  if (!(await pkg.has(stylesPath))) {
    throw new TableStyleAddError('The workbook has no stylesheet to hold a table style.')
  }
  const editor = new StylesheetEditor(await pkg.readText(stylesPath))
  for (const addition of additions) {
    validate(addition)
    editor.addTableStyle(
      addition.name,
      addition.elements
        .map((element) => ({ type: element.type, dxfXml: buildTableStyleDxf(element) }))
        .filter((entry) => entry.dxfXml !== '<dxf></dxf>'),
    )
  }
  if (editor.changed) {
    pkg.write(stylesPath, editor.serialize())
    touchedEntries.add(stylesPath)
  }
}

export { escapeAttribute as escapeTableStyleName }
