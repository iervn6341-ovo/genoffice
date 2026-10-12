import { type MutablePackage } from './xlsx-drawing-add'

/// Table Design on a table the file already has: the table part is patched in place — its
/// <tableStyleInfo> (the style by name and the four Table Style Options). The part's columns,
/// filter, totals row and everything else stay as the file has them. The table is found by its
/// identity: displayName (the structured-reference name, unique in a workbook), or @name when the
/// part has no displayName — what the reader hands the renderer as the table's name.

export class TableEditError extends Error {}

export interface TableEdit {
  /// the table's displayName
  readonly name: string
  /// a built-in style name, or the name of a custom one; absent = the style stays
  readonly style?: string | undefined
  readonly bandedRows?: boolean | undefined
  readonly bandedColumns?: boolean | undefined
  readonly firstColumn?: boolean | undefined
  readonly lastColumn?: boolean | undefined
}

const escapeAttribute = (input: string): string =>
  input.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

const decodeAttribute = (input: string): string =>
  input
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')

/// the <tableStyleInfo> element, empty (`<… />`) or written out (`<…></…>`)
const STYLE_INFO = /<tableStyleInfo\b[^>]*?(?:\/>|>\s*<\/tableStyleInfo\s*>)/

/// an attribute of an element's text, double- or single-quoted
function attributeOf(element: string, attribute: string): string | undefined {
  const match = new RegExp(`\\b${attribute}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(element)
  return match?.[1] ?? match?.[2]
}

/// the opening tag of the <table> element
const tableTag = (xml: string): string => /<table\b[^>]*>/.exec(xml)?.[0] ?? ''

/// the identity of a table part: its displayName, else its name, in lower case; null when it has neither
export function tableKeyOf(xml: string): string | null {
  const tag = tableTag(xml)
  const own = attributeOf(tag, 'displayName') ?? attributeOf(tag, 'name')
  return own === undefined ? null : decodeAttribute(own).toLowerCase()
}

/// the <tableStyleInfo> of a table part with `edit` applied; an absent flag keeps the file's
export function patchTableStyleInfo(tableXml: string, edit: TableEdit): string {
  const existing = STYLE_INFO.exec(tableXml)?.[0]
  const read = (attribute: string): string | undefined =>
    existing === undefined ? undefined : attributeOf(existing, attribute)
  const flag = (attribute: string, wanted: boolean | undefined): string => {
    if (wanted !== undefined) return wanted ? '1' : '0'
    const old = read(attribute)
    return old === undefined ? '0' : /^(1|true)$/i.test(old) ? '1' : '0'
  }
  const kept = read('name')
  const name = edit.style ?? (kept === undefined ? undefined : decodeAttribute(kept))
  const style = name === undefined ? undefined : escapeAttribute(name)
  const info =
    '<tableStyleInfo' +
    (style === undefined ? '' : ` name="${style}"`) +
    ` showFirstColumn="${flag('showFirstColumn', edit.firstColumn)}"` +
    ` showLastColumn="${flag('showLastColumn', edit.lastColumn)}"` +
    ` showRowStripes="${flag('showRowStripes', edit.bandedRows)}"` +
    ` showColumnStripes="${flag('showColumnStripes', edit.bandedColumns)}"/>`
  if (existing !== undefined) return tableXml.replace(existing, () => info)
  // schema order: autoFilter, sortState, tableColumns, tableStyleInfo, extLst — right after the
  // columns (an extLst can also sit inside a column or the filter, so it is not searched for)
  const columnsEnd = tableXml.indexOf('</tableColumns>')
  if (columnsEnd >= 0) {
    const at = columnsEnd + '</tableColumns>'.length
    return tableXml.slice(0, at) + info + tableXml.slice(at)
  }
  const close = tableXml.lastIndexOf('</table>')
  if (close < 0) throw new TableEditError('The table part is malformed.')
  return tableXml.slice(0, close) + info + tableXml.slice(close)
}

export async function applyTableEdits(
  pkg: MutablePackage,
  edits: readonly TableEdit[],
  touchedEntries: Set<string>,
): Promise<void> {
  if (edits.length === 0) return
  const paths = (await pkg.paths()).filter(isTablePartPath)
  const byKey = new Map<string, string>()
  for (const path of paths) {
    const key = tableKeyOf(await pkg.readText(path))
    if (key !== null) byKey.set(key, path)
  }
  for (const edit of edits) {
    const path = byKey.get(edit.name.toLowerCase())
    if (!path) throw new TableEditError(`Table "${edit.name}" is not in the workbook.`)
    pkg.write(path, patchTableStyleInfo(await pkg.readText(path), edit))
    touchedEntries.add(path)
  }
}

export interface TableInfo {
  readonly style?: string
  readonly showFirstColumn: boolean
  readonly showLastColumn: boolean
  readonly showRowStripes: boolean
  readonly showColumnStripes: boolean
}

/// a table part of the package (xl/tables/tableN.xml)
export const isTablePartPath = (path: string): boolean => /^xl\/tables\/[^/]+\.xml$/.test(path)

/// what a table part says about its style
export function tableStyleInfoOf(xml: string): TableInfo {
  const tag = STYLE_INFO.exec(xml)?.[0] ?? ''
  const on = (key: string): boolean => /^(1|true)$/i.test(attributeOf(tag, key) ?? '')
  const style = attributeOf(tag, 'name')
  return {
    ...(style === undefined ? {} : { style: decodeAttribute(style) }),
    showFirstColumn: on('showFirstColumn'),
    showLastColumn: on('showLastColumn'),
    showRowStripes: on('showRowStripes'),
    showColumnStripes: on('showColumnStripes'),
  }
}

/// what table part `xml` says about its style, when it is the table called `name` (any case); null otherwise
export function tableInfoOf(xml: string, name: string): TableInfo | null {
  return tableKeyOf(xml) === name.toLowerCase() ? tableStyleInfoOf(xml) : null
}
