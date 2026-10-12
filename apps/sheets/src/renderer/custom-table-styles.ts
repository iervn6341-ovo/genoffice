/**
 * Table Design → New Table Style: the table styles a user makes. Excel keeps one in the workbook's
 * styles.xml (`<tableStyle>` with one `<dxf>` per formatted element) and names it by its name alone;
 * a table refers to it by that name. Here the definitions are held by name for the open workbook —
 * they draw the tables of this session and are written when a table that uses one is saved.
 */
import type { WorkbookTableStyleAdd } from '../shared/desktop-api'
import type { BorderDef, BorderLine, ElementDef, TableStyleDef } from './table-style-defs'
import { setCustomTableStyleDef, setCustomTableStyleDefs, TABLE_STYLE_NAMES } from './table-styles'

export const CUSTOM_ELEMENT_KEYS = [
  'wholeTable',
  'headerRow',
  'totalRow',
  'firstColumn',
  'lastColumn',
  'firstRowStripe',
  'secondRowStripe',
  'firstColumnStripe',
  'secondColumnStripe',
] as const
export type CustomElementKey = (typeof CUSTOM_ELEMENT_KEYS)[number]

export const EDGE_KEYS = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const
export type EdgeKey = (typeof EDGE_KEYS)[number]

/** the edges an element can have: the whole table has the outline and the lines inside, the rest only their own sides */
export const EDGES_OF: Record<CustomElementKey, readonly EdgeKey[]> = {
  wholeTable: ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'],
  headerRow: ['top', 'bottom'],
  totalRow: ['top', 'bottom'],
  firstColumn: ['left', 'right'],
  lastColumn: ['left', 'right'],
  firstRowStripe: ['top', 'bottom'],
  secondRowStripe: ['top', 'bottom'],
  firstColumnStripe: ['left', 'right'],
  secondColumnStripe: ['left', 'right'],
}

export interface CustomElement {
  /** #RRGGBB */
  readonly fill?: string | undefined
  readonly font?: string | undefined
  readonly bold?: boolean | undefined
  /** the line the edges below are drawn with */
  readonly line?: BorderLine | undefined
  readonly lineColor?: string | undefined
  readonly edges?: readonly EdgeKey[] | undefined
}

export interface CustomTableStyle {
  readonly name: string
  readonly elements: Partial<Record<CustomElementKey, CustomElement>>
}

const HEX = /^#[0-9A-Fa-f]{6}$/

/** the element as the table-style engine draws it (see table-style-defs.ts) */
function elementDef(element: CustomElement): ElementDef {
  const def: { -readonly [K in keyof ElementDef]: ElementDef[K] } = {}
  if (element.fill && HEX.test(element.fill)) def.fill = { rgb: element.fill }
  if (element.font && HEX.test(element.font)) def.font = { rgb: element.font }
  if (element.bold) def.bold = true
  if (element.edges?.length && element.line) {
    const border: BorderDef = {
      l: element.line,
      c: { rgb: element.lineColor && HEX.test(element.lineColor) ? element.lineColor : '#000000' },
    }
    for (const edge of element.edges) def[edge] = border
  }
  return def
}

export function customStyleDef(style: CustomTableStyle): TableStyleDef {
  const e = style.elements
  const def: { -readonly [K in keyof TableStyleDef]: TableStyleDef[K] } = {}
  const set = (key: keyof TableStyleDef, element: CustomElement | undefined): void => {
    if (element) def[key] = elementDef(element)
  }
  set('whole', e.wholeTable)
  set('header', e.headerRow)
  set('total', e.totalRow)
  set('firstCol', e.firstColumn)
  set('lastCol', e.lastColumn)
  set('stripe1', e.firstRowStripe)
  set('stripe2', e.secondRowStripe)
  set('colStripe1', e.firstColumnStripe)
  set('colStripe2', e.secondColumnStripe)
  return def
}

export type StyleNameProblem = 'empty' | 'invalid' | 'taken' | 'builtin'

/** a name Excel accepts for a new style: not empty, not one of the built-ins (they all start with "TableStyle"), not in use */
export function styleNameProblem(name: string, taken: readonly string[]): StyleNameProblem | null {
  const trimmed = name.trim()
  if (trimmed === '') return 'empty'
  if (trimmed.length > 255 || /[\u0000-\u001f<>"]/.test(trimmed)) return 'invalid'
  if (/^TableStyle(?:Light|Medium|Dark)\d+$/i.test(trimmed) || TABLE_STYLE_NAMES.has(trimmed)) {
    return 'builtin'
  }
  if (taken.some((other) => other.toLowerCase() === trimmed.toLowerCase())) return 'taken'
  return null
}

/** the first free "Table Style 1", "Table Style 2" … (Excel's default for a new style) */
export function nextStyleName(base: string, taken: readonly string[]): string {
  for (let n = 1; ; n += 1) {
    const candidate = `${base} ${n}`
    if (!taken.some((other) => other.toLowerCase() === candidate.toLowerCase())) return candidate
  }
}

/** the style as the save request carries it */
export function toWireStyle(style: CustomTableStyle): WorkbookTableStyleAdd {
  return {
    name: style.name,
    elements: CUSTOM_ELEMENT_KEYS.flatMap((type) => {
      const element = style.elements[type]
      if (!element) return []
      const edges = element.edges?.filter((edge) => EDGES_OF[type].includes(edge)) ?? []
      const hasLine = edges.length > 0 && element.line !== undefined
      const entry = {
        type,
        ...(element.fill && HEX.test(element.fill) ? { fill: element.fill.toUpperCase() } : {}),
        ...(element.font && HEX.test(element.font) ? { font: element.font.toUpperCase() } : {}),
        ...(element.bold ? { bold: true } : {}),
        ...(hasLine
          ? {
              line: element.line!,
              lineColor: (element.lineColor && HEX.test(element.lineColor)
                ? element.lineColor
                : '#000000'
              ).toUpperCase(),
              edges,
            }
          : {}),
      }
      return Object.keys(entry).length > 1 ? [entry] : []
    }),
  }
}

// ── the workbook's custom styles ───────────────────────────────────────────────────────────────

const styles = new Map<string, CustomTableStyle>()
/// the file the styles belong to: opening another file drops them, saving this one (which reopens it) keeps them
let belongsTo: string | null = null

export function listCustomTableStyles(): CustomTableStyle[] {
  return [...styles.values()]
}

export function customTableStyle(name: string): CustomTableStyle | undefined {
  return styles.get(name)
}

export function addCustomTableStyle(style: CustomTableStyle): void {
  styles.set(style.name, style)
  setCustomTableStyleDef(style.name, customStyleDef(style))
}

/// the names of the styles the open file itself defines and its tables wear (a corporate template);
/// their definitions are not read here, but a new style must not take their name
let fileStyleNames: readonly string[] = []

/// every name a new style cannot have: the custom ones of this workbook and the ones the file defines
export function takenTableStyleNames(): string[] {
  return [...listCustomTableStyles().map((style) => style.name), ...fileStyleNames]
}

/** a workbook was opened: the styles of another file do not carry over */
export function openCustomTableStylesFor(
  path: string | null | undefined,
  definedInFile: readonly string[] = [],
): void {
  const key = path ?? null
  // a name outside the built-in set that a table of the file wears is the file's own style
  fileStyleNames = [...new Set(definedInFile)].filter((name) => !TABLE_STYLE_NAMES.has(name))
  if (key === belongsTo) return
  styles.clear()
  setCustomTableStyleDefs(new Map())
  belongsTo = key
}

/** a save wrote the workbook to `path` (and reopens it from there) */
export function saveCustomTableStylesTo(path: string | null | undefined): void {
  belongsTo = path ?? null
}

/** the styles a save must write: the custom ones the session's tables use */
export function customStylesUsedBy(
  tableStyleNames: readonly (string | undefined)[],
): CustomTableStyle[] {
  const used = new Set(tableStyleNames.filter((name): name is string => name !== undefined))
  return listCustomTableStyles().filter((style) => used.has(style.name))
}
