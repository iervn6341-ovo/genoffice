/**
 * Table markup from Excel (and from this app's own HTML flavor) → a `ForeignCopy`.
 *
 * Univer's HTML paste reads the visible text and a handful of CSS properties. Excel's markup
 * carries much more, and the browser drops what it does not know: the cell's real value
 * (`x:num`, `x:str`, `x:bool`), the formula (`x:fmla`), the number format (`mso-number-format`,
 * inline or through a class), `border:` shorthands, widths in points. This parser reads those
 * from the markup text itself and produces the same cell model a local copy has.
 */
import type { IRange, IStyleData } from '@univerjs/core'

import type { ForeignCopy } from './clipboard-foreign'
import type { PayloadLink } from './clipboard-payload'
import type { CopiedCell } from './paste-source'
import { typeField } from './clipboard-text'

const NUMBER = 2
const STRING = 1
const BOOLEAN = 3

/// declarations of one CSS block (`a:b; c:d`), property names lower-cased
export function parseCssDeclarations(text: string): Record<string, string> {
  const out: Record<string, string> = {}
  let depth = 0
  let quote = ''
  let start = 0
  const flush = (end: number): void => {
    const part = text.slice(start, end)
    const colon = part.indexOf(':')
    if (colon > 0) {
      const name = part.slice(0, colon).trim().toLowerCase()
      const value = part.slice(colon + 1).trim()
      if (name) out[name] = value
    }
  }
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i]!
    if (quote) {
      if (ch === '\\') i += 1
      else if (ch === quote) quote = ''
    } else if (ch === '"' || ch === "'") quote = ch
    else if (ch === '(') depth += 1
    else if (ch === ')') depth = Math.max(0, depth - 1)
    else if (ch === ';' && depth === 0) {
      flush(i)
      start = i + 1
    }
  }
  flush(text.length)
  return out
}

/// `.xl65 { … }`, `td { … }` blocks of a <style> element (comments and @rules ignored)
export function parseStyleSheet(css: string): Map<string, Record<string, string>> {
  const rules = new Map<string, Record<string, string>>()
  const clean = css.replace(/<!--|-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '')
  const re = /([^{}@]+)\{([^{}]*)\}/g
  let match: RegExpExecArray | null
  while ((match = re.exec(clean)) !== null) {
    const declarations = parseCssDeclarations(match[2]!)
    for (const raw of match[1]!.split(',')) {
      const selector = raw.trim()
      if (!selector) continue
      rules.set(selector, { ...(rules.get(selector) ?? {}), ...declarations })
    }
  }
  return rules
}

const NAMED_COLORS: Record<string, string> = {
  black: '#000000',
  white: '#FFFFFF',
  red: '#FF0000',
  lime: '#00FF00',
  blue: '#0000FF',
  yellow: '#FFFF00',
  fuchsia: '#FF00FF',
  magenta: '#FF00FF',
  aqua: '#00FFFF',
  cyan: '#00FFFF',
  maroon: '#800000',
  green: '#008000',
  navy: '#000080',
  olive: '#808000',
  purple: '#800080',
  teal: '#008080',
  silver: '#C0C0C0',
  gray: '#808080',
  grey: '#808080',
  orange: '#FFA500',
  windowtext: '#000000',
  window: '#FFFFFF',
}

/// a CSS color as #RRGGBB; null for none / unknown / transparent
export function cssColor(value: string | undefined): string | null {
  if (!value) return null
  const text = value.trim().toLowerCase()
  if (!text || text === 'none' || text === 'transparent' || text === 'auto') return null
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(text)
  if (hex) {
    const digits = hex[1]!
    const full =
      digits.length === 3
        ? digits
            .split('')
            .map((c) => c + c)
            .join('')
        : digits
    return `#${full.toUpperCase()}`
  }
  const rgb = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(text)
  if (rgb) {
    return (
      '#' +
      [rgb[1]!, rgb[2]!, rgb[3]!]
        .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase()
    )
  }
  return NAMED_COLORS[text] ?? null
}

/// CSS length → pixels (96 dpi)
export function cssLengthToPx(value: string | undefined): number | null {
  if (!value) return null
  const match = /^(-?\d*\.?\d+)\s*(px|pt|pc|in|cm|mm)?$/.exec(value.trim().toLowerCase())
  if (!match) return null
  const n = Number(match[1])
  switch (match[2] ?? 'px') {
    case 'pt':
      return (n * 96) / 72
    case 'pc':
      return n * 16
    case 'in':
      return n * 96
    case 'cm':
      return (n * 96) / 2.54
    case 'mm':
      return (n * 96) / 25.4
    default:
      return n
  }
}

/// the value of an `mso-number-format` declaration as an Excel pattern
export function msoNumberFormat(value: string | undefined): string | null {
  if (!value) return null
  let text = value.trim()
  if (text.startsWith('"') && text.endsWith('"')) text = text.slice(1, -1)
  // \0022 style code points, then \\ (a literal backslash), then \x (the character x)
  text = text.replace(/\\([0-9a-fA-F]{4})/g, (_m, hex: string) =>
    String.fromCharCode(Number.parseInt(hex, 16)),
  )
  text = text
    .replace(/\\\\/g, '\uE000')
    .replace(/\\(.)/g, '$1')
    .replace(/\uE000/g, '\\')
  switch (text.toLowerCase()) {
    case '':
    case 'general':
    case 'general number':
      return null
    case 'general date':
      return 'm/d/yyyy h:mm'
    case 'long date':
      return 'dddd, mmmm d, yyyy'
    case 'medium date':
      return 'd-mmm-yy'
    case 'short date':
      return 'm/d/yyyy'
    case 'long time':
      return 'h:mm:ss AM/PM'
    case 'medium time':
      return 'h:mm AM/PM'
    case 'short time':
      return 'h:mm'
    case 'currency':
      return '"$"#,##0.00'
    case 'euro currency':
      return '[$€-2] #,##0.00'
    case 'fixed':
      return '0.00'
    case 'standard':
      return '#,##0.00'
    case 'percent':
      return '0.00%'
    case 'scientific':
      return '0.00E+00'
    case 'yes/no':
      return '"Yes";"Yes";"No"'
    case 'true/false':
      return '"True";"True";"False"'
    case 'on/off':
      return '"On";"On";"Off"'
    default:
      return text
  }
}

/// one border side: CSS (`.5pt solid windowtext`) → Univer border {s, cl}
function borderSide(value: string | undefined): { s: number; cl: { rgb: string } } | null {
  if (!value) return null
  const text = value.trim().toLowerCase()
  if (!text || text === 'none' || text.startsWith('none ') || text === '0') return null
  const tokens = text.split(/\s+(?![^(]*\))/)
  let widthPx: number | null = null
  let kind = 'solid'
  let color: string | null = null
  for (const token of tokens) {
    const px = cssLengthToPx(token)
    if (px !== null) widthPx = px
    else if (
      /^(solid|dotted|dashed|dot-dash|dot-dot-dash|double|hairline|groove|ridge|inset|outset)$/.test(
        token,
      )
    ) {
      kind = token
    } else if (token === 'thin') widthPx = 1
    else if (token === 'medium') widthPx = 2
    else if (token === 'thick') widthPx = 3
    else {
      const c = cssColor(token)
      if (c) color = c
    }
  }
  const px = widthPx ?? 1
  if (px <= 0) return null
  // Excel writes thin as .5pt, medium as 1.0pt and thick as 1.5pt and more
  const level = px <= 1.1 ? 0 : px <= 1.7 ? 1 : 2
  let s: number
  if (kind === 'double') s = 7
  else if (kind === 'hairline') s = 2
  else if (kind === 'dot-dash') s = level === 0 ? 5 : 10
  else if (kind === 'dot-dot-dash') s = level === 0 ? 6 : 11
  else if (kind === 'dotted') s = level === 0 ? 3 : 11
  else if (kind === 'dashed') s = level === 0 ? 4 : 9
  else s = level === 0 ? 1 : level === 1 ? 8 : 13
  return { s, cl: { rgb: color ?? '#000000' } }
}

function fontFamily(value: string | undefined): string | null {
  if (!value) return null
  const first = value
    .split(',')[0]!
    .trim()
    .replace(/^["']|["']$/g, '')
  if (!first || /^(sans-serif|serif|monospace)$/i.test(first)) return null
  return first
}

/// the cell style a declaration set stands for
export function styleFromDeclarations(d: Record<string, string>): IStyleData | null {
  const style: Record<string, unknown> = {}
  const family = fontFamily(d['font-family'])
  if (family) style.ff = family
  const size = d['font-size']
  if (size) {
    const px = cssLengthToPx(size)
    if (px !== null && px > 0) style.fs = Math.round(((px * 72) / 96) * 2) / 2
  }
  const weight = d['font-weight']
  if (weight && (weight === 'bold' || weight === 'bolder' || Number(weight) >= 600)) style.bl = 1
  if (d['font-style'] === 'italic' || d['font-style'] === 'oblique') style.it = 1
  const decoration = `${d['text-decoration'] ?? ''} ${d['text-decoration-line'] ?? ''}`
  if (/underline/.test(decoration)) {
    const double = /double/.test(d['text-underline-style'] ?? '')
    style.ul = double ? { s: 1, t: 2 } : { s: 1 }
  }
  if (/line-through/.test(decoration)) style.st = { s: 1 }
  const color = cssColor(d.color)
  if (color && color !== '#000000') style.cl = { rgb: color }
  const bg = cssColor(d['background-color'] ?? d.background ?? d['mso-pattern']?.split(' ')[0])
  const backgroundText = (d['background-color'] ?? d.background ?? '').toLowerCase()
  if (bg && !/^(none|transparent)/.test(backgroundText)) style.bg = { rgb: bg }
  const align = d['text-align']
  if (align === 'left') style.ht = 1
  else if (align === 'center') style.ht = 2
  else if (align === 'right') style.ht = 3
  else if (align === 'justify') style.ht = 4
  const vertical = d['vertical-align']
  if (vertical === 'top') style.vt = 1
  else if (vertical === 'middle' || vertical === 'center') style.vt = 2
  // bottom is the default of every cell: Excel's markup spells it out for all of them
  const whiteSpace = d['white-space']
  if (whiteSpace === 'normal' || whiteSpace === 'pre-wrap' || whiteSpace === 'pre-line') {
    style.tb = 3
  }
  const rotate = d['mso-rotate']
  if (rotate && Number(rotate) !== 0 && Number.isFinite(Number(rotate))) {
    style.tr = { a: Number(rotate), v: 0 }
  }
  const pattern = msoNumberFormat(d['mso-number-format'])
  if (pattern) style.n = { pattern }
  const shorthand = d.border
  const bd: Record<string, unknown> = {}
  for (const [key, prop] of [
    ['t', 'border-top'],
    ['r', 'border-right'],
    ['b', 'border-bottom'],
    ['l', 'border-left'],
  ] as const) {
    const side = borderSide(d[prop] ?? shorthand)
    if (side) bd[key] = side
  }
  if (Object.keys(bd).length > 0) style.bd = bd
  return Object.keys(style).length > 0 ? (style as IStyleData) : null
}

interface ParsedCell {
  readonly value: string | number | boolean | null
  readonly t: number | null
  readonly formula: string | null
  readonly style: IStyleData | null
  readonly link: string | null
}

const decodeEntities = (text: string): string =>
  text
    .replace(/&nbsp;/gi, '\u00a0')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_m, n: string) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_m, n: string) => String.fromCodePoint(Number.parseInt(n, 16)))
    .replace(/&amp;/gi, '&')

/// The text a cell shows: line breaks only where the markup has <br>, runs of white space as one
/// space, no-break spaces kept as spaces.
export function cellText(element: Element): string {
  let out = ''
  const walk = (node: Node): void => {
    if (node.nodeType === 3) {
      out += (node.textContent ?? '').replace(/[\r\n\t ]+/g, ' ')
      return
    }
    if (node.nodeType !== 1) return
    const el = node as Element
    const tag = el.tagName.toLowerCase()
    if (tag === 'br') {
      out += '\n'
      return
    }
    if (tag === 'style' || tag === 'script') return
    // Excel pads accounting formats with mso-spacerun runs of no-break spaces
    if (/mso-spacerun:\s*yes/i.test(el.getAttribute('style') ?? '')) return
    el.childNodes.forEach(walk)
    if (tag === 'p' || tag === 'div') out += '\n'
  }
  element.childNodes.forEach(walk)
  out = out.replace(/\u00a0/g, ' ')
  // markup indentation around the text is not content
  out = out
    .replace(/^ +| +$/g, '')
    .replace(/ *\n */g, '\n')
    .replace(/\n+$/, (m) => (m.length > 1 ? '' : ''))
  return out
}

const attr = (el: Element, name: string): string | null => {
  const direct = el.getAttribute(name)
  if (direct !== null) return direct
  // case-insensitive lookup for x:num / X:Num
  const lower = name.toLowerCase()
  for (const a of Array.from(el.attributes)) {
    if (a.name.toLowerCase() === lower) return a.value
  }
  return null
}

function parseCell(
  td: Element,
  sheet: Map<string, Record<string, string>>,
  inheritRow: Record<string, string>,
): ParsedCell {
  const declarations: Record<string, string> = { ...(sheet.get('td') ?? {}), ...inheritRow }
  const classes = (td.getAttribute('class') ?? '').split(/\s+/).filter(Boolean)
  for (const name of classes) Object.assign(declarations, sheet.get(`.${name}`) ?? {})
  const id = td.getAttribute('id')
  if (id) Object.assign(declarations, sheet.get(`#${id}`) ?? {})
  // old attributes
  const align = td.getAttribute('align')
  if (align && !declarations['text-align']) declarations['text-align'] = align
  const valign = td.getAttribute('valign')
  if (valign && !declarations['vertical-align']) declarations['vertical-align'] = valign
  const bgcolor = td.getAttribute('bgcolor')
  if (bgcolor && !declarations.background) declarations.background = bgcolor
  Object.assign(declarations, parseCssDeclarations(td.getAttribute('style') ?? ''))
  // a cell whose inner element carries the font (<b>, <font>, <span style>) is the cell's font
  // when it wraps all of the text
  const sole =
    td.children.length === 1 && td.textContent?.trim() === td.children[0]!.textContent?.trim()
  if (sole) {
    let inner: Element | undefined = td.children[0]
    while (inner) {
      const tag = inner.tagName.toLowerCase()
      if (tag === 'b' || tag === 'strong') declarations['font-weight'] = 'bold'
      if (tag === 'i' || tag === 'em') declarations['font-style'] = 'italic'
      const decorate = (kind: string): void => {
        const have = declarations['text-decoration'] ?? ''
        if (!have.includes(kind)) declarations['text-decoration'] = `${have} ${kind}`.trim()
      }
      if (tag === 'u') decorate('underline')
      if (tag === 's' || tag === 'strike' || tag === 'del') decorate('line-through')
      const font = tag === 'font' ? inner.getAttribute('color') : null
      if (font) declarations.color = font
      Object.assign(declarations, parseCssDeclarations(inner.getAttribute('style') ?? ''))
      inner = inner.children.length === 1 ? inner.children[0] : undefined
    }
  }
  const style = styleFromDeclarations(declarations)

  const text = cellText(td)
  const anchor = td.querySelector('a[href]')
  const href = anchor?.getAttribute('href') ?? null
  const fmla = attr(td, 'x:fmla')
  const hasNum = attr(td, 'x:num') !== null
  const num = attr(td, 'x:num')
  const bool = attr(td, 'x:bool')
  const str = attr(td, 'x:str')
  const err = attr(td, 'x:err')
  const formula = fmla ? decodeEntities(fmla) : null

  let value: string | number | boolean | null
  let t: number | null
  if (err !== null) {
    value = err || text
    t = STRING
  } else if (bool !== null) {
    const raw = (bool || text).trim().toUpperCase()
    value = raw === 'TRUE' || raw === '1' ? 1 : 0
    t = BOOLEAN
  } else if (hasNum) {
    const n = num !== null && num !== '' ? Number(num) : Number(text.replace(/[,$%\s]/g, ''))
    if (Number.isFinite(n)) {
      value = n
      t = NUMBER
    } else {
      value = text
      t = STRING
    }
  } else if (str !== null) {
    // Excel writes x:str="'text" for text that looks like something else
    value = str === '' ? text : str.replace(/^'/, '')
    t = STRING
  } else if (text === '') {
    value = null
    t = null
  } else if (style?.n?.pattern === '@') {
    value = text
    t = STRING
  } else {
    const typed = typeField(text)
    if (typed.t === NUMBER || typed.t === BOOLEAN) {
      value = typed.v
      t = typed.t
      // a date / percent / currency typed from its text carries the format it was read with
      if (typed.pattern && !style?.n) {
        return {
          value,
          t,
          formula,
          style: { ...(style ?? {}), n: { pattern: typed.pattern } },
          link: href,
        }
      }
    } else {
      value = text
      t = STRING
    }
  }
  return { value, t, formula, style, link: href }
}

/// first number of an `x:` colgroup width etc.
function tableColumnWidths(
  table: Element,
  sheet: Map<string, Record<string, string>>,
): (number | null)[] {
  const widths: (number | null)[] = []
  for (const col of Array.from(table.querySelectorAll('col'))) {
    const span = Math.max(1, Number(col.getAttribute('span') ?? 1) || 1)
    const declared = parseCssDeclarations(col.getAttribute('style') ?? '')
    const px =
      cssLengthToPx(declared.width) ??
      (col.getAttribute('width') ? cssLengthToPx(col.getAttribute('width')!) : null) ??
      cssLengthToPx(sheet.get('col')?.width)
    for (let i = 0; i < span; i += 1) widths.push(px)
  }
  return widths
}

/// True when the markup is Excel's (or this app's own flavor of it).
export function looksLikeExcelHtml(html: string): boolean {
  return (
    /urn:schemas-microsoft-com:office:excel/i.test(html) ||
    /ProgId["']?\s*content=["']?Excel/i.test(html) ||
    /content=["']?Excel\.Sheet/i.test(html) ||
    /\bx:(num|str|fmla|bool)\b/i.test(html) ||
    /mso-(number-format|pattern|rotate|height-source|width-source)\s*:/i.test(html) ||
    /class=["']?xl\d+/i.test(html) ||
    /GenOffice Sheets/i.test(html)
  )
}

/// Parses clipboard HTML; null when it is not Excel-flavored table markup (the caller then
/// leaves the paste to Univer's own HTML reader).
export function foreignFromExcelHtml(html: string): ForeignCopy | null {
  if (!looksLikeExcelHtml(html)) return null
  const doc = new DOMParser().parseFromString(html, 'text/html')
  const table = doc.querySelector('table')
  if (!table) return null
  const sheet = new Map<string, Record<string, string>>()
  for (const style of Array.from(doc.querySelectorAll('style'))) {
    for (const [selector, declarations] of parseStyleSheet(style.textContent ?? '')) {
      sheet.set(selector, { ...(sheet.get(selector) ?? {}), ...declarations })
    }
  }
  const rowsEl = Array.from(table.querySelectorAll('tr'))
  const grid: (CopiedCell | null)[][] = []
  const occupied = new Set<string>()
  const merges: IRange[] = []
  const links: PayloadLink[] = []
  const rowHeights: (number | null)[] = []
  let columnCount = 0
  rowsEl.forEach((tr, r) => {
    const rowDeclared = parseCssDeclarations(tr.getAttribute('style') ?? '')
    const heightPx =
      cssLengthToPx(rowDeclared.height) ??
      (tr.getAttribute('height') ? cssLengthToPx(tr.getAttribute('height')!) : null)
    rowHeights[r] = heightPx
    grid[r] ??= []
    let c = 0
    for (const td of Array.from(tr.children)) {
      const tag = td.tagName.toLowerCase()
      if (tag !== 'td' && tag !== 'th') continue
      while (occupied.has(`${r}:${c}`)) c += 1
      const rowSpan = Math.max(1, Number(td.getAttribute('rowspan') ?? 1) || 1)
      const colSpan = Math.max(1, Number(td.getAttribute('colspan') ?? 1) || 1)
      const parsed = parseCell(td, sheet, {})
      const cell: CopiedCell = {
        ...(parsed.formula
          ? { f: parsed.formula }
          : parsed.value === null
            ? {}
            : { v: parsed.value }),
        ...(parsed.formula && parsed.value !== null ? { v: parsed.value } : {}),
        ...(parsed.t === null ? {} : { t: parsed.t }),
        s: parsed.style,
      }
      const blank = parsed.value === null && !parsed.formula && !parsed.style
      for (let dr = 0; dr < rowSpan; dr += 1) {
        for (let dc = 0; dc < colSpan; dc += 1) {
          occupied.add(`${r + dr}:${c + dc}`)
          grid[r + dr] ??= []
          // covered cells of a merge keep the style (borders / fill) but no value
          grid[r + dr]![c + dc] =
            dr === 0 && dc === 0 ? (blank ? null : cell) : parsed.style ? { s: parsed.style } : null
        }
      }
      if (rowSpan > 1 || colSpan > 1) {
        merges.push({
          startRow: r,
          endRow: r + rowSpan - 1,
          startColumn: c,
          endColumn: c + colSpan - 1,
        })
      }
      if (parsed.link) links.push({ row: r, column: c, target: parsed.link })
      c += colSpan
      columnCount = Math.max(columnCount, c)
    }
  })
  const rowCount = grid.length
  if (rowCount === 0 || columnCount === 0) return null
  const cells = Array.from({ length: rowCount }, (_, i) =>
    Array.from({ length: columnCount }, (_, j) => grid[i]?.[j] ?? null),
  )
  const widths = tableColumnWidths(table, sheet)
  return {
    origin: 'excel',
    copied: {
      unitId: '',
      sheetId: '',
      cut: false,
      rows: Array.from({ length: rowCount }, (_, i) => i),
      columns: Array.from({ length: columnCount }, (_, j) => j),
      cells,
    },
    sheetName: '',
    merges,
    columnWidths: Array.from({ length: columnCount }, (_, j) => widths[j] ?? null),
    rowHeights: Array.from({ length: rowCount }, (_, i) => rowHeights[i] ?? null),
    notes: [],
    validations: [],
    conditionalFormats: [],
    links,
    literalFormulas: true,
  }
}
