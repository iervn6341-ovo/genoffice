/**
 * Reading the look of HTML that comes from other programs (Word, Excel, web
 * pages) into the editor model: colors by name, highlight / shading, table
 * borders, fills, widths, paragraph alignment / indents / spacing, Word's
 * `mso-list` paragraphs, and `<style>` blocks whose class rules (Excel's
 * `.xl65`, Word's `p.MsoNormal`) are folded into inline styles first.
 *
 * Everything here works on parsed DOM / strings, never on layout, so it is the
 * same in the app and in tests.
 */

// ── CSS values ───────────────────────────────────────────────────────────

/** declarations of a style attribute / rule body, later ones overriding earlier */
export function parseDeclarations(text: string): Map<string, string> {
  const out = new Map<string, string>()
  let depth = 0
  let start = 0
  const parts: string[] = []
  for (let i = 0; i < text.length; i++) {
    const c = text[i]
    if (c === '(') depth++
    else if (c === ')') depth = Math.max(0, depth - 1)
    else if (c === ';' && depth === 0) {
      parts.push(text.slice(start, i))
      start = i + 1
    }
  }
  parts.push(text.slice(start))
  for (const part of parts) {
    const at = part.indexOf(':')
    if (at < 0) continue
    const name = part.slice(0, at).trim().toLowerCase()
    const value = part
      .slice(at + 1)
      .replace(/!important/i, '')
      .trim()
    if (name && value) out.set(name, value)
  }
  return out
}

export const declarationsOf = (el: Element): Map<string, string> =>
  parseDeclarations(el.getAttribute('style') ?? '')

const NAMED_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  lime: '00FF00',
  blue: '0000FF',
  yellow: 'FFFF00',
  aqua: '00FFFF',
  cyan: '00FFFF',
  fuchsia: 'FF00FF',
  magenta: 'FF00FF',
  silver: 'C0C0C0',
  gray: '808080',
  grey: '808080',
  maroon: '800000',
  olive: '808000',
  green: '008000',
  purple: '800080',
  teal: '008080',
  navy: '000080',
  orange: 'FFA500',
  pink: 'FFC0CB',
  brown: 'A52A2A',
  gold: 'FFD700',
  indigo: '4B0082',
  violet: 'EE82EE',
  crimson: 'DC143C',
  coral: 'FF7F50',
  salmon: 'FA8072',
  khaki: 'F0E68C',
  tan: 'D2B48C',
  beige: 'F5F5DC',
  ivory: 'FFFFF0',
  lavender: 'E6E6FA',
  turquoise: '40E0D0',
  skyblue: '87CEEB',
  lightblue: 'ADD8E6',
  lightgreen: '90EE90',
  lightgray: 'D3D3D3',
  lightgrey: 'D3D3D3',
  darkgray: 'A9A9A9',
  darkgrey: 'A9A9A9',
  darkblue: '00008B',
  darkgreen: '006400',
  darkred: '8B0000',
  lightyellow: 'FFFFE0',
  whitesmoke: 'F5F5F5',
  gainsboro: 'DCDCDC',
  orangered: 'FF4500',
  tomato: 'FF6347',
  steelblue: '4682B4',
  royalblue: '4169E1',
  dodgerblue: '1E90FF',
  slategray: '708090',
  slategrey: '708090',
}

/** `#abc`, `#aabbcc`, `rgb()/rgba()`, names → 'RRGGBB'; null for transparent / system colors / unknown */
export function cssColorToHex(value: string | undefined | null): string | null {
  if (!value) return null
  const v = value.trim().toLowerCase()
  if (!v || v === 'transparent' || v === 'windowtext' || v === 'window' || v === 'auto') return null
  if (v === 'inherit' || v === 'initial' || v === 'currentcolor') return null
  let m = /^#([0-9a-f]{6})$/.exec(v)
  if (m) return m[1]!.toUpperCase()
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v)
  if (m) return `${m[1]}${m[1]}${m[2]}${m[2]}${m[3]}${m[3]}`.toUpperCase()
  m = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+%?))?\s*\)$/.exec(v)
  if (m) {
    const alpha =
      m[4] === undefined ? 1 : m[4].endsWith('%') ? parseFloat(m[4]) / 100 : parseFloat(m[4])
    if (alpha === 0) return null
    return [m[1], m[2], m[3]]
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()
  }
  return NAMED_COLORS[v.replace(/[\s-]/g, '')] ?? null
}

/** length → points (pt/px/in/cm/mm/pc/q); null for % / auto / unknown. Bare numbers are px. */
export function lengthToPt(value: string | undefined | null): number | null {
  if (!value) return null
  const m = /^(-?\d*\.?\d+)\s*(pt|px|in|cm|mm|pc|q|em)?$/i.exec(value.trim())
  if (!m) return null
  const n = parseFloat(m[1]!)
  switch ((m[2] ?? 'px').toLowerCase()) {
    case 'pt':
      return n
    case 'px':
      return n * 0.75
    case 'in':
      return n * 72
    case 'cm':
      return (n * 72) / 2.54
    case 'mm':
      return (n * 72) / 25.4
    case 'pc':
      return n * 12
    case 'q':
      return (n * 72) / 101.6
    case 'em':
      return n * 12
    default:
      return null
  }
}

const twips = (pt: number): number => Math.round(pt * 20)

/** Word's mso-highlight names → OOXML w:highlight values */
const MSO_HIGHLIGHT: Record<string, string> = {
  yellow: 'yellow',
  lime: 'green',
  aqua: 'cyan',
  fuchsia: 'magenta',
  blue: 'blue',
  red: 'red',
  navy: 'darkBlue',
  teal: 'darkCyan',
  green: 'darkGreen',
  purple: 'darkMagenta',
  maroon: 'darkRed',
  olive: 'darkYellow',
  gray: 'darkGray',
  silver: 'lightGray',
  black: 'black',
  white: 'white',
}

// ── runs ─────────────────────────────────────────────────────────────────

/** highlight / shading / color of an inline element (color-name aware) */
export function foreignRunColors(el: Element): Record<string, string> {
  const d = declarationsOf(el)
  const out: Record<string, string> = {}
  const color = cssColorToHex(d.get('color'))
  if (color && color !== '000000') out.color = color
  // mso-highlight is not a CSS property: any style rewrite drops it, so the fold step keeps it as an attribute
  const mso = d.get('mso-highlight') ?? el.getAttribute('data-mso-highlight') ?? undefined
  if (mso && MSO_HIGHLIGHT[mso.trim().toLowerCase()])
    out.highlight = MSO_HIGHLIGHT[mso.trim().toLowerCase()]!
  else {
    const bg = cssColorToHex(d.get('background-color') ?? d.get('background')?.split(/\s+/)[0])
    if (bg && bg !== 'FFFFFF') out.shading = bg
  }
  return out
}

// ── tables ───────────────────────────────────────────────────────────────

export interface ForeignBorder {
  style: string
  szEighths?: number
  color?: string
}

const BORDER_STYLES: Record<string, string> = {
  solid: 'single',
  dashed: 'dashed',
  dotted: 'dotted',
  double: 'double',
  groove: 'threeDEngrave',
  ridge: 'threeDEmboss',
  inset: 'inset',
  outset: 'outset',
}
const WIDTH_WORDS: Record<string, number> = { thin: 0.75, medium: 1.5, thick: 3 }

/** one border value (`1.5pt solid red`, `solid red 1.5pt`, `none`) → model border; null = none/absent */
export function parseBorderValue(value: string | undefined): ForeignBorder | null {
  if (!value) return null
  let width: number | null = null
  let style: string | null = null
  let color: string | null = null
  const tokens = value.match(/rgba?\([^)]*\)|#[0-9a-fA-F]{3,6}|[^\s]+/g) ?? []
  for (const t of tokens) {
    const lower = t.toLowerCase()
    if (lower === 'none' || lower === 'hidden') return null
    if (BORDER_STYLES[lower]) style = BORDER_STYLES[lower]!
    else if (WIDTH_WORDS[lower] !== undefined) width = WIDTH_WORDS[lower]!
    else if (lengthToPt(t) !== null && /^-?[\d.]/.test(t)) width = lengthToPt(t)
    else {
      const c = cssColorToHex(t)
      if (c) color = c
    }
  }
  if (!style) return null
  if (width !== null && width <= 0) return null
  const sz = Math.max(2, Math.min(96, Math.round((width ?? 0.75) * 8)))
  return { style, szEighths: sz, color: color ?? 'auto' }
}

const SIDES = ['top', 'right', 'bottom', 'left'] as const
type Side = (typeof SIDES)[number]

/** per-side borders of an element's declarations (shorthand then longhand, in source order) */
export function foreignBorders(el: Element): Partial<Record<Side, ForeignBorder | 'none'>> {
  const out: Partial<Record<Side, ForeignBorder | 'none'>> = {}
  const raw = el.getAttribute('style') ?? ''
  const parts = raw.split(';')
  const partial: Record<Side, { w?: string; s?: string; c?: string }> = {
    top: {},
    right: {},
    bottom: {},
    left: {},
  }
  const apply = (side: Side, value: string) => {
    const b = parseBorderValue(value)
    out[side] = b ?? 'none'
    partial[side] = {}
  }
  for (const part of parts) {
    const at = part.indexOf(':')
    if (at < 0) continue
    const name = part.slice(0, at).trim().toLowerCase()
    const value = part
      .slice(at + 1)
      .replace(/!important/i, '')
      .trim()
    if (!value) continue
    if (name === 'border') for (const s of SIDES) apply(s, value)
    else if (
      name === 'border-top' ||
      name === 'border-right' ||
      name === 'border-bottom' ||
      name === 'border-left'
    )
      apply(name.slice(7) as Side, value)
    else {
      const m = /^border-(top|right|bottom|left)-(width|style|color)$/.exec(name)
      if (m) {
        const side = m[1] as Side
        const p = partial[side]
        if (m[2] === 'width') p.w = value
        else if (m[2] === 'style') p.s = value
        else p.c = value
        const current = out[side]
        const base = current && current !== 'none' ? current : null
        const merged = `${p.w ?? (base ? `${(base.szEighths ?? 4) / 8}pt` : '')} ${p.s ?? (base ? 'solid' : '')} ${p.c ?? (base?.color && base.color !== 'auto' ? `#${base.color}` : '')}`
        const b = parseBorderValue(merged)
        if (b) out[side] = b
        else if (p.s && /^(none|hidden)$/i.test(p.s)) out[side] = 'none'
      }
    }
  }
  return out
}

function widthPx(el: Element): number | null {
  const d = declarationsOf(el)
  const css = lengthToPt(d.get('width'))
  if (css !== null && !/%/.test(d.get('width') ?? '')) return Math.round((css * 4) / 3)
  const attr = el.getAttribute('width')
  if (attr && /^\d+(\.\d+)?$/.test(attr.trim())) return Math.round(parseFloat(attr))
  return null
}

/** td/th attrs from the cell's own presentation (not read from our own payload HTML) */
export function foreignCellAttrs(el: HTMLElement): Record<string, unknown> | null {
  const attrs: Record<string, unknown> = {}
  const d = declarationsOf(el)
  const fill = cssColorToHex(
    d.get('background-color') ?? d.get('background')?.split(/\s+/)[0] ?? el.getAttribute('bgcolor'),
  )
  if (fill && !(fill === 'FFFFFF' && !d.has('background') && !d.has('background-color')))
    attrs.fill = fill
  const borders = foreignBorders(el)
  const set: Record<string, ForeignBorder> = {}
  let any = false
  for (const side of SIDES) {
    const b = borders[side]
    if (!b) continue
    any = true
    const key = side
    set[key] = b === 'none' ? { style: 'nil' } : b
  }
  if (any) attrs.borders = set
  const va = (d.get('vertical-align') ?? el.getAttribute('valign') ?? '').toLowerCase()
  if (va === 'middle' || va === 'center') attrs.vAlign = 'center'
  else if (va === 'bottom') attrs.vAlign = 'bottom'
  const ta = (d.get('text-align') ?? el.getAttribute('align') ?? '').toLowerCase()
  if (ta === 'center' || ta === 'right' || ta === 'justify') attrs.align = ta
  const w = widthPx(el)
  const span = Math.max(1, parseInt(el.getAttribute('colspan') ?? '1', 10) || 1)
  if (w && w > 0) attrs.colwidth = Array.from({ length: span }, () => Math.round(w / span))
  return Object.keys(attrs).length > 0 ? attrs : null
}

export function foreignRowAttrs(el: HTMLElement): Record<string, unknown> | null {
  const d = declarationsOf(el)
  const pt = lengthToPt(d.get('height') ?? el.getAttribute('height'))
  if (pt && pt > 0) return { heightTwips: twips(pt), heightRule: 'atLeast' }
  return null
}

export function foreignTableAttrs(el: HTMLElement): Record<string, unknown> | null {
  const attrs: Record<string, unknown> = {}
  const w = widthPx(el)
  if (w && w > 0) attrs.widthPx = w
  const borders = foreignBorders(el)
  const declared = SIDES.filter((s) => borders[s] && borders[s] !== 'none')
  if (declared.length > 0) {
    const table: Record<string, ForeignBorder> = {}
    for (const s of SIDES) {
      const b = borders[s]
      if (b && b !== 'none') table[s] = b
    }
    attrs.borders = table
  }
  return Object.keys(attrs).length > 0 ? attrs : null
}

// ── paragraphs ───────────────────────────────────────────────────────────

/** paragraph attrs from a foreign block element's presentation (null when it declares nothing we read) */
export function foreignParaAttrs(el: HTMLElement): Record<string, unknown> | null {
  const d = declarationsOf(el)
  const attrs: Record<string, unknown> = {}
  const align = (d.get('text-align') ?? el.getAttribute('align') ?? '').toLowerCase()
  if (align === 'center' || align === 'right' || align === 'justify') attrs.align = align
  const box = (prop: string): number | null => {
    const side = d.get(prop)
    if (side) return lengthToPt(side)
    const short = d.get('margin')
    if (!short) return null
    const v = short.split(/\s+/).map((x) => lengthToPt(x))
    const idx = { 'margin-top': 0, 'margin-right': 1, 'margin-bottom': 2, 'margin-left': 3 }[prop]!
    const list =
      v.length === 1
        ? [v[0], v[0], v[0], v[0]]
        : v.length === 2
          ? [v[0], v[1], v[0], v[1]]
          : v.length === 3
            ? [v[0], v[1], v[2], v[1]]
            : v
    return list[idx] ?? null
  }
  const left = box('margin-left')
  const right = box('margin-right')
  if (left) attrs.indentLeft = twips(left)
  if (right) attrs.indentRight = twips(right)
  const indent = lengthToPt(d.get('text-indent'))
  if (indent) attrs.indentFirstLine = twips(indent)
  const before = box('margin-top')
  const after = box('margin-bottom')
  if (before !== null && before > 0) attrs.spaceBefore = twips(before)
  if (after !== null && after > 0) attrs.spaceAfter = twips(after)
  const lh = d.get('line-height')
  if (lh) {
    const pct = /^(\d+(?:\.\d+)?)%$/.exec(lh)
    const mult = /^\d+(\.\d+)?$/.test(lh) ? parseFloat(lh) : null
    if (pct) {
      attrs.lineSpacing = Math.round((parseFloat(pct[1]!) / 100) * 100) / 100
      attrs.lineRule = 'auto'
    } else if (mult) {
      attrs.lineSpacing = mult
      attrs.lineRule = 'auto'
    }
  }
  const bg = cssColorToHex(d.get('background-color') ?? d.get('background')?.split(/\s+/)[0])
  if (bg && bg !== 'FFFFFF' && el.tagName === 'P') attrs.shadingFill = bg
  return Object.keys(attrs).length > 0 ? attrs : null
}

// ── style sheets ─────────────────────────────────────────────────────────

interface Rule {
  selectors: string[]
  decls: string
}

function parseRules(css: string): Rule[] {
  const text = css.replace(/\/\*[\s\S]*?\*\//g, '').replace(/<!--|-->/g, '')
  const rules: Rule[] = []
  const re = /([^{}@]+)\{([^{}]*)\}/g
  let m: RegExpExecArray | null
  while ((m = re.exec(text))) {
    const selectors = m[1]!
      .split(',')
      .map((s) => s.trim())
      .filter((s) => /^[a-zA-Z0-9.#_*:-]+(?:\.[\w-]+)*$/.test(s) && !s.includes(':'))
    if (selectors.length > 0) rules.push({ selectors, decls: m[2]! })
  }
  return rules
}

/** the name of a Word style class on `el` (`class=MsoNormal` etc.) is not meaningful to the editor */
const SKIP_PROPS = /^(?:mso-(?!highlight|border|list|color-alt)|panose-|page-break|orphans|widows)/i

/**
 * Fold `<style>` rules into inline styles (simple tag / .class / tag.class / #id selectors — all Word
 * and Excel emit). Inline declarations keep priority over rule declarations.
 */
export function inlineStyleSheets(doc: Document): void {
  const sheets = [...doc.querySelectorAll('style')]
  if (sheets.length === 0) return
  const rules = sheets.flatMap((s) => parseRules(s.textContent ?? ''))
  for (const s of sheets) s.remove()
  if (rules.length === 0) return
  const body = doc.body
  for (const rule of rules) {
    const decls = [...parseDeclarations(rule.decls)].filter(([k]) => !SKIP_PROPS.test(k))
    if (decls.length === 0) continue
    for (const selector of rule.selectors) {
      let targets: Element[]
      try {
        targets = [...body.querySelectorAll(selector)]
      } catch {
        continue
      }
      for (const el of targets) {
        const own = declarationsOf(el)
        const merged = new Map<string, string>(decls)
        for (const [k, v] of own) merged.set(k, v)
        el.setAttribute('style', [...merged].map(([k, v]) => `${k}:${v}`).join(';'))
      }
    }
  }
}

// ── Word lists ───────────────────────────────────────────────────────────

/** carry non-CSS Word declarations that a later style rewrite would drop */
export function preserveMsoProps(doc: Document): void {
  for (const el of doc.body.querySelectorAll('[style*="mso-highlight"]')) {
    const v = declarationsOf(el).get('mso-highlight')
    if (v) el.setAttribute('data-mso-highlight', v)
  }
}

const BULLET_MARKERS = /^[·•●○▪■§o*\-–—➢❖]$/

/**
 * Word writes list items as `<p style="mso-list:l0 level2 lfo1">` with the
 * marker in a `mso-list:Ignore` span. Rewrite each as `<li data-list-kind
 * data-list-level>` so the schema's li rule builds a list item.
 */
export function convertMsoLists(doc: Document): void {
  for (const p of [...doc.body.querySelectorAll('p,div')]) {
    const style = p.getAttribute('style') ?? ''
    const m = /mso-list:\s*l(\d+)\s+level(\d+)/i.exec(style)
    if (!m || p.tagName !== 'P') continue
    const level = Math.max(0, Math.min(8, parseInt(m[2]!, 10) - 1))
    const marker = p.querySelector(
      'span[style*="mso-list:Ignore"], span[style*="mso-list: Ignore"]',
    )
    const markerText = (marker?.textContent ?? '').replace(/[\s\u00a0]+/g, '').trim()
    marker?.remove()
    // the conditional-comment wrapper Word emits around the marker is gone with the comments
    const li = doc.createElement('li')
    li.setAttribute(
      'data-list-kind',
      markerText && !BULLET_MARKERS.test(markerText) ? 'ordered' : 'bullet',
    )
    li.setAttribute('data-list-level', String(level))
    // keep the paragraph's own presentation (alignment, spacing) but drop the marker geometry
    li.setAttribute(
      'style',
      style
        .replace(/mso-list:[^;]*;?/i, '')
        .replace(/margin-left:[^;]*;?/i, '')
        .replace(/text-indent:[^;]*;?/i, ''),
    )
    while (p.firstChild) li.appendChild(p.firstChild)
    p.replaceWith(li)
  }
}
