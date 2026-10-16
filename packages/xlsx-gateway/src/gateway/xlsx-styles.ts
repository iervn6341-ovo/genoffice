import type { WorkbookStyleEdit } from '../shared/edit-schemas'
import { isGradientFill, type FillSpec, type StyleColor } from '../domain/style-color'
import { shortDateNumFmtId } from '../shared/short-date'

/// Copy-on-write editor for xl/styles.xml. Existing entries are never
/// modified — every changed cell gets a new cellXfs entry (deduped) derived
/// from its current one, so untouched cells keep their exact formatting.
export class StylesheetEditor {
  private readonly source: string
  private numFmts: string[]
  private readonly fonts: string[]
  private readonly fills: string[]
  private readonly borders: string[]
  private readonly hadBordersSection: boolean
  private readonly cellXfs: string[]
  private readonly styleXfs: string[]
  private readonly dxfs: string[]
  private readonly hadDxfsSection: boolean
  /// the <tableStyle> elements, and the attributes of their <tableStyles> section (defaults)
  private readonly tableStyles: string[]
  private readonly tableStylesAttributes: string
  private readonly originalCounts: {
    numFmts: number
    fonts: number
    fills: number
    borders: number
    cellXfs: number
    dxfs: number
    tableStyles: number
  }
  private readonly cache = new Map<string, number>()
  private nextNumFmtId: number

  constructor(stylesXml: string) {
    this.source = stylesXml
    this.numFmts = extractElements(sectionInner(stylesXml, 'numFmts') ?? '', 'numFmt')
    const fontsInner = sectionInner(stylesXml, 'fonts')
    const fillsInner = sectionInner(stylesXml, 'fills')
    const bordersInner = sectionInner(stylesXml, 'borders')
    const cellXfsInner = sectionInner(stylesXml, 'cellXfs')
    if (fontsInner === null || fillsInner === null || cellXfsInner === null) {
      throw new Error(
        'The workbook stylesheet is missing fonts, fills, or cellXfs — style edits cannot be saved.',
      )
    }
    this.fonts = extractElements(fontsInner, 'font')
    this.fills = extractElements(fillsInner, 'fill')
    this.hadBordersSection = bordersInner !== null
    this.borders = bordersInner === null ? ['<border/>'] : extractElements(bordersInner, 'border')
    if (this.borders.length === 0) this.borders.push('<border/>')
    this.cellXfs = extractElements(cellXfsInner, 'xf')
    this.styleXfs = extractElements(sectionInner(stylesXml, 'cellStyleXfs') ?? '', 'xf')
    const dxfsInner = sectionInner(stylesXml, 'dxfs')
    this.hadDxfsSection = dxfsInner !== null
    this.dxfs = dxfsInner === null ? [] : extractElements(dxfsInner, 'dxf')
    const tableStylesInner = sectionInner(stylesXml, 'tableStyles')
    this.tableStyles =
      tableStylesInner === null ? [] : extractElements(tableStylesInner, 'tableStyle')
    this.tableStylesAttributes = (/<tableStyles\b([^>]*?)\/?>/.exec(stylesXml)?.[1] ?? '')
      .replace(/\bcount="[^"]*"/, '')
      .trim()
    if (this.fonts.length === 0 || this.cellXfs.length === 0) {
      throw new Error(
        'The workbook stylesheet has no base font or cell format — style edits cannot be saved.',
      )
    }
    this.originalCounts = {
      numFmts: this.numFmts.length,
      fonts: this.fonts.length,
      fills: this.fills.length,
      borders: this.borders.length,
      cellXfs: this.cellXfs.length,
      dxfs: this.dxfs.length,
      tableStyles: this.tableStyles.length,
    }
    this.nextNumFmtId =
      this.numFmts.reduce(
        (maximum, entry) => Math.max(maximum, Number(readAttribute(entry, 'numFmtId') ?? 0)),
        163,
      ) + 1
  }

  get changed(): boolean {
    return (
      this.numFmts.length !== this.originalCounts.numFmts ||
      this.fonts.length !== this.originalCounts.fonts ||
      this.fills.length !== this.originalCounts.fills ||
      this.borders.length !== this.originalCounts.borders ||
      this.cellXfs.length !== this.originalCounts.cellXfs ||
      this.dxfs.length !== this.originalCounts.dxfs ||
      this.tableStyles.length !== this.originalCounts.tableStyles
    )
  }

  hasCellXf(index: number): boolean {
    return Number.isInteger(index) && index >= 0 && index < this.cellXfs.length
  }

  /// Conditional-formatting highlight styles; deduped like every other list.
  internDxf(dxfXml: string): number {
    return internElement(this.dxfs, dxfXml)
  }

  /// A custom table style: each element (wholeTable, headerRow, …) points at its dxf. A style the
  /// stylesheet already holds under the name — in any case, as Excel compares them — is left alone (a
  /// second save does not write it twice).
  addTableStyle(name: string, elements: readonly { type: string; dxfXml: string }[]): void {
    const wanted = escapeXmlAttribute(name)
    const lower = wanted.toLowerCase()
    if (this.tableStyles.some((entry) => readAttribute(entry, 'name')?.toLowerCase() === lower))
      return
    const items = elements.map(
      ({ type, dxfXml }) => `<tableStyleElement type="${type}" dxfId="${this.internDxf(dxfXml)}"/>`,
    )
    this.tableStyles.push(
      `<tableStyle name="${wanted}" pivot="0" count="${items.length}">${items.join('')}</tableStyle>`,
    )
  }

  /// Returns the cellXfs index of a format equal to the base format with the
  /// delta applied, appending new numFmt/font/fill/xf entries as needed.
  resolveStyle(baseXfIndex: number, delta: WorkbookStyleEdit): number {
    const cacheKey = `${baseXfIndex}|${stableJson(delta)}`
    const cached = this.cache.get(cacheKey)
    if (cached !== undefined) return cached

    const baseXf = this.cellXfs[baseXfIndex] ?? this.cellXfs[0] ?? '<xf/>'
    let fontId = Number(readAttribute(baseXf, 'fontId') ?? 0)
    let fillId = Number(readAttribute(baseXf, 'fillId') ?? 0)
    let numFmtId = Number(readAttribute(baseXf, 'numFmtId') ?? 0)
    let borderId = Number(readAttribute(baseXf, 'borderId') ?? 0)
    if (hasFontDelta(delta)) {
      fontId = this.internFont(buildFont(this.fonts[fontId] ?? '<font/>', delta))
    }
    // Fill index 0 is the stylesheet's mandatory "none" pattern.
    if (delta.fill !== undefined) {
      fillId = delta.fill === null ? 0 : this.internFill(buildFill(delta.fill))
    } else if (delta.fillColor !== undefined) {
      fillId = delta.fillColor === null ? 0 : this.internFill(buildSolidFill(delta.fillColor))
    }
    if (delta.numberFormat !== undefined) {
      numFmtId = this.internNumberFormat(delta.numberFormat)
    }
    if (hasBorderDelta(delta)) {
      borderId = internElement(
        this.borders,
        buildBorder(this.borders[borderId] ?? '<border/>', delta),
      )
    }
    // Preserve everything the neutral model does not expose: quotePrefix,
    // namespaced attributes, extension children and the original apply flags.
    let patched = baseXf
    if (hasFontDelta(delta)) patched = patchXfFont(patched, fontId)
    if (delta.fill !== undefined || delta.fillColor !== undefined) {
      patched = patchXfAttribute(patched, 'fillId', String(fillId))
      patched = patchXfAttribute(patched, 'applyFill', '1')
    }
    if (delta.numberFormat !== undefined) {
      patched = patchXfAttribute(patched, 'numFmtId', String(numFmtId))
      patched = patchXfAttribute(patched, 'applyNumberFormat', '1')
    }
    if (hasBorderDelta(delta)) patched = patchXfBorder(patched, borderId)
    if (hasAlignmentDelta(delta)) patched = patchXfAlignment(patched, delta)
    if (delta.protectionLocked !== undefined || delta.protectionHidden !== undefined) {
      const parentXf = this.styleXfs[Number(readAttribute(baseXf, 'xfId') ?? -1)] ?? ''
      patched = patchXfProtection(patched, parentXf, delta)
    }
    const index = internElement(this.cellXfs, patched)
    this.cache.set(cacheKey, index)
    return index
  }

  serialize(): string {
    let result = this.source
    result = replaceSection(result, 'fonts', this.fonts)
    result = replaceSection(result, 'fills', this.fills)
    if (this.hadBordersSection) {
      result = replaceSection(result, 'borders', this.borders)
    } else if (this.borders.length > this.originalCounts.borders) {
      // Schema order: borders comes immediately after fills.
      const section = `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`
      result = result.replace(/<\/fills>|<fills\b[^>]*\/>/, (match) => `${match}${section}`)
    }
    result = replaceSection(result, 'cellXfs', this.cellXfs)
    if (this.hadDxfsSection) {
      result = replaceSection(result, 'dxfs', this.dxfs)
    } else if (this.dxfs.length > 0) {
      const section = `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>`
      // Schema order: dxfs follows cellStyles (when present) or cellXfs.
      const anchor =
        /<\/cellStyles>|<cellStyles\b[^>]*\/>/.exec(result) ??
        /<\/cellXfs>|<cellXfs\b[^>]*\/>/.exec(result)
      if (anchor) {
        const at = anchor.index + anchor[0].length
        result = result.slice(0, at) + section + result.slice(at)
      }
    }
    if (this.tableStyles.length !== this.originalCounts.tableStyles) {
      const attributes = this.tableStylesAttributes === '' ? '' : ` ${this.tableStylesAttributes}`
      const section = `<tableStyles count="${this.tableStyles.length}"${attributes}>${this.tableStyles.join('')}</tableStyles>`
      const existing = /<tableStyles\b[^>]*>[\s\S]*?<\/tableStyles>|<tableStyles\b[^>]*\/>/
      if (existing.test(result)) {
        result = result.replace(existing, () => section)
      } else {
        // Schema order: tableStyles follows dxfs, then colors / extLst.
        const anchor =
          /<\/dxfs>|<dxfs\b[^>]*\/>/.exec(result) ??
          /<\/cellStyles>|<cellStyles\b[^>]*\/>/.exec(result) ??
          /<\/cellXfs>|<cellXfs\b[^>]*\/>/.exec(result)
        if (anchor) {
          const at = anchor.index + anchor[0].length
          result = result.slice(0, at) + section + result.slice(at)
        }
      }
    }
    if (this.numFmts.length > 0) {
      const section = `<numFmts count="${this.numFmts.length}">${this.numFmts.join('')}</numFmts>`
      if (sectionInner(result, 'numFmts') !== null) {
        result = replaceSection(result, 'numFmts', this.numFmts)
      } else {
        // Schema order: numFmts comes immediately before fonts.
        result = result.replace(/<fonts\b/, () => `${section}<fonts`)
      }
    }
    return result
  }

  private internFont(fontXml: string): number {
    return internElement(this.fonts, fontXml)
  }

  private internFill(fillXml: string): number {
    return internElement(this.fills, fillXml)
  }

  private internNumberFormat(pattern: string): number {
    const builtin = BUILTIN_NUMBER_FORMATS.get(pattern) ?? shortDateNumFmtId(pattern)
    if (builtin !== undefined) return builtin
    for (const entry of this.numFmts) {
      if (readAttribute(entry, 'formatCode') === escapeXmlAttribute(pattern)) {
        return Number(readAttribute(entry, 'numFmtId') ?? 0)
      }
    }
    const id = this.nextNumFmtId
    this.nextNumFmtId += 1
    this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${escapeXmlAttribute(pattern)}"/>`)
    return id
  }
}

export const BUILTIN_NUMBER_FORMATS = new Map<string, number>([
  ['General', 0],
  ['0', 1],
  ['0.00', 2],
  ['#,##0', 3],
  ['#,##0.00', 4],
  ['0%', 9],
  ['0.00%', 10],
  ['0.00E+00', 11],
  ['@', 49],
])

function hasFontDelta(delta: WorkbookStyleEdit): boolean {
  return (
    delta.bold !== undefined ||
    delta.italic !== undefined ||
    delta.underline !== undefined ||
    delta.underlineStyle !== undefined ||
    delta.fontVerticalAlign !== undefined ||
    delta.strikethrough !== undefined ||
    delta.fontFamily !== undefined ||
    delta.fontSize !== undefined ||
    delta.fontColor !== undefined
  )
}

function patchXfFont(baseXf: string, fontId: number): string {
  const opening = /^<xf\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/.exec(baseXf)?.[0]
  if (!opening) throw new Error('The original cell format has an invalid opening tag.')
  const set = (tag: string, name: string, value: string): string => {
    const existing = new RegExp(`(?<![\\w:.-])${name}=(?:"[^"]*"|'[^']*')`)
    return existing.test(tag)
      ? tag.replace(existing, `${name}="${value}"`)
      : tag.replace(/\/?>$/, (end) => ` ${name}="${value}"${end}`)
  }
  const patched = set(set(opening, 'fontId', String(fontId)), 'applyFont', '1')
  return `${patched}${baseXf.slice(opening.length)}`
}

function patchXfBorder(baseXf: string, borderId: number): string {
  const opening = /^<xf\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/.exec(baseXf)?.[0]
  if (!opening) throw new Error('The original cell format has an invalid opening tag.')
  const set = (tag: string, name: string, value: string): string => {
    const existing = new RegExp(`(?<![\\w:.-])${name}=(?:"[^"]*"|'[^']*')`)
    return existing.test(tag)
      ? tag.replace(existing, `${name}="${value}"`)
      : tag.replace(/\/?>$/, (end) => ` ${name}="${value}"${end}`)
  }
  const patched = set(set(opening, 'borderId', String(borderId)), 'applyBorder', '1')
  return `${patched}${baseXf.slice(opening.length)}`
}

/// Applies the delta to a copy of the base font XML. Only overridden child
/// elements are replaced; everything else (family, scheme, charset) is kept.
function buildFont(baseFontXml: string, delta: WorkbookStyleEdit): string {
  let inner = /<font\b[^>]*>([\s\S]*?)<\/font>/.exec(baseFontXml)?.[1] ?? ''
  const added: string[] = []
  const override = (pattern: RegExp, replacement: string): void => {
    inner = inner.replace(pattern, '')
    if (replacement !== '') added.push(replacement)
  }
  if (delta.bold !== undefined) override(/<b\b[^>]*\/?>/g, delta.bold ? '<b/>' : '')
  if (delta.italic !== undefined) override(/<i\b[^>]*\/?>/g, delta.italic ? '<i/>' : '')
  if (delta.underline !== undefined || delta.underlineStyle !== undefined) {
    const on = delta.underline ?? true
    override(
      /<u\b[^>]*\/?>/g,
      on
        ? delta.underlineStyle === undefined || delta.underlineStyle === 'single'
          ? '<u/>'
          : `<u val="${delta.underlineStyle}"/>`
        : '',
    )
  }
  if (delta.fontVerticalAlign !== undefined) {
    override(
      /<vertAlign\b[^>]*\/?>/g,
      delta.fontVerticalAlign === 'baseline' ? '' : `<vertAlign val="${delta.fontVerticalAlign}"/>`,
    )
  }
  if (delta.strikethrough !== undefined) {
    override(/<strike\b[^>]*\/?>/g, delta.strikethrough ? '<strike/>' : '')
  }
  if (delta.fontSize !== undefined) {
    override(/<sz\b[^>]*\/?>/g, `<sz val="${delta.fontSize}"/>`)
  }
  if (delta.fontColor !== undefined) {
    override(
      /<color\b[^>]*\/?>/g,
      delta.fontColor === null ? '' : `<color ${colorAttributes(delta.fontColor)}/>`,
    )
  }
  if (delta.fontFamily !== undefined) {
    override(/<name\b[^>]*\/?>/g, `<name val="${escapeXmlAttribute(delta.fontFamily)}"/>`)
  }
  const content = `${added.join('')}${inner}`
  return content === '' ? '<font/>' : `<font>${content}</font>`
}

function buildSolidFill(fillColor: StyleColor): string {
  return `<fill><patternFill patternType="solid"><fgColor ${colorAttributes(fillColor)}/><bgColor indexed="64"/></patternFill></fill>`
}

function buildFill(fill: FillSpec): string {
  if (isGradientFill(fill)) {
    const g = fill.gradient
    const attributes = [
      ...(g.type === 'path' ? ['type="path"'] : []),
      ...(g.type !== 'path' && g.angle !== undefined ? [`degree="${g.angle}"`] : []),
      ...(['left', 'right', 'top', 'bottom'] as const).flatMap((edge) =>
        g.type === 'path' && g[edge] !== undefined ? [`${edge}="${g[edge]}"`] : [],
      ),
    ]
    const stops = g.stops
      .map(
        (stop) =>
          `<stop position="${stop.position}"><color ${colorAttributes(stop.color)}/></stop>`,
      )
      .join('')
    const open =
      attributes.length === 0 ? '<gradientFill>' : `<gradientFill ${attributes.join(' ')}>`
    return `<fill>${open}${stops}</gradientFill></fill>`
  }
  if (fill.pattern === 'solid' && fill.bg === undefined) return buildSolidFill(fill.fg)
  const bg = fill.bg === undefined ? '' : `<bgColor ${colorAttributes(fill.bg)}/>`
  return `<fill><patternFill patternType="${fill.pattern}"><fgColor ${colorAttributes(fill.fg)}/>${bg}</patternFill></fill>`
}

/** CT_Color attributes: cached rgb for literals, theme index (+ tint) for theme slots */
function colorAttributes(color: StyleColor): string {
  if (typeof color === 'string') return `rgb="${toArgb(color)}"`
  const tint = color.tint === undefined || color.tint === 0 ? '' : ` tint="${color.tint}"`
  return `theme="${color.theme}"${tint}`
}

const BORDER_EDGE_TAGS = ['left', 'right', 'top', 'bottom', 'diagonal'] as const
const BORDER_DELTA_KEYS = {
  left: 'borderLeft',
  right: 'borderRight',
  top: 'borderTop',
  bottom: 'borderBottom',
  diagonal: 'borderDiagonal',
} as const

function hasBorderDelta(delta: WorkbookStyleEdit): boolean {
  return (
    delta.borderTop !== undefined ||
    delta.borderBottom !== undefined ||
    delta.borderLeft !== undefined ||
    delta.borderRight !== undefined ||
    delta.borderDiagonal !== undefined ||
    delta.diagonalUp !== undefined ||
    delta.diagonalDown !== undefined
  )
}

function patchXfAttribute(baseXf: string, name: string, value: string): string {
  const opening = /^<xf\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/.exec(baseXf)?.[0]
  if (!opening) throw new Error('The original cell format has an invalid opening tag.')
  return `${setXmlAttribute(opening, name, value)}${baseXf.slice(opening.length)}`
}

function setXmlAttribute(opening: string, name: string, value: string | null): string {
  const existing = new RegExp(`(?<![\\w:.-])${name}=(?:"[^"]*"|'[^']*')`)
  if (value === null) return opening.replace(new RegExp(`\\s+${existing.source}`), '')
  return existing.test(opening)
    ? opening.replace(existing, `${name}="${value}"`)
    : opening.replace(/\/?>$/, (end) => ` ${name}="${value}"${end}`)
}

/** Keep imported XF attributes and extension children when borders and alignment change together. */
function patchXfAlignment(baseXf: string, delta: WorkbookStyleEdit): string {
  const opening = /^<xf\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/.exec(baseXf)?.[0]
  if (!opening) throw new Error('The original cell format has an invalid opening tag.')
  const selfClosing = opening.endsWith('/>')
  const inner = selfClosing ? '' : baseXf.slice(opening.length, -'</xf>'.length)
  const direct =
    /^(\s*)(<alignment\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<alignment\b(?:[^>"']|"[^"]*"|'[^']*')*><\/alignment>)/.exec(
      inner,
    )
  const prefix = direct?.[1] ?? ''
  let alignment = (direct?.[2] ?? '<alignment/>').replace(/><\/alignment>$/, '/>')
  const changes: readonly (readonly [string, string | null | undefined])[] = [
    ['horizontal', delta.horizontalAlignment],
    [
      'vertical',
      delta.verticalAlignment === undefined ? undefined : XLSX_VERTICAL[delta.verticalAlignment],
    ],
    ['wrapText', delta.wrapText === undefined ? undefined : delta.wrapText ? '1' : null],
    ['shrinkToFit', delta.shrinkToFit === undefined ? undefined : delta.shrinkToFit ? '1' : null],
    [
      'textRotation',
      delta.textRotation === undefined
        ? undefined
        : delta.textRotation === 0
          ? null
          : String(delta.textRotation),
    ],
    [
      'indent',
      delta.indent === undefined ? undefined : delta.indent === 0 ? null : String(delta.indent),
    ],
  ]
  for (const [name, value] of changes) {
    if (value !== undefined) alignment = setXmlAttribute(alignment, name, value)
  }
  const rest = direct ? inner.slice(direct[0].length) : inner
  const content = `${prefix}${alignment}${rest}`
  const expanded = selfClosing
    ? `${opening.slice(0, -2)}>${content}</xf>`
    : `${opening}${content}</xf>`
  return patchXfAttribute(expanded, 'applyAlignment', '1')
}

function patchXfProtection(baseXf: string, parentXf: string, delta: WorkbookStyleEdit): string {
  const opening = /^<xf\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/.exec(baseXf)?.[0]
  if (!opening) throw new Error('The original cell format has an invalid opening tag.')
  const selfClosing = opening.endsWith('/>')
  const inner = selfClosing ? '' : baseXf.slice(opening.length, -'</xf>'.length)
  const extensionAt = inner.search(/<extLst\b|<mc:AlternateContent\b/)
  const directRegion = extensionAt < 0 ? inner : inner.slice(0, extensionAt)
  const direct =
    /<protection\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<protection\b(?:[^>"']|"[^"]*"|'[^']*')*><\/protection>/.exec(
      directRegion,
    )
  const parent = /<protection\b[^>]*\/?>/.exec(parentXf)?.[0] ?? ''
  const locked =
    delta.protectionLocked ??
    (readAttribute(direct?.[0] ?? '', 'locked') ?? readAttribute(parent, 'locked')) !== '0'
  const hidden =
    delta.protectionHidden ??
    (readAttribute(direct?.[0] ?? '', 'hidden') ?? readAttribute(parent, 'hidden')) === '1'
  if (locked && !hidden && !parent && !direct) return baseXf
  let protection = (direct?.[0] ?? '<protection/>').replace(/><\/protection>$/, '/>')
  protection = setXmlAttribute(protection, 'locked', locked ? '1' : '0')
  protection = setXmlAttribute(protection, 'hidden', hidden ? '1' : '0')
  const content = direct
    ? `${inner.slice(0, direct.index)}${protection}${inner.slice(direct.index + direct[0].length)}`
    : `${directRegion}${protection}${inner.slice(directRegion.length)}`
  const expanded = selfClosing
    ? `${opening.slice(0, -2)}>${content}</xf>`
    : `${opening}${content}</xf>`
  return patchXfAttribute(expanded, 'applyProtection', '1')
}

/// Applies edge deltas to a copy of the base border XML. Children rebuild in
/// schema order (left, right, top, bottom, diagonal); untouched edges and the
/// border element's own attributes (diagonalUp/Down) are kept verbatim.
function buildBorder(baseBorderXml: string, delta: WorkbookStyleEdit): string {
  const opening = /^<border\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/.exec(baseBorderXml)?.[0]
  if (!opening) throw new Error('The original border has an invalid opening tag.')
  let tag = opening
  for (const [attribute, value] of [
    ['diagonalUp', delta.diagonalUp],
    ['diagonalDown', delta.diagonalDown],
  ] as const) {
    if (value === undefined) continue
    const existing = new RegExp(`(?<![\\w:.-])${attribute}=(?:"[^"]*"|'[^']*')`)
    tag = existing.test(tag)
      ? tag.replace(existing, `${attribute}="${value ? '1' : '0'}"`)
      : tag.replace(/\/?>$/, (end) => ` ${attribute}="${value ? '1' : '0'}"${end}`)
  }
  let inner = baseBorderXml.slice(opening.length)
  if (opening.endsWith('/>')) inner = ''
  else if (inner.endsWith('</border>')) inner = inner.slice(0, -'</border>'.length)
  else throw new Error('The original border has no closing tag.')
  for (const [index, edgeTag] of BORDER_EDGE_TAGS.entries()) {
    const edge = delta[BORDER_DELTA_KEYS[edgeTag]]
    if (edge === undefined) continue
    const color =
      edge === null || edge.color === undefined ? '' : `<color ${colorAttributes(edge.color)}/>`
    const replacement =
      edge === null
        ? `<${edgeTag}/>`
        : color === ''
          ? `<${edgeTag} style="${edge.style}"/>`
          : `<${edgeTag} style="${edge.style}">${color}</${edgeTag}>`
    const existing = new RegExp(`<${edgeTag}\\b[^>]*/>|<${edgeTag}\\b[^>]*>[\\s\\S]*?</${edgeTag}>`)
    if (existing.test(inner)) inner = inner.replace(existing, () => replacement)
    else {
      const following = [...BORDER_EDGE_TAGS.slice(index + 1), 'vertical', 'horizontal']
        .map((name) => inner.search(new RegExp(`<${name}\\b`)))
        .filter((at) => at >= 0)
      const extension = inner.search(/<extLst\b/)
      if (extension >= 0) following.push(extension)
      const at = following.length > 0 ? Math.min(...following) : inner.length
      inner = `${inner.slice(0, at)}${replacement}${inner.slice(at)}`
    }
  }
  if (inner === '' && !/\s+[\w:.-]+=/.test(tag)) return '<border/>'
  return `${tag.replace(/\/?>$/, '>')}${inner}</border>`
}

function hasAlignmentDelta(delta: WorkbookStyleEdit): boolean {
  return (
    delta.horizontalAlignment !== undefined ||
    delta.verticalAlignment !== undefined ||
    delta.wrapText !== undefined ||
    delta.shrinkToFit !== undefined ||
    delta.textRotation !== undefined ||
    delta.indent !== undefined
  )
}

const XLSX_VERTICAL: Record<string, string> = {
  top: 'top',
  center: 'center',
  bottom: 'bottom',
  justify: 'justify',
  distributed: 'distributed',
}

/** key-sorted at every level: a replacer array would drop nested keys (border edges, fills) */
function stableJson(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`
  if (value !== null && typeof value === 'object') {
    const entries = Object.entries(value as Record<string, unknown>)
      .filter(([, v]) => v !== undefined)
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
      .map(([k, v]) => `${JSON.stringify(k)}:${stableJson(v)}`)
    return `{${entries.join(',')}}`
  }
  return JSON.stringify(value)
}

function internElement(list: string[], element: string): number {
  const existing = list.indexOf(element)
  if (existing !== -1) return existing
  list.push(element)
  return list.length - 1
}

function sectionInner(xml: string, tag: string): string | null {
  const selfClosing = new RegExp(`<${tag}\\b[^>]*/>`)
  if (selfClosing.test(xml)) return ''
  const match = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml)
  return match?.[1] ?? null
}

function replaceSection(xml: string, tag: string, elements: readonly string[]): string {
  const start = new RegExp(`<${tag}(?=[\\s/>])`).exec(xml)?.index
  if (start === undefined) return xml
  let quote: '"' | "'" | null = null
  let end = start + tag.length + 1
  for (; end < xml.length; end += 1) {
    const char = xml[end]
    if (quote) {
      if (char === quote) quote = null
    } else if (char === '"' || char === "'") quote = char
    else if (char === '>') break
  }
  if (end >= xml.length) throw new Error(`The stylesheet has an invalid <${tag}> tag.`)
  const opening = xml.slice(start, end + 1)
  const close = opening.endsWith('/>') ? end + 1 : xml.indexOf(`</${tag}>`, end + 1)
  if (close < 0) throw new Error(`The stylesheet has an unclosed <${tag}> tag.`)
  const after = opening.endsWith('/>') ? close : close + `</${tag}>`.length
  const attributes = setXmlAttribute(opening, 'count', String(elements.length))
  const expanded = attributes.endsWith('/>') ? `${attributes.slice(0, -2)}>` : attributes
  return `${xml.slice(0, start)}${expanded}${elements.join('')}</${tag}>${xml.slice(after)}`
}

function extractElements(inner: string, tag: string): string[] {
  const result: string[] = []
  const start = new RegExp(`<${tag}(?=[\\s/>])`, 'g')
  for (let match = start.exec(inner); match; match = start.exec(inner)) {
    let quote: '"' | "'" | null = null
    let end = match.index + match[0].length
    for (; end < inner.length; end += 1) {
      const char = inner[end]
      if (quote) {
        if (char === quote) quote = null
      } else if (char === '"' || char === "'") quote = char
      else if (char === '>') break
    }
    if (end >= inner.length) throw new Error(`The stylesheet has an invalid <${tag}> tag.`)
    if (inner[end - 1] === '/') {
      result.push(inner.slice(match.index, end + 1))
      start.lastIndex = end + 1
      continue
    }
    const closing = `</${tag}>`
    const close = inner.indexOf(closing, end + 1)
    if (close < 0) throw new Error(`The stylesheet has an unclosed <${tag}> tag.`)
    result.push(inner.slice(match.index, close + closing.length))
    start.lastIndex = close + closing.length
  }
  return result
}

function readAttribute(element: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(element)?.[1]
}

function toArgb(hexColor: string): string {
  return `FF${hexColor.slice(1).toUpperCase()}`
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;')
}
