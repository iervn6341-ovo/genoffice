/**
 * The HTML this app puts on the clipboard for a copied block.
 *
 * Univer's markup carries a few inline styles and the visible text; Excel pastes it with no number
 * formats, formulas or exact values, with wrong borders and with every cell's alignment and
 * wrapping lost. This writer produces Excel's own dialect: `x:num` / `x:str` / `x:bool` /
 * `x:fmla` attributes, `mso-number-format`, per-side borders in points, `mso-pattern` fills,
 * `<col>` / `<tr>` sizes, so Excel pastes the block as it was; other programs read the same
 * markup as an ordinary styled table. The table also keeps Univer's `data-copy-id` (this window's
 * own paste finds its cache entry by it) and the GenOffice payload attribute.
 */
import type { IColorStyle, IStyleData, Nullable } from '@univerjs/core'

import { PAYLOAD_ATTR } from './clipboard-payload'

export interface HtmlBlockCell {
  /// the text the cell displays
  readonly text: string
  readonly v?: string | number | boolean | null
  /// CellValueType: 1 string, 2 number, 3 boolean
  readonly t?: number | null
  readonly f?: string | null
  readonly style: IStyleData | null
  readonly rowSpan?: number
  readonly colSpan?: number
  /// the cell lies under a merged cell's top-left cell (no <td> of its own)
  readonly covered?: boolean
  readonly link?: string | null
}

export interface HtmlBlock {
  readonly copyId: string
  readonly payloadValue: string | null
  readonly cells: readonly (readonly HtmlBlockCell[])[]
  /// pixel sizes (null = default)
  readonly columnWidths: readonly (number | null)[]
  readonly rowHeights: readonly (number | null)[]
  readonly defaultColumnWidth?: number
  readonly defaultRowHeight?: number
  /// the workbook's default font, for cells that carry none of their own
  readonly defaultFont?: { readonly family: string; readonly size: number }
}

const escapeText = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
const escapeAttr = (text: string): string =>
  text.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/// px → pt with one decimal, the way Excel writes sizes
const pt = (px: number): string => `${((px * 72) / 96).toFixed(1)}pt`

const colorOf = (color: Nullable<IColorStyle>): string | null => {
  const rgb = color?.rgb
  if (!rgb) return null
  const hex = /^#?([0-9a-f]{6})([0-9a-f]{2})?$/i.exec(rgb)
  if (hex) return `#${hex[1]!.toUpperCase()}`
  return rgb
}

/// Excel's `mso-number-format` spelling of a number format pattern: every character that is not a
/// letter, digit, space or underscore is backslash-escaped, quotes become \0022.
export function msoNumberFormatOf(pattern: string): string {
  let out = ''
  for (const ch of pattern) {
    if (ch === '"') out += '\\0022'
    else if (ch === '\\') out += '\\\\'
    else if (/[A-Za-z0-9 _]/.test(ch)) out += ch
    else out += `\\${ch}`
  }
  return `"${out}"`
}

/// Border widths and line kinds by Univer border style (BorderStyleTypes).
const BORDER_CSS: Record<number, string> = {
  1: '.5pt solid',
  2: '.5pt hairline',
  3: '.5pt dotted',
  4: '.5pt dashed',
  5: '.5pt dot-dash',
  6: '.5pt dot-dot-dash',
  7: '2.25pt double',
  8: '1.0pt solid',
  9: '1.0pt dashed',
  10: '1.0pt dot-dash',
  11: '1.0pt dot-dot-dash',
  12: '1.0pt dot-dash',
  13: '2.0pt solid',
}

/// The inline CSS of one cell.
export function cellCss(
  style: IStyleData | null,
  value: { t?: number | null | undefined; hasNumber: boolean },
  defaultFont?: { family: string; size: number },
): string {
  const css: string[] = []
  const family = style?.ff ?? defaultFont?.family
  if (family) {
    const name = family.replace(/['"]/g, '')
    css.push(/\s/.test(name) ? `font-family:"${name}"` : `font-family:${name}`)
  }
  const size = style?.fs ?? defaultFont?.size
  if (size) css.push(`font-size:${size}pt`)
  if (style?.bl) css.push('font-weight:700')
  if (style?.it) css.push('font-style:italic')
  const lines: string[] = []
  if (style?.ul?.s) lines.push('underline')
  if (style?.st?.s) lines.push('line-through')
  if (lines.length) css.push(`text-decoration:${lines.join(' ')}`)
  if (style?.ul?.s && style.ul.t === 2) css.push('text-underline-style:double')
  const color = colorOf(style?.cl)
  if (color) css.push(`color:${color}`)
  const bg = colorOf(style?.bg)
  if (bg) css.push(`background:${bg}`, 'mso-pattern:black none')
  const ht = style?.ht
  if (ht === 1) css.push('text-align:left')
  else if (ht === 2) css.push('text-align:center')
  else if (ht === 3) css.push('text-align:right')
  else if (ht === 4) css.push('text-align:justify')
  else if (value.hasNumber) css.push('text-align:right')
  const vt = style?.vt
  if (vt === 1) css.push('vertical-align:top')
  else if (vt === 2) css.push('vertical-align:middle')
  else css.push('vertical-align:bottom')
  css.push(style?.tb === 3 ? 'white-space:normal' : 'white-space:nowrap')
  if (style?.tr?.a) css.push(`mso-rotate:${style.tr.a}`)
  const bd = style?.bd
  if (bd) {
    for (const [key, name] of [
      ['t', 'top'],
      ['r', 'right'],
      ['b', 'bottom'],
      ['l', 'left'],
    ] as const) {
      const side = bd[key]
      if (!side || !side.s) continue
      const kind = BORDER_CSS[side.s] ?? '.5pt solid'
      css.push(`border-${name}:${kind} ${colorOf(side.cl) ?? 'windowtext'}`)
    }
  }
  const pattern = style?.n?.pattern
  if (pattern && pattern !== 'General') css.push(`mso-number-format:${msoNumberFormatOf(pattern)}`)
  return css.join(';')
}

function cellMarkup(cell: HtmlBlockCell, defaultFont?: { family: string; size: number }): string {
  const attrs: string[] = []
  const hasNumber = cell.t === 2 && typeof cell.v === 'number'
  if (cell.rowSpan && cell.rowSpan > 1) attrs.push(`rowspan=${cell.rowSpan}`)
  if (cell.colSpan && cell.colSpan > 1) attrs.push(`colspan=${cell.colSpan}`)
  if (hasNumber) attrs.push(`x:num="${cell.v}"`)
  else if (cell.t === 3 && cell.v !== null && cell.v !== undefined) {
    attrs.push(`x:bool="${Number(cell.v) === 0 ? 'FALSE' : 'TRUE'}"`)
  } else if (cell.t === 1 && typeof cell.v === 'string' && cell.v !== '' && !cell.f) {
    // Excel reads the markup text as a typed entry; x:str keeps text that looks like a number,
    // a date or a boolean as the text it is
    if (
      /^[\s\d.,%$€£¥+\-/:()eE]+$/.test(cell.v) ||
      /^(true|false)$/i.test(cell.v.trim()) ||
      /^[=']/.test(cell.v)
    ) {
      attrs.push(`x:str="${escapeAttr(`'${cell.v}`)}"`)
    }
  }
  if (cell.f) attrs.push(`x:fmla="${escapeAttr(cell.f.startsWith('=') ? cell.f : `=${cell.f}`)}"`)
  const css = cellCss(cell.style, { t: cell.t, hasNumber }, defaultFont)
  attrs.push(`style='${css.replace(/'/g, '')}'`)
  let body = escapeText(cell.text).replace(
    /\r\n|\r|\n/g,
    "<br style='mso-data-placement:same-cell;'>",
  )
  if (cell.link) body = `<a href="${escapeAttr(cell.link)}">${body}</a>`
  return `<td ${attrs.join(' ')}>${body}</td>`
}

export function buildClipboardHtml(block: HtmlBlock): string {
  const colCount = block.cells[0]?.length ?? 0
  const defaultWidth = block.defaultColumnWidth ?? 69.33
  const defaultHeight = block.defaultRowHeight ?? 20
  const widths = Array.from({ length: colCount }, (_, j) => block.columnWidths[j] ?? defaultWidth)
  const total = widths.reduce((sum, w) => sum + w, 0)
  const cols = widths.map((w) => `<col width=${Math.round(w)} style='width:${pt(w)}'>`).join('')
  const rows = block.cells
    .map((line, i) => {
      const height = block.rowHeights[i] ?? defaultHeight
      const tds = line
        .map((cell) => (cell.covered ? '' : cellMarkup(cell, block.defaultFont)))
        .join('')
      return `<tr height=${Math.round(height)} style='height:${pt(height)}'>${tds}</tr>`
    })
    .join('\n')
  const payload = block.payloadValue ? ` ${PAYLOAD_ATTR}="${block.payloadValue}"` : ''
  return (
    '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    'xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta http-equiv=Content-Type content="text/html; charset=utf-8">' +
    '<meta name=ProgId content=Excel.Sheet><meta name=Generator content="GenOffice Sheets"></head>' +
    '<body><!--StartFragment-->' +
    `<table data-copy-id="${block.copyId}"${payload} border=0 cellpadding=0 cellspacing=0 width=${Math.round(total)} ` +
    `style='border-collapse:collapse;table-layout:fixed;width:${pt(total)}'>${cols}\n${rows}</table>` +
    '<!--EndFragment--></body></html>'
  )
}
