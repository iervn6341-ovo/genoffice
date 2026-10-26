/**
 * Plain-text / HTML renditions of copied elements, and HTML-table ingestion.
 *
 * Copying a table from a slide should paste into Excel / Numbers / Pages / a text editor as
 * a grid (HTML `<table>` + tab-separated text), and a table copied from Excel, Numbers or a
 * web page should paste into a slide as a native table — PowerPoint does both.
 */
import type { ElementClipboardItem } from './index'
import type { SlideElement, TableCell, TableElement, TextBody } from './types'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

const EMU_PER_PX = 9525
const BUILTIN_DEFAULT_STYLE = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

// ── plain text ──────────────────────────────────────────────────────────

function bodyText(text: TextBody | undefined): string {
  return (text?.paragraphs ?? [])
    .map((p) =>
      p.runs
        .filter((r) => !r.paraMark)
        .map((r) => r.text)
        .join(''),
    )
    .join('\n')
}

export function tableRowsText(el: TableElement): string[][] {
  return el.rows.map((row) =>
    row.map((c) => (c.merged ? '' : bodyText(c.text).replace(/\t/g, ' ').replace(/\n/g, ' '))),
  )
}

/** Tab/newline separated rendition of a table (what a spreadsheet pastes as a grid) */
export function tableToTsv(el: TableElement): string {
  return tableRowsText(el)
    .map((r) => r.join('\t'))
    .join('\n')
}

/** The text of copied elements: tables as TSV, shapes as their paragraphs, groups recursively */
export function elementPlainText(el: SlideElement): string {
  switch (el.type) {
    case 'table':
      return tableToTsv(el)
    case 'text':
    case 'shape':
      return bodyText(el.text)
    case 'group':
      return el.children.map(elementPlainText).filter(Boolean).join('\n')
    default:
      return ''
  }
}

// ── HTML rendition of a table ──────────────────────────────────────────

function css(cell: TableCell): string {
  const parts: string[] = []
  if (cell.fill?.type === 'solid') parts.push(`background:${cell.fill.color.slice(0, 7)}`)
  const run = cell.text?.paragraphs.flatMap((p) => p.runs).find((r) => !r.paraMark && r.text)
  if (run?.color) parts.push(`color:${run.color.slice(0, 7)}`)
  if (run?.bold) parts.push('font-weight:bold')
  if (run?.italic) parts.push('font-style:italic')
  if (run?.fontSize) parts.push(`font-size:${run.fontSize}pt`)
  if (run?.fontFamily) parts.push(`font-family:'${run.fontFamily.replace(/'/g, '')}'`)
  const align = cell.text?.paragraphs[0]?.align
  if (align) parts.push(`text-align:${align}`)
  const b = cell.borders
  const edge = (s: NonNullable<typeof b>['l'] | undefined) =>
    s && s.fill.type === 'solid'
      ? `${Math.max(1, Math.round(s.width / 12700))}pt solid ${s.fill.color.slice(0, 7)}`
      : ''
  for (const [k, side] of [
    ['l', 'left'],
    ['r', 'right'],
    ['t', 'top'],
    ['b', 'bottom'],
  ] as const) {
    const e = edge(b?.[k])
    if (e) parts.push(`border-${side}:${e}`)
  }
  return parts.join(';')
}

export function tableToHtml(el: TableElement): string {
  const rows = el.rows
    .map((row) => {
      const cells = row
        .map((c, ci) => {
          if (c.merged) return ''
          const span =
            (c.gridSpan && c.gridSpan > 1 ? ` colspan="${c.gridSpan}"` : '') +
            (c.rowSpan && c.rowSpan > 1 ? ` rowspan="${c.rowSpan}"` : '')
          const style = css(c)
          const w =
            ci < el.colWidths.length ? ` width="${Math.round(el.colWidths[ci]! / EMU_PER_PX)}"` : ''
          const body = bodyText(c.text)
            .split('\n')
            .map((l) => escapeXmlText(l))
            .join('<br>')
          return `<td${span}${w}${style ? ` style="${escapeXmlAttr(style)}"` : ''}>${body}</td>`
        })
        .join('')
      return `<tr>${cells}</tr>`
    })
    .join('')
  return `<table border="0" cellspacing="0" cellpadding="3">${rows}</table>`
}

// ── HTML table → native table ──────────────────────────────────────────

const ENTITIES: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: ' ',
  ensp: ' ',
  emsp: ' ',
  hellip: '…',
  ndash: '–',
  mdash: '—',
  copy: '©',
  reg: '®',
  trade: '™',
}

function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, name: string) => {
    if (name[0] === '#') {
      const code =
        name[1] === 'x' || name[1] === 'X'
          ? parseInt(name.slice(2), 16)
          : parseInt(name.slice(1), 10)
      return Number.isFinite(code) && code > 0 && code < 0x110000
        ? String.fromCodePoint(code)
        : whole
    }
    return ENTITIES[name.toLowerCase()] ?? whole
  })
}

function cellText(inner: string): string {
  const withBreaks = inner
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(script|style)\b[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
  return decodeEntities(withBreaks)
    .replace(/[ \t\r\f\v]+/g, ' ')
    .replace(/ ?\n ?/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}

function styleMap(style: string | undefined): Record<string, string> {
  const out: Record<string, string> = {}
  for (const decl of (style ?? '').split(';')) {
    const i = decl.indexOf(':')
    if (i > 0) out[decl.slice(0, i).trim().toLowerCase()] = decl.slice(i + 1).trim()
  }
  return out
}

const NAMED: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  grey: '808080',
  orange: 'FFA500',
  purple: '800080',
}

function cssColor(v: string | undefined): string | undefined {
  if (!v) return undefined
  const s = v.trim().toLowerCase()
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(s)
  if (hex) {
    const h = hex[1]!
    return (h.length === 3 ? [...h].map((c) => c + c).join('') : h).toUpperCase()
  }
  const rgb = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(s)
  if (rgb) {
    return [rgb[1], rgb[2], rgb[3]]
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase()
  }
  const named = NAMED[s]
  if (named) return named
  return undefined
}

function attr(attrs: string, name: string): string | undefined {
  const m = new RegExp(`\\b${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, 'i').exec(attrs)
  return m ? (m[1] ?? m[2] ?? m[3]) : undefined
}

function lengthEmu(v: string | undefined): number | undefined {
  const m = /^\s*([\d.]+)\s*(px|pt|in|cm|mm)?\s*$/i.exec(v ?? '')
  if (!m) return undefined
  const n = parseFloat(m[1]!)
  switch ((m[2] ?? 'px').toLowerCase()) {
    case 'pt':
      return Math.round(n * 12700)
    case 'in':
      return Math.round(n * 914400)
    case 'cm':
      return Math.round((n * 914400) / 2.54)
    case 'mm':
      return Math.round((n * 914400) / 25.4)
    default:
      return Math.round(n * EMU_PER_PX)
  }
}

interface ParsedCell {
  text: string
  header: boolean
  colSpan: number
  rowSpan: number
  fill?: string
  color?: string
  bold?: boolean
  italic?: boolean
  sizePt?: number
  align?: 'l' | 'ctr' | 'r'
  widthEmu?: number
}

/** The first `<table>` of an HTML fragment as rows of cells (null for none / nested-only) */
export function parseHtmlTable(
  html: string,
): { rows: ParsedCell[][]; hasFormatting: boolean } | null {
  const start = html.search(/<table\b/i)
  if (start < 0) return null
  // depth-aware end of the outermost table
  let depth = 0
  let end = -1
  const re = /<table\b|<\/table>/gi
  re.lastIndex = start
  let m: RegExpExecArray | null
  while ((m = re.exec(html))) {
    depth += m[0][1] === '/' ? -1 : 1
    if (depth === 0) {
      end = m.index + m[0].length
      break
    }
  }
  const table = html.slice(start, end < 0 ? html.length : end)
  const rows: ParsedCell[][] = []
  let hasFormatting = false
  for (const tr of table.matchAll(/<tr\b[^>]*>([\s\S]*?)(?=<tr\b|<\/tr>|<\/table>|$)/gi)) {
    const cells: ParsedCell[] = []
    for (const td of tr[1]!.matchAll(
      /<(t[dh])\b([^>]*)>([\s\S]*?)(?=<t[dh]\b|<\/t[dh]>|<\/tr>|$)/gi,
    )) {
      const attrs = td[2]!
      const st = styleMap(attr(attrs, 'style'))
      const fill = cssColor(st['background'] ?? st['background-color'] ?? attr(attrs, 'bgcolor'))
      const color = cssColor(st['color'])
      const weight = st['font-weight']
      const bold = weight
        ? /bold|[6-9]00/.test(weight)
        : td[1]!.toLowerCase() === 'th'
          ? true
          : undefined
      const italic = st['font-style'] ? /italic|oblique/.test(st['font-style']) : undefined
      const size = lengthEmu(st['font-size'])
      const align = (st['text-align'] ?? attr(attrs, 'align'))?.toLowerCase()
      const cell: ParsedCell = {
        text: cellText(td[3]!),
        header: td[1]!.toLowerCase() === 'th',
        colSpan: Math.max(1, parseInt(attr(attrs, 'colspan') ?? '1', 10) || 1),
        rowSpan: Math.max(1, parseInt(attr(attrs, 'rowspan') ?? '1', 10) || 1),
        ...(fill && fill !== 'FFFFFF' ? { fill } : {}),
        ...(color && color !== '000000' ? { color } : {}),
        ...(bold !== undefined ? { bold } : {}),
        ...(italic ? { italic } : {}),
        ...(size ? { sizePt: Math.round(size / 12700) } : {}),
        ...(align === 'center'
          ? { align: 'ctr' as const }
          : align === 'right'
            ? { align: 'r' as const }
            : align === 'left'
              ? { align: 'l' as const }
              : {}),
      }
      const w = lengthEmu(st['width'] ?? attr(attrs, 'width'))
      if (w) cell.widthEmu = w
      if (cell.fill || cell.color || (cell.bold && !cell.header) || cell.sizePt)
        hasFormatting = true
      cells.push(cell)
    }
    if (cells.length) rows.push(cells)
  }
  return rows.length ? { rows, hasFormatting } : null
}

function tcXml(cell: ParsedCell, fillExplicit: boolean, merged: 'h' | 'v' | null): string {
  const lines = (cell.text || '').split('\n')
  const rPr = (empty: boolean) => {
    const attrs = ` lang="en-US"${cell.sizePt ? ` sz="${cell.sizePt * 100}"` : ''}${cell.bold ? ' b="1"' : ''}${cell.italic ? ' i="1"' : ''}${empty ? '' : ' dirty="0"'}`
    return `<a:rPr${attrs}>${cell.color ? `<a:solidFill><a:srgbClr val="${cell.color}"/></a:solidFill>` : ''}</a:rPr>`
  }
  const paras = lines
    .map(
      (l) =>
        `<a:p>${cell.align ? `<a:pPr algn="${cell.align}"/>` : ''}${l ? `<a:r>${rPr(false)}<a:t>${escapeXmlText(l)}</a:t></a:r>` : ''}<a:endParaRPr lang="en-US"${cell.sizePt ? ` sz="${cell.sizePt * 100}"` : ''}/></a:p>`,
    )
    .join('')
  const edge = '<a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill>'
  const borders = fillExplicit
    ? ['L', 'R', 'T', 'B'].map((s) => `<a:ln${s} w="6350">${edge}</a:ln${s}>`).join('')
    : ''
  const fill = cell.fill
    ? `<a:solidFill><a:srgbClr val="${cell.fill}"/></a:solidFill>`
    : fillExplicit
      ? '<a:noFill/>'
      : ''
  const span =
    (cell.colSpan > 1 ? ` gridSpan="${cell.colSpan}"` : '') +
    (cell.rowSpan > 1 ? ` rowSpan="${cell.rowSpan}"` : '') +
    (merged === 'h' ? ' hMerge="1"' : merged === 'v' ? ' vMerge="1"' : '')
  return `<a:tc${span}><a:txBody><a:bodyPr/><a:lstStyle/>${paras}</a:txBody><a:tcPr${borders || fill ? '' : '/'}>${borders || fill ? `${borders}${fill}</a:tcPr>` : ''}</a:tc>`
}

/**
 * An HTML `<table>` (Excel, Numbers, Google Sheets, a web page) as a native table item.
 * Unformatted tables take the default table style like PowerPoint's own paste; tables with
 * explicit colours keep them (and get thin grid lines). Colspan/rowspan become merges.
 */
export function htmlTableToItem(
  html: string,
  opts: { maxWidthEmu?: number; maxHeightEmu?: number } = {},
): ElementClipboardItem | null {
  const parsed = parseHtmlTable(html)
  if (!parsed) return null
  const { rows, hasFormatting } = parsed
  // lay out onto a grid: spans claim following positions
  const nCols = Math.max(1, ...rows.map((r) => r.reduce((n, c) => n + c.colSpan, 0)))
  const grid: Array<Array<{ cell: ParsedCell; merged: 'h' | 'v' | null } | null>> = rows.map(() =>
    Array.from({ length: nCols }, () => null),
  )
  rows.forEach((row, r) => {
    let c = 0
    for (const cell of row) {
      while (c < nCols && grid[r]![c]) c++
      if (c >= nCols) break
      for (let dr = 0; dr < cell.rowSpan && r + dr < rows.length; dr++) {
        for (let dc = 0; dc < cell.colSpan && c + dc < nCols; dc++) {
          grid[r + dr]![c + dc] =
            dr === 0 && dc === 0 ? { cell, merged: null } : { cell, merged: dr === 0 ? 'h' : 'v' }
        }
      }
      c += cell.colSpan
    }
  })
  const firstHeader =
    rows[0]!.every((c) => c.header) ||
    (!hasFormatting && rows.length > 1 && rows[0]!.every((c) => c.bold))
  // column widths: declared widths where known, else the remaining share of a 640px table
  const declared: Array<number | undefined> = Array.from({ length: nCols }, () => undefined)
  rows.forEach((row, r) => {
    let c = 0
    for (const cell of row) {
      while (c < nCols && grid[r]![c]?.cell !== cell) c++
      if (cell.widthEmu && cell.colSpan === 1 && c < nCols && declared[c] === undefined)
        declared[c] = cell.widthEmu
      c += cell.colSpan
    }
  })
  const defaultW = 640 * EMU_PER_PX
  const known = declared.reduce<number>((n, w) => n + (w ?? 0), 0)
  const unknown = declared.filter((w) => w === undefined).length
  const fallbackW = unknown
    ? Math.max(
        60 * EMU_PER_PX,
        Math.round(Math.max(defaultW - known, unknown * 60 * EMU_PER_PX) / unknown),
      )
    : 0
  let colW = declared.map((w) => w ?? fallbackW)
  const total = colW.reduce((a, b) => a + b, 0)
  const maxW = opts.maxWidthEmu
  if (maxW && total > maxW) colW = colW.map((w) => Math.max(1, Math.round((w * maxW) / total)))
  const rowH = rows.map((row) => {
    const lines = Math.max(1, ...row.map((c) => c.text.split('\n').length))
    const pt = Math.max(...row.map((c) => c.sizePt ?? 18), 12)
    return Math.round((lines * pt * 1.2 + 7.2) * 12700)
  })
  const body = rows
    .map((_, r) => {
      const tcs = grid[r]!.map((slot) =>
        slot
          ? tcXml(
              slot.merged ? { ...slot.cell, text: '', colSpan: 1, rowSpan: 1 } : slot.cell,
              hasFormatting,
              slot.merged,
            )
          : tcXml({ text: '', header: false, colSpan: 1, rowSpan: 1 }, hasFormatting, null),
      ).join('')
      return `<a:tr h="${rowH[r]}">${tcs}</a:tr>`
    })
    .join('')
  const cx = colW.reduce((a, b) => a + b, 0)
  const cy = rowH.reduce((a, b) => a + b, 0)
  const tblPr = hasFormatting
    ? '<a:tblPr/>'
    : `<a:tblPr${firstHeader ? ' firstRow="1"' : ''} bandRow="1"><a:tableStyleId>${BUILTIN_DEFAULT_STYLE}</a:tableStyleId></a:tblPr>`
  const xml =
    '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>${tblPr}<a:tblGrid>${colW.map((w) => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>${body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  return {
    xml,
    rels: [],
    ...(hasFormatting ? {} : { tableStyles: [{ id: BUILTIN_DEFAULT_STYLE }] }),
  }
}
