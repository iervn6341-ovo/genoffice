/**
 * Bake a custom table style into explicit cell formatting.
 *
 * A custom table style (an entry in this deck's tableStyles.xml, e.g. Google Slides'
 * "Table_0") exists only inside its deck. PowerPoint's own clipboard drops the style id
 * when it is unknown to the receiving app and the table lands unstyled (measured: pasting
 * a Table ClipFormat whose tableStyleId is custom keeps cell text but loses fills/borders).
 * For the system clipboard the resolved look therefore has to live in the cells:
 * `a:tcPr` fills and borders, run colour and bold. The in-app clipboard keeps the style id
 * (plus its definition) instead, so the table stays style-driven inside GenOffice.
 */
import type { Fill, Stroke, TableElement, TextRun } from './types'
import { builtinTableStyleName } from './table-style'

function srgb(color: string): string {
  const m = /^#?([0-9a-f]{6})([0-9a-f]{2})?/i.exec(color)
  if (!m) return '<a:srgbClr val="000000"/>'
  const alpha = m[2] ? Math.round((parseInt(m[2], 16) / 255) * 100000) : 100000
  return alpha < 100000
    ? `<a:srgbClr val="${m[1]!.toUpperCase()}"><a:alpha val="${alpha}"/></a:srgbClr>`
    : `<a:srgbClr val="${m[1]!.toUpperCase()}"/>`
}

function fillXml(fill: Fill | undefined): string {
  return fill?.type === 'solid' ? `<a:solidFill>${srgb(fill.color)}</a:solidFill>` : ''
}

function lineXml(tag: string, s: Stroke): string {
  const w = Math.max(0, Math.round(s.width))
  const body =
    s.fill.type === 'solid'
      ? `<a:solidFill>${srgb(s.fill.color)}</a:solidFill><a:prstDash val="${s.dash ?? 'solid'}"/>`
      : '<a:noFill/>'
  return `<a:${tag} w="${w}" cap="flat" cmpd="${s.compound ?? 'sng'}" algn="ctr">${body}</a:${tag}>`
}

/** Fill + run formatting of one cell's text that the style supplied */
function bakeRuns(tcXml: string, runs: readonly TextRun[]): string {
  const real = runs.filter((r) => !r.paraMark && !r.rawXml)
  let i = 0
  return tcXml.replace(
    /<a:r>(?:<a:rPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:rPr>))?/g,
    (whole, attrs, close, inner) => {
      const run = real[i++]
      if (!run) return whole
      let a: string = attrs ?? ''
      let body: string = close === '/>' || close === undefined ? '' : (inner ?? '')
      if (run.bold && !/\sb=/.test(a)) a += ' b="1"'
      const needsFill =
        !!run.color && !/<a:(solidFill|gradFill|noFill|pattFill|blipFill)\b/.test(body)
      if (needsFill) {
        const fill = `<a:solidFill>${srgb(run.color!)}</a:solidFill>`
        const ln = /^(<a:ln\b[\s\S]*?<\/a:ln>)/.exec(body)
        body = ln ? ln[1] + fill + body.slice(ln[1]!.length) : fill + body
      }
      return `<a:r><a:rPr${a}>${body}</a:rPr>`
    },
  )
}

function bakeTcPr(
  tcXml: string,
  fill: Fill | undefined,
  borders: TableElement['rows'][number][number]['borders'],
): string {
  const m = /<a:tcPr\b([^>]*?)(\/>|>([\s\S]*?)<\/a:tcPr>)/.exec(tcXml)
  const attrs = m?.[1] ?? ''
  const inner = m && m[2] !== '/>' ? (m[3] ?? '') : ''
  let lines = ''
  for (const [key, tag] of [
    ['l', 'lnL'],
    ['r', 'lnR'],
    ['t', 'lnT'],
    ['b', 'lnB'],
  ] as const) {
    const s = borders?.[key]
    if (s && !new RegExp(`<a:${tag}\\b`).test(inner)) lines += lineXml(tag, s)
  }
  const hasFill = /<a:(solidFill|gradFill|noFill|pattFill|blipFill)\b/.test(
    inner.replace(/<a:ln[LRTB]\b[\s\S]*?<\/a:ln[LRTB]>/g, ''),
  )
  const fillPart = !hasFill ? fillXml(fill) : ''
  if (!lines && !fillPart) return tcXml
  // order: lnL lnR lnT lnB, then the rest of the existing lines, fill, cell3D…
  const own =
    /^((?:<a:ln(?:L|R|T|B|TlToBr|BlToTr)\b[\s\S]*?<\/a:ln(?:L|R|T|B|TlToBr|BlToTr)>)*)([\s\S]*)$/.exec(
      inner,
    )!
  const next = `<a:tcPr${attrs}>${lines}${own[1]}${fillPart}${own[2]}</a:tcPr>`
  if (m) return tcXml.replace(m[0], next)
  return tcXml.replace(/<\/a:tc>$/, `${next}</a:tc>`)
}

/** True when the table references a style only this deck defines */
export function hasCustomTableStyle(el: TableElement): boolean {
  return !!el.styleId && !builtinTableStyleName(el.styleId)
}

export function bakeTableStyleXml(el: TableElement, xml: string): string {
  if (!hasCustomTableStyle(el)) return xml
  const cells = el.rows.flat()
  let idx = 0
  const baked = xml.replace(/<a:tc\b[^>]*?>[\s\S]*?<\/a:tc>/g, (tc) => {
    const cell = cells[idx++]
    if (!cell || cell.merged) return tc
    let out = bakeTcPr(tc, cell.fill, cell.borders)
    if (cell.text)
      out = bakeRuns(
        out,
        cell.text.paragraphs.flatMap((p) => p.runs),
      )
    return out
  })
  return baked.replace(/<a:tableStyleId>[^<]*<\/a:tableStyleId>/, '')
}
