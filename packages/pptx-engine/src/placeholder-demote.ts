/**
 * Placeholder → free shape, for copy/paste and duplicate.
 *
 * A placeholder's geometry, text size/font/colour and bullets are inherited from the
 * slide's layout and master. Pasted as-is it either lands on a slide that already has
 * that placeholder (two titles) or moves/restyles to wherever the destination layout
 * puts it. PowerPoint demotes the pasted copy to an ordinary text box and bakes the
 * inherited look into the shape (the clipboard carries the full lstStyle) — this does
 * the same from the resolved model: the `<p:ph>` goes, the resolved frame becomes an
 * explicit `<a:xfrm>`, and the inherited paragraph/run defaults become `<a:lstStyle>`
 * levels so the copy looks identical on any slide of any deck.
 */
import type { Paragraph, SlideElement, TextElement, TextRun } from './types'
import { escapeXmlAttr } from './xml-utils'

const PH_RE = /<p:ph\b[^>]*\/>/

const ALIGN: Record<string, string> = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }
const ANCHOR: Record<string, string> = { top: 't', middle: 'ctr', bottom: 'b' }

function hex6(color: string | undefined): string | undefined {
  if (!color) return undefined
  const m = /^#?([0-9a-f]{6})/i.exec(color)
  return m ? m[1]!.toUpperCase() : undefined
}

/** Resolved typeface → declaration (theme references survive, so a destination theme still re-fonts the copy) */
function typefaceTags(run: TextRun): string {
  const src = run.styleSrc?.fontFamily ?? ''
  const theme = /theme major/i.test(src) ? 'mj' : /theme minor/i.test(src) ? 'mn' : null
  if (theme) {
    return `<a:latin typeface="+${theme}-lt"/><a:ea typeface="+${theme}-ea"/><a:cs typeface="+${theme}-cs"/>`
  }
  const face = run.latinFont ?? run.fontFamily
  if (!face) return ''
  const ea = run.eaFont ?? run.fontFamily ?? face
  return (
    `<a:latin typeface="${escapeXmlAttr(face)}"/>` +
    `<a:ea typeface="${escapeXmlAttr(ea)}"/>` +
    `<a:cs typeface="${escapeXmlAttr(run.csFont ?? face)}"/>`
  )
}

function inherited(run: TextRun, key: keyof NonNullable<TextRun['styleSrc']>): boolean {
  const src = run.styleSrc?.[key]
  return !!src && src !== 'run' && src !== 'default' && !run.paraMark && !run.field
}

function defRPrXml(runs: TextRun[]): string {
  const pick = (key: keyof NonNullable<TextRun['styleSrc']>) => runs.find((r) => inherited(r, key))
  const size = pick('fontSize')
  const bold = pick('bold')
  const italic = pick('italic')
  const color = pick('color')
  const face = pick('fontFamily')
  let attrs = ''
  if (size?.fontSize) attrs += ` sz="${Math.round(size.fontSize * 100)}"`
  if (bold) attrs += ` b="${bold.bold ? 1 : 0}"`
  if (italic) attrs += ` i="${italic.italic ? 1 : 0}"`
  const rgb = hex6(color?.color)
  const fill = rgb ? `<a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>` : ''
  const fonts = face ? typefaceTags(face) : ''
  return attrs || fill || fonts ? `<a:defRPr${attrs}>${fill}${fonts}</a:defRPr>` : ''
}

function spacingXml(tag: 'lnSpc' | 'spcBef' | 'spcAft', pts?: number, pct?: number): string {
  if (pts !== undefined) return `<a:${tag}><a:spcPts val="${Math.round(pts * 100)}"/></a:${tag}>`
  if (pct !== undefined) return `<a:${tag}><a:spcPct val="${Math.round(pct * 1000)}"/></a:${tag}>`
  return ''
}

function bulletXml(p: Paragraph): string {
  const b = p.bullet
  if (!b) return ''
  if (b.type === 'none') return '<a:buNone/>'
  const clr = hex6(b.color)
  const buClr = clr ? `<a:buClr><a:srgbClr val="${clr}"/></a:buClr>` : ''
  const size = b.sizePt
    ? `<a:buSzPts val="${Math.round(b.sizePt * 100)}"/>`
    : b.sizePct
      ? `<a:buSzPct val="${Math.round(b.sizePct * 1000)}"/>`
      : ''
  const font = b.font ? `<a:buFont typeface="${escapeXmlAttr(b.font)}"/>` : ''
  if (b.type === 'char' && b.char) {
    return `${buClr}${size}${font}<a:buChar char="${escapeXmlAttr(b.char)}"/>`
  }
  if (b.type === 'number' && b.numType) {
    return `${buClr}${size}${font}<a:buAutoNum type="${escapeXmlAttr(b.numType)}"${
      b.startAt && b.startAt !== 1 ? ` startAt="${b.startAt}"` : ''
    }/>`
  }
  return ''
}

/** `<a:lvlNpPr>` for one indent level, from the resolved paragraph + the inherited run defaults */
function levelXml(level: number, p: Paragraph, runs: TextRun[]): string {
  let attrs = ''
  if (p.marL !== undefined) attrs += ` marL="${Math.round(p.marL)}"`
  if (p.indent !== undefined) attrs += ` indent="${Math.round(p.indent)}"`
  if (p.align) attrs += ` algn="${ALIGN[p.align] ?? 'l'}"`
  const body =
    spacingXml('lnSpc', p.lineExact, p.lineHeight) +
    spacingXml('spcBef', p.spaceBefore, p.spaceBeforePct) +
    spacingXml('spcAft', p.spaceAfter, p.spaceAfterPct) +
    bulletXml(p) +
    defRPrXml(runs)
  return `<a:lvl${level + 1}pPr${attrs}>${body}</a:lvl${level + 1}pPr>`
}

function bakeLstStyle(xml: string, el: TextElement): string {
  const paragraphs = el.text?.paragraphs ?? []
  if (!paragraphs.length) return xml
  const byLevel = new Map<number, Paragraph[]>()
  for (const p of paragraphs) {
    const lvl = Math.min(8, Math.max(0, p.level ?? 0))
    const list = byLevel.get(lvl)
    if (list) list.push(p)
    else byLevel.set(lvl, [p])
  }
  const lst = /<a:lstStyle\s*\/>|<a:lstStyle>[\s\S]*?<\/a:lstStyle>/.exec(xml)
  if (!lst) return xml
  // a shape that already carries its own list style keeps it (it wins over the layout anyway)
  if (/<a:lvl\dpPr/.test(lst[0])) return xml
  const levels = [...byLevel.keys()]
    .sort((a, b) => a - b)
    .map((lvl) => {
      const paras = byLevel.get(lvl)!
      return levelXml(
        lvl,
        paras[0]!,
        paras.flatMap((p) => p.runs),
      )
    })
    .join('')
  return xml.replace(lst[0], `<a:lstStyle>${levels}</a:lstStyle>`)
}

function bakeBodyPr(xml: string, el: TextElement): string {
  const text = el.text
  if (!text) return xml
  const m = /<a:bodyPr\b([^>]*?)(\/?)>/.exec(xml)
  if (!m) return xml
  let attrs = m[1]!
  const add = (name: string, value: string | number) => {
    if (!new RegExp(`\\s${name}=`).test(attrs)) attrs += ` ${name}="${value}"`
  }
  if (text.anchor) add('anchor', ANCHOR[text.anchor] ?? 't')
  if (text.insets) {
    add('lIns', Math.round(text.insets.l))
    add('tIns', Math.round(text.insets.t))
    add('rIns', Math.round(text.insets.r))
    add('bIns', Math.round(text.insets.b))
  }
  if (text.wrap === false) add('wrap', 'none')
  else if (text.wrap === true) add('wrap', 'square')
  let out = `<a:bodyPr${attrs}${m[2]}>`
  let rest = xml.slice(m.index + m[0].length)
  // autofit is inherited from the layout body properties too
  if (m[2] === '/' && text.autofit === 'shrink') {
    out = `<a:bodyPr${attrs}><a:normAutofit/></a:bodyPr>`
  } else if (
    m[2] !== '/' &&
    text.autofit &&
    !/^\s*<a:(noAutofit|normAutofit|spAutoFit)\b/.test(rest)
  ) {
    const tag =
      text.autofit === 'shrink'
        ? 'normAutofit'
        : text.autofit === 'resize'
          ? 'spAutoFit'
          : 'noAutofit'
    // schema order: prstTxWarp, autofit, scene3d, sp3d — autofit goes before 3D / after warp
    const warp = /^\s*<a:prstTxWarp\b[\s\S]*?<\/a:prstTxWarp>/.exec(rest)
    if (warp) rest = rest.slice(0, warp[0].length) + `<a:${tag}/>` + rest.slice(warp[0].length)
    else out += `<a:${tag}/>`
  }
  return xml.slice(0, m.index) + out + rest
}

/**
 * Demote a top-level placeholder `<p:sp>`/`<p:pic>` slice to a free shape. Returns the
 * slice unchanged when it is not a placeholder.
 */
export function demotePlaceholderXml(el: SlideElement, xml: string): string {
  const head = /<p:nvPr>([\s\S]*?)<\/p:nvPr>/.exec(xml)
  if (!head || !PH_RE.test(head[1]!)) return xml
  let out = xml.replace(/<p:nvPr>\s*<p:ph\b[^>]*\/>\s*<\/p:nvPr>/, '<p:nvPr/>')
  if (out === xml) out = xml.replace(PH_RE, '')
  out = out.replace(/<a:spLocks\b[^>]*\bnoGrp="1"[^>]*\/>/, '')
  out = out.replace(/<p:cNvSpPr>\s*<\/p:cNvSpPr>/, '<p:cNvSpPr/>')

  // frame: inherited from the layout unless the shape declares its own
  const spPr = /<p:spPr\b([^>]*?)(\/?)>/.exec(out)
  if (spPr) {
    const hasXfrm = /<p:spPr\b[^>]*>(?:(?!<\/p:spPr>)[\s\S])*?<a:xfrm\b/.test(out)
    const t = el.transform
    let bake = ''
    if (!hasXfrm) {
      const o = t.offset
      const rot = t.rot ? ` rot="${Math.round(t.rot * 60000)}"` : ''
      const flips = `${t.flipH ? ' flipH="1"' : ''}${t.flipV ? ' flipV="1"' : ''}`
      bake = `<a:xfrm${rot}${flips}><a:off x="${Math.round(o.x)}" y="${Math.round(o.y)}"/><a:ext cx="${Math.round(o.cx)}" cy="${Math.round(o.cy)}"/></a:xfrm>`
    }
    const spBody =
      spPr[2] === '/'
        ? ''
        : out.slice(spPr.index + spPr[0].length, out.indexOf('</p:spPr>', spPr.index))
    const hasGeom = /<a:(prstGeom|custGeom)\b/.test(spBody)
    if (!hasGeom && el.type !== 'picture') {
      bake += `<a:prstGeom prst="${escapeXmlAttr((el as TextElement).presetGeometry ?? 'rect')}"><a:avLst/></a:prstGeom>`
    }
    if (bake) {
      if (spPr[2] === '/') {
        out =
          out.slice(0, spPr.index) +
          `<p:spPr${spPr[1]}>${bake}</p:spPr>` +
          out.slice(spPr.index + spPr[0].length)
      } else {
        // xfrm must come first, geometry right after it
        const at = spPr.index + spPr[0].length
        const geomAt = hasXfrm ? out.slice(at).search(/<a:(prstGeom|custGeom)\b/) : -1
        if (hasXfrm && geomAt < 0 && !hasGeom) {
          const x = /<a:xfrm\b[\s\S]*?<\/a:xfrm>/.exec(out.slice(at))
          const insertAt = x ? at + x.index + x[0].length : at
          out = out.slice(0, insertAt) + bake + out.slice(insertAt)
        } else out = out.slice(0, at) + bake + out.slice(at)
      }
    }
  }

  if (el.type === 'text' || el.type === 'shape') {
    out = bakeBodyPr(bakeLstStyle(out, el as TextElement), el as TextElement)
  }
  return out
}

/**
 * Bake the inherited text look (list style levels, body properties) of any text shape into
 * its slice, the way PowerPoint's own clipboard does: the receiving app then draws the same
 * sizes/fonts/colours without having this deck's master.
 */
export function bakeInheritedTextXml(el: SlideElement, xml: string): string {
  if (el.type !== 'text' && el.type !== 'shape') return xml
  return bakeBodyPr(bakeLstStyle(xml, el as TextElement), el as TextElement)
}
