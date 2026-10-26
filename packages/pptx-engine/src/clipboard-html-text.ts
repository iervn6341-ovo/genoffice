/**
 * Rich text from the HTML clipboard rendition (PowerPoint, Word, Pages, a web page) as a
 * text box with its character and paragraph formatting — bold / italic / underline, colour,
 * size, bullets and numbering — instead of a flat run of plain text.
 */
import type { ElementClipboardItem } from './index'
import type { Paragraph, TextRun } from './types'
import { generateParagraphXml } from './generate'

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
  bull: '•',
  middot: '·',
}

function decode(s: string): string {
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

function styleMap(style: string | undefined): Record<string, string> {
  const out: Record<string, string> = {}
  for (const decl of (style ?? '').split(';')) {
    const i = decl.indexOf(':')
    if (i > 0) out[decl.slice(0, i).trim().toLowerCase()] = decl.slice(i + 1).trim()
  }
  return out
}

function cssColor(v: string | undefined): string | undefined {
  if (!v) return undefined
  const s = v.trim().toLowerCase()
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(s)
  if (hex) {
    const h = hex[1]!
    return '#' + (h.length === 3 ? [...h].map((c) => c + c).join('') : h).toUpperCase()
  }
  const rgb = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(s)
  if (rgb) {
    return (
      '#' +
      [rgb[1], rgb[2], rgb[3]]
        .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase()
    )
  }
  return undefined
}

function sizePt(v: string | undefined): number | undefined {
  const m = /^\s*([\d.]+)\s*(pt|px)?\s*$/i.exec(v ?? '')
  if (!m) return undefined
  const n = parseFloat(m[1]!)
  const pt = (m[2] ?? 'px').toLowerCase() === 'px' ? n * 0.75 : n
  return pt >= 4 && pt <= 400 ? Math.round(pt * 2) / 2 : undefined
}

interface Fmt {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  color?: string
  size?: number
}

const BLOCK = new Set([
  'p',
  'div',
  'li',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'tr',
  'blockquote',
  'pre',
])
const SKIP = new Set(['script', 'style', 'head', 'title', 'meta', 'xml'])

/**
 * Returns an item only when the HTML carries formatting worth keeping (emphasis, colour,
 * several differently sized runs, lists); plain text returns null so the caller pastes it as before.
 */
export function htmlTextToItem(
  html: string,
  opts: { widthEmu: number },
): ElementClipboardItem | null {
  const frag = /<!--StartFragment-->([\s\S]*?)<!--EndFragment-->/i.exec(html)?.[1]
  const body = frag ?? /<body\b[^>]*>([\s\S]*?)(?:<\/body>|$)/i.exec(html)?.[1] ?? html
  if (/<table\b/i.test(body)) return null

  const paras: Array<{ runs: TextRun[]; bullet?: 'char' | 'num'; level: number }> = []
  type Para = { runs: TextRun[]; bullet?: 'char' | 'num'; level: number }
  const holder: { cur: Para | null } = { cur: null }
  const stack: Array<{ tag: string; fmt: Fmt }> = []
  const lists: Array<'ul' | 'ol'> = []
  let pendingBullet: 'char' | 'num' | undefined
  let fmt: Fmt = {}
  let skipDepth = 0

  const start = () => {
    const p: Para = {
      runs: [],
      level: Math.max(0, lists.length - 1),
      ...(pendingBullet ? { bullet: pendingBullet } : {}),
    }
    holder.cur = p
    pendingBullet = undefined
    paras.push(p)
  }
  const end = () => {
    holder.cur = null
  }
  const text = (raw: string) => {
    const t = decode(raw.replace(/<!--[\s\S]*?-->/g, '')).replace(/[\r\n\t\f\v ]+/g, ' ')
    if (!t || (!holder.cur && !t.trim())) return
    if (!holder.cur) start()
    const c = holder.cur!
    const run: TextRun = { text: c.runs.length ? t : t.replace(/^ /, '') }
    if (fmt.bold) run.bold = true
    if (fmt.italic) run.italic = true
    if (fmt.underline) run.underline = true
    if (fmt.color) run.color = fmt.color
    if (fmt.size) run.fontSize = fmt.size
    if (run.text) c.runs.push(run)
  }

  const tagRe = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)\b([^>]*)>|([^<]+)/g
  let m: RegExpExecArray | null
  while ((m = tagRe.exec(body))) {
    if (m[4] !== undefined) {
      if (!skipDepth) text(m[4])
      continue
    }
    const closing = m[1] === '/'
    const tag = m[2]!.toLowerCase()
    const attrs = m[3] ?? ''
    if (SKIP.has(tag)) {
      skipDepth += closing ? -1 : attrs.endsWith('/') ? 0 : 1
      if (skipDepth < 0) skipDepth = 0
      continue
    }
    if (skipDepth) continue
    if (tag === 'br') {
      if (holder.cur) holder.cur.runs.push({ text: '\n' })
      else start()
      continue
    }
    if (tag === 'ul' || tag === 'ol') {
      if (closing) lists.pop()
      else lists.push(tag)
      end()
      continue
    }
    if (BLOCK.has(tag)) {
      if (closing) end()
      else {
        end()
        if (tag === 'li') pendingBullet = lists[lists.length - 1] === 'ol' ? 'num' : 'char'
        start()
      }
    }
    if (closing) {
      // pop the matching formatting frame
      for (let i = stack.length - 1; i >= 0; i--) {
        if (stack[i]!.tag === tag) {
          fmt = stack[i]!.fmt
          stack.length = i
          break
        }
      }
      continue
    }
    if (attrs.endsWith('/')) continue
    // formatting tags push a frame (block tags too, so their style closes with them)
    const prev = fmt
    const next: Fmt = { ...fmt }
    const st = styleMap(
      /\bstyle\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(attrs)?.slice(1).find(Boolean),
    )
    if (tag === 'b' || tag === 'strong') next.bold = true
    if (tag === 'i' || tag === 'em') next.italic = true
    if (tag === 'u') next.underline = true
    if (st['font-weight']) next.bold = /bold|[6-9]00/.test(st['font-weight'])
    if (st['font-style']) next.italic = /italic|oblique/.test(st['font-style'])
    if (st['text-decoration'] || st['text-decoration-line']) {
      next.underline = /underline/.test(st['text-decoration'] ?? st['text-decoration-line'] ?? '')
    }
    const color = cssColor(st['color'] ?? /\bcolor\s*=\s*["']?([^"'\s>]+)/i.exec(attrs)?.[1])
    if (color) next.color = color === '#000000' ? undefined : color
    const size = sizePt(st['font-size'])
    if (size) next.size = size
    stack.push({ tag, fmt: prev })
    fmt = next
  }

  const real = paras.filter((p) => p.runs.some((r) => r.text.trim()))
  if (!real.length) return null
  const flat = real.flatMap((p) => p.runs)
  const sizes = new Set(flat.map((r) => r.fontSize).filter(Boolean))
  const worthKeeping =
    flat.some((r) => r.bold || r.italic || r.underline || r.color) ||
    real.some((p) => p.bullet) ||
    sizes.size > 1
  if (!worthKeeping) return null

  const paragraphs: Paragraph[] = real.map((p) => ({
    runs: p.runs.map((r, i) =>
      i === p.runs.length - 1 ? { ...r, text: r.text.replace(/ $/, '') || r.text } : r,
    ),
    ...(p.bullet
      ? {
          bullet:
            p.bullet === 'num'
              ? { type: 'number' as const, numType: 'arabicPeriod' }
              : { type: 'char' as const, char: '•', font: 'Arial' },
          marL: 285750 + p.level * 285750,
          indent: -285750,
          ...(p.level ? { level: p.level } : {}),
        }
      : {}),
  }))
  const lines = paragraphs.reduce((n, p) => n + 1 + p.runs.filter((r) => r.text === '\n').length, 0)
  const maxPt = Math.max(18, ...flat.map((r) => r.fontSize ?? 18))
  const cy = Math.round(lines * maxPt * 1.2 * 12700 + 91440)
  const xml =
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="TextBox 1"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${opts.widthEmu}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphs.map(generateParagraphXml).join('')}</p:txBody></p:sp>`
  return { xml, rels: [] }
}
