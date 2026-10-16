import { XMLValidator } from 'fast-xml-parser'
import {
  attrsOf,
  childrenOf,
  escapeXmlAttr,
  escapeXmlText,
  findChild,
  findChildren,
  nameOf,
  textOf,
  xmlParser,
  type XNode,
} from './xml-utils'

/**
 * OMML (Office Math Markup) <-> display / authoring support.
 *
 * - ommlFragmentsOf: pull the <m:oMath> fragments out of a paragraph
 * - ommlToMathML: convert OMML to MathML Core for native Chromium rendering
 * - latexToOmml: build OMML from a practical LaTeX subset (insert dialog / gallery)
 */

// ---- fragment extraction ----

/** all <m:oMath> fragments of a paragraph, in document order (oMathPara unwrapped) */
export function ommlFragmentsOf(paragraphXml: string): string[] {
  // `(?=[\s>])` keeps <m:oMathPara> from matching; oMath never nests
  const re = /<m:oMath(?=[\s>])(?:\s[^>]*)?>[\s\S]*?<\/m:oMath>/g
  return paragraphXml.match(re) ?? []
}

/** visible <m:t> leaf texts in document order (the editable formula tokens) */
export function mathTokensOf(xml: string): string[] {
  const tokens: string[] = []
  const re = /<m:t(?:\s[^>]*)?>([\s\S]*?)<\/m:t>/g
  let match: RegExpExecArray | null
  while ((match = re.exec(xml)) !== null) {
    tokens.push(
      match[1]!
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&'),
    )
  }
  return tokens
}

// ---- OMML -> MathML ----

/**
 * MathML Core markup for one or more <m:oMath> fragments (they may be passed
 * concatenated). Returns '' when nothing convertible is found.
 */
export function ommlToMathML(ommlXml: string): string {
  const fragments = ommlFragmentsOf(ommlXml)
  const sources = fragments.length > 0 ? fragments : [ommlXml]
  const parts: string[] = []
  for (const source of sources) {
    if (XMLValidator.validate(source) !== true) continue
    let parsed: XNode[]
    try {
      parsed = xmlParser.parse(source) as XNode[]
    } catch {
      continue
    }
    const root = parsed.find((n) => nameOf(n) === 'm:oMath')
    if (!root) continue
    const body = mmlChildren(childrenOf(root))
    if (body) parts.push(`<math display="block"><mrow>${body}</mrow></math>`)
  }
  return parts.join('')
}

/** child elements that are content (skip m:*Pr property bags and #text glue) */
function contentChildren(node: XNode): XNode[] {
  return childrenOf(node).filter((c) => {
    const name = nameOf(c)
    return !!name && !name.endsWith('Pr')
  })
}

function mmlChildren(nodes: XNode[]): string {
  return nodes.map(mmlOf).join('')
}

/** m:e (or m:num/m:den/...) container -> mrow content */
function mmlSlot(parent: XNode, name: string): string {
  const slot = findChild(parent, name)
  if (!slot) return '<mrow></mrow>'
  return `<mrow>${mmlChildren(contentChildren(slot))}</mrow>`
}

function propVal(node: XNode, prName: string, childName: string): string | undefined {
  const pr = findChild(node, prName)
  if (!pr) return undefined
  const child = findChild(pr, childName)
  if (!child) return undefined
  return attrsOf(child)['m:val']
}

function propOn(node: XNode, prName: string, childName: string): boolean {
  const val = propVal(node, prName, childName)
  return val !== undefined && !['0', 'false', 'off'].includes(val.toLowerCase())
}

function mo(ch: string, extra = ''): string {
  return `<mo${extra}>${escapeXmlText(ch)}</mo>`
}

function mmlOf(node: XNode): string {
  switch (nameOf(node)) {
    case 'm:r':
      return runToMml(node)
    case 'm:f': {
      const type = propVal(node, 'm:fPr', 'm:type')
      const attrs =
        type === 'noBar'
          ? ' linethickness="0"'
          : type === 'lin' || type === 'skw'
            ? ' bevelled="true"'
            : ''
      return `<mfrac${attrs}>${mmlSlot(node, 'm:num')}${mmlSlot(node, 'm:den')}</mfrac>`
    }
    case 'm:sSup':
      return `<msup>${mmlSlot(node, 'm:e')}${mmlSlot(node, 'm:sup')}</msup>`
    case 'm:sSub':
      return `<msub>${mmlSlot(node, 'm:e')}${mmlSlot(node, 'm:sub')}</msub>`
    case 'm:sSubSup':
      return `<msubsup>${mmlSlot(node, 'm:e')}${mmlSlot(node, 'm:sub')}${mmlSlot(node, 'm:sup')}</msubsup>`
    case 'm:sPre':
      return (
        '<mmultiscripts>' +
        mmlSlot(node, 'm:e') +
        '<mprescripts/>' +
        mmlSlot(node, 'm:sub') +
        mmlSlot(node, 'm:sup') +
        '</mmultiscripts>'
      )
    case 'm:rad': {
      const inner = mmlSlot(node, 'm:e')
      if (propOn(node, 'm:radPr', 'm:degHide') || !findChild(node, 'm:deg')) {
        return `<msqrt>${inner}</msqrt>`
      }
      return `<mroot>${inner}${mmlSlot(node, 'm:deg')}</mroot>`
    }
    case 'm:d': {
      const beg = propVal(node, 'm:dPr', 'm:begChr') ?? '('
      const end = propVal(node, 'm:dPr', 'm:endChr') ?? ')'
      const sep = propVal(node, 'm:dPr', 'm:sepChr') ?? '|'
      const slots = childrenOf(node).filter((c) => nameOf(c) === 'm:e')
      const body = slots
        .map((slot) => `<mrow>${mmlChildren(contentChildren(slot))}</mrow>`)
        .join(sep ? mo(sep) : '')
      const open = beg ? mo(beg, ' stretchy="true"') : ''
      const close = end ? mo(end, ' stretchy="true"') : ''
      return `<mrow>${open}${body}${close}</mrow>`
    }
    case 'm:nary': {
      const chr = propVal(node, 'm:naryPr', 'm:chr') ?? '\u222B'
      const limLoc =
        propVal(node, 'm:naryPr', 'm:limLoc') ?? (chr === '\u222B' ? 'subSup' : 'undOvr')
      const subHide = propOn(node, 'm:naryPr', 'm:subHide')
      const supHide = propOn(node, 'm:naryPr', 'm:supHide')
      const op = mo(chr, ' stretchy="false"')
      let scripted = op
      if (!subHide && !supHide) {
        const tag = limLoc === 'undOvr' ? 'munderover' : 'msubsup'
        scripted = `<${tag}>${op}${mmlSlot(node, 'm:sub')}${mmlSlot(node, 'm:sup')}</${tag}>`
      } else if (!subHide) {
        const tag = limLoc === 'undOvr' ? 'munder' : 'msub'
        scripted = `<${tag}>${op}${mmlSlot(node, 'm:sub')}</${tag}>`
      } else if (!supHide) {
        const tag = limLoc === 'undOvr' ? 'mover' : 'msup'
        scripted = `<${tag}>${op}${mmlSlot(node, 'm:sup')}</${tag}>`
      }
      return `<mrow>${scripted}${mmlSlot(node, 'm:e')}</mrow>`
    }
    case 'm:func':
      return `<mrow>${mmlSlot(node, 'm:fName')}<mo>\u2061</mo>${mmlSlot(node, 'm:e')}</mrow>`
    case 'm:limLow':
      return `<munder>${mmlSlot(node, 'm:e')}${mmlSlot(node, 'm:lim')}</munder>`
    case 'm:limUpp':
      return `<mover>${mmlSlot(node, 'm:e')}${mmlSlot(node, 'm:lim')}</mover>`
    case 'm:acc': {
      const chr = propVal(node, 'm:accPr', 'm:chr') ?? '\u0302'
      return `<mover accent="true">${mmlSlot(node, 'm:e')}${mo(chr)}</mover>`
    }
    case 'm:bar': {
      const pos = propVal(node, 'm:barPr', 'm:pos') ?? 'bot'
      const line = pos === 'top' ? '\u00AF' : '\u005F'
      const tag = pos === 'top' ? 'mover' : 'munder'
      return `<${tag}>${mmlSlot(node, 'm:e')}${mo(line, ' stretchy="true"')}</${tag}>`
    }
    case 'm:groupChr': {
      const chr = propVal(node, 'm:groupChrPr', 'm:chr') ?? '\u23DF'
      const pos = propVal(node, 'm:groupChrPr', 'm:pos') ?? 'bot'
      const tag = pos === 'top' ? 'mover' : 'munder'
      return `<${tag}>${mmlSlot(node, 'm:e')}${mo(chr, ' stretchy="true"')}</${tag}>`
    }
    case 'm:m': {
      const jc = attrsOf(
        findChild(
          findChild(
            findChild(findChild(findChild(node, 'm:mPr') ?? {}, 'm:mcs') ?? {}, 'm:mc') ?? {},
            'm:mcPr',
          ) ?? {},
          'm:mcJc',
        ) ?? {},
      )['m:val']
      // MathML Core has no columnalign: the UA centers mtd content, so the
      // alignment (and the inset on that side) is CSS on each cell
      const cellAttrs =
        jc === 'left' || jc === 'right' ? ` style="text-align:${jc};padding-${jc}:0"` : ''
      const rows = childrenOf(node)
        .filter((c) => nameOf(c) === 'm:mr')
        .map((row) => {
          const cells = childrenOf(row)
            .filter((c) => nameOf(c) === 'm:e')
            .map(
              (cell) => `<mtd${cellAttrs}><mrow>${mmlChildren(contentChildren(cell))}</mrow></mtd>`,
            )
            .join('')
          return `<mtr>${cells}</mtr>`
        })
        .join('')
      return `<mtable>${rows}</mtable>`
    }
    case 'm:eqArr': {
      const rows = childrenOf(node)
        .filter((c) => nameOf(c) === 'm:e')
        .map((row) => `<mtr><mtd><mrow>${mmlChildren(contentChildren(row))}</mrow></mtd></mtr>`)
        .join('')
      return `<mtable>${rows}</mtable>`
    }
    case 'm:borderBox': {
      // MathML Core has no menclose; the frame is CSS on the row
      const sides = (['Top', 'Right', 'Bot', 'Left'] as const)
        .filter((side) => !propOn(node, 'm:borderBoxPr', `m:hide${side}`))
        .map((side) => `border-${side === 'Bot' ? 'bottom' : side.toLowerCase()}:0.06em solid`)
      const style = sides.length ? ` style="${sides.join(';')};padding:0.15em"` : ''
      return `<mrow${style}>${mmlChildren(contentChildren(findChild(node, 'm:e') ?? {}))}</mrow>`
    }
    case 'm:box':
    case 'm:phant':
      return mmlSlot(node, 'm:e')
    case 'm:t':
      return runTextToMml(textOf(node), false)
    case undefined:
      return ''
    default:
      // unknown structure: render its content children so nothing disappears
      return mmlChildren(contentChildren(node))
  }
}

function runToMml(run: XNode): string {
  const sty = propVal(run, 'm:rPr', 'm:sty')
  const plain = sty === 'p' || !!findChild(findChild(run, 'm:rPr') ?? {}, 'm:nor')
  let out = ''
  for (const child of childrenOf(run)) {
    if (nameOf(child) === 'm:t') out += runTextToMml(textOf(child), plain)
  }
  return out
}

const OPERATOR_CHARS = new Set('+-−=<>±∓×÷·⋅∙*/!%&|,;:()[]{}′″∞→←↔⇒⇐⇔∈∉⊂⊃∪∩∀∃∧∨¬≤≥≠≈≡∼∝⊥∥°∂∇')

/** classify a text run into mn / mi / mo / mtext tokens */
function runTextToMml(text: string, plain: boolean): string {
  // a lone plain letter would otherwise take the italic single-char mi default
  if (plain)
    return text === ''
      ? ''
      : `<mi${[...text].length === 1 ? ' mathvariant="normal"' : ''}>${escapeXmlText(text)}</mi>`
  let out = ''
  let i = 0
  const chars = [...text]
  while (i < chars.length) {
    const ch = chars[i]!
    if (/[0-9.]/.test(ch)) {
      let num = ''
      while (i < chars.length && /[0-9.]/.test(chars[i]!)) num += chars[i++]!
      out += `<mn>${num}</mn>`
    } else if (/[A-Za-z\u0370-\u03FF\u{1D400}-\u{1D7FF}]/u.test(ch)) {
      out += `<mi>${escapeXmlText(ch)}</mi>`
      i++
    } else if (ch === ' ') {
      i++
    } else if (OPERATOR_CHARS.has(ch)) {
      // parens inside a plain run are literal characters; only m:d wraps
      // (explicit delimiters) may stretch; Word draws a math hyphen-minus as
      // a real minus sign
      const glyph = ch === '-' ? '\u2212' : ch
      out += '()[]{}|'.includes(ch) ? mo(ch, ' stretchy="false"') : mo(glyph)
      i++
    } else {
      out += `<mtext>${escapeXmlText(ch)}</mtext>`
      i++
    }
  }
  return out
}

// ---- LaTeX subset -> OMML ----

const LATEX_SYMBOLS: Record<string, string> = {
  alpha: 'α',
  beta: 'β',
  gamma: 'γ',
  delta: 'δ',
  epsilon: 'ε',
  varepsilon: 'ε',
  zeta: 'ζ',
  eta: 'η',
  theta: 'θ',
  vartheta: 'ϑ',
  iota: 'ι',
  kappa: 'κ',
  lambda: 'λ',
  mu: 'μ',
  nu: 'ν',
  xi: 'ξ',
  pi: 'π',
  rho: 'ρ',
  sigma: 'σ',
  tau: 'τ',
  upsilon: 'υ',
  phi: 'φ',
  varphi: 'ϕ',
  chi: 'χ',
  psi: 'ψ',
  omega: 'ω',
  Gamma: 'Γ',
  Delta: 'Δ',
  Theta: 'Θ',
  Lambda: 'Λ',
  Xi: 'Ξ',
  Pi: 'Π',
  Sigma: 'Σ',
  Upsilon: 'Υ',
  Phi: 'Φ',
  Psi: 'Ψ',
  Omega: 'Ω',
  infty: '∞',
  pm: '±',
  mp: '∓',
  times: '×',
  div: '÷',
  cdot: '⋅',
  ast: '*',
  le: '≤',
  leq: '≤',
  ge: '≥',
  geq: '≥',
  ne: '≠',
  neq: '≠',
  approx: '≈',
  equiv: '≡',
  sim: '∼',
  propto: '∝',
  to: '→',
  rightarrow: '→',
  leftarrow: '←',
  leftrightarrow: '↔',
  Rightarrow: '⇒',
  Leftarrow: '⇐',
  Leftrightarrow: '⇔',
  partial: '∂',
  nabla: '∇',
  in: '∈',
  notin: '∉',
  subset: '⊂',
  supset: '⊃',
  subseteq: '⊆',
  supseteq: '⊇',
  cup: '∪',
  cap: '∩',
  forall: '∀',
  exists: '∃',
  wedge: '∧',
  vee: '∨',
  neg: '¬',
  angle: '∠',
  perp: '⊥',
  parallel: '∥',
  ldots: '…',
  cdots: '⋯',
  vdots: '⋮',
  ddots: '⋱',
  prime: '′',
  circ: '∘',
  degree: '°',
  bullet: '∙',
  star: '⋆',
  emptyset: '∅',
  hbar: 'ℏ',
  ell: 'ℓ',
  Re: 'ℜ',
  Im: 'ℑ',
  aleph: 'ℵ',
  therefore: '∴',
  because: '∵',
}

const LATEX_FUNCTIONS = new Set([
  'sin',
  'cos',
  'tan',
  'cot',
  'sec',
  'csc',
  'sinh',
  'cosh',
  'tanh',
  'coth',
  'arcsin',
  'arccos',
  'arctan',
  'ln',
  'log',
  'exp',
  'max',
  'min',
  'sup',
  'inf',
  'arg',
  'det',
  'gcd',
  'deg',
  'dim',
  'ker',
  'mod',
  'sech',
  'csch',
  'lg',
  'lcm',
  'Pr',
  'hom',
  'limsup',
  'liminf',
])

/** functions whose subscript is a limit (drawn under the name), like Word's Limits & Logs gallery */
const LIMIT_FUNCTIONS = new Set([
  'lim',
  'max',
  'min',
  'sup',
  'inf',
  'limsup',
  'liminf',
  'det',
  'gcd',
  'Pr',
])

const NARY_OPS: Record<string, { chr: string; limLoc: 'undOvr' | 'subSup' }> = {
  sum: { chr: '∑', limLoc: 'undOvr' },
  prod: { chr: '∏', limLoc: 'undOvr' },
  coprod: { chr: '∐', limLoc: 'undOvr' },
  bigcup: { chr: '⋃', limLoc: 'undOvr' },
  bigcap: { chr: '⋂', limLoc: 'undOvr' },
  int: { chr: '∫', limLoc: 'subSup' },
  iint: { chr: '∬', limLoc: 'subSup' },
  iiint: { chr: '∭', limLoc: 'subSup' },
  oint: { chr: '∮', limLoc: 'subSup' },
  oiint: { chr: '∯', limLoc: 'subSup' },
  oiiint: { chr: '∰', limLoc: 'subSup' },
  intclockwise: { chr: '∱', limLoc: 'subSup' },
  varointclockwise: { chr: '∲', limLoc: 'subSup' },
  ointctrclockwise: { chr: '∳', limLoc: 'subSup' },
  bigvee: { chr: '⋁', limLoc: 'undOvr' },
  bigwedge: { chr: '⋀', limLoc: 'undOvr' },
  bigoplus: { chr: '⨁', limLoc: 'undOvr' },
  bigotimes: { chr: '⨂', limLoc: 'undOvr' },
  bigodot: { chr: '⨀', limLoc: 'undOvr' },
  biguplus: { chr: '⨄', limLoc: 'undOvr' },
  bigsqcup: { chr: '⨆', limLoc: 'undOvr' },
}

const ACCENT_CHARS: Record<string, string> = {
  hat: '\u0302',
  bar: '\u0304',
  vec: '\u20D7',
  dot: '\u0307',
  ddot: '\u0308',
  tilde: '\u0303',
  check: '\u030C',
  breve: '\u0306',
  acute: '\u0301',
  grave: '\u0300',
  dddot: '\u20DB',
  mathring: '\u030A',
  overleftarrow: '\u20D6',
  overleftrightarrow: '\u20E1',
  widehat: '\u0302',
  widetilde: '\u0303',
  overrightarrow: '\u20D7',
}

const MATRIX_DELIMS: Record<string, { beg: string; end: string } | null> = {
  matrix: null,
  pmatrix: { beg: '(', end: ')' },
  bmatrix: { beg: '[', end: ']' },
  Bmatrix: { beg: '{', end: '}' },
  vmatrix: { beg: '|', end: '|' },
  Vmatrix: { beg: '‖', end: '‖' },
  cases: { beg: '{', end: '' },
}

/** environments written as an OMML equation array (rows joined by alignment points) instead of a matrix */
const EQARRAY_ENVS = new Set(['cases', 'aligned', 'gathered', 'split'])
MATRIX_DELIMS.aligned = null
MATRIX_DELIMS.gathered = null
MATRIX_DELIMS.split = null

class LatexError extends Error {}

// ---- OMML -> LaTeX subset (inverse of latexToOmml) ----

class UnsupportedOmml extends Error {}

/** lazy inverse lookup tables (the forward tables are declared further down) */
function lazyMap(build: () => Map<string, string>): () => Map<string, string> {
  let map: Map<string, string> | null = null
  return () => (map ??= build())
}

/** char -> canonical \command (first name wins for aliased symbols) */
const symbolCommands = lazyMap(() => {
  const map = new Map<string, string>()
  for (const [name, ch] of Object.entries(LATEX_SYMBOLS)) {
    if (!map.has(ch)) map.set(ch, name)
  }
  return map
})

const naryCommands = lazyMap(() => {
  const map = new Map<string, string>()
  for (const [name, { chr }] of Object.entries(NARY_OPS)) {
    if (!map.has(chr)) map.set(chr, name)
  }
  return map
})

const accentCommands = lazyMap(() => {
  const map = new Map<string, string>()
  for (const [name, ch] of Object.entries(ACCENT_CHARS)) {
    if (!map.has(ch)) map.set(ch, name)
  }
  return map
})

/** delimiter char -> \left/\right token ('' = invisible '.') */
const delimTokens = lazyMap(() => {
  const map = new Map<string, string>([['', '.']])
  for (const [token, ch] of Object.entries(LEFT_RIGHT_CHARS)) {
    if (ch !== '' && !map.has(ch)) map.set(ch, token)
  }
  return map
})

/** "begChr endChr" -> matrix environment name */
const matrixEnvs = lazyMap(() => {
  const map = new Map<string, string>()
  for (const [env, delims] of Object.entries(MATRIX_DELIMS)) {
    if (env !== 'matrix' && delims) map.set(`${delims.beg} ${delims.end}`, env)
  }
  return map
})

/**
 * Decompile a single <m:oMath> fragment back into the LaTeX subset that
 * latexToOmml accepts. Returns null when the formula uses structures the
 * subset cannot express (the caller then keeps token-level editing only).
 * The result is semantically equivalent, not byte-identical: it is only used
 * when the user actually edits the formula, so regeneration is expected.
 */
export function ommlToLatex(ommlXml: string): string | null {
  const fragments = ommlFragmentsOf(ommlXml)
  const sources = fragments.length > 0 ? fragments : [ommlXml]
  if (sources.length !== 1 || XMLValidator.validate(sources[0]!) !== true) return null
  let parsed: XNode[]
  try {
    parsed = xmlParser.parse(sources[0]!) as XNode[]
  } catch {
    return null
  }
  const root = parsed.find((n) => nameOf(n) === 'm:oMath')
  if (!root) return null
  try {
    requireAuthorableProperties(root)
    return latexOfChildren(childrenOf(root))
      .trim()
      .replace(/\s{2,}/g, ' ')
  } catch (e) {
    if (e instanceof UnsupportedOmml) return null
    throw e
  }
}

/** w:rPr children Word writes on every math run / control character; they carry no information beyond the default math font. */
const BENIGN_WORD_RPR = new Set(['w:rFonts', 'w:i', 'w:iCs'])

/** A w:rPr that only restates Cambria Math (and italics) is safe to drop on regeneration. */
function benignWordRPr(rPr: XNode): boolean {
  return childrenOf(rPr).every((child) => {
    const name = nameOf(child)
    if (!name) return true
    if (!BENIGN_WORD_RPR.has(name)) return false
    if (name === 'w:rFonts') {
      const fonts = Object.entries(attrsOf(child)).filter(([key]) => key !== 'w:hint')
      return fonts.every(([, value]) => value === 'Cambria Math')
    }
    const value = attrsOf(child)['w:val']
    return value === undefined || value === '1' || value === 'true'
  })
}

const AUTHORABLE_CONTENT = new Set([
  'm:oMath',
  'm:r',
  'm:t',
  'm:f',
  'm:num',
  'm:den',
  'm:sSup',
  'm:sSub',
  'm:sSubSup',
  'm:sPre',
  'm:e',
  'm:sup',
  'm:sub',
  'm:rad',
  'm:deg',
  'm:d',
  'm:nary',
  'm:func',
  'm:fName',
  'm:limLow',
  'm:limUpp',
  'm:lim',
  'm:acc',
  'm:bar',
  'm:groupChr',
  'm:m',
  'm:mr',
  'm:eqArr',
  'm:borderBox',
  'm:phant',
])

/** Values each property bag may carry; anything else would be lost when the source is regenerated. */
const AUTHORABLE_PROPERTIES: Record<string, readonly string[]> = {
  'm:rPr': ['m:sty', 'm:nor', 'm:scr', 'm:aln'],
  'm:fPr': ['m:type'],
  'm:radPr': ['m:degHide'],
  'm:dPr': ['m:begChr', 'm:endChr', 'm:sepChr', 'm:grow'],
  'm:naryPr': ['m:chr', 'm:limLoc', 'm:subHide', 'm:supHide'],
  'm:accPr': ['m:chr'],
  'm:barPr': ['m:pos'],
  'm:groupChrPr': ['m:chr', 'm:pos', 'm:vertJc'],
  'm:borderBoxPr': [
    'm:hideTop',
    'm:hideBot',
    'm:hideLeft',
    'm:hideRight',
    'm:strikeH',
    'm:strikeV',
    'm:strikeBLTR',
    'm:strikeTLBR',
  ],
  'm:phantPr': ['m:show'],
  'm:sSupPr': [],
  'm:sSubPr': [],
  'm:sSubSupPr': [],
  'm:sPrePr': [],
  'm:funcPr': [],
  'm:limLowPr': [],
  'm:limUppPr': [],
  'm:eqArrPr': [],
  'm:mPr': ['m:mcs'],
  'm:mcPr': ['m:count', 'm:mcJc'],
}

/** Full source editing must not discard Office properties absent from our authoring subset. */
function requireAuthorableProperties(node: XNode): void {
  const nodeName = nameOf(node)
  if (nodeName && !AUTHORABLE_CONTENT.has(nodeName)) throw new UnsupportedOmml()
  for (const child of childrenOf(node)) {
    const name = nameOf(child)
    if (!name) continue
    if (name === 'w:rPr') {
      // Word's own run formatting: only the default math font may be dropped
      if (nodeName !== 'm:r' || !benignWordRPr(child)) throw new UnsupportedOmml()
      continue
    }
    if (name === 'm:mcs') continue // validated through m:mPr below
    if (name.endsWith('Pr')) {
      for (const property of childrenOf(child)) {
        const propertyName = nameOf(property)
        if (!propertyName) continue
        // Control-character formatting: only the default math font
        if (propertyName === 'm:ctrlPr') {
          const rPr = findChild(property, 'w:rPr')
          const extra = childrenOf(property).filter((c) => nameOf(c) && nameOf(c) !== 'w:rPr')
          if (extra.length > 0 || (rPr && !benignWordRPr(rPr))) throw new UnsupportedOmml()
          continue
        }
        if (!AUTHORABLE_PROPERTIES[name]?.includes(propertyName)) throw new UnsupportedOmml()
        const value = attrsOf(property)['m:val']
        if (propertyName === 'm:sty' && !['p', 'i', 'b', 'bi'].includes(value ?? ''))
          throw new UnsupportedOmml()
        if (propertyName === 'm:nor' && value !== undefined && value !== '1' && value !== 'true')
          throw new UnsupportedOmml()
        if (propertyName === 'm:aln' && value !== undefined && value !== '1' && value !== 'true')
          throw new UnsupportedOmml()
        if (propertyName === 'm:grow' && value !== undefined && value !== '1' && value !== 'true')
          throw new UnsupportedOmml()
        if (propertyName === 'm:mcs') requireAuthorableMatrixColumns(property)
      }
    } else requireAuthorableProperties(child)
  }
}

/** m:mcs column specs are only benign when every column keeps Word's default centered alignment. */
function requireAuthorableMatrixColumns(mcs: XNode): void {
  for (const column of childrenOf(mcs)) {
    if (nameOf(column) !== 'm:mc') continue
    for (const pr of childrenOf(column)) {
      if (nameOf(pr) !== 'm:mcPr') continue
      for (const property of childrenOf(pr)) {
        const propertyName = nameOf(property)
        if (!propertyName) continue
        if (propertyName === 'm:count') continue
        if (propertyName === 'm:mcJc' && attrsOf(property)['m:val'] === 'center') continue
        throw new UnsupportedOmml()
      }
    }
  }
}

function latexOfChildren(nodes: XNode[]): string {
  return nodes
    .filter((n) => !!nameOf(n))
    .map(latexOfNode)
    .join('')
}

/** m:e / m:num / ... container -> its content as LaTeX ('' when absent) */
function latexSlot(parent: XNode, name: string): string {
  const slot = findChild(parent, name)
  if (!slot) return ''
  return latexOfChildren(contentChildren(slot))
}

function latexOfNode(node: XNode): string {
  switch (nameOf(node)) {
    case 'm:r':
      return runToLatex(node)
    case 'm:t':
      return charsToLatex(textOf(node))
    case 'm:f': {
      const type = propVal(node, 'm:fPr', 'm:type')
      const num = latexSlot(node, 'm:num')
      const den = latexSlot(node, 'm:den')
      if (type === undefined || type === 'bar') return `\\frac{${num}}{${den}}`
      if (type === 'noBar') return `\\genfrac{}{}{0pt}{}{${num}}{${den}}`
      if (type === 'skw') return `\\nicefrac{${num}}{${den}}`
      if (type === 'lin') return `\\linfrac{${num}}{${den}}`
      throw new UnsupportedOmml()
    }
    case 'm:sPre': {
      const sub = latexSlot(node, 'm:sub')
      const sup = latexSlot(node, 'm:sup')
      return `{}_{${sub}}^{${sup}}{${latexSlot(node, 'm:e')}}`
    }
    case 'm:sSup':
      return `{${latexSlot(node, 'm:e')}}^{${latexSlot(node, 'm:sup')}}`
    case 'm:sSub':
      return `{${latexSlot(node, 'm:e')}}_{${latexSlot(node, 'm:sub')}}`
    case 'm:sSubSup':
      return `{${latexSlot(node, 'm:e')}}_{${latexSlot(node, 'm:sub')}}^{${latexSlot(node, 'm:sup')}}`
    case 'm:rad':
      if (propOn(node, 'm:radPr', 'm:degHide') && latexSlot(node, 'm:deg'))
        throw new UnsupportedOmml()
      if (propOn(node, 'm:radPr', 'm:degHide') || !latexSlot(node, 'm:deg')) {
        return `\\sqrt{${latexSlot(node, 'm:e')}}`
      }
      return `\\sqrt[${latexSlot(node, 'm:deg')}]{${latexSlot(node, 'm:e')}}`
    case 'm:d':
      return delimiterToLatex(node)
    case 'm:nary': {
      const chr = propVal(node, 'm:naryPr', 'm:chr') ?? '∫'
      const command = naryCommands().get(chr)
      if (!command) throw new UnsupportedOmml()
      const location = propVal(node, 'm:naryPr', 'm:limLoc')
      if (location !== undefined && location !== 'undOvr' && location !== 'subSup')
        throw new UnsupportedOmml()
      const explicit =
        location !== undefined && location !== NARY_OPS[command]?.limLoc
          ? location === 'undOvr'
            ? '\\limits'
            : '\\nolimits'
          : ''
      for (const slot of ['sub', 'sup']) {
        if (propOn(node, 'm:naryPr', `m:${slot}Hide`) && latexSlot(node, `m:${slot}`))
          throw new UnsupportedOmml()
      }
      const sub = propOn(node, 'm:naryPr', 'm:subHide') ? '' : `_{${latexSlot(node, 'm:sub')}}`
      const sup = propOn(node, 'm:naryPr', 'm:supHide') ? '' : `^{${latexSlot(node, 'm:sup')}}`
      return `\\${command}${explicit}${sub}${sup} {${latexSlot(node, 'm:e')}}`
    }
    case 'm:func': {
      const functionName = findChild(node, 'm:fName')
      if (!functionName) throw new UnsupportedOmml()
      const arg = `{${latexSlot(node, 'm:e')}}`
      const parts = contentChildren(functionName)
      const scripted = parts.length === 1 ? scriptedFunctionName(parts[0]!) : null
      if (scripted) return `${scripted} ${arg}`
      if (!hasOnlyPlainRuns(functionName)) throw new UnsupportedOmml()
      const name = plainTextOfRuns(functionName)
      if (LATEX_FUNCTIONS.has(name)) return `\\${name} ${arg}`
      if (name === 'lim') return `\\operatorname{lim} ${arg}`
      if (/^[A-Za-z]+$/.test(name)) return `\\operatorname{${name}} ${arg}`
      throw new UnsupportedOmml()
    }
    case 'm:limLow': {
      const base = findChild(node, 'm:e')
      if (base && hasOnlyPlainRuns(base) && LIMIT_FUNCTIONS.has(plainTextOfRuns(base)))
        return `${functionCommand(plainTextOfRuns(base))}_{${latexSlot(node, 'm:lim')}}`
      return `\\underset{${latexSlot(node, 'm:lim')}}{${latexSlot(node, 'm:e')}}`
    }
    case 'm:limUpp':
      return `\\overset{${latexSlot(node, 'm:lim')}}{${latexSlot(node, 'm:e')}}`
    case 'm:acc': {
      const chr = propVal(node, 'm:accPr', 'm:chr') ?? '̂'
      const command = accentCommands().get(chr)
      if (!command) throw new UnsupportedOmml()
      return `\\${command}{${latexSlot(node, 'm:e')}}`
    }
    case 'm:bar': {
      const pos = propVal(node, 'm:barPr', 'm:pos') ?? 'bot'
      if (pos !== 'top' && pos !== 'bot') throw new UnsupportedOmml()
      return `\\${pos === 'top' ? 'overline' : 'underline'}{${latexSlot(node, 'm:e')}}`
    }
    case 'm:groupChr': {
      const chr = propVal(node, 'm:groupChrPr', 'm:chr') ?? '⏟'
      const pos = propVal(node, 'm:groupChrPr', 'm:pos') ?? 'bot'
      const vertJc = propVal(node, 'm:groupChrPr', 'm:vertJc')
      if (vertJc !== undefined && vertJc !== 'bot') throw new UnsupportedOmml()
      const found = Object.entries(GROUP_CHARS).find(
        ([, group]) => group.chr === chr && group.pos === pos,
      )
      if (!found) throw new UnsupportedOmml()
      return `\\${found[0]}{${latexSlot(node, 'm:e')}}`
    }
    case 'm:m':
      return `\\begin{matrix} ${matrixBody(node)} \\end{matrix}`
    case 'm:eqArr':
      return `\\begin{aligned} ${matrixBody(node)} \\end{aligned}`
    case 'm:borderBox': {
      const content = latexSlot(node, 'm:e')
      if (
        !findChild(node, 'm:borderBoxPr') ||
        childrenOf(findChild(node, 'm:borderBoxPr')!).length === 0
      )
        return `\\boxed{${content}}`
      const on = (name: string) => propOn(node, 'm:borderBoxPr', `m:${name}`)
      const hidden = ['hideTop', 'hideBot', 'hideLeft', 'hideRight'].every(on)
      const strikes = ['strikeH', 'strikeV', 'strikeBLTR', 'strikeTLBR'].filter(on)
      if (!hidden || strikes.length === 0) throw new UnsupportedOmml()
      const key = strikes.join('+')
      const command: Record<string, string> = {
        strikeTLBR: 'cancel',
        strikeBLTR: 'bcancel',
        'strikeBLTR+strikeTLBR': 'xcancel',
        strikeH: 'hcancel',
        strikeV: 'vcancel',
      }
      if (!command[key]) throw new UnsupportedOmml()
      return `\\${command[key]}{${content}}`
    }
    case 'm:phant': {
      if (propVal(node, 'm:phantPr', 'm:show') !== '0') throw new UnsupportedOmml()
      return `\\phantom{${latexSlot(node, 'm:e')}}`
    }
    case 'm:box':
      throw new UnsupportedOmml()
    case undefined:
      return ''
    default:
      throw new UnsupportedOmml()
  }
}

/** rows of an m:m / m:eqArr as "a & b \\ c & d" */
let eqArrDepth = 0

function matrixBody(node: XNode): string {
  if (nameOf(node) !== 'm:eqArr') return matrixRows(node)
  eqArrDepth++
  try {
    return matrixRows(node)
  } finally {
    eqArrDepth--
  }
}

function matrixRows(node: XNode): string {
  const rowName = findChildren(node, 'm:mr').length > 0 ? 'm:mr' : 'm:e'
  const rows = childrenOf(node)
    .filter((c) => nameOf(c) === rowName)
    .map((row) =>
      rowName === 'm:mr'
        ? childrenOf(row)
            .filter((c) => nameOf(c) === 'm:e')
            .map((cell) => latexOfChildren(contentChildren(cell)))
            .join(' & ')
        : latexOfChildren(contentChildren(row)),
    )
  return rows.join(' \\\\ ')
}

function functionCommand(name: string): string {
  if (LATEX_FUNCTIONS.has(name) || name === 'lim') return `\\${name}`
  if (/^[A-Za-z]+$/.test(name)) return `\\operatorname{${name}}`
  throw new UnsupportedOmml()
}

/** `\\sin^{-1}`, `\\log_{2}`, `\\lim_{n\\to\\infty}`: the m:fName structure around a plain function name */
function scriptedFunctionName(node: XNode): string | null {
  const kind = nameOf(node)
  if (
    kind !== 'm:sSup' &&
    kind !== 'm:sSub' &&
    kind !== 'm:sSubSup' &&
    kind !== 'm:limLow' &&
    kind !== 'm:limUpp'
  )
    return null
  const base = findChild(node, 'm:e')
  if (!base || !hasOnlyPlainRuns(base)) throw new UnsupportedOmml()
  const head = functionCommand(plainTextOfRuns(base))
  if (kind === 'm:sSup') return `${head}^{${latexSlot(node, 'm:sup')}}`
  if (kind === 'm:sSub') return `${head}_{${latexSlot(node, 'm:sub')}}`
  if (kind === 'm:limLow') return `${head}_{${latexSlot(node, 'm:lim')}}`
  if (kind === 'm:limUpp') return `${head}^{${latexSlot(node, 'm:lim')}}`
  return `${head}_{${latexSlot(node, 'm:sub')}}^{${latexSlot(node, 'm:sup')}}`
}

function delimiterToLatex(node: XNode): string {
  const beg = propVal(node, 'm:dPr', 'm:begChr') ?? '('
  const end = propVal(node, 'm:dPr', 'm:endChr') ?? ')'
  const slots = childrenOf(node).filter((c) => nameOf(c) === 'm:e')
  if (slots.length === 0) throw new UnsupportedOmml()
  const inner = contentChildren(slots[0]!)

  if (slots.length === 1) {
    // \binom: parenthesised single noBar fraction
    if (
      beg === '(' &&
      end === ')' &&
      inner.length === 1 &&
      nameOf(inner[0]!) === 'm:f' &&
      propVal(inner[0]!, 'm:fPr', 'm:type') === 'noBar'
    ) {
      return `\\binom{${latexSlot(inner[0]!, 'm:num')}}{${latexSlot(inner[0]!, 'm:den')}}`
    }
    // matrix environments: delimiters wrapping a single matrix / equation array
    const innerName = inner.length === 1 ? nameOf(inner[0]!) : undefined
    if (innerName === 'm:m' || innerName === 'm:eqArr') {
      const env = matrixEnvs().get(`${beg} ${end}`)
      // an equation array only becomes `cases`; any other pairing stays a \\left...\\right group
      if (env && (innerName === 'm:m' || env === 'cases'))
        return `\\begin{${env}} ${matrixBody(inner[0]!)} \\end{${env}}`
    }
  }
  const begTok = delimTokens().get(beg)
  const endTok = delimTokens().get(end)
  if (begTok === undefined || endTok === undefined) throw new UnsupportedOmml()
  let joiner = ' '
  if (slots.length > 1) {
    const separator = propVal(node, 'm:dPr', 'm:sepChr') ?? '|'
    const sepTok = delimTokens().get(separator)
    if (sepTok === undefined || separator === '') throw new UnsupportedOmml()
    joiner = ` \\middle${sepTok} `
  }
  const body = slots.map((slot) => latexOfChildren(contentChildren(slot))).join(joiner)
  return `\\left${begTok} ${body} \\right${endTok}`
}

/** concatenated m:t text of all m:r children (for function names / lim) */
function hasOnlyPlainRuns(node: XNode): boolean {
  const runs = childrenOf(node)
  return (
    runs.length > 0 &&
    runs.every((run) => {
      if (nameOf(run) !== 'm:r') return false
      return (
        propVal(run, 'm:rPr', 'm:sty') === 'p' ||
        !!findChild(findChild(run, 'm:rPr') ?? {}, 'm:nor')
      )
    })
  )
}

function plainTextOfRuns(node: XNode | undefined): string {
  if (!node) return ''
  let out = ''
  for (const run of findChildren(node, 'm:r')) {
    for (const child of childrenOf(run)) {
      if (nameOf(child) === 'm:t') out += textOf(child)
    }
  }
  return out
}

function runToLatex(run: XNode): string {
  const rPr = findChild(run, 'm:rPr')
  const sty = propVal(run, 'm:rPr', 'm:sty')
  const scr = propVal(run, 'm:rPr', 'm:scr')
  const nor = !!(rPr && findChild(rPr, 'm:nor'))
  const aln = !!(rPr && findChild(rPr, 'm:aln'))
  let text = ''
  for (const child of childrenOf(run)) {
    if (nameOf(child) === 'm:t') text += textOf(child)
  }
  // an alignment point is only meaningful inside an equation array row
  if (aln && eqArrDepth === 0) throw new UnsupportedOmml()
  const prefix = aln ? '& ' : ''
  if (text === '') return aln ? prefix : ''
  if (scr !== undefined) {
    const command = fontCommandFor(scr, sty)
    if (!command || nor || /[{}\\]/.test(text)) throw new UnsupportedOmml()
    return `${prefix}\\${command}{${text}}`
  }
  if (sty === 'b' || sty === 'bi') {
    if (nor || /[{}\\]/.test(text)) throw new UnsupportedOmml()
    return `${prefix}\\${sty === 'b' ? 'mathbf' : 'boldsymbol'}{${text}}`
  }
  const plain = sty === 'p' || nor
  if (!plain) return prefix + charsToLatex(text)
  const trimmed = text.trim()
  if (trimmed === '') return `${prefix} `
  if (LATEX_FUNCTIONS.has(trimmed)) return `${prefix}\\${trimmed} `
  if (trimmed === 'lim') return `${prefix}\\lim `
  if (/[{}\\]/.test(text)) throw new UnsupportedOmml()
  return `${prefix}\\text{${text}}`
}

/** parser-special characters that need escaping in ordinary math text */
const CHAR_ESCAPES: Record<string, string> = {
  '{': '\\{ ',
  '}': '\\} ',
  _: '\\_ ',
  '^': '\\^ ',
  '&': '\\& ',
  '%': '\\% ',
  $: '\\$ ',
  '#': '\\# ',
}

function charsToLatex(text: string): string {
  let out = ''
  for (const ch of text) {
    if (ch === '\\' || ch === '\n') throw new UnsupportedOmml()
    if (ch in CHAR_ESCAPES) {
      out += CHAR_ESCAPES[ch]
      continue
    }
    const command = symbolCommands().get(ch)
    out += command ? `\\${command} ` : ch
  }
  return out
}

interface LatexParser {
  src: string
  pos: number
}

/**
 * Convert a practical LaTeX subset into OMML (the children of <m:oMath>).
 * Supported: \frac \binom \sqrt[n] ^ _ \sum/\int/... with limits,
 * \left(\right), matrix environments, accents, greek letters and common
 * symbols, \text{}, function names. Throws Error (with a Chinese message)
 * on syntax the subset does not cover.
 */
export function latexToOmml(latex: string): string {
  const parser: LatexParser = { src: latex, pos: 0 }
  const out = parseSequence(parser, () => parser.pos >= parser.src.length)
  if (parser.pos < parser.src.length) {
    throw new LatexError(`Cannot parse: "${parser.src.slice(parser.pos, parser.pos + 12)}"`)
  }
  return out
}

/** OMML for a display formula paragraph created by the editor */
export function mathParagraphXml(
  omml: string,
  align: 'left' | 'center' | 'right' = 'center',
): string {
  const jc = align === 'center' ? '' : `<w:pPr><w:jc w:val="${align}"/></w:pPr>`
  return (
    `<w:p>${jc}<m:oMathPara><m:oMathParaPr><m:jc m:val="${align === 'center' ? 'center' : align}"/></m:oMathParaPr>` +
    `<m:oMath>${omml}</m:oMath></m:oMathPara></w:p>`
  )
}

function mathRun(text: string, plain = false): string {
  if (text === '') return ''
  const rPr = plain ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''
  return `<m:r>${rPr}<m:t xml:space="preserve">${escapeXmlText(text)}</m:t></m:r>`
}

/** run with Word's math style (m:sty) and/or script family (m:scr), in schema order */
function styledRun(text: string, style: { sty?: string; scr?: string }): string {
  if (text === '') return ''
  const rPr =
    `<m:rPr>${style.scr ? `<m:scr m:val="${style.scr}"/>` : ''}` +
    `${style.sty ? `<m:sty m:val="${style.sty}"/>` : ''}</m:rPr>`
  return `<m:r>${rPr}<m:t xml:space="preserve">${escapeXmlText(text)}</m:t></m:r>`
}

/** LaTeX font command -> OMML run style */
const FONT_COMMANDS: Record<string, { sty?: string; scr?: string }> = {
  mathbf: { sty: 'b' },
  boldsymbol: { sty: 'bi' },
  bm: { sty: 'bi' },
  mathit: { sty: 'i' },
  mathcal: { scr: 'script', sty: 'p' },
  mathscr: { scr: 'script', sty: 'p' },
  mathfrak: { scr: 'fraktur', sty: 'p' },
  mathbb: { scr: 'double-struck', sty: 'p' },
  mathsf: { scr: 'sans-serif', sty: 'p' },
  mathtt: { scr: 'monospace', sty: 'p' },
}

/** command for an OMML (scr, sty) pair, or null when the pair is not authorable */
function fontCommandFor(scr: string | undefined, sty: string | undefined): string | null {
  for (const [command, style] of Object.entries(FONT_COMMANDS)) {
    if (command === 'mathscr' || command === 'bm') continue // aliases of mathcal / boldsymbol
    if (style.scr === scr && (style.scr ? (sty ?? 'p') === 'p' : style.sty === sty)) return command
  }
  return null
}

function peek(p: LatexParser): string {
  return p.src[p.pos] ?? ''
}

function skipSpaces(p: LatexParser): void {
  while (/\s/.test(peek(p))) p.pos++
}

/** \name control sequence at cursor (cursor just after the backslash) */
function readControlName(p: LatexParser): string {
  const m = /^[A-Za-z]+/.exec(p.src.slice(p.pos))
  if (m) {
    p.pos += m[0].length
    return m[0]
  }
  const ch = p.src[p.pos] ?? ''
  p.pos++
  return ch
}

/** required {...} group, or — LaTeX semantics — exactly one token, as OMML */
function parseGroup(p: LatexParser): string {
  skipSpaces(p)
  if (peek(p) === '{') {
    p.pos++
    const out = parseSequence(p, () => peek(p) === '}')
    if (peek(p) !== '}') throw new LatexError('Missing matching }')
    p.pos++
    return out
  }
  if (peek(p) === '\\') {
    p.pos++
    return parseControl(p)
  }
  const ch = peek(p)
  if (ch === '' || '{}^_&'.includes(ch)) throw new LatexError('An argument is required here')
  p.pos++
  return mathRun(ch)
}

/** raw text of a {...} group (for \text / \begin names) */
function readBraceText(p: LatexParser): string {
  skipSpaces(p)
  if (peek(p) !== '{') throw new LatexError('Expected { here')
  p.pos++
  let depth = 1
  let out = ''
  while (p.pos < p.src.length) {
    const ch = p.src[p.pos++]
    if (ch === '{') depth++
    else if (ch === '}') {
      depth--
      if (depth === 0) return out
    }
    if (depth > 0) out += ch
  }
  throw new LatexError('Missing matching }')
}

/** sequence until `stop` says done; applies ^/_ scripts to the previous atom */
function parseSequence(p: LatexParser, stop: () => boolean): string {
  const atoms: string[] = []
  for (;;) {
    skipSpaces(p)
    if (p.pos >= p.src.length || stop()) break
    const ch = peek(p)
    if (ch === '^' || ch === '_') {
      p.pos++
      const script = parseGroup(p)
      const other = peek(p)
      const base = atoms.pop() ?? mathRun('')
      if (base === '') {
        // `{}_{a}^{b}x`: scripts attached to nothing are left scripts of the next atom
        let sub = ch === '_' ? script : ''
        let sup = ch === '^' ? script : ''
        if ((other === '^' || other === '_') && other !== ch) {
          p.pos++
          const second = parseGroup(p)
          if (other === '_') sub = second
          else sup = second
        }
        skipSpaces(p)
        const operand = p.pos >= p.src.length || stop() ? '' : parseScriptOperand(p)
        atoms.push(
          `<m:sPre><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup><m:e>${operand}</m:e></m:sPre>`,
        )
        continue
      }
      if ((other === '^' || other === '_') && other !== ch) {
        p.pos++
        const second = parseGroup(p)
        const sub = ch === '_' ? script : second
        const sup = ch === '^' ? script : second
        atoms.push(
          `<m:sSubSup><m:e>${base}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`,
        )
      } else if (ch === '^') {
        atoms.push(`<m:sSup><m:e>${base}</m:e><m:sup>${script}</m:sup></m:sSup>`)
      } else {
        atoms.push(`<m:sSub><m:e>${base}</m:e><m:sub>${script}</m:sub></m:sSub>`)
      }
      continue
    }
    atoms.push(parseAtom(p))
  }
  return atoms.join('')
}

/** operand of a left-script: one braced group, command or single character */
function parseScriptOperand(p: LatexParser): string {
  const ch = peek(p)
  if (ch === '{' || ch === '\\') return parseAtom(p)
  if (ch === '' || '}^_&'.includes(ch)) return ''
  p.pos++
  return mathRun(ch)
}

/** one atom: control sequence, group, or a run of ordinary characters */
function parseAtom(p: LatexParser): string {
  skipSpaces(p)
  const ch = peek(p)
  if (ch === '') return ''
  if (ch === '{') return parseGroup(p)
  if (ch === '}') throw new LatexError('Unexpected }')
  if (ch === '\\') {
    p.pos++
    return parseControl(p)
  }
  // ordinary characters up to the next special char (kept as one m:r)
  let text = ''
  while (p.pos < p.src.length && !'\\{}^_&'.includes(peek(p)) && peek(p) !== '\n') {
    text += p.src[p.pos++]
  }
  if (text === '') throw new LatexError(`Cannot parse: "${ch}"`)
  // a following script binds to the last character only ("ab^2" = a·b²):
  // rewind so it becomes its own atom on the next pass
  const chars = [...text]
  if ((peek(p) === '^' || peek(p) === '_') && chars.length > 1) {
    p.pos -= chars[chars.length - 1]!.length
    return mathRun(chars.slice(0, -1).join(''))
  }
  return mathRun(text)
}

function naryOmml(p: LatexParser, chr: string, defaultLoc: 'undOvr' | 'subSup'): string {
  let limLoc = defaultLoc
  skipSpaces(p)
  if (p.src.startsWith('\\limits', p.pos)) {
    p.pos += '\\limits'.length
    limLoc = 'undOvr'
  } else if (p.src.startsWith('\\nolimits', p.pos)) {
    p.pos += '\\nolimits'.length
    limLoc = 'subSup'
  }
  let sub = ''
  let sup = ''
  for (let k = 0; k < 2; k++) {
    skipSpaces(p)
    const ch = peek(p)
    if (ch === '_' && sub === '') {
      p.pos++
      sub = parseGroup(p)
    } else if (ch === '^' && sup === '') {
      p.pos++
      sup = parseGroup(p)
    } else break
  }
  // operand: the next {...} group when present, else the nary stands alone
  skipSpaces(p)
  const operand = peek(p) === '{' ? parseGroup(p) : ''
  const pr =
    `<m:naryPr><m:chr m:val="${escapeXmlAttr(chr)}"/><m:limLoc m:val="${limLoc}"/>` +
    (sub === '' ? '<m:subHide m:val="1"/>' : '') +
    (sup === '' ? '<m:supHide m:val="1"/>' : '') +
    '</m:naryPr>'
  return (
    `<m:nary>${pr}` +
    (sub === '' ? '' : `<m:sub>${sub}</m:sub>`) +
    (sup === '' ? '' : `<m:sup>${sup}</m:sup>`) +
    `<m:e>${operand}</m:e></m:nary>`
  )
}

function matrixOmml(p: LatexParser, env: string): string {
  const delims = MATRIX_DELIMS[env]
  const rows: string[][] = [[]]
  let done = false
  while (!done) {
    const cell = parseSequence(p, () => {
      if (peek(p) === '&') return true
      if (p.src.startsWith('\\\\', p.pos)) return true
      if (atCommand(p, 'end')) return true
      return false
    })
    rows[rows.length - 1]!.push(cell)
    if (peek(p) === '&') {
      p.pos++
    } else if (p.src.startsWith('\\\\', p.pos)) {
      p.pos += 2
      rows.push([])
    } else if (atCommand(p, 'end')) {
      p.pos += 4
      const closing = readBraceText(p)
      if (closing !== env) throw new LatexError(`\\end{${closing}} does not match \\begin{${env}}`)
      done = true
    } else {
      throw new LatexError(`\\begin{${env}} is missing \\end{${env}}`)
    }
  }
  const kept = rows.filter((row) => row.length > 1 || row[0] !== '')
  // alignment points of an equation array are empty runs flagged m:aln, one before every `&` cell
  const aligned = (cell: string) =>
    `<m:r><m:rPr><m:aln/></m:rPr><m:t xml:space="preserve"></m:t></m:r>${cell}`
  const matrix = EQARRAY_ENVS.has(env)
    ? `<m:eqArr>${kept
        .map(
          (row) =>
            `<m:e>${row.map((cell, index) => (index === 0 ? cell : aligned(cell))).join('')}</m:e>`,
        )
        .join('')}</m:eqArr>`
    : `<m:m>${kept
        .map((row) => `<m:mr>${row.map((cell) => `<m:e>${cell}</m:e>`).join('')}</m:mr>`)
        .join('')}</m:m>`
  if (!delims) return matrix
  return (
    '<m:d><m:dPr>' +
    `<m:begChr m:val="${escapeXmlAttr(delims.beg)}"/>` +
    `<m:endChr m:val="${escapeXmlAttr(delims.end)}"/>` +
    `</m:dPr><m:e>${matrix}</m:e></m:d>`
  )
}

const LEFT_RIGHT_CHARS: Record<string, string> = {
  '(': '(',
  ')': ')',
  '[': '[',
  ']': ']',
  '|': '|',
  '.': '',
  '\\{': '{',
  '\\}': '}',
  '\\|': '‖',
  '\\langle': '⟨',
  '\\rangle': '⟩',
  '\\lfloor': '⌊',
  '\\rfloor': '⌋',
  '\\lceil': '⌈',
  '\\rceil': '⌉',
  '\\vert': '|',
  '\\Vert': '‖',
  '\\lbrace': '{',
  '\\rbrace': '}',
  '\\lbrack': '[',
  '\\rbrack': ']',
  '\\llbracket': '⟦',
  '\\rrbracket': '⟧',
  '\\ulcorner': '⌜',
  '\\urcorner': '⌝',
  '\\llcorner': '⌞',
  '\\lrcorner': '⌟',
  '\\Langle': '⟪',
  '\\Rangle': '⟫',
  '/': '/',
  '\\backslash': '\\',
}

function readDelimiter(p: LatexParser): string {
  skipSpaces(p)
  if (peek(p) === '\\') {
    const start = p.pos
    p.pos++
    const name = readControlName(p)
    const key = '\\' + name
    if (key in LEFT_RIGHT_CHARS) return LEFT_RIGHT_CHARS[key]!
    p.pos = start
    throw new LatexError(`Unsupported delimiter: \\${name}`)
  }
  const ch = peek(p)
  if (ch in LEFT_RIGHT_CHARS) {
    p.pos++
    return LEFT_RIGHT_CHARS[ch]!
  }
  throw new LatexError(`Unsupported delimiter: "${ch}"`)
}

/** true when the control sequence `\\name` starts at the cursor (not a longer command such as \\rightarrow) */
function atCommand(p: LatexParser, name: string): boolean {
  const token = `\\${name}`
  return p.src.startsWith(token, p.pos) && !/[A-Za-z]/.test(p.src[p.pos + token.length] ?? '')
}

/**
 * Keep a braced function argument as a native function rather than adjacent runs.
 * Scripts written right after the name (\\sin^{-1}{x}, \\log_{2}{x}, \\lim_{n\\to\\infty}{x}) become
 * the function-name structure Word uses (m:sSup / m:sSub / m:limLow around the name).
 */
function functionOmml(p: LatexParser, name: string): string {
  skipSpaces(p)
  const nameRun = mathRun(name, true)
  const first = peek(p)
  if (first === '_' || first === '^') {
    const save = p.pos
    p.pos++
    const one = parseGroup(p)
    let sub = first === '_' ? one : ''
    let sup = first === '^' ? one : ''
    skipSpaces(p)
    const other = peek(p)
    if ((other === '^' || other === '_') && other !== first) {
      p.pos++
      const two = parseGroup(p)
      if (other === '_') sub = two
      else sup = two
    }
    const limit = LIMIT_FUNCTIONS.has(name) && sub !== '' && sup === ''
    const scripted = limit
      ? `<m:limLow><m:e>${nameRun}</m:e><m:lim>${sub}</m:lim></m:limLow>`
      : sub !== '' && sup !== ''
        ? `<m:sSubSup><m:e>${nameRun}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`
        : sub !== ''
          ? `<m:sSub><m:e>${nameRun}</m:e><m:sub>${sub}</m:sub></m:sSub>`
          : `<m:sSup><m:e>${nameRun}</m:e><m:sup>${sup}</m:sup></m:sSup>`
    skipSpaces(p)
    if (peek(p) === '{')
      return `<m:func><m:fName>${scripted}</m:fName><m:e>${parseGroup(p)}</m:e></m:func>`
    // no argument: a limit keeps its native under-script, anything else stays plain scripts
    if (limit) return scripted
    p.pos = save
    return nameRun
  }
  if (peek(p) !== '{') return nameRun
  return `<m:func><m:fName>${nameRun}</m:fName><m:e>${parseGroup(p)}</m:e></m:func>`
}

const BORDER_BOX_HIDES =
  '<m:hideTop m:val="1"/><m:hideBot m:val="1"/><m:hideLeft m:val="1"/><m:hideRight m:val="1"/>'

/** crossed-out text: a border box with every side hidden and the chosen strike lines */
function strikeBoxOmml(content: string, strikes: string[]): string {
  return (
    `<m:borderBox><m:borderBoxPr>${BORDER_BOX_HIDES}` +
    strikes.map((name) => `<m:${name} m:val="1"/>`).join('') +
    `</m:borderBoxPr><m:e>${content}</m:e></m:borderBox>`
  )
}

/** braces / parens / brackets over or under a group (m:groupChr) */
const GROUP_CHARS: Record<string, { chr: string; pos: 'top' | 'bot' }> = {
  overbrace: { chr: '⏞', pos: 'top' },
  underbrace: { chr: '⏟', pos: 'bot' },
  overparen: { chr: '⏜', pos: 'top' },
  underparen: { chr: '⏝', pos: 'bot' },
  overbracket: { chr: '⎴', pos: 'top' },
  underbracket: { chr: '⎵', pos: 'bot' },
}

const EXTENSIBLE_ARROWS: Record<string, string> = {
  xrightarrow: '→',
  xleftarrow: '←',
  xRightarrow: '⇒',
  xLeftarrow: '⇐',
  xleftrightarrow: '↔',
  xLeftrightarrow: '⇔',
}

function fractionOmml(num: string, den: string, type?: 'noBar' | 'skw' | 'lin'): string {
  const pr = type ? `<m:fPr><m:type m:val="${type}"/></m:fPr>` : ''
  return `<m:f>${pr}<m:num>${num}</m:num><m:den>${den}</m:den></m:f>`
}

function parseControl(p: LatexParser): string {
  const name = readControlName(p)
  if (name in LATEX_SYMBOLS) return mathRun(LATEX_SYMBOLS[name]!)
  if (name in NARY_OPS) return naryOmml(p, NARY_OPS[name]!.chr, NARY_OPS[name]!.limLoc)
  if (name in ACCENT_CHARS) {
    const base = parseGroup(p)
    return `<m:acc><m:accPr><m:chr m:val="${escapeXmlAttr(ACCENT_CHARS[name]!)}"/></m:accPr><m:e>${base}</m:e></m:acc>`
  }
  if (LATEX_FUNCTIONS.has(name)) return functionOmml(p, name)
  if (name in FONT_COMMANDS) return styledRun(readBraceText(p), FONT_COMMANDS[name]!)
  if (name in GROUP_CHARS) {
    const { chr, pos } = GROUP_CHARS[name]!
    return (
      `<m:groupChr><m:groupChrPr><m:chr m:val="${chr}"/><m:pos m:val="${pos}"/></m:groupChrPr>` +
      `<m:e>${parseGroup(p)}</m:e></m:groupChr>`
    )
  }
  if (name in EXTENSIBLE_ARROWS) {
    const above = parseGroup(p)
    return `<m:limUpp><m:e>${mathRun(EXTENSIBLE_ARROWS[name]!)}</m:e><m:lim>${above}</m:lim></m:limUpp>`
  }
  switch (name) {
    case 'frac':
    case 'dfrac':
    case 'tfrac': {
      const num = parseGroup(p)
      const den = parseGroup(p)
      return `<m:f><m:num>${num}</m:num><m:den>${den}</m:den></m:f>`
    }
    case 'nicefrac':
    case 'sfrac':
      return fractionOmml(parseGroup(p), parseGroup(p), 'skw')
    case 'linfrac':
      return fractionOmml(parseGroup(p), parseGroup(p), 'lin')
    case 'genfrac': {
      // \\genfrac{left}{right}{thickness}{style}{numerator}{denominator}
      const left = readBraceText(p).trim()
      const right = readBraceText(p).trim()
      const thickness = readBraceText(p).trim()
      readBraceText(p)
      const num = parseGroup(p)
      const den = parseGroup(p)
      const frac = fractionOmml(
        num,
        den,
        /^0(?:\.0*)?[a-z]*$/i.test(thickness) ? 'noBar' : undefined,
      )
      if (left === '' && right === '') return frac
      const delimiter = (value: string) => (value === '.' ? '' : value)
      return (
        '<m:d><m:dPr>' +
        `<m:begChr m:val="${escapeXmlAttr(delimiter(left))}"/><m:endChr m:val="${escapeXmlAttr(delimiter(right))}"/>` +
        `</m:dPr><m:e>${frac}</m:e></m:d>`
      )
    }
    case 'phantom':
      return `<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr><m:e>${parseGroup(p)}</m:e></m:phant>`
    case 'cancel':
      return strikeBoxOmml(parseGroup(p), ['strikeTLBR'])
    case 'bcancel':
      return strikeBoxOmml(parseGroup(p), ['strikeBLTR'])
    case 'xcancel':
      return strikeBoxOmml(parseGroup(p), ['strikeBLTR', 'strikeTLBR'])
    case 'hcancel':
      return strikeBoxOmml(parseGroup(p), ['strikeH'])
    case 'vcancel':
      return strikeBoxOmml(parseGroup(p), ['strikeV'])
    case 'binom': {
      const top = parseGroup(p)
      const bottom = parseGroup(p)
      return (
        '<m:d><m:e><m:f><m:fPr><m:type m:val="noBar"/></m:fPr>' +
        `<m:num>${top}</m:num><m:den>${bottom}</m:den></m:f></m:e></m:d>`
      )
    }
    case 'sqrt': {
      skipSpaces(p)
      let deg = ''
      if (peek(p) === '[') {
        // ordinary-char runs do not stop at ']', so isolate the degree source
        // up to the closing bracket and parse it as its own sequence
        p.pos++
        const close = p.src.indexOf(']', p.pos)
        if (close === -1) throw new LatexError('Missing matching ]')
        const sub: LatexParser = { src: p.src.slice(p.pos, close), pos: 0 }
        deg = parseSequence(sub, () => sub.pos >= sub.src.length)
        p.pos = close + 1
      }
      const inner = parseGroup(p)
      if (deg === '') {
        return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${inner}</m:e></m:rad>`
      }
      return `<m:rad><m:deg>${deg}</m:deg><m:e>${inner}</m:e></m:rad>`
    }
    case 'boxed':
      return `<m:borderBox><m:e>${parseGroup(p)}</m:e></m:borderBox>`
    case 'overset':
    case 'stackrel':
    case 'underset': {
      const limit = parseGroup(p)
      const base = parseGroup(p)
      const tag = name === 'underset' ? 'm:limLow' : 'm:limUpp'
      return `<${tag}><m:e>${base}</m:e><m:lim>${limit}</m:lim></${tag}>`
    }
    case 'overline':
      return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${parseGroup(p)}</m:e></m:bar>`
    case 'underline':
      return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${parseGroup(p)}</m:e></m:bar>`
    case 'underbrace':
      return (
        '<m:groupChr><m:groupChrPr><m:chr m:val="⏟"/><m:pos m:val="bot"/></m:groupChrPr>' +
        `<m:e>${parseGroup(p)}</m:e></m:groupChr>`
      )
    case 'overbrace':
      return (
        '<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/><m:pos m:val="top"/></m:groupChrPr>' +
        `<m:e>${parseGroup(p)}</m:e></m:groupChr>`
      )
    case 'operatorname':
      return functionOmml(p, readBraceText(p))
    case 'text':
    case 'mathrm':
      return mathRun(readBraceText(p), true)
    case 'lim':
      return functionOmml(p, 'lim')
    case 'left': {
      const beg = readDelimiter(p)
      const slots: string[] = []
      let separator: string | null = null
      for (;;) {
        slots.push(parseSequence(p, () => atCommand(p, 'right') || atCommand(p, 'middle')))
        if (!atCommand(p, 'middle')) break
        p.pos += '\\middle'.length
        const middle = readDelimiter(p)
        if (separator !== null && separator !== middle)
          throw new LatexError('A delimiter pair supports a single separator character')
        separator = middle
      }
      if (!atCommand(p, 'right')) throw new LatexError('\\left is missing a matching \\right')
      p.pos += '\\right'.length
      const end = readDelimiter(p)
      return (
        '<m:d><m:dPr>' +
        `<m:begChr m:val="${escapeXmlAttr(beg)}"/>` +
        (separator !== null ? `<m:sepChr m:val="${escapeXmlAttr(separator)}"/>` : '') +
        `<m:endChr m:val="${escapeXmlAttr(end)}"/>` +
        `</m:dPr>${slots.map((slot) => `<m:e>${slot}</m:e>`).join('')}</m:d>`
      )
    }
    case 'begin': {
      const env = readBraceText(p)
      if (!(env in MATRIX_DELIMS)) throw new LatexError(`Unsupported environment: \\begin{${env}}`)
      return matrixOmml(p, env)
    }
    case ',':
    case ';':
    case ' ':
    case 'quad':
    case 'qquad':
      return mathRun(' ')
    case '\\':
      throw new LatexError('\\\\ is only allowed inside matrix environments')
    case '{':
      return mathRun('{')
    case '}':
      return mathRun('}')
    case '%':
    case '&':
    case '$':
    case '#':
    case '_':
    case '^':
      return mathRun(name)
    default:
      throw new LatexError(`Unsupported command: \\${name}`)
  }
}
