import { XMLValidator } from 'fast-xml-parser'
import {
  attrsOf,
  childrenOf,
  escapeXmlAttr,
  escapeXmlText,
  findChild,
  nameOf,
  textOf,
  xmlParser,
  type XNode,
} from './xml-utils'
import { ommlFragmentsOf } from './math'

/**
 * UnicodeMath (Word's "linear" equation format) <-> OMML.
 *
 * `ommlToUnicodeMath` writes the linear form of the structures the gallery produces (fractions, scripts,
 * radicals, n-ary operators with the `▒` operand mark, brackets, function apply, accents, bars, over/under
 * limits, matrices `■(a&b@c&d)`, equation arrays `█(...)`, boxes and quoted text) and returns null for
 * anything it cannot express. `unicodeMathToOmml` reads the same subset back. The two are inverse on the
 * supported subset, which is what Professional <-> Linear conversion needs.
 */

export class UnicodeMathError extends Error {}

const FUNCTION_APPLY = '⁡'
const NARY_OPERAND = '▒' // ▒
const MATRIX_OPEN = '■' // ■
const ARRAY_OPEN = '█' // █
const BOX_OPEN = '▭' // ▭
const INVISIBLE_OPEN = '├' // ├
const INVISIBLE_CLOSE = '┤' // ┤
const BELOW = '┬' // ┬
const ABOVE = '┴' // ┴
const STACK = '¦' // ¦
const SKEWED = '⊘' // ⊘
const LINEAR = '∕' // ∕

const NARY_CHARS = new Set([
  '∑',
  '∏',
  '∐',
  '∫',
  '∬',
  '∭',
  '∮',
  '∯',
  '∰',
  '∱',
  '∲',
  '∳',
  '⋃',
  '⋂',
  '⋁',
  '⋀',
  '⨁',
  '⨂',
  '⨀',
  '⨄',
  '⨆',
])
/** n-ary characters whose limits sit beside the glyph by default (integrals) */
const SIDE_LIMIT_NARY = new Set(['∫', '∬', '∭', '∮', '∯', '∰', '∱', '∲', '∳'])

const FUNCTION_NAMES = new Set([
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
  'sech',
  'csch',
  'arcsin',
  'arccos',
  'arctan',
  'ln',
  'log',
  'lg',
  'exp',
  'lim',
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
  'limsup',
  'liminf',
])

const BRACKET_PAIRS: Record<string, string> = {
  '(': ')',
  '[': ']',
  '{': '}',
  '⟨': '⟩',
  '⌊': '⌋',
  '⌈': '⌉',
  '⟦': '⟧',
  '⌜': '⌝',
  '⟪': '⟫',
}
const CLOSERS = new Set(Object.values(BRACKET_PAIRS))
const SAME_SIDE = new Set(['|', '‖'])

const isCombining = (ch: string): boolean => /[̀-ͯ⃐-⃿]/.test(ch)

// ---------------------------------------------------------------- OMML -> UnicodeMath

class Unsupported extends Error {}

function kids(node: XNode): XNode[] {
  return childrenOf(node).filter((c) => {
    const name = nameOf(c)
    return !!name && !name.endsWith('Pr') && name !== 'w:rPr'
  })
}

function prop(node: XNode, pr: string, child: string): string | undefined {
  const bag = findChild(node, pr)
  const value = bag ? findChild(bag, child) : undefined
  return value ? attrsOf(value)['m:val'] : undefined
}

function propOn(node: XNode, pr: string, child: string): boolean {
  const value = prop(node, pr, child)
  return value !== undefined && !['0', 'false', 'off'].includes(value.toLowerCase())
}

function slot(node: XNode, name: string): string {
  const child = findChild(node, name)
  return child ? sequence(kids(child)) : ''
}

function sequence(nodes: XNode[]): string {
  return nodes.map(item).join('')
}

/** operand text: bare when it is one token, parenthesised otherwise */
function operand(text: string): string {
  if (text === '') return '()'
  if ([...text].length === 1 || /^\d+(?:\.\d+)?$/.test(text)) return text
  return `(${text})`
}

function runText(run: XNode): string {
  let text = ''
  for (const child of childrenOf(run)) if (nameOf(child) === 'm:t') text += textOf(child)
  return text
}

function item(node: XNode): string {
  switch (nameOf(node)) {
    case 'm:r': {
      const raw = runText(node)
      const pr = findChild(node, 'm:rPr')
      const aln = !!(pr && findChild(pr, 'm:aln'))
      const nor = !!(pr && findChild(pr, 'm:nor'))
      const sty = prop(node, 'm:rPr', 'm:sty')
      // ordinary math text ignores spaces in linear form; quoted text keeps them
      const text = nor || sty === 'p' ? raw : raw.replace(/\s+/g, '')
      if (prop(node, 'm:rPr', 'm:scr') !== undefined) throw new Unsupported()
      if ((sty === 'b' || sty === 'bi') && text !== '') throw new Unsupported()
      const body = nor || (sty === 'p' && /\s|[^\p{L}\p{N}]/u.test(text)) ? `"${text}"` : text
      if (/["@&]/.test(text) && !nor) throw new Unsupported()
      return `${aln ? '&' : ''}${body}`
    }
    case 'm:f': {
      const type = prop(node, 'm:fPr', 'm:type')
      const op =
        type === undefined || type === 'bar'
          ? '/'
          : type === 'noBar'
            ? STACK
            : type === 'skw'
              ? SKEWED
              : type === 'lin'
                ? LINEAR
                : null
      if (!op) throw new Unsupported()
      return `${operand(slot(node, 'm:num'))}${op}${operand(slot(node, 'm:den'))}`
    }
    case 'm:sSup':
      return `${operand(slot(node, 'm:e'))}^${operand(slot(node, 'm:sup'))}`
    case 'm:sSub':
      return `${operand(slot(node, 'm:e'))}_${operand(slot(node, 'm:sub'))}`
    case 'm:sSubSup':
      return `${operand(slot(node, 'm:e'))}_${operand(slot(node, 'm:sub'))}^${operand(slot(node, 'm:sup'))}`
    case 'm:rad': {
      const hidden = propOn(node, 'm:radPr', 'm:degHide') || slot(node, 'm:deg') === ''
      const base = slot(node, 'm:e')
      if (hidden) return `√(${base})`
      const degree = slot(node, 'm:deg')
      if (degree === '3') return `∛(${base})`
      if (degree === '4') return `∜(${base})`
      return `√(${degree}&${base})`
    }
    case 'm:nary': {
      const chr = prop(node, 'm:naryPr', 'm:chr') ?? '∫'
      if (!NARY_CHARS.has(chr)) throw new Unsupported()
      const location = prop(node, 'm:naryPr', 'm:limLoc')
      // limits above/below on an integral (or beside a sum) need an explicit hint UnicodeMath lacks
      if (location !== undefined && location !== (SIDE_LIMIT_NARY.has(chr) ? 'subSup' : 'undOvr'))
        throw new Unsupported()
      const sub = propOn(node, 'm:naryPr', 'm:subHide') ? '' : slot(node, 'm:sub')
      const sup = propOn(node, 'm:naryPr', 'm:supHide') ? '' : slot(node, 'm:sup')
      const body = slot(node, 'm:e')
      return (
        chr +
        (sub ? `_${operand(sub)}` : '') +
        (sup ? `^${operand(sup)}` : '') +
        (body ? `${NARY_OPERAND}${operand(body)}` : '')
      )
    }
    case 'm:d': {
      const beg = prop(node, 'm:dPr', 'm:begChr') ?? '('
      const end = prop(node, 'm:dPr', 'm:endChr') ?? ')'
      const sep = prop(node, 'm:dPr', 'm:sepChr') ?? '|'
      const slots = childrenOf(node)
        .filter((c) => nameOf(c) === 'm:e')
        .map((c) => sequence(kids(c)))
      if (slots.length === 0) throw new Unsupported()
      const open = beg === '' ? INVISIBLE_OPEN : beg
      const close = end === '' ? INVISIBLE_CLOSE : end
      const same = SAME_SIDE.has(beg) && beg === end
      if (!same && !(beg in BRACKET_PAIRS) && beg !== '') throw new Unsupported()
      if (!same && !CLOSERS.has(end) && end !== '') throw new Unsupported()
      if (slots.length > 1 && sep !== '|') throw new Unsupported()
      return `${open}${slots.join('|')}${close}`
    }
    case 'm:func': {
      const name = findChild(node, 'm:fName')
      if (!name) throw new Unsupported()
      return `${sequence(kids(name))}${FUNCTION_APPLY}${operand(slot(node, 'm:e'))}`
    }
    case 'm:acc': {
      const chr = prop(node, 'm:accPr', 'm:chr') ?? '̂'
      if (!isCombining(chr)) throw new Unsupported()
      return `${operand(slot(node, 'm:e'))}${chr}`
    }
    case 'm:bar': {
      const pos = prop(node, 'm:barPr', 'm:pos') ?? 'bot'
      return `${operand(slot(node, 'm:e'))}${pos === 'top' ? '̅' : '̲'}`
    }
    case 'm:limLow':
      return `${operand(slot(node, 'm:e'))}${BELOW}${operand(slot(node, 'm:lim'))}`
    case 'm:limUpp':
      return `${operand(slot(node, 'm:e'))}${ABOVE}${operand(slot(node, 'm:lim'))}`
    case 'm:m': {
      const rows = childrenOf(node)
        .filter((c) => nameOf(c) === 'm:mr')
        .map((row) =>
          childrenOf(row)
            .filter((c) => nameOf(c) === 'm:e')
            .map((cell) => sequence(kids(cell)))
            .join('&'),
        )
      return `${MATRIX_OPEN}(${rows.join('@')})`
    }
    case 'm:eqArr': {
      const rows = childrenOf(node)
        .filter((c) => nameOf(c) === 'm:e')
        .map((row) => sequence(kids(row)))
      return `${ARRAY_OPEN}(${rows.join('@')})`
    }
    case 'm:borderBox': {
      const pr = findChild(node, 'm:borderBoxPr')
      if (pr && childrenOf(pr).some((c) => nameOf(c))) throw new Unsupported()
      return `${BOX_OPEN}(${slot(node, 'm:e')})`
    }
    default:
      throw new Unsupported()
  }
}

/** Linear (UnicodeMath) text of one <m:oMath> fragment, or null when it holds a structure outside the subset. */
export function ommlToUnicodeMath(ommlXml: string): string | null {
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
    return sequence(kids(root))
  } catch (error) {
    if (error instanceof Unsupported) return null
    throw error
  }
}

// ---------------------------------------------------------------- UnicodeMath -> OMML

interface Cursor {
  src: string[]
  pos: number
}

const run = (text: string, plain = false): string =>
  text === ''
    ? ''
    : `<m:r>${plain ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''}<m:t xml:space="preserve">${escapeXmlText(text)}</m:t></m:r>`

const norRun = (text: string): string =>
  `<m:r><m:rPr><m:nor/></m:rPr><m:t xml:space="preserve">${escapeXmlText(text)}</m:t></m:r>`

const alignRun = (): string => '<m:r><m:rPr><m:aln/></m:rPr><m:t xml:space="preserve"></m:t></m:r>'

/** characters that end a plain run of ordinary text */
const SPECIAL = new Set([
  '^',
  '_',
  '/',
  STACK,
  SKEWED,
  LINEAR,
  BELOW,
  ABOVE,
  '(',
  ')',
  '[',
  ']',
  '{',
  '}',
  '⟨',
  '⟩',
  '⌊',
  '⌋',
  '⌈',
  '⌉',
  '⟦',
  '⟧',
  '⌜',
  '⌝',
  '⟪',
  '⟫',
  '|',
  '‖',
  '&',
  '@',
  '"',
  '√',
  '∛',
  '∜',
  MATRIX_OPEN,
  ARRAY_OPEN,
  BOX_OPEN,
  INVISIBLE_OPEN,
  INVISIBLE_CLOSE,
  NARY_OPERAND,
  FUNCTION_APPLY,
  ...NARY_CHARS,
])

const peek = (c: Cursor): string => c.src[c.pos] ?? ''

/** parse until one of `stops` (not consumed) or the end of input */
function parseSeq(c: Cursor, stops: ReadonlySet<string>): string {
  let out = ''
  while (c.pos < c.src.length && !stops.has(peek(c))) {
    const before = c.pos
    out += parseTerm(c, stops)
    if (c.pos === before) throw new UnicodeMathError(`Unexpected "${peek(c)}"`)
  }
  return out
}

/** result of an operand: OMML plus whether it came from bare parentheses (a grouping, not a bracket) */
interface Operand {
  omml: string
  grouped: boolean
}

const FRACTION_OPS = new Set(['/', STACK, SKEWED, LINEAR])

function fractionOmml(op: string, num: string, den: string): string {
  const type = op === STACK ? 'noBar' : op === SKEWED ? 'skw' : op === LINEAR ? 'lin' : undefined
  const pr = type ? `<m:fPr><m:type m:val="${type}"/></m:fPr>` : ''
  return `<m:f>${pr}<m:num>${num}</m:num><m:den>${den}</m:den></m:f>`
}

function parseTerm(c: Cursor, stops: ReadonlySet<string>): string {
  let base = parsePrimary(c, stops)
  for (;;) {
    const ch = peek(c)
    if (isCombining(ch)) {
      c.pos++
      base = { omml: accentOmml(ch, base.omml), grouped: false }
    } else if (ch === '^' || ch === '_') {
      c.pos++
      const first = parseScriptOperand(c)
      let sub = ch === '_' ? first : ''
      let sup = ch === '^' ? first : ''
      const other = peek(c)
      if ((other === '^' || other === '_') && other !== ch) {
        c.pos++
        const second = parseScriptOperand(c)
        if (other === '_') sub = second
        else sup = second
      }
      base = { omml: scriptOmml(base.omml, sub, sup), grouped: false }
    } else if (ch === BELOW || ch === ABOVE) {
      c.pos++
      const limit = parseScriptOperand(c)
      const tag = ch === BELOW ? 'm:limLow' : 'm:limUpp'
      base = {
        omml: `<${tag}><m:e>${base.omml}</m:e><m:lim>${limit}</m:lim></${tag}>`,
        grouped: false,
      }
    } else if (ch === FUNCTION_APPLY) {
      // a scripted function name (sin^(-1), log_2, lim┬(n→∞)) applied to its argument
      c.pos++
      const argument = parsePrimaryWithScripts(c, stops, true)
      base = {
        omml: `<m:func><m:fName>${base.omml}</m:fName><m:e>${argument}</m:e></m:func>`,
        grouped: false,
      }
    } else if (FRACTION_OPS.has(ch)) {
      c.pos++
      const denominator = parsePrimaryWithScripts(c, stops)
      base = { omml: fractionOmml(ch, base.omml, denominator), grouped: false }
    } else break
  }
  return finish(base)
}

/** a parenthesised group that never became an operand is a real bracket */
function finish(operandValue: Operand): string {
  return operandValue.grouped
    ? `<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>${operandValue.omml}</m:e></m:d>`
    : operandValue.omml
}

function parsePrimaryWithScripts(c: Cursor, stops: ReadonlySet<string>, single = false): string {
  let base = parsePrimary(c, stops, single)
  for (;;) {
    const ch = peek(c)
    if (isCombining(ch)) {
      c.pos++
      base = { omml: accentOmml(ch, base.omml), grouped: false }
    } else if (ch === '^' || ch === '_') {
      c.pos++
      const first = parseScriptOperand(c)
      let sub = ch === '_' ? first : ''
      let sup = ch === '^' ? first : ''
      const other = peek(c)
      if ((other === '^' || other === '_') && other !== ch) {
        c.pos++
        const second = parseScriptOperand(c)
        if (other === '_') sub = second
        else sup = second
      }
      base = { omml: scriptOmml(base.omml, sub, sup), grouped: false }
    } else break
  }
  return base.omml
}

function scriptOmml(base: string, sub: string, sup: string): string {
  if (sub && sup)
    return `<m:sSubSup><m:e>${base}</m:e><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`
  if (sub) return `<m:sSub><m:e>${base}</m:e><m:sub>${sub}</m:sub></m:sSub>`
  return `<m:sSup><m:e>${base}</m:e><m:sup>${sup}</m:sup></m:sSup>`
}

function accentOmml(chr: string, base: string): string {
  if (chr === '̅') return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>${base}</m:e></m:bar>`
  if (chr === '̲') return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr><m:e>${base}</m:e></m:bar>`
  return `<m:acc><m:accPr><m:chr m:val="${escapeXmlAttr(chr)}"/></m:accPr><m:e>${base}</m:e></m:acc>`
}

/** operand of a script / limit: a parenthesised group, a number, or one character */
function parseScriptOperand(c: Cursor): string {
  const ch = peek(c)
  if (ch === '(') {
    c.pos++
    const inner = parseSeq(c, new Set([')']))
    expect(c, ')')
    return inner
  }
  if (/\d/.test(ch)) {
    let digits = ''
    while (/[\d.]/.test(peek(c))) digits += c.src[c.pos++]
    return run(digits)
  }
  if (ch === '' || SPECIAL.has(ch) || isCombining(ch)) {
    if (NARY_CHARS.has(ch) || ch === '√') return parsePrimary(c, new Set()).omml
    throw new UnicodeMathError('An operand is required here')
  }
  c.pos++
  return run(ch)
}

function expect(c: Cursor, ch: string): void {
  if (peek(c) !== ch) throw new UnicodeMathError(`Missing "${ch}"`)
  c.pos++
}

const matrixOf = (omml: string[][]): string =>
  `<m:m>${omml.map((row) => `<m:mr>${row.map((cell) => `<m:e>${cell}</m:e>`).join('')}</m:mr>`).join('')}</m:m>`

function parseRows(c: Cursor, closing: string): string[][] {
  const rows: string[][] = [[]]
  const stops = new Set(['&', '@', closing])
  for (;;) {
    rows[rows.length - 1]!.push(parseSeq(c, stops))
    const ch = peek(c)
    if (ch === '&') c.pos++
    else if (ch === '@') {
      c.pos++
      rows.push([])
    } else break
  }
  expect(c, closing)
  return rows
}

function parsePrimary(c: Cursor, stops: ReadonlySet<string>, single = false): Operand {
  const ch = peek(c)
  if (ch === '') throw new UnicodeMathError('Unexpected end of input')
  if (ch === ' ' || ch === '\t' || ch === '\n') {
    c.pos++
    return { omml: '', grouped: false }
  }
  if (ch === '"') {
    c.pos++
    let text = ''
    while (c.pos < c.src.length && peek(c) !== '"') text += c.src[c.pos++]
    expect(c, '"')
    return { omml: norRun(text), grouped: false }
  }
  if (ch === '√' || ch === '∛' || ch === '∜') {
    c.pos++
    skipSpace(c)
    let degree = ch === '∛' ? run('3') : ch === '∜' ? run('4') : ''
    let body: string
    if (peek(c) === '(') {
      c.pos++
      const rows = parseRows(c, ')')
      if (rows.length > 1) throw new UnicodeMathError('A radical takes one argument')
      if (rows[0]!.length === 2 && ch === '√') {
        degree = rows[0]![0]!
        body = rows[0]![1]!
      } else if (rows[0]!.length === 1) body = rows[0]![0]!
      else throw new UnicodeMathError('A radical takes a degree and a base')
    } else body = parseScriptOperand(c)
    const pr = degree === '' ? '<m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>' : ''
    return {
      omml: `<m:rad>${pr}${degree === '' ? '' : `<m:deg>${degree}</m:deg>`}<m:e>${body}</m:e></m:rad>`,
      grouped: false,
    }
  }
  if (NARY_CHARS.has(ch)) {
    c.pos++
    let sub = ''
    let sup = ''
    for (let i = 0; i < 2; i++) {
      if (peek(c) === '_' && sub === '') {
        c.pos++
        sub = parseScriptOperand(c)
      } else if (peek(c) === '^' && sup === '') {
        c.pos++
        sup = parseScriptOperand(c)
      } else break
    }
    let body = ''
    if (peek(c) === NARY_OPERAND) {
      c.pos++
      body = parsePrimaryWithScripts(c, stops, true)
    }
    const limLoc = SIDE_LIMIT_NARY.has(ch) ? 'subSup' : 'undOvr'
    return {
      omml:
        `<m:nary><m:naryPr><m:chr m:val="${escapeXmlAttr(ch)}"/><m:limLoc m:val="${limLoc}"/>` +
        `${sub === '' ? '<m:subHide m:val="1"/>' : ''}${sup === '' ? '<m:supHide m:val="1"/>' : ''}</m:naryPr>` +
        `${sub === '' ? '' : `<m:sub>${sub}</m:sub>`}${sup === '' ? '' : `<m:sup>${sup}</m:sup>`}<m:e>${body}</m:e></m:nary>`,
      grouped: false,
    }
  }
  if (ch === MATRIX_OPEN || ch === ARRAY_OPEN || ch === BOX_OPEN) {
    c.pos++
    expect(c, '(')
    const rows = parseRows(c, ')')
    if (ch === MATRIX_OPEN) return { omml: matrixOf(rows), grouped: false }
    if (ch === BOX_OPEN) {
      if (rows.length !== 1 || rows[0]!.length !== 1)
        throw new UnicodeMathError('A box takes one argument')
      return {
        omml: `<m:borderBox><m:e>${rows[0]![0]}</m:e></m:borderBox>`,
        grouped: false,
      }
    }
    const lines = rows.map((row) => {
      return row.map((cell, index) => (index === 0 ? cell : alignRun() + cell)).join('')
    })
    return {
      omml: `<m:eqArr>${lines.map((line) => `<m:e>${line}</m:e>`).join('')}</m:eqArr>`,
      grouped: false,
    }
  }
  if (ch === INVISIBLE_OPEN || ch in BRACKET_PAIRS || SAME_SIDE.has(ch) || ch === '(') {
    return parseDelimited(c, ch)
  }
  if (CLOSERS.has(ch) || ch === INVISIBLE_CLOSE || ch === '&' || ch === '@')
    throw new UnicodeMathError(`Unexpected "${ch}"`)
  if (isCombining(ch)) throw new UnicodeMathError('An accent needs an operand')
  // function name followed by function-apply
  const word = /^[A-Za-z]+/.exec(c.src.slice(c.pos).join(''))?.[0]
  if (word && FUNCTION_NAMES.has(word) && c.src[c.pos + word.length] === FUNCTION_APPLY) {
    c.pos += word.length + 1
    const argument = parsePrimaryWithScripts(c, stops, true)
    return {
      omml: `<m:func><m:fName>${run(word, true)}</m:fName><m:e>${argument}</m:e></m:func>`,
      grouped: false,
    }
  }
  // a recognised function name stays upright even without the function-apply mark
  if (!single && word && FUNCTION_NAMES.has(word)) {
    c.pos += word.length
    return { omml: run(word, true), grouped: false }
  }
  // an n-ary operand is one token: a number or a single character
  if (single) {
    let token = ''
    if (/\d/.test(ch)) while (/[\d.]/.test(peek(c))) token += c.src[c.pos++]
    else {
      token = ch
      c.pos++
    }
    return { omml: run(token), grouped: false }
  }
  // ordinary characters
  let text = ''
  while (c.pos < c.src.length) {
    const next = peek(c)
    if (SPECIAL.has(next) || isCombining(next) || next === ' ' || next === '\n' || next === '\t')
      break
    // a script / combining mark binds to the last character only
    if (
      text !== '' &&
      (c.src[c.pos + 1] === '^' || c.src[c.pos + 1] === '_' || isCombining(c.src[c.pos + 1] ?? ''))
    ) {
      text += next
      c.pos++
      break
    }
    text += next
    c.pos++
  }
  if (text === '') throw new UnicodeMathError(`Unexpected "${ch}"`)
  // keep only the final character for a following script: "ab^2" = a·b²
  const chars = [...text]
  if (chars.length > 1 && (peek(c) === '^' || peek(c) === '_' || isCombining(peek(c)))) {
    c.pos -= chars[chars.length - 1]!.length
    return { omml: run(chars.slice(0, -1).join('')), grouped: false }
  }
  return { omml: run(text), grouped: false }
}

function skipSpace(c: Cursor): void {
  while (peek(c) === ' ') c.pos++
}

function parseDelimited(c: Cursor, open: string): Operand {
  c.pos++
  const same = SAME_SIDE.has(open)
  const close = same ? open : open === INVISIBLE_OPEN ? '' : (BRACKET_PAIRS[open] ?? ')')
  const closers = new Set<string>([...CLOSERS, INVISIBLE_CLOSE, ...(same ? [open] : [])])
  const slots: string[] = []
  const stops = new Set<string>([...closers, ...(same ? [] : ['|'])])
  for (;;) {
    slots.push(parseSeq(c, stops))
    if (peek(c) === '|' && !same) {
      c.pos++
      continue
    }
    break
  }
  const end = peek(c)
  if (end === '') throw new UnicodeMathError('Missing closing bracket')
  if (!closers.has(end) || (close !== '' && end !== close && end !== INVISIBLE_CLOSE))
    throw new UnicodeMathError(`Unexpected "${end}"`)
  c.pos++
  const begChr = open === INVISIBLE_OPEN ? '' : open
  const endChr = end === INVISIBLE_CLOSE ? '' : end
  // bare parentheses with a single slot are grouping until proven otherwise
  if (begChr === '(' && endChr === ')' && slots.length === 1) {
    return { omml: slots[0]!, grouped: true }
  }
  return {
    omml:
      `<m:d><m:dPr><m:begChr m:val="${escapeXmlAttr(begChr)}"/>` +
      (slots.length > 1 ? '<m:sepChr m:val="|"/>' : '') +
      `<m:endChr m:val="${escapeXmlAttr(endChr)}"/></m:dPr>${slots.map((s) => `<m:e>${s}</m:e>`).join('')}</m:d>`,
    grouped: false,
  }
}

/** OMML (children of <m:oMath>) for UnicodeMath text; throws UnicodeMathError outside the supported subset. */
export function unicodeMathToOmml(text: string): string {
  const cursor: Cursor = { src: [...text], pos: 0 }
  const out = parseSeq(cursor, new Set())
  if (cursor.pos < cursor.src.length)
    throw new UnicodeMathError(`Unexpected "${cursor.src[cursor.pos]}"`)
  return out
}
