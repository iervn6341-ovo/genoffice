/// Rewrite an imported Office equation in place. DrawingML shapes can carry
/// transforms, formatting, names, and adjacent text that a new shape cannot
/// faithfully reconstruct; only the math payload and its matched text fallback
/// are edited here.

import { createHash } from 'node:crypto'
import { latexToOmml, ommlToLatex } from '@genoffice/office-math'
import { equationSourceDescription, type EquationAdd } from './xlsx-drawing-add'

export class EquationEditError extends Error {}

interface XmlMatch {
  readonly xml: string
  readonly index: number
}

function firstMatch(xml: string, pattern: RegExp): XmlMatch | null {
  const match = pattern.exec(xml)
  return match ? { xml: match[0], index: match.index } : null
}

function replaceMatch(xml: string, match: XmlMatch, replacement: string): string {
  return xml.slice(0, match.index) + replacement + xml.slice(match.index + match.xml.length)
}

function nativeMath(xml: string): { wrapper: XmlMatch; omath: XmlMatch; omml: string } | null {
  const wrappers = [...xml.matchAll(/<a14:m\b[^>]*>[\s\S]*?<\/a14:m>/g)]
  const mathOpenings = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?oMath\b/g)]
  if (wrappers.length !== 1 || mathOpenings.length !== 1) return null
  const match = wrappers[0]!
  const wrapper = { xml: match[0], index: match.index }
  const math = /<([A-Za-z_][\w.-]*:)?oMath\b[^>]*>([\s\S]*?)<\/\1oMath>/.exec(wrapper.xml)
  if (!math) return null
  return { wrapper, omath: { xml: math[0], index: math.index }, omml: math[2] ?? '' }
}

function rewriteMath(xml: string, omml: string): string {
  const native = nativeMath(xml)
  if (!native) throw new EquationEditError('The equation has no supported native math payload.')
  const openEnd = native.omath.xml.indexOf('>') + 1
  const closeStart = native.omath.xml.lastIndexOf('</')
  const revisedMath = native.omath.xml.slice(0, openEnd) + omml + native.omath.xml.slice(closeStart)
  const revisedWrapper = replaceMatch(native.wrapper.xml, native.omath, revisedMath)
  return replaceMatch(xml, native.wrapper, revisedWrapper)
}

function storedSource(xml: string): string | null {
  const metadata =
    /<(?:[\w.-]+:)?cNvPr\b[^>]*\bdescr="GenOfficeEquationV2:([A-Za-z0-9+/=]+):([a-f0-9]{64})"/.exec(
      xml,
    )
  const encoded = metadata?.[1]
  return encoded && encoded.length <= 24_000
    ? Buffer.from(encoded, 'base64').toString('utf8')
    : null
}

function ownSource(xml: string, omml: string): string | null {
  const metadata =
    /<(?:[\w.-]+:)?cNvPr\b[^>]*\bdescr="GenOfficeEquationV2:[A-Za-z0-9+/=]+:([a-f0-9]{64})"/.exec(
      xml,
    )
  if (!metadata || metadata[1] !== createHash('sha256').update(omml).digest('hex')) return null
  const source = storedSource(xml)
  if (!source) return null
  try {
    return latexToOmml(source) === omml ? source : null
  } catch {
    return null
  }
}

function currentSource(xml: string, omml: string): string | null {
  return ownSource(xml, omml) ?? ommlToLatex(`<m:oMath>${omml}</m:oMath>`)
}

function rewriteOwnedDescription(xml: string, description: string): string {
  const tag = firstMatch(xml, /<(?:[\w.-]+:)?cNvPr\b[^>]*\/?>/)
  if (!tag) return xml
  const existing = /\bdescr="([^"]*)"/.exec(tag.xml)
  if (existing && !existing[1]?.startsWith('GenOfficeEquationV2:')) return xml
  const revised = existing
    ? tag.xml.replace(/\bdescr="[^"]*"/, `descr="${description}"`)
    : tag.xml.replace(/(\/?)>$/, ` descr="${description}"$1>`)
  return replaceMatch(xml, tag, revised)
}

function decodeText(value: string): string {
  return value.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (entity) => {
    switch (entity) {
      case '&amp;':
        return '&'
      case '&lt;':
        return '<'
      case '&gt;':
        return '>'
      case '&quot;':
        return '"'
      case '&apos;':
        return "'"
      default: {
        const code = entity.startsWith('&#x')
          ? Number.parseInt(entity.slice(3, -1), 16)
          : Number.parseInt(entity.slice(2, -1), 10)
        return Number.isFinite(code) && code >= 0 && code <= 0x10ffff
          ? String.fromCodePoint(code)
          : entity
      }
    }
  })
}

function encodeText(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

interface FallbackNode {
  readonly match: XmlMatch
  readonly text: string
  readonly open: string
  readonly close: string
  readonly start: number
  readonly end: number
}

function fallbackNodes(xml: string): FallbackNode[] {
  const nodes: FallbackNode[] = []
  let offset = 0
  for (const match of xml.matchAll(/(<a:t\b[^>]*>)([\s\S]*?)(<\/a:t>)/g)) {
    const text = decodeText(match[2] ?? '')
    nodes.push({
      match: { xml: match[0], index: match.index },
      text,
      open: match[1] ?? '',
      close: match[3] ?? '',
      start: offset,
      end: offset + text.length,
    })
    offset += text.length
  }
  return nodes
}

function fallbackRange(xml: string, oldLatex: string): { start: number; end: number } | null {
  if (/<(?:[\w.-]+:)?(?:pic|graphicFrame)\b|<a:blip\b/.test(xml)) return null
  const nodes = fallbackNodes(xml)
  if (nodes.length === 0 || !oldLatex) return null
  const text = nodes.map((node) => node.text).join('')
  const index = text.indexOf(oldLatex)
  const end = index + oldLatex.length
  if (
    index >= 0 &&
    text.indexOf(oldLatex, index + 1) < 0 &&
    nodes.some((node) => node.start === index) &&
    nodes.some((node) => node.end === end)
  ) {
    return { start: index, end }
  }
  return null
}

function rewriteFallback(xml: string, oldLatex: string, newLatex: string): string {
  const range = fallbackRange(xml, oldLatex)
  if (!range) {
    throw new EquationEditError('The equation fallback cannot be updated without losing content.')
  }
  const nodes = fallbackNodes(xml)
  let inserted = false
  let result = xml
  for (const node of [...nodes].reverse()) {
    const overlapStart = Math.max(range.start, node.start)
    const overlapEnd = Math.min(range.end, node.end)
    if (overlapEnd <= overlapStart) continue
    const before = node.text.slice(0, overlapStart - node.start)
    const after = node.text.slice(overlapEnd - node.start)
    // On reverse traversal, insert once at the earliest overlapping run.
    const isFirst = node.start <= range.start && range.start < node.end
    const text = before + (isFirst ? newLatex : '') + after
    if (isFirst) inserted = true
    result = replaceMatch(result, node.match, node.open + encodeText(text) + node.close)
  }
  if (!inserted) throw new EquationEditError('The equation fallback text is ambiguous.')
  return result
}

function alternateParts(anchorXml: string): {
  alternate: XmlMatch
  choice: XmlMatch
  fallback: XmlMatch
} | null {
  const alternate = firstMatch(
    anchorXml,
    /<([A-Za-z_][\w.-]*:)?AlternateContent\b[^>]*>[\s\S]*?<\/\1AlternateContent>/,
  )
  if (!alternate) return null
  const choices = [
    ...alternate.xml.matchAll(/<([A-Za-z_][\w.-]*:)?Choice\b[^>]*>[\s\S]*?<\/\1Choice>/g),
  ]
  const choiceMatch = choices.find((match) => nativeMath(match[0]) !== null)
  const fallback = firstMatch(
    alternate.xml,
    /<([A-Za-z_][\w.-]*:)?Fallback\b[^>]*>[\s\S]*?<\/\1Fallback>/,
  )
  if (!choiceMatch || !fallback) return null
  return {
    alternate,
    choice: { xml: choiceMatch[0], index: choiceMatch.index },
    fallback,
  }
}

export function canEditEquationAnchor(anchorXml: string, latex: string | null): boolean {
  if (!latex || !nativeMath(anchorXml)) return false
  const parts = alternateParts(anchorXml)
  if (!parts) return !/<(?:[\w.-]+:)?AlternateContent\b/.test(anchorXml)
  return [latex, storedSource(parts.choice.xml)].some(
    (candidate) => candidate !== null && fallbackRange(parts.fallback.xml, candidate) !== null,
  )
}

export function rewriteEquationAnchor(anchorXml: string, equation: EquationAdd): string {
  const description = equationSourceDescription(equation)
  const selectedMath = nativeMath(anchorXml)
  if (!selectedMath) {
    throw new EquationEditError('The selected drawing does not have one supported native equation.')
  }
  const parts = alternateParts(anchorXml)
  if (parts) {
    const native = nativeMath(parts.choice.xml)
    if (!native) throw new EquationEditError('The selected drawing has no native equation.')
    const authoritativeSource = currentSource(parts.choice.xml, native.omml)
    if (!authoritativeSource) {
      throw new EquationEditError('The original Office equation cannot be reversed safely.')
    }
    const previous = [authoritativeSource, storedSource(parts.choice.xml)].find(
      (candidate) => candidate && fallbackRange(parts.fallback.xml, candidate) !== null,
    )
    if (!previous) {
      throw new EquationEditError('The equation fallback cannot be updated without losing content.')
    }
    const fallback = rewriteOwnedDescription(
      rewriteFallback(parts.fallback.xml, previous, equation.latex),
      description,
    )
    const choice = rewriteOwnedDescription(
      rewriteMath(parts.choice.xml, equation.omml),
      description,
    )
    // Both replacements use source offsets in the original AlternateContent.
    let revised = parts.alternate.xml
    for (const part of [
      { match: parts.choice, replacement: choice },
      { match: parts.fallback, replacement: fallback },
    ].sort((left, right) => right.match.index - left.match.index)) {
      revised = replaceMatch(revised, part.match, part.replacement)
    }
    return replaceMatch(anchorXml, parts.alternate, revised)
  }
  if (/<(?:[\w.-]+:)?AlternateContent\b/.test(anchorXml)) {
    throw new EquationEditError('The selected equation has no editable text fallback.')
  }
  if (!currentSource(anchorXml, selectedMath.omml)) {
    throw new EquationEditError('The original Office equation cannot be reversed safely.')
  }
  return rewriteOwnedDescription(rewriteMath(anchorXml, equation.omml), description)
}
