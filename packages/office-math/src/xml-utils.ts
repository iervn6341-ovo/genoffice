import { XMLParser } from 'fast-xml-parser'

/** preserveOrder node shape from fast-xml-parser */
export type XNode = Record<string, unknown>

const parserOptions = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
} as const

export const xmlParser = new XMLParser(parserOptions)

/**
 * Table XML only: deep nesting is legitimate there (POI stress files nest 5000 table
 * levels, far past the default 100-tag cap). fxp parses iteratively, but callers must
 * cap their own recursion when walking the result.
 */
export const deepXmlParser = new XMLParser({ ...parserOptions, maxNestedTags: 100_000 })

export function nameOf(node: XNode): string | undefined {
  return Object.keys(node).find((k) => k !== ':@' && k !== '#text')
}

export function childrenOf(node: XNode): XNode[] {
  const name = nameOf(node)
  if (!name) return []
  const value = node[name]
  return Array.isArray(value) ? (value as XNode[]) : []
}

export function attrsOf(node: XNode): Record<string, string> {
  return (node[':@'] as Record<string, string>) ?? {}
}

export function textOf(node: XNode): string {
  let out = ''
  for (const child of childrenOf(node)) {
    if ('#text' in child) out += String(child['#text'])
    else out += textOf(child)
  }
  return out
}

export function findChild(node: XNode, name: string): XNode | undefined {
  return childrenOf(node).find((c) => nameOf(c) === name)
}

export function findChildren(node: XNode, name: string): XNode[] {
  return childrenOf(node).filter((c) => nameOf(c) === name)
}

// Control characters outside \t \n \r are illegal in XML 1.0 even when escaped
// eslint-disable-next-line no-control-regex
const ILLEGAL_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g

export function escapeXmlText(text: string): string {
  return text
    .replace(ILLEGAL_XML_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

export function escapeXmlAttr(text: string): string {
  return escapeXmlText(text).replace(/"/g, '&quot;')
}
