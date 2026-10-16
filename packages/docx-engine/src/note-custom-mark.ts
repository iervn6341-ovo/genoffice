import { decodeEntities } from './parse-xml-text'
import { symbolGlyph } from './symbol-fonts'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

export interface NoteCustomMark {
  text: string
  symbol?: { font: string; char: string }
}

// Word's Symbol-font encodings observed in the Mac custom-mark dialog.
const TEXT_SYMBOLS: Record<string, string> = {
  '♣': 'F0A7',
  '♦': 'F0A8',
  '♥': 'F0A9',
  '♠': 'F0AA',
}

export function wordSymbolForMark(text: string): { font: string; char: string } | undefined {
  return TEXT_SYMBOLS[text] ? { font: 'Symbol', char: TEXT_SYMBOLS[text] } : undefined
}

export function noteCustomMarkFromRun(runXml: string): NoteCustomMark | null {
  const symbol = /<w:sym\b([^>]*)\/>/.exec(runXml)
  if (symbol) {
    const font = /\bw:font="([^"]+)"/.exec(symbol[1])?.[1]
    const char = /\bw:char="([0-9A-Fa-f]+)"/.exec(symbol[1])?.[1]?.toUpperCase()
    if (!font || !char) return null
    const text = symbolGlyph(font, char) ?? '�'
    return { text, symbol: { font: decodeEntities(font), char } }
  }
  const text = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/.exec(runXml)?.[1]
  return text === undefined ? null : { text: decodeEntities(text) }
}

export function noteCustomMarkXml(mark: NoteCustomMark): string {
  const symbol = mark.symbol ?? wordSymbolForMark(mark.text)
  return symbol
    ? `<w:sym w:font="${escapeXmlAttr(symbol.font)}" w:char="${escapeXmlAttr(symbol.char)}"/>`
    : `<w:t xml:space="preserve">${escapeXmlText(mark.text)}</w:t>`
}

/** The first run of a custom note entry is its mark, not editable note text. */
export function firstNoteMarkRun(entryXml: string): string | null {
  return /<w:p\b[^>]*>[\s\S]*?(<w:r\b[^>]*>[\s\S]*?<\/w:r>)/.exec(entryXml)?.[1] ?? null
}
