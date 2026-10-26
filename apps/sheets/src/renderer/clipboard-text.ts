/**
 * Plain-text clipboard import. Excel splits pasted text at every tab and every
 * line break; a field that starts with a double quote runs to the closing quote
 * so it can carry tabs, line breaks and doubled quotes (`""`). Univer's own
 * plain-text path either put the whole text into one cell or turned it into
 * an unescaped `<table>` first, which lost quoted line breaks, trimmed leading
 * and trailing empty cells and broke on `<` / `&`.
 */
import { CellValueType, getNumfmtParseValueFilter } from '@univerjs/core'

/// A pasted text becomes a grid when it has a tab or a line break; a plain run of
/// text is one cell and stays on the ordinary single-cell path.
export function isGridText(text: string): boolean {
  return /[\t\r\n]/.test(text)
}

/// Splits text into rows and fields the way Excel reads pasted text. One trailing line
/// break ends the last row instead of adding an empty one.
export function parseDelimitedText(text: string, delimiter = '\t'): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let i = 0
  const n = text.length
  let atFieldStart = true
  const endField = (): void => {
    row.push(field)
    field = ''
    atFieldStart = true
  }
  const endRow = (): void => {
    endField()
    rows.push(row)
    row = []
  }
  while (i < n) {
    const ch = text[i]!
    if (atFieldStart && ch === '"') {
      // quoted field: runs to the next lone quote that ends the field
      let j = i + 1
      let value = ''
      let closed = false
      while (j < n) {
        const c = text[j]!
        if (c === '"') {
          if (text[j + 1] === '"') {
            value += '"'
            j += 2
            continue
          }
          closed = true
          j += 1
          break
        }
        value += c
        j += 1
      }
      const after = text[j]
      if (closed && (j >= n || after === delimiter || after === '\n' || after === '\r')) {
        field = value.replace(/\r\n|\r/g, '\n')
        i = j
        atFieldStart = false
        continue
      }
      // not a well-formed quoted field: the quote is plain text
      field += ch
      i += 1
      atFieldStart = false
      continue
    }
    if (ch === delimiter) {
      endField()
      i += 1
      continue
    }
    if (ch === '\r' || ch === '\n') {
      endRow()
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1
      continue
    }
    field += ch
    atFieldStart = false
    i += 1
  }
  // a final line break ends the last row; it does not start another one
  if (n === 0 || !(text.endsWith('\n') || text.endsWith('\r'))) endRow()
  return rows
}

/// One pasted field typed the way a typed-in cell would be: numbers, dates (with their
/// format), percentages, currency and TRUE/FALSE become values; a leading `=` is a formula.
export interface TypedField {
  readonly v: string | number | boolean | null
  readonly t: number | null
  readonly f: string | null
  readonly pattern: string | null
}

export function typeField(text: string): TypedField {
  if (text === '') return { v: null, t: null, f: null, pattern: null }
  if (text.length > 1 && text[0] === '=' && !/^=\s*$/.test(text)) {
    return { v: null, t: null, f: text, pattern: null }
  }
  const trimmed = text.trim()
  if (/^(true|false)$/i.test(trimmed)) {
    return {
      v: /^true$/i.test(trimmed) ? 1 : 0,
      t: CellValueType.BOOLEAN,
      f: null,
      pattern: null,
    }
  }
  // text that must stay text: a leading apostrophe (Excel's quote prefix) or digits Excel would
  // not keep precisely
  if (text.startsWith("'") && text.length > 1) {
    return { v: text.slice(1), t: CellValueType.STRING, f: null, pattern: null }
  }
  const parsed = getNumfmtParseValueFilter(trimmed)
  if (parsed && typeof parsed.v === 'number' && Number.isFinite(parsed.v)) {
    // a long digit string (account numbers, ids) would lose digits as a number
    if (!parsed.z && /^\d{16,}$/.test(trimmed)) {
      return { v: text, t: CellValueType.STRING, f: null, pattern: null }
    }
    // zero-padded digits ("00123") are codes: Excel pastes them as the number 123, but a
    // sheet user pasting ids wants them kept; Excel itself shows 123, so follow Excel
    return { v: parsed.v, t: CellValueType.NUMBER, f: null, pattern: parsed.z ?? null }
  }
  return { v: text, t: CellValueType.STRING, f: null, pattern: null }
}
