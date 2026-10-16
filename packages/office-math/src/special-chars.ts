/**
 * Word's "Special Characters" list and its default AutoCorrect replacements, shared so Docs, Sheets
 * and Slides can offer the same typographic characters. Names are the canonical Unicode names.
 */

export interface SpecialCharacter {
  id: string
  char: string
  /** Unicode name (also what the symbol search matches) */
  name: string
  /** the OOXML construct Word writes for it when it is not a plain character */
  note?: 'noBreakHyphen' | 'softHyphen'
}

export const SPECIAL_CHARACTERS: readonly SpecialCharacter[] = [
  { id: 'emDash', char: '—', name: 'EM DASH' },
  { id: 'enDash', char: '–', name: 'EN DASH' },
  { id: 'nonBreakingHyphen', char: '‑', name: 'NON-BREAKING HYPHEN', note: 'noBreakHyphen' },
  { id: 'optionalHyphen', char: '­', name: 'SOFT HYPHEN', note: 'softHyphen' },
  { id: 'emSpace', char: ' ', name: 'EM SPACE' },
  { id: 'enSpace', char: ' ', name: 'EN SPACE' },
  { id: 'quarterEmSpace', char: ' ', name: 'FOUR-PER-EM SPACE' },
  { id: 'nonBreakingSpace', char: ' ', name: 'NO-BREAK SPACE' },
  { id: 'copyright', char: '©', name: 'COPYRIGHT SIGN' },
  { id: 'registered', char: '®', name: 'REGISTERED SIGN' },
  { id: 'trademark', char: '™', name: 'TRADE MARK SIGN' },
  { id: 'section', char: '§', name: 'SECTION SIGN' },
  { id: 'paragraph', char: '¶', name: 'PILCROW SIGN' },
  { id: 'ellipsis', char: '…', name: 'HORIZONTAL ELLIPSIS' },
  { id: 'singleOpenQuote', char: '‘', name: 'LEFT SINGLE QUOTATION MARK' },
  { id: 'singleCloseQuote', char: '’', name: 'RIGHT SINGLE QUOTATION MARK' },
  { id: 'doubleOpenQuote', char: '“', name: 'LEFT DOUBLE QUOTATION MARK' },
  { id: 'doubleCloseQuote', char: '”', name: 'RIGHT DOUBLE QUOTATION MARK' },
]

/** Word's default "replace as you type" entries (text typed, then a space or punctuation). */
export const AUTOCORRECT_DEFAULTS: Readonly<Record<string, string>> = {
  '(c)': '©',
  '(r)': '®',
  '(tm)': '™',
  '(e)': '€',
  '...': '…',
  '-->': '→',
  '<--': '←',
  '==>': '⇒',
  '<==': '⇐',
  ':)': '☺',
  ':(': '☹',
  '1/2': '½',
  '1/4': '¼',
  '3/4': '¾',
}

/**
 * Replacement for the text typed before a trigger character, or null. User entries win over the
 * defaults; the match must start the line or follow whitespace.
 */
export function autoCorrectReplacement(
  before: string,
  user: Readonly<Record<string, string>> = {},
): { length: number; with: string } | null {
  const table = { ...AUTOCORRECT_DEFAULTS, ...user }
  let best: { length: number; with: string } | null = null
  for (const [key, value] of Object.entries(table)) {
    if (!key || !before.endsWith(key)) continue
    const start = before.length - key.length
    if (start > 0 && !/\s/.test(before[start - 1] ?? '')) continue
    if (!best || key.length > best.length) best = { length: key.length, with: value }
  }
  return best
}
