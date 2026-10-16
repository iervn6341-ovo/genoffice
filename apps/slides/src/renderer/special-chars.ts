/**
 * "Special Characters" tab of the Symbol dialog (typographic characters that
 * PowerPoint/Word list separately from the Symbols grid). The names are Unicode
 * character names (technical identifiers, not UI copy).
 */
export interface SpecialChar {
  readonly char: string
  readonly name: string
}

export const SPECIAL_CHARS: readonly SpecialChar[] = [
  { char: '—', name: 'EM DASH' },
  { char: '–', name: 'EN DASH' },
  { char: '‑', name: 'NON-BREAKING HYPHEN' },
  { char: '­', name: 'SOFT HYPHEN' },
  { char: ' ', name: 'EM SPACE' },
  { char: ' ', name: 'EN SPACE' },
  { char: ' ', name: 'FOUR-PER-EM SPACE' },
  { char: ' ', name: 'NO-BREAK SPACE' },
  { char: '©', name: 'COPYRIGHT SIGN' },
  { char: '®', name: 'REGISTERED SIGN' },
  { char: '™', name: 'TRADE MARK SIGN' },
  { char: '§', name: 'SECTION SIGN' },
  { char: '¶', name: 'PILCROW SIGN' },
  { char: '…', name: 'HORIZONTAL ELLIPSIS' },
  { char: '‘', name: 'LEFT SINGLE QUOTATION MARK' },
  { char: '’', name: 'RIGHT SINGLE QUOTATION MARK' },
  { char: '“', name: 'LEFT DOUBLE QUOTATION MARK' },
  { char: '”', name: 'RIGHT DOUBLE QUOTATION MARK' },
  { char: '•', name: 'BULLET' },
  { char: '†', name: 'DAGGER' },
  { char: '‡', name: 'DOUBLE DAGGER' },
  { char: '‰', name: 'PER MILLE SIGN' },
]

/** Printable stand-in for invisible/whitespace characters in the grid. */
export function specialCharGlyph(char: string): string {
  switch (char) {
    case ' ':
    case ' ':
    case ' ':
    case ' ':
      return '␣'
    case '­':
      return '-'
    default:
      return char
  }
}
