/**
 * Unicode-aware text folding for Find, with an offset map back to the source.
 *
 * Word's Find without "Match case" treats `straße` = `STRASSE`, final `ς` = `σ`,
 * `İ` = `i`, and a precomposed `é` = `e` + U+0301. Each of those changes the
 * string length, so the folded haystack carries, per folded char, the source
 * span of the segment it came from. A segment is one base character plus its
 * combining marks, so a match can never start or end inside `é` — searching
 * `e` does not hit the `e` of a decomposed `é`.
 */

export interface FoldedText {
  /** the searchable string */
  readonly folded: string
  /** per folded UTF-16 unit: source index where its segment starts */
  readonly srcStart: readonly number[]
  /** per folded UTF-16 unit: source index where its segment ends (exclusive) */
  readonly srcEnd: readonly number[]
  /** per folded UTF-16 unit: first unit of its segment's folded form */
  readonly segStart: readonly boolean[]
  /** per folded UTF-16 unit: last unit of its segment's folded form */
  readonly segEnd: readonly boolean[]
}

// base char (any non-mark code point) followed by combining marks
const SEGMENT_RE = /[^\p{M}][\p{M}]*|[\p{M}]+/gsu

/** case-insensitive fold of one NFC segment (full folding, not length-preserving) */
function foldSegment(segment: string): string {
  return segment
    .toLowerCase()
    .replace(/i̇/g, 'i') // İ lowercases to i + combining dot above
    .replace(/ß/g, 'ss')
    .replace(/ẞ/g, 'ss')
    .replace(/ς/g, 'σ')
    .replace(/ſ/g, 's')
    .normalize('NFC')
}

export function foldForSearch(text: string, matchCase: boolean): FoldedText {
  let folded = ''
  const srcStart: number[] = []
  const srcEnd: number[] = []
  const segStart: boolean[] = []
  const segEnd: boolean[] = []
  for (const match of text.matchAll(SEGMENT_RE)) {
    const start = match.index
    const end = start + match[0].length
    const nfc = match[0].normalize('NFC')
    const out = matchCase ? nfc : foldSegment(nfc)
    for (let k = 0; k < out.length; k++) {
      srcStart.push(start)
      srcEnd.push(end)
      segStart.push(k === 0)
      segEnd.push(k === out.length - 1)
    }
    folded += out
  }
  return { folded, srcStart, srcEnd, segStart, segEnd }
}

/** source [from, to) spans of every match of `query` in `text` */
export function findFoldedSpans(
  text: string,
  query: string,
  matchCase: boolean,
  accept: (from: number, to: number) => boolean = () => true,
): Array<{ from: number; to: number }> {
  const needle = foldForSearch(query, matchCase).folded
  if (!needle) return []
  const hay = foldForSearch(text, matchCase)
  const spans: Array<{ from: number; to: number }> = []
  let i = 0
  while ((i = hay.folded.indexOf(needle, i)) !== -1) {
    const last = i + needle.length - 1
    const from = hay.srcStart[i]!
    const to = hay.srcEnd[last]!
    if (hay.segStart[i] && hay.segEnd[last] && accept(from, to)) {
      spans.push({ from, to })
      i += needle.length
    } else {
      i += 1
    }
  }
  return spans
}
