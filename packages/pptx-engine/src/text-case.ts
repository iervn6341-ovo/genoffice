/**
 * PowerPoint's Home → Font → Change Case, as a pure function over the text segments of one
 * paragraph (runs in the model, text nodes in the editor). Segments keep their boundaries, so
 * formatting stays on the characters it was on; "start of a sentence" / "start of a word" is
 * decided on the joined paragraph text, not per segment.
 */
export type TextCaseMode = 'sentence' | 'lower' | 'upper' | 'title' | 'toggle'

export const TEXT_CASE_MODES: readonly TextCaseMode[] = [
  'sentence',
  'lower',
  'upper',
  'title',
  'toggle',
]

/** Characters after which the next letter starts a sentence (Latin and CJK full stops) */
const SENTENCE_END = /[.!?。！？]/
const isSpace = (ch: string) => /\s/.test(ch)
const isLetter = (ch: string) => ch.toLowerCase() !== ch.toUpperCase()

/**
 * Change the case of `parts` (consecutive segments of one paragraph) and return the new
 * segments. `atSentenceStart` says whether the first character opens a sentence (true for a
 * paragraph start; pass false when the parts continue text that came before them).
 * A character whose case mapping changes its length (e.g. "ß" → "SS") is left as it is, so
 * segment lengths — and the editor's selection offsets — never move.
 */
export function changeTextCase(
  parts: readonly string[],
  mode: TextCaseMode,
  atSentenceStart = true,
): string[] {
  let sentenceStart = atSentenceStart
  let wordStart = true
  return parts.map((part) => {
    let out = ''
    for (const ch of part) {
      let next = ch
      if (isLetter(ch)) {
        switch (mode) {
          case 'lower':
            next = ch.toLowerCase()
            break
          case 'upper':
            next = ch.toUpperCase()
            break
          case 'toggle':
            next = ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase()
            break
          case 'sentence':
            next = sentenceStart ? ch.toUpperCase() : ch.toLowerCase()
            break
          case 'title':
            next = wordStart ? ch.toUpperCase() : ch.toLowerCase()
            break
        }
        if (next.length !== ch.length) next = ch
        sentenceStart = false
        wordStart = false
      } else if (SENTENCE_END.test(ch)) {
        sentenceStart = true
        wordStart = true
      } else if (isSpace(ch)) {
        wordStart = true
      } else if (/[\p{L}\p{N}]/u.test(ch)) {
        // digits and caseless letters (CJK) continue a word/sentence
        sentenceStart = false
        wordStart = false
      } else {
        // punctuation such as quotes or hyphens starts a new word (PowerPoint capitalises
        // "well-known" as "Well-Known")
        wordStart = true
      }
      out += next
    }
    return out
  })
}
