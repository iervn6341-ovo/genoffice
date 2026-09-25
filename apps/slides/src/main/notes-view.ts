/**
 * Notes pane view model: the notes body's model paragraphs → what the pane renders.
 * Only display-relevant, resolved formatting crosses IPC; the pane returns EditParagraphs
 * traced by index (srcPara / srcRun), and the setNotes op merges them back into the model.
 */
import { NOTES_DEFAULT_FONT_PT, type Paragraph } from '@genoffice/pptx-engine'
import type { NotesParagraphView } from '../shared/ipc'

const ROMAN = ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x']

/** Visible label of an auto-number bullet (the common arabic / alpha / roman families) */
function autoNumLabel(numType: string | undefined, n: number): string {
  const t = numType ?? 'arabicPeriod'
  const body = t.startsWith('alphaLc')
    ? String.fromCharCode(96 + (((n - 1) % 26) + 1))
    : t.startsWith('alphaUc')
      ? String.fromCharCode(64 + (((n - 1) % 26) + 1))
      : t.startsWith('romanLc')
        ? (ROMAN[n - 1] ?? String(n))
        : t.startsWith('romanUc')
          ? (ROMAN[n - 1] ?? String(n)).toUpperCase()
          : String(n)
  if (t.endsWith('ParenBoth')) return `(${body})`
  if (t.endsWith('ParenR')) return `${body})`
  if (t.endsWith('Plain')) return body
  return `${body}.`
}

export function notesParagraphViews(paras: Paragraph[]): NotesParagraphView[] {
  // numbering restarts whenever a paragraph at that level is not numbered; deeper levels
  // restart under every new parent
  const counters: Array<number | undefined> = []
  return paras.map((p) => {
    const level = p.level ?? 0
    let bullet: string | undefined
    if (p.bullet?.type === 'number') {
      const n = (counters[level] ?? (p.bullet.startAt ?? 1) - 1) + 1
      counters[level] = n
      bullet = autoNumLabel(p.bullet.numType, n)
    } else {
      counters[level] = undefined
      if (p.bullet?.type === 'char') bullet = p.bullet.char ?? '•'
    }
    counters.length = level + 1
    return {
      level,
      ...(p.align ? { align: p.align } : {}),
      ...(bullet ? { bullet } : {}),
      runs: p.runs.map((r) => ({
        text: r.text,
        bold: !!r.bold,
        italic: !!r.italic,
        underline: !!r.underline,
        strike: !!r.strike,
        baseline: r.baseline ?? 0,
        fontSizePt: r.fontSize ?? NOTES_DEFAULT_FONT_PT,
        ...(r.fontFamily ? { fontFamily: r.fontFamily } : {}),
        fontExplicit: !r.fontImplicit,
        ...(r.color && !r.colorInherited ? { color: r.color } : {}),
        ...(r.highlight ? { highlight: r.highlight } : {}),
      })),
    }
  })
}
