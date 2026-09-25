/**
 * Plain-text view of a header/footer part, shared by the on-canvas editing
 * surface and the AI set_header_footer tool: paragraphs edit as lines, the
 * invisible PAGE / NUMPAGES field sentinels as visible {PAGE} / {NUMPAGES}
 * tokens. Layout-table rows (cells) and images stay out of the text flow and
 * keep their original content.
 */
import {
  PAGE_MARK,
  TOTAL_PAGES_MARK,
  type HeaderFooter,
  type HfParagraph,
  type Run,
} from '@genoffice/docx-engine'

export const PAGE_TOKEN = '{PAGE}'
export const TOTAL_TOKEN = '{NUMPAGES}'

/** effective paragraphs: rich paras when present, else the legacy single line */
export function hfParasOf(value: HeaderFooter): HfParagraph[] {
  if (value.paras?.length) return value.paras
  const runs: Run[] = value.text ? [{ text: value.text }] : []
  if (value.pageNumber && !value.text.includes('#') && !value.text.includes(PAGE_MARK)) {
    runs.push({ text: runs.length > 0 ? ` ${PAGE_MARK}` : PAGE_MARK })
  }
  return [{ align: 'center', runs }]
}

/** editable text of the part: one line per text paragraph, field sentinels as tokens */
export function hfEditText(value: HeaderFooter): string {
  return hfParasOf(value)
    .filter((p) => !p.cells)
    .map((p) => p.runs.map((r) => r.text).join(''))
    .join('\n')
    .replaceAll(PAGE_MARK, PAGE_TOKEN)
    .replaceAll(TOTAL_PAGES_MARK, TOTAL_TOKEN)
}

/**
 * Replace only what changed in a line: the common start and end keep their runs (a bold
 * red word next to a one-character edit stays bold red), the new middle takes the
 * formatting of the character before it — the first character's when typed at the start —
 * as Word does. Rebuilding the line as one run flattened mixed formatting on any edit.
 */
function spliceLineRuns(runs: Run[], line: string): Run[] {
  const old = runs.map((r) => r.text).join('')
  if (old === line) return runs
  let from = 0
  while (from < old.length && from < line.length && old[from] === line[from]) from++
  let oldTo = old.length
  let newTo = line.length
  while (oldTo > from && newTo > from && old[oldTo - 1] === line[newTo - 1]) {
    oldTo--
    newTo--
  }
  const slice = (start: number, end: number): Run[] => {
    const out: Run[] = []
    let at = 0
    for (const run of runs) {
      const next = at + run.text.length
      const lo = Math.max(start, at)
      const hi = Math.min(end, next)
      if (hi > lo) out.push({ ...run, text: run.text.slice(lo - at, hi - at) })
      at = next
    }
    return out
  }
  const inserted = line.slice(from, newTo)
  const styleRun = (from > 0 ? slice(from - 1, from)[0] : slice(0, 1)[0]) ?? runs[0]
  const pieces = [
    ...slice(0, from),
    ...(inserted ? [{ ...styleRun, text: inserted }] : []),
    ...slice(oldTo, old.length),
  ]
  // rejoin pieces of one formatting (a run cut by the edit, or typed text continuing it)
  const out: Run[] = []
  for (const piece of pieces) {
    const last = out[out.length - 1]
    if (last && sameFormat(last, piece))
      out[out.length - 1] = { ...last, text: last.text + piece.text }
    else out.push(piece)
  }
  return out
}

const sameFormat = (a: Run, b: Run): boolean =>
  JSON.stringify({ ...a, text: '' }) === JSON.stringify({ ...b, text: '' })

/**
 * Map edited lines back onto the part: each line keeps its original
 * paragraph's format and the runs the edit did not touch (extra lines reuse
 * the last template's first-run styling), cells rows are spliced back at
 * their original positions.
 */
export function applyHfText(value: HeaderFooter | null, text: string): HeaderFooter {
  const base = value ?? { text: '' }
  const paras = hfParasOf(base)
  const lines = text
    .replace(/\n+$/, '')
    .replaceAll(PAGE_TOKEN, PAGE_MARK)
    .replaceAll(TOTAL_TOKEN, TOTAL_PAGES_MARK)
    .split('\n')
  const textParas = paras.filter((p) => !p.cells)
  const templates: HfParagraph[] =
    textParas.length > 0 ? textParas : [{ align: 'center', runs: [] }]
  const edited: HfParagraph[] = lines.map((line, i) => {
    const template = templates[Math.min(i, templates.length - 1)]
    if (line === '') return { ...template, runs: [] }
    // the line's own paragraph: splice the edit into its runs; an extra line starts
    // from the last paragraph's first-run styling
    if (i < textParas.length) return { ...template, runs: spliceLineRuns(template.runs, line) }
    return { ...template, runs: [{ ...(template.runs[0] ?? {}), text: line }] }
  })
  const nextParas: HfParagraph[] = []
  let ei = 0
  for (const p of paras) {
    if (p.cells) nextParas.push(p)
    else if (ei < edited.length) nextParas.push(edited[ei++])
  }
  nextParas.push(...edited.slice(ei))
  const nextText = edited.map((p) => p.runs.map((r) => r.text).join('')).join('')
  return { ...base, text: nextText, paras: nextParas }
}
