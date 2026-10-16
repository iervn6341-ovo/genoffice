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

export class HfProtectedFieldEditError extends Error {
  constructor() {
    super('Header/footer fields are protected; keep each field token unchanged.')
  }
}

const editablePara = (para: HfParagraph) => !para.cells && !para.runs.some((run) => run.image)

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
    .filter(editablePara)
    .map((p) => p.runs.map((r) => r.hfField?.token ?? r.text).join(''))
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
function insertionStyle(run: Run | undefined): Partial<Run> {
  if (!run) return {}
  const {
    text: _text,
    hfField: _hfField,
    image: _image,
    math: _math,
    sym: _sym,
    ruby: _ruby,
    noteRef: _noteRef,
    xeTerm: _xeTerm,
    refField: _refField,
    refInstr: _refInstr,
    instrField: _instrField,
    fldBeginXml: _fldBeginXml,
    fldDirty: _fldDirty,
    sdtCheckboxXml: _sdtCheckboxXml,
    zoteroFieldId: _zoteroFieldId,
    zoteroFieldPart: _zoteroFieldPart,
    ...style
  } = run
  return style
}

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
  const pieces: Run[] = [
    ...slice(0, from),
    ...(inserted ? [{ ...insertionStyle(styleRun), text: inserted }] : []),
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
  JSON.stringify(Object.entries({ ...a, text: '' }).sort(([a], [b]) => a.localeCompare(b))) ===
  JSON.stringify(Object.entries({ ...b, text: '' }).sort(([a], [b]) => a.localeCompare(b)))

/**
 * Map edited lines back onto the part: each line keeps its original
 * paragraph's format and the runs the edit did not touch (extra lines reuse
 * the last template's first-run styling), cells rows are spliced back at
 * their original positions.
 */
export function applyHfText(value: HeaderFooter | null, text: string): HeaderFooter {
  const base = value ?? { text: '' }
  const paras = hfParasOf(base)
  if (paras.some((para) => para.protectedFields)) throw new HfProtectedFieldEditError()
  const source = text.replace(/\n+$/, '')
  const textParas = paras.filter(editablePara)
  const templates: HfParagraph[] =
    textParas.length > 0 ? textParas : [{ align: 'center', runs: [] }]
  const flat: Run[] = textParas.flatMap((para, i) => [
    ...(i > 0 ? [{ text: '\n' }] : []),
    ...para.runs.map((run) => ({
      ...run,
      text:
        run.hfField?.token ??
        run.text.replaceAll(PAGE_MARK, PAGE_TOKEN).replaceAll(TOTAL_PAGES_MARK, TOTAL_TOKEN),
    })),
  ])
  const fields = flat.filter((run) => run.hfField)
  if (fields.length === 0) {
    const lines = source
      .replaceAll(PAGE_TOKEN, PAGE_MARK)
      .replaceAll(TOTAL_TOKEN, TOTAL_PAGES_MARK)
      .split('\n')
    const edited = lines.map((line, i): HfParagraph => {
      const template = templates[Math.min(i, templates.length - 1)]
      if (line === '') return { ...template, runs: [] }
      if (i < textParas.length) return { ...template, runs: spliceLineRuns(template.runs, line) }
      return { ...template, runs: [{ ...insertionStyle(template.runs[0]), text: line }] }
    })
    return finishHfEdit(base, paras, edited)
  }
  const segments: Run[][] = [[]]
  for (const run of flat) {
    if (run.hfField) segments.push([])
    else segments[segments.length - 1].push(run)
  }
  const editedSegments: string[] = []
  let cursor = 0
  for (const field of fields) {
    const token = field.hfField!.token
    const at = source.indexOf(token, cursor)
    if (at < 0 || source.indexOf(token, at + token.length) >= 0)
      throw new HfProtectedFieldEditError()
    editedSegments.push(source.slice(cursor, at))
    cursor = at + token.length
  }
  editedSegments.push(source.slice(cursor))
  if (
    fields.some((field) => editedSegments.some((segment) => segment.includes(field.hfField!.token)))
  ) {
    throw new HfProtectedFieldEditError()
  }
  const flatEdited: Run[] = []
  for (let i = 0; i < editedSegments.length; i++) {
    flatEdited.push(...spliceLineRuns(segments[i], editedSegments[i]))
    if (i < fields.length) flatEdited.push(fields[i])
  }
  const lineRuns: Run[][] = [[]]
  for (const run of flatEdited) {
    if (run.hfField) {
      lineRuns[lineRuns.length - 1].push({ ...run, text: run.hfField.displayText })
      continue
    }
    const parts = run.text.split('\n')
    parts.forEach((part, i) => {
      if (i > 0) lineRuns.push([])
      if (part)
        lineRuns[lineRuns.length - 1].push({
          ...run,
          text: part.replaceAll(PAGE_TOKEN, PAGE_MARK).replaceAll(TOTAL_TOKEN, TOTAL_PAGES_MARK),
        })
    })
  }
  const edited: HfParagraph[] = lineRuns.map((runs, i) => ({
    ...templates[Math.min(i, templates.length - 1)],
    runs: runs.reduce<Run[]>((joined, run) => {
      const last = joined[joined.length - 1]
      if (last && !last.hfField && !run.hfField && sameFormat(last, run)) {
        joined[joined.length - 1] = { ...last, text: last.text + run.text }
      } else joined.push(run)
      return joined
    }, []),
  }))
  return finishHfEdit(base, paras, edited)
}

function finishHfEdit(
  base: HeaderFooter,
  paras: HfParagraph[],
  edited: HfParagraph[],
): HeaderFooter {
  const nextParas: HfParagraph[] = []
  let ei = 0
  let lastEditable = -1
  for (let index = 0; index < paras.length; index++) {
    if (editablePara(paras[index])) lastEditable = index
  }
  for (const [index, p] of paras.entries()) {
    if (!editablePara(p)) nextParas.push(p)
    else if (index === lastEditable) {
      nextParas.push(...edited.slice(ei))
      ei = edited.length
    } else if (ei < edited.length) nextParas.push(edited[ei++])
  }
  nextParas.push(...edited.slice(ei))
  const nextText = edited.map((p) => p.runs.map((r) => r.text).join('')).join('')
  const pageInCells = nextParas.some((p) =>
    p.cells?.some((c) => c.paras.some((runs) => runs.some((run) => run.text.includes(PAGE_MARK)))),
  )
  const pageNumber =
    nextText.includes(PAGE_MARK) ||
    pageInCells ||
    (base.pageNumber === true && !hfEditText(base).includes(PAGE_TOKEN) && nextText.includes('#'))
  return { ...base, text: nextText, pageNumber, paras: nextParas }
}

export type HfNativeFieldKind = 'DATE' | 'FILENAME' | 'AUTHOR' | 'PAGE' | 'NUMPAGES'
export const HF_DATE_FORMATS = ['yyyy-MM-dd', 'MM/dd/yyyy', 'dd/MM/yyyy', 'yyyy/MM/dd'] as const
export type HfDateFormat = (typeof HF_DATE_FORMATS)[number]

export function formatHfDate(now: Date, format: HfDateFormat): string {
  const yyyy = String(now.getFullYear())
  const MM = String(now.getMonth() + 1).padStart(2, '0')
  const dd = String(now.getDate()).padStart(2, '0')
  return format.replace(/yyyy|MM|dd/g, (part) => ({ yyyy, MM, dd })[part] ?? part)
}

const xmlText = (text: string) =>
  text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
const xmlAttr = (text: string) => xmlText(text).replaceAll('"', '&quot;')

/** Insert one atomic run at the textarea selection after validating existing field tokens. */
export function insertHfRunAtSelection(
  value: HeaderFooter | null,
  editText: string,
  start: number,
  end: number,
  atom: Run,
  token: string,
): { value: HeaderFooter; text: string; caret: number } {
  const base = applyHfText(value, editText)
  const from = Math.max(0, Math.min(start, editText.length))
  const to = Math.max(from, Math.min(end, editText.length))
  let serial = 0
  let marker: string
  do marker = `\uE102hf-insert-${++serial}\uE103`
  while (editText.includes(marker))
  const insertedText = editText.slice(0, from) + marker + editText.slice(to)
  const inserted = applyHfText(base, insertedText)
  let found = 0
  const paras = inserted.paras?.map((para) => ({
    ...para,
    runs: para.runs.flatMap((run): Run[] => {
      const at = run.text.indexOf(marker)
      if (at < 0) return [run]
      found++
      return [
        ...(at > 0 ? [{ ...run, text: run.text.slice(0, at) }] : []),
        atom,
        ...(at + marker.length < run.text.length
          ? [{ ...run, text: run.text.slice(at + marker.length) }]
          : []),
      ]
    }),
  }))
  if (found !== 1 || !paras) throw new HfProtectedFieldEditError()
  const text = insertedText.replace(marker, token)
  const nextValue = {
    ...inserted,
    paras,
    text: inserted.text.replace(marker, atom.text),
    pageNumber: inserted.pageNumber || atom.text === PAGE_MARK,
  }
  return { value: nextValue, text, caret: from + token.length }
}

/** Supported Word field instructions are fixed here; arbitrary instruction/XML input is excluded. */
export function insertHfNativeField(
  value: HeaderFooter | null,
  editText: string,
  start: number,
  end: number,
  kind: HfNativeFieldKind,
  fileName = '',
  now = new Date(),
  dateFormat: HfDateFormat = 'yyyy-MM-dd',
  author = '',
): { value: HeaderFooter; text: string; caret: number } {
  const instruction = kind === 'DATE' ? ` DATE \\@ "${dateFormat}" ` : ` ${kind} `
  const cache =
    kind === 'DATE'
      ? formatHfDate(now, dateFormat)
      : kind === 'FILENAME'
        ? fileName
        : kind === 'AUTHOR'
          ? author
          : kind === 'PAGE' || kind === 'NUMPAGES'
            ? '1'
            : ''
  let serial = 1
  let token = `⟦${kind}·${serial}⟧`
  while (editText.includes(token)) token = `⟦${kind}·${++serial}⟧`
  const displayText = kind === 'PAGE' ? PAGE_MARK : kind === 'NUMPAGES' ? TOTAL_PAGES_MARK : cache
  const atom: Run = {
    text: displayText,
    hfField: {
      xml: `<w:fldSimple w:instr="${xmlAttr(instruction)}" w:dirty="true"><w:r><w:t xml:space="preserve">${xmlText(cache)}</w:t></w:r></w:fldSimple>`,
      instruction,
      token,
      displayText,
      resultRuns: cache ? [{ text: cache }] : [],
    },
  }
  return insertHfRunAtSelection(value, editText, start, end, atom, token)
}

/** Place a picture paragraph after the textarea line at the selection. */
export function insertHfPictureAtSelection(
  value: HeaderFooter | null,
  editText: string,
  start: number,
  end: number,
  image: {
    base64: string
    mime: 'image/png' | 'image/jpeg' | 'image/gif'
    widthPx: number
    heightPx: number
  },
): { value: HeaderFooter; text: string; caret: number } {
  const from = Math.max(0, Math.min(start, editText.length))
  const to = Math.max(from, Math.min(end, editText.length))
  const split = from > 0 && to < editText.length
  const nextText = editText.slice(0, from) + (split ? '\n' : '') + editText.slice(to)
  const base = applyHfText(value, nextText)
  const line = editText.slice(0, from).split('\n').length - 1
  const paras = [...(base.paras ?? [])]
  let editable = -1
  let insertAt = paras.length
  for (let index = 0; index < paras.length; index++) {
    if (!editablePara(paras[index])) continue
    editable++
    if (editable === line) {
      insertAt = index + 1
      break
    }
  }
  paras.splice(insertAt, 0, {
    align: 'left',
    runs: [
      {
        text: '',
        image: {
          dataUrl: `data:${image.mime};base64,${image.base64}`,
          widthPx: image.widthPx,
          heightPx: image.heightPx,
          xml: '',
          pendingMedia: image,
        },
      },
    ],
  })
  const nextValue = { ...base, paras }
  return { value: nextValue, text: hfEditText(nextValue), caret: from + (split ? 1 : 0) }
}

export type HfLayoutKind = 'blank' | 'threeColumn' | 'datePage'

/** Append an editable native layout; existing fields, tables and pictures stay in place. */
export function appendHfLayout(
  value: HeaderFooter | null,
  editText: string,
  kind: HfLayoutKind,
  fileName = '',
  now = new Date(),
  dateFormat: HfDateFormat = 'yyyy-MM-dd',
): { value: HeaderFooter; text: string; caret: number } {
  const current = applyHfText(value, editText)
  const prefix = editText.length > 0 ? '\n' : ''
  const seed = kind === 'threeColumn' ? '\t\t' : ''
  let text = editText + prefix + seed
  let next = applyHfText(current, text)
  if (kind === 'datePage') {
    const date = insertHfNativeField(
      next,
      text,
      text.length,
      text.length,
      'DATE',
      fileName,
      now,
      dateFormat,
    )
    text = date.text + '\t'
    next = applyHfText(date.value, text)
    const page = insertHfNativeField(next, text, text.length, text.length, 'PAGE', fileName, now)
    text = page.text
    next = page.value
  }
  if (kind !== 'blank') {
    const paras = [...(next.paras ?? [])]
    let last = -1
    for (let index = 0; index < paras.length; index++) {
      if (editablePara(paras[index])) last = index
    }
    if (last >= 0) {
      paras[last] = {
        ...paras[last],
        align: 'left',
        tabStops:
          kind === 'threeColumn'
            ? [
                { pos: 4680, val: 'center' },
                { pos: 9360, val: 'right' },
              ]
            : [{ pos: 9360, val: 'right' }],
      }
      next = { ...next, paras }
    }
  }
  return { value: next, text, caret: text.length }
}
