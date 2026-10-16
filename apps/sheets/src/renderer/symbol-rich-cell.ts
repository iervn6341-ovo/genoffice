import { CellValueType, type ICellData, type IStyleData } from '@univerjs/core'
import { appendSymbol } from './SymbolDialog'

interface RichBody {
  readonly dataStream: string
  readonly textRuns?: readonly unknown[]
  readonly paragraphs?: readonly unknown[]
  readonly sectionBreaks?: readonly unknown[]
}

/// Text style in force at `offset`: the run that ends at it (typing continues the previous run),
/// else the run that starts at it.
function textStyleAt(runs: readonly unknown[], offset: number): IStyleData {
  let following: IStyleData | undefined
  for (let index = runs.length - 1; index >= 0; index -= 1) {
    const run = runs[index] as { st?: unknown; ed?: unknown; ts?: unknown } | undefined
    if (
      !run ||
      typeof run.st !== 'number' ||
      typeof run.ed !== 'number' ||
      run.ed <= run.st ||
      typeof run.ts !== 'object' ||
      run.ts === null
    )
      continue
    if (offset > 0 && run.st < offset && run.ed >= offset) return run.ts as IStyleData
    if (run.st <= offset && run.ed > offset) following = run.ts as IStyleData
  }
  return following ?? {}
}

/// Where in the cell text the symbol goes: the editor caret / selection captured when the dialog
/// opened (offsets into the document stream), or the end of the text when absent.
export interface SymbolInsertPoint {
  readonly start: number
  readonly end: number
}

/// Insert a font-specific run at the caret (replacing the selection) without changing the font of
/// the surrounding content. Univer's document stream ends in a paragraph + section break; the
/// symbol is inserted before those terminators.
export function symbolCellPatch(
  existing: string | number | boolean | null | undefined,
  cellData: ICellData | null | undefined,
  char: string,
  font: string,
  baseFont: IStyleData = {},
  point?: SymbolInsertPoint | null,
): ICellData {
  const currentDocument =
    cellData?.p && typeof cellData.p === 'object'
      ? (cellData.p as { body?: RichBody; documentStyle?: unknown })
      : undefined
  const currentBody = currentDocument?.body
  const stream = currentBody?.dataStream
  const prefix =
    typeof stream === 'string'
      ? stream.replace(/\r\n$/, '').replace(/\r$/, '')
      : String(existing ?? '').replace(/\n/g, '\r')
  const from = point
    ? Math.max(0, Math.min(prefix.length, Math.min(point.start, point.end)))
    : prefix.length
  const to = point
    ? Math.max(from, Math.min(prefix.length, Math.max(point.start, point.end)))
    : prefix.length
  const merged = `${prefix.slice(0, from)}${char}${prefix.slice(to)}`
  const value =
    typeof stream === 'string' || point
      ? merged.replace(/\r\n?/g, '\n')
      : appendSymbol(existing, char)
  const dataStream = `${merged}\r\n`
  const priorRuns = Array.isArray(currentBody?.textRuns) ? [...currentBody.textRuns] : []
  if (prefix.length > 0 && priorRuns.length === 0 && Object.keys(baseFont).length > 0) {
    priorRuns.push({ st: 0, ed: prefix.length, ts: baseFont })
  }
  const inserted = { ...baseFont, ...textStyleAt(priorRuns, from), ff: font }
  const delta = char.length - (to - from)
  const shifted: unknown[] = []
  for (const raw of priorRuns) {
    const run = raw as { st?: unknown; ed?: unknown; ts?: unknown }
    if (typeof run.st !== 'number' || typeof run.ed !== 'number') {
      shifted.push(raw)
      continue
    }
    // part before the replaced span keeps its place, the part after it moves by `delta`
    if (run.st < from) shifted.push({ ...run, st: run.st, ed: Math.min(run.ed, from) })
    if (run.ed > to) shifted.push({ ...run, st: Math.max(run.st, to) + delta, ed: run.ed + delta })
  }
  const textRuns = [...shifted, { st: from, ed: from + char.length, ts: inserted }].sort(
    (left, right) => (left as { st: number }).st - (right as { st: number }).st,
  )
  const paragraphs: { startIndex: number }[] = []
  for (let index = 0; index < dataStream.length; index += 1) {
    if (dataStream[index] === '\r') paragraphs.push({ startIndex: index })
  }
  return {
    v: value,
    t: CellValueType.STRING,
    f: null,
    p: {
      id: 'rich-cell',
      body: {
        dataStream,
        textRuns,
        paragraphs,
        sectionBreaks: [{ startIndex: dataStream.length - 1 }],
      },
      documentStyle: currentDocument?.documentStyle ?? {},
    },
  } as ICellData
}
