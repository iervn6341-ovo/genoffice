/**
 * A copy that did not come from this window's own sheet clipboard cache: another GenOffice
 * workbook window (payload in the HTML), Excel or a web page (table markup), or plain text.
 * Each is turned into the same shape as a local copy (`CopiedCells` plus the sheet features that
 * travel with the block) so Paste and Paste Special run one code path for all of them.
 */
import type { IRange, IStyleData } from '@univerjs/core'

import { type ClipboardPayload, type PayloadLink, type PayloadNote } from './clipboard-payload'
import { nameReferences } from './clipboard-formula'
import { parseDelimitedText, typeField } from './clipboard-text'
import type { CopiedCell, CopiedCells } from './paste-source'

export interface ForeignCopy {
  readonly origin: 'genoffice' | 'excel' | 'html' | 'text'
  readonly copied: CopiedCells
  readonly sheetName: string
  /// absolute source coordinates, like `copied.rows` / `copied.columns`
  readonly merges: readonly IRange[]
  readonly columnWidths: readonly (number | null)[]
  readonly rowHeights: readonly (number | null)[]
  readonly notes: readonly PayloadNote[]
  readonly validations: readonly unknown[]
  readonly conditionalFormats: readonly unknown[]
  readonly links: readonly PayloadLink[]
  /// The source address is unknown (Excel's HTML carries none), so formulas land exactly as
  /// written instead of being moved by the distance between source and destination.
  readonly literalFormulas: boolean
  /// the block was cut in another window (its copy id); that window clears it after the paste
  readonly cutCopyId?: string
}

/// Sheet names a formula points to by `Name!A1` / `'My Sheet'!A1` (not `#REF!`, not inside text).
export function hasSheetReference(formula: string): boolean {
  let inText = false
  for (let i = 0; i < formula.length; i += 1) {
    const ch = formula[i]!
    if (ch === '"') {
      inText = !inText
      continue
    }
    if (inText) continue
    if (ch === "'") {
      // quoted sheet name: runs to the closing quote ('' escapes a quote)
      let j = i + 1
      while (j < formula.length) {
        if (formula[j] === "'") {
          if (formula[j + 1] === "'") {
            j += 2
            continue
          }
          break
        }
        j += 1
      }
      if (formula[j + 1] === '!') return true
      i = j
      continue
    }
    if (ch === '!') {
      const before = formula.slice(Math.max(0, i - 4), i).toUpperCase()
      if (before.endsWith('#REF')) continue
      return true
    }
  }
  return false
}

export function foreignFromPayload(payload: ClipboardPayload): ForeignCopy {
  const cells: (CopiedCell | null)[][] = payload.cells.map((line) =>
    line.map((cell) => {
      if (cell === 0) return null
      const style = cell.s === undefined ? null : (payload.styles[cell.s] ?? null)
      return {
        ...(cell.v === undefined ? {} : { v: cell.v }),
        ...(cell.t === undefined ? {} : { t: cell.t }),
        ...(cell.f === undefined ? {} : { f: cell.f }),
        ...(cell.p === undefined ? {} : { p: cell.p }),
        s: style,
      } satisfies CopiedCell
    }),
  )
  return {
    origin: 'genoffice',
    copied: {
      unitId: payload.bookId,
      sheetId: payload.sheetId,
      cut: false,
      rows: payload.rows,
      columns: payload.columns,
      cells,
    },
    sheetName: payload.sheetName,
    merges: payload.merges,
    columnWidths: payload.columnWidths,
    rowHeights: payload.rowHeights,
    notes: payload.notes,
    validations: payload.validations,
    conditionalFormats: payload.conditionalFormats,
    links: payload.links,
    literalFormulas: false,
    ...(payload.cut && payload.copyId ? { cutCopyId: payload.copyId } : {}),
  }
}

/// Plain text pasted the way Excel reads it: tab-separated columns, one row per line.
export function foreignFromText(text: string): ForeignCopy {
  const grid = parseDelimitedText(text)
  const columnCount = Math.max(1, ...grid.map((line) => line.length))
  const cells: (CopiedCell | null)[][] = grid.map((line) => {
    const out: (CopiedCell | null)[] = []
    for (let j = 0; j < columnCount; j += 1) {
      const raw = line[j]
      if (raw === undefined || raw === '') {
        out.push(null)
        continue
      }
      const typed = typeField(raw)
      const style: IStyleData | null = typed.pattern ? { n: { pattern: typed.pattern } } : null
      out.push({
        ...(typed.f ? { f: typed.f } : { v: typed.v }),
        ...(typed.t === null ? {} : { t: typed.t }),
        s: style,
      })
    }
    return out
  })
  return {
    origin: 'text',
    copied: {
      unitId: '',
      sheetId: '',
      cut: false,
      rows: grid.map((_, i) => i),
      columns: Array.from({ length: columnCount }, (_, j) => j),
      cells,
    },
    sheetName: '',
    merges: [],
    columnWidths: [],
    rowHeights: [],
    notes: [],
    validations: [],
    conditionalFormats: [],
    links: [],
    literalFormulas: true,
  }
}

/// A formula of a block pasted into another workbook: references to other sheets cannot follow
/// (Excel makes an external link to the source file, which this app cannot keep), so such a
/// formula lands as its last result.
export function foreignFormula(
  formula: string | null | undefined,
  knownNames: ReadonlySet<string> = new Set(),
): string | null {
  if (!formula) return null
  if (hasSheetReference(formula)) return null
  // a defined name of the source workbook means nothing here unless this workbook has it too
  return nameReferences(formula).every((name) => knownNames.has(name)) ? formula : null
}

/// A minimal stand-in for the facade worksheet the paste extras read from, backed by a copy that
/// lives in another window.
export interface VirtualSourceSheet {
  getMergedRanges(): {
    getRow(): number
    getColumn(): number
    getHeight(): number
    getWidth(): number
  }[]
  getNotes(): { row: number; col: number; note: string }[]
  getDataValidations(): { rule: { uid: string; ranges?: IRange[] } & Record<string, unknown> }[]
  getConditionalFormattingRules(): ({ ranges: IRange[] } & Record<string, unknown>)[]
}

export function virtualSourceSheet(copy: ForeignCopy): VirtualSourceSheet {
  return {
    getMergedRanges: () =>
      copy.merges.map((range) => ({
        getRow: () => range.startRow,
        getColumn: () => range.startColumn,
        getHeight: () => range.endRow - range.startRow + 1,
        getWidth: () => range.endColumn - range.startColumn + 1,
      })),
    getNotes: () => copy.notes.map((n) => ({ row: n.row, col: n.column, note: n.text })),
    getDataValidations: () =>
      (copy.validations as ({ uid: string; ranges?: IRange[] } & Record<string, unknown>)[]).map(
        (rule) => ({ rule }),
      ),
    getConditionalFormattingRules: () =>
      copy.conditionalFormats as ({ ranges: IRange[] } & Record<string, unknown>)[],
  }
}
