import {
  IUniverInstanceService,
  UniverInstanceType,
  type IStyleData,
  type Workbook,
} from '@univerjs/core'
import { ISheetClipboardService } from '@univerjs/sheets-ui'
import { HTML_CLIPBOARD_MIME_TYPE, IClipboardInterfaceService } from '@univerjs/ui'

import type { UniverRuntime } from './univer-state'

/// One copied cell as the sheet clipboard cache holds it (style resolved to the style object).
export interface CopiedCell {
  readonly v?: string | number | boolean | null
  readonly t?: number | null
  readonly f?: string | null
  readonly si?: string | null
  /// the cache stores a style id; readCopiedCells hands back the style object
  readonly s?: IStyleData | string | null
  readonly p?: unknown
  readonly rowSpan?: number
  readonly colSpan?: number
}

export interface CopiedCells {
  readonly unitId: string
  readonly sheetId: string
  readonly cut: boolean
  /// absolute rows / columns of the copied range, in order (filtered rows are not in the list)
  readonly rows: readonly number[]
  readonly columns: readonly number[]
  /// cells[i][j] belongs to rows[i] × columns[j]; null = empty
  readonly cells: readonly (readonly (CopiedCell | null)[])[]
}

/**
 * What Copy (or Cut) left in the sheet clipboard cache — the source Paste Special works from.
 * `copyId` picks the entry the system clipboard names; without it the last copy of this window is
 * used. Null when there is none (a paste from another program has no cache).
 */
export function readCopiedCells(
  runtime: UniverRuntime,
  copyId?: string | null,
): CopiedCells | null {
  try {
    const cache = runtime.univer.__getInjector().get(ISheetClipboardService).copyContentCache()
    const id = copyId ?? cache.getLastCopyId()
    const entry = id ? cache.get(id) : undefined
    if (!entry?.matrix) return null
    const { rows, cols } = entry.range as { rows: number[]; cols: number[] }
    const matrix = entry.matrix
    const injector = runtime.univer.__getInjector()
    const workbook = injector
      .get(IUniverInstanceService)
      .getUnit<Workbook>(entry.unitId, UniverInstanceType.UNIVER_SHEET)
    const styles = workbook?.getStyles()
    const cells = rows.map((_, i) =>
      cols.map((_, j) => {
        const cell = (matrix.getValue(i, j) as CopiedCell | undefined | null) ?? null
        if (cell === null || typeof cell.s !== 'string') return cell
        return { ...cell, s: styles?.get(cell.s) ?? null }
      }),
    )
    return {
      unitId: entry.unitId,
      sheetId: entry.subUnitId,
      cut: String(entry.copyType) === 'CUT',
      rows: [...rows],
      columns: [...cols],
      cells,
    }
  } catch {
    return null
  }
}

/** the id Univer writes into the html it puts on the system clipboard (`data-copy-id="…"`) */
export function extractCopyId(html: string): string | null {
  const match = /data-copy-id="([^\s"]+)"/.exec(html)
  return match?.[1] ?? null
}

/**
 * The cells the system clipboard stands for, like Univer's own paste: when the clipboard holds
 * html from another program (or other text) the earlier sheet copy is no longer what was copied,
 * so nothing is returned. When the clipboard cannot be read (no permission, no focus) the last
 * copy of this window stands in.
 */
export async function readClipboardCopy(runtime: UniverRuntime): Promise<CopiedCells | null> {
  let html: string | null | undefined
  try {
    const items = await runtime.univer.__getInjector().get(IClipboardInterfaceService).read()
    html = null
    for (const item of items) {
      if (item.types.includes(HTML_CLIPBOARD_MIME_TYPE)) {
        html = await (await item.getType(HTML_CLIPBOARD_MIME_TYPE)).text()
        break
      }
    }
  } catch {
    html = undefined
  }
  if (html === undefined) return readCopiedCells(runtime)
  const id = html === null ? null : extractCopyId(html)
  return id === null ? null : readCopiedCells(runtime, id)
}
