import {
  IUniverInstanceService,
  UniverInstanceType,
  type IStyleData,
  type Workbook,
} from '@univerjs/core'
import { ISheetClipboardService } from '@univerjs/sheets-ui'
import { HTML_CLIPBOARD_MIME_TYPE, IClipboardInterfaceService } from '@univerjs/ui'

import { foreignFromPayload, foreignFromText, type ForeignCopy } from './clipboard-foreign'
import { extractPayload } from './clipboard-payload'
import { isGridText } from './clipboard-text'
import { foreignFromExcelHtml } from './clipboard-html-import'
import { foreignFromExcelEmbed, parseExcelLink } from './excel-embed'
import { toUniverStyle } from './univer-sync'

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

/**
 * What the system clipboard stands for, local or not: a copy made in this window (the sheet
 * clipboard cache entry the HTML names), a block copied in another GenOffice window (the payload
 * in the HTML), Excel / web table markup, or plain text. `foreign` is set for every source that is
 * not this window's cache.
 */
export interface ClipboardSource {
  readonly copied: CopiedCells
  readonly foreign: ForeignCopy | null
}

/// The foreign copy `html` / `text` stand for, or null when they are this window's own copy
/// (or nothing a sheet can use).
export function foreignFromClipboardData(
  html: string | null | undefined,
  text: string | null | undefined,
  isLocalCopyId: (id: string) => boolean,
): ForeignCopy | null {
  if (html) {
    const id = extractCopyId(html)
    if (id !== null && isLocalCopyId(id)) return null
    const payload = extractPayload(html)
    if (payload) return foreignFromPayload(payload)
    const parsed = foreignFromExcelHtml(html)
    if (parsed) return parsed
    // other table markup (web pages, Docs) is left to Univer's own HTML reader
    if (/<table\b/i.test(html)) return null
  }
  if (text && isGridText(text)) return foreignFromText(text)
  return null
}

/// What Excel copied, read from its embedded workbook: formulas, exact values and real styles.
export async function readExcelCopy(): Promise<ForeignCopy | null> {
  try {
    const api = (
      window as unknown as { desktopApi?: { readExcelClipboard?: () => Promise<unknown> } }
    ).desktopApi
    const raw = (await api?.readExcelClipboard?.()) as { embed?: unknown; link?: unknown } | null
    if (!raw || typeof raw.embed !== 'string' || typeof raw.link !== 'string') return null
    const link = parseExcelLink(raw.link)
    if (!link) return null
    const binary = atob(raw.embed)
    const bytes = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
    return await foreignFromExcelEmbed(bytes, link, { toStyle: (style) => toUniverStyle(style) })
  } catch {
    return null
  }
}

/// `foreignFromClipboardData`, with Excel's embedded copy preferred over its markup.
export async function resolveForeignCopy(
  html: string | null | undefined,
  text: string | null | undefined,
  isLocalCopyId: (id: string) => boolean,
): Promise<ForeignCopy | null> {
  const id = html ? extractCopyId(html) : null
  const own = html ? extractPayload(html) : null
  if (!own && !(id !== null && isLocalCopyId(id))) {
    const excel = await readExcelCopy()
    if (excel) return excel
  }
  return foreignFromClipboardData(html, text, isLocalCopyId)
}

/// The markup and text on the system clipboard exactly as the pasteboard holds them (the async
/// Clipboard API sanitizes markup: Excel's number formats and border colors would be lost).
export async function readSystemClipboard(): Promise<{ html: string; text: string } | null> {
  try {
    const api = (
      window as unknown as { desktopApi?: { readSystemClipboard?: () => Promise<unknown> } }
    ).desktopApi
    const raw = (await api?.readSystemClipboard?.()) as
      { html?: unknown; text?: unknown } | undefined
    if (!raw) return null
    return {
      html: typeof raw.html === 'string' ? raw.html : '',
      text: typeof raw.text === 'string' ? raw.text : '',
    }
  } catch {
    return null
  }
}

export async function readClipboardSource(runtime: UniverRuntime): Promise<ClipboardSource | null> {
  let html: string | null | undefined
  let text: string | null | undefined
  const system = await readSystemClipboard()
  if (system && (system.html || system.text)) {
    html = system.html
    text = system.text
  } else
    try {
      const items = await runtime.univer.__getInjector().get(IClipboardInterfaceService).read()
      html = null
      text = null
      for (const item of items) {
        if (item.types.includes(HTML_CLIPBOARD_MIME_TYPE)) {
          html = await (await item.getType(HTML_CLIPBOARD_MIME_TYPE)).text()
        }
        if (item.types.includes('text/plain')) {
          text = await (await item.getType('text/plain')).text()
        }
        if (html || text) break
      }
    } catch {
      html = undefined
    }
  if (html === undefined) {
    const local = readCopiedCells(runtime)
    return local ? { copied: local, foreign: null } : null
  }
  const cache = runtime.univer.__getInjector().get(ISheetClipboardService).copyContentCache()
  const foreign = await resolveForeignCopy(html, text, (id) => cache.get(id) !== undefined)
  if (foreign) return { copied: foreign.copied, foreign }
  const id = html ? extractCopyId(html) : null
  const local = id === null ? null : readCopiedCells(runtime, id)
  return local ? { copied: local, foreign: null } : null
}
