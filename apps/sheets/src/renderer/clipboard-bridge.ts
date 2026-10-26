/**
 * Cross-window clipboard bridge.
 *
 * Copy: the block's cells, formulas, styles, merges, sizes, notes, validation, conditional formats
 * and links go into the HTML as a GenOffice payload, and the HTML itself is written in Excel's
 * dialect (see clipboard-html.ts) so Excel pastes numbers, formats and borders as they were.
 *
 * Paste: Univer pastes from its per-window cache when the HTML names an entry of it, and
 * otherwise reads the markup. A copy made in another workbook window, in Excel, or as plain text
 * is therefore parsed here instead (payload, Excel markup, tab-separated text) and pasted through
 * the Paste Special plan, which keeps formulas, formats, merges and the sheet features, and
 * undoes in one step.
 */
import {
  CellModeEnum,
  IUniverInstanceService,
  UniverInstanceType,
  extractPureTextFromCell,
  type ICellData,
  type IRange,
  type IStyleData,
  type Workbook,
} from '@univerjs/core'
import { ISheetClipboardService } from '@univerjs/sheets-ui'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import { SheetsSelectionsService } from '@univerjs/sheets'
import { IClipboardInterfaceService } from '@univerjs/ui'

import type { ForeignCopy } from './clipboard-foreign'
import { buildClipboardHtml, type HtmlBlockCell } from './clipboard-html'
import { readSystemClipboard, resolveForeignCopy } from './paste-source'
import {
  PAYLOAD_VERSION,
  createStyleTable,
  encodePayload,
  usedExtent,
  type ClipboardPayload,
  type PayloadCell,
  type PayloadLink,
  type PayloadNote,
} from './clipboard-payload'
import { clipboardField, plainTextOfBlock, withFullDisplay } from './clipboard-tsv'
import { multiSelectionLayout } from './clipboard-multi'
import { extractVisualPayload, pasteVisualPayload } from './visual-clipboard-actions'
import { t } from './i18n/locale'
import { clearWithExtras, hyperlinkTargetAt, setHyperlinksUndoable } from './clear-actions'
import { formulaForExcelClipboard, nameReferences } from './clipboard-formula'
import { hasSheetReference } from './clipboard-foreign'
import { makeSharedFormulaResolver } from './shared-formula-journal'
import { pasteSpecial } from './paste-special-actions'
import { runAsOneUndoStepAsync } from './undo-group'
import { linkEditsForLocalPaste } from './clipboard-links'
import { DEFAULT_PASTE_SPECIAL, type PasteWhat } from './paste-special'
import type { RibbonCommandContext } from './ribbon-actions'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

type ClipboardCell = ICellData & { displayV?: string; rowSpan?: number; colSpan?: number }

interface CopyContent {
  copyId: string
  plain: string
  html: string
  discreteRange: { rows: number[]; cols: number[] }
}

const intersects = (a: IRange, b: IRange): boolean =>
  a.startRow <= b.endRow &&
  b.startRow <= a.endRow &&
  a.startColumn <= b.endColumn &&
  b.startColumn <= a.endColumn

/// A selection of whole rows / columns spans the sheet's full width / height; only the used part
/// belongs on the clipboard (Excel writes the used range of a whole-row copy, not 16,384 columns).
const WHOLE_AXIS_COLUMNS = 1000
const WHOLE_AXIS_ROWS = 100_000

export interface BridgeDeps {
  getContext(): RibbonCommandContext
  getState(): LazyWorkbookState | null
}

const hookToWhat: Record<string, PasteWhat> = {
  'default-paste': 'all',
  'special-paste-value': 'values',
  'special-paste-formula': 'formulas',
  'special-paste-format': 'formats',
  'special-paste-col-width': 'columnWidths',
  'special-paste-besides-border': 'allExceptBorders',
}

export function installClipboardBridge(
  runtime: UniverRuntime,
  deps: BridgeDeps,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const instances = injector.get(IUniverInstanceService)
  const service = injector.get(ISheetClipboardService) as unknown as {
    generateCopyContent(
      workbookId: string,
      worksheetId: string,
      range: IRange,
      options?: { copyHookType?: string; copyType?: string },
    ): CopyContent | null | undefined
    copy(options?: { copyType?: string; copyHookType?: string }): Promise<boolean>
    paste(item: ClipboardItem, pasteType?: string): Promise<boolean>
    legacyPaste(html?: string, text?: string, files?: File[]): Promise<boolean>
    copyContentCache(): { get(id: string): unknown }
    removeMarkSelection?: () => void
  }
  const lexer = injector.get(LexerTreeBuilder)
  const originalGenerate = service.generateCopyContent.bind(service)
  const originalCopy = service.copy.bind(service)
  const originalPaste = service.paste.bind(service)
  const originalLegacyPaste = service.legacyPaste.bind(service)

  // Univer's own paste of a block copied in this window: hyperlinks live in the edit journal,
  // not in the cells, so they ride along here
  const internal = service as unknown as {
    _pasteInternal(copyId: string, pasteType: unknown): Promise<boolean>
    _getPastedRange(cellMatrix: unknown): unknown
  }
  const originalPasteInternal = internal._pasteInternal.bind(service)
  const originalGetPastedRange = internal._getPastedRange.bind(service)
  let lastTarget: {
    unitId: string
    subUnitId: string
    pastedRange: { rows: number[]; cols: number[] }
  } | null = null
  internal._getPastedRange = (cellMatrix: unknown): unknown => {
    const target = originalGetPastedRange(cellMatrix) as typeof lastTarget | undefined
    lastTarget = target ?? null
    return target
  }
  internal._pasteInternal = async (copyId: string, pasteType: unknown): Promise<boolean> => {
    const entry = service.copyContentCache().get(copyId) as
      | {
          subUnitId?: string
          range?: { rows: number[]; cols: number[] }
          copyType?: string
        }
      | undefined
    const state = deps.getState()
    const carries =
      pasteType === undefined ||
      pasteType === 'default-paste' ||
      pasteType === 'special-paste-besides-border'
    if (!entry?.range || !entry.subUnitId || !state || !carries) {
      return originalPasteInternal(copyId, pasteType)
    }
    let result = false
    lastTarget = null
    await runAsOneUndoStepAsync(runtime, async () => {
      result = await originalPasteInternal(copyId, pasteType)
      const target = lastTarget as typeof lastTarget
      if (!result || !target) return
      const edits = linkEditsForLocalPaste({
        sourceRows: entry.range!.rows,
        sourceColumns: entry.range!.cols,
        targetRows: target.pastedRange.rows,
        targetColumns: target.pastedRange.cols,
        cut: entry.copyType === 'CUT',
        sourceTarget: (row, column) => hyperlinkTargetAt(state, entry.subUnitId!, row, column),
        sourceSheetId: entry.subUnitId!,
        targetSheetId: target.subUnitId,
        targetLink: (row, column) => hyperlinkTargetAt(state, target.subUnitId, row, column),
      })
      for (const [sheetId, links] of edits) {
        await setHyperlinksUndoable(deps.getContext(), runtime, sheetId, links)
      }
    })
    return result
  }

  const hasLocalCopy = (id: string): boolean => service.copyContentCache().get(id) !== undefined

  /// blocks cut in this window and not pasted elsewhere yet, by copy id
  const cutPending = new Map<string, { sheetId: string; range: IRange }>()

  service.generateCopyContent = (workbookId, worksheetId, range, options) => {
    const content = originalGenerate(workbookId, worksheetId, range, options)
    if (!content || (options?.copyHookType && options.copyHookType !== 'default-copy')) {
      return content
    }
    try {
      const html = buildCopyHtml(
        workbookId,
        worksheetId,
        range,
        content,
        options?.copyType === 'CUT',
      )
      return html ? { ...content, html } : content
    } catch {
      return content
    }
  }

  function buildCopyHtml(
    workbookId: string,
    worksheetId: string,
    range: IRange,
    content: CopyContent,
    isCut: boolean,
  ): string | null {
    const workbook = instances.getUniverSheetInstance(workbookId)
    const worksheet = workbook?.getSheetBySheetId(worksheetId)
    const facade = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(worksheetId)
    if (!workbook || !worksheet) return null
    let { rows, cols } = content.discreteRange
    if (rows.length === 0 || cols.length === 0) return null
    const matrix = worksheet.getMatrixWithMergedCells(
      Math.min(...rows),
      Math.min(...cols),
      Math.max(...rows),
      Math.max(...cols),
      CellModeEnum.Both,
    )
    const state = deps.getState()
    const resolveShared = makeSharedFormulaResolver(
      () => worksheet.getCellMatrix(),
      (formula, dx, dy) => lexer.moveFormulaRefOffset(formula, dx, dy),
    )
    const rawAt = (row: number, column: number): ClipboardCell | null =>
      (matrix.getValue(row, column) as ClipboardCell | undefined) ?? null
    const wholeAxis = cols.length >= WHOLE_AXIS_COLUMNS || rows.length >= WHOLE_AXIS_ROWS
    if (wholeAxis) {
      const rowIndex = new Map(rows.map((row, i) => [row, i] as const))
      const columnIndex = new Map(cols.map((column, j) => [column, j] as const))
      const used: [number, number][] = []
      matrix.forValue((row, column, cell) => {
        const value = cell as ClipboardCell
        const holds =
          (value.v !== undefined && value.v !== null && value.v !== '') ||
          !!value.f ||
          !!value.p ||
          (typeof value.s === 'object' && value.s !== null && Object.keys(value.s).length > 0)
        const i = rowIndex.get(row)
        const j = columnIndex.get(column)
        if (holds && i !== undefined && j !== undefined) used.push([i, j])
      })
      const extent = usedExtent(used)
      rows = rows.slice(0, extent.rows)
      cols = cols.slice(0, extent.columns)
    }
    const styles = createStyleTable()
    const payloadCells: (PayloadCell | 0)[][] = []
    const htmlCells: HtmlBlockCell[][] = []
    const merges: IRange[] = []
    const block: IRange = {
      startRow: rows[0]!,
      endRow: rows[rows.length - 1]!,
      startColumn: cols[0]!,
      endColumn: cols[cols.length - 1]!,
    }
    for (const merge of worksheet.getMergeData()) {
      if (!intersects(merge, block)) continue
      if (
        merge.startRow >= block.startRow &&
        merge.endRow <= block.endRow &&
        merge.startColumn >= block.startColumn &&
        merge.endColumn <= block.endColumn
      ) {
        merges.push({ ...merge })
      }
    }
    const covered = new Set<string>()
    for (const merge of merges) {
      for (let r = merge.startRow; r <= merge.endRow; r += 1) {
        for (let c = merge.startColumn; c <= merge.endColumn; c += 1) {
          if (r !== merge.startRow || c !== merge.startColumn) covered.add(`${r}:${c}`)
        }
      }
    }
    const mergeAt = new Map<string, IRange>()
    for (const merge of merges) mergeAt.set(`${merge.startRow}:${merge.startColumn}`, merge)
    const links: PayloadLink[] = []
    for (const row of rows) {
      const payloadLine: (PayloadCell | 0)[] = []
      const htmlLine: HtmlBlockCell[] = []
      for (const column of cols) {
        const cell = rawAt(row, column)
        const raw = worksheet.getCellRaw(row, column)
        const isCovered = covered.has(`${row}:${column}`)
        // visual style (conditional formats and number-format colors included) for the markup,
        // the cell's own composed style for the payload
        const visual = worksheet.getComposedCellStyle(row, column) as IStyleData | undefined
        const own =
          (worksheet.getComposedCellStyleByCellData(row, column, raw) as IStyleData | undefined) ??
          null
        const formula =
          raw?.f ||
          (raw?.si ? (resolveShared(row, column, raw.si) ?? undefined) : undefined) ||
          undefined
        const merge = mergeAt.get(`${row}:${column}`)
        const link = state ? hyperlinkTargetAt(state, worksheetId, row, column) : null
        if (link) links.push({ row, column, target: link })
        const styleIndex = styles.indexOf(own)
        const hasContent =
          raw !== null &&
          raw !== undefined &&
          ((raw.v !== undefined && raw.v !== null && raw.v !== '') || !!formula || !!raw.p)
        if (!hasContent && styleIndex === undefined) payloadLine.push(0)
        else {
          payloadLine.push({
            ...(raw?.v === undefined || raw.v === null ? {} : { v: raw.v }),
            ...(raw?.t === undefined || raw.t === null ? {} : { t: raw.t }),
            ...(formula ? { f: formula } : {}),
            ...(styleIndex === undefined ? {} : { s: styleIndex }),
            ...(raw?.p ? { p: raw.p } : {}),
          })
        }
        // Excel can keep a formula only when it points at nothing outside this sheet
        const excelFormula =
          formula && !hasSheetReference(formula) && nameReferences(formula).length === 0
            ? formulaForExcelClipboard(formula, block.startRow, block.startColumn)
            : null
        htmlLine.push({
          text: cell ? displayText(withFullDisplay(cell, own?.n?.pattern) ?? cell) : '',
          ...(raw?.v === undefined || raw.v === null ? {} : { v: raw.v }),
          ...(raw?.t === undefined || raw.t === null ? {} : { t: raw.t }),
          ...(excelFormula ? { f: excelFormula } : {}),
          style: visual ?? own,
          ...(merge
            ? {
                rowSpan: merge.endRow - merge.startRow + 1,
                colSpan: merge.endColumn - merge.startColumn + 1,
              }
            : {}),
          ...(isCovered ? { covered: true } : {}),
          ...(link ? { link } : {}),
        })
      }
      payloadCells.push(payloadLine)
      htmlCells.push(htmlLine)
    }
    const notes: PayloadNote[] = []
    const validations: unknown[] = []
    const conditionalFormats: unknown[] = []
    try {
      for (const note of facade?.getNotes() ?? []) {
        if (
          note.row >= block.startRow &&
          note.row <= block.endRow &&
          note.col >= block.startColumn &&
          note.col <= block.endColumn
        ) {
          notes.push({ row: note.row, column: note.col, text: note.note })
        }
      }
      for (const validation of facade?.getDataValidations() ?? []) {
        const rule = validation.rule as { ranges?: IRange[] }
        if ((rule.ranges ?? []).some((r) => intersects(r, block))) {
          validations.push(JSON.parse(JSON.stringify(rule)))
        }
      }
      for (const rule of facade?.getConditionalFormattingRules() ?? []) {
        if ((rule.ranges ?? []).some((r: IRange) => intersects(r, block))) {
          conditionalFormats.push(JSON.parse(JSON.stringify(rule)))
        }
      }
    } catch {
      // a sheet feature that cannot be read just does not travel
    }
    const payload: ClipboardPayload = {
      version: PAYLOAD_VERSION,
      bookId: workbookId,
      sheetId: worksheetId,
      sheetName: worksheet.getName(),
      rows,
      columns: cols,
      styles: styles.styles,
      cells: payloadCells,
      merges,
      columnWidths: cols.map((c) => worksheet.getColumnWidth(c)),
      rowHeights: rows.map((r) => worksheet.getRowHeight(r)),
      notes,
      validations,
      conditionalFormats,
      links,
      ...(isCut ? { cut: true, copyId: content.copyId } : {}),
    }
    if (isCut) {
      cutPending.set(content.copyId, { sheetId: worksheetId, range: block })
      // only the last few cuts can still be waiting for a paste
      while (cutPending.size > 10) cutPending.delete(cutPending.keys().next().value as string)
    }
    const defaultStyle = worksheet.getDefaultCellStyleInternal()
    return buildClipboardHtml({
      copyId: content.copyId,
      payloadValue: encodePayload(payload),
      cells: htmlCells,
      columnWidths: payload.columnWidths,
      rowHeights: payload.rowHeights,
      ...(defaultStyle?.ff && defaultStyle.fs
        ? { defaultFont: { family: defaultStyle.ff, size: defaultStyle.fs } }
        : {}),
    })
  }

  /// the text a copied cell shows
  function displayText(cell: ClipboardCell): string {
    const field = clipboardField(cell)
    // clipboardField quotes for TSV; the markup wants the raw text
    if (field.startsWith('"') && field.endsWith('"') && field.length >= 2) {
      return field.slice(1, -1).replace(/""/g, '"')
    }
    return field || extractPureTextFromCell(cell)
  }

  async function pasteForeign(
    foreign: ForeignCopy,
    pasteType: string | undefined,
  ): Promise<boolean> {
    const what = hookToWhat[pasteType ?? 'default-paste'] ?? 'all'
    const pasted = await pasteSpecial(
      deps.getContext(),
      runtime,
      { ...DEFAULT_PASTE_SPECIAL, what },
      foreign,
    )
    // Excel's cut-and-paste moves the cells: the window that cut the block clears it now
    if (pasted && foreign.cutCopyId) {
      try {
        window.desktopApi?.notifyClipboardCutDone?.({ copyId: foreign.cutCopyId })
      } catch {
        // the cut stays where it is
      }
    }
    return true
  }

  // Several selected areas: Excel copies them together when they line up (same rows or same
  // columns) and refuses otherwise. Univer copies only the last area.
  service.copy = async (options) => {
    const selections = injector.get(SheetsSelectionsService).getCurrentSelections() ?? []
    if (selections.length < 2) return originalCopy(options)
    const ranges = selections.map((selection) => selection.range as IRange)
    const layout = multiSelectionLayout(ranges)
    if (!layout || options?.copyType === 'CUT') {
      deps.getContext().setMessage(t('appCopyMultiSelection'))
      return false
    }
    const workbook = instances.getCurrentUnitOfType<Workbook>(UniverInstanceType.UNIVER_SHEET)
    const worksheet = workbook?.getActiveSheet()
    if (!workbook || !worksheet) return false
    const unitId = workbook.getUnitId()
    const sheetId = worksheet.getSheetId()
    const content = originalGenerate(unitId, sheetId, layout.bounds, {
      copyHookType: 'default-copy',
    })
    if (!content) return false
    const rows = layout.rows(content.discreteRange.rows)
    const cols = layout.columns(content.discreteRange.cols)
    const combined = {
      ...content,
      // not in Univer's cache: a paste in this window reads the payload like any other window's
      copyId: Math.random().toString(36).slice(2, 8),
      discreteRange: { rows, cols },
    }
    const html = buildCopyHtml(unitId, sheetId, layout.bounds, combined, false)
    if (!html) return false
    await injector
      .get(IClipboardInterfaceService)
      .write(plainTextOfBlock(worksheet, rows, cols), html)
    return true
  }

  service.paste = async (item, pasteType) => {
    try {
      const system = await readSystemClipboard()
      const types = item.types
      const html =
        system?.html ||
        (types.includes('text/html') ? await (await item.getType('text/html')).text() : '')
      const text =
        system?.text ||
        (types.includes('text/plain') ? await (await item.getType('text/plain')).text() : '')
      const visual = extractVisualPayload(html)
      if (visual) return pasteVisualPayload(deps.getContext().visualContext(), visual)
      const foreign = await resolveForeignCopy(html, text, hasLocalCopy)
      if (foreign) return await pasteForeign(foreign, pasteType)
    } catch {
      // anything unreadable is left to Univer's own paste
    }
    return originalPaste(item, pasteType)
  }

  service.legacyPaste = async (html, text, files) => {
    try {
      const visual = extractVisualPayload(html)
      if (visual) return pasteVisualPayload(deps.getContext().visualContext(), visual)
      const foreign = await resolveForeignCopy(html, text, hasLocalCopy)
      if (foreign) return await pasteForeign(foreign, undefined)
    } catch {
      // fall through to Univer
    }
    return originalLegacyPaste(html, text, files)
  }

  /// another window pasted a block this window cut: take it out of the sheet
  const stopCutListener =
    window.desktopApi?.onClipboardCutConsumed?.((info) => {
      const entry = cutPending.get(info.copyId)
      if (!entry) return
      cutPending.delete(info.copyId)
      const workbook = runtime.univerAPI.getActiveWorkbook()
      const sheet = workbook?.getSheetBySheetId(entry.sheetId)
      if (!workbook || !sheet) return
      sheet.activate()
      sheet
        .getRange(
          entry.range.startRow,
          entry.range.startColumn,
          entry.range.endRow - entry.range.startRow + 1,
          entry.range.endColumn - entry.range.startColumn + 1,
        )
        .activate()
      void clearWithExtras(deps.getContext(), runtime, 'all').then(() =>
        service.removeMarkSelection?.(),
      )
    }) ?? (() => undefined)

  return {
    dispose() {
      stopCutListener()
      service.generateCopyContent = originalGenerate
      service.copy = originalCopy
      service.paste = originalPaste
      service.legacyPaste = originalLegacyPaste
      internal._pasteInternal = originalPasteInternal
      internal._getPastedRange = originalGetPastedRange
    },
  }
}
