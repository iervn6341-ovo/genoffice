import { IUniverInstanceService, UniverInstanceType, type DocumentDataModel } from '@univerjs/core'
import { DocSelectionManagerService } from '@univerjs/docs'
import { IEditorBridgeService } from '@univerjs/sheets-ui'

import { indexedFormulaText } from './formula-view'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'

export interface FormulaSymbolDraft {
  readonly workbookId: string
  readonly sheetId: string
  readonly row: number
  readonly column: number
  readonly source: string
  readonly selectionStart: number
  readonly selectionEnd: number
}

export function insertIntoFormulaSource(
  source: string,
  start: number,
  end: number,
  char: string,
): { source: string; caret: number } {
  const from = Math.max(0, Math.min(source.length, Math.min(start, end)))
  const to = Math.max(from, Math.min(source.length, Math.max(start, end)))
  return {
    source: `${source.slice(0, from)}${char}${source.slice(to)}`,
    caret: from + char.length,
  }
}

export function formulaSymbolSourceProblem(
  source: string,
): 'start' | 'parentheses' | 'length' | null {
  if (!source.startsWith('=')) return 'start'
  if (source.length > 8_192) return 'length'
  let quoted = false
  let depth = 0
  for (let index = 1; index < source.length; index += 1) {
    const char = source[index]
    if (char === '"') {
      if (quoted && source[index + 1] === '"') index += 1
      else quoted = !quoted
    } else if (!quoted && char === '(') depth += 1
    else if (!quoted && char === ')') {
      depth -= 1
      if (depth < 0) return 'parentheses'
    }
  }
  return depth === 0 && !quoted ? null : 'parentheses'
}

/// Capture the live document before a ribbon click can blur and commit the
/// cell editor. The normal facade formula also covers shared-formula cells;
/// the harvested index covers formulas outside Univer's resident window.
export function captureFormulaSymbolDraft(
  runtime: UniverRuntime | null,
  lazy: LazyWorkbookState | null,
): { draft: FormulaSymbolDraft; wasEditing: boolean } | null {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!runtime || !workbook || !sheet || !range) return null
  const row = range.getRow()
  const column = range.getColumn()
  const sheetId = sheet.getSheetId()
  const cell = sheet.getRange(row, column, 1, 1)
  const saved =
    cell.getFormula() ||
    indexedFormulaText(lazy, sheetId, row, column) ||
    cell.getCellDatas()[0]?.[0]?.f ||
    ''
  const injector = runtime.univer.__getInjector()
  const bridge = injector.get(IEditorBridgeService)
  const location = bridge.getEditCellState()
  const editing =
    bridge.isVisible().visible &&
    location?.sheetId === sheetId &&
    location.row === row &&
    location.column === column
  let source = saved
  let selectionStart = saved.length
  let selectionEnd = saved.length
  if (editing) {
    const editorId = bridge.getCurrentEditorId()
    const doc = injector
      .get(IUniverInstanceService)
      .getUnit<DocumentDataModel>(editorId, UniverInstanceType.UNIVER_DOC)
    const live = doc?.getBody()?.dataStream?.replace(/\r\n$/, '') ?? ''
    // A live plain-text edit wins over an older formula in the cell model.
    // The caller commits that editor before opening the ordinary symbol picker.
    if (!live.startsWith('=')) return null
    source = live
    const selections = injector.get(DocSelectionManagerService)
    // Univer keeps document selections in either the editor unit's own
    // sub-unit or the default '' sub-unit, depending on the editing surface.
    // Its current selection identifies the focused document; prefer that
    // over a stale range left in another sub-unit by a previous edit.
    const current = selections.__getCurrentSelection()
    const currentRanges = current?.unitId === editorId ? selections.getTextRanges() : null
    const ownRanges = selections.getTextRanges({ unitId: editorId, subUnitId: editorId })
    const selected = currentRanges?.length
      ? currentRanges
      : ownRanges?.length
        ? ownRanges
        : selections.getTextRanges({ unitId: editorId, subUnitId: '' })
    const active = selected?.find((range) => range.isActive) ?? selected?.[0]
    selectionStart = active?.startOffset ?? live.length
    selectionEnd = active?.endOffset ?? live.length
  }
  if (!source.startsWith('=')) return null
  return {
    draft: {
      workbookId: workbook.getId(),
      sheetId,
      row,
      column,
      source,
      selectionStart,
      selectionEnd,
    },
    wasEditing: editing,
  }
}

/// The caret / selection of the live cell editor for the active cell (offsets into its document
/// stream), or null while the cell is not being edited. Captured before a ribbon click can blur the
/// editor so Insert Symbol lands at the caret instead of the end of the text.
export function captureEditorCaret(
  runtime: UniverRuntime | null,
): { start: number; end: number } | null {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  const range = workbook?.getActiveRange()
  if (!runtime || !workbook || !sheet || !range) return null
  try {
    return readEditorCaret(runtime, sheet.getSheetId(), range.getRow(), range.getColumn())
  } catch {
    // no live editor service (or it is mid-dispose): the symbol is simply appended
    return null
  }
}

function readEditorCaret(
  runtime: UniverRuntime,
  sheetId: string,
  row: number,
  column: number,
): { start: number; end: number } | null {
  const injector = runtime.univer.__getInjector()
  const bridge = injector.get(IEditorBridgeService)
  const location = bridge.getEditCellState()
  if (
    !bridge.isVisible().visible ||
    location?.sheetId !== sheetId ||
    location.row !== row ||
    location.column !== column
  )
    return null
  const editorId = bridge.getCurrentEditorId()
  const selections = injector.get(DocSelectionManagerService)
  const current = selections.__getCurrentSelection()
  const currentRanges = current?.unitId === editorId ? selections.getTextRanges() : null
  const ownRanges = selections.getTextRanges({ unitId: editorId, subUnitId: editorId })
  const selected = currentRanges?.length
    ? currentRanges
    : ownRanges?.length
      ? ownRanges
      : selections.getTextRanges({ unitId: editorId, subUnitId: '' })
  const active = selected?.find((candidate) => candidate.isActive) ?? selected?.[0]
  return active ? { start: active.startOffset, end: active.endOffset } : null
}

/// Read-only instrumentation for the opt-in built-app E2E hook. The live
/// editor can retain several selection maps; expose each one so a focus/keys
/// failure is distinguishable from a lookup failure without mutating either.
export function formulaSymbolSelectionDiagnostic(
  runtime: UniverRuntime | null,
  lazy: LazyWorkbookState | null,
) {
  if (!runtime) return null
  const injector = runtime.univer.__getInjector()
  const bridge = injector.get(IEditorBridgeService)
  const selections = injector.get(DocSelectionManagerService)
  const editorId = bridge.getCurrentEditorId()
  const current = selections.__getCurrentSelection()
  const ranges = (items: ReturnType<typeof selections.getTextRanges>) =>
    items?.map((item) => ({
      startOffset: item.startOffset,
      endOffset: item.endOffset,
      isActive: item.isActive ?? false,
    })) ?? null
  const activeElement = document.activeElement
  return {
    editorId,
    liveDraft: captureFormulaSymbolDraft(runtime, lazy)?.draft ?? null,
    currentSelection: current ? { unitId: current.unitId, subUnitId: current.subUnitId } : null,
    activeRanges: ranges(selections.getTextRanges()),
    ownRanges: ranges(selections.getTextRanges({ unitId: editorId, subUnitId: editorId })),
    defaultRanges: ranges(selections.getTextRanges({ unitId: editorId, subUnitId: '' })),
    activeElement: activeElement
      ? {
          tag: activeElement.tagName,
          className: typeof activeElement.className === 'string' ? activeElement.className : '',
        }
      : null,
  }
}

export function applyFormulaSymbolDraft(
  runtime: UniverRuntime | null,
  draft: FormulaSymbolDraft,
  source: string,
): boolean {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook || workbook.getId() !== draft.workbookId) return false
  const sheet = workbook.getSheetBySheetId(draft.sheetId)
  if (!sheet || formulaSymbolSourceProblem(source)) return false
  // The modal's destination is fixed even if a background selection changes.
  // The facade passes an ICellData object through unchanged. Explicit nulls
  // clear any prior rich document, cached result, or shared-formula group.
  sheet.getRange(draft.row, draft.column, 1, 1).setValue({ f: source, v: null, p: null, si: null })
  return true
}
