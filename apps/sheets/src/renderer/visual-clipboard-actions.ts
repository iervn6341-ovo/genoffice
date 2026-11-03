/**
 * Copy / cut / paste of floating objects: the runtime half of visual-clipboard.ts.
 */
import {
  applyChartStateEdit,
  type ChartVisualState,
} from '@genoffice/xlsx-gateway/domain/chart-visual'
import { isMetafileMime, metafileToDataUrl } from '@genoffice/docx-engine/metafile'

import type { WorkbookVisualObject } from '../shared/desktop-api'
import { journalSize } from './edit-journal'
import { t } from './i18n/locale'
import {
  extractVisualPayload,
  visualClipboardHtml,
  visualFromPayload,
  visualToPayload,
  type VisualPayload,
} from './visual-clipboard'
import {
  pushVisualAddUndo,
  queueCtxVisualInstall,
  type VisualActionContext,
} from './visual-actions'
import type { LazyWorkbookState } from './univer-state'
import {
  clampExplicitTo,
  markerFrom,
  markerSpan,
  requestVisualFocus,
  EMU_PER_PIXEL,
} from './WorkbookVisuals'

/// The object as the save would write it now: the file's own visual with this session's anchor
/// and chart edits applied, or the session-added one.
export function effectiveVisual(
  state: LazyWorkbookState,
  visual: WorkbookVisualObject,
): WorkbookVisualObject {
  const added = state.editJournal.visualAdds.find((candidate) => candidate.id === visual.id)
  if (added) return added
  const edit = state.editJournal.visualEdits.get(visual.id)
  const chart =
    visual.chart && visual.chartPath
      ? (applyChartStateEdit(
          visual.chart as ChartVisualState,
          state.editJournal.chartEdits.get(visual.chartPath),
        ) as WorkbookVisualObject['chart'])
      : visual.chart
  return {
    ...visual,
    ...(edit?.anchor ? { anchor: edit.anchor } : {}),
    ...(chart ? { chart } : {}),
  }
}

function base64ToBytes(base64: string): Uint8Array {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
  return bytes
}

async function mediaDataUrlOf(
  state: LazyWorkbookState,
  visual: WorkbookVisualObject,
): Promise<string | undefined> {
  if (visual.mediaDataUrl) return visual.mediaDataUrl
  if (visual.kind !== 'image') return undefined
  try {
    const media = await window.desktopApi.readWorkbookMedia({
      sessionId: state.file.sessionId,
      visualId: visual.id,
    })
    return isMetafileMime(media.mediaType)
      ? ((await metafileToDataUrl(base64ToBytes(media.base64), media.mediaType)) ?? undefined)
      : `data:${media.mediaType};base64,${media.base64}`
  } catch {
    return undefined
  }
}

async function pngBlobOf(dataUrl: string): Promise<Blob | null> {
  try {
    const comma = dataUrl.indexOf(',')
    if (comma === -1) return null
    const mimeMatch = /^data:([^;]+)/.exec(dataUrl)
    const mime = (mimeMatch && mimeMatch[1]) ? mimeMatch[1] : 'application/octet-stream'
    const bytes = base64ToBytes(dataUrl.slice(comma + 1))
    const arrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
    const blob = new Blob([arrayBuffer as ArrayBuffer], { type: mime })
    if (blob.type === 'image/png') return blob
    const bitmap = await createImageBitmap(blob)
    const canvas = document.createElement('canvas')
    canvas.width = bitmap.width
    canvas.height = bitmap.height
    canvas.getContext('2d')?.drawImage(bitmap, 0, 0)
    return await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'))
  } catch {
    return null
  }
}

/// Puts the object on the clipboard (and removes it after a Cut). Returns whether it did.
export async function copyVisualObject(
  ctx: VisualActionContext,
  visual: WorkbookVisualObject,
  cut: boolean,
): Promise<boolean> {
  const state = ctx.lazyWorkbookRef.current
  const runtime = ctx.univerRef.current
  if (!state || !runtime) return false
  const effective = effectiveVisual(state, visual)
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getSheetBySheetId(effective.sheetId)
  const sheetName = worksheet?.getSheetName() ?? ''
  const dataUrl = await mediaDataUrlOf(state, effective)

  const columnWidth = (index: number): number => Math.max(worksheet?.getColumnWidth(index) ?? 72, 1)
  const rowHeight = (index: number): number => Math.max(worksheet?.getRowHeight(index) ?? 20, 1)
  const marginX = Math.max(0, effective.anchor.fromColumnOffset / EMU_PER_PIXEL)
  const marginY = Math.max(0, effective.anchor.fromRowOffset / EMU_PER_PIXEL)
  const explicitTo = effective.anchor.explicitTo === true
  const pixelWidth = effective.frameWidth
    ? Math.round(effective.frameWidth / EMU_PER_PIXEL)
    : markerSpan(
        { index: effective.anchor.fromColumn, offset: marginX },
        clampExplicitTo(
          markerFrom(effective.anchor.toColumn, effective.anchor.toColumnOffset),
          explicitTo,
          columnWidth,
        ),
        columnWidth,
      )
  const pixelHeight = effective.frameHeight
    ? Math.round(effective.frameHeight / EMU_PER_PIXEL)
    : markerSpan(
        { index: effective.anchor.fromRow, offset: marginY },
        clampExplicitTo(
          markerFrom(effective.anchor.toRow, effective.anchor.toRowOffset),
          explicitTo,
          rowHeight,
        ),
        rowHeight,
      )

  const payload = visualToPayload(
    effective,
    sheetName,
    state.file.sessionId,
    dataUrl,
    pixelWidth > 0 ? pixelWidth : undefined,
    pixelHeight > 0 ? pixelHeight : undefined,
  )
  if (!payload) {
    ctx.setMessage(t('appVisualCannotCopy'))
    return false
  }
  const label = effective.name ?? effective.kind
  const items: Record<string, Blob> = {
    'text/html': new Blob([visualClipboardHtml(payload, label)], { type: 'text/html' }),
    'text/plain': new Blob([label], { type: 'text/plain' }),
  }
  if (effective.kind === 'image' && dataUrl) {
    const png = await pngBlobOf(dataUrl)
    if (png) items['image/png'] = png
  }
  try {
    await navigator.clipboard.write([new ClipboardItem(items)])
  } catch (err) {
    if (items['image/png']) {
      try {
        delete items['image/png']
        await navigator.clipboard.write([new ClipboardItem(items)])
      } catch (retryErr) {
        console.error('[Sheets/Visual] Clipboard write retry failed:', retryErr)
        ctx.setMessage(t('appVisualCannotCopy'))
        return false
      }
    } else {
      console.error('[Sheets/Visual] Clipboard write failed:', err)
      ctx.setMessage(t('appVisualCannotCopy'))
      return false
    }
  }
  if (cut) ctx.shapeEditRef.current(visual.id, { remove: true })
  return true
}

/// Pastes a copied object at the active cell of the active sheet. Returns whether it did.
export function pasteVisualPayload(ctx: VisualActionContext, payload: VisualPayload): boolean {
  const runtime = ctx.univerRef.current
  const state = ctx.lazyWorkbookRef.current
  if (!runtime || !state) {
    ctx.setMessage(t('appShapeNeedsFile'))
    return false
  }
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sheet = workbook?.getActiveSheet()
  if (!workbook || !sheet) return false
  const active = workbook.getActiveRange()

  const columnWidth = (index: number): number => Math.max(sheet.getColumnWidth(index), 1)
  const rowHeight = (index: number): number => Math.max(sheet.getRowHeight(index), 1)
  const maxColumns = sheet.getMaxColumns()
  const maxRows = sheet.getMaxRows()

  const visual = visualFromPayload(
    // a chart's cell references name a sheet of the workbook it was copied from
    {
      ...payload,
      visual: payload.bookId === state.file.sessionId ? payload.visual : withoutReferences(payload),
    },
    {
      id: `added-${payload.visual.kind}-${Date.now().toString(36)}-${state.editJournal.visualAdds.length + 1}`,
      sheetId: sheet.getSheetId(),
      row: active?.getRow() ?? 0,
      column: active?.getColumn() ?? 0,
    },
    { columnWidth, rowHeight, maxColumns, maxRows },
  )
  pushVisualAddUndo(ctx, runtime, state, visual)
  requestVisualFocus(visual.id)
  ctx.setPendingEdits(journalSize(state.editJournal))
  queueCtxVisualInstall(ctx, runtime)
  return true
}

/// A chart pasted into another workbook keeps its numbers but not the references to the cells it
/// was drawn from.
function withoutReferences(payload: VisualPayload): VisualPayload['visual'] {
  const chart = payload.visual.chart
  if (!chart) return payload.visual
  return {
    ...payload.visual,
    chart: {
      ...chart,
      series: chart.series.map((series) => {
        const { valuesRef: _v, categoriesRef: _c, nameRef: _n, ...rest } = series
        return rest
      }),
    },
  }
}

export { extractVisualPayload }
