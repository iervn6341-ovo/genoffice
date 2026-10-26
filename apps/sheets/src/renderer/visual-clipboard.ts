/**
 * Copying floating objects (charts, pictures, shapes, equations) between workbook windows.
 *
 * The object travels as JSON inside the clipboard HTML (`data-genoffice-visual`), so another
 * GenOffice window rebuilds it as a session-added object at the active cell. A picture also puts
 * its PNG on the clipboard for other programs. File-bound identity (drawing part, chart part,
 * media path, ids) is dropped: the copy is a new object of the pasting workbook.
 */
import type { WorkbookVisualObject } from '../shared/desktop-api'

export const VISUAL_ATTR = 'data-genoffice-visual'
export const VISUAL_PAYLOAD_VERSION = 1

type Anchor = WorkbookVisualObject['anchor']

export interface VisualPayload {
  readonly version: number
  /// the object without its place in a file
  readonly visual: Omit<WorkbookVisualObject, 'id' | 'sheetId' | 'anchor'>
  /// the frame's size in cells, and the offsets inside its first and last cell
  readonly spanRows: number
  readonly spanColumns: number
  readonly anchor: Pick<
    Anchor,
    'fromRowOffset' | 'fromColumnOffset' | 'toRowOffset' | 'toColumnOffset'
  > & { readonly explicitTo?: boolean }
  /// the sheet the object was copied from (charts keep references to it)
  readonly sheetName: string
  /// the workbook session it was copied from (references only stay valid inside it)
  readonly bookId: string
}

/// Kinds a pasted copy can stand for: a read-only embed has no data to carry
const COPYABLE_KINDS = new Set(['chart', 'image', 'shape', 'equation'])

export function isCopyableVisual(visual: WorkbookVisualObject): boolean {
  return COPYABLE_KINDS.has(visual.kind)
}

/// The payload for a visual; `mediaDataUrl` supplies a picture's bytes when the object itself
/// holds only a media path.
export function visualToPayload(
  visual: WorkbookVisualObject,
  sheetName: string,
  bookId: string,
  mediaDataUrl?: string,
): VisualPayload | null {
  if (!isCopyableVisual(visual)) return null
  const {
    id: _id,
    sheetId: _sheetId,
    anchor,
    chartPath: _chartPath,
    mediaPath: _mediaPath,
    fillMediaPath: _fillMediaPath,
    fillMediaType: _fillMediaType,
    drawingPath: _drawingPath,
    drawingIndex: _drawingIndex,
    ...rest
  } = visual
  const carried: Omit<WorkbookVisualObject, 'id' | 'sheetId' | 'anchor'> = {
    ...rest,
    ...(mediaDataUrl ? { mediaDataUrl } : {}),
  }
  if (visual.kind === 'image' && !carried.mediaDataUrl) return null
  return {
    version: VISUAL_PAYLOAD_VERSION,
    visual: carried,
    spanRows: Math.max(1, anchor.toRow - anchor.fromRow),
    spanColumns: Math.max(1, anchor.toColumn - anchor.fromColumn),
    anchor: {
      fromRowOffset: anchor.fromRowOffset,
      fromColumnOffset: anchor.fromColumnOffset,
      toRowOffset: anchor.toRowOffset,
      toColumnOffset: anchor.toColumnOffset,
      ...(anchor.explicitTo === undefined ? {} : { explicitTo: anchor.explicitTo }),
    },
    sheetName,
    bookId,
  }
}

/// A pasted copy at `row` / `column` of `sheetId`, as a session-added object.
export function visualFromPayload(
  payload: VisualPayload,
  place: { id: string; sheetId: string; row: number; column: number },
): WorkbookVisualObject {
  const a = payload.anchor
  return {
    ...payload.visual,
    id: place.id,
    sheetId: place.sheetId,
    anchor: {
      fromRow: place.row,
      fromColumn: place.column,
      fromRowOffset: a.fromRowOffset,
      fromColumnOffset: a.fromColumnOffset,
      toRow: place.row + payload.spanRows,
      toColumn: place.column + payload.spanColumns,
      toRowOffset: a.toRowOffset,
      toColumnOffset: a.toColumnOffset,
      ...(a.explicitTo === undefined ? {} : { explicitTo: a.explicitTo }),
    },
  } as WorkbookVisualObject
}

function toBase64(text: string): string {
  const bytes = new TextEncoder().encode(text)
  let binary = ''
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  }
  return btoa(binary)
}

function fromBase64(encoded: string): string {
  const binary = atob(encoded)
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i)
  return new TextDecoder().decode(bytes)
}

export function encodeVisualPayload(payload: VisualPayload): string {
  return toBase64(JSON.stringify(payload))
}

export function extractVisualPayload(html: string | null | undefined): VisualPayload | null {
  if (!html) return null
  const match = new RegExp(`${VISUAL_ATTR}="([A-Za-z0-9+/=]+)"`).exec(html)
  if (!match) return null
  try {
    const raw = JSON.parse(fromBase64(match[1]!)) as Partial<VisualPayload> | null
    if (!raw || raw.version !== VISUAL_PAYLOAD_VERSION || !raw.visual || !raw.anchor) return null
    if (!COPYABLE_KINDS.has(String(raw.visual.kind))) return null
    if (typeof raw.spanRows !== 'number' || typeof raw.spanColumns !== 'number') return null
    return raw as VisualPayload
  } catch {
    return null
  }
}

/// The clipboard markup of a copied object: a labelled placeholder other programs show as text.
export function visualClipboardHtml(payload: VisualPayload, label: string): string {
  const safe = label.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  return `<span ${VISUAL_ATTR}="${encodeVisualPayload(payload)}">${safe}</span>`
}
