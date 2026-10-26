/**
 * Entering cell edit (double-click / F2 / typing) must leave the text exactly
 * where the grid canvas drew it — Excel overlays the in-cell editor so only
 * the caret is new. Stock Univer shifts it, because the editor is a separate
 * document layout:
 *
 * 1. The editor document keeps the paragraph terminator as a visible glyph
 *    (a space-wide advance). Centered text therefore sits half that width
 *    left of the grid position and right-aligned text a full width left
 *    (about 2px / 4px at 11pt, growing with the font size and zoom) — the
 *    visible "jump" on double-click. The grid's rich-text path already marks
 *    its documents `zeroWidthParagraphBreak`; the editor document of a plain
 *    cell must do the same.
 * 2. Cell paddings are partial (an indent is stored as `pd: { l }` only).
 *    The grid fills the missing sides from the defaults, the editor document
 *    did not (undefined margins), and the editor's page width ignores the
 *    left/right padding altogether — so an indented cell's text wrapped onto
 *    a second line the moment editing began. The page must be at least the
 *    text plus BOTH margins (stock Univer only reserves 5px, enough for the
 *    default 2px + 2px paddings and nothing more).
 *
 * 3. The grid canvas paints cell content translated by half a scene pixel
 *    (FIX_ONE_PIXEL_BLUR_OFFSET, so 1px lines stay crisp); the editor DOM
 *    overlay sits on the exact cell origin. Left unshifted the text steps
 *    up-left by 0.5 x zoom px (~0.5px at 100%, ~0.7px at 140%, 1px at
 *    200%) when editing starts. Vertically centered text additionally sits
 *    MIDDLE_CELL_POS_MAGIC_NUMBER (1) unit lower in the editor.
 * 4. `fitTextSize` rounds the editor width UP to a whole pixel, which moves
 *    right-aligned text left by the fraction (up to 1px) once the editor has
 *    to grow past the cell; the widths here are exact.
 *
 * All patches live on the class prototypes because the injector hands out
 * lazy proxies: assigning onto a resolved instance would only shadow them.
 */
import {
  DOCS_NORMAL_EDITOR_UNIT_ID_KEY,
  HorizontalAlign,
  UniverInstanceType,
  VerticalAlign,
  WrapStrategy,
} from '@univerjs/core'
import {
  DEFAULT_PADDING_DATA,
  FIX_ONE_PIXEL_BLUR_OFFSET,
  MIDDLE_CELL_POS_MAGIC_NUMBER,
  convertTextRotation,
} from '@univerjs/engine-render'
import {
  CellEditorManagerService,
  EditorBridgeService,
  SheetCellEditorResizeService,
} from '@univerjs/sheets-ui'

export interface PaddingData {
  t?: number
  b?: number
  l?: number
  r?: number
}

/** Univer's per-side padding defaults for the sides a cell style leaves out. */
export function fullPadding(
  pd: PaddingData | null | undefined,
  defaults: Required<PaddingData> = DEFAULT_PADDING_DATA as Required<PaddingData>,
): Required<PaddingData> {
  return {
    t: pd?.t ?? defaults.t,
    b: pd?.b ?? defaults.b,
    l: pd?.l ?? defaults.l,
    r: pd?.r ?? defaults.r,
  }
}

/**
 * Slack past the text and both margins so float noise never wraps the line.
 * (Stock Univer reserves a full unit, which also pushes right-aligned text a
 * unit off the cell edge — the caret fits inside the right padding anyway.)
 */
export const EDITOR_CARET_ROOM = 0.05

/**
 * Editor document page width (px, in the editor document's own unscaled
 * space ×scale) for a non-wrapping cell: the cell width, widened when the
 * text plus both paddings and the caret room do not fit.
 */
export function editorPageWidth(
  cellWidth: number,
  textWidth: number,
  padding: Required<PaddingData>,
  scaleX: number,
): number {
  return Math.max(cellWidth, textWidth + (padding.l + padding.r + EDITOR_CARET_ROOM) * scaleX)
}

export interface EditorOverlayRect {
  startX?: number
  startY?: number
  endX?: number
  endY?: number
  show?: boolean
}

/**
 * How far (css px) the editor overlay must move to land on the text the grid
 * painted: the grid's half-pixel content translation in the current zoom, and
 * for vertically centered cells Univer's one-unit centering constant.
 */
export function editorOverlayShift(
  scaleX: number,
  scaleY: number,
  verticalAlign: number | undefined,
): { dx: number; dy: number } {
  const dy =
    verticalAlign === VerticalAlign.MIDDLE
      ? (FIX_ONE_PIXEL_BLUR_OFFSET - MIDDLE_CELL_POS_MAGIC_NUMBER) * scaleY
      : FIX_ONE_PIXEL_BLUR_OFFSET * scaleY
  return { dx: FIX_ONE_PIXEL_BLUR_OFFSET * scaleX, dy }
}

export function shiftEditorOverlay<T extends EditorOverlayRect>(
  rect: T,
  shift: { dx: number; dy: number },
): T {
  const next = { ...rect }
  if (next.startX != null) next.startX += shift.dx
  if (next.endX != null) next.endX += shift.dx
  if (next.startY != null) next.startY += shift.dy
  if (next.endY != null) next.endY += shift.dy
  return next
}

export interface EditorFitInput {
  /** cell size on screen (merged range included), css px */
  cellWidth: number
  cellHeight: number
  /** laid-out text size on screen (already x scale), css px */
  textWidth: number
  textHeight: number
  padding: Required<PaddingData>
  scaleX: number
  scaleY: number
  horizontalAlign: number
  verticalAlign: number
}

export interface EditorFitResult {
  /** editor box size on screen, css px */
  width: number
  height: number
  /** document margins (document units) that place the text inside the box */
  offsetTop: number
  offsetLeft: number
}

/**
 * Editor box and document margins — Univer's `fitTextSize` arithmetic with the
 * width kept exact (see header, point 4).
 */
export function computeEditorFit(input: EditorFitInput): EditorFitResult {
  const { padding, scaleX, scaleY, horizontalAlign, verticalAlign } = input
  const actualWidth = input.textWidth + (padding.l + padding.r) * scaleX
  const actualHeight = input.textHeight + (padding.t + padding.b) * scaleY
  const width = Math.max(input.cellWidth, actualWidth)
  const height = input.cellHeight < actualHeight ? Math.ceil(actualHeight) : input.cellHeight
  let offsetTop: number
  if (verticalAlign === VerticalAlign.MIDDLE) offsetTop = (height - actualHeight) / 2 / scaleY
  else if (verticalAlign === VerticalAlign.TOP) offsetTop = padding.t
  else offsetTop = (height - actualHeight) / scaleY
  let offsetLeft: number
  if (horizontalAlign === HorizontalAlign.CENTER) offsetLeft = 0
  else if (horizontalAlign === HorizontalAlign.RIGHT) offsetLeft = (width - actualWidth) / scaleX
  else offsetLeft = padding.l
  return {
    width,
    height,
    offsetTop: Math.max(offsetTop, padding.t),
    offsetLeft: Math.max(offsetLeft, padding.l),
  }
}

interface DocStyleLike {
  marginTop?: number
  marginBottom?: number
  marginLeft?: number
  marginRight?: number
  renderConfig?: Record<string, unknown>
}

interface LayoutObjectLike {
  paddingData?: PaddingData
  documentModel?: { documentStyle?: DocStyleLike } | null
}

/** Normalise the editor's document layout object in place (see the header). */
export function normalizeEditorLayoutObject(layout: LayoutObjectLike | null | undefined): void {
  if (!layout) return
  const padding = fullPadding(layout.paddingData)
  layout.paddingData = padding
  const style = layout.documentModel?.documentStyle
  if (!style) return
  style.marginTop = padding.t
  style.marginBottom = padding.b
  style.marginLeft = padding.l
  style.marginRight = padding.r
  style.renderConfig = { ...style.renderConfig, zeroWidthParagraphBreak: 1 }
}

interface BridgeProto {
  getLatestEditCellState(): { documentLayoutObject?: LayoutObjectLike } | undefined
}

interface ManagerProto {
  setState(param: EditorOverlayRect): void
}

interface ResizeProto {
  _editAreaProcessing(...args: unknown[]): unknown
  fitTextSize(callback?: () => void): void
  _getEditorSkeleton(): { calculate(): void } | null | undefined
  _editorBridgeService: {
    getEditCellState(): {
      position: { startX: number; startY: number; endX: number; endY: number }
      documentLayoutObject: {
        verticalAlign: number
        horizontalAlign: number
        paddingData: PaddingData
        fill?: string
        textRotation?: unknown
        wrapStrategy?: number
      }
      canvasOffset: unknown
      scaleX: number
      scaleY: number
    } | null
  }
  _predictingSize(
    position: { startX: number; endX: number },
    canvasOffset: unknown,
    documentSkeleton: unknown,
    documentLayoutObject: {
      textRotation?: unknown
      wrapStrategy?: number
      paddingData?: PaddingData
    },
    scaleX?: number,
    scaleY?: number,
  ): { actualWidth: number; actualHeight: number } | undefined
  _univerInstanceService: {
    getUnit(
      id: string,
      type: number,
    ):
      | {
          updateDocumentDataPageSize(width: number): void
          updateDocumentDataMargin(margin: { t: number; l: number }): void
        }
      | null
      | undefined
  }
}

let installed = false

export function installCellEditorOverlayFix(): void {
  if (installed) return
  installed = true

  const bridge = EditorBridgeService.prototype as unknown as BridgeProto
  const originalLatest = bridge.getLatestEditCellState
  if (typeof originalLatest === 'function') {
    bridge.getLatestEditCellState = function (this: BridgeProto) {
      const state = originalLatest.call(this)
      normalizeEditorLayoutObject(state?.documentLayoutObject)
      return state
    }
  }

  // _editAreaProcessing is the one caller that positions the overlay (scaleX /
  // scaleY are its 6th / 7th arguments); shift exactly the state it publishes.
  let pendingShift: { dx: number; dy: number } | null = null
  const manager = CellEditorManagerService.prototype as unknown as ManagerProto
  const originalSetState = manager.setState
  if (typeof originalSetState === 'function') {
    manager.setState = function (this: ManagerProto, param: EditorOverlayRect) {
      const shift = pendingShift
      pendingShift = null
      return originalSetState.call(
        this,
        shift && param.startX != null ? shiftEditorOverlay(param, shift) : param,
      )
    }
  }

  const resize = SheetCellEditorResizeService.prototype as unknown as ResizeProto
  const originalEditArea = resize._editAreaProcessing
  if (typeof originalEditArea === 'function' && typeof originalSetState === 'function') {
    resize._editAreaProcessing = function (this: ResizeProto, ...args: unknown[]) {
      pendingShift = editorOverlayShift(
        (args[5] as number) ?? 1,
        (args[6] as number) ?? 1,
        this._editorBridgeService.getEditCellState()?.documentLayoutObject.verticalAlign,
      )
      try {
        return originalEditArea.apply(this, args)
      } finally {
        pendingShift = null
      }
    }
  }

  if (typeof resize.fitTextSize === 'function') {
    resize.fitTextSize = function (this: ResizeProto, callback?: () => void) {
      const param = this._editorBridgeService.getEditCellState()
      if (!param) return
      const { position, documentLayoutObject, canvasOffset, scaleX, scaleY } = param
      const doc = this._univerInstanceService.getUnit(
        DOCS_NORMAL_EDITOR_UNIT_ID_KEY,
        UniverInstanceType.UNIVER_DOC,
      )
      if (!doc) return
      const skeleton = this._getEditorSkeleton()
      if (!skeleton) return
      const info = this._predictingSize(
        position,
        canvasOffset,
        skeleton,
        documentLayoutObject,
        scaleX,
        scaleY,
      )
      if (!info || info.actualWidth <= 0) return
      const fit = computeEditorFit({
        cellWidth: position.endX - position.startX,
        cellHeight: position.endY - position.startY,
        textWidth: info.actualWidth,
        textHeight: info.actualHeight,
        padding: fullPadding(documentLayoutObject.paddingData),
        scaleX,
        scaleY,
        horizontalAlign: documentLayoutObject.horizontalAlign,
        verticalAlign: documentLayoutObject.verticalAlign,
      })
      doc.updateDocumentDataMargin({ t: fit.offsetTop, l: fit.offsetLeft })
      skeleton.calculate()
      this._editAreaProcessing(
        fit.width,
        fit.height,
        position,
        canvasOffset,
        documentLayoutObject.fill,
        scaleX,
        scaleY,
        documentLayoutObject.horizontalAlign,
        callback,
      )
    }
  }

  const originalPredict = resize._predictingSize
  if (typeof originalPredict === 'function') {
    resize._predictingSize = function (
      this: ResizeProto,
      position,
      canvasOffset,
      documentSkeleton,
      documentLayoutObject,
      scaleX = 1,
      scaleY = 1,
    ) {
      const info = originalPredict.call(
        this,
        position,
        canvasOffset,
        documentSkeleton,
        documentLayoutObject,
        scaleX,
        scaleY,
      )
      if (!info) return info
      const { vertexAngle } = convertTextRotation(
        documentLayoutObject.textRotation as Parameters<typeof convertTextRotation>[0],
      )
      if (documentLayoutObject.wrapStrategy === WrapStrategy.WRAP && vertexAngle === 0) return info
      const doc = this._univerInstanceService.getUnit(
        DOCS_NORMAL_EDITOR_UNIT_ID_KEY,
        UniverInstanceType.UNIVER_DOC,
      )
      if (!doc) return info
      const width = editorPageWidth(
        position.endX - position.startX,
        info.actualWidth,
        fullPadding(documentLayoutObject.paddingData),
        scaleX,
      )
      doc.updateDocumentDataPageSize(width / scaleX)
      return info
    }
  }
}
