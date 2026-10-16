/// Accounting underlines for the two cell kinds Univer draws through its document renderer:
/// rich-text cells (`p` documents) and rotated text. The plain-cell hook in
/// `accounting-underline.ts` cannot reach them, and a hook on the private document-line path only
/// produced 3–4 px strokes (clipped to the glyph run). Instead of touching the renderer, this
/// layer anchors a tiny SVG to the cell through the public float-DOM API, so the rule scrolls,
/// zooms and clips with the grid and is derived from the same marker (`custom`) that the saved
/// `<u val="…Accounting">` produces. Native Univer underline of the runs stays drawn underneath;
/// the saved format is never changed by this layer.

import type { IStyleData } from '@univerjs/core'

import { ACCOUNTING_UNDERLINE_KEY, type AccountingUnderline } from './accounting-underline'
import type { createUniver } from './create-univer'

type UniverRuntime = ReturnType<typeof createUniver>

export interface AccountingOverlayInput {
  readonly width: number
  readonly height: number
  readonly mode: AccountingUnderline
  /** Rotation in degrees, counter-clockwise positive (Excel textRotation 0–90 / 91–180 mapped). */
  readonly angle: number
  /** Measured text length in px (only used for rotated cells). */
  readonly textLength: number
  readonly horizontal: 'left' | 'center' | 'right'
  readonly vertical: 'top' | 'middle' | 'bottom'
}

export interface OverlayStroke {
  readonly x1: number
  readonly y1: number
  readonly x2: number
  readonly y2: number
}

const INSET = 1
const DOUBLE_GAP = 2

/**
 * Stroke geometry inside the cell box. Horizontal (rich) cells get the full-width rule that the
 * Excel Mac capture shows (E1: nearly the whole cell); rotated cells get a rule that follows the
 * rotated baseline (F1 at 45°: diagonal, as long as the text, clipped by the cell).
 */
export function accountingOverlayStrokes(input: AccountingOverlayInput): OverlayStroke[] {
  const { width, height } = input
  if (!(width > 2) || !(height > 2)) return []
  const count = input.mode === 'doubleAccounting' ? 2 : 1
  const out: OverlayStroke[] = []
  if (!input.angle) {
    for (let index = 0; index < count; index += 1) {
      const y = height - INSET - index * (DOUBLE_GAP + 1)
      out.push({ x1: INSET, y1: y, x2: width - INSET, y2: y })
    }
    return out
  }
  const radians = (input.angle * Math.PI) / 180
  const cos = Math.cos(radians)
  const sin = Math.sin(radians)
  const length = Math.max(4, input.textLength)
  // the baseline starts where the rotated text block starts for the cell's alignment
  const spanX = Math.abs(cos) * length
  const spanY = Math.abs(sin) * length
  const startX =
    input.horizontal === 'right'
      ? width - INSET - spanX
      : input.horizontal === 'center'
        ? (width - spanX) / 2
        : INSET
  const startY =
    input.vertical === 'top'
      ? INSET + spanY
      : input.vertical === 'middle'
        ? (height + spanY) / 2
        : height - INSET
  const directionX = cos >= 0 ? 1 : -1
  const directionY = sin >= 0 ? -1 : 1
  for (let index = 0; index < count; index += 1) {
    const offset = index * (DOUBLE_GAP + 1)
    const x1 = startX + (directionX < 0 ? spanX : 0)
    const y1 = startY + offset
    out.push({ x1, y1, x2: x1 + directionX * spanX, y2: y1 + directionY * spanY })
  }
  return out
}

/** Univer stores rotation as `{ a: degrees (−90..90), v: vertical-stack flag }`. */
export function overlayAngleOf(style: IStyleData | undefined): number {
  const rotation = style?.tr
  if (!rotation || rotation.v) return 0
  return Number.isFinite(rotation.a) ? rotation.a : 0
}

const HORIZONTAL = { 1: 'left', 2: 'center', 3: 'right' } as const
const VERTICAL = { 1: 'top', 2: 'middle', 3: 'bottom' } as const

export interface AccountingOverlayHandle {
  refresh(): void
  dispose(): void
}

const MAX_OVERLAYS = 200

function measure(text: string, style: IStyleData | undefined): number {
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  if (!context) return text.length * 7
  const size = style?.fs ?? 11
  context.font = `${style?.bl ? 'bold ' : ''}${style?.it ? 'italic ' : ''}${size}pt ${style?.ff ?? 'sans-serif'}`
  return context.measureText(text).width
}

/** Installs/refreshes float layers for the active sheet's rich or rotated accounting cells. */
export function installAccountingUnderlineOverlay(runtime: UniverRuntime): AccountingOverlayHandle {
  let disposables: { dispose(): void }[] = []
  let signature = ''
  const clear = (): void => {
    for (const item of disposables) item.dispose()
    disposables = []
  }
  const refresh = (): void => {
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const facade = workbook?.getActiveSheet()
    if (!workbook || !facade) return
    const sheet = facade.getSheet()
    const found: {
      row: number
      column: number
      mode: AccountingUnderline
      style: IStyleData
      text: string
    }[] = []
    sheet.getCellMatrix().forValue((row, column, cell) => {
      const mode = cell?.custom?.[ACCOUNTING_UNDERLINE_KEY]
      if (mode !== 'singleAccounting' && mode !== 'doubleAccounting') return true
      const style = sheet.getComposedCellStyle(row, column)
      if (!cell?.p && !overlayAngleOf(style)) return true
      if (found.length < MAX_OVERLAYS) {
        const text =
          cell?.p?.body?.dataStream?.replace(/[\r\n]+$/, '') ??
          (typeof cell?.v === 'string' ? cell.v : String(cell?.v ?? ''))
        found.push({ row, column, mode, style, text })
      }
      return true
    })
    const next = found
      .map((item) => `${item.row}:${item.column}:${item.mode}:${item.style.tr?.a ?? 0}`)
      .join('|')
    // float layers are keyed by cell: re-installing identical ones every mutation would flicker
    if (next === signature && (next === '' || disposables.length > 0)) return
    signature = next
    clear()
    for (const item of found) {
      const range = facade.getRange(item.row, item.column, 1, 1)
      const merged = sheet.getMergedCell(item.row, item.column)
      if (merged) continue
      const width = facade.getColumnWidth(item.column)
      const height = facade.getRowHeight(item.row)
      const angle = overlayAngleOf(item.style)
      const strokes = accountingOverlayStrokes({
        width,
        height,
        mode: item.mode,
        angle,
        textLength: angle ? measure(item.text, item.style) : 0,
        horizontal: HORIZONTAL[(item.style.ht as 1 | 2 | 3) ?? 1] ?? 'left',
        vertical: VERTICAL[(item.style.vt as 1 | 2 | 3) ?? 3] ?? 'bottom',
      })
      if (strokes.length === 0) continue
      const key = `accounting-overlay-${item.row}-${item.column}-${item.mode}`
      const color = item.style.cl?.rgb ?? '#000000'
      disposables.push(
        runtime.univerAPI.registerComponent(key, () => (
          <svg
            className="accounting-underline-overlay"
            viewBox={`0 0 ${width} ${height}`}
            width={width}
            height={height}
            aria-hidden="true"
            data-accounting-overlay={item.mode}
          >
            {strokes.map((stroke, index) => (
              <line key={index} {...stroke} stroke={color} strokeWidth={1} />
            ))}
          </svg>
        )),
      )
      const floating = facade.addFloatDomToRange(
        range,
        { componentKey: key, allowTransform: false, eventPassThrough: true },
        {},
        key,
      )
      if (floating) disposables.push(floating)
    }
  }
  return {
    refresh,
    dispose: () => {
      signature = ''
      clear()
    },
  }
}
