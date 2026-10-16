import { TextDecoration, VerticalAlign } from '@univerjs/core'
import { DocSimpleSkeleton, Font, Text } from '@univerjs/engine-render'

import type { WorkbookCellStyle } from '../shared/desktop-api'

export const ACCOUNTING_UNDERLINE_KEY = 'genofficeAccountingUnderline'
export type AccountingUnderline = 'singleAccounting' | 'doubleAccounting'

export function accountingUnderlineMode(
  style: Pick<WorkbookCellStyle, 'underline' | 'underlineStyle'> | undefined,
): AccountingUnderline | null {
  return style?.underline &&
    (style.underlineStyle === 'singleAccounting' || style.underlineStyle === 'doubleAccounting')
    ? style.underlineStyle
    : null
}

interface LineGeometry {
  readonly height: number
}

/** Keep the accounting rule under descenders, inside the cell's text area. */
export function accountingUnderlineLines(
  lines: readonly LineGeometry[],
  bounds: { left: number; top: number; width: number; height: number },
  totalHeight: number,
  vertical: VerticalAlign,
  mode: AccountingUnderline,
): { x: number; y: number; width: number; lineType: TextDecoration }[] {
  if (
    lines.length === 0 ||
    lines.length > 256 ||
    !Number.isFinite(bounds.width) ||
    bounds.width <= 0 ||
    !Number.isFinite(bounds.height) ||
    bounds.height <= 0
  )
    return []
  const offset =
    vertical === VerticalAlign.TOP
      ? 0
      : vertical === VerticalAlign.MIDDLE
        ? (bounds.height - totalHeight) / 2
        : bounds.height - totalHeight
  const lineType = mode === 'doubleAccounting' ? TextDecoration.DOUBLE : TextDecoration.SINGLE
  let top = bounds.top + offset
  return lines.map((line) => {
    // Univer's ordinary underline sits at lineHeight - 1 and crosses the
    // descenders. Excel's accounting line sits below them and spans the cell.
    const y = Math.min(
      top + line.height + 1,
      bounds.top + bounds.height - (mode === 'doubleAccounting' ? 3 : 1),
    )
    top += line.height
    return { x: bounds.left, y, width: bounds.width, lineType }
  })
}

let installed = false
let active: AccountingUnderline | null = null

/** Replace only plain-cell underline strokes; all other Univer text stays native. */
export function installAccountingUnderlineRender(): void {
  if (installed) return
  installed = true
  const fontProto = Font.prototype as unknown as {
    _renderText(
      ctx: unknown,
      row: number,
      column: number,
      renderFontCtx: {
        fontCache?: { vertexAngle?: number }
        spreadsheetSkeleton?: {
          worksheet?: {
            getCell?(row: number, column: number): { custom?: Record<string, unknown> } | null
          }
        }
      },
      overflowCache: unknown,
    ): void
  }
  const originalFontRender = fontProto._renderText
  fontProto._renderText = function (ctx, row, column, renderFontCtx, overflowCache) {
    const marker = renderFontCtx.spreadsheetSkeleton?.worksheet?.getCell?.(row, column)?.custom?.[
      ACCOUNTING_UNDERLINE_KEY
    ]
    if (
      (marker !== 'singleAccounting' && marker !== 'doubleAccounting') ||
      renderFontCtx.fontCache?.vertexAngle
    ) {
      originalFontRender.call(this, ctx, row, column, renderFontCtx, overflowCache)
      return
    }
    const previous = active
    active = marker
    try {
      originalFontRender.call(this, ctx, row, column, renderFontCtx, overflowCache)
    } finally {
      active = previous
    }
  }

  const originalTextDraw = Text.drawWith
  const decoration = Text as unknown as {
    _drawTextDecoration(
      ctx: unknown,
      options: {
        x: number
        y: number
        width: number
        color: string
        lineWidth: number
        lineType: TextDecoration
      },
    ): void
  }
  Text.drawWith = function (ctx, props, skeleton) {
    const mode = active
    if (!mode || !props.underline) return originalTextDraw.call(this, ctx, props, skeleton)
    const laidOut =
      skeleton ??
      new DocSimpleSkeleton(
        props.text,
        props.fontStyle,
        Boolean(props.warp),
        props.width,
        props.vAlign === VerticalAlign.TOP ? props.height : Number.POSITIVE_INFINITY,
      )
    const lines = laidOut.calculate()
    const strokes = accountingUnderlineLines(
      lines,
      {
        left: props.left ?? 0,
        top: props.top ?? 0,
        width: props.width,
        height: props.height,
      },
      laidOut.getTotalHeight(),
      props.vAlign ?? VerticalAlign.BOTTOM,
      mode,
    )
    const result = originalTextDraw.call(this, ctx, { ...props, underline: false }, laidOut)
    for (const stroke of strokes) {
      decoration._drawTextDecoration(ctx, {
        ...stroke,
        color: props.color ?? '#000000',
        lineWidth: 1,
      })
    }
    return result
  }
}
