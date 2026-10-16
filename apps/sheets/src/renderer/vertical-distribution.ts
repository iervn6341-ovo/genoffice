import { VerticalAlign, WrapStrategy } from '@univerjs/core'
import { DocSimpleSkeleton, Documents, Font, Text } from '@univerjs/engine-render'

export const VERTICAL_DISTRIBUTION_KEY = 'genofficeVerticalDistribution'
export type VerticalDistribution = 'justify' | 'distributed'

interface RenderLine {
  top: number
  lineHeight: number
}

interface RenderPage {
  sections: { columns: { lines: RenderLine[] }[] }[]
  renderConfig?: { verticalAlign?: VerticalAlign } | undefined
  skeTables?: { size: number }
}

interface RenderDocument {
  getSkeleton?: () => { getSkeletonData?: () => { pages: RenderPage[] } | null } | null
}

interface DistributionContext {
  mode: VerticalDistribution
  innerHeight: number
}

const active = new WeakMap<object, DistributionContext>()
let activePlainMode: VerticalDistribution | null = null
let insidePlainLine = false

/** Extra offsets for each visual line; Justify pins the ends, Distributed gives half-gaps at them. */
export function verticalDistributionOffsets(
  lines: readonly RenderLine[],
  innerHeight: number,
  mode: VerticalDistribution,
): number[] | null {
  if (lines.length < 2 || lines.length > 256 || !Number.isFinite(innerHeight)) return null
  const first = lines[0]
  const last = lines[lines.length - 1]
  if (!first || !last || lines.some((line) => !Number.isFinite(line.top + line.lineHeight))) {
    return null
  }
  const naturalSpan = last.top + last.lineHeight - first.top
  if (!(naturalSpan > 0)) return null
  const remaining = Math.max(0, innerHeight - naturalSpan)
  const gap = remaining / (mode === 'justify' ? lines.length - 1 : lines.length)
  return lines.map((_line, index) => (mode === 'justify' ? index * gap : (index + 0.5) * gap))
}

let installed = false

/** Distributes the already-laid-out document lines only for tagged spreadsheet cells. */
export function installVerticalDistributionRender(): void {
  if (installed) return
  installed = true
  const fontProto = Font.prototype as unknown as {
    getDocuments(): object | null
    _renderDocuments(
      ctx: unknown,
      row: number,
      column: number,
      renderFontCtx: {
        startY: number
        endY: number
        fontCache?: { vertexAngle?: number; style?: { pd?: { t?: number; b?: number } } }
        spreadsheetSkeleton?: {
          worksheet?: {
            getCell?(row: number, column: number): { custom?: Record<string, unknown> } | null
          }
        }
      },
      overflowCache: unknown,
    ): void
    _renderText(
      ctx: unknown,
      row: number,
      column: number,
      renderFontCtx: {
        fontCache?: { vertexAngle?: number; wrapStrategy?: WrapStrategy }
        spreadsheetSkeleton?: {
          worksheet?: {
            getCell?(row: number, column: number): { custom?: Record<string, unknown> } | null
          }
        }
      },
      overflowCache: unknown,
    ): void
  }
  const originalTextRender = fontProto._renderText
  fontProto._renderText = function (ctx, row, column, renderFontCtx, overflowCache) {
    const mode = renderFontCtx.spreadsheetSkeleton?.worksheet?.getCell?.(row, column)?.custom?.[
      VERTICAL_DISTRIBUTION_KEY
    ]
    if (
      (mode !== 'justify' && mode !== 'distributed') ||
      renderFontCtx.fontCache?.wrapStrategy !== WrapStrategy.WRAP ||
      renderFontCtx.fontCache?.vertexAngle
    ) {
      originalTextRender.call(this, ctx, row, column, renderFontCtx, overflowCache)
      return
    }
    const previous = activePlainMode
    activePlainMode = mode
    try {
      originalTextRender.call(this, ctx, row, column, renderFontCtx, overflowCache)
    } finally {
      activePlainMode = previous
    }
  }
  const originalTextDraw = Text.drawWith
  Text.drawWith = function (ctx, props, skeleton) {
    if (!activePlainMode || insidePlainLine)
      return originalTextDraw.call(this, ctx, props, skeleton)
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
    let naturalTop = 0
    const geometry = lines.map((line) => {
      const current = { top: naturalTop, lineHeight: line.height }
      naturalTop += line.height
      return current
    })
    const offsets = verticalDistributionOffsets(geometry, props.height, activePlainMode)
    if (!offsets) return originalTextDraw.call(this, ctx, props, skeleton)
    insidePlainLine = true
    try {
      for (const [index, line] of lines.entries()) {
        originalTextDraw.call(this, ctx, {
          ...props,
          text: line.text,
          warp: false,
          vAlign: VerticalAlign.TOP,
          top: (props.top ?? 0) + (geometry[index]?.top ?? 0) + (offsets[index] ?? 0),
          height: line.height,
        })
      }
    } finally {
      insidePlainLine = false
    }
    return laidOut.getTotalHeight()
  }
  const originalFontRender = fontProto._renderDocuments
  fontProto._renderDocuments = function (ctx, row, column, renderFontCtx, overflowCache) {
    const mode = renderFontCtx.spreadsheetSkeleton?.worksheet?.getCell?.(row, column)?.custom?.[
      VERTICAL_DISTRIBUTION_KEY
    ]
    const docs = this.getDocuments()
    if (
      (mode !== 'justify' && mode !== 'distributed') ||
      !docs ||
      renderFontCtx.fontCache?.vertexAngle
    ) {
      originalFontRender.call(this, ctx, row, column, renderFontCtx, overflowCache)
      return
    }
    const padding = renderFontCtx.fontCache?.style?.pd
    active.set(docs, {
      mode,
      innerHeight:
        renderFontCtx.endY - renderFontCtx.startY - (padding?.t ?? 2) - (padding?.b ?? 2),
    })
    try {
      originalFontRender.call(this, ctx, row, column, renderFontCtx, overflowCache)
    } finally {
      active.delete(docs)
    }
  }

  const documentsProto = Documents.prototype as unknown as {
    draw(ctx: unknown, bounds?: unknown): void
  }
  const originalDraw = documentsProto.draw
  documentsProto.draw = function (ctx, bounds) {
    const context = active.get(this)
    if (!context) {
      originalDraw.call(this, ctx, bounds)
      return
    }
    const pages = (this as RenderDocument).getSkeleton?.()?.getSkeletonData?.()?.pages
    const page = pages?.length === 1 ? pages[0] : undefined
    const section = page?.sections.length === 1 ? page.sections[0] : undefined
    const column = section?.columns.length === 1 ? section.columns[0] : undefined
    const lines = column?.lines
    if (!page || !lines || page.skeTables?.size || lines.length < 2) {
      originalDraw.call(this, ctx, bounds)
      return
    }
    const offsets = verticalDistributionOffsets(lines, context.innerHeight, context.mode)
    if (!offsets) {
      originalDraw.call(this, ctx, bounds)
      return
    }
    const originalTops = lines.map((line) => line.top)
    const originalConfig = page.renderConfig
    page.renderConfig = { ...originalConfig, verticalAlign: VerticalAlign.TOP }
    for (const [index, line] of lines.entries()) line.top += offsets[index] ?? 0
    try {
      originalDraw.call(this, ctx, bounds)
    } finally {
      for (const [index, line] of lines.entries()) line.top = originalTops[index] ?? line.top
      page.renderConfig = originalConfig
    }
  }
}
