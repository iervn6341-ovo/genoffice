/**
 * Format painter (renderer side) — extracts formatting from a RenderNode and converts it into
 * IPC call parameters.
 *
 * Format scope:
 *  - shape/text: fill, stroke, default text run (first run of first paragraph), paragraph alignment
 *  - picture: stroke
 *  - cross-type application takes the intersection
 *
 * This module has no side effects; App.tsx calls it.
 */
import type {
  ShapeRenderNode,
  PictureRenderNode,
  RenderNode,
  RenderFill,
} from '@genoffice/pptx-render'
import type {
  EditFillOp,
  EditStrokeOp,
  SetElementFontOp,
  SetElementParagraphFormatOp,
  SetEffectsPatch,
  SlidesApi,
} from '../shared/ipc'
import type { Op } from '@genoffice/pptx-ops'

/** Format container captured by the render layer (pure data, serializable) */
export interface BrushFormat {
  sourceType: 'shape' | 'text' | 'picture' | 'group'
  children?: Array<BrushFormat | null>
  effects?: SetEffectsPatch
  body?: {
    anchor: 'top' | 'middle' | 'bottom'
    props: Parameters<SlidesApi['setTextBodyProps']>[0]['props']
  }
  paragraph?: Omit<SetElementParagraphFormatOp, 'slideIndex' | 'sourceIds'>
  /** solid fill color (#RRGGBB) or 'none'; undefined = no fill info */
  fillColor?: string
  fill?: EditFillOp['fill']
  /** stroke; undefined = no stroke info */
  stroke?: EditStrokeOp['stroke']
  /** Default text run format (extracted from text.lines[0].runs[0]) */
  run?: {
    fontFamily?: string
    fontSize?: number // pt
    bold?: boolean
    italic?: boolean
    underline?: boolean
    color?: string
    strike?: boolean
    highlight?: string | null
    baseline?: number
    letterSpacingPt?: number
  }
  /** Paragraph alignment (first paragraph) */
  align?: 'left' | 'center' | 'right' | 'justify'
}

/** px→pt conversion (96dpi) */
function pxToPt(px: number): number {
  return Math.round(((px * 72) / 96) * 100) / 100
}

/** RenderFill → fillColor string or undefined */
function renderFillToColor(fill: RenderFill | undefined): string | undefined {
  if (!fill || fill.kind === 'none') return 'none'
  if (fill.kind === 'solid') return fill.color
  // Non-solid fills are handled separately.
  return undefined
}

/** Extract format from a ShapeRenderNode / text node */
export function extractBrushFormat(node: RenderNode, scale = 1): BrushFormat | null {
  if (node.type === 'group') {
    const children = node.children.map((child) => extractBrushFormat(child, scale))
    return children.some(Boolean) ? { sourceType: 'group', children } : null
  }
  if (node.type === 'shape' || node.type === 'text') {
    const shape = node as ShapeRenderNode
    const fmt: BrushFormat = { sourceType: node.type as 'shape' | 'text' }
    fmt.effects = brushEffects(shape, scale)

    // fill
    const fc = renderFillToColor(shape.fill)
    if (fc !== undefined) fmt.fillColor = fc
    if (shape.fill?.kind === 'gradient') {
      const g = shape.fill
      fmt.fill = {
        gradient: {
          from: g.stops[0]?.color ?? '#000000',
          to: g.stops.at(-1)?.color ?? '#000000',
          stops: g.stops.map((s) => ({ ...s })),
          angleDeg: g.angleDeg,
          ...(g.path ? { path: g.path } : g.radial ? { radial: true } : {}),
          ...(g.center ? { center: { ...g.center } } : {}),
        },
      }
    }

    // stroke
    if (shape.stroke) {
      fmt.stroke = brushStroke(shape.stroke)
    } else fmt.stroke = null

    // text: first run of first line (para start)
    if (shape.text?.lines?.length) {
      const firstLine = shape.text.lines.find((l) => l.paraStart !== false) ?? shape.text.lines[0]
      const firstRun = firstLine?.runs?.find((r) => !r.isBullet) ?? firstLine?.runs?.[0]
      if (firstRun) {
        fmt.run = {
          fontFamily: firstRun.srcFontFamily ?? firstRun.fontFamily,
          fontSize: pxToPt(firstRun.fontSizePx / scale / (shape.text.fontScale || 1)),
          bold: firstRun.bold,
          italic: firstRun.italic,
          underline: firstRun.underline,
          color: firstRun.color,
          strike: firstRun.strike ?? false,
          highlight: firstRun.highlight ?? null,
          baseline: firstRun.baselinePct ?? 0,
          letterSpacingPt: pxToPt((firstRun.letterSpacingPx ?? 0) / scale),
        }
      }
      // align from first line
      const firstParaLine =
        shape.text.lines.find((l) => l.paraStart !== false) ?? shape.text.lines[0]
      fmt.align = firstParaLine?.align ?? 'left'
      const bullet = firstParaLine?.runs.find((r) => r.isBullet)
      fmt.paragraph = {
        align: fmt.align,
        rtl: firstParaLine?.rtl ?? false,
        ...(bullet?.numType
          ? { bullet: 'number', numType: bullet.numType }
          : bullet
            ? { bullet: 'char', bulletChar: bullet.text, bulletFont: bullet.fontFamily }
            : { bullet: 'none' }),
      }
      const { insets, anchor, wrap, autofit, vert } = shape.text
      fmt.body = {
        anchor,
        props: {
          wrap,
          autofit: autofit ?? 'none',
          vert: vert ?? 'horz',
          ...(insets
            ? {
                insets: Object.fromEntries(
                  Object.entries(insets).map(([key, px]) => [key, Math.round((px * 9525) / scale)]),
                ),
              }
            : {}),
        },
      }
    }

    return fmt
  }

  if (node.type === 'picture') {
    const pic = node as PictureRenderNode
    const fmt: BrushFormat = { sourceType: 'picture' }
    fmt.effects = brushEffects(pic, scale)
    if (pic.stroke) {
      fmt.stroke = brushStroke(pic.stroke)
    } else fmt.stroke = null
    return fmt
  }

  return null
}

function brushEffects(node: ShapeRenderNode | PictureRenderNode, scale: number): SetEffectsPatch {
  const emu = (px: number) => Math.round((px * 9525) / scale)
  const shadow = node.shadow
  return {
    shadow: shadow
      ? {
          color: shadow.color,
          blurRad: emu(shadow.blurPx),
          dist: emu(shadow.distPx ?? Math.hypot(shadow.offsetX, shadow.offsetY)),
          dirDeg: shadow.dirDeg ?? (Math.atan2(shadow.offsetY, shadow.offsetX) * 180) / Math.PI,
          inner: shadow.inner,
          ...(shadow.scaleX != null ? { sx: shadow.scaleX } : {}),
          ...(shadow.scaleY != null ? { sy: shadow.scaleY } : {}),
          ...(shadow.skewXDeg != null ? { kxDeg: shadow.skewXDeg } : {}),
          ...(shadow.skewYDeg != null ? { kyDeg: shadow.skewYDeg } : {}),
          ...(shadow.algn ? { algn: shadow.algn } : {}),
        }
      : null,
    glow: node.glow ? { color: node.glow.color, radius: emu(node.glow.blurPx) } : null,
    reflection: node.reflection
      ? {
          blurRad: emu(node.reflection.blurPx),
          startA: node.reflection.startAlpha,
          endPos: node.reflection.endPos,
          dist: emu(node.reflection.distPx),
        }
      : null,
    softEdge: 'softEdgePx' in node && node.softEdgePx ? emu(node.softEdgePx) : null,
  }
}

function brushStroke(
  stroke: NonNullable<ShapeRenderNode['stroke']>,
): NonNullable<EditStrokeOp['stroke']> {
  return {
    color: stroke.color,
    widthPt: stroke.widthPt,
    dash: stroke.dashPreset ?? 'solid',
    cap: stroke.cap === 'round' ? 'rnd' : stroke.cap === 'square' ? 'sq' : 'flat',
    join: stroke.join ?? 'miter',
    compound: stroke.compound ?? 'sng',
    ...(stroke.gradient ? { gradient: structuredClone(stroke.gradient) } : {}),
  }
}

/** Render group children are not top-level nodes; retain their owning group for native edits. */
export function findBrushTarget(
  nodes: RenderNode[],
  id: string,
  groupId?: string,
): { node: RenderNode; groupId?: string } | undefined {
  for (const node of nodes) {
    if (node.sourceId === id) return { node, ...(groupId ? { groupId } : {}) }
    if (node.type === 'group') {
      const found = findBrushTarget(node.children, id, node.sourceId)
      if (found) return found
    }
  }
}

/** Format application result (caller invokes IPC as needed) */
export interface BrushApplyPlan {
  /** editFill parameters */
  fillColor?: string
  fill?: EditFillOp['fill']
  /** editStroke parameters */
  stroke?: EditStrokeOp['stroke']
  /** editText parameters (replaces format only, text unchanged; needs to read the current node's text) */
  runFormat?: BrushFormat['run']
  /** Paragraph alignment */
  align?: 'left' | 'center' | 'right' | 'justify'
}

/**
 * Apply BrushFormat to the target node type, producing an IPC call plan.
 * Cross-type takes the intersection: picture only receives stroke.
 */
export function computeBrushApply(fmt: BrushFormat, targetNode: RenderNode): BrushApplyPlan {
  const plan: BrushApplyPlan = {}

  if (targetNode.type === 'picture') {
    if (fmt.stroke !== undefined) plan.stroke = fmt.stroke
    return plan
  }

  if (targetNode.type === 'shape' || targetNode.type === 'text') {
    if (fmt.fillColor !== undefined) plan.fillColor = fmt.fillColor
    if (fmt.fill !== undefined) plan.fill = fmt.fill
    if (fmt.stroke !== undefined) plan.stroke = fmt.stroke
    if (fmt.run) plan.runFormat = fmt.run
    if (fmt.align) plan.align = fmt.align
    return plan
  }

  return plan
}

/** Native font patches preserve fields, equations, hyperlinks, soft breaks and paragraph metadata. */
export function brushFontPatch(
  run: NonNullable<BrushFormat['run']>,
): Omit<SetElementFontOp, 'slideIndex' | 'sourceIds'> {
  const { fontSize, ...rest } = run
  return { ...rest, ...(fontSize != null ? { fontSizePt: fontSize } : {}) }
}

export function buildBrushOps(
  brush: BrushFormat,
  node: RenderNode,
  slideIndex: number,
  groupId?: string,
): Op[] {
  if (node.type === 'group')
    return node.children.flatMap((child, i) => {
      const format = brush.sourceType === 'group' ? brush.children?.[i] : brush
      return format ? buildBrushOps(format, child, slideIndex, node.sourceId) : []
    })
  if (brush.sourceType === 'group') return []
  const plan = computeBrushApply(brush, node)
  const target = { slide: slideIndex, el: node.sourceId }
  const group = groupId ? { group: groupId } : {}
  const ops: Op[] = []
  const fill = plan.fill ?? plan.fillColor
  if (fill !== undefined) {
    const g = typeof fill === 'object' ? fill.gradient : undefined
    ops.push({
      op: 'setFill',
      target,
      ...group,
      fill: g
        ? {
            stops: g.stops ?? [
              { pos: 0, color: g.from },
              { pos: 1, color: g.to },
            ],
            ...(g.path || g.radial
              ? {
                  path: g.path ?? 'circle',
                  ...(g.center
                    ? {
                        fillTo: {
                          l: g.center.x,
                          t: g.center.y,
                          r: 1 - g.center.x,
                          b: 1 - g.center.y,
                        },
                      }
                    : {}),
                }
              : { angle: Math.round((g.angleDeg ?? 0) * 60000) }),
          }
        : fill,
    })
  }
  if (plan.stroke !== undefined) {
    const stroke = plan.stroke
    ops.push({
      op: 'setStroke',
      target,
      ...group,
      stroke: stroke
        ? {
            ...stroke,
            widthEmu: Math.round(stroke.widthPt * 12700),
            ...(stroke.gradient
              ? {
                  gradient: {
                    stops: stroke.gradient.stops,
                    angle: Math.round(stroke.gradient.angleDeg * 60000),
                  },
                }
              : {}),
          }
        : null,
    })
  }
  if (
    plan.runFormat &&
    (node.type === 'shape' || node.type === 'text') &&
    node.text?.lines.length
  ) {
    ops.push({ op: 'setFont', target, ...group, font: brushFontPatch(plan.runFormat) })
  }
  if (plan.align && (node.type === 'shape' || node.type === 'text') && node.text?.lines.length) {
    ops.push({
      op: 'setParagraphFormat',
      target,
      ...group,
      format: brush.paragraph ?? { align: plan.align },
    })
  }
  if (
    !groupId &&
    brush.body &&
    (node.type === 'shape' || node.type === 'text') &&
    node.text?.lines.length
  ) {
    ops.push({ op: 'setTextAnchor', target, anchor: brush.body.anchor })
    ops.push({ op: 'setTextBodyProps', target, props: brush.body.props })
  }
  if (
    !groupId &&
    brush.effects &&
    (node.type === 'shape' || node.type === 'text' || node.type === 'picture')
  ) {
    ops.push({ op: 'setEffects', target, effects: brush.effects })
  }
  return ops
}
