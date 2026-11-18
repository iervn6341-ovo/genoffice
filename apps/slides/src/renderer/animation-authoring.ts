import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'
import type { AnimationItem } from '../shared/ipc'

/** Click order, shared by the pane and canvas. 0 means automatic on page entry. */
export function animationStepNumbers(items: AnimationItem[]): number[] {
  let step = 0
  return items.map((item) => {
    if (item.trigger === 'onClick') step++
    return step
  })
}

type Point = { x: number; y: number }
type Box = RenderNode['box']
export function animationTargetSelection(
  slide: RenderSlide,
  sourceId: string,
): { selectedId: string; groupId: string | null } {
  for (const node of slide.nodes) {
    if (node.sourceId === sourceId) return { selectedId: sourceId, groupId: null }
    if (node.type !== 'group') continue
    if (node.children.some((child) => child.sourceId === sourceId))
      return { selectedId: sourceId, groupId: node.sourceId }
    const contains = (children: RenderNode[]): boolean =>
      children.some(
        (child) =>
          child.sourceId === sourceId || (child.type === 'group' && contains(child.children)),
      )
    if (contains(node.children)) return { selectedId: node.sourceId, groupId: null }
  }
  return { selectedId: sourceId, groupId: null }
}
function toParent(point: Point, box: Box): Point {
  const x = (point.x - box.w / 2) * (box.flipH ? -1 : 1)
  const y = (point.y - box.h / 2) * (box.flipV ? -1 : 1)
  const angle = ((box.rotationDeg ?? 0) * Math.PI) / 180
  return {
    x: box.x + box.w / 2 + x * Math.cos(angle) - y * Math.sin(angle),
    y: box.y + box.h / 2 + x * Math.sin(angle) + y * Math.cos(angle),
  }
}

export interface AnimationMarker {
  index: number
  step: number
  x: number
  y: number
  sourceId: string
  /** The top group to select when a nested animated target isn't entered. */
  groupId?: string
}

/** Group-local coordinates have already been scaled by the renderer; apply rotation/flip once. */
export function animationMarkers(
  slide: RenderSlide,
  items: AnimationItem[],
  zoom: number,
): AnimationMarker[] {
  const targets = new Map<string, { node: RenderNode; parents: RenderNode[] }>()
  const walk = (nodes: RenderNode[], parents: RenderNode[]) => {
    for (const node of nodes) {
      if (node.decoration) continue
      targets.set(node.sourceId, { node, parents })
      if (node.type === 'group') walk(node.children, [...parents, node])
    }
  }
  walk(slide.nodes, [])
  const steps = animationStepNumbers(items)
  const scale = Math.max(0.1, zoom)
  const width = 26 / scale
  const height = 20 / scale
  const gap = 3 / scale
  const markers: AnimationMarker[] = []
  items.forEach((item, index) => {
    const target = targets.get(item.sourceId)
    if (!target) return
    const { node, parents } = target
    let y0 = 0
    let y1 = node.box.h
    if (item.paragraph != null && (node.type === 'text' || node.type === 'shape') && node.text) {
      let paragraph = -1
      const lines = node.text.lines.filter((line) => {
        if (line.paraStart !== false) paragraph++
        return paragraph === item.paragraph
      })
      if (lines.length) {
        y0 = node.text.insets.t + lines[0]!.top
        y1 = node.text.insets.t + lines.at(-1)!.top + lines.at(-1)!.height
      }
    }
    const points = [
      { x: 0, y: y0 },
      { x: node.box.w, y: y0 },
      { x: 0, y: y1 },
      { x: node.box.w, y: y1 },
    ].map((p) => {
      let point = toParent(p, node.box)
      for (let i = parents.length - 1; i >= 0; i--) point = toParent(point, parents[i]!.box)
      return point
    })
    const x = Math.max(
      0,
      Math.min(slide.widthPx - width, Math.min(...points.map((p) => p.x)) - width - gap),
    )
    let y = Math.max(0, Math.min(slide.heightPx - height, Math.min(...points.map((p) => p.y))))
    // Stack colliding badges, including multiple effects on one target.
    for (let tries = 0; tries < items.length; tries++) {
      const hit = markers.find(
        (m) => Math.abs(m.x - x) < width + gap && Math.abs(m.y - y) < height + gap,
      )
      if (!hit) break
      y = hit.y + height + gap
    }
    markers.push({
      index,
      step: steps[index]!,
      x,
      y,
      sourceId: item.sourceId,
      ...(parents[0] ? { groupId: parents[0].sourceId } : {}),
    })
  })
  return markers
}
