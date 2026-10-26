/**
 * Minimum size of a selection of render nodes (single element, group, table or a multi-selection),
 * shared by the canvas Transformer and the numeric size fields so a dragged handle and a typed
 * value obey the same floor.
 */
import type {
  GroupRenderNode,
  RenderNode,
  ShapeRenderNode,
  TableRenderNode,
} from '@genoffice/pptx-render'
import { minSizeFor, minSizeForMembers, type MinNode, type MinSize } from './min-size'

/** Whether a node is a connector (read-only, no Transformer attached). */
export function isConnectorNode(node: RenderNode): boolean {
  return (node.type === 'shape' || node.type === 'text') && !!(node as ShapeRenderNode).line
}

const gaps = (g: number[]): number[] => g.slice(1).map((v, i) => v - g[i]!)

/** A render node reduced to what the minimum-size rules need (tables: grid; groups: children) */
export function toMinNode(n: RenderNode): MinNode {
  const base: MinNode = {
    type: n.type,
    isConnector: isConnectorNode(n),
    w: n.box.w,
    h: n.box.h,
  }
  if (n.type === 'table') {
    const t = n as TableRenderNode
    return {
      ...base,
      rows: t.gridY.length - 1,
      cols: t.gridX.length - 1,
      colWidths: gaps(t.gridX),
      rowHeights: gaps(t.gridY),
      ...(t.rowNeedPx ? { rowNeeds: t.rowNeedPx } : {}),
      ...(t.colNeedPx ? { colNeeds: t.colNeedPx } : {}),
      ...(t.fitMinWPx ? { fitMinW: t.fitMinWPx } : {}),
    }
  }
  if (n.type === 'group') {
    return { ...base, children: (n as GroupRenderNode).children.map(toMinNode) }
  }
  return base
}

/** Axis-aligned bounds of the selected boxes (rotation ignored: the Transformer scales per axis) */
export function selectionBounds(nodes: RenderNode[]): { w: number; h: number } {
  const x0 = Math.min(...nodes.map((n) => n.box.x))
  const y0 = Math.min(...nodes.map((n) => n.box.y))
  const x1 = Math.max(...nodes.map((n) => n.box.x + n.box.w))
  const y1 = Math.max(...nodes.map((n) => n.box.y + n.box.h))
  return { w: x1 - x0, h: y1 - y0 }
}

/** Floor of resizing the given selection as one unit */
export function minSizeOfNodes(nodes: RenderNode[]): MinSize {
  if (nodes.length === 0) return minSizeFor({ type: 'multi' })
  if (nodes.length === 1) return minSizeFor(toMinNode(nodes[0]!))
  const bb = selectionBounds(nodes)
  return minSizeForMembers(nodes.map(toMinNode), bb.w, bb.h)
}
