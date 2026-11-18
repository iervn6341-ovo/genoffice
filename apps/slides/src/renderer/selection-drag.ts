type Point = { x: number; y: number }
import type { RenderNode } from '@genoffice/pptx-render'
import type { TransformReq } from './history-group'

export function stopSelectionDrags(
  nodes: Array<{ isDragging(): boolean; stopDrag(): void } | null | undefined>,
): void {
  for (const node of nodes) if (node?.isDragging()) node.stopDrag()
}

export function selectionMoveRequests(
  nodes: RenderNode[],
  delta: Point,
  groupId?: string,
  toLocal: (id: string, delta: Point) => Point = (_id, d) => d,
): TransformReq[] {
  return nodes.map((node) => {
    const local = toLocal(node.sourceId, delta)
    return {
      sourceId: node.sourceId,
      box: { ...node.box, x: node.box.x + local.x, y: node.box.y + local.y },
      ...(groupId ? { groupId } : {}),
    }
  })
}

/** Every member uses the same pointer displacement, independent of Konva proxy-drag offsets. */
export function createSelectionDrag() {
  let gesture: {
    pointer: Point
    positions: Map<string, Point>
    constrain: (delta: Point) => Point
    last?: { pointer: Point; delta: Point }
  } | null = null
  return {
    begin(pointer: Point, positions: Map<string, Point>, constrain: (delta: Point) => Point) {
      gesture = {
        pointer: { ...pointer },
        positions: new Map([...positions].map(([id, p]) => [id, { ...p }])),
        constrain,
      }
    },
    reset() {
      gesture = null
    },
    finish(): Point | null {
      const delta = gesture?.last?.delta ?? null
      gesture = null
      return delta && { ...delta }
    },
    bound(id: string, pointer: Point): Point | null {
      const start = gesture?.positions.get(id)
      if (!gesture || !start) return null
      if (
        !gesture.last ||
        gesture.last.pointer.x !== pointer.x ||
        gesture.last.pointer.y !== pointer.y
      ) {
        gesture.last = {
          pointer: { ...pointer },
          delta: gesture.constrain({
            x: pointer.x - gesture.pointer.x,
            y: pointer.y - gesture.pointer.y,
          }),
        }
      }
      const { delta } = gesture.last
      return { x: start.x + delta.x, y: start.y + delta.y }
    },
  }
}

export type SelectionDrag = ReturnType<typeof createSelectionDrag>
