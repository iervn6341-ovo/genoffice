import { expect, it } from 'vitest'
import { Group } from 'konva/lib/Group'
import { DD } from 'konva/lib/DragAndDrop'
import { Stage } from 'konva/lib/Stage'
import { stopSelectionDrags } from '../src/renderer/selection-drag'

it('ends newly created Konva proxy drags even when they missed the mouseup pointer-id filter', () => {
  expect(Stage).toBeTypeOf('function') // Register the runtime class used by DD's instanceof guard.
  let pointer = { id: 999, x: 0, y: 0 }
  const stage = {
    get _changedPointerPositions() {
      return [pointer]
    },
    _getPointerById: () => pointer,
    batchDraw: () => {},
    setPointersPositions: (evt: MouseEvent) => {
      pointer = { id: 999, x: evt.clientX, y: evt.clientY }
    },
  } as unknown as Stage
  const nodes = [
    new Group({ x: 10, y: 20 }),
    new Group({ x: 110, y: 20 }),
    new Group({ x: 210, y: 20 }),
  ]
  for (const node of nodes) node.getStage = () => stage
  try {
    nodes[0]!.on('dragmove', (event) => {
      for (const peer of nodes.slice(1)) if (!peer.isDragging()) peer.startDrag(event)
    })
    nodes[0]!.startDrag()
    const move = new MouseEvent('mousemove', { clientX: 20, clientY: 30 })
    DD._drag(move)
    const up = new MouseEvent('mouseup', { clientX: 20, clientY: 30 })
    DD._endDragBefore(up)
    DD._endDragAfter(up)
    expect(nodes[0]!.isDragging()).toBe(false)
    expect(nodes.slice(1).some((n) => n.isDragging())).toBe(true)
    stopSelectionDrags(nodes)
    expect(nodes.some((n) => n.isDragging())).toBe(false)
    const positions = nodes.map((n) => n.position())
    DD._drag(new MouseEvent('mousemove', { clientX: 900, clientY: 800 }))
    expect(nodes.map((n) => n.position())).toEqual(positions)
  } finally {
    DD._dragElements.clear()
    nodes.forEach((node) => node.destroy())
  }
})
