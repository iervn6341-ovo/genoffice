/**
 * Live move/resize gestures of the canvas nodes. The model only learns the new box when
 * the gesture ends, so DOM overlays pinned to a node (the table cell-selection layer)
 * listen here to follow — or step aside — while it is in flight.
 */
export interface NodeGestureDetail {
  sourceId: string
  /** 'move' carries the live offset from the model position; 'resize' has none; 'end' = the pointer
   * let go (the commit is still on its way); 'committed' = the commit's reply reached the app */
  kind: 'move' | 'resize' | 'end' | 'committed'
  dx?: number
  dy?: number
}

const EVENT = 'slides:node-gesture'

export function emitNodeGesture(detail: NodeGestureDetail): void {
  window.dispatchEvent(new CustomEvent<NodeGestureDetail>(EVENT, { detail }))
}

export function onNodeGesture(handler: (detail: NodeGestureDetail) => void): () => void {
  const listener = (e: Event) => handler((e as CustomEvent<NodeGestureDetail>).detail)
  window.addEventListener(EVENT, listener)
  return () => window.removeEventListener(EVENT, listener)
}
