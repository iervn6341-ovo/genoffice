/**
 * Geometry helpers of a live Transformer resize that the node view needs when the DRAWN size is not
 * the pointer's size (a table whose rows grew with re-wrapped text). The Konva node is pivoted on its
 * box centre, so changing its drawn size while the handle's opposite edge stays put moves the centre.
 */

/**
 * Shift of the centre when the drawn size goes from `from` to `to` while the edge opposite to the
 * dragged `anchor` ('top-left', 'middle-right', …) stays where it is. `rotDeg` rotates the shift into
 * stage space. A handle-less axis (the middle handles' other axis) keeps its centre.
 */
export function centerShift(
  anchor: string | undefined,
  rotDeg: number,
  from: { w: number; h: number },
  to: { w: number; h: number },
): { dx: number; dy: number } {
  const a = anchor ?? ''
  // dragging a left/top handle pins the right/bottom edge, and vice versa
  const lx = a.includes('left')
    ? (from.w - to.w) / 2
    : a.includes('right')
      ? (to.w - from.w) / 2
      : 0
  const ly = a.includes('top')
    ? (from.h - to.h) / 2
    : a.includes('bottom')
      ? (to.h - from.h) / 2
      : 0
  const th = (rotDeg * Math.PI) / 180
  return {
    dx: lx * Math.cos(th) - ly * Math.sin(th),
    dy: lx * Math.sin(th) + ly * Math.cos(th),
  }
}

/**
 * Height a table is drawn at while its handle is dragged: the pointer's height, unless the latest
 * layout (rows grown to fit re-wrapped text) is taller — then the table is that tall and the frame
 * follows the content instead of squashing it. `contentH` is 0 when the layout is not content-bound.
 */
export function drawnTableHeight(pointerH: number, contentH: number): number {
  return Math.max(pointerH, contentH)
}

/**
 * The layout a preview reply carries is content-bound when it came back taller than what was asked
 * for (the rows grew to fit their text): it then names the height the table cannot go under.
 */
export function contentBoundHeight(askedH: number | null | undefined, repliedH: number): number {
  return askedH != null && repliedH > askedH + 0.5 ? repliedH : 0
}
