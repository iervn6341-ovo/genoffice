/** Konva's first click can mount the Transformer and change its hit node before
 * the second. Selection is resolved by Konva itself before the native click. */
export function selectedEquationAtCanvasClick(
  stage: HTMLElement,
  target: EventTarget | null,
  selectedEquationId: string | null,
): string | null {
  if (!(target instanceof HTMLCanvasElement) || !stage.contains(target)) return null
  return selectedEquationId
}

export interface EquationClick {
  id: string
  x: number
  y: number
  at: number
}

/** A double-click should survive Konva hit-node and browser canvas-layer changes. */
export function trackEquationClick(
  previous: EquationClick | null,
  current: EquationClick,
  equationSelected: boolean,
): { next: EquationClick | null; activate: boolean } {
  const matched =
    previous?.id === current.id &&
    current.at >= previous.at &&
    current.at - previous.at <= 500 &&
    Math.hypot(current.x - previous.x, current.y - previous.y) <= 8
  if (matched) return { next: null, activate: equationSelected }
  return { next: current, activate: false }
}
