/**
 * Keep a fixed-position popup menu fully inside the window, using its MEASURED size (a hard-coded
 * height guess drifts from the real one whenever padding, font or the item count changes).
 */
export function clampMenuToViewport(
  point: { x: number; y: number },
  size: { width: number; height: number },
  viewport: { width: number; height: number },
  margin = 4,
): { x: number; y: number } {
  const maxX = viewport.width - size.width - margin
  const maxY = viewport.height - size.height - margin
  return {
    x: Math.max(margin, Math.min(point.x, maxX)),
    y: Math.max(margin, Math.min(point.y, maxY)),
  }
}
