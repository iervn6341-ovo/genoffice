/**
 * Resize gestures scale the node's whole Konva group (the Transformer works on scaleX/scaleY).
 * Text must not follow that scale: PowerPoint keeps the font size and only re-wraps, so while a
 * handle is dragged the glyphs have to stay unstretched. Between two layout previews (an IPC
 * round trip each) the group is stretched by the pointer; without a counter-scale the glyphs are
 * squashed/stretched for those frames and snap back when the re-laid-out text arrives — the
 * "text flickers while resizing" effect. The glyph layer (NodeBody names it) gets the inverse
 * scale instead, anchored at the box's top-left, so only fills/outlines stretch.
 */
import type Konva from 'konva'

export const GLYPH_LAYER_NAME = 'glyph-layer'

/**
 * Scale for the glyph layer so that group × layer stays at the net orientation +1 (the layer
 * counter-flips mirrored geometry already: a flipped group at rest has scale -1, layer -1).
 */
export function counterScale(groupScale: number): number {
  // a group squeezed to ~0 (handle dragged across the opposite edge) has no usable inverse
  if (!Number.isFinite(groupScale) || Math.abs(groupScale) < 1e-3) return 1
  return 1 / groupScale
}

/** Re-apply the counter-scale from the group's current scale (also restores the layer at rest). */
export function unstretchGlyphLayer(group: Konva.Group | null | undefined): void {
  if (!group) return
  const layers =
    typeof group.find === 'function'
      ? group.find<Konva.Group>(`.${GLYPH_LAYER_NAME}`)
      : typeof group.findOne === 'function'
        ? [group.findOne<Konva.Group>(`.${GLYPH_LAYER_NAME}`)].filter(
            (node): node is Konva.Group => Boolean(node),
          )
        : []
  if (!layers || layers.length === 0) return
  const nx = counterScale(group.scaleX())
  const ny = counterScale(group.scaleY())
  for (const layer of layers) {
    if (layer.scaleX() !== nx) layer.scaleX(nx)
    if (layer.scaleY() !== ny) layer.scaleY(ny)
  }
}
