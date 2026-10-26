import { describe, expect, it } from 'vitest'
import type Konva from 'konva'
import { counterScale, GLYPH_LAYER_NAME, unstretchGlyphLayer } from '../src/renderer/text-unstretch'

/** Minimal stand-ins for the two Konva groups involved (scale getters/setters + findOne). */
function fakeGroups(groupScale: { x: number; y: number }) {
  const layer = {
    sx: 1,
    sy: 1,
    scaleX(v?: number) {
      if (v === undefined) return this.sx
      this.sx = v
      return this.sx
    },
    scaleY(v?: number) {
      if (v === undefined) return this.sy
      this.sy = v
      return this.sy
    },
  }
  const group = {
    scaleX: () => groupScale.x,
    scaleY: () => groupScale.y,
    findOne: (sel: string) => (sel === `.${GLYPH_LAYER_NAME}` ? layer : undefined),
  }
  return { group: group as unknown as Konva.Group, layer }
}

describe('text-unstretch', () => {
  it('counterScale inverts the group scale so glyphs keep their size', () => {
    expect(counterScale(1)).toBe(1)
    expect(counterScale(2)).toBeCloseTo(0.5)
    expect(counterScale(0.5)).toBeCloseTo(2)
    // mirrored group (flipped shape / dragged through the opposite edge): the layer mirrors back
    expect(counterScale(-1)).toBe(-1)
    expect(counterScale(-1.86)).toBeCloseTo(-1 / 1.86)
  })

  it('counterScale ignores a collapsed or invalid scale instead of exploding', () => {
    expect(counterScale(0)).toBe(1)
    expect(counterScale(0.0004)).toBe(1)
    expect(counterScale(Number.NaN)).toBe(1)
  })

  it('applies group × layer = 1 on both axes through a stretch and back to rest', () => {
    const gs = { x: 1.86, y: 0.58 }
    const { group, layer } = fakeGroups(gs)
    unstretchGlyphLayer(group)
    expect(gs.x * layer.sx).toBeCloseTo(1)
    expect(gs.y * layer.sy).toBeCloseTo(1)
    gs.x = 1
    gs.y = 1
    unstretchGlyphLayer(group)
    expect(layer.sx).toBe(1)
    expect(layer.sy).toBe(1)
  })

  it('keeps a flipped node readable at rest (group -1 → layer -1)', () => {
    const { group, layer } = fakeGroups({ x: -1, y: 1 })
    unstretchGlyphLayer(group)
    expect(layer.sx).toBe(-1)
    expect(layer.sy).toBe(1)
  })

  it('unstretches multiple glyph layers (e.g. all cells of a table)', () => {
    const makeLayer = () => ({
      sx: 1,
      sy: 1,
      scaleX(v?: number) {
        if (v === undefined) return this.sx
        this.sx = v
        return this.sx
      },
      scaleY(v?: number) {
        if (v === undefined) return this.sy
        this.sy = v
        return this.sy
      },
    })
    const layer1 = makeLayer()
    const layer2 = makeLayer()
    const group = {
      scaleX: () => 2,
      scaleY: () => 0.5,
      find: (sel: string) => (sel === `.${GLYPH_LAYER_NAME}` ? [layer1, layer2] : []),
    } as unknown as Konva.Group
    unstretchGlyphLayer(group)
    expect(layer1.sx).toBeCloseTo(0.5)
    expect(layer1.sy).toBeCloseTo(2)
    expect(layer2.sx).toBeCloseTo(0.5)
    expect(layer2.sy).toBeCloseTo(2)
  })
})
