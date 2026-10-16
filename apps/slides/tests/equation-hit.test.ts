// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { selectedEquationAtCanvasClick, trackEquationClick } from '../src/renderer/equation-hit'

describe('native equation double-click routing', () => {
  it('uses the selected equation even when the Konva canvas layer changes after selection', () => {
    const stage = document.createElement('div')
    const contentCanvas = document.createElement('canvas')
    const transformerCanvas = document.createElement('canvas')
    const overlay = document.createElement('div')
    stage.append(contentCanvas, transformerCanvas, overlay)
    expect(selectedEquationAtCanvasClick(stage, contentCanvas, 'equation-on-slide-1')).toBe(
      'equation-on-slide-1',
    )
    expect(selectedEquationAtCanvasClick(stage, transformerCanvas, 'equation-on-slide-1')).toBe(
      'equation-on-slide-1',
    )
    expect(selectedEquationAtCanvasClick(stage, transformerCanvas, null)).toBeNull()
    expect(selectedEquationAtCanvasClick(stage, overlay, 'equation-on-slide-1')).toBeNull()
    stage.remove()
  })

  it('opens only after the second nearby click on the same equation', () => {
    const first = trackEquationClick(null, { id: 'math-1', x: 200, y: 120, at: 100 }, false)
    expect(first.activate).toBe(false)
    const second = trackEquationClick(first.next, { id: 'math-1', x: 201, y: 120, at: 310 }, true)
    expect(second).toEqual({ next: null, activate: true })
    expect(
      trackEquationClick(second.next, { id: 'math-1', x: 201, y: 120, at: 320 }, true).activate,
    ).toBe(false)
  })

  it('does not edit math under a foreground shape selected by the first click', () => {
    const first = trackEquationClick(
      null,
      { id: 'math-under-shape', x: 200, y: 120, at: 100 },
      false,
    )
    // The rectangle still covers the canvas point, but the actual selected node is the foreground shape.
    const second = trackEquationClick(
      first.next,
      { id: 'math-under-shape', x: 200, y: 120, at: 280 },
      false,
    )
    expect(second).toEqual({ next: null, activate: false })
  })

  it('does not join distant, delayed, or different-equation clicks', () => {
    const first = { id: 'math-1', x: 200, y: 120, at: 100 }
    expect(
      trackEquationClick(first, { id: 'math-2', x: 200, y: 120, at: 200 }, true).activate,
    ).toBe(false)
    expect(
      trackEquationClick(first, { id: 'math-1', x: 220, y: 120, at: 200 }, true).activate,
    ).toBe(false)
    expect(
      trackEquationClick(first, { id: 'math-1', x: 200, y: 120, at: 700 }, true).activate,
    ).toBe(false)
  })
})
