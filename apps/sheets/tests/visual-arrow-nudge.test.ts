import { describe, expect, it } from 'vitest'

import { ARROW_NUDGE, ARROW_NUDGE_PX } from '../src/renderer/WorkbookVisuals'

describe('arrow-key nudge of a selected object', () => {
  it('maps exactly the four arrows to unit directions', () => {
    expect(Object.keys(ARROW_NUDGE).sort()).toEqual([
      'ArrowDown',
      'ArrowLeft',
      'ArrowRight',
      'ArrowUp',
    ])
    expect(ARROW_NUDGE['ArrowLeft']).toEqual([-1, 0])
    expect(ARROW_NUDGE['ArrowDown']).toEqual([0, 1])
    expect(ARROW_NUDGE['Enter']).toBeUndefined()
  })

  it('a plain arrow moves further than the one-pixel Option nudge', () => {
    expect(ARROW_NUDGE_PX).toBeGreaterThan(1)
  })
})
