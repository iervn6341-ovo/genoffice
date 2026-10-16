import { describe, expect, it } from 'vitest'
import { planCompression } from '../src/renderer/picture-tools-actions'

describe('picture quality planning', () => {
  it('downsamples to the target ppi for the frame size', () => {
    const plan = planCompression({ w: 4000, h: 2000 }, undefined, { w: 4, h: 2 }, 150, false)
    expect(plan).toMatchObject({ width: 600, height: 300, sx: 0, sy: 0, sw: 4000, sh: 2000 })
  })

  it('never upsamples and reports nothing to do for an already small image', () => {
    expect(planCompression({ w: 300, h: 150 }, undefined, { w: 4, h: 2 }, 150, false)).toBeNull()
  })

  it('scales the whole image when cropped pixels are kept', () => {
    // a 50% horizontal crop shown in 2in x 2in: the full image spans 4in x 2in
    const plan = planCompression(
      { w: 4000, h: 2000 },
      { l: 0.25, r: 0.25, t: 0, b: 0 },
      { w: 2, h: 2 },
      100,
      false,
    )
    expect(plan).toMatchObject({ width: 400, height: 200, sx: 0, sw: 4000 })
  })

  it('crops to the visible region when cropped areas are deleted', () => {
    const plan = planCompression(
      { w: 4000, h: 2000 },
      { l: 0.25, r: 0.25, t: 0, b: 0 },
      { w: 2, h: 2 },
      100,
      true,
    )
    expect(plan).toMatchObject({ sx: 1000, sy: 0, sw: 2000, sh: 2000, width: 200, height: 200 })
  })

  it('still deletes the crop of an image that is already small', () => {
    const plan = planCompression(
      { w: 400, h: 200 },
      { l: 0.5, r: 0, t: 0, b: 0 },
      { w: 2, h: 2 },
      100,
      true,
    )
    expect(plan).toMatchObject({ sx: 200, sw: 200, sh: 200, width: 200, height: 200 })
  })
})
