/**
 * "After Animation" (hide on next click / hide after animation): read from PowerPoint's
 * <p:subTnLst>, written back by the engine, and kept stable across a read → write → read cycle.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import {
  buildTimingXml,
  patchSlideTimingXml,
  readSlideTimingXml,
  type SlideAnimation,
} from '../src/animation'

// slide 12 of a real PowerPoint deck: four Appear effects; all but the last are hidden when the next click starts
const REAL = readFileSync(
  new URL('./fixtures/appear-hide-on-next-click-timing.xml', import.meta.url),
  'utf8',
)

describe('After Animation', () => {
  it('reads hide-on-next-click from a real PowerPoint timeline', () => {
    const anims = readSlideTimingXml(`<p:sld>${REAL}</p:sld>`)
    expect(anims.map((a) => a.after)).toEqual([
      'hideOnNextClick',
      'hideOnNextClick',
      'hideOnNextClick',
      undefined, // the last item stays on screen
    ])
    expect(anims.map((a) => a.effect)).toEqual(['appear', 'appear', 'appear', 'appear'])
  })

  it('writes the setting and reads it back unchanged', () => {
    const anims: SlideAnimation[] = [
      {
        spid: 4,
        effect: 'appear',
        trigger: 'onClick',
        durationMs: 0,
        delayMs: 0,
        after: 'hideOnNextClick',
      },
      {
        spid: 5,
        effect: 'fade',
        trigger: 'onClick',
        durationMs: 500,
        delayMs: 0,
        after: 'hideAfter',
      },
      { spid: 6, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    ]
    const xml = buildTimingXml(anims)
    expect(xml).toContain('masterRel="nextClick"')
    const back = readSlideTimingXml(`<p:sld>${xml}</p:sld>`)
    expect(back.map((a) => a.after)).toEqual(['hideOnNextClick', 'hideAfter', undefined])
  })

  it('a rebuild of an unchanged list keeps every hide setting', () => {
    const suffix = `<p:sld>${REAL}</p:sld>`
    const anims = readSlideTimingXml(suffix)
    const patched = patchSlideTimingXml(suffix, anims)
    expect(readSlideTimingXml(patched).map((a) => a.after)).toEqual(anims.map((a) => a.after))
  })
})
