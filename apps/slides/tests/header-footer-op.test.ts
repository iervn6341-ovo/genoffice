import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  getFirstSlideNumber,
  openPptx,
  readDeckHeaderFooter,
  readHeaderFooter,
  savePptx,
} from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'

const fixture = join(
  dirname(fileURLToPath(import.meta.url)),
  '../../../packages/pptx-engine/tests/fixtures/01_standard_business.pptx',
)

describe('Slides header/footer operation payload', () => {
  it('passes slideIndexes and date mode through the operation registry and saved file', async () => {
    const opened = await openPptx(readFileSync(fixture))
    expect(opened.deck.slides.length).toBeGreaterThan(1)
    const slidePaths = opened.deck.slides.map((slide) => slide.path)
    const current = runTxn(opened, {
      ops: [
        {
          op: 'applyHeaderFooter',
          slideIndexes: [1],
          settings: {
            footer: 'Current only',
            date: 'January 2, 2001',
            dateAuto: false,
          },
        },
      ],
    })
    expect(current.applied).toBe(true)
    let reopened = await openPptx(await savePptx(opened))
    expect(readHeaderFooter(reopened.deck.slides[1]!)).toMatchObject({
      footer: 'Current only',
      dateAuto: false,
    })
    expect(readHeaderFooter(reopened.deck.slides[0]!)).not.toMatchObject({ footer: 'Current only' })

    const all = runTxn(reopened, {
      ops: [
        {
          op: 'applyHeaderFooter',
          settings: {
            footer: 'Every slide',
            date: 'Tuesday',
            dateAuto: true,
            dateFormat: 'datetime4',
            slideNum: true,
            firstSlideNum: 0,
          },
        },
      ],
    })
    expect(all.applied).toBe(true)
    reopened = await openPptx(await savePptx(reopened))
    expect(reopened.deck.slides.map((slide) => slide.path)).toEqual(slidePaths)
    expect(getFirstSlideNumber(reopened.archive)).toBe(0)
    expect(readDeckHeaderFooter(reopened, 0).slideNumber).toBe(0)
    expect(readDeckHeaderFooter(reopened, 1).slideNumber).toBe(1)
    for (const slide of reopened.deck.slides) {
      expect(readHeaderFooter(slide)).toMatchObject({
        footer: 'Every slide',
        dateAuto: true,
        dateFormat: 'datetime4',
        slideNum: true,
      })
    }
  })
})
