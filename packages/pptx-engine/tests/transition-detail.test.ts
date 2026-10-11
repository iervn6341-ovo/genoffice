/** Transition direction / duration / advance settings (Transitions tab: Effect Options, Duration, Advance). */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  getSlideTransition,
  getSlideTransitionDetail,
  openPptx,
  savePptx,
  setSlideAdvanceTime,
  setSlideTransition,
  setSlideTransitionDetail,
} from '../src/index'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))
const open = async () => (await openPptx(fx('01_standard_business.pptx'))).deck.slides[0]!

describe('transition detail', () => {
  it('a freshly applied effect starts with PowerPoint defaults', async () => {
    const slide = await open()
    setSlideTransition(slide, 'push')
    expect(getSlideTransitionDetail(slide)).toMatchObject({
      kind: 'push',
      dir: 'u',
      durationMs: null,
      advClick: true,
      advTm: null,
    })
    // a bare effect stays as bare as before: no AlternateContent, no attributes
    expect(slide.bodySuffix).toContain('<p:transition><p:push dir="u"/></p:transition>')
  })

  it('direction changes only the effect element', async () => {
    const slide = await open()
    setSlideTransition(slide, 'cover')
    setSlideTransitionDetail(slide, { dir: 'rd' })
    expect(slide.bodySuffix).toContain('<p:cover dir="rd"/>')
    expect(getSlideTransition(slide)).toBe('cover')
    // a direction the effect does not know falls back to its default
    setSlideTransitionDetail(slide, { dir: 'in' })
    expect(slide.bodySuffix).toContain('<p:cover dir="l"/>')
  })

  it('a duration wraps the effect in AlternateContent with p14:dur and the matching speed word', async () => {
    const slide = await open()
    setSlideTransition(slide, 'wipe')
    setSlideTransitionDetail(slide, { durationMs: 1200 })
    expect(slide.bodySuffix).toContain('Requires="p14"')
    expect(slide.bodySuffix).toContain('<p:transition spd="slow" p14:dur="1200"><p:wipe dir="l"/>')
    expect(slide.bodySuffix).toContain('<mc:Fallback><p:transition spd="slow"><p:wipe dir="l"/>')
    expect(getSlideTransitionDetail(slide).durationMs).toBe(1200)
    expect(getSlideTransition(slide)).toBe('wipe')
    setSlideTransitionDetail(slide, { durationMs: 400 })
    expect(slide.bodySuffix).toContain('spd="fast" p14:dur="400"')
    // clearing returns to the bare form
    setSlideTransitionDetail(slide, { durationMs: null })
    expect(slide.bodySuffix).not.toContain('<mc:AlternateContent')
    expect(getSlideTransitionDetail(slide).durationMs).toBeNull()
  })

  it('on-click off and auto-advance time live on the transition element', async () => {
    const slide = await open()
    setSlideTransition(slide, 'fade')
    setSlideTransitionDetail(slide, { advClick: false })
    setSlideAdvanceTime(slide, 4500)
    const d = getSlideTransitionDetail(slide)
    expect(d).toMatchObject({ kind: 'fade', advClick: false, advTm: 4500 })
    expect(slide.bodySuffix).toMatch(
      /<p:transition [^>]*advClick="0"[^>]*advTm="4500"|<p:transition [^>]*advTm="4500"[^>]*advClick="0"/,
    )
    // the auto-advance time survives a direction / duration patch and an effect change
    setSlideTransitionDetail(slide, { durationMs: 900 })
    expect(getSlideTransitionDetail(slide)).toMatchObject({
      advClick: false,
      advTm: 4500,
      durationMs: 900,
    })
    setSlideTransition(slide, 'push')
    expect(getSlideTransitionDetail(slide)).toMatchObject({
      kind: 'push',
      dir: 'u',
      advClick: false,
      advTm: 4500,
      durationMs: 900,
    })
  })

  it('split takes orientation + in/out, fade takes through-black', async () => {
    const slide = await open()
    setSlideTransition(slide, 'split')
    setSlideTransitionDetail(slide, { orient: 'vert', dir: 'in' })
    expect(slide.bodySuffix).toContain('<p:split orient="vert" dir="in"/>')
    expect(getSlideTransitionDetail(slide)).toMatchObject({ orient: 'vert', dir: 'in' })
    setSlideTransition(slide, 'fade')
    setSlideTransitionDetail(slide, { thruBlk: true })
    expect(slide.bodySuffix).toContain('<p:fade thruBlk="1"/>')
    expect(getSlideTransitionDetail(slide).thruBlk).toBe(true)
    setSlideTransitionDetail(slide, { thruBlk: false })
    expect(slide.bodySuffix).toContain('<p:fade/>')
  })

  it('morph keeps its own duration and takes the flags', async () => {
    const slide = await open()
    setSlideTransition(slide, 'morph')
    expect(getSlideTransitionDetail(slide)).toMatchObject({ kind: 'morph', durationMs: 800 })
    setSlideTransitionDetail(slide, { durationMs: 1500, advClick: false })
    expect(getSlideTransitionDetail(slide)).toMatchObject({
      kind: 'morph',
      durationMs: 1500,
      advClick: false,
    })
    expect(slide.bodySuffix).toContain('<p159:morph option="byObject"/>')
    // leaving morph does not drag its duration along
    setSlideTransition(slide, 'zoom')
    expect(getSlideTransitionDetail(slide).durationMs).toBeNull()
  })

  it('a slide without an effect only takes the on-click flag (and a PowerPoint-only effect survives)', async () => {
    const slide = await open()
    setSlideTransition(slide, 'none')
    setSlideTransitionDetail(slide, { durationMs: 900, dir: 'u' })
    expect(slide.bodySuffix).not.toContain('<p:transition')
    setSlideTransitionDetail(slide, { advClick: false })
    expect(slide.bodySuffix).toContain('<p:transition advClick="0"/>')
    expect(getSlideTransitionDetail(slide)).toMatchObject({ kind: 'none', advClick: false })
    setSlideTransitionDetail(slide, { advClick: true })
    expect(slide.bodySuffix).not.toContain('advClick')

    // an effect we do not model (here p14:vortex) must not be rewritten by a duration patch
    slide.bodySuffix = slide.bodySuffix.replace(
      '</p:clrMapOvr>',
      '</p:clrMapOvr><p:transition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="slow" p14:dur="2000"><p14:vortex dir="r"/></p:transition>',
    )
    const before = slide.bodySuffix
    setSlideTransitionDetail(slide, { durationMs: 700 })
    expect(slide.bodySuffix).toBe(before)
  })

  it('survives save and reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide = opened.deck.slides[0]!
    setSlideTransition(slide, 'push')
    setSlideTransitionDetail(slide, { dir: 'l', durationMs: 1000, advClick: false })
    setSlideAdvanceTime(slide, 3000)
    const reopened = await openPptx(await savePptx(opened))
    expect(getSlideTransitionDetail(reopened.deck.slides[0]!)).toMatchObject({
      kind: 'push',
      dir: 'l',
      durationMs: 1000,
      advClick: false,
      advTm: 3000,
    })
  })
})
