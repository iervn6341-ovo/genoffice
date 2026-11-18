import { describe, it, expect } from 'vitest'
import { openPptx, getSlideAnimations, savePptx } from '@genoffice/pptx-engine'
import { runTxn } from '../src/ops'
import fs from 'fs'

const deckExists = fs.existsSync('第四次估咪簡報.pptx')

describe('repro animation reorder and group child support', () => {
  it.skipIf(!deckExists)('tests moving animation 3 to position 0 on slide 11 of 第四次估咪簡報.pptx', async () => {
    const p = '第四次估咪簡報.pptx'
    const opened = await openPptx(fs.readFileSync(p))
    const s = opened.deck.slides[11]!
    const anims = getSlideAnimations(s, true)
    expect(anims).toHaveLength(4)

    // User moves anim 3 to the top
    const items = [
      { sourceId: 'sp_6d', effect: 'appear', trigger: 'onClick', durationMs: 0, delayMs: 0, nativeId: 17 },
      { sourceId: 'sp_6a', effect: 'appear', trigger: 'onClick', durationMs: 0, delayMs: 0, nativeId: 5, after: 'hideOnNextClick' },
      { sourceId: 'sp_6b', effect: 'appear', trigger: 'onClick', durationMs: 0, delayMs: 0, nativeId: 9, after: 'hideOnNextClick' },
      { sourceId: 'sp_6c', effect: 'appear', trigger: 'onClick', durationMs: 0, delayMs: 0, nativeId: 13, after: 'hideOnNextClick' },
    ]

    const txn = runTxn(opened, {
      ops: [{ op: 'setAnimations', target: { slide: 11 }, items }],
    })
    expect(txn.applied).toBe(true)
    expect(txn.failures).toBeUndefined()

    const savedBuf = await savePptx(opened)
    const reopened = await openPptx(savedBuf)
    const reSlide = reopened.deck.slides[11]!
    const reAnims = getSlideAnimations(reSlide, true)
    expect(reAnims.map((a) => a.spid)).toEqual([8, 5, 6, 7])
    expect(reAnims[0]!.after).toBeUndefined()
    expect(reAnims[1]!.after).toBe('hideOnNextClick')
    expect(reAnims[2]!.after).toBe('hideOnNextClick')
    expect(reAnims[3]!.after).toBe('hideOnNextClick')

    // animBg="1" is preserved on all bldP items
    const timing = reSlide.bodySuffix.match(/<p:timing>[\s\S]*?<\/p:timing>/)?.[0]
    expect(timing).toContain('<p:bldP spid="8" grpId="0" animBg="1"/>')
    expect(timing).toContain('<p:bldP spid="5" grpId="0" animBg="1"/>')
    expect(timing).toContain('<p:bldP spid="6" grpId="0" animBg="1"/>')
    expect(timing).toContain('<p:bldP spid="7" grpId="0" animBg="1"/>')
  }, 30000)

  it.skipIf(!deckExists)('supports animating group child elements', async () => {
    const p = '第四次估咪簡報.pptx'
    const opened = await openPptx(fs.readFileSync(p))
    const s = opened.deck.slides[0]!
    // Slide 0 has elements that can be grouped
    const el1 = s.elements[0]!.id
    const el2 = s.elements[1]!.id
    const groupTxn = runTxn(opened, {
      ops: [{ op: 'groupElements', target: { slide: 0 }, els: [el1, el2] }],
    })
    expect(groupTxn.applied).toBe(true)

    // After groupElements, read the updated slide from the deck
    const currentSlide = opened.deck.slides[0]!
    const grp = currentSlide.elements.find((e) => e.type === 'group') as any
    expect(grp).toBeDefined()
    expect(grp.children).toHaveLength(2)
    const child0 = grp.children[0].id
    const child1 = grp.children[1].id

    const animTxn = runTxn(opened, {
      ops: [{
        op: 'setAnimations',
        target: { slide: 0 },
        items: [
          { sourceId: child0, effect: 'fade', trigger: 'onClick', durationMs: 500, delayMs: 0 },
          { sourceId: child1, effect: 'flyIn', trigger: 'onClick', durationMs: 500, delayMs: 0 },
        ],
      }],
    })
    expect(animTxn.applied).toBe(true)
    expect(animTxn.failures).toBeUndefined()

    const savedBuf = await savePptx(opened)
    const reopened = await openPptx(savedBuf)
    const reSlide = reopened.deck.slides[0]!
    const reAnims = getSlideAnimations(reSlide)
    expect(reAnims).toHaveLength(2)
    expect(reAnims[0]!.effect).toBe('fade')
    expect(reAnims[1]!.effect).toBe('flyIn')
  }, 30000)
})
