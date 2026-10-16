import { describe, expect, it } from 'vitest'
import {
  addPicture,
  createBlankPptx,
  materializeSlide,
  openPptx,
  savePptx,
} from '@genoffice/pptx-engine'
import type { PictureElement } from '@genoffice/pptx-engine'
import { runTxn } from '../src/index'

const PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

async function fixture() {
  const made = await openPptx(await createBlankPptx())
  addPicture(made, made.deck.slides[0]!, {
    bytes: PNG,
    ext: 'png',
    offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
  })
  const opened = await openPptx(await savePptx(made))
  const pic = opened.deck.slides[0]!.elements.find((el) => el.type === 'picture') as PictureElement
  return { opened, target: { slide: 0, el: pic.id } }
}

function picture(opened: Awaited<ReturnType<typeof fixture>>['opened']) {
  return materializeSlide(opened, 0)!.elements.find((el) => el.type === 'picture') as PictureElement
}

describe('picture alt text and aspect lock ops', () => {
  it('writes alt text and the decorative flag, survives save/reopen, and keeps creationId', async () => {
    const { opened, target } = await fixture()
    expect(
      runTxn(opened, { ops: [{ op: 'setAltText', target, text: 'A "red" <box> & more\nline' }] })
        .applied,
    ).toBe(true)
    let reopened = await openPptx(await savePptx(opened))
    let xml = reopened.archive.readText('ppt/slides/slide1.xml')!
    expect(xml).toContain('descr="A &quot;red&quot; &lt;box&gt; &amp; more&#10;line"')
    expect(xml).toContain('a16:creationId')
    expect(picture(reopened).descr).toBe('A "red" <box> & more\nline')

    const t2 = { slide: 0, el: picture(reopened).id }
    expect(
      runTxn(reopened, { ops: [{ op: 'setAltText', target: t2, decorative: true }] }).applied,
    ).toBe(true)
    reopened = await openPptx(await savePptx(reopened))
    xml = reopened.archive.readText('ppt/slides/slide1.xml')!
    expect(xml).toContain('C183D7F6-B498-43B3-948B-1728B52AA6E4')
    expect(xml).toContain('a16:creationId')
    expect(xml).not.toContain('descr=')
    expect(picture(reopened).decorative).toBe(true)
    expect(picture(reopened).descr).toBeUndefined()

    const t3 = { slide: 0, el: picture(reopened).id }
    expect(runTxn(reopened, { ops: [{ op: 'setAltText', target: t3, text: '' }] }).applied).toBe(
      true,
    )
    reopened = await openPptx(await savePptx(reopened))
    expect(reopened.archive.readText('ppt/slides/slide1.xml')!).not.toContain('C183D7F6')
    expect(reopened.archive.readText('ppt/slides/slide1.xml')!).toContain('a16:creationId')
  })

  it('persists the aspect lock and rejects non-picture targets', async () => {
    const { opened, target } = await fixture()
    expect(runTxn(opened, { ops: [{ op: 'setLockAspect', target, locked: false }] }).applied).toBe(
      true,
    )
    let reopened = await openPptx(await savePptx(opened))
    expect(picture(reopened).lockAspect).toBe(false)
    const t2 = { slide: 0, el: picture(reopened).id }
    expect(
      runTxn(reopened, { ops: [{ op: 'setLockAspect', target: t2, locked: true }] }).applied,
    ).toBe(true)
    reopened = await openPptx(await savePptx(reopened))
    expect(picture(reopened).lockAspect).toBe(true)
    expect(reopened.archive.readText('ppt/slides/slide1.xml')!).toContain('noChangeAspect="1"')
    reopened = await openPptx(await savePptx(reopened))
    expect(picture(reopened).lockAspect).toBe(true)
    const bad = runTxn(reopened, {
      ops: [{ op: 'setLockAspect', target: { slide: 0, el: 'nope' }, locked: true }],
    })
    expect(bad.applied).toBe(false)
  })

  it('requires text or decorative', async () => {
    const { opened, target } = await fixture()
    const r = runTxn(opened, { ops: [{ op: 'setAltText', target }] })
    expect(r.applied).toBe(false)
    expect(r.failures?.[0]?.error).toContain('setAltText')
  })
})
