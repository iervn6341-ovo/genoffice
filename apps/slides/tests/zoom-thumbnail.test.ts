import { describe, expect, it } from 'vitest'
import { createBlankPptx, openPptx, savePptx } from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'

const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

describe('slide zoom thumbnail ops', () => {
  it('adds a described picture and links it to the target slide, surviving save/reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(runTxn(opened, { ops: [{ op: 'addBlankSlide', target: { slide: 0 } }] }).applied).toBe(
      true,
    )
    const added = runTxn(opened, {
      ops: [
        {
          op: 'addPicture',
          target: { slide: 0 },
          bytes: PNG,
          ext: 'png',
          name: 'Slide Zoom 2',
          descr: 'Slide Zoom: slide 2',
          offset: { x: 100000, y: 100000, cx: 2000000, cy: 1200000 },
        },
      ],
    })
    expect(added.applied).toBe(true)
    const id = added.records![0]!.created![0]!
    const linked = runTxn(opened, {
      ops: [
        {
          op: 'setLink',
          target: { slide: 0, el: id },
          link: { kind: 'slide', slideIndex: 1 },
        },
      ],
    })
    expect(linked.applied).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const xml = reopened.archive.readText('ppt/slides/slide1.xml')!
    expect(xml).toContain('action="ppaction://hlinksldjump"')
    expect(xml).toContain('descr="Slide Zoom: slide 2"')
  })
})
