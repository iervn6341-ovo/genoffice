import { describe, expect, it } from 'vitest'
import {
  addPicture,
  createBlankPptx,
  materializeSlide,
  openPptx,
  savePptx,
  type PictureElement,
} from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'
import { PICTURE_STYLES, pictureStyleOps } from '../src/renderer/picture-styles'

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
  return { opened, id: pic.id }
}

describe('picture styles', () => {
  it('applies every style as one atomic transaction that survives save and reopen', async () => {
    for (const style of PICTURE_STYLES) {
      const { opened, id } = await fixture()
      const result = runTxn(opened, { ops: pictureStyleOps(0, id, style) })
      expect(result.failures, style.id).toBeUndefined()
      expect(result.applied, style.id).toBe(true)
      const reopened = await openPptx(await savePptx(opened))
      const xml = reopened.archive.readText('ppt/slides/slide1.xml')!
      expect(xml).toContain(`<a:prstGeom prst="${style.prst}">`)
      const pic = materializeSlide(reopened, 0)!.elements.find(
        (el) => el.type === 'picture',
      ) as PictureElement
      expect(pic.presetGeometry ?? 'rect').toBe(style.prst)
      expect(/<a:effectLst>/.test(xml), style.id).toBe(
        !!(style.effects.shadow || style.effects.reflection || style.effects.softEdge),
      )
      expect(/<a:ln\b[^>]*>\s*<a:solidFill>/.test(xml), style.id).toBe(!!style.stroke)
    }
  })

  it('switching styles clears the previous style effects', async () => {
    const { opened, id } = await fixture()
    const byId = (key: string) => PICTURE_STYLES.find((s) => s.id === key)!
    expect(runTxn(opened, { ops: pictureStyleOps(0, id, byId('simple-frame')) }).applied).toBe(true)
    expect(runTxn(opened, { ops: pictureStyleOps(0, id, byId('none')) }).applied).toBe(true)
    const reopened = await openPptx(await savePptx(opened))
    const xml = reopened.archive.readText('ppt/slides/slide1.xml')!
    expect(xml).not.toContain('<a:effectLst>')
    expect(xml).toContain('<a:prstGeom prst="rect">')
  })
})
