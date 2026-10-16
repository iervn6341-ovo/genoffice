import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  addPicture,
  createBlankPptx,
  openPptx,
  savePptx,
  setPictureAdjustments,
} from '../src/index'
import type { PictureElement } from '../src/types'

const PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

async function importedPicture() {
  const created = await openPptx(await createBlankPptx())
  addPicture(created, created.deck.slides[0]!, {
    bytes: PNG,
    ext: 'png',
    offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
  })
  const zip = await JSZip.loadAsync(await savePptx(created))
  const path = 'ppt/slides/slide1.xml'
  const xml = await zip.file(path)!.async('string')
  expect(xml).toContain('<a:blip r:embed=')
  zip.file(
    path,
    xml.replace(
      /(<a:blip\b[^>]*)(?:\/>|><\/a:blip>)/,
      '$1><a:alphaModFix amt="75000"/><a:blur rad="25000"/><a:extLst data-preserve="yes"/></a:blip>',
    ),
  )
  const opened = await openPptx(await zip.generateAsync({ type: 'uint8array' }))
  const slide = opened.deck.slides[0]!
  const picture = slide.elements.find((el) => el.type === 'picture') as PictureElement
  return { opened, slide, picture }
}

describe('native picture corrections and color', () => {
  it('edits only requested a:blip effects and retains media, alpha, blur and extensions', async () => {
    const { opened, slide, picture } = await importedPicture()
    const id = picture.id
    expect(setPictureAdjustments(slide, id, { lum: { bright: 0.2, contrast: -0.1 } })).toBe(true)
    expect(setPictureAdjustments(slide, id, { color: 'sepia' })).toBe(true)
    const zip = await JSZip.loadAsync(await savePptx(opened))
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toContain('<a:alphaModFix amt="75000"/>')
    expect(xml).toContain('<a:blur rad="25000"/>')
    expect(xml).toContain('<a:extLst data-preserve="yes"/>')
    expect(xml).toMatch(
      /<a:lum bright="20000" contrast="-10000"\/><a:duotone>[\s\S]*?<\/a:duotone><a:extLst/,
    )
    expect((await zip.file('ppt/media/image1.png')!.async('uint8array')).length).toBe(PNG.length)

    const reopened = await openPptx(await savePptx(opened))
    const pic = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'picture',
    ) as PictureElement
    expect(pic.lum).toEqual({ bright: 0.2, contrast: -0.1 })
    expect(pic.duotone).toEqual(['#704214', '#F5DEB3'])
    expect(pic.opacity).toBeCloseTo(0.75)
    expect(xml).toContain('<p:cNvPr id="2" name="Picture 2"')
  })

  it('recolors, then clears modeled effects without touching extensions', async () => {
    const { opened, slide, picture } = await importedPicture()
    expect(setPictureAdjustments(slide, picture.id, { color: 'blackWhite' })).toBe(true)
    expect(picture.biLevel).toBe(0.5)
    expect(setPictureAdjustments(slide, picture.id, { color: 'original', lum: null })).toBe(true)
    const zip = await JSZip.loadAsync(await savePptx(opened))
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).not.toContain('<a:biLevel')
    expect(xml).not.toContain('<a:lum')
    expect(xml).toContain('data-preserve="yes"')
    const reopened = await openPptx(await savePptx(opened))
    const pic = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'picture',
    ) as PictureElement
    expect(pic.biLevel).toBeUndefined()
    expect(pic.lum).toBeUndefined()
  })

  it('expands a self-closing image blip when applying corrections', async () => {
    const created = await openPptx(await createBlankPptx())
    addPicture(created, created.deck.slides[0]!, {
      bytes: PNG,
      ext: 'png',
      offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
    })
    const opened = await openPptx(await savePptx(created))
    const slide = opened.deck.slides[0]!
    const pic = slide.elements.find((el) => el.type === 'picture') as PictureElement
    expect(setPictureAdjustments(slide, pic.id, { color: 'grayscale' })).toBe(true)
    const zip = await JSZip.loadAsync(await savePptx(opened))
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toMatch(
      /<a:blip r:embed="rId2"><a:grayscl\/><\/a:blip>/,
    )
  })
})
