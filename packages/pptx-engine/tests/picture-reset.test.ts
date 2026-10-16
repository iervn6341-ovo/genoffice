import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import {
  addPicture,
  copyElementData,
  createBlankPptx,
  editPictureSrcRect,
  getElementLink,
  openPptx,
  replacePictureBytes,
  resetPictureFromBaseline,
  savePptx,
  setPictureAdjustments,
  setElementLink,
} from '../src/index'
import type { PictureElement } from '../src/types'

const PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)
const GIF = Uint8Array.from(
  Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64'),
)

async function fixture() {
  const made = await openPptx(await createBlankPptx())
  addPicture(made, made.deck.slides[0]!, {
    bytes: PNG,
    ext: 'png',
    offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
  })
  const zip = await JSZip.loadAsync(await savePptx(made))
  const path = 'ppt/slides/slide1.xml'
  const xml = await zip.file(path)!.async('string')
  zip.file(
    path,
    xml.replace(
      /(<a:blip\b[^>]*)(?:\/>|><\/a:blip>)/,
      '$1><a:alphaModFix amt="70000"/><a:extLst data-reset-keep="yes"/></a:blip>',
    ),
  )
  const opened = await openPptx(await zip.generateAsync({ type: 'uint8array' }))
  const slide = opened.deck.slides[0]!
  const pic = slide.elements.find((el) => el.type === 'picture') as PictureElement
  return { opened, slide, pic }
}

describe('Reset Picture native saved baseline', () => {
  it('ignores frame-only changes without touching archive parts or history-worthy state', async () => {
    const { opened, slide, pic } = await fixture()
    const baseline = copyElementData(opened, slide, pic)
    pic.transform.offset.x = 2000000
    pic.dirtyTransform = true
    const entriesBefore = new Map(opened.archive.entries)
    expect(resetPictureFromBaseline(opened, 0, pic.id, baseline)).toBe(false)
    expect(opened.archive.entries).toEqual(entriesBefore)
    expect(pic.dirtyTransform).toBe(true)
  })

  it('restores image, crop, corrections and unknown XML while retaining the current frame and nonvisual metadata', async () => {
    const { opened, slide, pic } = await fixture()
    const baseline = copyElementData(opened, slide, pic)
    expect(baseline.parts?.['ppt/media/image1.png']).toBeTruthy()
    expect(
      setPictureAdjustments(slide, pic.id, {
        lum: { bright: 0.25, contrast: 0.1 },
        color: 'sepia',
      }),
    ).toBe(true)
    expect(editPictureSrcRect(slide, pic.id, { l: 0.1, t: 0, r: 0, b: 0 })).toBe(true)
    expect(replacePictureBytes(opened, slide, pic.id, GIF, 'gif')).toBe(true)
    pic.transform.offset.x = 2000000
    pic.dirtyTransform = true
    pic.anchor.originalXml =
      pic.anchor.originalXml?.replace('name="Picture 2"', 'name="Retitled picture"') ?? ''
    expect(resetPictureFromBaseline(opened, 0, pic.id, baseline)).toBe(true)
    const bytes = await savePptx(opened)
    const out = await JSZip.loadAsync(bytes)
    const xml = await out.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toContain('name="Retitled picture"')
    expect(xml).toContain('x="2000000"')
    expect(xml).toContain('data-reset-keep="yes"')
    expect(xml).toContain('<a:alphaModFix amt="70000"/>')
    expect(xml).not.toContain('<a:lum ')
    expect(xml).not.toContain('<a:duotone>')
    expect(xml).not.toContain('<a:srcRect')
    expect(xml).toMatch(/<a:blip r:embed="rId\d+"/)
    const reopened = await openPptx(bytes)
    const restored = reopened.deck.slides[0]!.elements.find(
      (el) => el.type === 'picture',
    ) as PictureElement
    expect(restored.mediaRef).toBe('ppt/media/image1.png')
    expect(restored.opacity).toBeCloseTo(0.7)
    expect(restored.transform.offset.x).toBe(2000000)
    expect(await out.file('ppt/media/image1.png')?.async('uint8array')).toEqual(PNG)
  })

  it('retains a real external hyperlink relationship through media replacement and reset', async () => {
    const { opened, slide, pic } = await fixture()
    const baseline = copyElementData(opened, slide, pic)
    const target = 'https://example.org/picture?state=saved'
    const linkedSlide = setElementLink(opened, 0, pic.id, { kind: 'url', url: target })!
    const linked = linkedSlide.elements.find((item) => item.type === 'picture')!
    const hrefRid = /<a:hlinkClick[^>]*r:id="(rId\d+)"/.exec(linked.anchor.originalXml)?.[1]
    expect(hrefRid).toBeTruthy()
    expect(replacePictureBytes(opened, linkedSlide, linked.id, GIF, 'gif')).toBe(true)
    expect(setPictureAdjustments(linkedSlide, linked.id, { color: 'grayscale' })).toBe(true)
    expect(resetPictureFromBaseline(opened, 0, linked.id, baseline)).toBe(true)
    const saved = await savePptx(opened)
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
    expect(xml).toContain(`r:id="${hrefRid}"`)
    expect(rels).toContain(`Id="${hrefRid}"`)
    expect(rels).toContain('TargetMode="External"')
    expect(rels).toContain('Target="https://example.org/picture?state=saved"')
    const reopened = await openPptx(saved)
    const restored = reopened.deck.slides[0]!.elements.find((item) => item.type === 'picture')!
    expect(getElementLink(reopened, 0, restored.id)).toEqual({ kind: 'url', url: target })
  })
})
