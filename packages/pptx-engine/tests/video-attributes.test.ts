import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import type { PictureElement } from '../src/types'

const theme: any = { colors: { accent1: '#4472C4', lt1: '#FFFFFF', dk1: '#000000' } }

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

describe('video attributes parsing (p14:media autoPlay, loop, trim)', () => {
  const picWithExt = (extXml: string) =>
    '<p:pic><p:nvPicPr><p:cNvPr id="4" name="movie.mp4"/><p:cNvPicPr/>' +
    `<p:nvPr><a:videoFile r:link="rId2"/>${extXml}</p:nvPr></p:nvPicPr>` +
    '<p:blipFill><a:blip r:embed="rId3"/></p:blipFill>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>'

  it('parses autoPlay, loop, and trim when present in p14:media', () => {
    const ext =
      '<p:extLst><p:ext uri="{DAA4B390-3A2F-4340-BE83-50A949437C0C}">' +
      '<p14:media r:embed="rId2" autoPlay="1" loop="1">' +
      '<p14:trim st="1200" end="5500"/>' +
      '</p14:media>' +
      '</p:ext></p:extLst>'

    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(picWithExt(ext)),
      ctx: {
        theme,
        mediaRels: new Map([['rId3', 'ppt/media/poster.png']]),
        avRels: new Map([['rId2', { target: 'ppt/media/video.mp4' }]]),
      },
    })

    const el = slide.elements[0] as PictureElement
    expect(el.type).toBe('picture')
    expect(el.media).toBeDefined()
    expect(el.media?.kind).toBe('video')
    expect(el.media?.target).toBe('ppt/media/video.mp4')
    expect(el.media?.autoPlay).toBe(true)
    expect(el.media?.loop).toBe(true)
    expect(el.media?.trim).toEqual({ startMs: 1200, endMs: 5500 })
  })

  it('handles video without p14:media attributes gracefully', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(picWithExt('')),
      ctx: {
        theme,
        avRels: new Map([['rId2', { target: 'ppt/media/video.mp4' }]]),
      },
    })

    const el = slide.elements[0] as PictureElement
    expect(el.media).toBeDefined()
    expect(el.media?.kind).toBe('video')
    expect(el.media?.target).toBe('ppt/media/video.mp4')
    expect(el.media?.autoPlay).toBeUndefined()
    expect(el.media?.loop).toBeUndefined()
    expect(el.media?.trim).toBeUndefined()
  })

  it('parses media when video is only referenced via p14:media embed without a:videoFile', () => {
    const picOnlyP14 =
      '<p:pic><p:nvPicPr><p:cNvPr id="5" name="media2.mp4"/><p:cNvPicPr/>' +
      '<p:nvPr><p:extLst><p:ext uri="{DAA4B390-3A2F-4340-BE83-50A949437C0C}">' +
      '<p14:media r:embed="rId5" autoPlay="true">' +
      '<p14:trim st="500"/>' +
      '</p14:media>' +
      '</p:ext></p:extLst></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId3"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>'

    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(picOnlyP14),
      ctx: {
        theme,
        avRels: new Map([['rId5', { target: 'ppt/media/clip.mp4' }]]),
      },
    })

    const el = slide.elements[0] as PictureElement
    expect(el.media).toBeDefined()
    expect(el.media?.kind).toBe('video')
    expect(el.media?.target).toBe('ppt/media/clip.mp4')
    expect(el.media?.autoPlay).toBe(true)
    expect(el.media?.trim).toEqual({ startMs: 500 })
  })
})
