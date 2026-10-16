import { describe, expect, it } from 'vitest'
import {
  createBlankPptx,
  getShowSetup,
  openPptx,
  savePptx,
  setShowSetup,
  setSlideSize,
} from '../src/index'

describe('native slide show setup', () => {
  it('creates a valid presProps relationship and roundtrips kiosk/range options', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(getShowSetup(opened)).toEqual({
      mode: 'speaker',
      loop: false,
      showNarration: true,
      showAnimation: true,
      useTimings: true,
      slides: { kind: 'all' },
    })
    expect(
      setShowSetup(opened, {
        mode: 'kiosk',
        loop: true,
        showNarration: false,
        showAnimation: false,
        useTimings: false,
        slides: { kind: 'range', start: 1, end: 1 },
      }),
    ).toBe(true)
    expect(opened.archive.readText('ppt/_rels/presentation.xml.rels')).toContain(
      'relationships/presProps',
    )
    expect(opened.archive.readText('[Content_Types].xml')).toContain('presentationml.presProps+xml')
    const reopened = await openPptx(await savePptx(opened))
    expect(getShowSetup(reopened)).toMatchObject({
      mode: 'kiosk',
      loop: true,
      showNarration: false,
      showAnimation: false,
      useTimings: false,
      slides: { kind: 'range', start: 1, end: 1 },
    })
  })

  it('changes known choices while preserving imported extension children and attributes', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(
      setShowSetup(opened, {
        mode: 'speaker',
        loop: false,
        showNarration: true,
        showAnimation: true,
        useTimings: true,
        slides: { kind: 'all' },
      }),
    ).toBe(true)
    const path = 'ppt/presProps.xml'
    const original = opened.archive.readText(path)!
    opened.archive.entries.set(
      path,
      Buffer.from(
        original
          .replace('<p:showPr ', '<p:showPr customVendor="keep" ')
          .replace(
            '</p:showPr>',
            '<p:extLst><p:ext uri="vendor"><x:opaque xmlns:x="urn:keep"/></p:ext></p:extLst></p:showPr>',
          ),
        'utf8',
      ),
    )
    expect(
      setShowSetup(opened, {
        mode: 'browse',
        loop: true,
        showNarration: true,
        showAnimation: false,
        useTimings: false,
        slides: { kind: 'all' },
      }),
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('customVendor="keep"')
    expect(xml).toContain('<x:opaque xmlns:x="urn:keep"/>')
    expect(xml).toContain('<p:browse/>')
    expect(xml).not.toContain('<p:present/>')
  })

  it('expands an imported self-closing presentation-properties root without losing its attributes', async () => {
    const opened = await openPptx(await createBlankPptx())
    const defaults = getShowSetup(opened)
    expect(setShowSetup(opened, defaults)).toBe(true)
    const path = 'ppt/presProps.xml'
    const root =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:vendor" x:custom="keep"/>'
    opened.archive.entries.set(path, Buffer.from(root, 'utf8'))
    const presentation = opened.archive.readBytes('ppt/presentation.xml')
    const relationships = opened.archive.readBytes('ppt/_rels/presentation.xml.rels')
    const contentTypes = opened.archive.readBytes('[Content_Types].xml')
    const settings = {
      mode: 'browse' as const,
      loop: true,
      showNarration: false,
      showAnimation: true,
      useTimings: false,
      slides: { kind: 'range' as const, start: 1, end: 1 },
    }
    expect(setShowSetup(opened, settings)).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('xmlns:x="urn:vendor" x:custom="keep">')
    expect(xml).toContain('<p:showPr')
    expect(xml).toContain('<p:browse/><p:sldRg st="1" end="1"/>')
    expect(xml).toContain('</p:presentationPr>')
    expect(opened.archive.readBytes('ppt/presentation.xml')).toBe(presentation)
    expect(opened.archive.readBytes('ppt/_rels/presentation.xml.rels')).toBe(relationships)
    expect(opened.archive.readBytes('[Content_Types].xml')).toBe(contentTypes)
    const reopened = await openPptx(await savePptx(opened))
    expect(getShowSetup(reopened)).toEqual(settings)
    expect(reopened.archive.readText(path)).toContain('x:custom="keep"')
  })

  it('expands a self-closing imported showPr while keeping its unknown attributes', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(setShowSetup(opened, getShowSetup(opened))).toBe(true)
    const path = 'ppt/presProps.xml'
    opened.archive.entries.set(
      path,
      Buffer.from(
        opened.archive
          .readText(path)!
          .replace(/<p:showPr\b[\s\S]*?<\/p:showPr>/, '<p:showPr vendorFlag="keep"/>'),
        'utf8',
      ),
    )
    expect(
      setShowSetup(opened, {
        mode: 'kiosk',
        loop: true,
        showNarration: false,
        showAnimation: false,
        useTimings: false,
        slides: { kind: 'all' },
      }),
    ).toBe(true)
    const xml = opened.archive.readText(path)!
    expect(xml).toContain('vendorFlag="keep"')
    expect(xml).toContain('<p:kiosk/><p:sldAll/>')
    expect(xml).toContain('</p:showPr>')
  })

  it('rejects invalid slide ranges without changing the archive', async () => {
    const opened = await openPptx(await createBlankPptx())
    const before = new Map(opened.archive.entries)
    expect(
      setShowSetup(opened, {
        mode: 'speaker',
        loop: false,
        showNarration: true,
        showAnimation: true,
        useTimings: true,
        slides: { kind: 'range', start: 1, end: 2 },
      }),
    ).toBe(false)
    expect(opened.archive.entries).toEqual(before)
  })

  it('marks arbitrary slide dimensions as a native custom size', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(setSlideSize(opened, 10000000, 7000000)).toBe(true)
    const xml = opened.archive.readText('ppt/presentation.xml')!
    expect(xml).toMatch(/<p:sldSz\b[^>]*cx="10000000"[^>]*cy="7000000"[^>]*type="custom"/)
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.deck.size).toEqual({ cx: 10000000, cy: 7000000 })
  })
})
