/** Master/layout editing: enumeration/parsing, byte-surgery write-back, inheritance taking effect (6.16). */
import { describe, it, expect } from 'vitest'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import {
  openPptx,
  savePptx,
  patchSlideXml,
  listMasterParts,
  parseMasterPart,
  materializeSlide,
  createBlankPptx,
  ensureNotesMaster,
  ensureHandoutMaster,
  masterPartSize,
  readNotesMasterPlaceholders,
  setNotesMasterPlaceholderEnabled,
  setNotesPageOrientation,
  setNotesMasterBodyStyle,
  getSlideNotesParagraphs,
  setSlideNotes,
} from '../src/index'
import type { TextElement } from '../src/types'

const here = dirname(fileURLToPath(import.meta.url))
const fx = (name: string) => readFileSync(join(here, 'fixtures', name))

describe('listMasterParts / parseMasterPart', () => {
  it('writes Notes Master body formatting into native notesStyle inherited by notes pages', async () => {
    const opened = await openPptx(await createBlankPptx())
    const path = ensureNotesMaster(opened.archive)!
    setSlideNotes(opened, 0, 'Body inherits style')
    expect(
      setNotesMasterBodyStyle(opened.archive, path, {
        fontFamily: 'Arial',
        fontSizePt: 20,
        color: '#336699',
        bold: true,
        align: 'right',
        lineSpacingPct: 145,
      }),
    ).toBe(true)
    const native = opened.archive.readText(path)!
    expect(native).toContain('<p:notesStyle><a:lvl1pPr algn="r">')
    expect(native).toContain('<a:lnSpc><a:spcPct val="145000"/></a:lnSpc>')
    const paragraph = getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)[0]!
    expect(paragraph.align).toBe('right')
    expect(paragraph.runs[0]).toMatchObject({
      text: 'Body inherits style',
      fontFamily: 'Arial',
      fontSize: 20,
      color: '#336699',
      bold: true,
    })
    const reopened = await openPptx(await savePptx(opened))
    expect(
      getSlideNotesParagraphs(reopened.archive, reopened.deck.slides[0]!.path)[0]?.runs[0]
        ?.fontSize,
    ).toBe(20)
  })
  it('parses imported decorative picture and gradient background from a notes master for print artwork', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const path = ensureNotesMaster(opened.archive)!
    const xml = opened.archive.readText(path)!
    const gradient =
      '<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs><a:gs pos="100000"><a:srgbClr val="EEDDCC"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>'
    const picture =
      '<p:pic><p:nvPicPr><p:cNvPr id="99" name="Native logo"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId99"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="98" name="Native Logo Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="914400" y="365760"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${picture}</p:grpSp>`
    opened.archive.entries.set(
      path,
      Buffer.from(
        xml
          .replace('<p:cSld>', `<p:cSld>${gradient}`)
          .replace('</p:spTree>', `${group}</p:spTree>`),
      ),
    )
    const relsPath = 'ppt/notesMasters/_rels/notesMaster1.xml.rels'
    const rels = opened.archive.readText(relsPath)!
    opened.archive.entries.set(
      relsPath,
      Buffer.from(
        rels.replace(
          '</Relationships>',
          '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>',
        ),
      ),
    )
    const master = parseMasterPart(opened.archive, path)!
    expect(master.background?.type).toBe('gradient')
    const logo = master.elements
      .find((element) => element.type === 'group')
      ?.children.find((element) => element.type === 'picture')
    expect(logo).toBeDefined()
    expect(JSON.stringify(logo)).toContain('ppt/media/image1.png')
  })
  it('lists native notes and handout masters on their notes canvas with editable placeholders', async () => {
    const opened = await openPptx(await createBlankPptx())
    const notesPath = ensureNotesMaster(opened.archive)!
    const handoutPath = ensureHandoutMaster(opened.archive)
    const parts = listMasterParts(opened.archive)
    expect(parts.some((p) => p.kind === 'notesMaster' && p.partPath === notesPath)).toBe(true)
    expect(parts.some((p) => p.kind === 'handoutMaster' && p.partPath === handoutPath)).toBe(true)
    const notesSize = masterPartSize(opened.archive, notesPath, opened.deck.size)
    expect(notesSize).toMatchObject({ cx: 6858000, cy: 9144000 })
    const notes = parseMasterPart(opened.archive, notesPath)!
    expect(notes.elements.some((e) => e.placeholder === 'sldImg')).toBe(true)
    expect(notes.elements.some((e) => e.placeholder === 'body')).toBe(true)
    expect(readNotesMasterPlaceholders(opened.archive, notesPath)).toEqual({
      sldImg: true,
      body: true,
    })
    const originalBody = opened.archive
      .readText(notesPath)!
      .match(/<p:sp\b[^>]*>[\s\S]*?<p:ph type="body"[\s\S]*?<\/p:sp>/)?.[0]
    expect(setNotesMasterPlaceholderEnabled(opened.archive, notesPath, 'body', false)).toBe(true)
    expect(readNotesMasterPlaceholders(opened.archive, notesPath)?.body).toBe(false)
    expect(setNotesMasterPlaceholderEnabled(opened.archive, notesPath, 'body', true)).toBe(true)
    expect(readNotesMasterPlaceholders(opened.archive, notesPath)?.body).toBe(true)
    expect(originalBody).toContain('idx="3"')
    expect(parseMasterPart(opened.archive, handoutPath)).not.toBeNull()
    expect(setNotesPageOrientation(opened.archive, 'landscape')).toBe(true)
    expect(masterPartSize(opened.archive, notesPath, opened.deck.size)).toEqual({
      cx: 9144000,
      cy: 6858000,
    })
    const reopened = await openPptx(await savePptx(opened))
    expect(listMasterParts(reopened.archive).map((p) => p.kind)).toContain('notesMaster')
  })
  it('enumeration: master first, layouts after; both part kinds parse into elements', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const parts = listMasterParts(opened.archive)
    expect(parts[0]!.kind).toBe('master')
    expect(parts.filter((p) => p.kind === 'layout').length).toBeGreaterThan(1)

    const master = parseMasterPart(opened.archive, parts[0]!.partPath)!
    expect(master.elements.length).toBeGreaterThan(0)
    expect(master.elements.some((e) => e.placeholder === 'title')).toBe(true)

    const layout = parseMasterPart(opened.archive, parts[1]!.partPath)!
    expect(layout.elements.length).toBeGreaterThan(0)
    expect(layout.elements.every((e) => e.transform.offset.cx >= 0)).toBe(true)
  })
})

describe('slide master edit write-back (byte surgery)', () => {
  it('edit master title placeholder text: dirty element regenerated, other elements pass through byte-for-byte, persists after savePptx', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const parts = listMasterParts(opened.archive)
    const masterPath = parts[0]!.partPath
    const layoutPath = parts[1]!.partPath
    const layoutBefore = opened.archive.readText(layoutPath)!
    const master = parseMasterPart(opened.archive, masterPath)!

    const title = master.elements.find((e) => e.placeholder === 'title') as TextElement
    title.text!.paragraphs[0]!.runs[0]!.text = 'MASTER-EDITED'
    title.dirty = true

    const patched = patchSlideXml(master)
    expect(patched).toContain('MASTER-EDITED')
    // Untouched elements preserved byte-for-byte
    for (const el of master.elements) {
      if (el === title) continue
      expect(patched).toContain(el.anchor.originalXml)
    }
    opened.archive.entries.set(masterPath, Buffer.from(patched, 'utf8'))

    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.archive.readText(masterPath)).toContain('MASTER-EDITED')
    // Unedited layout parts are byte-identical
    expect(reopened.archive.readText(layoutPath)).toBe(layoutBefore)
  })

  it('edit layout placeholder geometry: slide placeholders without explicit xfrm inherit the new geometry, still there after save and reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const slide0 = opened.deck.slides[0]!
    const slideTitle = slide0.elements.find((e) => e.placeholder === 'ctrTitle')!
    expect(slideTitle.anchor.originalXml).not.toContain('<a:xfrm') // precondition: geometry is inherited

    const layoutPath = slide0.layoutPath!
    const layout = parseMasterPart(opened.archive, layoutPath)!
    const layoutTitle = layout.elements.find((e) => e.placeholder === 'ctrTitle')!
    const newX = 1234567
    layoutTitle.transform = {
      ...layoutTitle.transform,
      offset: { ...layoutTitle.transform.offset, x: newX },
    }
    layoutTitle.dirtyTransform = true
    opened.archive.entries.set(layoutPath, Buffer.from(patchSlideXml(layout), 'utf8'))

    // Inheritance takes effect immediately after reparse
    const fresh = materializeSlide(opened, 0)!
    const freshTitle = fresh.elements.find((e) => e.placeholder === 'ctrTitle')!
    expect(freshTitle.transform.offset.x).toBe(newX)

    // Still in effect after save and reopen
    const reopened = await openPptx(await savePptx(opened))
    const t2 = reopened.deck.slides[0]!.elements.find((e) => e.placeholder === 'ctrTitle')!
    expect(t2.transform.offset.x).toBe(newX)
  })

  it('edit master placeholder text color (run-level patch): part parses with the new color after save and reopen', async () => {
    const opened = await openPptx(fx('01_standard_business.pptx'))
    const masterPath = listMasterParts(opened.archive)[0]!.partPath
    const master = parseMasterPart(opened.archive, masterPath)!
    const title = master.elements.find((e) => e.placeholder === 'title') as TextElement
    for (const p of title.text!.paragraphs)
      for (const r of p.runs) {
        r.color = '#FF0000'
        // user explicitly recolored: no longer follows the theme or inheritance
        delete r.colorFollowsTheme
        delete r.colorInherited
      }
    title.dirty = true
    opened.archive.entries.set(masterPath, Buffer.from(patchSlideXml(master), 'utf8'))

    const reopened = await openPptx(await savePptx(opened))
    const m2 = parseMasterPart(reopened.archive, masterPath)!
    const t2 = m2.elements.find((e) => e.placeholder === 'title') as TextElement
    expect(t2.text!.paragraphs[0]!.runs[0]!.color?.toUpperCase()).toBe('#FF0000')
  })
})
