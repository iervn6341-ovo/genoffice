import { describe, expect, it } from 'vitest'
import {
  applyNotesHandoutsHeaderFooter,
  createBlankPptx,
  getFirstSlideNumber,
  openPptx,
  readNotesHandoutsHeaderFooter,
  readNotesHandoutsPrintFields,
  readMasterHeaderFooterFlags,
  setMasterHeaderFooterFlag,
  savePptx,
  setFirstSlideNumber,
  setSlideNotes,
  setNotesMasterPlaceholderEnabled,
} from '../src/index'

const settings = {
  headerEnabled: true,
  header: 'Course notes',
  footerEnabled: true,
  footer: 'Private',
  dateEnabled: true,
  date: '10/13/2026',
  dateAuto: true,
  dateFormat: 'datetime8',
  pageNum: true,
}

describe('native notes and handouts header/footer', () => {
  it('resolves local and inherited Notes Master picture-bullet relationships', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, 'Picture bullet')
    const local = 'ppt/notesSlides/notesSlide1.xml'
    const master = 'ppt/notesMasters/notesMaster1.xml'
    const png = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL/nwAAAABJRU5ErkJggg==',
      'base64',
    )
    opened.archive.entries.set('ppt/media/local-bullet.png', png)
    opened.archive.entries.set('ppt/media/master-bullet.png', png)
    const localRels = 'ppt/notesSlides/_rels/notesSlide1.xml.rels'
    const masterRels = 'ppt/notesMasters/_rels/notesMaster1.xml.rels'
    const rel = (id: string, target: string) =>
      `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${target}"/>`
    opened.archive.entries.set(
      localRels,
      Buffer.from(
        opened.archive
          .readText(localRels)!
          .replace('</Relationships>', `${rel('rId80', 'local-bullet.png')}</Relationships>`),
      ),
    )
    opened.archive.entries.set(
      masterRels,
      Buffer.from(
        opened.archive
          .readText(masterRels)!
          .replace('</Relationships>', `${rel('rId81', 'master-bullet.png')}</Relationships>`),
      ),
    )
    const notesXml = opened.archive.readText(local)!
    const masterXml = opened.archive.readText(master)!
    const localBullet = '<a:pPr><a:buBlip><a:blip r:embed="rId80"/></a:buBlip></a:pPr>'
    opened.archive.entries.set(
      local,
      Buffer.from(notesXml.replace('<a:p><a:r>', `<a:p>${localBullet}<a:r>`)),
    )
    opened.archive.entries.set(
      master,
      Buffer.from(
        masterXml.replace(
          '</p:notesMaster>',
          '<p:notesStyle><a:lvl1pPr><a:buBlip><a:blip r:embed="rId81"/></a:buBlip></a:lvl1pPr></p:notesStyle></p:notesMaster>',
        ),
      ),
    )
    const localPicture =
      readNotesHandoutsPrintFields(opened).notesPages[0]?.notesParagraphs?.[0]?.bullet
    expect(localPicture?.type).toBe('blip')
    expect(localPicture?.imageDataUrl).toBe(`data:image/png;base64,${png.toString('base64')}`)
    opened.archive.entries.set(local, Buffer.from(notesXml))
    const inheritedPicture =
      readNotesHandoutsPrintFields(opened).notesPages[0]?.notesParagraphs?.[0]?.bullet
    expect(inheritedPicture?.type).toBe('blip')
    expect(inheritedPicture?.imageDataUrl).toBe(`data:image/png;base64,${png.toString('base64')}`)
  })
  it('uses the synthesized notes-page fallback when no local notesSlide exists', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    const master = 'ppt/notesMasters/notesMaster1.xml'
    setNotesMasterPlaceholderEnabled(opened.archive, master, 'sldImg', false)
    setNotesMasterPlaceholderEnabled(opened.archive, master, 'body', false)
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]).toMatchObject({
      slideImageVisible: true,
      notesBodyVisible: true,
    })
  })

  it('lets local notes placeholders survive master removal, but omits locally deleted regions after reopen', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    setSlideNotes(opened, 0, 'Retained speaker text')
    const master = 'ppt/notesMasters/notesMaster1.xml'
    const local = 'ppt/notesSlides/notesSlide1.xml'
    const before = opened.archive.readText(local)!
    expect(setNotesMasterPlaceholderEnabled(opened.archive, master, 'sldImg', false)).toBe(true)
    expect(setNotesMasterPlaceholderEnabled(opened.archive, master, 'body', false)).toBe(true)
    expect(opened.archive.readText(local)).toBe(before)
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]).toMatchObject({
      slideImageVisible: true,
      notesBodyVisible: true,
    })
    const localShapes = [...before.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].map(
      (match) => match[0],
    )
    const removed = localShapes.filter((shape) =>
      /<p:ph\b[^>]*\btype=["'](?:sldImg|body)["']/.test(shape),
    )
    expect(removed).toHaveLength(2)
    opened.archive.entries.set(
      local,
      Buffer.from(removed.reduce((xml, shape) => xml.replace(shape, ''), before)),
    )
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]).toMatchObject({
      slideImageVisible: false,
      notesBodyVisible: false,
    })
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.archive.readText(local)).not.toContain('Retained speaker text')
    expect(readNotesHandoutsPrintFields(reopened).notesPages[0]).toMatchObject({
      slideImageVisible: false,
      notesBodyVisible: false,
    })
  })
  it('toggles one master field with native p:hf while retaining field identity and geometry', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    const path = 'ppt/notesMasters/notesMaster1.xml'
    const before = opened.archive.readText(path)!
    const id = /<a:fld id="([^"]+)" type="datetime8"/.exec(before)![1]
    expect(setMasterHeaderFooterFlag(opened.archive, path, 'dt', false)).toBe(true)
    expect(readMasterHeaderFooterFlags(opened.archive, path)?.dt).toBe(false)
    expect(opened.archive.readText(path)).toContain(`<a:fld id="${id}" type="datetime8"`)
    expect(setMasterHeaderFooterFlag(opened.archive, path, 'dt', true)).toBe(true)
    const printed = readNotesHandoutsPrintFields(opened)
    expect(printed.notesPages[0]!.date?.box).toEqual({
      x: 457200,
      y: 8595360,
      cx: 1828800,
      cy: 320040,
    })
    expect(printed.notesPages[0]!.canvas).toEqual({ cx: 6858000, cy: 9144000 })
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.archive.readText(path)).toContain(`<a:fld id="${id}" type="datetime8"`)
    expect(readMasterHeaderFooterFlags(reopened.archive, path)?.dt).toBe(true)
  })

  it('resolves native theme-linked field color without replacing the scheme reference', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    const path = 'ppt/notesMasters/notesMaster1.xml'
    const xml = opened.archive.readText(path)!
    const themed = xml.replace(
      'type="datetime8"><a:rPr lang="en-US" sz="1200"/>',
      'type="datetime8"><a:rPr lang="en-US" sz="1200"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr>',
    )
    opened.archive.entries.set(path, Buffer.from(themed))
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]?.date?.color).toMatch(
      /^#[0-9A-F]{6}$/,
    )
    expect(opened.archive.readText(path)).toContain('<a:schemeClr val="accent1"/>')
  })

  it('prints native transform flags and inherited notesStyle with local run overrides', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    setSlideNotes(opened, 0, 'Inherited style')
    const masterPath = 'ppt/notesMasters/notesMaster1.xml'
    const original = opened.archive.readText(masterPath)!
    const styled = original
      .replace(
        '</p:notesMaster>',
        '<p:notesStyle><a:lvl1pPr algn="r"><a:defRPr sz="1800"><a:solidFill><a:srgbClr val="225588"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>',
      )
      .replace(
        '<a:xfrm><a:off x="457200" y="8595360"/>',
        '<a:xfrm rot="5400000" flipH="1"><a:off x="457200" y="8595360"/>',
      )
    opened.archive.entries.set(masterPath, Buffer.from(styled))
    const inherited = readNotesHandoutsPrintFields(opened).notesPages[0]!
    expect(inherited.date?.box).toMatchObject({ rotationDeg: 90, flipH: true })
    expect(inherited.notesParagraphs?.[0]).toMatchObject({
      align: 'right',
      runs: [{ text: 'Inherited style', fontSizePt: 18, color: '#225588', fontFamily: 'Arial' }],
    })
    const notesPath = 'ppt/notesSlides/notesSlide1.xml'
    const notes = opened.archive.readText(notesPath)!
    opened.archive.entries.set(
      notesPath,
      Buffer.from(
        notes.replace(
          '<a:rPr lang="zh-CN" dirty="0"/>',
          '<a:rPr lang="zh-CN" sz="2300" dirty="0"/>',
        ),
      ),
    )
    const local = readNotesHandoutsPrintFields(opened).notesPages[0]!
    expect(local.notesParagraphs?.[0]?.runs[0]?.fontSizePt).toBe(23)
    const reopened = await openPptx(await savePptx(opened))
    expect(readNotesHandoutsPrintFields(reopened).notesPages[0]?.date?.box).toMatchObject({
      rotationDeg: 90,
      flipH: true,
    })
  })

  it('projects authored notes paragraph direction, bullet, spacing, indents and tabs for print', async () => {
    const opened = await openPptx(await createBlankPptx())
    setSlideNotes(opened, 0, 'אב\tNotes')
    const path = 'ppt/notesSlides/notesSlide1.xml'
    const original = opened.archive.readText(path)!
    const pPr =
      '<a:pPr rtl="1" lvl="1" marL="914400" marR="127000" indent="-228600" defTabSz="457200" algn="r"><a:lnSpc><a:spcPts val="1800"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPct val="50000"/></a:spcAft><a:buClr><a:srgbClr val="AA3300"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="★"/><a:tabLst><a:tab pos="457200" algn="r"/><a:tab pos="914400" algn="ctr"/></a:tabLst></a:pPr>'
    opened.archive.entries.set(
      path,
      Buffer.from(original.replace('<a:p><a:r>', `<a:p>${pPr}<a:r>`)),
    )
    const paragraph = readNotesHandoutsPrintFields(opened).notesPages[0]?.notesParagraphs?.[0]
    expect(paragraph).toMatchObject({
      align: 'right',
      rtl: true,
      level: 1,
      lineExactPt: 18,
      spaceBeforePt: 6,
      spaceAfterPct: 50,
      marginLeftPt: 72,
      marginRightPt: 10,
      indentPt: -18,
      tabStopsPt: [36, 72],
      tabStops: [
        { posPt: 36, align: 'r' },
        { posPt: 72, align: 'ctr' },
      ],
      defaultTabPt: 36,
      bullet: { type: 'char', char: '★', color: '#AA3300', fontFamily: 'Arial' },
    })
    const reopened = await openPptx(await savePptx(opened))
    expect(
      readNotesHandoutsPrintFields(reopened).notesPages[0]?.notesParagraphs?.[0]?.bullet?.char,
    ).toBe('★')
  })

  it('keeps a local static date static even when the master date is dynamic', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    setSlideNotes(opened, 0, 'Speaker notes')
    const master = opened.archive.readText('ppt/notesMasters/notesMaster1.xml')!
    const dateShape = [...master.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)]
      .map((match) => match[0])
      .find((shape) => shape.includes('<p:ph type="dt"'))!
    const localStatic = dateShape
      .replace(/<a:fld\b[^>]*>/, '<a:r>')
      .replace('</a:fld>', '</a:r>')
      .replace('10/13/2026', 'Fixed local')
      .replace(/<a:xfrm>[\s\S]*?<\/a:xfrm>/, '')
    const path = 'ppt/notesSlides/notesSlide1.xml'
    opened.archive.entries.set(
      path,
      Buffer.from(
        opened.archive.readText(path)!.replace('</p:spTree>', `${localStatic}</p:spTree>`),
      ),
    )
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]?.date).toMatchObject({
      text: 'Fixed local',
      box: { x: 457200, y: 8595360, cx: 1828800, cy: 320040 },
    })
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]?.date?.field).toBeUndefined()
  })
  it('registers both masters and round-trips dynamic fields and slide number start', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(setFirstSlideNumber(opened.archive, 0)).toBe(true)
    expect(applyNotesHandoutsHeaderFooter(opened, settings)).toBe(true)
    const presentation = opened.archive.readText('ppt/presentation.xml')!
    expect(presentation).toContain('firstSlideNum="0"')
    expect(presentation).toContain('<p:notesMasterIdLst>')
    expect(presentation).toContain('<p:handoutMasterIdLst>')
    for (const path of [
      'ppt/notesMasters/notesMaster1.xml',
      'ppt/handoutMasters/handoutMaster1.xml',
    ]) {
      const xml = opened.archive.readText(path)!
      expect(xml).toContain('<p:hf hdr="1" ftr="1" dt="1" sldNum="1"/>')
      expect(xml).toContain('type="datetime8"')
      expect(xml).toContain('type="slidenum"')
      expect(xml).toContain('Course notes')
    }
    const reopened = await openPptx(await savePptx(opened))
    expect(getFirstSlideNumber(reopened.archive)).toBe(0)
    expect(readNotesHandoutsHeaderFooter(reopened)).toMatchObject(settings)
    expect(readNotesHandoutsPrintFields(reopened).handout.date?.field).toBe('datetime8')
  })

  it('keeps existing placeholder geometry and run properties while updating local notes overrides', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    setSlideNotes(opened, 0, 'Speaker text')
    const masterPath = 'ppt/notesMasters/notesMaster1.xml'
    const master = opened.archive.readText(masterPath)!
    const footer = [...master.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)]
      .map((match) => match[0])
      .find((shape) => shape.includes('<p:ph type="ftr"'))
    expect(footer).toBeTruthy()
    const linked = footer!.replace(
      '<a:rPr lang="en-US" sz="1200"/>',
      '<a:rPr lang="en-US" sz="1800" b="1"><a:hlinkClick r:id="rId45"/></a:rPr>',
    )
    opened.archive.entries.set(masterPath, Buffer.from(master.replace(footer!, linked)))
    const notesPath = 'ppt/notesSlides/notesSlide1.xml'
    const notes = opened.archive.readText(notesPath)!
    opened.archive.entries.set(
      notesPath,
      Buffer.from(
        notes.replace('</p:spTree>', `${linked.replace('Private', 'Old local')}</p:spTree>`),
      ),
    )
    expect(applyNotesHandoutsHeaderFooter(opened, { ...settings, footer: 'Updated' })).toBe(true)
    const nextMaster = opened.archive.readText(masterPath)!
    const nextLocal = opened.archive.readText(notesPath)!
    for (const xml of [nextMaster, nextLocal]) {
      expect(xml).toContain('sz="1800" b="1"')
      expect(xml).toContain('r:id="rId45"')
      expect(xml).toContain('Updated')
      expect(xml).toContain('<a:off x="2286000" y="8595360"/>')
    }
    const print = readNotesHandoutsPrintFields(opened)
    expect(print.notesPages[0]!.footer).toMatchObject({
      text: 'Updated',
      fontSizePt: 18,
      bold: true,
    })
    expect(print.handout.footer?.text).toBe('Updated')
  })

  it('retains imported dynamic field identity and attributes when changing its cached text', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    const path = 'ppt/notesMasters/notesMaster1.xml'
    const xml = opened.archive.readText(path)!
    const imported = xml.replace(
      /<a:fld id="[^"]+" type="datetime8">/,
      '<a:fld id="{EXISTING-FIELD}" type="datetime8" dirty="1">',
    )
    opened.archive.entries.set(path, Buffer.from(imported))
    expect(applyNotesHandoutsHeaderFooter(opened, { ...settings, date: 'New cache' })).toBe(true)
    const next = opened.archive.readText(path)!
    expect(next).toContain('<a:fld id="{EXISTING-FIELD}" type="datetime8" dirty="1">')
    expect(next).toContain('<a:t>New cache</a:t>')
    const reopened = await openPptx(await savePptx(opened))
    expect(reopened.archive.readText(path)).toContain(
      '<a:fld id="{EXISTING-FIELD}" type="datetime8" dirty="1">',
    )
  })

  it('switches native notes and handout dates between dynamic and fixed without stale field types', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    expect(
      applyNotesHandoutsHeaderFooter(opened, {
        ...settings,
        dateAuto: false,
        date: 'Board meeting',
      }),
    ).toBe(true)
    for (const path of [
      'ppt/notesMasters/notesMaster1.xml',
      'ppt/handoutMasters/handoutMaster1.xml',
    ]) {
      const xml = opened.archive.readText(path)!
      const dateShape = xml.match(/<p:sp\b[^>]*>[\s\S]*?<p:ph type="dt"[\s\S]*?<\/p:sp>/)?.[0]
      expect(dateShape).toContain('Board meeting')
      expect(dateShape).not.toContain('type="datetime8"')
    }
    const reopened = await openPptx(await savePptx(opened))
    expect(readNotesHandoutsHeaderFooter(reopened)).toMatchObject({
      dateAuto: false,
      date: 'Board meeting',
    })
    expect(readNotesHandoutsPrintFields(reopened).handout.date).toMatchObject({
      text: 'Board meeting',
    })
  })

  it('hides local dynamic notes fields by master flags without flattening their IDs or cache', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    setSlideNotes(opened, 0, 'Speaker text')
    const master = opened.archive.readText('ppt/notesMasters/notesMaster1.xml')!
    const shapes = [...master.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].map((match) => match[0])
    const dateShape = shapes
      .find((shape) => shape.includes('<p:ph type="dt"'))!
      .replace(
        /<a:fld id="[^"]+" type="datetime8">/,
        '<a:fld id="{LOCAL-DATE}" type="datetime8" dirty="1">',
      )
    const numberShape = shapes
      .find((shape) => shape.includes('<p:ph type="sldNum"'))!
      .replace(
        /<a:fld id="[^"]+" type="slidenum">/,
        '<a:fld id="{LOCAL-NUMBER}" type="slidenum" dirty="1">',
      )
    const notesPath = 'ppt/notesSlides/notesSlide1.xml'
    const notes = opened.archive.readText(notesPath)!
    const local = notes.replace('</p:spTree>', `${dateShape}${numberShape}</p:spTree>`)
    opened.archive.entries.set(notesPath, Buffer.from(local))

    expect(
      applyNotesHandoutsHeaderFooter(opened, { ...settings, dateEnabled: false, pageNum: false }),
    ).toBe(true)
    expect(opened.archive.readText(notesPath)).toBe(local)
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]?.date).toBeUndefined()
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]?.pageNum).toBeUndefined()
    const disabled = await openPptx(await savePptx(opened))
    expect(disabled.archive.readText(notesPath)).toBe(local)
    expect(disabled.archive.readText('ppt/notesMasters/notesMaster1.xml')).toMatch(
      /<p:hf\b[^>]*dt="0"[^>]*sldNum="0"/,
    )

    expect(applyNotesHandoutsHeaderFooter(disabled, settings)).toBe(true)
    expect(disabled.archive.readText(notesPath)).toBe(local)
    const restored = await openPptx(await savePptx(disabled))
    const restoredLocal = restored.archive.readText(notesPath)!
    expect(restoredLocal).toContain('<a:fld id="{LOCAL-DATE}" type="datetime8" dirty="1">')
    expect(restoredLocal).toContain('<a:fld id="{LOCAL-NUMBER}" type="slidenum" dirty="1">')
    expect(restoredLocal).toContain('<a:t>10/13/2026</a:t>')
    expect(readNotesHandoutsPrintFields(restored).notesPages[0]?.date?.field).toBe('datetime8')
    expect(readNotesHandoutsPrintFields(restored).notesPages[0]?.pageNum?.field).toBe('slidenum')
  })

  it('uses the notes page relationship when an imported deck has multiple notes masters', async () => {
    const opened = await openPptx(await createBlankPptx())
    applyNotesHandoutsHeaderFooter(opened, settings)
    setSlideNotes(opened, 0, 'Speaker text')
    const master = opened.archive.readText('ppt/notesMasters/notesMaster1.xml')!
    opened.archive.entries.set(
      'ppt/notesMasters/notesMaster2.xml',
      Buffer.from(master.replace('Course notes', 'Alternate master')),
    )
    const relsPath = 'ppt/notesSlides/_rels/notesSlide1.xml.rels'
    const rels = opened.archive.readText(relsPath)!
    opened.archive.entries.set(
      relsPath,
      Buffer.from(
        rels.replace('../notesMasters/notesMaster1.xml', '../notesMasters/notesMaster2.xml'),
      ),
    )
    expect(readNotesHandoutsPrintFields(opened).notesPages[0]?.header?.text).toBe(
      'Alternate master',
    )
    expect(readNotesHandoutsPrintFields(opened).handout.header?.text).toBe('Course notes')
  })

  it('does not create masters for an unchanged empty state and supports negative starts', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(
      applyNotesHandoutsHeaderFooter(opened, {
        ...settings,
        headerEnabled: false,
        footerEnabled: false,
        dateEnabled: false,
        pageNum: false,
      }),
    ).toBe(false)
    expect(opened.archive.has('ppt/notesMasters/notesMaster1.xml')).toBe(false)
    expect(setFirstSlideNumber(opened.archive, -3)).toBe(true)
    expect(getFirstSlideNumber((await openPptx(await savePptx(opened))).archive)).toBe(-3)
  })
})
