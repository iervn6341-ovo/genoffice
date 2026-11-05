import { expect, it } from 'vitest'
import {
  applyHeaderFooter,
  createBlankPptx,
  insertBlankSlide,
  deleteSlide,
  openPptx,
  reparseDeck,
  moveSlide,
  savePptx,
  setElementFont,
  materializeSlide,
  readHeaderFooter,
  getFirstSlideNumber,
  getSlideNumberingRule,
} from '../src/index'
const values = (o: Awaited<ReturnType<typeof openPptx>>) =>
  o.deck.slides.map((s) =>
    s.elements
      .filter((e) => e.placeholder === 'sldNum')
      .map((e) =>
        e.type === 'text' || e.type === 'shape'
          ? e.text?.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('')
          : '',
      )
      .join(''),
  )
it('persists offset fields with valid PowerPoint start numbering and recalculates after insert/delete', async () => {
  let o = await openPptx(await createBlankPptx())
  for (let i = 0; i < 5; i++) insertBlankSlide(o, i)
  o = reparseDeck(o)
  applyHeaderFooter(o, {
    footer: 'Keep footer',
    numbering: { startSlide: 5, startNumber: 1, hideBefore: true, enabled: true },
  })
  expect(values(o)).toEqual(['', '', '', '', '1', '2'])
  o = await openPptx(await savePptx(o))
  expect(getSlideNumberingRule(o.archive)).toEqual({
    startSlide: 5,
    startNumber: 1,
    hideBefore: true,
    enabled: true,
  })
  expect(getFirstSlideNumber(o.archive)).toBe(1)
  expect(o.archive.readText(o.deck.slides[4]!.path)).toContain('type="genofficeSlideNumber"')
  expect(values(o)).toEqual(['', '', '', '', '1', '2'])
  insertBlankSlide(o, 0)
  expect(values(o)).toEqual(['', '', '', '', '1', '2', '3'])
  o = reparseDeck(o)
  expect(values(o)).toEqual(['', '', '', '', '1', '2', '3'])
  deleteSlide(o, 0)
  expect(values(o)).toEqual(['', '', '', '', '1', '2'])
  moveSlide(o, 5, 0)
  expect(values(o)).toEqual(['', '', '', '', '1', '2'])
  o = reparseDeck(o)
  expect(values(o)).toEqual(['', '', '', '', '1', '2'])
  const untouchedFooter = o.deck.slides[5]!.elements.filter((e) => e.placeholder === 'ftr').map(
    (e) => e.anchor.originalXml,
  )
  applyHeaderFooter(
    o,
    { numbering: { startSlide: 5, startNumber: 9, hideBefore: false, enabled: true } },
    [4],
  )
  expect(values(o)).toEqual(['1', '2', '3', '4', '9', '10'])
  expect(
    o.deck.slides[5]!.elements.filter((e) => e.placeholder === 'ftr').map(
      (e) => e.anchor.originalXml,
    ),
  ).toEqual(untouchedFooter)
})
it('migrates legacy negative starting numbers on save', async () => {
  const o = await openPptx(await createBlankPptx())
  for (let i = 0; i < 5; i++) insertBlankSlide(o, i)
  applyHeaderFooter(o, { firstSlideNum: -3, slideNum: true }, [4, 5])
  const reopened = await openPptx(await savePptx(o))
  expect(values(reopened)).toEqual(['', '', '', '', '1', '2'])
  expect(reopened.archive.readText('ppt/presentation.xml')).not.toContain('firstSlideNum="-3"')
})

it('keeps title exclusion through save and restores built-in fields when numbering from slide one', async () => {
  let o = await openPptx(await createBlankPptx())
  const path = o.deck.slides[0]!.layoutPath!
  const xml = o.archive.readText(path)!
  const open = xml.match(/<p:sldLayout\b[^>]*>/)![0]
  const title = /\btype="[^"]*"/.test(open)
    ? open.replace(/\btype="[^"]*"/, 'type="title"')
    : open.replace('>', ' type="title">')
  o.archive.entries.set(path, Buffer.from(xml.replace(open, title)))
  o = reparseDeck(o)
  applyHeaderFooter(
    o,
    {
      dontShowOnTitleSlide: true,
      numbering: { startSlide: 1, startNumber: 7, hideBefore: true, enabled: true },
    },
    [0],
  )
  o = await openPptx(await savePptx(o))
  expect(values(o)).toEqual([''])
  expect(getSlideNumberingRule(o.archive)?.excludeTitleSlide).toBe(true)
  applyHeaderFooter(o, {
    dontShowOnTitleSlide: false,
    numbering: { startSlide: 1, startNumber: 7, hideBefore: true, enabled: true },
  })
  o = await openPptx(await savePptx(o))
  expect(values(o)).toEqual(['7'])
  expect(o.archive.readText(o.deck.slides[0]!.path)).toContain('type="slidenum"')
})

it('retains a chosen field color through renumbering and save/reopen', async () => {
  let o = await openPptx(await createBlankPptx())
  insertBlankSlide(o, 0)
  applyHeaderFooter(o, {
    numbering: { startSlide: 2, startNumber: 1, hideBefore: true, enabled: true },
  })
  const slide = o.deck.slides[1]!,
    field = slide.elements.find((e) => e.placeholder === 'sldNum')!
  setElementFont(slide, field.id, { color: '#156082' })
  applyHeaderFooter(o, {
    numbering: { startSlide: 2, startNumber: 7, hideBefore: true, enabled: true },
  })
  o = await openPptx(await savePptx(o))
  const saved = o.deck.slides[1]!.elements.find((e) => e.placeholder === 'sldNum')!
  if (saved.type !== 'text' && saved.type !== 'shape') throw new Error('Missing field')
  expect(saved.text?.paragraphs[0]?.runs[0]).toMatchObject({
    field: 'genofficeSlideNumber',
    color: '#156082',
    text: '7',
  })
})

it('hides empty local placeholders in PowerPoint and restores the same styled objects', async () => {
  let o = await openPptx(await createBlankPptx())
  insertBlankSlide(o, 0)
  applyHeaderFooter(o, { slideNum: true, footer: 'Footer', date: 'November 4' })
  const first = o.deck.slides[0]!
  const original = new Map(
    first.elements.filter((e) => e.placeholder).map((e) => [e.placeholder, e.anchor.originalXml]),
  )
  for (const e of first.elements.filter((e) => e.placeholder))
    setElementFont(first, e.id, { color: '#156082' })
  applyHeaderFooter(o, {
    numbering: { startSlide: 2, startNumber: 1, hideBefore: true, enabled: true },
  })
  o = await openPptx(await savePptx(o))
  expect(values(o)).toEqual(['', '1'])
  expect(readHeaderFooter(o.deck.slides[0]!)).toMatchObject({
    slideNum: false,
    footer: null,
    date: null,
  })
  for (const e of o.deck.slides[0]!.elements.filter((e) => e.placeholder)) {
    expect(e).toMatchObject({ type: 'passthrough', noChip: true })
    expect(e.anchor.originalXml).toMatch(/<p:cNvPr\b[^>]*hidden="1"/)
  }
  // Simulate an already-saved file from before the fix: ordinary Save must migrate it.
  const path = o.deck.slides[0]!.path
  o.archive.entries.set(path, Buffer.from(o.archive.readText(path)!.replace(/ hidden="1"/g, '')))
  o = reparseDeck(o)
  o = await openPptx(await savePptx(o))
  expect(
    o.deck.slides[0]!.elements.find((e) => e.placeholder === 'sldNum')!.anchor.originalXml,
  ).toContain('hidden="1"')
  applyHeaderFooter(o, {
    numbering: { startSlide: 1, startNumber: 7, hideBefore: true, enabled: true },
    footer: 'Restored footer',
    date: 'November 5',
  })
  o = await openPptx(await savePptx(o))
  expect(values(o)).toEqual(['7', '8'])
  for (const type of ['sldNum', 'dt', 'ftr'] as const) {
    const matching = o.deck.slides[0]!.elements.filter((e) => e.placeholder === type)
    expect(matching).toHaveLength(1)
    const xml = matching[0]!.anchor.originalXml
    expect(xml).not.toContain('hidden="1"')
    expect(xml).toContain('val="156082"')
    expect(xml.match(/<p:cNvPr\b[^>]*\bid="(\d+)"/)?.[1]).toBe(
      original.get(type)!.match(/<p:cNvPr\b[^>]*\bid="(\d+)"/)?.[1],
    )
    expect(xml.match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/)?.[0]).toBe(
      original.get(type)!.match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/)?.[0],
    )
  }
})

it('keeps hidden inherited-number masks after reopen and restores their original layout style', async () => {
  let o = await openPptx(await createBlankPptx())
  insertBlankSlide(o, 0)
  applyHeaderFooter(o, { slideNum: true }, [0])
  const first = o.deck.slides[0]!
  const field = first.elements.find((e) => e.placeholder === 'sldNum')!
  setElementFont(first, field.id, { color: '#156082', fontSizePt: 11 })
  materializeSlide(o, 0)
  const template = o.deck.slides[0]!.elements.find(
    (e) => e.placeholder === 'sldNum',
  )!.anchor.originalXml.replace(/(<p:cNvPr\b[^>]*\bid=)"\d+"/, '$1"999"')
  for (let i = 0; i < o.deck.slides.length; i++) {
    const s = o.deck.slides[i]!
    s.elements = s.elements.filter((e) => e.placeholder !== 'sldNum')
    s.structureDirty = true
    materializeSlide(o, i)
  }
  for (const path of new Set(o.deck.slides.map((s) => s.layoutPath!))) {
    let xml = o.archive
      .readText(path)!
      .replace(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g, (shape) =>
        shape.includes('type="sldNum"') ? '' : shape,
      )
      .replace(/<p:hf\b[^>]*\/>/g, '')
      .replace('</p:spTree>', `${template}</p:spTree>`)
      .replace('</p:sldLayout>', '<p:hf sldNum="1"/></p:sldLayout>')
    o.archive.entries.set(path, Buffer.from(xml))
  }
  o = reparseDeck(o)
  expect(o.deck.slides[0]!.decorations?.some((e) => e.placeholder === 'sldNum')).toBe(true)
  applyHeaderFooter(o, {
    numbering: { startSlide: 2, startNumber: 1, hideBefore: true, enabled: true },
  })
  o = await openPptx(await savePptx(o))
  expect(values(o)).toEqual(['', '1'])
  expect(o.deck.slides[0]!.decorations?.some((e) => e.placeholder === 'sldNum')).toBeFalsy()
  const mask = o.deck.slides[0]!.elements.find((e) => e.placeholder === 'sldNum')!
  expect(mask).toMatchObject({ type: 'passthrough', noChip: true })
  expect(mask.anchor.originalXml).toContain('hidden="1"')
  applyHeaderFooter(o, {
    numbering: { startSlide: 1, startNumber: 1, hideBefore: true, enabled: true },
  })
  o = await openPptx(await savePptx(o))
  expect(values(o)).toEqual(['1', '2'])
  const restored = o.deck.slides[0]!.elements.filter((e) => e.placeholder === 'sldNum')
  expect(restored).toHaveLength(1)
  expect(restored[0]!.anchor.originalXml).toContain('val="156082"')
  expect(restored[0]!.anchor.originalXml).not.toContain('hidden="1"')
  expect(restored[0]!.anchor.originalXml.match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/)?.[0]).toBe(
    template.match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/)?.[0],
  )
})
