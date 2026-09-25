/**
 * Formatted speaker notes: the notes body placeholder reads and writes through the same text
 * model as slide text, so font / size / colour survive, untouched runs keep their bytes, and
 * the size inherited from the notes master's <p:notesStyle> is not baked into runs.
 */
import { describe, it, expect } from 'vitest'
import {
  createBlankPptx,
  getSlideNotes,
  getSlideNotesParagraphs,
  NOTES_DEFAULT_FONT_PT,
  notesPathForSlide,
  openPptx,
  savePptx,
  setSlideNotes,
  setSlideNotesParagraphs,
  type OpenedPptx,
} from '../src/index'

/** A deck whose slide 1 carries PowerPoint-authored notes: notesStyle 14pt, a bold red run */
async function deckWithAuthoredNotes(): Promise<OpenedPptx> {
  const opened = await openPptx(await createBlankPptx())
  setSlideNotes(opened, 0, 'seed')
  const { archive } = opened
  const notesPath = notesPathForSlide(archive, opened.deck.slides[0]!.path)!
  const master = [...archive.entries.keys()].find((p) => /notesMaster\d+\.xml$/.test(p))!
  archive.entries.set(
    master,
    Buffer.from(
      archive
        .readText(master)!
        .replace(
          '</p:notesMaster>',
          '<p:notesStyle><a:lvl1pPr marL="0"><a:defRPr sz="1400"/></a:lvl1pPr></p:notesStyle></p:notesMaster>',
        ),
    ),
  )
  const xml = archive.readText(notesPath)!
  archive.entries.set(
    notesPath,
    Buffer.from(
      xml.replace(
        /<p:txBody>[\s\S]*?<\/p:txBody>/,
        '<p:txBody><a:bodyPr/><a:lstStyle/>' +
          '<a:p><a:r><a:rPr lang="en-US" b="1" dirty="0"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr><a:t>Key point</a:t></a:r>' +
          '<a:r><a:rPr lang="en-US" dirty="0"/><a:t> then detail</a:t></a:r></a:p>' +
          '<a:p><a:r><a:rPr lang="en-US" sz="2000" dirty="0"><a:latin typeface="Georgia"/></a:rPr><a:t>Big line</a:t></a:r></a:p>' +
          '</p:txBody>',
      ),
    ),
  )
  return opened
}

describe('formatted speaker notes', () => {
  it('reads run formatting and the notes-master size', async () => {
    const opened = await deckWithAuthoredNotes()
    const paras = getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)
    expect(paras).toHaveLength(2)
    const [key, detail] = paras[0]!.runs
    expect(key).toMatchObject({ text: 'Key point', bold: true, color: '#C00000' })
    expect(detail).toMatchObject({ text: ' then detail', fontSize: 14, fontSizeImplicit: true })
    expect(paras[1]!.runs[0]).toMatchObject({
      text: 'Big line',
      fontSize: 20,
      fontFamily: 'Georgia',
    })
  })

  it('writing the unchanged paragraphs back keeps the notes part byte-identical', async () => {
    const opened = await deckWithAuthoredNotes()
    const path = notesPathForSlide(opened.archive, opened.deck.slides[0]!.path)!
    const before = opened.archive.readText(path)
    const paras = getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)
    expect(setSlideNotesParagraphs(opened, 0, paras)).toBe(true)
    expect(opened.archive.readText(path)).toBe(before)
  })

  it('a changed size / colour / font lands on that run only and survives save', async () => {
    const opened = await deckWithAuthoredNotes()
    const paras = getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)
    const detail = paras[0]!.runs[1]!
    detail.fontSize = 24
    delete detail.fontSizeImplicit
    detail.color = '#0070C0'
    detail.fontFamily = 'Courier New'
    delete detail.fontImplicit
    setSlideNotesParagraphs(opened, 0, paras)
    const reopened = await openPptx(await savePptx(opened))
    const back = getSlideNotesParagraphs(reopened.archive, reopened.deck.slides[0]!.path)
    expect(back[0]!.runs[0]).toMatchObject({ text: 'Key point', bold: true, color: '#C00000' })
    expect(back[0]!.runs[1]).toMatchObject({
      text: ' then detail',
      fontSize: 24,
      color: '#0070C0',
      fontFamily: 'Courier New',
    })
    expect(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path)).toBe(
      'Key point then detail\nBig line',
    )
  })

  it('an inherited size is not baked into runs on a structural rewrite', async () => {
    const opened = await deckWithAuthoredNotes()
    const paras = getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)
    paras.push({ ...paras[0]!, runs: [{ ...paras[0]!.runs[1]!, text: 'Added' }] })
    setSlideNotesParagraphs(opened, 0, paras)
    const xml = opened.archive.readText(
      notesPathForSlide(opened.archive, opened.deck.slides[0]!.path)!,
    )!
    expect(xml).toContain('Added')
    expect(/<a:rPr[^>]*sz="1400"/.test(xml)).toBe(false)
  })

  it('a fresh deck gets a notes page and the default size', async () => {
    const opened = await openPptx(await createBlankPptx())
    expect(getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)).toEqual([])
    setSlideNotesParagraphs(opened, 0, [
      { runs: [{ text: 'Hello', bold: true, italic: false, underline: false, fontSize: 18 }] },
    ])
    const reopened = await openPptx(await savePptx(opened))
    const back = getSlideNotesParagraphs(reopened.archive, reopened.deck.slides[0]!.path)
    expect(back[0]!.runs[0]).toMatchObject({ text: 'Hello', bold: true, fontSize: 18 })
    setSlideNotes(reopened, 0, 'plain')
    expect(
      getSlideNotesParagraphs(reopened.archive, reopened.deck.slides[0]!.path)[0]!.runs[0],
    ).toMatchObject({ text: 'plain', fontSize: NOTES_DEFAULT_FONT_PT })
  })
})
