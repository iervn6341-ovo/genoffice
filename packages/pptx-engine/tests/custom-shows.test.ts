import { describe, expect, it } from 'vitest'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  deleteSlide,
  getCustomShows,
  moveSlide,
  openPptx,
  savePptx,
  setCustomShows,
} from '../src/index'

const fixture = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), 'fixtures/01_standard_business.pptx'),
)
const presPath = 'ppt/presentation.xml'
const open = () => openPptx(fixture)

describe('native custom shows', () => {
  it('writes ordered, repeated slide references and round-trips names and IDs without touching rels', async () => {
    const deck = await open()
    const rels = deck.archive.readText('ppt/_rels/presentation.xml.rels')
    expect(getCustomShows(deck)).toEqual([])
    const shows = [
      { id: '0', name: 'Review & sign-off', slideIndices: [3, 0, 3, 1] },
      { id: '17', name: 'Board "short"', slideIndices: [4, 2] },
    ]
    expect(setCustomShows(deck, shows)).toBe(true)
    const xml = deck.archive.readText(presPath)!
    expect(xml).toContain('name="Review &amp; sign-off"')
    expect(xml).toContain('name="Board &quot;short&quot;"')
    expect(xml.match(/<p:sld r:id=/g) ?? []).toHaveLength(6)
    expect(deck.archive.readText('ppt/_rels/presentation.xml.rels')).toBe(rels)
    const reopened = await openPptx(await savePptx(deck))
    expect(getCustomShows(reopened)).toEqual(shows)
  })

  it('retains untouched imported extensions, dangling refs, and duplicate slide refs through edits', async () => {
    const deck = await open()
    const original = deck.archive.readText(presPath)!
    const rid = /<p:sldId\b[^>]*r:id="([^"]+)"/.exec(original)![1]!
    const imported = `<p:custShowLst custom="keep"><p:custShow name="Imported" id="23" ext="keep"><p:sldLst><p:sld r:id="${rid}"/><p:sld r:id="rMissing"/><p:sld r:id="${rid}"/></p:sldLst><p:extLst><p:ext uri="custom"/></p:extLst></p:custShow><p:extLst><p:ext uri="list"/></p:extLst></p:custShowLst>`
    deck.archive.entries.set(
      presPath,
      Buffer.from(original.replace('</p:presentation>', `${imported}</p:presentation>`)),
    )
    expect(getCustomShows(deck)).toEqual([{ id: '23', name: 'Imported', slideIndices: [0, 0] }])
    expect(
      setCustomShows(deck, [...getCustomShows(deck), { id: '24', name: 'New', slideIndices: [2] }]),
    ).toBe(true)
    const xml = deck.archive.readText(presPath)!
    expect(xml).toContain('<p:sld r:id="rMissing"/>')
    expect(xml).toContain('<p:ext uri="custom"/>')
    expect(xml).toContain('<p:ext uri="list"/>')
    expect(getCustomShows(await openPptx(await savePptx(deck)))).toEqual([
      { id: '23', name: 'Imported', slideIndices: [0, 0] },
      { id: '24', name: 'New', slideIndices: [2] },
    ])
  })

  it('retains imported extension children within sldLst when an existing show is edited', async () => {
    const deck = await open()
    const original = deck.archive.readText(presPath)!
    const rid = /<p:sldId\b[^>]*r:id="([^"]+)"/.exec(original)![1]!
    const list = `<p:custShowLst><p:custShow name="Ext" id="2"><p:sldLst native="keep"><p:sld r:id="${rid}"/><p:extLst><p:ext uri="inside"/></p:extLst></p:sldLst></p:custShow></p:custShowLst>`
    deck.archive.entries.set(
      presPath,
      Buffer.from(original.replace('</p:presentation>', `${list}</p:presentation>`)),
    )
    expect(setCustomShows(deck, [{ id: '2', name: 'Edited', slideIndices: [3, 1] }])).toBe(true)
    const xml = deck.archive.readText(presPath)!
    expect(xml).toContain('<p:sldLst native="keep">')
    expect(xml).toContain('<p:ext uri="inside"/>')
    expect(getCustomShows(await openPptx(await savePptx(deck)))).toEqual([
      { id: '2', name: 'Edited', slideIndices: [3, 1] },
    ])
  })

  it('uses slide relationship identity across reorder and rejects invalid edits without mutation', async () => {
    const deck = await open()
    expect(setCustomShows(deck, [{ id: '5', name: 'Order', slideIndices: [0, 4, 2] }])).toBe(true)
    const before = deck.archive.readText(presPath)!
    expect(setCustomShows(deck, [{ id: '5', name: 'Bad', slideIndices: [99] }])).toBe(false)
    expect(
      setCustomShows(deck, [
        { id: '5', name: 'One', slideIndices: [] },
        { id: '5', name: 'Two', slideIndices: [] },
      ]),
    ).toBe(false)
    expect(setCustomShows(deck, [{ id: 'cs-abc', name: 'Legacy', slideIndices: [] }])).toBe(false)
    expect(deck.archive.readText(presPath)).toBe(before)
    expect(moveSlide(deck, 0, 3)).toBeTruthy()
    expect(getCustomShows(deck)[0]!.slideIndices).toEqual([3, 4, 1])
    expect(getCustomShows(await openPptx(await savePptx(deck)))[0]!.slideIndices).toEqual([3, 4, 1])
  })

  it('removes a deleted slide reference without disturbing the other playback entries', async () => {
    const deck = await open()
    expect(setCustomShows(deck, [{ id: '3', name: 'Deletion', slideIndices: [1, 3, 1, 4] }])).toBe(
      true,
    )
    expect(deleteSlide(deck, 1)).toBe(true)
    expect(getCustomShows(deck)).toEqual([{ id: '3', name: 'Deletion', slideIndices: [2, 3] }])
    expect(getCustomShows(await openPptx(await savePptx(deck)))).toEqual([
      { id: '3', name: 'Deletion', slideIndices: [2, 3] },
    ])
  })

  it('deleting the last custom show removes the list and preserves unrelated presentation XML', async () => {
    const deck = await open()
    const original = deck.archive.readText(presPath)!
    expect(setCustomShows(deck, [{ id: '0', name: 'Transient', slideIndices: [0] }])).toBe(true)
    expect(setCustomShows(deck, [])).toBe(true)
    expect(deck.archive.readText(presPath)).toBe(original)
  })
})
