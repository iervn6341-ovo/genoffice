import { describe, expect, it } from 'vitest'
import { computeSectionedSlicesF2, type BlockBox } from '../src/renderer/pagination'

/**
 * Geometry of the Word Mac probes (docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/
 * section-end-endnotes): A4, 1in margins -> 930.67px of content; one 25pt body line pitch.
 */
const PAGE = 930.67
const LINE = 33.33
const NOTE = 122.67
const SEP = 16

function layout(opts: { notes1: number; secondStartsNewPage: boolean }) {
  const blocks: BlockBox[] = []
  let y = 0
  for (let i = 0; i < 15; i++) {
    blocks.push({
      top: y,
      height: LINE,
      section: 0,
      lineBoxes: [{ offsetInBlock: 0, height: LINE }],
    })
    y += LINE
  }
  const rows = (count: number, section: number): number[] => {
    const tops: number[] = []
    for (let i = 0; i < count; i++) {
      const h = NOTE + (i === 0 ? SEP : 0)
      tops.push(y)
      blocks.push({
        top: y,
        height: h,
        section,
        isEndnotes: true,
        widowControl: false,
        lineBoxes: [{ offsetInBlock: 0, height: h }],
      })
      y += h
    }
    return tops
  }
  const section1Rows = rows(opts.notes1, 0)
  const bodyTwoTop = y
  for (let i = 0; i < 2; i++) {
    blocks.push({
      top: y,
      height: LINE,
      section: 1,
      lineBoxes: [{ offsetInBlock: 0, height: LINE }],
    })
    y += LINE
  }
  const slices = computeSectionedSlicesF2(
    blocks,
    [
      { contentHeight: PAGE, forceBreak: false },
      { contentHeight: PAGE, forceBreak: opts.secondStartsNewPage },
    ],
    y,
  )
  return { slices, section1Rows, bodyTwoTop }
}

describe('section-end endnote areas in the page slicer', () => {
  it('a short area follows its section on the same page; the next page-type section starts a new page', () => {
    const { slices, bodyTwoTop } = layout({ notes1: 2, secondStartsNewPage: true })
    expect(slices).toHaveLength(2)
    expect(slices[0].section).toBe(0)
    expect(slices[1]).toMatchObject({ section: 1, start: bodyTwoTop })
  })

  it('a long area spills across pages at entry boundaries before the next section begins', () => {
    const { slices, section1Rows, bodyTwoTop } = layout({
      notes1: 14,
      secondStartsNewPage: true,
    })
    // the first cuts land on row tops (whole entries move, like the document-end area)
    const rowTops = new Set(section1Rows.map((t) => Math.round(t * 100)))
    const cuts = slices.slice(1, -1).map((s) => Math.round(s.start * 100))
    expect(cuts.length).toBeGreaterThanOrEqual(2)
    for (const cut of cuts) expect(rowTops.has(cut)).toBe(true)
    // the last page belongs to section 2 and begins exactly at its first body line
    expect(slices.at(-1)).toMatchObject({ section: 1, start: bodyTwoTop })
    expect(slices.filter((s) => s.section === 0).length).toBeGreaterThanOrEqual(3)
    // every page-0 section slice stays inside the page capacity
    for (const s of slices.filter((x) => x.section === 0))
      expect(s.end - s.start).toBeLessThanOrEqual(PAGE + 0.01)
  })

  it('a continuous next section continues on the same page right after the notes', () => {
    const { slices, bodyTwoTop } = layout({ notes1: 2, secondStartsNewPage: false })
    expect(slices).toHaveLength(1)
    expect(slices[0].end).toBeGreaterThan(bodyTwoTop)
  })
})
