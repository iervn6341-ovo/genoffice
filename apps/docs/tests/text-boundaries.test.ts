import { describe, expect, it } from 'vitest'

import { textBoundaryMarks, type PageFrame } from '../src/renderer/editor/pagination-gaps'

const frame = (over: Partial<PageFrame> = {}): PageFrame => ({
  top: 100,
  bottom: 1156,
  left: 50,
  width: 816,
  margins: { top: 96, bottom: 96, left: 96, right: 96 },
  ...over,
})

describe('textBoundaryMarks', () => {
  it('marks the four corners of the text area (page minus margins)', () => {
    expect(textBoundaryMarks(frame())).toEqual([
      { corner: 'tl', x: 146, y: 196 },
      { corner: 'tr', x: 770, y: 196 },
      { corner: 'bl', x: 146, y: 1060 },
      { corner: 'br', x: 770, y: 1060 },
    ])
  })

  it('follows per-page margins (a section with different margins)', () => {
    const marks = textBoundaryMarks(
      frame({ margins: { top: 48, bottom: 24, left: 20, right: 40 } }),
    )
    expect(marks.find((m) => m.corner === 'tl')).toEqual({ corner: 'tl', x: 70, y: 148 })
    expect(marks.find((m) => m.corner === 'br')).toEqual({ corner: 'br', x: 826, y: 1132 })
  })

  it('draws nothing without margins or when the margins swallow the page', () => {
    expect(textBoundaryMarks(frame({ margins: undefined }))).toEqual([])
    expect(
      textBoundaryMarks(frame({ margins: { top: 600, bottom: 600, left: 10, right: 10 } })),
    ).toEqual([])
  })
})
