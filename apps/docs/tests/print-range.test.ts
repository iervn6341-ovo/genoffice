import { describe, expect, it } from 'vitest'
import { parsePrintRange } from '../src/renderer/print-range'

describe('parsePrintRange', () => {
  it('parses single pages and ranges into 0-based indices', () => {
    expect(parsePrintRange('1,3,5-8', 10)).toEqual([0, 2, 4, 5, 6, 7])
  })

  it('dedupes overlaps and sorts', () => {
    expect(parsePrintRange('5-6, 2, 5', 10)).toEqual([1, 4, 5])
  })

  it('accepts CJK separators and en-dash', () => {
    expect(parsePrintRange('1，3、5–6', 6)).toEqual([0, 2, 4, 5])
  })

  it('accepts spaces around dashes', () => {
    expect(parsePrintRange('1 - 3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1 – 3, 5', 5)).toEqual([0, 1, 2, 4])
  })

  it('accepts em-dash, tilde, fullwidth tilde and wave dash (IME variants)', () => {
    expect(parsePrintRange('1—3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1~3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1～3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1〜3', 5)).toEqual([0, 1, 2])
    expect(parsePrintRange('1 — 3, 5～6', 6)).toEqual([0, 1, 2, 4, 5])
  })

  it('rejects empty, malformed, reversed, and out-of-bounds input', () => {
    expect(parsePrintRange('', 5)).toBeNull()
    expect(parsePrintRange('a-b', 5)).toBeNull()
    expect(parsePrintRange('4-2', 5)).toBeNull()
    expect(parsePrintRange('0', 5)).toBeNull()
    expect(parsePrintRange('6', 5)).toBeNull()
    expect(parsePrintRange('1-9', 5)).toBeNull()
  })
})

import { parsePrintRangeSections } from '../src/renderer/print-range'

describe('parsePrintRangeSections', () => {
  // 6 pages: section 1 = pages 1-2, section 2 = pages 3-5, section 3 = page 6
  const pages = [0, 0, 1, 1, 1, 2].map((section) => ({ section }))

  it('keeps the plain page syntax', () => {
    expect(parsePrintRangeSections('1,3,5-6', pages)).toEqual([0, 2, 4, 5])
    expect(parsePrintRangeSections('2 – 3', pages)).toEqual([1, 2])
  })

  it('prints whole sections', () => {
    expect(parsePrintRangeSections('s2', pages)).toEqual([2, 3, 4])
    expect(parsePrintRangeSections('s1,s3', pages)).toEqual([0, 1, 5])
    expect(parsePrintRangeSections('s2-s3', pages)).toEqual([2, 3, 4, 5])
  })

  it('addresses a page within a section', () => {
    expect(parsePrintRangeSections('p2s2', pages)).toEqual([3])
    expect(parsePrintRangeSections('p2s1-p1s3', pages)).toEqual([1, 2, 3, 4, 5])
    expect(parsePrintRangeSections('P3S2', pages)).toEqual([4])
  })

  it('mixes page numbers and sections in one range', () => {
    expect(parsePrintRangeSections('2-s2', pages)).toEqual([1, 2, 3, 4])
  })

  it('rejects unknown sections, pages past the section end, reversed and malformed input', () => {
    expect(parsePrintRangeSections('s4', pages)).toBeNull()
    expect(parsePrintRangeSections('p3s1', pages)).toBeNull()
    expect(parsePrintRangeSections('s3-s1', pages)).toBeNull()
    expect(parsePrintRangeSections('7', pages)).toBeNull()
    expect(parsePrintRangeSections('', pages)).toBeNull()
    expect(parsePrintRangeSections('s', pages)).toBeNull()
    expect(parsePrintRangeSections('1-2-3', pages)).toBeNull()
  })
})
