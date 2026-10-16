import { describe, expect, it } from 'vitest'
import { clampCustomScale, pageFit, selectPrintPages } from '../src/renderer/print-plan'

describe('selectPrintPages', () => {
  const pages = [1, 2, 3, 4, 5, 6]

  it('keeps page order by default', () => {
    expect(selectPrintPages([3, 1, 2], { parity: 'all', reverse: false })).toEqual([1, 2, 3])
  })

  it('narrows to odd or even page numbers', () => {
    expect(selectPrintPages(pages, { parity: 'odd', reverse: false })).toEqual([1, 3, 5])
    expect(selectPrintPages(pages, { parity: 'even', reverse: false })).toEqual([2, 4, 6])
    // parity follows the page numbers, not the position inside the chosen range
    expect(selectPrintPages([2, 3, 4], { parity: 'odd', reverse: false })).toEqual([3])
  })

  it('reverses after filtering and de-duplicates', () => {
    expect(selectPrintPages([5, 5, 2, 3], { parity: 'all', reverse: true })).toEqual([5, 3, 2])
    expect(selectPrintPages(pages, { parity: 'even', reverse: true })).toEqual([6, 4, 2])
  })

  it('can come out empty (one odd page chosen with even-only)', () => {
    expect(selectPrintPages([3], { parity: 'even', reverse: false })).toEqual([])
  })
})

describe('pageFit', () => {
  it('leaves the default fit-to-paper layout alone', () => {
    expect(pageFit('fit', 100, 612)).toEqual({ clip: false, style: {} })
    expect(pageFit('actual', 100, 0)).toEqual({ clip: false, style: {} })
  })

  it('shrinks only oversized pages: real size, capped by the sheet', () => {
    expect(pageFit('shrink', 100, 612)).toMatchObject({
      clip: true,
      style: { width: '612pt', maxWidth: '100%', maxHeight: '100%' },
    })
  })

  it('prints actual size and custom percentages from the page size in points', () => {
    expect(pageFit('actual', 55, 595).style.width).toBe('595pt')
    expect(pageFit('custom', 50, 595).style).toMatchObject({ width: '297.5pt', maxWidth: 'none' })
    expect(pageFit('custom', 150, 600).style.width).toBe('900pt')
  })

  it('clamps the custom scale', () => {
    expect(clampCustomScale(1)).toBe(10)
    expect(clampCustomScale(9999)).toBe(400)
    expect(clampCustomScale(Number.NaN)).toBe(100)
    expect(pageFit('custom', 5, 100).style.width).toBe('10pt')
  })
})
