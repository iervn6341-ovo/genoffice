import { describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  resolveOutlineDestination,
  resolveOutlineDestinations,
} from '../src/renderer/outline-navigation'

function document() {
  return {
    numPages: 8,
    getDestination: vi.fn(async () => [{ num: 20, gen: 0 }, { name: 'XYZ' }, 40, 720, null]),
    getPageIndex: vi.fn(async () => 3),
  }
}

describe('PDF outline navigation', () => {
  it('resolves named bookmarks, page references and within-page coordinates', async () => {
    const doc = document()
    expect(await resolveOutlineDestination(doc as unknown as PDFDocumentProxy, 'section')).toEqual({
      pageIndex: 3,
      left: 40,
      top: 720,
    })
    expect(
      await resolveOutlineDestination(doc as unknown as PDFDocumentProxy, [1, { name: 'Fit' }]),
    ).toEqual({ pageIndex: 1, left: null, top: null })
    expect(
      await resolveOutlineDestination(doc as unknown as PDFDocumentProxy, [
        1,
        { name: 'FitH' },
        400,
      ]),
    ).toEqual({ pageIndex: 1, left: null, top: 400 })
    expect(
      await resolveOutlineDestination(doc as unknown as PDFDocumentProxy, [
        1,
        { name: 'FitR' },
        10,
        20,
        30,
        40,
      ]),
    ).toEqual({ pageIndex: 1, left: 10, top: 40 })
  })

  it('ignores broken and out-of-range destinations', async () => {
    const doc = document()
    for (const dest of [
      null,
      [],
      [99, { name: 'Fit' }],
      [-1, { name: 'Fit' }],
      [0.5, { name: 'Fit' }],
    ]) {
      expect(await resolveOutlineDestination(doc as unknown as PDFDocumentProxy, dest)).toBeNull()
    }
    doc.getDestination.mockRejectedValueOnce(new Error('bad bookmark'))
    expect(await resolveOutlineDestination(doc as unknown as PDFDocumentProxy, 'broken')).toBeNull()
  })

  it('caches repeated names, excludes external links and abandons stale documents', async () => {
    const doc = document()
    const items = [
      { id: 'a', node: { title: 'A', dest: 'section' } },
      { id: 'b', node: { title: 'B', dest: 'section' } },
      { id: 'link', node: { title: 'Link', dest: 'section', url: 'https://example.com' } },
    ]
    expect(
      (
        await resolveOutlineDestinations(doc as unknown as PDFDocumentProxy, items, () => false)
      ).map((target) => target.id),
    ).toEqual(['a', 'b'])
    expect(doc.getDestination).toHaveBeenCalledOnce()
    expect(
      await resolveOutlineDestinations(doc as unknown as PDFDocumentProxy, items, () => true),
    ).toEqual([])
    expect(doc.getDestination).toHaveBeenCalledOnce()
  })
})
