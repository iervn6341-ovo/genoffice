import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { printPdf, printScaleForAreas } from '../src/renderer/print'

function fakeDoc(numPages: number, render = vi.fn(() => ({ promise: Promise.resolve() }))) {
  const getPage = vi.fn(async () => ({
    getViewport: ({ scale }: { scale: number }) => ({ width: 612 * scale, height: 792 * scale }),
    render,
  }))
  return { doc: { numPages, getPage } as unknown as PDFDocumentProxy, getPage, render }
}

beforeEach(() => {
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue('data:image/jpeg;base64,FAKE')
  // jsdom does not implement img.decode or window.print
  Object.defineProperty(HTMLImageElement.prototype, 'decode', {
    configurable: true,
    writable: true,
    value: vi.fn(() => Promise.resolve()),
  })
  window.print = vi.fn(() => {
    window.dispatchEvent(new Event('afterprint'))
  })
})

afterEach(() => {
  vi.restoreAllMocks()
  document.body.innerHTML = ''
})

describe('printPdf', () => {
  it('renders one image per page into a print root and opens the print dialog', async () => {
    const { doc, getPage, render } = fakeDoc(3)
    let imgsAtPrintTime = 0
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      imgsAtPrintTime = document.querySelectorAll('.pdf-print-root img').length
      window.dispatchEvent(new Event('afterprint'))
    })

    await printPdf(doc)

    expect(getPage).toHaveBeenCalledTimes(6)
    expect(getPage).toHaveBeenNthCalledWith(1, 1)
    expect(getPage).toHaveBeenNthCalledWith(2, 2)
    expect(getPage).toHaveBeenNthCalledWith(3, 3)
    expect(getPage).toHaveBeenNthCalledWith(4, 1)
    expect(getPage).toHaveBeenNthCalledWith(5, 2)
    expect(getPage).toHaveBeenNthCalledWith(6, 3)
    expect(render).toHaveBeenCalledTimes(3)
    expect(imgsAtPrintTime).toBe(3)
    expect(window.print).toHaveBeenCalledTimes(1)
  })

  it('removes the print root after the dialog closes', async () => {
    const { doc } = fakeDoc(2)
    await printPdf(doc)
    expect(document.querySelector('.pdf-print-root')).toBeNull()
  })

  it('renders only the requested pages, sorted, when given a subset', async () => {
    const { doc, getPage, render } = fakeDoc(5)
    let imgsAtPrintTime = 0
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      imgsAtPrintTime = document.querySelectorAll('.pdf-print-root img').length
      window.dispatchEvent(new Event('afterprint'))
    })

    await printPdf(doc, [4, 2, 99, 0])

    // Out-of-range entries are dropped; the rest render in document order.
    // Measure pass fetches each target first, then the render pass refetches.
    expect(getPage).toHaveBeenCalledTimes(4)
    expect(getPage).toHaveBeenNthCalledWith(1, 2)
    expect(getPage).toHaveBeenNthCalledWith(2, 4)
    expect(getPage).toHaveBeenNthCalledWith(3, 2)
    expect(getPage).toHaveBeenNthCalledWith(4, 4)
    expect(render).toHaveBeenCalledTimes(2)
    expect(imgsAtPrintTime).toBe(2)
    expect(window.print).toHaveBeenCalledTimes(1)
  })

  it('drops non-integer page targets instead of aborting the job', async () => {
    const { doc, getPage } = fakeDoc(5)
    await printPdf(doc, [1.5, 2, 99])
    // 1.5 would throw inside pdf.js getPage and abort the whole print;
    // only the valid integer target renders (measure + render passes).
    expect(getPage).toHaveBeenCalledTimes(2)
    expect(getPage).toHaveBeenNthCalledWith(1, 2)
    expect(getPage).toHaveBeenNthCalledWith(2, 2)
    expect(window.print).toHaveBeenCalledTimes(1)
  })

  it('prints the pages in the given order when ordered (reverse printing)', async () => {
    const { doc, getPage } = fakeDoc(5)
    await printPdf(doc, [5, 3, 1], undefined, true)
    // measure pass then render pass, both in the requested order
    expect(getPage.mock.calls.map((c) => (c as unknown as [number])[0])).toEqual([5, 3, 1, 5, 3, 1])
  })

  it('sizes page images by the chosen scale mode and wraps them in a clipping box', async () => {
    const { doc } = fakeDoc(2)
    let html = ''
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      html = document.querySelector('.pdf-print-root')!.innerHTML
      window.dispatchEvent(new Event('afterprint'))
    })
    await printPdf(doc, undefined, { scaleMode: 'custom', customScale: 50 })
    expect(html.match(/pdf-print-page/g)).toHaveLength(2)
    expect(html).toContain('width: 306pt')
  })

  it('keeps plain images (no wrapper) for fit-to-paper', async () => {
    const { doc } = fakeDoc(1)
    let html = ''
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      html = document.querySelector('.pdf-print-root')!.innerHTML
      window.dispatchEvent(new Event('afterprint'))
    })
    await printPdf(doc, undefined, { scaleMode: 'fit' })
    expect(html).not.toContain('pdf-print-page')
  })

  it('turns annotations off only when asked to', async () => {
    const off = fakeDoc(1)
    await printPdf(off.doc, undefined, { annotations: false })
    expect(off.render).toHaveBeenCalledWith(expect.objectContaining({ annotationMode: 0 }))
    const on = fakeDoc(1)
    await printPdf(on.doc, undefined, { annotations: true })
    expect(on.render).toHaveBeenCalledWith(
      expect.not.objectContaining({ annotationMode: expect.anything() }),
    )
  })

  it('waits for afterprint before resolving', async () => {
    const { doc } = fakeDoc(1)
    let fireAfterPrint: () => void = () => {}
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      fireAfterPrint = () => window.dispatchEvent(new Event('afterprint'))
    })

    let resolved = false
    const done = printPdf(doc).then(() => {
      resolved = true
    })
    // Let rendering and the print call complete
    await vi.waitFor(() => expect(window.print).toHaveBeenCalled())
    expect(resolved).toBe(false)
    expect(document.querySelector('.pdf-print-root')).not.toBeNull()

    fireAfterPrint()
    await done
    expect(resolved).toBe(true)
    expect(document.querySelector('.pdf-print-root')).toBeNull()
  })

  it('cleans up the print root even when image decoding fails', async () => {
    const { doc } = fakeDoc(1)
    ;(HTMLImageElement.prototype.decode as ReturnType<typeof vi.fn>).mockRejectedValue(
      new Error('decode failed'),
    )
    await expect(printPdf(doc)).rejects.toThrow('decode failed')
    expect(document.querySelector('.pdf-print-root')).toBeNull()
    expect(window.print).not.toHaveBeenCalled()
  })
})

/** US-Letter page area at scale 1 (PDF points²) */
const LETTER_AREA = 612 * 792

describe('printScaleForAreas', () => {
  it('uses the 200 DPI target while the document fits the pixel budget', () => {
    expect(printScaleForAreas([LETTER_AREA])).toBeCloseTo(200 / 72, 10)
    // ~40 letter pages at 200 DPI ≈ the 150 MP budget
    expect(printScaleForAreas(new Array(40).fill(LETTER_AREA))).toBeCloseTo(200 / 72, 10)
  })

  it('falls back to the target for empty or degenerate input', () => {
    expect(printScaleForAreas([])).toBeCloseTo(200 / 72, 10)
    expect(printScaleForAreas([0])).toBeCloseTo(200 / 72, 10)
    expect(printScaleForAreas([Number.NaN])).toBeCloseTo(200 / 72, 10)
    expect(printScaleForAreas([Number.POSITIVE_INFINITY])).toBeCloseTo(200 / 72, 10)
  })

  it('scales down proportionally over budget, never below the 150 DPI baseline', () => {
    const mid = printScaleForAreas(new Array(60).fill(LETTER_AREA))
    expect(mid).toBeLessThan(200 / 72)
    expect(mid).toBeGreaterThan(150 / 72)
    expect(printScaleForAreas(new Array(500).fill(LETTER_AREA))).toBeCloseTo(150 / 72, 10)
  })
})

describe('printPdf render scale', () => {
  function scalesDoc(numPages: number) {
    const scales: number[] = []
    const getViewport = vi.fn(({ scale }: { scale: number }) => {
      scales.push(scale)
      return { width: 612 * scale, height: 792 * scale }
    })
    const render = vi.fn(() => ({ promise: Promise.resolve() }))
    const getPage = vi.fn(async () => ({ getViewport, render }))
    const doc = { numPages, getPage } as unknown as PDFDocumentProxy
    return { doc, scales, getPage, render }
  }

  it('renders small documents at 200 DPI (all pages measured before the render pass)', async () => {
    const { doc, scales, getPage } = scalesDoc(2)
    await printPdf(doc)
    expect(getPage).toHaveBeenCalledTimes(4)
    expect(scales).toEqual([1, 1, 200 / 72, 200 / 72])
  })

  it('floors huge documents at the 150 DPI baseline', async () => {
    const { doc, scales } = scalesDoc(200)
    await printPdf(doc)
    expect(scales.filter((s) => s !== 1)).toHaveLength(200)
    for (const s of scales) {
      expect(s === 1 || s === 150 / 72).toBe(true)
    }
  })

  it('budgets a page subset on its own: two pages out of a huge document print at 200 DPI', async () => {
    const { doc, scales, getPage } = scalesDoc(200)
    await printPdf(doc, [7, 3])
    expect(getPage).toHaveBeenCalledTimes(4)
    expect(getPage).toHaveBeenNthCalledWith(1, 3)
    expect(getPage).toHaveBeenNthCalledWith(2, 7)
    expect(getPage).toHaveBeenNthCalledWith(3, 3)
    expect(getPage).toHaveBeenNthCalledWith(4, 7)
    expect(scales).toEqual([1, 1, 200 / 72, 200 / 72])
  })

  it('streams pages and cleans up each one, bounding a corrupt huge page', async () => {
    const cleanups: number[] = []
    const viewports: Array<{ width: number; height: number }> = [
      { width: 612, height: 792 },
      { width: Number.NaN, height: Number.NaN },
      { width: 1e6, height: 1e6 },
    ]
    const getPage = vi.fn(async (n: number) => {
      const vp = viewports[n - 1]!
      return {
        getViewport: vi.fn(({ scale }: { scale: number }) => ({
          width: vp.width * scale,
          height: vp.height * scale,
        })),
        render: vi.fn(() => ({ promise: Promise.resolve() })),
        cleanup: vi.fn(() => cleanups.push(1)),
      }
    })
    const doc = { numPages: 3, getPage } as unknown as PDFDocumentProxy
    // The print root is removed on afterprint, so snapshot the image count
    // at print time like the other tests do.
    let imgsAtPrintTime = 0
    ;(window.print as ReturnType<typeof vi.fn>).mockImplementation(() => {
      imgsAtPrintTime = document.querySelectorAll('.pdf-print-root img').length
      window.dispatchEvent(new Event('afterprint'))
    })
    await printPdf(doc)
    // each page cleaned up twice: once after measure, once after render/skip
    expect(cleanups).toHaveLength(6)
    // NaN pages are skipped and huge pages render at their own reduced scale
    // rather than aborting the whole print
    expect(imgsAtPrintTime).toBe(2)
  })
})
