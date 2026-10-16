/**
 * pdf.js glue for the print sheet's preview: opens the PDF segments the main process printed
 * (main/pdf-export.ts renderPrintPreview) and draws one page into a canvas. Document content:
 * the paper is always white, whatever the UI theme.
 */
import type { PDFDocumentProxy } from 'pdfjs-dist'

/** Paper colour of the preview bitmap (document data, not UI chrome). */
const PAPER = 'rgb(255, 255, 255)'

type PdfJs = typeof import('pdfjs-dist/legacy/build/pdf.mjs')
let pdfjsPromise: Promise<PdfJs> | null = null

function loadPdfJs(): Promise<PdfJs> {
  pdfjsPromise ??= (async () => {
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const { default: workerUrl } = await import('pdfjs-dist/legacy/build/pdf.worker.min.mjs?url')
    pdfjs.GlobalWorkerOptions.workerSrc = workerUrl
    return pdfjs
  })()
  return pdfjsPromise
}

export async function openPreviewPdf(bytes: Uint8Array): Promise<PDFDocumentProxy> {
  const pdfjs = await loadPdfJs()
  return pdfjs.getDocument({ data: bytes.slice(), useWasm: false }).promise
}

export interface PreviewDraw {
  /** Resolves when the page is on the canvas; rejects (or resolves) quietly after cancel(). */
  readonly done: Promise<void>
  cancel(): void
}

/** Fits page `pageNumber` (1-based) of `doc` into `maxWidth` x `maxHeight` CSS px. */
export function drawPreviewPage(
  doc: PDFDocumentProxy,
  pageNumber: number,
  canvas: HTMLCanvasElement,
  maxWidth: number,
  maxHeight: number,
): PreviewDraw {
  let task: { cancel(): void; promise: Promise<unknown> } | null = null
  let canceled = false
  const done = (async () => {
    const page = await doc.getPage(pageNumber)
    if (canceled) return
    const base = page.getViewport({ scale: 1 })
    const fit = Math.max(0.05, Math.min(maxWidth / base.width, maxHeight / base.height))
    const ratio = Math.min(window.devicePixelRatio || 1, 2)
    const viewport = page.getViewport({ scale: fit * ratio })
    canvas.width = Math.max(1, Math.round(viewport.width))
    canvas.height = Math.max(1, Math.round(viewport.height))
    canvas.style.width = `${Math.round(viewport.width / ratio)}px`
    canvas.style.height = `${Math.round(viewport.height / ratio)}px`
    task = page.render({ canvas, viewport, background: PAPER })
    try {
      await task.promise
    } catch (error: unknown) {
      if (!canceled) throw error
    } finally {
      page.cleanup()
    }
  })()
  return {
    done,
    cancel() {
      canceled = true
      task?.cancel()
    },
  }
}
