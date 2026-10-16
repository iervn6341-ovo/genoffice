/**
 * Print options of the in-app print dialog (Windows / Linux; macOS hands the saved file to the
 * system print panel instead). Pure helpers: which pages print, in what order, and how each page
 * image is sized on the sheet.
 */

export type PrintParity = 'all' | 'odd' | 'even'
export type PrintScaleMode = 'fit' | 'shrink' | 'actual' | 'custom'

export interface PdfPrintOptions {
  parity: PrintParity
  reverse: boolean
  scaleMode: PrintScaleMode
  /** percent, used by the 'custom' scale mode */
  customScale: number
  /** comments, highlights and other annotations */
  annotations: boolean
}

export const DEFAULT_PDF_PRINT_OPTIONS: PdfPrintOptions = {
  parity: 'all',
  reverse: false,
  scaleMode: 'fit',
  customScale: 100,
  annotations: true,
}

export const MIN_CUSTOM_SCALE = 10
export const MAX_CUSTOM_SCALE = 400

export function clampCustomScale(value: number): number {
  if (!Number.isFinite(value)) return 100
  return Math.min(MAX_CUSTOM_SCALE, Math.max(MIN_CUSTOM_SCALE, Math.round(value)))
}

/**
 * Narrows the chosen 1-based pages to the odd / even page numbers and orders them. Parity follows
 * the page numbers shown in the viewer (page 1 is odd), like Acrobat's "Odd pages only".
 */
export function selectPrintPages(
  pages: readonly number[],
  options: Pick<PdfPrintOptions, 'parity' | 'reverse'>,
): number[] {
  const sorted = [...new Set(pages)].sort((a, b) => a - b)
  const kept =
    options.parity === 'odd'
      ? sorted.filter((n) => n % 2 === 1)
      : options.parity === 'even'
        ? sorted.filter((n) => n % 2 === 0)
        : sorted
  return options.reverse ? kept.reverse() : kept
}

export interface PageFit {
  /** wrap the image in a clipping, centering box (needed once the page can outgrow the sheet) */
  clip: boolean
  /** inline style for the page image */
  style: Record<string, string>
}

/**
 * How a page image sits on the sheet. 'fit' keeps the long-standing behavior (scaled to the
 * printable width, capped to its height). The others size the image from the page's own size in
 * points (1pt = 1/72in on paper): 'shrink' only reduces pages larger than the sheet, 'actual'
 * and 'custom' print at 100% / N% and crop whatever falls outside, as Acrobat does.
 */
export function pageFit(mode: PrintScaleMode, customScale: number, widthPt: number): PageFit {
  if (mode === 'fit' || !(widthPt > 0)) return { clip: false, style: {} }
  if (mode === 'shrink') {
    return {
      clip: true,
      style: {
        width: `${widthPt}pt`,
        maxWidth: '100%',
        maxHeight: '100%',
        objectFit: 'contain',
      },
    }
  }
  const pct = mode === 'actual' ? 100 : clampCustomScale(customScale)
  return {
    clip: true,
    style: {
      width: `${Math.round(widthPt * pct) / 100}pt`,
      maxWidth: 'none',
      maxHeight: 'none',
    },
  }
}
