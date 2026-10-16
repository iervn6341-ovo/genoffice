/**
 * Print job helpers shared by the Markdown and HTML apps (pure; no Electron import).
 *
 * macOS hands the rendered PDF to the system print panel (see mac-print.ts). Everywhere else the
 * renderer's print sheet collects these options and the main process prints the self-contained
 * document from a hidden window with `webContents.print`, so every option below changes the output.
 */

export type PrintPaper = 'A3' | 'A4' | 'A5' | 'Letter' | 'Legal' | 'Tabloid'
export type PrintMargin = 'default' | 'none' | 'narrow' | 'wide'

export const PRINT_PAPERS: readonly PrintPaper[] = ['A3', 'A4', 'A5', 'Letter', 'Legal', 'Tabloid']
export const PRINT_MARGINS: readonly PrintMargin[] = ['default', 'none', 'narrow', 'wide']

export interface HtmlPrintOptions {
  paper: PrintPaper
  landscape: boolean
  margin: PrintMargin
  /** percent, 10-200 */
  scale: number
  /** browser-style title / page-number header and footer */
  headerFooter: boolean
  /** background colors and images */
  background: boolean
  /** "1-3,5"; empty prints every page */
  pages: string
}

export const DEFAULT_HTML_PRINT_OPTIONS: HtmlPrintOptions = {
  paper: 'A4',
  landscape: false,
  margin: 'default',
  scale: 100,
  headerFooter: false,
  background: true,
  pages: '',
}

/** Margin widths in inches; "default" matches the PDF export's 0.6in. */
const MARGIN_INCHES: Record<PrintMargin, number> = {
  default: 0.6,
  none: 0,
  narrow: 0.25,
  wide: 1,
}

export const MIN_PRINT_SCALE = 10
export const MAX_PRINT_SCALE = 200

/** Letter-using regions default to Letter, everyone else to A4. */
export function defaultPaperFor(countryCode: string | undefined): PrintPaper {
  return countryCode &&
    ['US', 'CA', 'MX', 'PH', 'CL', 'CO', 'VE'].includes(countryCode.toUpperCase())
    ? 'Letter'
    : 'A4'
}

/** Coerces untrusted IPC input into a complete, valid option set. */
export function normalizeHtmlPrintOptions(input: unknown): HtmlPrintOptions {
  const raw = (input && typeof input === 'object' ? input : {}) as Record<string, unknown>
  const d = DEFAULT_HTML_PRINT_OPTIONS
  const scale = Number(raw.scale)
  return {
    paper: PRINT_PAPERS.includes(raw.paper as PrintPaper) ? (raw.paper as PrintPaper) : d.paper,
    landscape: typeof raw.landscape === 'boolean' ? raw.landscape : d.landscape,
    margin: PRINT_MARGINS.includes(raw.margin as PrintMargin)
      ? (raw.margin as PrintMargin)
      : d.margin,
    scale: Number.isFinite(scale)
      ? Math.min(MAX_PRINT_SCALE, Math.max(MIN_PRINT_SCALE, Math.round(scale)))
      : d.scale,
    headerFooter: typeof raw.headerFooter === 'boolean' ? raw.headerFooter : d.headerFooter,
    background: typeof raw.background === 'boolean' ? raw.background : d.background,
    pages: typeof raw.pages === 'string' ? raw.pages.slice(0, 200) : d.pages,
  }
}

export interface PrintRange {
  /** 1-based, inclusive */
  from: number
  to: number
}

/**
 * "1-3,5,8-" style page text → ranges (1-based). Empty text means every page (`[]`); malformed
 * text (zero, reversed ranges, junk) is null. The page count is unknown before printing, so an
 * open-ended "8-" runs to the last page.
 */
export function parsePrintRanges(text: string): PrintRange[] | null {
  const out: PrintRange[] = []
  for (const part of text.split(/[,，;；]/)) {
    const s = part.trim()
    if (!s) continue
    const m = /^(\d+)\s*[-–]\s*(\d*)$|^(\d+)$/.exec(s)
    if (!m) return null
    const from = Number(m[1] ?? m[3])
    const to = m[3] !== undefined ? from : m[2] ? Number(m[2]) : 100_000
    if (from < 1 || to < from) return null
    out.push({ from, to })
  }
  return out
}

/** The CSS the print document must carry for margins (the dialog reads the page's own @page). */
export function printPageCss(options: HtmlPrintOptions): string {
  return `@page { margin: ${MARGIN_INCHES[options.margin]}in; }`
}

/** Inserts the @page rule into a self-contained HTML document. */
export function withPrintPageCss(html: string, options: HtmlPrintOptions): string {
  const style = `<style>${printPageCss(options)}</style>`
  if (/<\/head\s*>/i.test(html)) return html.replace(/<\/head\s*>/i, `${style}</head>`)
  return style + html
}

export interface WebContentsPrintArgs {
  silent: false
  printBackground: boolean
  landscape: boolean
  scaleFactor: number
  pageSize: PrintPaper
  pageRanges?: Array<{ from: number; to: number }>
  header?: string
  footer?: string
}

/**
 * Options for `webContents.print`. Electron counts pages from 0. Null: the page range text is
 * malformed (the sheet validates first, so this only guards a stale or hand-built request).
 */
export function webContentsPrintArgs(
  options: HtmlPrintOptions,
  title: string,
): WebContentsPrintArgs | null {
  const ranges = parsePrintRanges(options.pages)
  if (!ranges) return null
  return {
    silent: false,
    printBackground: options.background,
    landscape: options.landscape,
    scaleFactor: options.scale,
    pageSize: options.paper,
    ...(ranges.length > 0
      ? { pageRanges: ranges.map((r) => ({ from: r.from - 1, to: r.to - 1 })) }
      : {}),
    ...(options.headerFooter ? { header: title, footer: ' ' } : {}),
  }
}

export type HtmlPrintOutcome =
  | { readonly status: 'printed' }
  | { readonly status: 'canceled' }
  | { readonly status: 'error'; readonly message: string }

/** The slice of a hidden BrowserWindow that printing needs. */
export interface PrintableContents {
  print(
    options: WebContentsPrintArgs,
    callback: (success: boolean, failureReason: string) => void,
  ): void
}

/** Opens the system print dialog for an already-loaded document and reports how it ended. */
export function printLoadedDocument(
  contents: PrintableContents,
  options: HtmlPrintOptions,
  title: string,
): Promise<HtmlPrintOutcome> {
  const args = webContentsPrintArgs(options, title)
  if (!args) return Promise.resolve({ status: 'error', message: 'the page range is not valid' })
  return new Promise((resolve) => {
    try {
      contents.print(args, (success, failureReason) => {
        if (success) resolve({ status: 'printed' })
        else if (/cancel/i.test(failureReason)) resolve({ status: 'canceled' })
        else resolve({ status: 'error', message: failureReason || 'the print job failed' })
      })
    } catch (error: unknown) {
      resolve({ status: 'error', message: error instanceof Error ? error.message : String(error) })
    }
  })
}

/** Renderer side: macOS goes straight to the system panel, other platforms open the print sheet. */
export type PrintEntry = 'native-panel' | 'in-app'

export function printEntryFor(platform: string): PrintEntry {
  return platform.toLowerCase().includes('mac') ? 'native-panel' : 'in-app'
}
