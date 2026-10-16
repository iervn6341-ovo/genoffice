/// PDF export: renders the print HTML (laid out by the renderer) in a hidden
/// scripting-disabled window and writes webContents.printToPDF's output where
/// the save dialog points.

import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { BrowserWindow, app, dialog, shell } from 'electron'

import {
  isHeadlessMode,
  printPdfWithSystemPanel,
  showSaveDialogWithMemory,
} from '@genoffice/electron-utils'

import { evenPageRanges, stitchPlan, type PageVariant } from './pdf-page-variants'
import { hasHeaderFooterPageOffset, headerFooterTemplateForPage } from './pdf-header-offsets'
import { directPrintOptions, type PrinterJob } from '@genoffice/electron-utils'
import { printOptionsFor } from './print-options'
export { resolvePrintRequest, storePrintHtmlChunk } from './print-transfer'

import type { IpcMainInvokeEvent, WebContents } from 'electron'
import type { PDFDocument } from 'pdf-lib'
import type {
  WorkbookExportPdfRequest,
  WorkbookExportPdfResult,
  WorkbookPrintPreviewResult,
  WorkbookPrintResult,
} from '../shared/desktop-api'

export async function exportPdf(
  event: IpcMainInvokeEvent,
  request: WorkbookExportPdfRequest,
): Promise<WorkbookExportPdfResult> {
  const parent = BrowserWindow.fromWebContents(event.sender)
  const dialogOptions = {
    defaultPath: request.fileName,
    filters: [{ name: 'PDF', extensions: ['pdf'] }],
  }
  // Headless export has no dialog to authorize a path; the CLI already chose one.
  const selection =
    isHeadlessMode() && request.outPath
      ? { canceled: false as const, filePath: request.outPath }
      : await showSaveDialogWithMemory(dialog, parent, dialogOptions)
  if (selection.canceled || !selection.filePath) return { canceled: true }

  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-pdf-'))
  const htmlPath = join(workDir, 'print.html')
  const window = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, javascript: false },
  })
  try {
    await writeFile(htmlPath, request.html, 'utf8')
    await window.loadFile(htmlPath)
    const pdf = await renderPdf(window.webContents, request)
    await writeFile(selection.filePath, pdf)
    return { canceled: false, path: selection.filePath }
  } finally {
    window.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}

/// macOS with no printer chosen in the app's own sheet: the standard macOS print panel *with its page
/// preview*, driven natively from a PDF (Chromium's own panel, what webContents.print opens, has no
/// preview). Elsewhere, or when the sheet picked a printer, the existing routes apply.
function usesMacPrintPanel(job: { printer?: string | undefined } | undefined): boolean {
  return process.platform === 'darwin' && !job?.printer
}

async function printPdfViaMacPanel(pdfPath: string): Promise<WorkbookPrintResult> {
  const outcome = await printPdfWithSystemPanel(pdfPath, 'GenOffice', {
    resourcesPath: app.isPackaged ? process.resourcesPath : undefined,
    devDir: join(app.getAppPath(), '../../packages/electron-utils/native'),
  })
  if (outcome.status === 'printed') return { ok: true }
  if (outcome.status === 'canceled') return { ok: false }
  return { ok: false, error: outcome.message }
}

/// Print: the same laid-out HTML in a hidden window, handed to the system
/// print dialog with the sheet's page setup preselected. Header / footer
/// templates are a printToPDF feature; the dialog's own header option stands
/// in for them.
export async function printWorkbook(
  event: IpcMainInvokeEvent,
  request: WorkbookExportPdfRequest,
  job?: PrinterJob & { pageFrom?: number | undefined; pageTo?: number | undefined },
): Promise<WorkbookPrintResult> {
  const owner = BrowserWindow.fromWebContents(event.sender)
  const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-print-'))
  const htmlPath = join(workDir, 'print.html')
  const window = new BrowserWindow({
    show: false,
    ...(owner && !owner.isDestroyed() ? { parent: owner } : {}),
    // Chromium attaches the native Windows print dialog to the printed window;
    // a hidden owner hides the dialog too, so Windows gets a real one
    ...(process.platform === 'win32'
      ? { width: 900, height: 700, autoHideMenuBar: true, closable: false, skipTaskbar: true }
      : {}),
    webPreferences: { sandbox: true, javascript: false },
  })
  try {
    await writeFile(htmlPath, request.html, 'utf8')
    await window.loadFile(htmlPath)
    if (usesMacPrintPanel(job)) {
      // the panel prints what the PDF holds, so the sheet's page setup is baked into it here
      const pdfPath = join(workDir, 'print.pdf')
      await writeFile(pdfPath, await renderPdf(window.webContents, request))
      return await printPdfViaMacPanel(pdfPath)
    }
    if (process.platform === 'win32') {
      window.show()
      window.focus()
    }
    const outcome = await new Promise<{ success: boolean; failureReason: string }>((resolve) => {
      window.webContents.print(
        {
          ...printOptionsFor(request),
          ...directPrintOptions(job),
          // Excel's Pages from / to (Electron counts pages from 0)
          ...(job?.pageFrom === undefined
            ? {}
            : { pageRanges: [{ from: job.pageFrom - 1, to: (job.pageTo ?? 100_000) - 1 }] }),
        },
        (success, failureReason) => resolve({ success, failureReason }),
      )
    })
    if (outcome.success) return { ok: true }
    return outcome.failureReason === 'Print job canceled'
      ? { ok: false }
      : { ok: false, error: outcome.failureReason }
  } catch (error: unknown) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  } finally {
    window.destroy()
    await rm(workDir, { recursive: true, force: true })
  }
}

/// One hidden print window kept alive while the print sheet pages through its preview, so each
/// further segment does not re-load a large laid-out sheet.
interface PreviewWindow {
  readonly id: string
  readonly window: BrowserWindow
  readonly workDir: string
  readonly request: WorkbookExportPdfRequest
  timer: ReturnType<typeof setTimeout> | null
}

const PREVIEW_IDLE_MS = 45_000
let previewWindow: PreviewWindow | null = null

async function disposePreview(entry: PreviewWindow | null): Promise<void> {
  if (!entry) return
  if (entry.timer) clearTimeout(entry.timer)
  if (!entry.window.isDestroyed()) entry.window.destroy()
  await rm(entry.workDir, { recursive: true, force: true })
}

/// Drops the cached preview window (print sheet closed, or `id` superseded).
export async function releasePrintPreview(id?: string): Promise<void> {
  if (id !== undefined && previewWindow?.id !== id) return
  const entry = previewWindow
  previewWindow = null
  await disposePreview(entry)
}

/// The print sheet's preview: pages `from`..`to` (1-based) of the laid-out sheet as a PDF, from
/// the same hidden-window printToPDF path the export uses. `request` is sent with the first
/// segment of a preview `id`; later segments reuse the cached window (`expired` asks for a
/// resend). `pdf: null` means the range lies past the last page.
export function renderPrintPreview(
  id: string,
  request: WorkbookExportPdfRequest | undefined,
  from: number,
  to: number,
): Promise<WorkbookPrintPreviewResult> {
  // one preview render at a time: a newer id disposes the window an older call still prints
  const run = previewQueue.then(() => renderPrintPreviewNow(id, request, from, to))
  previewQueue = run.then(
    () => undefined,
    () => undefined,
  )
  return run
}

let previewQueue: Promise<void> = Promise.resolve()

async function renderPrintPreviewNow(
  id: string,
  request: WorkbookExportPdfRequest | undefined,
  from: number,
  to: number,
): Promise<WorkbookPrintPreviewResult> {
  try {
    let entry = previewWindow?.id === id ? previewWindow : null
    if (!entry) {
      if (request === undefined) return { ok: false, expired: true }
      await releasePrintPreview()
      const workDir = await mkdtemp(join(tmpdir(), 'ai-excel-preview-'))
      const window = new BrowserWindow({
        show: false,
        webPreferences: { sandbox: true, javascript: false },
      })
      entry = { id, window, workDir, request, timer: null }
      previewWindow = entry
      const htmlPath = join(workDir, 'print.html')
      await writeFile(htmlPath, request.html, 'utf8')
      await window.loadFile(htmlPath)
    }
    const live = entry
    if (live.timer) clearTimeout(live.timer)
    live.timer = setTimeout(() => void releasePrintPreview(live.id), PREVIEW_IDLE_MS)
    const full = needsFullRender(live.request)
    let pdf: Buffer
    try {
      pdf = await renderPdf(
        live.window.webContents,
        live.request,
        full ? undefined : `${from}-${to}`,
      )
    } catch (error: unknown) {
      // a range wholly past the last page is rejected by Chromium: that is just "no more pages"
      if (!full && from > 1) return { ok: true, pdf: null, full: false }
      throw error
    }
    return { ok: true, pdf: new Uint8Array(pdf), full }
  } catch (error: unknown) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

/// ---- band-by-band printing of very large sheets ----
/// The renderer lays out and sends one row band at a time (each band's HTML is rendered to its own
/// PDF on disk and dropped), then asks for the merge. Memory stays one band wide however large
/// the sheet is.

interface BandJob {
  readonly dir: string
  readonly files: string[]
}
const bandJobs = new Map<string, BandJob>()

async function bandJobFor(jobId: string): Promise<BandJob> {
  let job = bandJobs.get(jobId)
  if (!job) {
    job = { dir: await mkdtemp(join(tmpdir(), 'ai-excel-bands-')), files: [] }
    bandJobs.set(jobId, job)
  }
  return job
}

export async function cancelPrintBands(jobId: string): Promise<void> {
  const job = bandJobs.get(jobId)
  bandJobs.delete(jobId)
  if (job) await rm(job.dir, { recursive: true, force: true })
}

export async function addPrintBand(
  jobId: string,
  request: WorkbookExportPdfRequest,
): Promise<{ ok: true } | { ok: false; error: string }> {
  try {
    const job = await bandJobFor(jobId)
    const index = job.files.length
    const htmlPath = join(job.dir, `band-${index}.html`)
    const window = new BrowserWindow({
      show: false,
      webPreferences: { sandbox: true, javascript: false },
    })
    try {
      await writeFile(htmlPath, request.html, 'utf8')
      await window.loadFile(htmlPath)
      const pdf = await renderPdf(window.webContents, request)
      const pdfPath = join(job.dir, `band-${index}.pdf`)
      await writeFile(pdfPath, pdf)
      job.files.push(pdfPath)
    } finally {
      window.destroy()
      await rm(htmlPath, { force: true })
    }
    return { ok: true }
  } catch (error: unknown) {
    await cancelPrintBands(jobId)
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

/// Merges the bands into one PDF and prints it. Printing a PDF document goes through Chromium's
/// PDF viewer; when that cannot print, the merged PDF opens in the system viewer instead.
export async function finishPrintBands(
  event: IpcMainInvokeEvent,
  jobId: string,
  job?: PrinterJob & { pageFrom?: number | undefined; pageTo?: number | undefined },
): Promise<WorkbookPrintResult> {
  const entry = bandJobs.get(jobId)
  if (!entry || entry.files.length === 0) return { ok: false, error: 'No pages were prepared.' }
  let keep = false
  try {
    const { PDFDocument: PdfDocument } = await import('pdf-lib')
    const merged = await PdfDocument.create()
    for (const file of entry.files) {
      const part = await PdfDocument.load(await readFile(file))
      for (const page of await merged.copyPages(part, part.getPageIndices())) merged.addPage(page)
    }
    const mergedPath = join(entry.dir, 'print.pdf')
    await writeFile(mergedPath, await merged.save())
    if (usesMacPrintPanel(job)) return await printPdfViaMacPanel(mergedPath)
    const owner = BrowserWindow.fromWebContents(event.sender)
    const window = new BrowserWindow({
      show: false,
      ...(owner && !owner.isDestroyed() ? { parent: owner } : {}),
      ...(process.platform === 'win32'
        ? { width: 900, height: 700, autoHideMenuBar: true, closable: false, skipTaskbar: true }
        : {}),
      webPreferences: { sandbox: true, plugins: true },
    })
    try {
      await window.loadFile(mergedPath)
      if (process.platform === 'win32') {
        window.show()
        window.focus()
      }
      const outcome = await new Promise<{ success: boolean; failureReason: string }>((resolve) => {
        window.webContents.print(
          {
            printBackground: true,
            ...directPrintOptions(job),
            ...(job?.pageFrom === undefined
              ? {}
              : { pageRanges: [{ from: job.pageFrom - 1, to: (job.pageTo ?? 100_000) - 1 }] }),
          },
          (success, failureReason) => resolve({ success, failureReason }),
        )
      })
      if (outcome.success) return { ok: true }
      if (outcome.failureReason === 'Print job canceled') return { ok: false }
      keep = true
    } finally {
      window.destroy()
    }
    // the PDF viewer could not print it: hand the merged PDF to the system viewer
    const failed = await shell.openPath(mergedPath)
    return failed === '' ? { ok: true, openedPdf: mergedPath } : { ok: false, error: failed }
  } catch (error: unknown) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  } finally {
    bandJobs.delete(jobId)
    if (!keep) await rm(entry.dir, { recursive: true, force: true })
  }
}

interface TemplatePair {
  readonly headerTemplate?: string | undefined
  readonly footerTemplate?: string | undefined
}

/// One printToPDF pass; `pageRanges` undefined prints every page. Chromium
/// falls back to its own date/title header when a template is missing, so
/// both templates are always passed once headers/footers are shown.
async function printPass(
  contents: WebContents,
  request: WorkbookExportPdfRequest,
  templates: TemplatePair | undefined,
  pageRanges?: string,
): Promise<Buffer> {
  return contents.printToPDF({
    landscape: request.landscape,
    pageSize: request.pageSize,
    margins: request.margins,
    scale: request.scale,
    printBackground: true,
    ...(pageRanges === undefined ? {} : { pageRanges }),
    ...(templates
      ? {
          displayHeaderFooter: true,
          headerTemplate: templates.headerTemplate ?? '<span></span>',
          footerTemplate: templates.footerTemplate ?? '<span></span>',
        }
      : {}),
  })
}

/// True when the request needs the stitched multi-pass export (first-page / even-page header and
/// footer variants, or page-number offsets) — those cannot be rendered one page range at a time.
export function needsFullRender(request: WorkbookExportPdfRequest): boolean {
  return (
    request.firstPage !== undefined ||
    request.evenPages !== undefined ||
    hasHeaderFooterPageOffset(request.headerTemplate) ||
    hasHeaderFooterPageOffset(request.footerTemplate)
  )
}

/// Chromium prints one header/footer template pair for every page. Excel's
/// differentFirst / differentOddEven need extra passes — page 1 with the
/// first-page templates, the even pages with the even ones — stitched into
/// the odd-page print by page index (pdf-lib). Without variants the odd pass
/// is the whole export (single-pass fast path). `range` (preview only) limits
/// the single-pass path to those pages ("1-8"); variants always render in full.
async function renderPdf(
  contents: WebContents,
  request: WorkbookExportPdfRequest,
  range?: string,
): Promise<Buffer> {
  const oddTemplates: TemplatePair | undefined =
    request.headerTemplate !== undefined || request.footerTemplate !== undefined
      ? { headerTemplate: request.headerTemplate, footerTemplate: request.footerTemplate }
      : undefined
  const flags = {
    hasFirst: request.firstPage !== undefined,
    hasEven: request.evenPages !== undefined,
  }
  const templatesByVariant: Record<PageVariant, TemplatePair | undefined> = {
    odd: oddTemplates,
    first: request.firstPage,
    even: request.evenPages,
  }
  const hasOffsets = Object.values(templatesByVariant).some(
    (pair) =>
      hasHeaderFooterPageOffset(pair?.headerTemplate) ||
      hasHeaderFooterPageOffset(pair?.footerTemplate),
  )
  const showHeaderFooter = oddTemplates !== undefined || flags.hasFirst || flags.hasEven
  const single = !flags.hasFirst && !flags.hasEven && !hasOffsets
  const odd = await printPass(
    contents,
    request,
    showHeaderFooter ? (oddTemplates ?? {}) : undefined,
    single ? range : undefined,
  )
  if (single) return odd

  const { PDFDocument: PdfDocument } = await import('pdf-lib')
  const oddDocument = await PdfDocument.load(odd)
  const total = oddDocument.getPageCount()
  const passes: Partial<Record<PageVariant, PDFDocument>> = { odd: oddDocument }
  if (request.firstPage !== undefined && total >= 1) {
    const first = await printPass(contents, request, request.firstPage, '1')
    passes.first = await PdfDocument.load(first)
  }
  const evenRanges = evenPageRanges(total)
  if (request.evenPages !== undefined && evenRanges !== '') {
    const even = await printPass(contents, request, request.evenPages, evenRanges)
    passes.even = await PdfDocument.load(even)
  }
  const merged = await PdfDocument.create()
  for (const step of stitchPlan(total, flags)) {
    const template = templatesByVariant[step.source]
    if (
      hasHeaderFooterPageOffset(template?.headerTemplate) ||
      hasHeaderFooterPageOffset(template?.footerTemplate)
    ) {
      const onePage = await PdfDocument.load(
        await printPass(
          contents,
          request,
          {
            headerTemplate: headerFooterTemplateForPage(template?.headerTemplate, step.page, total),
            footerTemplate: headerFooterTemplateForPage(template?.footerTemplate, step.page, total),
          },
          String(step.page),
        ),
      )
      if (onePage.getPageCount() !== 1) {
        throw new Error(`Expected one PDF page for header/footer page ${step.page}.`)
      }
      const [page] = await merged.copyPages(onePage, [0])
      merged.addPage(page!)
      continue
    }
    // A pass that came back with fewer pages than planned (Chromium and
    // pdf-lib disagreeing about a range) falls back to the odd print of
    // that page rather than failing the export.
    const source = passes[step.source]
    const [page] =
      source !== undefined && step.index < source.getPageCount()
        ? await merged.copyPages(source, [step.index])
        : await merged.copyPages(oddDocument, [step.page - 1])
    if (page) merged.addPage(page)
  }
  return Buffer.from(await merged.save())
}
