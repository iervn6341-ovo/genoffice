/**
 * macOS print panel with a page preview.
 *
 * Electron's `webContents.print()` opens Chromium's own system panel, which has no preview. The
 * `print-pdf` helper (native/print-pdf.swift) drives NSPrintOperation through PDFKit instead, so the
 * standard macOS panel appears with its preview pane, the PDF menu, presets and AirPrint. Callers
 * render the document to a PDF first and hand the file to the helper.
 *
 * Packaged apps ship the compiled helper at Resources/print/print-pdf; a source checkout compiles
 * it with swiftc on first use (and again when the source is newer).
 */
import { execFile, spawn } from 'node:child_process'
import { existsSync, statSync } from 'node:fs'
import { join } from 'node:path'

export type SystemPrintOutcome =
  | { readonly status: 'printed' }
  | { readonly status: 'canceled' }
  | { readonly status: 'error'; readonly message: string }

/** The helper's exit code → outcome (see print-pdf.swift). */
export function outcomeForExit(code: number | null, stderr: string): SystemPrintOutcome {
  if (code === 0) return { status: 'printed' }
  if (code === 1) return { status: 'canceled' }
  const detail = stderr.trim()
  const reason =
    code === 2 ? 'the PDF could not be opened' : code === 3 ? 'the print job could not start' : ''
  return {
    status: 'error',
    message: detail || reason || `the print helper exited with code ${String(code)}`,
  }
}

export interface HelperLocation {
  /** `process.resourcesPath` of a packaged app, when there is one */
  readonly resourcesPath?: string | undefined
  /** folder holding print-pdf.swift (and the binary compiled from it) in a source checkout */
  readonly devDir?: string | undefined
}

export interface HelperPlan {
  readonly binary: string
  /** present when the binary must be (re)built from this source first */
  readonly compileFrom?: string
}

/** Which binary to run, and whether it has to be compiled first. Null: no helper available. */
export function planHelper(location: HelperLocation): HelperPlan | null {
  if (location.resourcesPath) {
    const shipped = join(location.resourcesPath, 'print', 'print-pdf')
    if (existsSync(shipped)) return { binary: shipped }
  }
  if (!location.devDir) return null
  const source = join(location.devDir, 'print-pdf.swift')
  const binary = join(location.devDir, 'print-pdf')
  if (!existsSync(source)) return existsSync(binary) ? { binary } : null
  const stale = !existsSync(binary) || statSync(binary).mtimeMs < statSync(source).mtimeMs
  return stale ? { binary, compileFrom: source } : { binary }
}

function compile(source: string, binary: string): Promise<void> {
  return new Promise((resolve, reject) => {
    execFile('swiftc', ['-O', source, '-o', binary], (error, _stdout, stderr) =>
      error ? reject(new Error(stderr.trim() || error.message)) : resolve(),
    )
  })
}

/**
 * Shows the macOS print panel for a PDF and resolves once the panel is dismissed. 'error' also
 * covers a missing helper, so the caller can fall back to another route.
 */
export async function printPdfWithSystemPanel(
  pdfPath: string,
  title: string,
  location: HelperLocation,
): Promise<SystemPrintOutcome> {
  if (process.platform !== 'darwin') {
    return { status: 'error', message: 'the macOS print panel is only available on macOS' }
  }
  const plan = planHelper(location)
  if (!plan) return { status: 'error', message: 'the macOS print helper was not found' }
  try {
    if (plan.compileFrom) await compile(plan.compileFrom, plan.binary)
  } catch (error: unknown) {
    return {
      status: 'error',
      message: `could not build the macOS print helper: ${error instanceof Error ? error.message : String(error)}`,
    }
  }
  return new Promise((resolve) => {
    const child = spawn(plan.binary, [pdfPath, title], { stdio: ['ignore', 'ignore', 'pipe'] })
    let stderr = ''
    child.stderr.on('data', (chunk: Buffer) => {
      // AppKit logs layout warnings on stderr even on success; only the tail is useful in an error
      stderr = (stderr + chunk.toString()).slice(-2000)
    })
    child.on('error', (error) => resolve({ status: 'error', message: error.message }))
    child.on('close', (code) =>
      resolve(outcomeForExit(code, code === 0 || code === 1 ? '' : stderr)),
    )
  })
}
