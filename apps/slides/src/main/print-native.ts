import { app } from 'electron'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { printPdfWithSystemPanel } from '@genoffice/electron-utils'
import { normalizePrintPrefs, type SlidesPrintPrefs } from '../shared/print-prefs'

export { usesMacPrintPanel } from '../shared/print-prefs'

export type NativePrintResult = { ok: true } | { ok: false; error?: string }

/** Hand a finished PDF to the macOS print panel. Cancel is a normal completion (no error). */
export async function printPdfViaMacPanel(
  pdf: Uint8Array,
  title: string,
): Promise<NativePrintResult> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-slides-nativeprint-'))
  try {
    const pdfPath = join(dir, 'print.pdf')
    await writeFile(pdfPath, pdf, { mode: 0o600 })
    const outcome = await printPdfWithSystemPanel(pdfPath, title, {
      resourcesPath: app.isPackaged ? process.resourcesPath : undefined,
      devDir: join(app.getAppPath(), '../../packages/electron-utils/native'),
    })
    if (outcome.status === 'printed') return { ok: true }
    if (outcome.status === 'canceled') return { ok: false }
    return { ok: false, error: outcome.message }
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

// ── persisted preferences (the File menu owns them on macOS) ──

let cache: SlidesPrintPrefs | null = null
const prefsPath = () => join(app.getPath('userData'), 'slides-print-prefs.json')

export async function loadPrintPrefs(): Promise<SlidesPrintPrefs> {
  if (cache) return cache
  try {
    cache = normalizePrintPrefs(JSON.parse(await readFile(prefsPath(), 'utf8')))
  } catch {
    cache = normalizePrintPrefs(null)
  }
  return cache
}

/** Synchronous view for building the menu (loaded at startup; defaults until then) */
export function currentPrintPrefs(): SlidesPrintPrefs {
  return cache ?? normalizePrintPrefs(null)
}

export async function updatePrintPrefs(
  patch: Partial<SlidesPrintPrefs>,
): Promise<SlidesPrintPrefs> {
  const next = normalizePrintPrefs({ ...(await loadPrintPrefs()), ...patch })
  cache = next
  try {
    await writeFile(prefsPath(), JSON.stringify(next), 'utf8')
  } catch {
    // preferences are a convenience; a read-only profile must not break printing
  }
  return next
}
