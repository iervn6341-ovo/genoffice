import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

/** Load large print HTML from a private temporary file instead of a length-limited data URL. */
export async function loadPrintHtmlFile(
  window: { loadFile(path: string): Promise<void> },
  html: string,
): Promise<() => Promise<void>> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-slides-print-'))
  try {
    const path = join(dir, 'print.html')
    await writeFile(path, html, { encoding: 'utf8', mode: 0o600 })
    await window.loadFile(path)
    return () => rm(dir, { recursive: true, force: true })
  } catch (error) {
    await rm(dir, { recursive: true, force: true })
    throw error
  }
}

/** Electron errors can quote the entire embedded page; never show its image bytes in UI. */
export function printErrorMessage(error: unknown): string {
  return String(error)
    .replace(/data:[A-Za-z0-9+;/=,_:%.-]+/g, '[embedded document]')
    .slice(0, 512)
}
