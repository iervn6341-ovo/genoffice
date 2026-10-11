/**
 * Print job settings from the renderers' print sheets → webContents.print()
 * options. Pure (no Electron import): the caller spreads the result into its
 * own print call. Without a printer the system dialog opens as before.
 */
export interface PrinterJob {
  printer?: string | undefined
  copies?: number | undefined
  duplex?: 'off' | 'long' | 'short' | undefined
  collate?: boolean | undefined
}

export interface DirectPrintOptions {
  silent: boolean
  deviceName?: string
  copies?: number
  duplexMode?: 'simplex' | 'shortEdge' | 'longEdge'
  collate?: boolean
}

export function directPrintOptions(job: PrinterJob | undefined): DirectPrintOptions {
  if (!job?.printer) return { silent: false }
  const copies = Math.min(Math.max(Math.floor(Number(job.copies ?? 1)) || 1, 1), 999)
  return {
    silent: true,
    deviceName: job.printer,
    copies,
    duplexMode:
      job.duplex === 'long' ? 'longEdge' : job.duplex === 'short' ? 'shortEdge' : 'simplex',
    collate: job.collate !== false,
  }
}

export interface PrinterListEntry {
  name: string
  displayName: string
  isDefault: boolean
}

/** Installed printers (`isDefault` exists at runtime but is missing from Electron's typings) */
export async function listSystemPrinters(contents: {
  getPrintersAsync(): Promise<Array<{ name: string; displayName?: string }>>
}): Promise<PrinterListEntry[]> {
  try {
    const list = (await contents.getPrintersAsync()) as Array<{
      name: string
      displayName?: string
      isDefault?: boolean
    }>
    return list.map((p) => ({
      name: p.name,
      displayName: p.displayName || p.name,
      isDefault: p.isDefault === true,
    }))
  } catch {
    return []
  }
}
