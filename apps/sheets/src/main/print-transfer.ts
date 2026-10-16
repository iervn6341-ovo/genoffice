import type { WorkbookExportPdfRequest } from '../shared/desktop-api'

/// ---- chunked HTML transfer ----
/// The preload ships print HTML larger than one IPC message in chunks; the request that follows
/// names the transfer instead of carrying the text.
const htmlTransfers = new Map<string, { parts: string[]; touched: number }>()

export function storePrintHtmlChunk(id: string, index: number, text: string): void {
  const now = Date.now()
  for (const [key, entry] of htmlTransfers) {
    if (now - entry.touched > 10 * 60_000) htmlTransfers.delete(key)
  }
  const entry = htmlTransfers.get(id) ?? { parts: [], touched: now }
  entry.parts[index] = text
  entry.touched = now
  htmlTransfers.set(id, entry)
}

/// The request with its HTML reassembled from the chunks it names.
export function resolvePrintRequest(request: WorkbookExportPdfRequest): WorkbookExportPdfRequest {
  if (request.htmlTransferId === undefined) {
    if (request.html.length === 0) throw new Error('Empty print request.')
    return request
  }
  const entry = htmlTransfers.get(request.htmlTransferId)
  htmlTransfers.delete(request.htmlTransferId)
  if (
    !entry ||
    entry.parts.length === 0 ||
    Array.from(entry.parts).some((part) => part === undefined)
  ) {
    throw new Error('Print data was not received completely.')
  }
  const { htmlTransferId: _id, ...rest } = request
  return { ...rest, html: entry.parts.join('') }
}
