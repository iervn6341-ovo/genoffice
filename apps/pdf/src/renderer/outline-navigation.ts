import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { OutlineNode } from './OutlinePanel'

export interface OutlineDestination {
  pageIndex: number
  left: number | null
  top: number | null
}

/** Resolve both named bookmarks and explicit page references, including same-page sections. */
export async function resolveOutlineDestination(
  doc: PDFDocumentProxy,
  dest: unknown,
): Promise<OutlineDestination | null> {
  try {
    const arr = typeof dest === 'string' ? await doc.getDestination(dest) : dest
    if (!Array.isArray(arr) || arr.length < 2) return null
    const ref = arr[0]
    const pageIndex = typeof ref === 'number' ? ref : await doc.getPageIndex(ref)
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= doc.numPages) return null
    const mode = arr[1]?.name
    const number = (value: unknown) =>
      typeof value === 'number' && Number.isFinite(value) ? value : null
    return {
      pageIndex,
      left: mode === 'XYZ' || mode === 'FitR' ? number(arr[2]) : null,
      top:
        mode === 'XYZ'
          ? number(arr[3])
          : mode === 'FitH' || mode === 'FitBH'
            ? number(arr[2])
            : mode === 'FitR'
              ? number(arr[5])
              : null,
    }
  } catch {
    return null
  }
}

/** Bounded worker requests and per-document cache; no destination lookup on scroll. */
export async function resolveOutlineDestinations(
  doc: PDFDocumentProxy,
  items: readonly { id: string; node: OutlineNode }[],
  cancelled: () => boolean,
): Promise<(OutlineDestination & { id: string })[]> {
  const targets: (OutlineDestination & { id: string })[] = []
  const cache = new Map<unknown, Promise<OutlineDestination | null>>()
  for (let start = 0; start < items.length; start += 8) {
    if (cancelled()) return []
    const batch = await Promise.all(
      items.slice(start, start + 8).map(async ({ id, node }) => {
        if (node.url || node.dest == null) return null
        const key = node.dest
        let pending = cache.get(key)
        if (!pending) {
          pending = resolveOutlineDestination(doc, node.dest)
          cache.set(key, pending)
        }
        const target = await pending
        return target ? { id, ...target } : null
      }),
    )
    for (const target of batch) if (target) targets.push(target)
  }
  return cancelled() ? [] : targets
}
