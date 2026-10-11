/**
 * Word's "Pages per sheet" / "Scale to paper size" / reverse order: the printed pages are
 * clones of the pagination preview's .pv-page boxes (same styles, so they render exactly like
 * the preview) laid out on sheets inside the preview root. Print CSS swaps the preview's own
 * page list for this sheet list while it exists (see styles.css .pv-has-sheets).
 */

/** Word's pages-per-sheet grids: columns × rows, filled row by row */
export const PAGES_PER_SHEET = [1, 2, 4, 6, 8, 16] as const
export type PagesPerSheet = (typeof PAGES_PER_SHEET)[number]

export function sheetGrid(n: PagesPerSheet): { cols: number; rows: number } {
  switch (n) {
    case 2:
      return { cols: 2, rows: 1 }
    case 4:
      return { cols: 2, rows: 2 }
    case 6:
      return { cols: 3, rows: 2 }
    case 8:
      return { cols: 4, rows: 2 }
    case 16:
      return { cols: 4, rows: 4 }
    default:
      return { cols: 1, rows: 1 }
  }
}

/** CSS px per twip (96 dpi) */
export const PX_PER_TWIP = 96 / 1440

export interface SheetOptions {
  perSheet: PagesPerSheet
  /** Sheet size in CSS px; undefined = the first page's own size */
  sheetW?: number
  sheetH?: number
}

/** Scale that fits a w×h page into a cell */
export function fitScale(pageW: number, pageH: number, cellW: number, cellH: number): number {
  if (!(pageW > 0) || !(pageH > 0)) return 1
  return Math.min(cellW / pageW, cellH / pageH)
}

/**
 * Builds the sheets for `pages` (already filtered and ordered) and appends them to `host`.
 * Returns the sheet size in CSS px so the caller can hand the paper size to the printer.
 */
export function buildSheets(
  host: HTMLElement,
  pages: readonly HTMLElement[],
  o: SheetOptions,
): { sheetW: number; sheetH: number; count: number } {
  const first = pages[0]
  const baseW = first?.offsetWidth ?? 816
  const baseH = first?.offsetHeight ?? 1056
  // pages per sheet ≥ 2 keeps the sheet orientation of the document's pages
  const sheetW = o.sheetW ?? baseW
  const sheetH = o.sheetH ?? baseH
  const { cols, rows } = sheetGrid(o.perSheet)
  const cellW = sheetW / cols
  const cellH = sheetH / rows
  const wrap = document.createElement('div')
  wrap.className = 'pv-sheets'
  wrap.style.setProperty('--sheet-w', `${sheetW}px`)
  wrap.style.setProperty('--sheet-h', `${sheetH}px`)
  let count = 0
  for (let i = 0; i < pages.length; i += o.perSheet) {
    const sheet = document.createElement('div')
    sheet.className = 'pv-sheet'
    pages.slice(i, i + o.perSheet).forEach((src, k) => {
      const pageW = src.offsetWidth
      const pageH = src.offsetHeight
      const scale = fitScale(pageW, pageH, cellW, cellH)
      const col = k % cols
      const row = Math.floor(k / cols)
      const cell = document.createElement('div')
      cell.className = 'pv-cell'
      cell.style.width = `${pageW * scale}px`
      cell.style.height = `${pageH * scale}px`
      cell.style.left = `${col * cellW + (cellW - pageW * scale) / 2}px`
      cell.style.top = `${row * cellH + (cellH - pageH * scale) / 2}px`
      const clone = src.cloneNode(true) as HTMLElement
      clone.classList.remove('pv-print-skip')
      clone.style.transformOrigin = '0 0'
      clone.style.transform = `scale(${scale})`
      clone.style.width = `${pageW}px`
      clone.style.height = `${pageH}px`
      cell.appendChild(clone)
      sheet.appendChild(cell)
    })
    wrap.appendChild(sheet)
    count++
  }
  host.appendChild(wrap)
  return { sheetW, sheetH, count }
}

/** 1-based page numbers kept by Word's "Only print odd / even pages" */
export function keepParity(index: number, parity: 'all' | 'odd' | 'even'): boolean {
  if (parity === 'all') return true
  const pageNo = index + 1
  return parity === 'odd' ? pageNo % 2 === 1 : pageNo % 2 === 0
}
