/**
 * Print page assembly shared by the main process (hidden print window) and the
 * renderer's print-preview pane, so the preview is pixel-identical to the printout.
 * Pure string building — no Electron or DOM imports.
 */

export type PrintLayout =
  | 'full'
  | 'notes'
  | 'outline'
  | 'handout1'
  | 'handout2'
  | 'handout3'
  | 'handout4h'
  | 'handout4v'
  | 'handout6'
  | 'handout6h'
  | 'handout6v'
  | 'handout9h'
  | 'handout9v'
export type PrintOrientation = 'portrait' | 'landscape'
/** PowerPoint's "Output": color, grayscale, or pure black & white */
export type PrintOutput = 'color' | 'grayscale' | 'bw'
/** Paper for handout / notes pages (full-slide pages always follow the slide size) */
export type PrintPaper =
  'slide' | 'a4' | 'a3' | 'a5' | 'b4' | 'b5' | 'letter' | 'legal' | 'tabloid' | 'executive'
/** Slides per sheet for the full-slide layout (the system print panel's "Pages per sheet") */
export const PRINT_PER_SHEET = [1, 2, 4, 6, 9, 16] as const
export type PrintPerSheet = (typeof PRINT_PER_SHEET)[number]
/** Fill order of a multi-up sheet: Z (rows, left→right), reverse Z, N (columns, top→bottom), reverse N */
export type PrintDirection = 'z' | 'rz' | 'n' | 'rn'

export const PRINT_PAPERS: Record<PrintPaper, { w: number; h: number; label: string }> = {
  // 'slide' = the page is exactly slide-sized (full-page layout only); sized like A4 elsewhere
  slide: { w: 8.27, h: 11.69, label: '' },
  a4: { w: 8.27, h: 11.69, label: 'A4' },
  a3: { w: 11.69, h: 16.54, label: 'A3' },
  a5: { w: 5.83, h: 8.27, label: 'A5' },
  b4: { w: 9.84, h: 13.9, label: 'B4' },
  b5: { w: 6.93, h: 9.84, label: 'B5' },
  letter: { w: 8.5, h: 11, label: 'Letter' },
  legal: { w: 8.5, h: 14, label: 'Legal' },
  tabloid: { w: 11, h: 17, label: 'Tabloid' },
  executive: { w: 7.25, h: 10.5, label: 'Executive' },
}

/** Thumbnails per handout page and whether they fill column-first ("Vertical") */
export function handoutSpec(layout: PrintLayout): { per: number; vertical: boolean } | null {
  switch (layout) {
    case 'handout1':
      return { per: 1, vertical: false }
    case 'handout2':
      return { per: 2, vertical: false }
    case 'handout3':
      return { per: 3, vertical: false }
    case 'handout4h':
      return { per: 4, vertical: false }
    case 'handout4v':
      return { per: 4, vertical: true }
    case 'handout6':
    case 'handout6h':
      return { per: 6, vertical: false }
    case 'handout6v':
      return { per: 6, vertical: true }
    case 'handout9h':
      return { per: 9, vertical: false }
    case 'handout9v':
      return { per: 9, vertical: true }
    default:
      return null
  }
}

/** One slide's text for the outline layout */
export interface PrintOutlineItem {
  /** 1-based slide number in the deck */
  number: number
  title: string
  lines: string[]
}

/** One comment for the comment pages */
export interface PrintCommentItem {
  number: number
  author: string
  text: string
}

export interface PrintDocOptions {
  /** One image URL per slide, in print order (data: or blob: URLs) */
  srcs: string[]
  /** Slide aspect ratio (width / height); sets the page size of the 'full' layout */
  ratio: number
  layout: PrintLayout
  /** Per-slide speaker notes for the 'notes' layout (same order as srcs) */
  notes?: string[]
  /** Paper orientation for handouts/notes pages ('full' pages always follow the slide ratio) */
  orientation?: PrintOrientation
  /** Draw a thin border around full-page slides (handout/notes thumbnails are always framed) */
  frame?: boolean
  /** Color handling of the slide images (default color) */
  output?: PrintOutput
  /** Paper size of handout/notes/multi-up pages (default A4) */
  paper?: PrintPaper
  /** Full layout only: slides per sheet (default 1 = one slide per page) */
  perSheet?: PrintPerSheet
  /** Fill order for perSheet > 1 (default Z) */
  direction?: PrintDirection
  /** Rotate every page content by 180° */
  rotate180?: boolean
  /** Mirror every page content horizontally */
  flipH?: boolean
  /** Outline layout: the text of each printed slide */
  outline?: PrintOutlineItem[]
  /** Comment pages appended after the slides (heading text in `commentsTitle`) */
  comments?: PrintCommentItem[]
  commentsTitle?: string
  /** Full layout on a real paper: scale the slide up to the paper (off = natural size, centered) */
  scaleFit?: boolean
  /** Preview mode: page gaps/shadows + page-number badges instead of a printable document */
  preview?: boolean
}

/** "1,3,5-8" → 0-based slide indices (1-based input over the whole deck); null = invalid */
export function parsePrintRange(text: string, max: number): number[] | null {
  const out = new Set<number>()
  const tokenRe = /\d+\s*[-–—~～〜]\s*\d+|\d+/g
  const parts = text.match(tokenRe) ?? []
  if (parts.length === 0) return null
  if (text.replace(tokenRe, '').replace(/[,，、;；\s]/g, '') !== '') return null
  for (const part of parts) {
    const m = /^(\d+)\s*[-–—~～〜]\s*(\d+)$|^(\d+)$/.exec(part)
    if (!m) return null
    const a = Number(m[1] ?? m[3])
    const b = Number(m[2] ?? m[3])
    if (a < 1 || b > max || a > b) return null
    for (let i = a; i <= b; i++) out.add(i - 1)
  }
  return [...out].sort((x, y) => x - y)
}

export function printPageCount(
  slideCount: number,
  layout: PrintLayout,
  perSheet: number = 1,
): number {
  if (layout === 'outline') return Math.max(1, Math.ceil(slideCount / 10))
  const perPage = handoutSpec(layout)?.per ?? (layout === 'full' ? perSheet : 1)
  return Math.ceil(slideCount / perPage)
}

/** A4 in inches; slide height for the 'full' layout follows the PDF export (7.5in) */
const SLIDE_H = 7.5

/** [columns, rows] of an N-up sheet */
function gridFor(n: number, landscape: boolean): [number, number] {
  const side = Math.round(Math.sqrt(n))
  if (side * side === n) return [side, side]
  // 2 → 1×2, 6 → 2×3 on portrait; transposed on landscape
  const [a, b] = n === 2 ? [1, 2] : [2, 3]
  return landscape ? [b, a] : [a, b]
}

/** [columns, rows] of a handout page */
function handoutGrid(per: number, landscape: boolean): [number, number] {
  if (per === 1) return [1, 1]
  if (per === 2) return landscape ? [2, 1] : [1, 2]
  if (per === 4) return [2, 2]
  if (per === 9) return [3, 3]
  // 3 stays a single column; 6 is 2×3 portrait, 3×2 landscape
  if (per === 3) return [1, 3]
  return landscape ? [3, 2] : [2, 3]
}

/**
 * Fallback slide ratio (16:9) when layout state carries a corrupt ratio.
 * Ratios come from slide dimensions that can be zeroed by a failed load or
 * NaN/Infinity from a degenerate transform; emitting those into @page yields
 * NaNin/Infinityin/0in and a broken print. Bounds [0.2, 5] keep real formats
 * (4:3, 16:9, 16:10, portrait variants) exact while clamping absurd values.
 */
export const DEFAULT_PRINT_RATIO = 16 / 9
export const MIN_PRINT_RATIO = 0.2
export const MAX_PRINT_RATIO = 5

export function normalizePrintRatio(ratio: number): number {
  if (!Number.isFinite(ratio) || ratio <= 0) return DEFAULT_PRINT_RATIO
  return Math.min(Math.max(ratio, MIN_PRINT_RATIO), MAX_PRINT_RATIO)
}

export function buildPrintDocumentHtml(o: PrintDocOptions): string {
  const layout = o.layout
  const nUp = layout === 'full' ? (o.perSheet ?? 1) : 1
  // a plain full page is exactly slide-sized; multi-up sheets are paper pages
  const realPaper = o.paper !== undefined && o.paper !== 'slide'
  const isFull = layout === 'full' && nUp === 1 && !realPaper
  // one slide per paper page (a real paper chosen for the full layout)
  const onPaper = layout === 'full' && nUp === 1 && realPaper
  const landscape = !isFull && o.orientation === 'landscape'
  const slideW = Math.round(normalizePrintRatio(o.ratio) * SLIDE_H * 1000) / 1000
  const paper = PRINT_PAPERS[o.paper ?? 'a4'] ?? PRINT_PAPERS.a4
  const pageW = isFull ? slideW : landscape ? paper.h : paper.w
  const pageH = isFull ? SLIDE_H : landscape ? paper.w : paper.h
  const spec = handoutSpec(layout)
  const perPage = spec?.per ?? 1
  // 2 / 3 / 6 slides filled row-first keep the dedicated handout styles; the rest use a grid
  const gridHandout = spec !== null && (spec.vertical || ![2, 3, 6].includes(spec.per))
  const esc = (x: string) =>
    x.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!)
  const escAttr = (x: string) =>
    x.replace(
      /[&<>"']/g,
      (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!,
    )

  let body: string
  if (layout === 'full' && nUp > 1) {
    // Multi-up sheet: a cols×rows grid, cells placed by the chosen fill order
    const [cols, rows] = gridFor(nUp, landscape)
    const dir = o.direction ?? 'z'
    const pages: string[] = []
    for (let i = 0; i < o.srcs.length; i += nUp) {
      const cells = o.srcs
        .slice(i, i + nUp)
        .map((src, k) => {
          const major = dir === 'z' || dir === 'rz'
          let r = major ? Math.floor(k / cols) : k % rows
          let c = major ? k % cols : Math.floor(k / rows)
          if (dir === 'rz' || dir === 'rn') c = cols - 1 - c
          r = Math.min(r, rows - 1)
          return `<div class="cell" style="grid-row:${r + 1};grid-column:${c + 1}"><img src="${escAttr(src)}"></div>`
        })
        .join('')
      pages.push(`<div class="page nup"><div class="sheet">${cells}</div></div>`)
    }
    body = pages.join('')
  } else if (isFull) {
    body = o.srcs.map((src) => `<div class="page"><img src="${escAttr(src)}"></div>`).join('')
  } else if (onPaper) {
    body = o.srcs
      .map(
        (src) =>
          `<div class="page onpaper${o.scaleFit ? ' fit' : ''}"><img src="${escAttr(src)}"></div>`,
      )
      .join('')
  } else if (layout === 'outline') {
    const items = o.outline ?? []
    body =
      '<div class="page flow outline">' +
      items
        .map(
          (it) =>
            `<div class="ol-slide"><div class="ol-num">${it.number}</div><div class="ol-body">` +
            `<div class="ol-title">${esc(it.title)}</div>` +
            it.lines.map((l) => `<div class="ol-line">${esc(l)}</div>`).join('') +
            '</div></div>',
        )
        .join('') +
      '</div>'
  } else if (layout === 'notes') {
    // Notes page: slide on top + notes text below
    body = o.srcs
      .map(
        (src, i) =>
          `<div class="page notes"><img src="${escAttr(src)}">` +
          `<div class="note">${esc(o.notes?.[i] ?? '').replace(/\n/g, '<br>')}</div></div>`,
      )
      .join('')
  } else {
    // Handouts: perPage thumbnails per page (with 3, ruled lines on the right for handwriting)
    const pages: string[] = []
    const [hCols, hRows] = handoutGrid(perPage, landscape)
    for (let i = 0; i < o.srcs.length; i += perPage) {
      const cells = o.srcs
        .slice(i, i + perPage)
        .map((src, k) => {
          const place = gridHandout
            ? ` style="grid-row:${(spec!.vertical ? k % hRows : Math.floor(k / hCols)) + 1};grid-column:${(spec!.vertical ? Math.floor(k / hRows) : k % hCols) + 1}"`
            : ''
          return (
            `<div class="cell"${place}><img src="${escAttr(src)}">` +
            (perPage === 3 ? '<div class="rules"></div>' : '') +
            '</div>'
          )
        })
        .join('')
      const grid = gridHandout
        ? ` hg" style="grid-template-columns:repeat(${hCols},1fr);grid-template-rows:repeat(${hRows},1fr)`
        : ''
      pages.push(`<div class="page handout h${perPage}${grid}">${cells}</div>`)
    }
    body = pages.join('')
  }

  if (o.comments && o.comments.length > 0) {
    body +=
      '<div class="page flow cmts">' +
      `<h2>${esc(o.commentsTitle ?? 'Comments')}</h2>` +
      o.comments
        .map(
          (c) =>
            `<div class="cm"><b>${c.number}</b> <span class="cm-by">${esc(c.author)}</span>` +
            `<div class="cm-text">${esc(c.text).replace(/\n/g, '<br>')}</div></div>`,
        )
        .join('') +
      '</div>'
  }

  // h6 grid: 2×3 portrait, 3×2 landscape
  const h6Cols = landscape ? '1fr 1fr 1fr' : '1fr 1fr'
  const outputCss =
    o.output === 'grayscale'
      ? '.page img { filter: grayscale(1); }'
      : o.output === 'bw'
        ? '.page img { filter: grayscale(1) contrast(1000%); }'
        : ''
  const frameCss = o.frame
    ? isFull
      ? '.page > img { border: 1px solid #bbb; }'
      : nUp > 1
        ? '.page.nup .cell img { border: 1px solid #333; }'
        : ''
    : ''
  const transforms = [o.rotate180 ? 'rotate(180deg)' : '', o.flipH ? 'scaleX(-1)' : '']
    .filter(Boolean)
    .join(' ')
  const gridCols = nUp > 1 ? gridFor(nUp, landscape) : [1, 1]
  const nupCss =
    nUp > 1
      ? `.page.nup { padding: 0.4in; }
.page.nup .sheet { width: 100%; height: 100%; display: grid; gap: 0.15in; grid-template-columns: repeat(${gridCols[0]}, 1fr); grid-template-rows: repeat(${gridCols[1]}, 1fr); }
.page.nup .cell { display: flex; align-items: center; justify-content: center; min-width: 0; min-height: 0; }
.page.nup .cell img { display: block; max-width: 100%; max-height: 100%; width: auto; height: auto; }`
      : ''
  const transformCss = transforms
    ? `.page > img, .page.handout, .page.notes, .page.nup .sheet { transform: ${transforms}; }`
    : ''
  // flowing text pages (outline / comments) are always paper pages
  const flowLandscape = o.orientation === 'landscape' && !isFull
  const flowW = flowLandscape ? paper.h : paper.w
  const flowH = flowLandscape ? paper.w : paper.h

  // Preview chrome (never printed): gray gaps, page shadows, PowerPoint-style page badges
  const total = printPageCount(o.srcs.length, layout, nUp)
  const previewCss = o.preview
    ? `
body { counter-reset: pg; background: transparent; padding: 18px 0 6px; }
.page { counter-increment: pg; position: relative; margin: 0 auto 8px; background: #fff; box-shadow: 0 1px 5px rgba(0, 0, 0, 0.35); }
.page.flow { width: ${flowW}in; padding: 0.6in; box-sizing: border-box; min-height: ${flowH}in; }
.page.flow::after { display: none; }
.page::after {
  content: counter(pg) ' / ${total}';
  position: absolute; left: 50%; bottom: 3%; transform: translateX(-50%) ${transforms};
  padding: 3px 12px; border-radius: 999px;
  background: rgba(32, 33, 36, 0.65); color: #fff;
  font-size: 10px; white-space: nowrap;
}
`
    : ''

  return `<!doctype html><html><head><meta charset="utf-8"><style>
@page { size: ${pageW}in ${pageH}in; margin: 0; }
@page flow { size: ${flowW}in ${flowH}in; margin: 0.6in; }
html, body { margin: 0; padding: 0; font-family: -apple-system, 'Segoe UI', sans-serif; }
.page { width: ${pageW}in; height: ${pageH}in; overflow: hidden; page-break-after: always; box-sizing: border-box; }
.page:last-of-type { page-break-after: auto; }
.page > img { display: block; width: 100%; height: 100%; }
.page.handout { padding: 0.4in; display: flex; flex-direction: ${landscape && perPage === 2 ? 'row' : 'column'}; gap: 0.24in; }
.page.handout .cell { display: flex; gap: 0.2in; align-items: center; flex: 1; min-width: 0; min-height: 0; }
.page.handout .cell img { border: 1px solid #bbb; object-fit: contain; max-height: 100%; }
.page.handout.h2 .cell img, .page.handout.h6 .cell img { width: 100%; height: auto; max-height: 100%; }
.page.handout.h3 .cell img { width: 55%; height: auto; }
.page.handout.h3 .rules {
  flex: 1; align-self: stretch;
  background: repeating-linear-gradient(#fff 0 0.28in, #ccc 0.28in calc(0.28in + 1px));
}
.page.handout.h6 { display: grid; grid-template-columns: ${h6Cols}; grid-auto-rows: 1fr; }
.page.handout.hg { display: grid; gap: 0.24in; }
.page.handout.hg .cell { justify-content: center; }
.page.handout.hg .cell img { width: 100%; height: auto; max-height: 100%; }
.page.onpaper { display: flex; align-items: center; justify-content: center; padding: 0.4in; }
.page.onpaper > img { display: block; width: ${slideW}in; height: ${SLIDE_H}in; max-width: 100%; max-height: 100%; object-fit: contain; }
.page.onpaper.fit > img { width: 100%; height: auto; }
.page.flow { page: flow; width: auto; height: auto; min-height: 0; overflow: visible; page-break-before: always; }
.page.flow:first-child { page-break-before: auto; }
.page.outline .ol-slide { display: flex; gap: 0.25in; margin: 0 0 0.18in; break-inside: avoid; }
.page.outline .ol-num { width: 0.4in; text-align: right; font-weight: 700; font-size: 12pt; color: #555; flex: none; }
.page.outline .ol-title { font-weight: 700; font-size: 13pt; margin-bottom: 0.04in; }
.page.outline .ol-line { font-size: 11pt; line-height: 1.45; padding-left: 0.2in; }
.page.cmts h2 { font-size: 15pt; margin: 0 0 0.2in; }
.page.cmts .cm { margin: 0 0 0.14in; break-inside: avoid; font-size: 11pt; }
.page.cmts .cm-by { color: #555; }
.page.cmts .cm-text { margin-top: 0.03in; white-space: normal; }
.page.notes { padding: 0.5in; display: flex; flex-direction: column; }
.page.notes img { width: ${landscape ? 'auto' : '100%'}; ${landscape ? 'max-height: 55%; align-self: center;' : 'height: auto;'} border: 1px solid #bbb; }
.page.notes .note { margin-top: 0.3in; font-size: 11pt; line-height: 1.5; white-space: pre-wrap; }
${nupCss}
${frameCss}
${transformCss}
${outputCss}
${previewCss}
</style></head><body>${body}</body></html>`
}
