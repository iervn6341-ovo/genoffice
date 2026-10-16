/**
 * Print page assembly shared by the main process (hidden print window) and the
 * renderer's print-preview pane, so the preview is pixel-identical to the printout.
 * Pure string building — no Electron or DOM imports.
 */
import type {
  PrintHeaderFooterField,
  PrintHeaderFooterPage,
  PrintNotesParagraph,
  PrintPageBox,
} from '@genoffice/pptx-engine'
import { formatHeaderFooterDate } from './header-footer-date'

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
  /** Native master fields, resolved against each notes page's local placeholders. */
  notesFields?: PrintHeaderFooterPage[]
  handoutFields?: PrintHeaderFooterPage
  /** Rasterized native master background and non-placeholder artwork, keyed by part path. */
  masterArtwork?: Record<string, string>
  /** Original zero-based indexes of srcs (for notes-page number fields). */
  slideIndexes?: number[]
  /** Slide-number origin; handout page numbers always start at 1. */
  firstSlideNum?: number
  dateLocale?: string
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

  const fieldText = (field: PrintHeaderFooterField | undefined, number?: number): string => {
    if (!field) return ''
    if (field.field === 'slidenum') return String(number ?? 1)
    if (field.field?.startsWith('datetime'))
      return formatHeaderFooterDate(new Date(), field.field, o.dateLocale ?? 'en-US') ?? field.text
    return field.text
  }
  const fieldHtml = (
    field:
      | (PrintHeaderFooterField & {
          strike?: boolean
          baseline?: number
          highlight?: string
          letterSpacingPt?: number
        })
      | undefined,
    number?: number,
  ): string => {
    const text = fieldText(field, number)
    if (!text) return ''
    const style = [
      field?.fontSizePt != null && Number.isFinite(field.fontSizePt) && field.fontSizePt > 0
        ? `font-size:${field.fontSizePt}pt`
        : '',
      field?.color && /^#[0-9A-Fa-f]{6}(?:[0-9A-Fa-f]{2})?$/.test(field.color)
        ? `color:${field.color}`
        : '',
      field?.bold ? 'font-weight:700' : '',
      field?.italic ? 'font-style:italic' : '',
      field?.underline || field?.strike
        ? `text-decoration:${[field.underline && 'underline', field.strike && 'line-through'].filter(Boolean).join(' ')}`
        : '',
      field?.fontFamily ? `font-family:${JSON.stringify(field.fontFamily)}` : '',
      field?.baseline ? `vertical-align:${field.baseline > 0 ? 'super' : 'sub'}` : '',
      field?.highlight && /^#[0-9A-Fa-f]{6}$/.test(field.highlight)
        ? `background-color:${field.highlight}`
        : '',
      field?.letterSpacingPt != null ? `letter-spacing:${field.letterSpacingPt}pt` : '',
    ]
      .filter(Boolean)
      .join(';')
    return `<span${style ? ` style="${escAttr(style)}"` : ''}>${esc(text)}</span>`
  }
  const boxTransform = (box: PrintPageBox | undefined): string => {
    if (!box) return ''
    const transforms = [
      box.rotationDeg ? `rotate(${box.rotationDeg}deg)` : '',
      box.flipH ? 'scaleX(-1)' : '',
      box.flipV ? 'scaleY(-1)' : '',
    ].filter(Boolean)
    return transforms.length ? `;transform:${transforms.join(' ')}` : ''
  }
  const masterArtwork = (fields: PrintHeaderFooterPage | undefined): string => {
    const src = fields?.masterPath ? o.masterArtwork?.[fields.masterPath] : undefined
    return src ? `<img class="master-artwork" src="${escAttr(src)}" alt="">` : ''
  }
  const notesHtml = (fields: PrintHeaderFooterPage | undefined, fallback: string): string => {
    const paragraphs = fields?.notesParagraphs
    if (!paragraphs?.length) return esc(fallback).replace(/\n/g, '<br>')
    const numbered = new Map<number, number>()
    const roman = (number: number): string => {
      let value = Math.max(1, Math.min(3999, number))
      let out = ''
      for (const [digit, glyph] of [
        [1000, 'M'],
        [900, 'CM'],
        [500, 'D'],
        [400, 'CD'],
        [100, 'C'],
        [90, 'XC'],
        [50, 'L'],
        [40, 'XL'],
        [10, 'X'],
        [9, 'IX'],
        [5, 'V'],
        [4, 'IV'],
        [1, 'I'],
      ] as const) {
        while (value >= digit) {
          out += glyph
          value -= digit
        }
      }
      return out
    }
    const bulletText = (paragraph: PrintNotesParagraph): string => {
      const bullet = paragraph.bullet
      if (!bullet || bullet.type === 'blip') return ''
      if (bullet.type === 'char') return bullet.char ?? '•'
      const level = paragraph.level ?? 0
      const number = bullet.startAt ?? (numbered.get(level) ?? 0) + 1
      numbered.set(level, number)
      const kind = bullet.numberType ?? 'arabicPeriod'
      let digits = String(number)
      if (/roman/i.test(kind))
        digits = /Uc/.test(kind) ? roman(number) : roman(number).toLowerCase()
      else if (/alpha/i.test(kind))
        digits = String.fromCharCode((/Uc/.test(kind) ? 65 : 97) + ((number - 1) % 26))
      return `${/ParenBoth/.test(kind) ? '(' : ''}${digits}${/ParenR|ParenBoth/.test(kind) ? ')' : /Period/.test(kind) ? '.' : ''}`
    }
    const tabbedRuns = (paragraph: PrintNotesParagraph, marker: string, basePt: number): string => {
      if (!paragraph.runs.some((run) => run.text.includes('\t')))
        return marker + paragraph.runs.map((run) => fieldHtml(run)).join('')
      type NoteRun = PrintNotesParagraph['runs'][number]
      const segments: NoteRun[][] = [[]]
      for (const run of paragraph.runs) {
        const pieces = run.text.split('\t')
        pieces.forEach((piece, index) => {
          if (index) segments.push([])
          if (piece) segments.at(-1)!.push({ ...run, text: piece })
        })
      }
      const stops = (paragraph.tabStops ?? paragraph.tabStopsPt?.map((posPt) => ({ posPt })) ?? [])
        .filter((stop) => Number.isFinite(stop.posPt) && stop.posPt >= 0)
        .sort((a, b) => a.posPt - b.posPt)
      const defaultPt = Math.max(1, paragraph.defaultTabPt || 72)
      const side = paragraph.rtl ? 'right' : 'left'
      const shift = (align: string | undefined): string => {
        if (align === 'ctr') return paragraph.rtl ? 'translateX(50%)' : 'translateX(-50%)'
        if (paragraph.rtl) return align === 'l' ? 'translateX(100%)' : ''
        return align === 'r' ? 'translateX(-100%)' : ''
      }
      const decimalParts = (runs: NoteRun[]): [NoteRun[], NoteRun[]] => {
        const text = runs.map((run) => run.text).join('')
        const at = text.search(/[.,]/)
        if (at < 0) return [runs, []]
        const before: NoteRun[] = []
        const after: NoteRun[] = []
        let offset = 0
        for (const run of runs) {
          const split = Math.max(0, Math.min(run.text.length, at - offset))
          if (split) before.push({ ...run, text: run.text.slice(0, split) })
          if (split < run.text.length) after.push({ ...run, text: run.text.slice(split) })
          offset += run.text.length
        }
        return [before, after]
      }
      // The print HTML is also loaded in a script-disabled preview iframe. Resolve
      // tabs here, using the run font sizes and a conservative advance estimate,
      // rather than placing every segment at the next stop by segment index.
      const runWidth = (run: NoteRun): number => {
        const size = run.fontSizePt ?? 11
        let em = 0
        for (const ch of run.text) {
          if (/[\p{Mn}\p{Me}\p{Cf}]/u.test(ch)) continue
          if (/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(ch))
            em += 1
          else if (/[iIlj!|.,:;']/u.test(ch)) em += 0.3
          else if (/[ ftr]/u.test(ch)) em += 0.36
          else if (/[mwMW]/u.test(ch)) em += 0.85
          else if (/[A-Z0-9]/u.test(ch)) em += 0.62
          else if (ch === '_') em += 0.55
          else em += 0.65
        }
        return (
          em * size * (run.bold ? 1.16 : 1.12) +
          Math.max(0, run.letterSpacingPt ?? 0) * [...run.text].length
        )
      }
      const runsWidth = (runs: NoteRun[]): number =>
        runs.reduce((sum, run) => sum + runWidth(run), 0)
      const margin = paragraph.rtl ? (paragraph.marginRightPt ?? 0) : (paragraph.marginLeftPt ?? 0)
      let cursorEnd = 0
      let lastStop = 0
      let nextNative = 0
      const placed = segments.map((runs, index) => {
        const width = runsWidth(runs) + (index === 0 && marker ? basePt * 1.5 : 0)
        const [before, after] = decimalParts(runs)
        const beforeWidth = runsWidth(before)
        const afterWidth = runsWidth(after)
        const bounds = (pos: number, align?: string): [number, number] => {
          if (align === 'ctr') return [pos - width / 2, pos + width / 2]
          if (align === 'dec') return [pos - beforeWidth, pos + afterWidth]
          if (align === (paragraph.rtl ? 'l' : 'r')) return [pos - width, pos]
          return [pos, pos + width]
        }
        let stop: { posPt: number; align?: string } = { posPt: 0 }
        let pos = 0
        if (index) {
          while (nextNative < stops.length) {
            const candidate: { posPt: number; align?: string } = stops[nextNative++]!
            const candidatePos = Math.max(0, candidate.posPt - margin)
            if (candidatePos <= lastStop + 0.5) continue
            if (bounds(candidatePos, candidate.align)[0] < cursorEnd + 1) continue
            stop = candidate
            pos = candidatePos
            break
          }
          if (!pos) {
            // Native stops are exhausted or consumed by earlier text. Continue on
            // the paragraph's default grid until the entire aligned run fits.
            pos = (Math.floor(Math.max(lastStop, cursorEnd) / defaultPt) + 1) * defaultPt
            while (bounds(pos)[0] < cursorEnd + 1) pos += defaultPt
            stop = { posPt: pos + margin, align: paragraph.rtl ? 'r' : 'l' }
          }
        }
        cursorEnd = Math.max(cursorEnd, bounds(pos, stop.align)[1])
        lastStop = pos
        const position = `${side}:${pos}pt`
        if (stop.align === 'dec') {
          const beforeShift = paragraph.rtl
            ? 'transform:translateX(100%);'
            : 'transform:translateX(-100%);'
          return `<span class="note-tab-segment note-tab-decimal-before" style="position:absolute;${position};${beforeShift}white-space:pre">${index ? '' : marker}${before.map((run) => fieldHtml(run)).join('')}</span><span class="note-tab-segment note-tab-decimal-after" style="position:absolute;${position};white-space:pre">${after.map((run) => fieldHtml(run)).join('')}</span>`
        }
        return `<span class="note-tab-segment" data-align="${escAttr(stop.align ?? (paragraph.rtl ? 'r' : 'l'))}" style="position:absolute;${position}${shift(stop.align) ? `;transform:${shift(stop.align)}` : ''};white-space:pre">${index ? '' : marker}${runs.map((run) => fieldHtml(run)).join('')}</span>`
      })
      return `<span class="note-tab-line" style="position:relative;display:block;min-height:${paragraph.lineExactPt ?? (basePt * (paragraph.lineHeight ?? 120)) / 100}pt">${placed.join('')}</span>`
    }
    return paragraphs
      .map((paragraph) => {
        const basePt = paragraph.runs.find((run) => run.fontSizePt)?.fontSizePt ?? 12
        const bullet = paragraph.bullet
        const before =
          paragraph.spaceBeforePt ??
          (paragraph.spaceBeforePct != null ? (basePt * paragraph.spaceBeforePct) / 100 : undefined)
        const after =
          paragraph.spaceAfterPt ??
          (paragraph.spaceAfterPct != null ? (basePt * paragraph.spaceAfterPct) / 100 : undefined)
        const tabPt = paragraph.tabStopsPt?.[0] ?? paragraph.defaultTabPt
        const style = [
          paragraph.align ? `text-align:${paragraph.align}` : '',
          paragraph.lineExactPt != null
            ? `line-height:${paragraph.lineExactPt}pt`
            : paragraph.lineHeight
              ? `line-height:${paragraph.lineHeight}%`
              : '',
          before != null ? `margin-top:${before}pt` : '',
          after != null ? `margin-bottom:${after}pt` : '',
          paragraph.marginLeftPt != null
            ? `margin-left:${paragraph.marginLeftPt}pt`
            : paragraph.level
              ? `margin-left:${paragraph.level * 18}pt`
              : '',
          paragraph.marginRightPt != null ? `margin-right:${paragraph.marginRightPt}pt` : '',
          paragraph.indentPt != null ? `text-indent:${paragraph.indentPt}pt` : '',
          tabPt != null ? `tab-size:${Math.max(1, Math.round(tabPt / (basePt * 0.5)))}` : '',
        ]
          .filter(Boolean)
          .join(';')
        const bulletStyle = [
          bullet?.fontFamily ? `font-family:${JSON.stringify(bullet.fontFamily)}` : '',
          bullet?.color && /^#[0-9A-Fa-f]{6}$/.test(bullet.color) ? `color:${bullet.color}` : '',
          bullet?.sizePt
            ? `font-size:${bullet.sizePt}pt`
            : bullet?.sizePct
              ? `font-size:${bullet.sizePct}%`
              : '',
        ]
          .filter(Boolean)
          .join(';')
        const marker = bullet
          ? `<span class="note-bullet"${bulletStyle ? ` style="${escAttr(bulletStyle)}"` : ''}>${bullet.imageDataUrl ? `<img src="${escAttr(bullet.imageDataUrl)}" alt="">` : esc(bulletText(paragraph))}</span>`
          : ''
        return `<p${paragraph.rtl ? ' dir="rtl"' : ''}${style ? ` style="${escAttr(style)}"` : ''}${paragraph.tabStopsPt?.length ? ` data-tab-stops="${escAttr(paragraph.tabStopsPt.join(','))}"` : ''}>${tabbedRuns(paragraph, marker, basePt) || '&nbsp;'}</p>`
      })
      .join('')
  }
  const pageFields = (fields: PrintHeaderFooterPage | undefined, number?: number): string => {
    if (!fields) return ''
    const defaults: Record<
      keyof Pick<PrintHeaderFooterPage, 'header' | 'date' | 'footer' | 'pageNum'>,
      [number, number, number, number]
    > = {
      header: [5, 2, 90, 4],
      date: [5, 93, 30, 4],
      footer: [35, 93, 40, 4],
      pageNum: [78, 93, 17, 4],
    }
    return (['header', 'date', 'footer', 'pageNum'] as const)
      .map((kind) => {
        const field = fields[kind]
        const value = fieldHtml(field, number)
        if (!value) return ''
        const box = field?.box
        const size = fields.canvas
        const [x, y, w, h] =
          box && size
            ? [
                (box.x / size.cx) * 100,
                (box.y / size.cy) * 100,
                (box.cx / size.cx) * 100,
                (box.cy / size.cy) * 100,
              ]
            : defaults[kind]
        const style = `left:${x}%;top:${y}%;width:${w}%;height:${h}%;text-align:${field?.align ?? (kind === 'footer' ? 'center' : kind === 'pageNum' ? 'right' : 'left')}${boxTransform(box)}`
        return `<div class="hf-field hf-${kind}" style="${escAttr(style)}">${value}</div>`
      })
      .join('')
  }

  const boxStyle = (
    box: PrintPageBox | undefined,
    fields: PrintHeaderFooterPage | undefined,
  ): string => {
    if (!box || !fields?.canvas) return ''
    const { cx, cy } = fields.canvas
    return `position:absolute;left:${(box.x / cx) * 100}%;top:${(box.y / cy) * 100}%;width:${(box.cx / cx) * 100}%;height:${(box.cy / cy) * 100}%;margin:0;box-sizing:border-box;object-fit:contain${boxTransform(box)}`
  }
  const handoutContentStyle = (fields: PrintHeaderFooterPage | undefined): string => {
    const canvas = fields?.canvas
    const fieldBoxes = [fields?.header, fields?.date, fields?.footer, fields?.pageNum]
      .map((f) => f?.box)
      .filter((b): b is NonNullable<typeof b> => !!b)
    if (!canvas || !fieldBoxes.length) return ''
    const top = Math.max(
      5,
      ...fieldBoxes
        .filter((b) => b.y < canvas.cy / 2)
        .map((b) => ((b.y + b.cy) / canvas.cy) * 100 + 1),
    )
    const bottom = Math.min(
      93,
      ...fieldBoxes.filter((b) => b.y >= canvas.cy / 2).map((b) => (b.y / canvas.cy) * 100 - 1),
    )
    return `position:absolute;left:5%;right:5%;top:${top}%;bottom:${Math.max(0, 100 - bottom)}%;width:auto;height:auto`
  }

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
      .map((src, i) => {
        const fields = o.notesFields?.[i]
        const slideImage =
          fields?.slideImageVisible === false
            ? ''
            : `<img class="notes-slide-image" src="${escAttr(src)}"${fields?.slideImage ? ` style="${escAttr(boxStyle(fields.slideImage, fields))}"` : ''}>`
        const notesBody =
          fields?.notesBodyVisible === false
            ? ''
            : `<div class="note"${fields?.notesBody ? ` style="${escAttr(boxStyle(fields.notesBody, fields))}"` : ''}>${notesHtml(fields, o.notes?.[i] ?? '')}</div>`
        return `<div class="page notes"${fields?.background ? ` style="background:${escAttr(fields.background)}"` : ''}>${masterArtwork(fields)}${slideImage}${notesBody}${pageFields(fields, (o.firstSlideNum ?? 1) + (o.slideIndexes?.[i] ?? i))}</div>`
      })
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
        ? `grid-template-columns:repeat(${hCols},1fr);grid-template-rows:repeat(${hRows},1fr);`
        : ''
      const contentStyle = `${grid}${handoutContentStyle(o.handoutFields)}`
      pages.push(
        `<div class="page handout h${perPage}"${o.handoutFields?.background ? ` style="background:${escAttr(o.handoutFields.background)}"` : ''}>${masterArtwork(o.handoutFields)}<div class="handout-content h${perPage}${gridHandout ? ' hg' : ''}"${contentStyle ? ` style="${escAttr(contentStyle)}"` : ''}>${cells}</div>${pageFields(o.handoutFields, Math.floor(i / perPage) + 1)}</div>`,
      )
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
.page.handout { position: relative; padding: 0.42in 0.4in; }
.page.handout .handout-content { width: 100%; height: 100%; display: flex; flex-direction: ${landscape && perPage === 2 ? 'row' : 'column'}; gap: 0.24in; min-height: 0; }
.page.handout .cell { display: flex; gap: 0.2in; align-items: center; flex: 1; min-width: 0; min-height: 0; }
.page.handout .cell img { border: 1px solid #bbb; object-fit: contain; max-height: 100%; }
.page.handout.h2 .cell img, .page.handout.h6 .cell img { width: 100%; height: auto; max-height: 100%; }
.page.handout.h3 .cell img { width: 55%; height: auto; }
.page.handout.h3 .rules {
  flex: 1; align-self: stretch;
  background: repeating-linear-gradient(#fff 0 0.28in, #ccc 0.28in calc(0.28in + 1px));
}
.page.handout .handout-content.h6 { display: grid; grid-template-columns: ${h6Cols}; grid-auto-rows: 1fr; }
.page.handout .handout-content.hg { display: grid; gap: 0.24in; }
.page.handout .handout-content.hg .cell { justify-content: center; }
.page.handout .handout-content.hg .cell img { width: 100%; height: auto; max-height: 100%; }
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
.page.notes { position: relative; padding: 0.5in; display: flex; flex-direction: column; }
.page.notes .notes-slide-image { position:relative; z-index:1; width: ${landscape ? 'auto' : '100%'}; ${landscape ? 'max-height: 55%; align-self: center;' : 'height: auto;'} border: 1px solid #bbb; }
.page.notes .note { position:relative; z-index:2; margin-top: 0.3in; font-size: 11pt; line-height: 1.5; white-space: pre-wrap; }
.page.notes .note p { margin:0; }
.page.notes .note .note-bullet { display:inline-block; margin-inline-end:0.4em; }
.page.notes .note .note-bullet img { display:inline; width:1em; height:1em; object-fit:contain; border:0; }
.page .master-artwork { position:absolute; inset:0; width:100%; height:100%; z-index:0; border:0; pointer-events:none; }
.page.handout .handout-content { z-index:1; }
.hf-field { position: absolute; z-index: 3; font-size: 10pt; overflow: visible; white-space: nowrap; }
${nupCss}
${frameCss}
${transformCss}
${outputCss}
${previewCss}
</style></head><body>${body}</body></html>`
}
