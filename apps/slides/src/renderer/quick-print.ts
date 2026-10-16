/**
 * macOS: "Print" skips the in-app sheet and opens the system print panel (with its preview). The
 * layout and options come from the File menu preferences, so this builds the same job the in-app
 * sheet would have built for those settings and the main process turns it into the PDF the panel prints.
 */
import type { RenderSlide } from '@genoffice/pptx-render'
import { backstageLabels, printOfficeLabels } from '@genoffice/ui'
import type { Lang } from '@genoffice/i18n'
import type { PrintSlidesOp } from '../shared/ipc'
import type { PrintCommentItem, PrintLayout, PrintOutlineItem } from '../shared/print-html'
import {
  NATIVE_HANDOUT_LAYOUTS,
  nativePrintParams,
  selectPrintSlides,
  type SlidesPrintPrefs,
} from '../shared/print-prefs'
import { loadSlidesPrintData, type SlidesPrintData } from './print-data'
import { outlineOf } from './outline-of'

export interface NativePrintInput {
  slides: RenderSlide[]
  data: SlidesPrintData
  prefs: SlidesPrintPrefs
  layout: PrintLayout
  dateLocale: string
  commentsTitle: string
}

/** The print job for a layout + preferences; null when there is nothing to print */
export function buildNativePrintOp(input: NativePrintInput): PrintSlidesOp | null {
  const { slides, data, prefs, layout } = input
  const selected = selectPrintSlides(slides, prefs.includeHidden)
  const first = selected[0]
  if (first === undefined) return null
  const masterCanvas = data.nativeFields.notesPages[0]?.canvas ?? data.nativeFields.handout.canvas
  const params = nativePrintParams(
    layout,
    prefs,
    masterCanvas && masterCanvas.cx > masterCanvas.cy ? 'landscape' : 'portrait',
  )
  const outline: PrintOutlineItem[] =
    layout === 'outline' ? selected.map((i) => ({ number: i + 1, ...outlineOf(slides[i]!) })) : []
  const picked = new Set(selected)
  const comments: PrintCommentItem[] = prefs.comments
    ? data.comments.filter((c) => picked.has(c.number - 1))
    : []
  return {
    pngsBase64: selected.map((i) => data.rendered[i]!),
    widthPx: slides[first]!.widthPx,
    heightPx: slides[first]!.heightPx,
    layout,
    slideIndexes: selected,
    dateLocale: input.dateLocale,
    masterArtworkBase64: data.masterArtworkBase64,
    ...(layout === 'notes' ? { notes: selected.map((i) => data.notes[i] ?? '') } : {}),
    ...(layout === 'outline' ? { outline } : {}),
    ...(comments.length > 0 ? { comments, commentsTitle: input.commentsTitle } : {}),
    ...(params.orientation ? { orientation: params.orientation } : {}),
    ...(params.frame ? { frame: true } : {}),
    ...(params.output !== 'color' ? { output: params.output } : {}),
    paper: params.paper,
    ...(params.scaleFit ? { scaleFit: true } : {}),
    nativePanel: true,
  }
}

export type NativePrintOutcome = 'printed' | 'canceled' | 'empty' | { error: string }

/** Render the deck, build the job and open the macOS print panel. Cancel is not an error. */
export async function runNativePrint(args: {
  slides: RenderSlide[]
  images: Map<string, HTMLImageElement>
  current: number
  /** layout to print; omitted = the last-used one */
  layout?: PrintLayout
  dateLocale: string
  commentsTitle: string
}): Promise<NativePrintOutcome> {
  try {
    const stored = await window.slidesApi.getPrintPrefs()
    const layout = args.layout ?? stored.layout
    const prefs = args.layout ? await window.slidesApi.setPrintPrefs({ layout }) : stored
    const data = await loadSlidesPrintData(args.slides, args.images, args.current)
    const op = buildNativePrintOp({
      slides: args.slides,
      data,
      prefs,
      layout,
      dateLocale: args.dateLocale,
      commentsTitle: args.commentsTitle,
    })
    if (!op) return 'empty'
    const r = await window.slidesApi.printSlides(op)
    if (r.ok) return 'printed'
    return r.error ? { error: r.error } : 'canceled'
  } catch (error) {
    return { error: String(error) }
  }
}

/** One Backstage entry per print layout (macOS has no in-app sheet to choose from) */
export function nativePrintEntries(lang: Lang): Array<{ layout: PrintLayout; label: string }> {
  const L = printOfficeLabels(lang)
  const print = backstageLabels(lang).navPrint
  const handout: Record<(typeof NATIVE_HANDOUT_LAYOUTS)[number], string> = {
    handout1: L.h1,
    handout2: L.h2,
    handout3: L.h3,
    handout4h: L.h4h,
    handout4v: L.h4v,
    handout6h: L.h6h,
    handout6v: L.h6v,
    handout9h: L.h9h,
    handout9v: L.h9v,
  }
  return [
    { layout: 'full', label: `${print} · ${L.plFull}` },
    { layout: 'notes', label: `${print} · ${L.plNotes}` },
    { layout: 'outline', label: `${print} · ${L.plOutline}` },
    ...NATIVE_HANDOUT_LAYOUTS.map((layout) => ({
      layout,
      label: `${print} · ${L.plHandouts} · ${handout[layout]}`,
    })),
  ]
}
