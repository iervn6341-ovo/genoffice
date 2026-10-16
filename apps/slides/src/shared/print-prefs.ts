/**
 * Print preferences for the macOS native-panel route. On macOS "Print" goes straight to the system
 * print panel, so there is no in-app sheet to pick the layout or the PowerPoint options; they live
 * in the File menu (main process owns and persists them) and are baked into the PDF the panel prints.
 * Pure helpers: no Electron or DOM imports.
 */
import {
  PRINT_PAPERS,
  handoutSpec,
  type PrintLayout,
  type PrintOrientation,
  type PrintOutput,
  type PrintPaper,
} from './print-html'

export interface SlidesPrintPrefs {
  /** Layout the next plain "Print…" uses (the last one printed) */
  layout: PrintLayout
  /** Print hidden slides too */
  includeHidden: boolean
  /** Frame slides */
  frame: boolean
  /** Color / grayscale / pure black & white */
  output: PrintOutput
  /** Scale the slide up to the paper (full layout on real paper) */
  scaleFit: boolean
  /** Paper: 'slide' = the PDF page is slide-sized (full layout only) */
  paper: PrintPaper
  /** Print comment pages after the slides */
  comments: boolean
}

export const DEFAULT_PRINT_PREFS: SlidesPrintPrefs = {
  layout: 'full',
  includeHidden: false,
  frame: false,
  output: 'color',
  scaleFit: false,
  paper: 'slide',
  comments: false,
}

/** Layouts reachable from the menu, in menu order */
export const NATIVE_HANDOUT_LAYOUTS = [
  'handout1',
  'handout2',
  'handout3',
  'handout4h',
  'handout4v',
  'handout6h',
  'handout6v',
  'handout9h',
  'handout9v',
] as const satisfies readonly PrintLayout[]

const OUTPUTS: readonly PrintOutput[] = ['color', 'grayscale', 'bw']

export function isPrintLayout(value: unknown): value is PrintLayout {
  return (
    typeof value === 'string' &&
    (value === 'full' ||
      value === 'notes' ||
      value === 'outline' ||
      handoutSpec(value as PrintLayout) !== null)
  )
}

/** Anything read from disk or IPC → a complete, valid preference object */
export function normalizePrintPrefs(raw: unknown): SlidesPrintPrefs {
  const r = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>
  const flag = (key: keyof SlidesPrintPrefs, fallback: boolean): boolean =>
    typeof r[key] === 'boolean' ? (r[key] as boolean) : fallback
  const d = DEFAULT_PRINT_PREFS
  return {
    layout: isPrintLayout(r.layout) ? r.layout : d.layout,
    includeHidden: flag('includeHidden', d.includeHidden),
    frame: flag('frame', d.frame),
    output: OUTPUTS.includes(r.output as PrintOutput) ? (r.output as PrintOutput) : d.output,
    scaleFit: flag('scaleFit', d.scaleFit),
    paper:
      typeof r.paper === 'string' && Object.prototype.hasOwnProperty.call(PRINT_PAPERS, r.paper)
        ? (r.paper as PrintPaper)
        : d.paper,
    comments: flag('comments', d.comments),
  }
}

/** Slide indices that print: hidden slides drop out unless "print hidden slides" is on */
export function selectPrintSlides(
  slides: ReadonlyArray<{ hidden?: boolean }>,
  includeHidden: boolean,
): number[] {
  const out: number[] = []
  slides.forEach((s, i) => {
    if (includeHidden || !s.hidden) out.push(i)
  })
  return out
}

/** What the PDF assembly receives for a layout + preferences (mirrors the in-app dialog's rules) */
export interface NativePrintParams {
  layout: PrintLayout
  /** 'slide' only for the plain full layout; every other layout prints on real paper */
  paper: PrintPaper
  frame: boolean
  output: PrintOutput
  scaleFit: boolean
  /** handout / notes / outline pages only; a slide-sized page follows the slide ratio */
  orientation: PrintOrientation | null
}

export function nativePrintParams(
  layout: PrintLayout,
  prefs: SlidesPrintPrefs,
  /** orientation of the deck's own notes / handout master */
  masterOrientation: PrintOrientation = 'portrait',
): NativePrintParams {
  const slidePage = layout === 'full'
  const paper: PrintPaper = prefs.paper === 'slide' && !slidePage ? 'a4' : prefs.paper
  const onRealPaper = paper !== 'slide'
  return {
    layout,
    paper,
    // Outline and notes pages never draw the frame (the thumbnails are framed already)
    frame: prefs.frame && layout !== 'outline' && layout !== 'notes',
    output: prefs.output,
    scaleFit: prefs.scaleFit && slidePage && onRealPaper,
    // a real-paper full page keeps portrait; the master's orientation drives the rest
    orientation: onRealPaper || layout !== 'full' ? masterOrientation : null,
  }
}

/**
 * macOS: "Print" goes straight to the standard system print panel with its page preview (Chromium's
 * own panel, what webContents.print opens, has none). Elsewhere the in-app sheet applies.
 */
export function usesMacPrintPanel(platform: string, op: { nativePanel?: boolean }): boolean {
  return platform === 'darwin' && op.nativePanel === true
}
