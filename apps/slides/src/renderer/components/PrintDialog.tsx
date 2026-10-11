/**
 * Print dialog modeled on PowerPoint's print sheet: live page preview on the left
 * (the exact pages that will print, built by the shared print-html assembly),
 * range / layout / orientation / framing options on the right. "Print" hands the
 * assembled pages to the system print dialog.
 */
import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { RenderSlide } from '@genoffice/pptx-render'
import type { SectionInfo } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'
import { outlineOf } from '../outline-of'
import {
  buildPrintDocumentHtml,
  parsePrintRange,
  printPageCount,
  handoutSpec,
  type PrintCommentItem,
  type PrintLayout,
  type PrintOrientation,
  type PrintOutput,
  type PrintPaper,
  type PrintPerSheet,
  type PrintDirection,
  PRINT_PER_SHEET,
  PRINT_PAPERS,
} from '../../shared/print-html'
import { renderSlidesToPngBase64 } from '../export-render'
import {
  DEFAULT_PRINTER_SETTINGS,
  PrintCheck,
  PrinterSettingsFields,
  PrintRow,
  PrintSection,
  printerJob,
  printOfficeLabels,
  type PrinterSettings,
} from '@genoffice/ui'

/** Arrow glyphs of the four fill orders (2×2 cells in a 20px box) */
const DIR_PATHS: Record<PrintDirection, string> = {
  z: 'M3 4h8M11 4l-8 6h14M3 16h14M15 14l2 2-2 2',
  rz: 'M17 4H9M9 4l8 6H3M17 16H3M5 14l-2 2 2 2',
  n: 'M4 3v8M4 11l6-8v14M16 3v14M14 15l2 2 2-2',
  rn: 'M4 17V9M4 9l6 8V3M16 17V3M14 5l2-2 2 2',
}

type RangeMode = 'all' | 'selection' | 'current' | 'custom' | 'section'

export function PrintDialog({
  slides,
  images,
  current,
  selection,
  sections,
  onEditHeaderFooter,
  onClose,
  setStatus,
}: {
  /** The whole deck, hidden pages included (the hidden toggle filters at print time) */
  slides: RenderSlide[]
  images: Map<string, HTMLImageElement>
  /** Current page index, for the "current slide" range */
  current: number
  /** Slides highlighted in the thumbnail list (the "Print Selected Slides" range) */
  selection: number[]
  /** Deck sections ("Print Section" range) */
  sections: SectionInfo[]
  /** "Edit Header & Footer": closes this sheet and opens the header/footer dialog */
  onEditHeaderFooter: () => void
  onClose: () => void
  setStatus: (s: string) => void
}) {
  const { t, lang } = useI18n()
  const L = printOfficeLabels(lang)
  const [pngs, setPngs] = useState<string[] | null>(null)
  const [notes, setNotes] = useState<string[]>([])
  const [layout, setLayout] = useState<PrintLayout>('full')
  const [rangeMode, setRangeMode] = useState<RangeMode>('all')
  const [customRange, setCustomRange] = useState('')
  const [orientation, setOrientation] = useState<PrintOrientation>('portrait')
  const [frame, setFrame] = useState(false)
  const [includeHidden, setIncludeHidden] = useState(false)
  const [output, setOutput] = useState<PrintOutput>('color')
  const [paper, setPaper] = useState<PrintPaper>('slide')
  const [scaleFit, setScaleFit] = useState(false)
  const [highQuality, setHighQuality] = useState(false)
  const [withComments, setWithComments] = useState(false)
  const [comments, setComments] = useState<PrintCommentItem[]>([])
  const [sectionId, setSectionId] = useState(sections[0]?.id ?? '')
  const [perSheet, setPerSheet] = useState<PrintPerSheet>(1)
  const [direction, setDirection] = useState<PrintDirection>('z')
  const [rotate180, setRotate180] = useState(false)
  const [flipH, setFlipH] = useState(false)
  const [printerSettings, setPrinterSettings] = useState<PrinterSettings>(DEFAULT_PRINTER_SETTINGS)
  const [printing, setPrinting] = useState(false)
  const blobUrlsRef = useRef<string[]>([])
  const frameRef = useRef<HTMLIFrameElement | null>(null)
  const paneRef = useRef<HTMLDivElement | null>(null)
  const dialogRef = useRef<HTMLDivElement>(null)
  const previouslyFocused = useRef<HTMLElement | null>(null)

  // Focus the first field on mount; return focus to the opener on unmount
  useEffect(() => {
    previouslyFocused.current = document.activeElement as HTMLElement | null
    const root = dialogRef.current
    if (root && !root.contains(document.activeElement)) {
      root.querySelector<HTMLElement>('input, textarea, select, button')?.focus()
    }
    return () => {
      previouslyFocused.current?.focus?.()
    }
  }, [])

  // One offscreen render of the whole deck on open (print quality, 2x); the same
  // bitmaps feed the preview (as blob URLs) and the print job (as base64).
  useEffect(() => {
    let alive = true
    void (async () => {
      try {
        const rendered = await renderSlidesToPngBase64(slides, images)
        const noteTexts = await Promise.all(slides.map((_s, i) => window.slidesApi.getNotes(i)))
        const commentLists = await Promise.all(
          slides.map((_s, i) => window.slidesApi.getComments(i)),
        )
        if (!alive) return
        setComments(
          commentLists.flatMap((list, i) =>
            list.map((c) => ({ number: i + 1, author: c.author, text: c.text })),
          ),
        )
        blobUrlsRef.current = rendered.map((b64) =>
          URL.createObjectURL(
            new Blob([Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))], { type: 'image/png' }),
          ),
        )
        setNotes(noteTexts)
        setPngs(rendered)
      } catch (err) {
        if (alive) setStatus(t('appPrintFailed', { error: String(err) }))
      }
    })()
    return () => {
      alive = false
      for (const url of blobUrlsRef.current) URL.revokeObjectURL(url)
      blobUrlsRef.current = []
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        onClose()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose])

  /** 0-based deck indices selected by the range options (before the hidden filter) */
  const rangeIndices = useMemo<number[] | null>(() => {
    if (rangeMode === 'all') return slides.map((_s, i) => i)
    if (rangeMode === 'current') return [current]
    if (rangeMode === 'selection')
      return [...new Set(selection.length > 0 ? selection : [current])].sort((a, b) => a - b)
    if (rangeMode === 'section')
      return [...(sections.find((x) => x.id === sectionId)?.slideIndices ?? [])].sort(
        (a, b) => a - b,
      )
    return parsePrintRange(customRange, slides.length)
  }, [rangeMode, customRange, slides, current, selection, sections, sectionId])

  const selected = useMemo(
    () => (rangeIndices ?? []).filter((i) => includeHidden || !slides[i].hidden),
    [rangeIndices, includeHidden, slides],
  )
  const rangeInvalid = rangeMode === 'custom' && rangeIndices === null
  const pageCount = printPageCount(selected.length, layout, perSheet)
  /** "Slide size" paper only exists for one full-page slide per sheet; every other layout prints on paper */
  const slidePage = layout === 'full' && perSheet === 1
  const effPaper: PrintPaper = paper === 'slide' && !slidePage ? 'a4' : paper
  const onRealPaper = effPaper !== 'slide'

  const outline = useMemo(
    () =>
      layout === 'outline' ? selected.map((i) => ({ number: i + 1, ...outlineOf(slides[i]) })) : [],
    [layout, selected, slides],
  )
  const pageComments = useMemo(() => {
    if (!withComments) return []
    const set = new Set(selected)
    return comments.filter((c) => set.has(c.number - 1))
  }, [withComments, comments, selected])

  /** Everything the page assembly needs except the slide bitmaps (shared by preview and print) */
  const docOptions = () => ({
    layout,
    orientation,
    frame,
    output,
    paper: effPaper,
    perSheet,
    direction,
    rotate180,
    flipH,
    scaleFit,
    ...(layout === 'outline' ? { outline } : {}),
    ...(pageComments.length > 0 ? { comments: pageComments, commentsTitle: L.comments } : {}),
  })

  const previewHtml = useMemo(() => {
    if (!pngs || selected.length === 0) return null
    const first = slides[selected[0]]
    return buildPrintDocumentHtml({
      srcs: selected.map((i) => blobUrlsRef.current[i]),
      ratio: first.widthPx / first.heightPx,
      notes: selected.map((i) => notes[i] ?? ''),
      ...docOptions(),
      preview: true,
    })
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [
    pngs,
    selected,
    slides,
    layout,
    notes,
    orientation,
    frame,
    output,
    effPaper,
    perSheet,
    direction,
    rotate180,
    flipH,
    scaleFit,
    outline,
    pageComments,
  ])

  /** Scale the preview document so one page fits the pane width (iframe body zoom) */
  const applyZoom = useCallback(() => {
    const doc = frameRef.current?.contentDocument
    const pane = paneRef.current
    const page = doc?.querySelector('.page') as HTMLElement | null
    if (!doc || !pane || !page) return
    doc.body.style.zoom = ''
    const pageW = page.getBoundingClientRect().width
    if (pageW > 0) doc.body.style.zoom = String(Math.min((pane.clientWidth - 44) / pageW, 1))
  }, [])

  useEffect(() => {
    const pane = paneRef.current
    if (!pane) return
    const ro = new ResizeObserver(applyZoom)
    ro.observe(pane)
    return () => ro.disconnect()
  }, [applyZoom])

  const doPrint = async () => {
    if (!pngs || selected.length === 0 || printing) return
    setPrinting(true)
    try {
      const first = slides[selected[0]]
      // "High Quality": re-render the printed slides at a higher pixel ratio
      const bitmaps = highQuality
        ? await renderSlidesToPngBase64(
            selected.map((i) => slides[i]),
            images,
            3,
          )
        : selected.map((i) => pngs[i])
      const doc = docOptions()
      const r = await window.slidesApi.printSlides({
        pngsBase64: bitmaps,
        widthPx: first.widthPx,
        heightPx: first.heightPx,
        layout,
        ...(layout === 'notes' ? { notes: selected.map((i) => notes[i] ?? '') } : {}),
        ...(layout === 'outline' ? { outline } : {}),
        ...(pageComments.length > 0 ? { comments: pageComments, commentsTitle: L.comments } : {}),
        ...(onRealPaper || layout !== 'full' ? { orientation } : {}),
        ...(frame ? { frame: true } : {}),
        ...(layout === 'full' && perSheet > 1 ? { perSheet, direction } : {}),
        ...(rotate180 ? { rotate180: true } : {}),
        ...(flipH ? { flipH: true } : {}),
        ...(output !== 'color' ? { output } : {}),
        paper: doc.paper,
        ...(scaleFit ? { scaleFit: true } : {}),
        ...printerJob(printerSettings),
      })
      if (r.ok) onClose()
      else if (r.error) setStatus(t('appPrintFailed', { error: r.error }))
      // ok=false without an error = canceled in the system dialog: keep the dialog open
    } finally {
      setPrinting(false)
    }
  }

  /** One value for the "Print Layout" select */
  const handoutOptions: Array<[PrintLayout, string]> = [
    ['handout1', L.h1],
    ['handout2', L.h2],
    ['handout3', L.h3],
    ['handout4h', L.h4h],
    ['handout6h', L.h6h],
    ['handout9h', L.h9h],
    ['handout4v', L.h4v],
    ['handout6v', L.h6v],
    ['handout9v', L.h9v],
  ]
  const layoutValue = layout === 'handout6' ? 'handout6h' : layout
  const hasSections = sections.length > 0

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div
        ref={dialogRef}
        className="modal print-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('appPrintTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h2>{t('appPrintTitle')}</h2>
        <div className="print-dialog-body">
          <div className="print-preview-pane" ref={paneRef}>
            {previewHtml ? (
              <iframe
                ref={frameRef}
                title="print-preview"
                sandbox="allow-same-origin"
                srcDoc={previewHtml}
                onLoad={applyZoom}
              />
            ) : (
              <div className="print-preview-empty">
                {pngs ? t('appPrintNoPages') : t('appPrintRendering')}
              </div>
            )}
          </div>
          <div className="print-options">
            <PrintSection title={L.secPrinter}>
              <PrinterSettingsFields
                lang={lang}
                value={printerSettings}
                onChange={setPrinterSettings}
                listPrinters={() => window.slidesApi.listPrinters()}
              />
            </PrintSection>
            <PrintSection title={L.secSettings}>
              <select
                className="ps-select"
                aria-label={L.secSettings}
                data-testid="print-range"
                value={rangeMode}
                onChange={(e) => setRangeMode(e.target.value as RangeMode)}
              >
                <option value="all">{L.rangeAllSlides}</option>
                <option value="selection">{L.rangeSelSlides}</option>
                <option value="current">{L.rangeCurSlide}</option>
                <option value="custom">{L.rangeCustom}</option>
                {hasSections && <option value="section">{L.rangeSection}</option>}
              </select>
              {rangeMode === 'custom' && (
                <PrintRow label={L.slides}>
                  <input
                    type="text"
                    className={`print-range-input ${rangeInvalid ? 'invalid' : ''}`}
                    placeholder={L.pagesHint}
                    value={customRange}
                    onChange={(e) => setCustomRange(e.target.value)}
                  />
                </PrintRow>
              )}
              {rangeMode === 'section' && (
                <select
                  className="ps-select"
                  aria-label={L.rangeSection}
                  value={sectionId}
                  onChange={(e) => setSectionId(e.target.value)}
                >
                  {sections.map((x) => (
                    <option key={x.id} value={x.id}>
                      {x.name}
                    </option>
                  ))}
                </select>
              )}
              <PrintRow label={L.plLayout}>
                <select
                  className="ps-select"
                  data-testid="print-layout"
                  value={layoutValue}
                  onChange={(e) => setLayout(e.target.value as PrintLayout)}
                >
                  <option value="full">{L.plFull}</option>
                  <option value="notes">{L.plNotes}</option>
                  <option value="outline">{L.plOutline}</option>
                  <optgroup label={L.plHandouts}>
                    {handoutOptions.map(([k, label]) => (
                      <option key={k} value={k}>
                        {label}
                      </option>
                    ))}
                  </optgroup>
                </select>
              </PrintRow>
              <PrintRow label={L.colorLabel}>
                <select
                  className="ps-select"
                  data-testid="print-color"
                  value={output}
                  onChange={(e) => setOutput(e.target.value as PrintOutput)}
                >
                  <option value="color">{L.colColor}</option>
                  <option value="grayscale">{L.colGray}</option>
                  <option value="bw">{L.colBW}</option>
                </select>
              </PrintRow>
            </PrintSection>
            <PrintSection title={L.secPageSetup} defaultOpen={false}>
              <PrintRow label={L.orientation}>
                <select
                  className="ps-select"
                  data-testid="print-orientation"
                  value={orientation}
                  disabled={!onRealPaper && layout === 'full'}
                  onChange={(e) => setOrientation(e.target.value as PrintOrientation)}
                >
                  <option value="portrait">{L.portrait}</option>
                  <option value="landscape">{L.landscape}</option>
                </select>
              </PrintRow>
              <PrintRow label={L.paperSize}>
                <select
                  className="ps-select"
                  data-testid="print-paper"
                  value={effPaper}
                  onChange={(e) => setPaper(e.target.value as PrintPaper)}
                >
                  {slidePage && <option value="slide">{L.slideSizePaper}</option>}
                  {(Object.keys(PRINT_PAPERS) as PrintPaper[])
                    .filter((k) => k !== 'slide')
                    .map((k) => (
                      <option key={k} value={k}>
                        {PRINT_PAPERS[k].label}
                      </option>
                    ))}
                </select>
              </PrintRow>
              <PrintRow label={L.pagesPerSheet}>
                <select
                  className="ps-select"
                  data-testid="print-per-sheet"
                  value={perSheet}
                  disabled={layout !== 'full'}
                  onChange={(e) => setPerSheet(Number(e.target.value) as PrintPerSheet)}
                >
                  {PRINT_PER_SHEET.map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
              </PrintRow>
              <div className="ps-field">
                <span>{t('appPrintDirection')}</span>
                <span className="print-dirs" role="group" aria-label={t('appPrintDirection')}>
                  {(['z', 'rz', 'n', 'rn'] as const).map((d) => (
                    <button
                      key={d}
                      type="button"
                      className={`print-dir ${direction === d ? 'active' : ''}`}
                      aria-pressed={direction === d}
                      disabled={layout !== 'full' || perSheet === 1}
                      title={d.toUpperCase()}
                      onClick={() => setDirection(d)}
                    >
                      <svg width="20" height="20" viewBox="0 0 20 20" aria-hidden="true">
                        <path d={DIR_PATHS[d]} />
                      </svg>
                    </button>
                  ))}
                </span>
              </div>
              <PrintCheck
                label={t('appPrintRotate180')}
                checked={rotate180}
                onChange={setRotate180}
              />
              <PrintCheck label={t('appPrintFlipH')} checked={flipH} onChange={setFlipH} />
            </PrintSection>
            <PrintSection title={L.secOptions} defaultOpen={false}>
              <PrintCheck
                label={L.frameSlides}
                checked={frame}
                disabled={layout === 'outline' || layout === 'notes'}
                onChange={setFrame}
              />
              <PrintCheck
                label={L.scaleFit}
                checked={scaleFit}
                disabled={!(layout === 'full' && perSheet === 1 && onRealPaper)}
                onChange={setScaleFit}
              />
              <PrintCheck label={L.highQuality} checked={highQuality} onChange={setHighQuality} />
              <PrintCheck
                label={L.printComments}
                checked={withComments}
                disabled={comments.length === 0}
                onChange={setWithComments}
              />
              <PrintCheck
                label={L.printHiddenSlides}
                checked={includeHidden}
                onChange={setIncludeHidden}
              />
              <button type="button" className="ps-link" onClick={onEditHeaderFooter}>
                {L.headerFooterBtn}
              </button>
            </PrintSection>
            <div className="print-page-count">
              {selected.length > 0 && layout !== 'outline'
                ? t('appPrintPageCount', { n: pageCount })
                : ''}
            </div>
          </div>
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('appSettingsCancel')}</button>
          <button
            className="primary"
            disabled={!pngs || selected.length === 0 || printing}
            onClick={() => void doPrint()}
          >
            {printing ? t('appPrintProgress') : t('appPrintTitle')}
          </button>
        </div>
      </div>
    </div>
  )
}
