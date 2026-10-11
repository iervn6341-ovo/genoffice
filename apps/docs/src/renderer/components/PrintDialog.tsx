/**
 * Print dialog modeled on Word's print pane: live page preview on the left (a scaled clone of
 * the pagination-preview sheet that will actually print), the settings on the right grouped the
 * way Word groups them — Printer / Settings / Page Setup / Options. "Print" hands the preview
 * pages to the system print dialog.
 *
 * The pagination preview must be mounted (the App opens it, visually hidden,
 * when the dialog opens): its .pv-page boxes are both the print source and the
 * preview source, so what the dialog shows is exactly what prints. Orientation, paper size and
 * margins edit the document's page setup (like Word's pane) and re-slice that preview.
 */
import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { SectionSettings } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { parsePrintRange } from '../print-range'
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
import { clearPrintZoom, setPrintZoom } from '../print-zoom'
import {
  PAGES_PER_SHEET,
  PX_PER_TWIP,
  buildSheets,
  keepParity,
  type PagesPerSheet,
} from '../print-sheets'
import { MARGIN_PRESETS, PAPER_SIZES } from './ribbon-layout-tab'

type RangeMode = 'all' | 'current' | 'custom'
type Parity = 'all' | 'odd' | 'even'

const TWIPS_PER_CM = 1440 / 2.54

const pvPages = (): HTMLElement[] => [
  ...document.querySelectorAll<HTMLElement>('.pagination-preview .pv-page'),
]
const pvRoot = (): HTMLElement | null => document.querySelector<HTMLElement>('.pagination-preview')

/** The paper entry matching a page size (either orientation), or null for a custom size */
function paperKeyOf(s: SectionSettings | null): string {
  if (!s) return ''
  const w = Math.min(s.pageWidth, s.pageHeight)
  const h = Math.max(s.pageWidth, s.pageHeight)
  return PAPER_SIZES.find((p) => Math.abs(p.w - w) < 12 && Math.abs(p.h - h) < 12)?.key ?? ''
}

function marginKeyOf(s: SectionSettings | null): string {
  if (!s) return 'normal'
  return (
    MARGIN_PRESETS.find(
      (m) =>
        m.top === s.marginTop &&
        m.right === s.marginRight &&
        m.bottom === s.marginBottom &&
        m.left === s.marginLeft,
    )?.key ?? 'custom'
  )
}

export function PrintDialog({
  onClose,
  setStatus,
  section,
  onPageSetup,
}: {
  onClose: () => void
  setStatus: (s: string) => void
  /** The document's page setup (first section) shown by Orientation / Paper size / Margins */
  section: SectionSettings | null
  /** Applies a page setup change to every section, like Word's print pane */
  onPageSetup: (change: (s: SectionSettings) => SectionSettings) => void
}) {
  const { t, lang } = useI18n()
  const L = printOfficeLabels(lang)
  const [pageCount, setPageCount] = useState(0)
  /** Changes whenever the pagination preview re-slices (page count or sheet size) */
  const [sig, setSig] = useState('')
  const [page, setPage] = useState(0)
  const [rangeMode, setRangeMode] = useState<RangeMode>('all')
  const [customRange, setCustomRange] = useState('')
  const [parity, setParity] = useState<Parity>('all')
  const [reverse, setReverse] = useState(false)
  const [perSheet, setPerSheet] = useState<PagesPerSheet>(1)
  const [scalePaper, setScalePaper] = useState('none')
  const [pageColor, setPageColor] = useState(false)
  const [printerSettings, setPrinterSettings] = useState<PrinterSettings>(DEFAULT_PRINTER_SETTINGS)
  const [printing, setPrinting] = useState(false)
  const paneRef = useRef<HTMLDivElement | null>(null)
  const hostRef = useRef<HTMLDivElement | null>(null)

  // The pagination preview measures and slices asynchronously; poll until the
  // .pv-page count is stable (and keep tracking it, e.g. re-slices on resize).
  useEffect(() => {
    let last = -1
    const timer = window.setInterval(() => {
      const pages = pvPages()
      const n = pages.length
      if (n > 0 && n === last) {
        setPageCount(n)
        setSig(`${n}:${pages[0]!.offsetWidth}x${pages[0]!.offsetHeight}`)
      }
      last = n
    }, 250)
    return () => window.clearInterval(timer)
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

  const ready = pageCount > 0
  const pageNow = Math.min(page, Math.max(0, pageCount - 1))

  /** 0-based preview pages in print order: range → odd/even → reverse (null = invalid custom range) */
  const printOrder: number[] | null = useMemo(() => {
    const base =
      rangeMode === 'all'
        ? Array.from({ length: pageCount }, (_x, i) => i)
        : rangeMode === 'current'
          ? [pageNow]
          : parsePrintRange(customRange, pageCount)
    if (base === null) return null
    const kept = base.filter((i) => keepParity(i, parity))
    return reverse ? kept.reverse() : kept
  }, [rangeMode, customRange, pageCount, pageNow, parity, reverse])
  const rangeInvalid = rangeMode === 'custom' && printOrder === null
  const total = printOrder?.length ?? 0
  const sheetCount = Math.ceil(total / perSheet)
  const sheetNow = Math.min(page, Math.max(0, sheetCount - 1))

  /** Sheet size in CSS px for pages-per-sheet / scale-to-paper (undefined = the page's own size) */
  const sheetSizeFor = useCallback(
    (first: HTMLElement | undefined): { sheetW?: number; sheetH?: number } => {
      const paper = PAPER_SIZES.find((p) => p.key === scalePaper)
      if (!paper || !first) return {}
      const landscape = first.offsetWidth > first.offsetHeight
      return {
        sheetW: (landscape ? paper.h : paper.w) * PX_PER_TWIP,
        sheetH: (landscape ? paper.w : paper.h) * PX_PER_TWIP,
      }
    },
    [scalePaper],
  )

  /** Scale the preview sheet so the whole sheet fits the preview pane (Word-style single-page fit) */
  const applyZoom = useCallback(() => {
    const pane = paneRef.current
    const host = hostRef.current
    const sheet = host?.querySelector<HTMLElement>('.pv-sheet') ?? null
    if (!pane || !host || !sheet) return
    host.style.zoom = '1'
    const w = sheet.offsetWidth
    const h = sheet.offsetHeight
    if (w > 0 && h > 0) {
      host.style.zoom = String(Math.min((pane.clientWidth - 36) / w, (pane.clientHeight - 36) / h))
    }
  }, [])

  // Show one printed sheet at a time: clones of the pages that land on it (styles apply
  // in-document, so the clones render identically) fitted to the pane.
  useEffect(() => {
    const host = hostRef.current
    if (!host) return
    host.replaceChildren()
    if (!printOrder || printOrder.length === 0) return
    const pages = pvPages()
    const chosen = printOrder
      .slice(sheetNow * perSheet, (sheetNow + 1) * perSheet)
      .map((i) => pages[i])
      .filter((el): el is HTMLElement => el !== undefined)
    if (chosen.length === 0) return
    const { sheetW, sheetH } = sheetSizeFor(pages[printOrder[0]!])
    buildSheets(host, chosen, {
      perSheet,
      ...(sheetW !== undefined ? { sheetW } : {}),
      ...(sheetH !== undefined ? { sheetH } : {}),
    })
    host.querySelectorAll('.pv-pageno').forEach((el) => el.remove())
    applyZoom()
  }, [sig, printOrder, sheetNow, perSheet, sheetSizeFor, applyZoom])

  // `ready` in the deps: the pane only mounts once pages exist, so the observer
  // must (re)attach after the first pageCount update, not just on dialog mount
  useEffect(() => {
    const pane = paneRef.current
    if (!pane) return
    const ro = new ResizeObserver(applyZoom)
    ro.observe(pane)
    return () => ro.disconnect()
  }, [applyZoom, ready])

  const doPrint = async () => {
    const els = pvPages()
    const root = pvRoot()
    if (!printOrder || printOrder.length === 0 || printing || !root) return
    setPrinting(true)
    const chosen = printOrder.map((i) => els[i]).filter((el): el is HTMLElement => el !== undefined)
    // plain ranges keep the exact preview page list (unselected sheets are skipped in print CSS);
    // pages per sheet / scale to paper / reverse order print freshly laid-out sheets
    const needSheets = perSheet > 1 || scalePaper !== 'none' || reverse
    let pageSize: { widthIn: number; heightIn: number } | undefined
    let sheetsEl: Element | null = null
    if (needSheets) {
      const { sheetW, sheetH } = sheetSizeFor(chosen[0])
      const built = buildSheets(root, chosen, {
        perSheet,
        ...(sheetW !== undefined ? { sheetW } : {}),
        ...(sheetH !== undefined ? { sheetH } : {}),
      })
      sheetsEl = root.lastElementChild
      root.classList.add('pv-has-sheets')
      pageSize = { widthIn: built.sheetW / 96, heightIn: built.sheetH / 96 }
    } else {
      // pv-print-skip only takes effect in the print stylesheet: the on-screen
      // (hidden) preview keeps its layout while unselected sheets don't print.
      const sel = new Set(printOrder)
      els.forEach((el, i) => el.classList.toggle('pv-print-skip', !sel.has(i)))
      const first = chosen[0]
      if (
        first &&
        chosen.every(
          (el) => el.offsetWidth === first.offsetWidth && el.offsetHeight === first.offsetHeight,
        )
      )
        pageSize = { widthIn: first.offsetWidth / 96, heightIn: first.offsetHeight / 96 }
    }
    root.classList.toggle('pv-no-pagecolor', !pageColor)
    const scale = setPrintZoom()
    try {
      const r = await window.desktop.print(scale, printerJob(printerSettings), pageSize)
      if (r.ok) {
        onClose()
        return
      }
      if (r.error) setStatus(t('appPrintFailed', { error: r.error }))
      // not ok without an error = canceled in the system dialog: keep the dialog open
    } finally {
      clearPrintZoom()
      sheetsEl?.remove()
      root.classList.remove('pv-has-sheets', 'pv-no-pagecolor')
      els.forEach((el) => el.classList.remove('pv-print-skip'))
      setPrinting(false)
    }
  }

  const portrait = section ? section.orientation === 'portrait' : true
  const paperKey = paperKeyOf(section)
  const marginKey = marginKeyOf(section)
  const cm = (twips: number | undefined) =>
    String(Math.round(((twips ?? 0) / TWIPS_PER_CM) * 100) / 100)

  const setOrientation = (orientation: 'portrait' | 'landscape') =>
    onPageSetup((s) =>
      s.orientation === orientation
        ? s
        : { ...s, orientation, pageWidth: s.pageHeight, pageHeight: s.pageWidth },
    )
  const setPaper = (key: string) => {
    const p = PAPER_SIZES.find((x) => x.key === key)
    if (!p) return
    onPageSetup((s) => {
      const landscape = s.orientation === 'landscape'
      return { ...s, pageWidth: landscape ? p.h : p.w, pageHeight: landscape ? p.w : p.h }
    })
  }
  const setMarginPreset = (key: string) => {
    const m = MARGIN_PRESETS.find((x) => x.key === key)
    if (!m) return
    onPageSetup((s) => ({
      ...s,
      marginTop: m.top,
      marginRight: m.right,
      marginBottom: m.bottom,
      marginLeft: m.left,
    }))
  }
  const setMarginSide = (
    side: 'marginTop' | 'marginBottom' | 'marginLeft' | 'marginRight',
    raw: string,
  ) => {
    const v = Math.round(Math.max(0, Number(raw) || 0) * TWIPS_PER_CM)
    onPageSetup((s) => ({ ...s, [side]: v }))
  }

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal print-dialog" onClick={(e) => e.stopPropagation()}>
        <h2>{t('appPrintTitle')}</h2>
        <div className="print-dialog-body">
          <div className="print-preview-pane">
            {ready ? (
              <>
                <div className="print-preview-page" ref={paneRef}>
                  <div
                    className={`print-preview-zoom${pageColor ? '' : ' pv-no-pagecolor'}`}
                    ref={hostRef}
                  />
                </div>
                <div className="print-preview-nav">
                  <button
                    aria-label={t('appPrintPrevPage')}
                    title={t('appPrintPrevPage')}
                    disabled={sheetNow <= 0}
                    onClick={() => setPage(sheetNow - 1)}
                  >
                    ‹
                  </button>
                  <span>
                    {sheetCount === 0 ? 0 : sheetNow + 1} / {sheetCount}
                  </span>
                  <button
                    aria-label={t('appPrintNextPage')}
                    title={t('appPrintNextPage')}
                    disabled={sheetNow >= sheetCount - 1}
                    onClick={() => setPage(sheetNow + 1)}
                  >
                    ›
                  </button>
                </div>
              </>
            ) : (
              <div className="print-preview-empty">{t('appPrintRendering')}</div>
            )}
          </div>
          <div className="print-options">
            <PrintSection title={L.secPrinter}>
              <PrinterSettingsFields
                lang={lang}
                value={printerSettings}
                onChange={setPrinterSettings}
                listPrinters={() => window.desktop.listPrinters()}
              />
            </PrintSection>
            <PrintSection title={L.secSettings}>
              <select
                className="ps-select"
                aria-label={L.secSettings}
                data-testid="print-range"
                value={rangeMode}
                onChange={(e) => {
                  setRangeMode(e.target.value as RangeMode)
                  setPage(0)
                }}
              >
                <option value="all">{L.rangeAllPages}</option>
                <option value="current">{L.rangeCurrentPage}</option>
                <option value="custom">{L.rangeCustom}</option>
              </select>
              {rangeMode === 'custom' && (
                <PrintRow label={L.pages}>
                  <input
                    type="text"
                    className={`print-range-input ${rangeInvalid ? 'invalid' : ''}`}
                    placeholder={t('appPrintRangeHint')}
                    value={customRange}
                    onChange={(e) => setCustomRange(e.target.value)}
                  />
                </PrintRow>
              )}
              <select
                className="ps-select"
                aria-label={L.pagesInRange}
                data-testid="print-parity"
                value={parity}
                onChange={(e) => setParity(e.target.value as Parity)}
              >
                <option value="all">{L.pagesInRange}</option>
                <option value="odd">{L.pagesOdd}</option>
                <option value="even">{L.pagesEven}</option>
              </select>
            </PrintSection>
            <PrintSection title={L.secPageSetup} defaultOpen={false}>
              <PrintRow label={L.orientation}>
                <select
                  className="ps-select"
                  data-testid="print-orientation"
                  value={portrait ? 'portrait' : 'landscape'}
                  disabled={!section}
                  onChange={(e) => setOrientation(e.target.value as 'portrait' | 'landscape')}
                >
                  <option value="portrait">{L.portrait}</option>
                  <option value="landscape">{L.landscape}</option>
                </select>
              </PrintRow>
              <PrintRow label={L.paperSize}>
                <select
                  className="ps-select"
                  data-testid="print-paper"
                  value={paperKey}
                  disabled={!section}
                  onChange={(e) => setPaper(e.target.value)}
                >
                  {paperKey === '' && <option value="">{L.marginCustom}</option>}
                  {PAPER_SIZES.map((p) => (
                    <option key={p.key} value={p.key}>
                      {p.name}
                    </option>
                  ))}
                </select>
              </PrintRow>
              <PrintRow label={L.margins}>
                <select
                  className="ps-select"
                  data-testid="print-margins"
                  value={marginKey}
                  disabled={!section}
                  onChange={(e) => setMarginPreset(e.target.value)}
                >
                  {MARGIN_PRESETS.map((m) => (
                    <option key={m.key} value={m.key}>
                      {t(m.nameKey)}
                    </option>
                  ))}
                  {marginKey === 'custom' && <option value="custom">{L.marginCustom}</option>}
                </select>
              </PrintRow>
              {section && (
                <div className="ps-margins">
                  {(
                    [
                      ['marginTop', L.marginTop],
                      ['marginBottom', L.marginBottom],
                      ['marginLeft', L.marginLeft],
                      ['marginRight', L.marginRight],
                    ] as const
                  ).map(([side, label]) => (
                    <label key={side} className="ps-field">
                      <span>{label} (cm)</span>
                      <input
                        type="number"
                        min={0}
                        step={0.1}
                        className="ps-num"
                        key={`${side}:${section[side]}`}
                        defaultValue={cm(section[side])}
                        onBlur={(e) => setMarginSide(side, e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                        }}
                      />
                    </label>
                  ))}
                </div>
              )}
              <PrintRow label={L.pagesPerSheet}>
                <select
                  className="ps-select"
                  data-testid="print-per-sheet"
                  value={perSheet}
                  onChange={(e) => {
                    setPerSheet(Number(e.target.value) as PagesPerSheet)
                    setPage(0)
                  }}
                >
                  {PAGES_PER_SHEET.map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
              </PrintRow>
              <PrintRow label={L.scaleToPaper}>
                <select
                  className="ps-select"
                  data-testid="print-scale-paper"
                  value={scalePaper}
                  onChange={(e) => setScalePaper(e.target.value)}
                >
                  <option value="none">{L.noScaling}</option>
                  {PAPER_SIZES.map((p) => (
                    <option key={p.key} value={p.key}>
                      {p.name}
                    </option>
                  ))}
                </select>
              </PrintRow>
            </PrintSection>
            <PrintSection title={L.secOptions} defaultOpen={false}>
              <PrintCheck label={L.reverseOrder} checked={reverse} onChange={setReverse} />
              <PrintCheck label={L.optBackground} checked={pageColor} onChange={setPageColor} />
            </PrintSection>
            <div className="print-page-count">
              {ready && printOrder && total > 0 ? t('appPrintPageCount', { n: sheetCount }) : ''}
            </div>
          </div>
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button
            className="primary"
            disabled={!ready || !printOrder || total === 0 || printing}
            onClick={() => void doPrint()}
          >
            {printing ? t('appPrintProgress') : t('appPrintTitle')}
          </button>
        </div>
      </div>
    </div>
  )
}
