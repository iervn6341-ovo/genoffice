import { useCallback, useEffect, useRef, useState } from 'react'
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

import { useI18n } from './i18n/locale'
import { MARGIN_PRESETS, type EffectivePageSetup } from './print-settings'
import type { PageSetupJournalState } from './edit-journal'
import type { PrintScopeOptions } from './page-layout-actions'

/// OOXML paper-size codes offered by the pane (Excel's Paper size list).
const PAPERS: ReadonlyArray<readonly [number, string]> = [
  [8, 'A3'],
  [9, 'A4'],
  [11, 'A5'],
  [12, 'B4 (JIS)'],
  [13, 'B5 (JIS)'],
  [1, 'Letter'],
  [5, 'Legal'],
  [7, 'Executive'],
  [3, 'Tabloid'],
]

type Scope = PrintScopeOptions['scope']
type Scaling = 'none' | 'fitSheet' | 'fitCols' | 'fitRows' | 'custom'

function scalingOf(setup: EffectivePageSetup): Scaling {
  if (setup.fitToPage) {
    if (setup.fitToWidth === 1 && setup.fitToHeight === 1) return 'fitSheet'
    if (setup.fitToWidth === 1 && setup.fitToHeight === 0) return 'fitCols'
    if (setup.fitToWidth === 0 && setup.fitToHeight === 1) return 'fitRows'
    return 'custom'
  }
  return setup.scale === 100 ? 'none' : 'custom'
}

function marginKeyOf(setup: EffectivePageSetup): 'normal' | 'wide' | 'narrow' | 'custom' {
  for (const key of ['normal', 'wide', 'narrow'] as const) {
    const m = MARGIN_PRESETS[key]
    const c = setup.margins
    if (
      Math.abs(m.left - c.left) < 0.005 &&
      Math.abs(m.right - c.right) < 0.005 &&
      Math.abs(m.top - c.top) < 0.005 &&
      Math.abs(m.bottom - c.bottom) < 0.005
    )
      return key
  }
  return 'custom'
}

/// Print sheet grouped like Excel's File → Print pane: Printer / Settings (what to print,
/// pages) / Page Setup (orientation, paper, margins, scaling) / Options (gridlines, headings,
/// header & footer). Page Setup edits go to the active sheet's page layout, exactly as in Excel;
/// without a known printer the system dialog opens.
export function PrintDialog({
  readSetup,
  applyPatch,
  onPrint,
  onHeaderFooter,
  onClose,
}: {
  readonly readSetup: () => EffectivePageSetup | null
  readonly applyPatch: (patch: PageSetupJournalState) => Promise<boolean>
  readonly onPrint: (
    job: ReturnType<typeof printerJob>,
    scope: PrintScopeOptions,
  ) => Promise<boolean>
  readonly onHeaderFooter: () => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t, lang } = useI18n()
  const L = printOfficeLabels(lang)
  const [settings, setSettings] = useState<PrinterSettings>(DEFAULT_PRINTER_SETTINGS)
  const [busy, setBusy] = useState(false)
  const [scope, setScope] = useState<Scope>('active')
  const [ignoreArea, setIgnoreArea] = useState(false)
  const [pageFrom, setPageFrom] = useState('')
  const [pageTo, setPageTo] = useState('')
  const [setup, setSetup] = useState<EffectivePageSetup | null>(() => readSetup())
  const [customScale, setCustomScale] = useState('')
  // "Custom scaling" picked while the sheet still prints at 100%: keep the % field on screen
  const [customPicked, setCustomPicked] = useState(false)
  const readRef = useRef(readSetup)
  readRef.current = readSetup

  const refresh = useCallback(() => setSetup(readRef.current()), [])
  useEffect(() => {
    // the page-setup ops apply asynchronously: keep the pane's echo in step
    const timer = window.setInterval(refresh, 300)
    return () => window.clearInterval(timer)
  }, [refresh])

  useEffect(() => {
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        event.stopPropagation()
        onClose()
      }
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  }, [onClose])

  const patch = (p: PageSetupJournalState): void => {
    void applyPatch(p).then(refresh)
  }

  const print = async (): Promise<void> => {
    setBusy(true)
    try {
      const from = Math.floor(Number(pageFrom))
      const to = Math.floor(Number(pageTo))
      const pages =
        scope !== 'workbook' && from >= 1
          ? { pageFrom: from, ...(to >= from ? { pageTo: to } : {}) }
          : {}
      // a canceled system dialog keeps this sheet open
      const sent = await onPrint(printerJob(settings), {
        scope,
        ...(ignoreArea ? { ignorePrintArea: true } : {}),
        ...pages,
      })
      if (sent) onClose()
    } finally {
      setBusy(false)
    }
  }

  const scaling: Scaling =
    customPicked && setup && !setup.fitToPage && setup.scale === 100
      ? 'custom'
      : setup
        ? scalingOf(setup)
        : 'none'
  const margin = setup ? marginKeyOf(setup) : 'normal'
  const paperKnown = setup ? PAPERS.some(([code]) => code === setup.paperSize) : true

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog print-sheet-dialog"
        role="dialog"
        aria-label={t('appPrintCheck')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appPrintCheck')}</header>
        <section className="dialog-body">
          <PrintSection title={L.secPrinter}>
            <PrinterSettingsFields
              lang={lang}
              value={settings}
              onChange={setSettings}
              listPrinters={() => window.desktopApi.listPrinters()}
            />
          </PrintSection>
          <PrintSection title={L.secSettings}>
            <select
              className="ps-select"
              aria-label={L.secSettings}
              data-testid="print-scope"
              value={scope}
              onChange={(e) => setScope(e.target.value as Scope)}
            >
              <option value="active">{L.scopeActive}</option>
              <option value="workbook">{L.scopeWorkbook}</option>
              <option value="selection">{L.rangeSelection}</option>
            </select>
            <PrintCheck
              label={L.ignorePrintArea}
              checked={ignoreArea}
              disabled={scope === 'selection'}
              onChange={setIgnoreArea}
            />
            <div className="ps-field">
              <span>{L.pageFrom}</span>
              <span className="ps-pages">
                <input
                  type="number"
                  min={1}
                  className="ps-num"
                  data-testid="print-page-from"
                  value={pageFrom}
                  disabled={scope === 'workbook'}
                  onChange={(e) => setPageFrom(e.target.value)}
                />
                <span>{L.pageTo}</span>
                <input
                  type="number"
                  min={1}
                  className="ps-num"
                  data-testid="print-page-to"
                  value={pageTo}
                  disabled={scope === 'workbook'}
                  onChange={(e) => setPageTo(e.target.value)}
                />
              </span>
            </div>
          </PrintSection>
          <PrintSection title={L.secPageSetup} defaultOpen={false}>
            <PrintRow label={L.orientation}>
              <select
                className="ps-select"
                data-testid="print-orientation"
                value={setup?.orientation ?? 'portrait'}
                onChange={(e) => patch({ orientation: e.target.value as 'portrait' | 'landscape' })}
              >
                <option value="portrait">{L.portrait}</option>
                <option value="landscape">{L.landscape}</option>
              </select>
            </PrintRow>
            <PrintRow label={L.paperSize}>
              <select
                className="ps-select"
                data-testid="print-paper"
                value={setup?.paperSize ?? 9}
                onChange={(e) => patch({ paperSize: Number(e.target.value) })}
              >
                {!paperKnown && setup && (
                  <option value={setup.paperSize}>#{setup.paperSize}</option>
                )}
                {PAPERS.map(([code, name]) => (
                  <option key={code} value={code}>
                    {name}
                  </option>
                ))}
              </select>
            </PrintRow>
            <PrintRow label={L.margins}>
              <select
                className="ps-select"
                data-testid="print-margins"
                value={margin}
                onChange={(e) => {
                  if (e.target.value !== 'custom')
                    patch({ margins: e.target.value as 'normal' | 'wide' | 'narrow' })
                }}
              >
                <option value="normal">{L.marginNormal}</option>
                <option value="wide">{L.marginWide}</option>
                <option value="narrow">{L.marginNarrow}</option>
                {margin === 'custom' && <option value="custom">{L.marginCustom}</option>}
              </select>
            </PrintRow>
            <PrintRow label={L.scaling}>
              <select
                className="ps-select"
                data-testid="print-scaling"
                value={scaling}
                onChange={(e) => {
                  const v = e.target.value as Scaling
                  setCustomPicked(v === 'custom')
                  if (v === 'none') patch({ scale: 100 })
                  else if (v === 'fitSheet') patch({ fitToWidth: 1, fitToHeight: 1 })
                  else if (v === 'fitCols') patch({ fitToWidth: 1, fitToHeight: 0 })
                  else if (v === 'fitRows') patch({ fitToWidth: 0, fitToHeight: 1 })
                  else
                    patch({
                      scale: Math.min(
                        400,
                        Math.max(10, Number(customScale) || setup?.scale || 100),
                      ),
                    })
                }}
              >
                <option value="none">{L.noScaling}</option>
                <option value="fitSheet">{L.scFitSheet}</option>
                <option value="fitCols">{L.scFitCols}</option>
                <option value="fitRows">{L.scFitRows}</option>
                <option value="custom">{L.scCustom}</option>
              </select>
            </PrintRow>
            {scaling === 'custom' && !setup?.fitToPage && (
              <PrintRow label={L.adjustTo}>
                <input
                  type="number"
                  min={10}
                  max={400}
                  className="ps-num"
                  data-testid="print-scale"
                  key={setup?.scale}
                  defaultValue={setup?.scale ?? 100}
                  onChange={(e) => setCustomScale(e.target.value)}
                  onBlur={(e) => {
                    const n = Math.round(Number(e.target.value))
                    if (n >= 10 && n <= 400) patch({ scale: n })
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                  }}
                />
              </PrintRow>
            )}
          </PrintSection>
          <PrintSection title={L.secOptions} defaultOpen={false}>
            <PrintCheck
              label={L.gridlines}
              checked={setup?.printGridlines ?? false}
              onChange={(v) => patch({ printGridlines: v })}
            />
            <PrintCheck
              label={L.headings}
              checked={setup?.printHeadings ?? false}
              onChange={(v) => patch({ printHeadings: v })}
            />
            <button type="button" className="ps-link" onClick={onHeaderFooter}>
              {L.headerFooterBtn}
            </button>
          </PrintSection>
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="primary-action" disabled={busy} onClick={() => void print()}>
            {t('appPrintCheck')}
          </button>
        </div>
      </div>
    </div>
  )
}
