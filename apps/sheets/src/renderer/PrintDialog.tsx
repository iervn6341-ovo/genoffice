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
import { isDefaultPrintOptions, type PrintErrorsAs, type PrintOnlyOptions } from './print-options'
import { MARGIN_PRESETS, type EffectivePageSetup, type PrintMargins } from './print-settings'
import type { PageSetupJournalState } from './edit-journal'
import type { PrintScopeOptions, PrintTargets } from './page-layout-actions'
import { PrintPreviewPane } from './PrintPreviewPane'
import { previewKey } from './print-preview-plan'

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
type Scaling = 'none' | 'fitSheet' | 'fitCols' | 'fitRows' | 'fitCustom' | 'custom'

function scalingOf(setup: EffectivePageSetup): Scaling {
  if (setup.fitToPage) {
    if (setup.fitToWidth === 1 && setup.fitToHeight === 1) return 'fitSheet'
    if (setup.fitToWidth === 1 && setup.fitToHeight === 0) return 'fitCols'
    if (setup.fitToWidth === 0 && setup.fitToHeight === 1) return 'fitRows'
    return 'fitCustom'
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

/// Print sheet like Excel's File → Print pane: a page preview next to Printer / Settings (what to
/// print, pages) / Page Setup (orientation, paper, margins, scaling) / Options (gridlines,
/// headings, header & footer). Page Setup edits go to the active sheet's page layout, exactly as
/// in Excel. Windows and Linux only: macOS goes straight to the system print panel (print-entry.ts).
export function PrintDialog({
  getTargets,
  readSetup,
  applyPatch,
  onPrint,
  onHeaderFooter,
  onClose,
}: {
  readonly getTargets: (scope: PrintScopeOptions) => PrintTargets
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
  // print-only options (not saved to the file): see print-options.ts
  const [customMargins, setCustomMargins] = useState<PrintMargins | null>(null)
  const [centerH, setCenterH] = useState(false)
  const [centerV, setCenterV] = useState(false)
  const [errorsAs, setErrorsAs] = useState<PrintErrorsAs>('displayed')
  const [blackWhite, setBlackWhite] = useState(false)
  const [draft, setDraft] = useState(false)
  const [firstPageNumber, setFirstPageNumber] = useState('')
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

  const firstPage = Math.floor(Number(firstPageNumber))
  const options: PrintOnlyOptions = {
    ...(customMargins ? { margins: customMargins } : {}),
    ...(centerH ? { centerHorizontally: true } : {}),
    ...(centerV ? { centerVertically: true } : {}),
    ...(errorsAs !== 'displayed' ? { errorsAs } : {}),
    ...(blackWhite ? { blackAndWhite: true } : {}),
    ...(draft ? { draft: true } : {}),
    ...(firstPageNumber !== '' && firstPage >= 1 && firstPage <= 9999 && firstPage !== 1
      ? { firstPageNumber: firstPage }
      : {}),
  }

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
      const jobScope: PrintScopeOptions = {
        scope,
        ...(ignoreArea ? { ignorePrintArea: true } : {}),
        ...(isDefaultPrintOptions(options) ? {} : { options }),
        ...pages,
      }
      // a canceled system dialog keeps this sheet open
      const sent = await onPrint(printerJob(settings), jobScope)
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
  const jobKey = previewKey({ scope, ignoreArea, setup, options })
  const paperKnown = setup ? PAPERS.some(([code]) => code === setup.paperSize) : true

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog print-sheet-dialog with-preview"
        role="dialog"
        aria-label={t('appPrintCheck')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appPrintCheck')}</header>
        <div className="print-sheet-layout">
          <PrintPreviewPane
            getTargets={() =>
              getTargets({
                scope,
                ignorePrintArea: ignoreArea,
                ...(isDefaultPrintOptions(options) ? {} : { options }),
              })
            }
            jobKey={jobKey}
          />
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
              {
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
              }
            </PrintSection>
            <PrintSection title={L.secPageSetup} defaultOpen={false}>
              <PrintRow label={L.orientation}>
                <select
                  className="ps-select"
                  data-testid="print-orientation"
                  value={setup?.orientation ?? 'portrait'}
                  onChange={(e) =>
                    patch({ orientation: e.target.value as 'portrait' | 'landscape' })
                  }
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
                  value={customMargins ? 'custom' : margin}
                  onChange={(e) => {
                    if (e.target.value === 'custom') {
                      setCustomMargins(customMargins ?? setup?.margins ?? MARGIN_PRESETS.normal)
                    } else {
                      setCustomMargins(null)
                      patch({ margins: e.target.value as 'normal' | 'wide' | 'narrow' })
                    }
                  }}
                >
                  <option value="normal">{L.marginNormal}</option>
                  <option value="wide">{L.marginWide}</option>
                  <option value="narrow">{L.marginNarrow}</option>
                  <option value="custom">{L.marginCustom}</option>
                </select>
              </PrintRow>
              {(customMargins ?? (margin === 'custom' ? setup?.margins : null)) && (
                <div className="print-margin-fields" data-testid="print-margin-fields">
                  {(
                    [
                      ['top', 'appPrintMarginTop'],
                      ['bottom', 'appPrintMarginBottom'],
                      ['left', 'appPrintMarginLeft'],
                      ['right', 'appPrintMarginRight'],
                      ['header', 'appPrintMarginHeader'],
                      ['footer', 'appPrintMarginFooter'],
                    ] as const
                  ).map(([edge, label]) => {
                    const current = customMargins ?? setup?.margins ?? MARGIN_PRESETS.normal
                    return (
                      <PrintRow key={edge} label={t(label)}>
                        <input
                          type="number"
                          min={0}
                          max={3}
                          step={0.05}
                          className="ps-num"
                          data-testid={`print-margin-${edge}`}
                          key={`${edge}${current[edge]}`}
                          defaultValue={current[edge]}
                          onBlur={(e) => {
                            const n = Number(e.target.value)
                            if (!Number.isFinite(n) || e.target.value === '') return
                            setCustomMargins({
                              ...current,
                              [edge]: Math.round(Math.min(3, Math.max(0, n)) * 100) / 100,
                            })
                          }}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                          }}
                        />
                      </PrintRow>
                    )
                  })}
                </div>
              )}
              <PrintCheck label={t('appPrintCenterH')} checked={centerH} onChange={setCenterH} />
              <PrintCheck label={t('appPrintCenterV')} checked={centerV} onChange={setCenterV} />
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
                    else if (v === 'fitCustom')
                      patch({
                        fitToWidth: setup?.fitToWidth || 1,
                        fitToHeight: setup?.fitToHeight || 2,
                      })
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
                  <option value="fitCustom">{t('appPrintFitCustom')}</option>
                  <option value="custom">{L.scCustom}</option>
                </select>
              </PrintRow>
              {scaling === 'fitCustom' && setup && (
                <div className="ps-field print-fit-fields" data-testid="print-fit">
                  <input
                    type="number"
                    min={0}
                    max={99}
                    className="ps-num"
                    aria-label={t('appPrintFitWide')}
                    data-testid="print-fit-wide"
                    key={`w${setup.fitToWidth}`}
                    defaultValue={setup.fitToWidth}
                    onBlur={(e) => {
                      const n = Math.round(Number(e.target.value))
                      if (n >= 0 && n <= 99) patch({ fitToWidth: n })
                    }}
                  />
                  <span>{t('appPrintFitWide')}</span>
                  <input
                    type="number"
                    min={0}
                    max={99}
                    className="ps-num"
                    aria-label={t('appPrintFitTall')}
                    data-testid="print-fit-tall"
                    key={`h${setup.fitToHeight}`}
                    defaultValue={setup.fitToHeight}
                    onBlur={(e) => {
                      const n = Math.round(Number(e.target.value))
                      if (n >= 0 && n <= 99) patch({ fitToHeight: n })
                    }}
                  />
                  <span>{t('appPrintFitTall')}</span>
                </div>
              )}
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
              <PrintRow label={t('appPrintErrorsAs')}>
                <select
                  className="ps-select"
                  data-testid="print-errors-as"
                  value={errorsAs}
                  onChange={(e) => setErrorsAs(e.target.value as PrintErrorsAs)}
                >
                  <option value="displayed">{t('appPrintErrDisplayed')}</option>
                  <option value="blank">{t('appPrintErrBlank')}</option>
                  <option value="dash">{t('appPrintErrDash')}</option>
                  <option value="na">{t('appPrintErrNA')}</option>
                </select>
              </PrintRow>
              <PrintCheck
                label={t('appPrintBlackWhite')}
                checked={blackWhite}
                onChange={setBlackWhite}
              />
              <PrintCheck label={t('appPrintDraft')} checked={draft} onChange={setDraft} />
              <PrintRow label={t('appPrintFirstPage')}>
                <input
                  type="number"
                  min={1}
                  max={9999}
                  className="ps-num"
                  data-testid="print-first-page"
                  placeholder="1"
                  value={firstPageNumber}
                  onChange={(e) => setFirstPageNumber(e.target.value)}
                />
              </PrintRow>
              <p className="ps-note">{t('appPrintOnlyNote')}</p>
              <button type="button" className="ps-link" onClick={onHeaderFooter}>
                {L.headerFooterBtn}
              </button>
            </PrintSection>
          </section>
        </div>
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
