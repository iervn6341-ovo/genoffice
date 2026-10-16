import { useEffect, useRef, useState } from 'react'
import {
  DEFAULT_HTML_PRINT_OPTIONS,
  MAX_PRINT_SCALE,
  MIN_PRINT_SCALE,
  PRINT_MARGINS,
  PRINT_PAPERS,
  defaultPaperFor,
  parsePrintRanges,
  type HtmlPrintOptions,
  type PrintMargin,
  type PrintPaper,
} from '@genoffice/electron-utils/print-job'
import { useI18n } from '../i18n/locale'

/** Choices survive within the session so the next print starts where the last one ended */
let lastOptions: HtmlPrintOptions | null = null

function initialOptions(): HtmlPrintOptions {
  return (
    lastOptions ?? {
      ...DEFAULT_HTML_PRINT_OPTIONS,
      paper: defaultPaperFor(navigator.language.split('-')[1]),
    }
  )
}

const MARGIN_LABEL = {
  default: 'printSheetMarginDefault',
  none: 'printSheetMarginNone',
  narrow: 'printSheetMarginNarrow',
  wide: 'printSheetMarginWide',
} as const satisfies Record<PrintMargin, string>

/**
 * Print sheet for platforms without a native print panel with preview (Windows, Linux): paper,
 * orientation, margins, scale, header/footer, background and page range.
 */
export function PrintSheet({
  busy,
  onCancel,
  onPrint,
}: {
  busy: boolean
  onCancel: () => void
  onPrint: (options: HtmlPrintOptions) => void
}) {
  const { t } = useI18n()
  const [options, setOptions] = useState<HtmlPrintOptions>(initialOptions)
  const [rangeMode, setRangeMode] = useState<'all' | 'range'>(options.pages ? 'range' : 'all')
  const [invalid, setInvalid] = useState(false)
  const rootRef = useRef<HTMLDivElement>(null)
  useEffect(() => rootRef.current?.querySelector('button')?.focus(), [])

  const patch = (next: Partial<HtmlPrintOptions>) => {
    setOptions((o) => ({ ...o, ...next }))
    setInvalid(false)
  }
  const confirm = () => {
    const pages = rangeMode === 'all' ? '' : options.pages.trim()
    const ranges = parsePrintRanges(pages)
    if (!ranges || (rangeMode === 'range' && ranges.length === 0)) {
      setInvalid(true)
      return
    }
    const final = { ...options, pages }
    lastOptions = final
    onPrint(final)
  }

  return (
    <div className="print-sheet-mask" onMouseDown={onCancel}>
      <div
        ref={rootRef}
        className="print-sheet"
        role="dialog"
        aria-modal="true"
        aria-label={t('printSheetTitle')}
        onMouseDown={(e) => e.stopPropagation()}
        onKeyDown={(e) => {
          if (e.key === 'Escape') onCancel()
          else if (e.key === 'Enter' && (e.target as HTMLElement).tagName !== 'BUTTON') confirm()
        }}
      >
        <div className="print-sheet-title">{t('printSheetTitle')}</div>
        <label className="print-sheet-row">
          <span>{t('printSheetPaper')}</span>
          <select
            value={options.paper}
            onChange={(e) => patch({ paper: e.target.value as PrintPaper })}
          >
            {PRINT_PAPERS.map((p) => (
              <option key={p} value={p}>
                {p}
              </option>
            ))}
          </select>
        </label>
        <label className="print-sheet-row">
          <span>{t('printSheetOrientation')}</span>
          <select
            value={options.landscape ? 'landscape' : 'portrait'}
            onChange={(e) => patch({ landscape: e.target.value === 'landscape' })}
          >
            <option value="portrait">{t('printSheetPortrait')}</option>
            <option value="landscape">{t('printSheetLandscape')}</option>
          </select>
        </label>
        <label className="print-sheet-row">
          <span>{t('printSheetMargins')}</span>
          <select
            value={options.margin}
            onChange={(e) => patch({ margin: e.target.value as PrintMargin })}
          >
            {PRINT_MARGINS.map((m) => (
              <option key={m} value={m}>
                {t(MARGIN_LABEL[m])}
              </option>
            ))}
          </select>
        </label>
        <label className="print-sheet-row">
          <span>{t('printSheetScale')}</span>
          <input
            type="number"
            min={MIN_PRINT_SCALE}
            max={MAX_PRINT_SCALE}
            value={options.scale}
            onChange={(e) => patch({ scale: Number(e.target.value) || 100 })}
          />
        </label>
        <fieldset className="print-sheet-group">
          <legend>{t('printSheetPages')}</legend>
          <label className="print-sheet-check">
            <input
              type="radio"
              name="print-sheet-range"
              checked={rangeMode === 'all'}
              onChange={() => setRangeMode('all')}
            />
            <span>{t('printSheetPagesAll')}</span>
          </label>
          <label className="print-sheet-check">
            <input
              type="radio"
              name="print-sheet-range"
              checked={rangeMode === 'range'}
              onChange={() => setRangeMode('range')}
            />
            <span>{t('printSheetPagesRange')}</span>
          </label>
          {rangeMode === 'range' && (
            <input
              className={`print-sheet-input${invalid ? ' invalid' : ''}`}
              value={options.pages}
              placeholder={t('printSheetPagesHint')}
              aria-invalid={invalid}
              autoFocus
              onChange={(e) => patch({ pages: e.target.value })}
            />
          )}
          {invalid && <div className="print-sheet-error">{t('printSheetPagesInvalid')}</div>}
        </fieldset>
        <label className="print-sheet-check">
          <input
            type="checkbox"
            checked={options.headerFooter}
            onChange={(e) => patch({ headerFooter: e.target.checked })}
          />
          <span>{t('printSheetHeaderFooter')}</span>
        </label>
        <label className="print-sheet-check">
          <input
            type="checkbox"
            checked={options.background}
            onChange={(e) => patch({ background: e.target.checked })}
          />
          <span>{t('printSheetBackground')}</span>
        </label>
        <div className="print-sheet-actions">
          <button type="button" className="print-sheet-btn" onClick={onCancel}>
            {t('printSheetCancel')}
          </button>
          <button
            type="button"
            className="print-sheet-btn primary"
            disabled={busy}
            onClick={confirm}
          >
            {t('printSheetPrint')}
          </button>
        </div>
      </div>
    </div>
  )
}
