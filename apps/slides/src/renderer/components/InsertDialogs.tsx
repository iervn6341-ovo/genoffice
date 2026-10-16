/**
 * The Insert tab's three dialogs: hyperlink / header & footer / equation.
 * Reuses SettingsModal's .modal-backdrop/.modal styles.
 */
import React, { useEffect, useMemo, useState } from 'react'
import {
  Dropdown,
  MathWorkbench,
  SymbolPicker,
  isSymbolFontFamily,
  mathPreview,
} from '@genoffice/ui'
import { SYMBOL_NAMES, symbolCodePoint } from '@genoffice/office-math'
import { SPECIAL_CHARS, specialCharGlyph } from '../special-chars'
import { useSystemFontFamilies } from '../system-fonts'
import { FONT_FAMILIES } from './ribbon-shared'
import type { LinkTargetOp } from '../../shared/ipc'
import type {
  NotesHandoutsHeaderFooterOptions,
  NotesHandoutsHeaderFooterState,
} from '@genoffice/pptx-engine'
import { useI18n } from '../i18n/locale'
import { formatHeaderFooterDate } from '../header-footer-date'

// ── Hyperlink ─────────────────────────────────────────────────────────────

interface LinkDialogProps {
  /** Current link (for display; null = none) */
  initial: LinkTargetOp | null
  /** Document page count (page-jump dropdown) */
  slideCount: number
  currentSlide: number
  onApply: (target: LinkTargetOp | null) => void
  onClose: () => void
}

export function LinkDialog({
  initial,
  slideCount,
  currentSlide,
  onApply,
  onClose,
}: LinkDialogProps) {
  const { t } = useI18n()
  const [mode, setMode] = useState<'url' | 'slide'>(initial?.kind === 'slide' ? 'slide' : 'url')
  const [url, setUrl] = useState(initial?.kind === 'url' ? initial.url : 'https://')
  const [slideIndex, setSlideIndex] = useState(
    initial?.kind === 'slide' ? initial.slideIndex : (currentSlide + 1) % Math.max(slideCount, 1),
  )

  const apply = () => {
    if (mode === 'url') {
      const trimmed = url.trim()
      if (!trimmed || trimmed === 'https://') return
      onApply({ kind: 'url', url: trimmed })
    } else {
      onApply({ kind: 'slide', slideIndex })
    }
  }

  return (
    // data-keep-edit: opening over a text-edit session must not commit it — the run-level
    // link applies to the saved editor selection after the dialog closes
    <div className="modal-backdrop" data-keep-edit="" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonDlgHyperlink')}</h2>
        <div className="dlg-radio-row">
          <label className="dlg-radio">
            <input type="radio" checked={mode === 'url'} onChange={() => setMode('url')} />
            {t('ribbonDlgWebAddress')}
          </label>
          <label className="dlg-radio">
            <input type="radio" checked={mode === 'slide'} onChange={() => setMode('slide')} />
            {t('ribbonDlgPlaceInDoc')}
          </label>
        </div>
        {mode === 'url' ? (
          <label>
            {t('ribbonDlgAddress')}
            <input
              type="text"
              value={url}
              autoFocus
              onChange={(e) => setUrl(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && apply()}
              placeholder="https://example.com"
            />
          </label>
        ) : (
          <label>
            {t('ribbonDlgGoTo')}
            <Dropdown
              value={String(slideIndex)}
              ariaLabel={t('ribbonDlgGoTo')}
              options={Array.from({ length: slideCount }, (_, i) => ({
                value: String(i),
                label:
                  t('ribbonSlideN', { n: i + 1 }) +
                  (i === currentSlide ? t('ribbonCurrentSlideSuffix') : ''),
              }))}
              onPick={(v) => setSlideIndex(Number(v))}
            />
          </label>
        )}
        <div className="modal-actions">
          {initial && (
            <button className="dlg-danger" onClick={() => onApply(null)}>
              {t('ribbonRemoveLink')}
            </button>
          )}
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button className="primary" onClick={apply}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

// ── Header & footer ───────────────────────────────────────────────────────

export type HeaderFooterDialogApply =
  | ({ tab: 'slides' } & {
      footer: string | null
      slideNum: boolean
      date: string | null
      dateAuto: boolean
      dateFormat: string
      dontShowOnTitleSlide: boolean
      firstSlideNum: number
      scope: 'current' | 'all'
    })
  | { tab: 'notes-handouts'; settings: NotesHandoutsHeaderFooterOptions }

interface HeaderFooterDialogProps {
  onApply: (opts: HeaderFooterDialogApply) => void
  onClose: () => void
  /** Display: current page state */
  initial: {
    footer: string | null
    slideNum: boolean
    date: string | null
    dateAuto: boolean
    dateFormat: string
    dontShowOnTitleSlide: boolean
    firstSlideNum: number
    isTitleSlide: boolean
    slideNumber: number
    notesHandouts: NotesHandoutsHeaderFooterState
  }
}

export function HeaderFooterDialog({ initial, onApply, onClose }: HeaderFooterDialogProps) {
  const { t, dateLocale } = useI18n()
  const [tab, setTab] = useState<'slides' | 'notes-handouts'>('slides')
  const [dateOn, setDateOn] = useState(!!initial.date)
  const [dateAuto, setDateAuto] = useState(initial.date ? initial.dateAuto : true)
  const [dateFormat, setDateFormat] = useState(initial.dateFormat)
  const [dateText, setDateText] = useState(initial.date ?? new Date().toLocaleDateString())
  const [autoDateText, setAutoDateText] = useState(
    initial.dateAuto && initial.date
      ? initial.date
      : (formatHeaderFooterDate(new Date(), initial.dateFormat, dateLocale) ?? initial.date ?? ''),
  )
  const [numOn, setNumOn] = useState(initial.slideNum)
  const [footerOn, setFooterOn] = useState(!!initial.footer)
  const [footerText, setFooterText] = useState(initial.footer ?? '')
  const [excludeTitle, setExcludeTitle] = useState(initial.dontShowOnTitleSlide)
  const [firstSlideNumText, setFirstSlideNumText] = useState(String(initial.firstSlideNum))
  const firstSlideNum = Number(firstSlideNumText)
  const firstSlideNumValid =
    firstSlideNumText.trim() !== '' &&
    Number.isInteger(firstSlideNum) &&
    firstSlideNum >= -9999 &&
    firstSlideNum <= 9999
  const [nh, setNh] = useState<NotesHandoutsHeaderFooterState>(initial.notesHandouts)
  const [nhFixedDate, setNhFixedDate] = useState(initial.notesHandouts.date)
  const [nhAutoDate, setNhAutoDate] = useState(
    initial.notesHandouts.dateAuto && initial.notesHandouts.date
      ? initial.notesHandouts.date
      : (formatHeaderFooterDate(new Date(), initial.notesHandouts.dateFormat, dateLocale) ??
          initial.notesHandouts.date),
  )
  const dateFormats = Array.from({ length: 13 }, (_, i) => `datetime${i + 1}`)
  if (!dateFormats.includes(dateFormat)) dateFormats.push(dateFormat)
  const exampleDate = new Date(2007, 9, 12, 16, 28, 34)
  const hiddenOnPreview = excludeTitle && initial.isTitleSlide
  const dateValue = (dateAuto ? autoDateText : dateText).trim()
  const apply = (scope: 'current' | 'all') =>
    onApply({
      tab: 'slides',
      footer: footerOn && footerText.trim() ? footerText.trim() : null,
      slideNum: numOn,
      date: dateOn && dateValue ? dateValue : null,
      dateAuto,
      dateFormat,
      dontShowOnTitleSlide: excludeTitle,
      firstSlideNum,
      scope,
    })
  const applyNotes = () =>
    onApply({
      tab: 'notes-handouts',
      settings: {
        ...nh,
        date: nh.dateAuto ? nhAutoDate : nhFixedDate,
      },
    })

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonDlgHeaderFooter')}</h2>
        <div className="hf-tabs" role="tablist" aria-label={t('ribbonDlgHeaderFooter')}>
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'slides'}
            data-testid="hf-tab-slides"
            onClick={() => setTab('slides')}
          >
            {t('ribbonDlgSlidesTab')}
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'notes-handouts'}
            data-testid="hf-tab-notes-handouts"
            onClick={() => setTab('notes-handouts')}
          >
            {t('ribbonDlgNotesHandoutsTab')}
          </button>
        </div>
        {tab === 'slides' ? (
          <>
            <label className="dlg-check">
              <input
                data-testid="hf-date-enabled"
                type="checkbox"
                checked={dateOn}
                onChange={(e) => setDateOn(e.target.checked)}
              />
              {t('ribbonDlgDateTime')}
            </label>
            {dateOn && (
              <div className="dlg-indent">
                <label className="dlg-check">
                  <input
                    data-testid="hf-date-auto"
                    type="radio"
                    name="hf-date-mode"
                    checked={dateAuto}
                    onChange={() => {
                      setDateAuto(true)
                      if (!initial.dateAuto)
                        setAutoDateText(
                          formatHeaderFooterDate(new Date(), dateFormat, dateLocale) ??
                            autoDateText,
                        )
                    }}
                  />
                  {t('ribbonDlgAutoUpdate')}
                </label>
                {dateAuto && (
                  <label className="dlg-check">
                    {t('ribbonDlgDateFormat')}
                    <select
                      data-testid="hf-date-format"
                      value={dateFormat}
                      onChange={(e) => {
                        const format = e.target.value
                        setDateFormat(format)
                        setAutoDateText(
                          formatHeaderFooterDate(new Date(), format, dateLocale) ??
                            initial.date ??
                            autoDateText,
                        )
                      }}
                    >
                      {dateFormats.map((format, i) => (
                        <option key={format} value={format}>
                          {i < 13
                            ? formatHeaderFooterDate(exampleDate, format, dateLocale)
                            : t('ribbonDlgImportedDateFormat')}
                        </option>
                      ))}
                    </select>
                  </label>
                )}
                <label className="dlg-check">
                  <input
                    data-testid="hf-date-fixed"
                    type="radio"
                    name="hf-date-mode"
                    checked={!dateAuto}
                    onChange={() => setDateAuto(false)}
                  />
                  {t('ribbonDlgDateFixed')}
                </label>
                <input
                  data-testid="hf-date-text"
                  type="text"
                  value={dateAuto ? autoDateText : dateText}
                  onChange={(e) => setDateText(e.target.value)}
                  placeholder={t('ribbonDlgDateExample')}
                  readOnly={dateAuto}
                />
              </div>
            )}
            <label className="dlg-check">
              <input
                data-testid="hf-slide-number"
                type="checkbox"
                checked={numOn}
                onChange={(e) => setNumOn(e.target.checked)}
              />
              {t('ribbonDlgSlideNum')}
            </label>
            <label className="dlg-check">
              {t('ribbonDlgSlideStartAt')}
              <input
                data-testid="hf-slide-start-at"
                type="number"
                min={-9999}
                max={9999}
                step={1}
                value={firstSlideNumText}
                onChange={(e) => setFirstSlideNumText(e.target.value)}
                aria-invalid={!firstSlideNumValid}
              />
            </label>
            <label className="dlg-check">
              <input
                data-testid="hf-footer-enabled"
                type="checkbox"
                checked={footerOn}
                onChange={(e) => setFooterOn(e.target.checked)}
              />
              {t('ribbonDlgFooter')}
            </label>
            {footerOn && (
              <div className="dlg-indent">
                <input
                  data-testid="hf-footer-text"
                  type="text"
                  value={footerText}
                  autoFocus
                  onChange={(e) => setFooterText(e.target.value)}
                  placeholder={t('ribbonDlgFooterPlaceholder')}
                />
              </div>
            )}
            <label className="dlg-check">
              <input
                data-testid="hf-exclude-title"
                type="checkbox"
                checked={excludeTitle}
                onChange={(e) => setExcludeTitle(e.target.checked)}
              />
              {t('ribbonDlgNoTitleFooter')}
            </label>
            <div style={{ marginTop: 12 }}>
              <div>{t('ribbonDlgPreview')}</div>
              <div
                data-testid="hf-preview"
                style={{
                  position: 'relative',
                  aspectRatio: '16 / 9',
                  width: 240,
                  maxWidth: '100%',
                  border: '1px solid var(--border)',
                  background: 'var(--surface)',
                  color: 'var(--text-secondary)',
                  fontSize: 10,
                }}
              >
                {!hiddenOnPreview && (
                  <div
                    style={{
                      position: 'absolute',
                      left: '5%',
                      right: '5%',
                      bottom: '5%',
                      display: 'flex',
                      alignItems: 'end',
                      justifyContent: 'space-between',
                      gap: 4,
                    }}
                  >
                    <span data-testid="hf-preview-date">
                      {dateOn ? (dateAuto ? autoDateText : dateText) : ''}
                    </span>
                    <span data-testid="hf-preview-footer">{footerOn ? footerText : ''}</span>
                    <span data-testid="hf-preview-number">
                      {numOn && firstSlideNumValid
                        ? firstSlideNum + initial.slideNumber - initial.firstSlideNum
                        : ''}
                    </span>
                  </div>
                )}
              </div>
            </div>
          </>
        ) : (
          <>
            <label className="dlg-check">
              <input
                data-testid="nh-header-enabled"
                type="checkbox"
                checked={nh.headerEnabled}
                onChange={(e) => setNh({ ...nh, headerEnabled: e.target.checked })}
              />
              {t('ribbonDlgHeader')}
            </label>
            {nh.headerEnabled && (
              <div className="dlg-indent">
                <input
                  data-testid="nh-header-text"
                  type="text"
                  value={nh.header}
                  onChange={(e) => setNh({ ...nh, header: e.target.value })}
                  placeholder={t('ribbonDlgHeaderPlaceholder')}
                />
              </div>
            )}
            <label className="dlg-check">
              <input
                data-testid="nh-date-enabled"
                type="checkbox"
                checked={nh.dateEnabled}
                onChange={(e) => setNh({ ...nh, dateEnabled: e.target.checked })}
              />
              {t('ribbonDlgDateTime')}
            </label>
            {nh.dateEnabled && (
              <div className="dlg-indent">
                <label className="dlg-check">
                  <input
                    data-testid="nh-date-auto"
                    type="radio"
                    name="nh-date-mode"
                    checked={nh.dateAuto}
                    onChange={() => setNh({ ...nh, dateAuto: true })}
                  />
                  {t('ribbonDlgAutoUpdate')}
                </label>
                {nh.dateAuto && (
                  <label className="dlg-check">
                    {t('ribbonDlgDateFormat')}
                    <select
                      data-testid="nh-date-format"
                      value={nh.dateFormat}
                      onChange={(e) => {
                        const format = e.target.value
                        setNh({ ...nh, dateFormat: format })
                        setNhAutoDate(
                          formatHeaderFooterDate(new Date(), format, dateLocale) ?? nhAutoDate,
                        )
                      }}
                    >
                      {[
                        ...dateFormats,
                        ...(dateFormats.includes(nh.dateFormat) ? [] : [nh.dateFormat]),
                      ].map((format, i) => (
                        <option key={format} value={format}>
                          {i < 13
                            ? formatHeaderFooterDate(exampleDate, format, dateLocale)
                            : t('ribbonDlgImportedDateFormat')}
                        </option>
                      ))}
                    </select>
                  </label>
                )}
                <label className="dlg-check">
                  <input
                    data-testid="nh-date-fixed"
                    type="radio"
                    name="nh-date-mode"
                    checked={!nh.dateAuto}
                    onChange={() => setNh({ ...nh, dateAuto: false })}
                  />
                  {t('ribbonDlgDateFixed')}
                </label>
                <input
                  data-testid="nh-date-text"
                  type="text"
                  value={nh.dateAuto ? nhAutoDate : nhFixedDate}
                  readOnly={nh.dateAuto}
                  onChange={(e) => setNhFixedDate(e.target.value)}
                  placeholder={t('ribbonDlgDateExample')}
                />
              </div>
            )}
            <label className="dlg-check">
              <input
                data-testid="nh-page-number"
                type="checkbox"
                checked={nh.pageNum}
                onChange={(e) => setNh({ ...nh, pageNum: e.target.checked })}
              />
              {t('ribbonDlgPageNum')}
            </label>
            <label className="dlg-check">
              <input
                data-testid="nh-footer-enabled"
                type="checkbox"
                checked={nh.footerEnabled}
                onChange={(e) => setNh({ ...nh, footerEnabled: e.target.checked })}
              />
              {t('ribbonDlgFooter')}
            </label>
            {nh.footerEnabled && (
              <div className="dlg-indent">
                <input
                  data-testid="nh-footer-text"
                  type="text"
                  value={nh.footer}
                  onChange={(e) => setNh({ ...nh, footer: e.target.value })}
                  placeholder={t('ribbonDlgFooterPlaceholder')}
                />
              </div>
            )}
            <div style={{ marginTop: 12 }}>
              {t('ribbonDlgPreview')}
              <div data-testid="nh-preview" className="nh-preview">
                <span data-testid="nh-preview-header">{nh.headerEnabled ? nh.header : ''}</span>
                <div className="nh-preview-bottom">
                  <span data-testid="nh-preview-date">
                    {nh.dateEnabled ? (nh.dateAuto ? nhAutoDate : nhFixedDate) : ''}
                  </span>
                  <span data-testid="nh-preview-footer">{nh.footerEnabled ? nh.footer : ''}</span>
                  <span data-testid="nh-preview-number">
                    {nh.pageNum ? initial.slideNumber : ''}
                  </span>
                </div>
              </div>
            </div>
          </>
        )}
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          {tab === 'slides' && (
            <button
              data-testid="hf-apply-current"
              disabled={!firstSlideNumValid}
              onClick={() => apply('current')}
            >
              {t('ribbonApplyBtn')}
            </button>
          )}
          <button
            data-testid={tab === 'slides' ? 'hf-apply-all' : 'nh-apply-all'}
            className="primary"
            disabled={tab === 'slides' && !firstSlideNumValid}
            onClick={tab === 'slides' ? () => apply('all') : applyNotes}
          >
            {t('ribbonApplyToAllBtn')}
          </button>
        </div>
      </div>
    </div>
  )
}

// ── Equation ───────────────────────────────────────────────────────

interface EquationDialogProps {
  onInsert: (text: string) => void
  onClose: () => void
  initial?: string
  unsupported?: boolean
}

export function EquationDialog({
  onInsert,
  onClose,
  initial = '',
  unsupported = false,
}: EquationDialogProps) {
  const { t } = useI18n()
  const [text, setText] = useState(initial)
  const preview = mathPreview(text)

  // Esc closes
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="modal-backdrop" onClick={onClose} data-testid="equation-dialog">
      <div className="modal modal-wide" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonDlgInsertEquation')}</h2>
        {unsupported ? (
          <p role="status">{t('appStatusEquationPreview')}</p>
        ) : (
          <MathWorkbench
            latex={text}
            onLatexChange={setText}
            labels={{
              input: t('ribbonMathInput'),
              structures: t('ribbonMathStructures'),
              professional: t('ribbonMathProfessional'),
              linear: t('ribbonMathLinear'),
              preview: t('ribbonDlgPreview'),
              empty: t('ribbonMathInput'),
              symbols: t('ribbonSymbol'),
              symbolSearch: t('ribbonSymbolSearch'),
              symbolSubsetNames: {
                greek: t('ribbonSymbolGreek'),
                math: t('ribbonSymbolMath'),
                arrows: t('ribbonSymbolArrows'),
              },
              categoryNames: {
                fraction: t('ribbonMathCatFraction'),
                script: t('ribbonMathCatScript'),
                radical: t('ribbonMathCatRadical'),
                integral: t('ribbonMathCatIntegral'),
                largeOperator: t('ribbonMathCatLargeOperator'),
                delimiter: t('ribbonMathCatDelimiter'),
                function: t('ribbonMathCatFunction'),
                accent: t('ribbonMathCatAccent'),
                limitLog: t('ribbonMathCatLimitLog'),
                operator: t('ribbonMathCatOperator'),
                matrix: t('ribbonMathCatMatrix'),
              },
            }}
          />
        )}
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button
            className="primary"
            data-testid="equation-commit"
            disabled={unsupported || !preview || !('mathml' in preview)}
            onClick={() => onInsert(text.trim())}
          >
            {t('ribbonInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

export function SymbolDialog({
  onInsert,
  onClose,
}: {
  onInsert: (char: string, font: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [tab, setTab] = useState<'symbols' | 'special'>('symbols')
  const { families, load } = useSystemFontFamilies()
  useEffect(load, [load])
  const fonts = useMemo(
    () => [
      ...new Set([
        'Cambria Math',
        'Segoe UI Symbol',
        ...FONT_FAMILIES,
        ...families.filter((f) => !isSymbolFontFamily(f)),
      ]),
    ],
    [families],
  )
  const [info, setInfo] = useState<string>('')
  const [specialSel, setSpecialSel] = useState<string>(SPECIAL_CHARS[0]!.char)
  const shown = tab === 'symbols' ? info : specialSel
  const shownName = SPECIAL_CHARS.find((s) => s.char === shown)?.name ?? SYMBOL_NAMES[shown] ?? ''
  return (
    <div className="modal-backdrop" data-keep-edit data-testid="symbol-dialog" onClick={onClose}>
      <div className="modal modal-wide" data-keep-edit onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonSymbol')}</h2>
        <div className="go-math-view-switch" role="tablist" data-keep-edit>
          {(['symbols', 'special'] as const).map((id) => (
            <button
              key={id}
              type="button"
              role="tab"
              aria-selected={tab === id}
              className={tab === id ? 'active' : ''}
              data-keep-edit
              data-testid={`symbol-tab-${id}`}
              onClick={() => setTab(id)}
            >
              {id === 'symbols' ? t('ribbonSymbolTabSymbols') : t('ribbonSymbolTabSpecial')}
            </button>
          ))}
        </div>
        {tab === 'symbols' ? (
          <SymbolPicker
            onPick={onInsert}
            fontFamilies={fonts}
            onSelect={(char) => setInfo(char)}
            labels={{
              font: t('ribbonSymbolFont'),
              subset: t('ribbonSymbolSubset'),
              search: t('ribbonSymbolSearch'),
              codePoint: t('ribbonSymbolCodePoint'),
              recent: t('ribbonSymbolRecent'),
              insert: t('ribbonInsert'),
              empty: t('ribbonSymbolEmpty'),
              invalidCodePoint: t('ribbonSymbolInvalidCodePoint'),
              subsetNames: {
                latin: t('ribbonSymbolLatin'),
                greek: t('ribbonSymbolGreek'),
                arrows: t('ribbonSymbolArrows'),
                math: t('ribbonSymbolMath'),
                currency: t('ribbonSymbolCurrency'),
                misc: t('ribbonSymbolMisc'),
              },
            }}
          />
        ) : (
          <div className="go-symbol-picker" data-testid="special-chars">
            <div className="go-symbol-grid" role="listbox" aria-label={t('ribbonSymbolTabSpecial')}>
              {SPECIAL_CHARS.map((s) => (
                <button
                  key={s.char}
                  type="button"
                  role="option"
                  data-keep-edit
                  aria-selected={specialSel === s.char}
                  className={specialSel === s.char ? 'active' : ''}
                  title={`${s.name} (U+${symbolCodePoint(s.char)})`}
                  onClick={() => setSpecialSel(s.char)}
                  onDoubleClick={() => onInsert(s.char, 'Calibri')}
                >
                  {specialCharGlyph(s.char)}
                </button>
              ))}
            </div>
            <div className="go-symbol-bottom">
              <button
                type="button"
                className="go-symbol-insert"
                data-keep-edit
                data-testid="special-char-insert"
                onClick={() => onInsert(specialSel, 'Calibri')}
              >
                {t('ribbonInsert')}
              </button>
            </div>
          </div>
        )}
        <p className="go-symbol-info" data-testid="symbol-info" role="status">
          {shown
            ? `${t('ribbonSymbolInfoName')}: ${shownName || '-'} · ${t('ribbonSymbolInfoUnicode')}: U+${symbolCodePoint(shown)} · ${t('ribbonSymbolInfoDecimal')}: ${shown.codePointAt(0)}`
            : t('ribbonSymbolInfoNone')}
        </p>
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
        </div>
      </div>
    </div>
  )
}

// ── Insert table (explicit size beyond the 8×10 hover grid) ──────────────

const MAX_TABLE_SIDE = 50

export function TableInsertDialog({
  onInsert,
  onClose,
}: {
  onInsert: (rows: number, cols: number) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [cols, setCols] = useState(5)
  const [rows, setRows] = useState(2)

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const insert = () => onInsert(rows, cols)

  const countInput = (label: string, value: number, set: (v: number) => void, focus = false) => (
    <label>
      {label}
      <input
        type="number"
        min={1}
        max={MAX_TABLE_SIDE}
        value={value}
        autoFocus={focus}
        onChange={(e) => {
          const v = Math.round(Number(e.target.value))
          set(Number.isFinite(v) ? Math.min(MAX_TABLE_SIDE, Math.max(1, v)) : 1)
        }}
        onKeyDown={(e) => e.key === 'Enter' && insert()}
      />
    </label>
  )

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonTableInsertDialog')}</h2>
        <div className="dlg-two-col">
          {countInput(t('ribbonTableColsLabel'), cols, setCols, true)}
          {countInput(t('ribbonTableRowsLabel'), rows, setRows)}
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button className="primary" onClick={insert}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
