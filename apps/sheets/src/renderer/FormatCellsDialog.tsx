import { useEffect, useRef, useState } from 'react'

import { Dropdown } from '@genoffice/ui'

import { ColorDropdown } from './ColorDropdown'
import {
  accountingUnderlineChangeBlocked,
  accountingUnderlineUnavailable,
  degreesToOoxmlRotation,
  borderPresetEdges,
  draftFromSelection,
  formatCellsCommands,
  ooxmlRotationDegrees,
  type FormatCellsDraft,
} from './format-cells'
import { useI18n, type StringKey } from './i18n/locale'
import { FILL_PATTERN_TYPES, patternPreviewCss, type CellFillPattern } from './fill-pattern'
import {
  clampDecimals,
  CURRENCY_SYMBOLS,
  datePatterns,
  FRACTION_PATTERNS,
  NEGATIVE_STYLES,
  NUMFMT_CATEGORIES,
  numfmtOptionsOf,
  numfmtPattern,
  numfmtPreview,
  sampleValue,
  TIME_PATTERNS,
  todaySerial,
  type NumfmtCategory,
  type NumfmtOptions,
  SPECIAL_LOCALES,
} from './numfmt-dialog'

import type { SelectionFormat } from './selection-format'
import type { SelectionProtection } from './selection-protection'
import { BORDER_CONTROLS } from './border-format'
import { borderControlHasVisibleSeam, canEditSelectionBorder } from './merged-border-geometry'
import { fontFamilyGroups, useSystemFontFamilies } from './system-fonts'

/// Excel's Format Cells dialog (⌘1), scoped to what the save pipeline can
/// persist today: number format, alignment, font, border, fill, and
/// protection flags. It opens prefilled with the selection's current format
/// and emits the same command strings the ribbon uses — but only for
/// settings the user actually changed.

const TABS = ['Number', 'Alignment', 'Font', 'Border', 'Fill', 'Protection'] as const
type Tab = (typeof TABS)[number]

const TAB_LABELS: Record<Tab, StringKey> = {
  Number: 'dlgFcTabNumber',
  Alignment: 'dlgFcTabAlignment',
  Font: 'dlgFcTabFont',
  Border: 'dlgFcTabBorder',
  Fill: 'dlgFcTabFill',
  Protection: 'dlgFcTabProtection',
}

const H_ALIGNMENTS = [
  'general',
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
] as const
const V_ALIGNMENTS = ['top', 'middle', 'bottom', 'justify', 'distributed'] as const
const H_ALIGN_LABELS: Record<(typeof H_ALIGNMENTS)[number], StringKey> = {
  general: 'dlgFcAlignGeneral',
  left: 'dlgFcAlignLeft',
  center: 'dlgFcAlignCenter',
  right: 'dlgFcAlignRight',
  fill: 'dlgFcAlignFill',
  justify: 'dlgFcAlignJustify',
  centerContinuous: 'dlgFcCenterAcross',
  distributed: 'dlgFcAlignDistributed',
}
const V_ALIGN_LABELS: Record<(typeof V_ALIGNMENTS)[number], StringKey> = {
  top: 'dlgFcAlignTop',
  middle: 'dlgFcAlignMiddle',
  bottom: 'dlgFcAlignBottom',
  justify: 'dlgFcAlignJustify',
  distributed: 'dlgFcAlignDistributed',
}
const ROTATIONS: { readonly labelKey: StringKey; readonly value: string }[] = [
  { labelKey: 'dlgFcRotNone', value: '0' },
  { labelKey: 'dlgFcRotCcw', value: '45' },
  { labelKey: 'dlgFcRotCw', value: '135' },
  { labelKey: 'dlgFcRotUp', value: '90' },
  { labelKey: 'dlgFcRotDown', value: '180' },
  { labelKey: 'dlgFcRotVertical', value: 'vertical' },
]

const NUMFMT_CATEGORY_LABELS: Record<NumfmtCategory, StringKey> = {
  general: 'dlgFcNumGeneral',
  number: 'dlgFcNumNumber',
  currency: 'dlgFcNumCurrency',
  accounting: 'appNumFmtAccounting',
  date: 'dlgFcNumDate',
  time: 'dlgFcNumTime',
  percentage: 'dlgFcNumPercent',
  fraction: 'appNumFmtFraction',
  scientific: 'dlgFcNumScientific',
  text: 'dlgFcNumText',
  special: 'dlgFcNumSpecial',
  custom: 'dlgFcNumCustomName',
}

const PATTERN_LABELS: Record<CellFillPattern['pattern'], StringKey> = {
  darkHorizontal: 'dlgFcPatDarkHorizontal',
  darkVertical: 'dlgFcPatDarkVertical',
  darkDown: 'dlgFcPatDarkDown',
  darkUp: 'dlgFcPatDarkUp',
  darkGrid: 'dlgFcPatDarkGrid',
  darkTrellis: 'dlgFcPatDarkTrellis',
  lightHorizontal: 'dlgFcPatLightHorizontal',
  lightVertical: 'dlgFcPatLightVertical',
  lightDown: 'dlgFcPatLightDown',
  lightUp: 'dlgFcPatLightUp',
  lightGrid: 'dlgFcPatLightGrid',
  lightTrellis: 'dlgFcPatLightTrellis',
  gray125: 'dlgFcPat12',
  gray0625: 'dlgFcPat6',
  lightGray: 'dlgFcPat25',
  mediumGray: 'dlgFcPat50',
  darkGray: 'dlgFcPat75',
}

const FRACTION_LABELS: readonly StringKey[] = [
  'dlgFcFracD1',
  'dlgFcFracD2',
  'dlgFcFracD3',
  'dlgFcFracHalf',
  'dlgFcFracQuarter',
  'dlgFcFracEighth',
  'dlgFcFracSixteenth',
  'dlgFcFracTenth',
]

const DECIMAL_CATEGORIES: readonly NumfmtCategory[] = [
  'number',
  'currency',
  'accounting',
  'percentage',
  'scientific',
]

const FONT_SIZES = ['9', '10', '11', '12', '14', '16', '18', '22', '26']
/// OOXML ST_BorderStyle names offered by the line-style picker; the
/// journal maps them 1:1 to Univer BorderStyleTypes and back to the file.
const BORDER_LINE_STYLES: { readonly labelKey: StringKey; readonly value: string }[] = [
  { labelKey: 'dlgFcStyleThin', value: 'thin' },
  { labelKey: 'dlgFcStyleMedium', value: 'medium' },
  { labelKey: 'dlgFcStyleThick', value: 'thick' },
  { labelKey: 'dlgFcStyleDouble', value: 'double' },
  { labelKey: 'dlgFcStyleHair', value: 'hair' },
  { labelKey: 'dlgFcStyleDashed', value: 'dashed' },
  { labelKey: 'dlgFcStyleDotted', value: 'dotted' },
  { labelKey: 'dlgFcStyleDashDot', value: 'dashDot' },
  { labelKey: 'dlgFcStyleDashDotDot', value: 'dashDotDot' },
  { labelKey: 'dlgFcStyleMediumDashed', value: 'mediumDashed' },
  { labelKey: 'dlgFcStyleMediumDashDot', value: 'mediumDashDot' },
  { labelKey: 'dlgFcStyleMediumDashDotDot', value: 'mediumDashDotDot' },
  { labelKey: 'dlgFcStyleSlantDashDot', value: 'slantDashDot' },
]

const BORDER_PRESETS: { readonly labelKey: StringKey; readonly value: string }[] = [
  { labelKey: 'dlgFcBorderNone', value: 'none' },
  { labelKey: 'dlgFcBorderAll', value: 'all' },
  { labelKey: 'dlgFcBorderInside', value: 'inside' },
  { labelKey: 'dlgFcBorderOutline', value: 'outer' },
  { labelKey: 'dlgFcBorderThickOutline', value: 'thick-outer' },
  { labelKey: 'dlgFcBorderTop', value: 'top' },
  { labelKey: 'dlgFcBorderBottom', value: 'bottom' },
  { labelKey: 'dlgFcBorderLeft', value: 'left' },
  { labelKey: 'dlgFcBorderRight', value: 'right' },
]

export function FormatCellsDialog({
  selectionFormat,
  selectionProtection,
  anchorValue,
  onCommand,
  onClose,
}: {
  readonly selectionFormat: SelectionFormat | null
  readonly selectionProtection: SelectionProtection | null
  /// Value of the selection's top-left cell, for the number-format preview.
  readonly anchorValue: number | string | null
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [tab, setTab] = useState<Tab>('Number')
  useEffect(() => {
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])
  // The selection format at open time is the change baseline; the live echo
  // must not move it while the dialog is up.
  const initialRef = useRef(draftFromSelection(selectionFormat, selectionProtection))
  const initial = initialRef.current
  const [draft, setDraft] = useState(initial)
  const accountingUnavailable = accountingUnderlineUnavailable(selectionProtection, draft.rotation)
  const accountingChangeBlocked = accountingUnderlineChangeBlocked(
    initial,
    draft,
    selectionProtection,
  )
  const borderEditable = canEditSelectionBorder(selectionProtection)
  const hasHorizontalSeam =
    selectionProtection !== null &&
    borderEditable &&
    borderControlHasVisibleSeam(selectionProtection.range, selectionProtection.merges ?? [], 'down')
  const hasVerticalSeam =
    selectionProtection !== null &&
    borderEditable &&
    borderControlHasVisibleSeam(
      selectionProtection.range,
      selectionProtection.merges ?? [],
      'right',
    )
  const set = <K extends keyof FormatCellsDraft>(key: K, value: FormatCellsDraft[K]): void =>
    setDraft((previous) => ({ ...previous, [key]: value }))

  const [numOptions, setNumOptions] = useState<NumfmtOptions>(() =>
    numfmtOptionsOf(initial.pattern),
  )
  const updateNumfmt = (patch: Partial<NumfmtOptions>): void => {
    const next = { ...numOptions, ...patch }
    setNumOptions(next)
    set('pattern', numfmtPattern(next))
  }
  // Freeze the preview inputs at open time so the sample doesn't tick.
  const previewBase = useRef({ anchor: anchorValue, serial: todaySerial() })
  const sample = sampleValue(
    numOptions.category,
    previewBase.current.anchor,
    previewBase.current.serial,
  )
  const negativeSample = -Math.abs(typeof sample === 'number' ? sample : 1234.56)

  const { families: systemFontFamilies, load: loadSystemFonts } = useSystemFontFamilies()
  // the dialog opens from a click, so activation is still live here
  useEffect(() => loadSystemFonts(), [loadSystemFonts])
  const fontGroups = fontFamilyGroups(systemFontFamilies, draft.family)
  const sizeOptions =
    !draft.size || FONT_SIZES.includes(draft.size)
      ? FONT_SIZES
      : [...FONT_SIZES, draft.size].sort((a, b) => Number(a) - Number(b))

  function handleApply(): void {
    if (accountingChangeBlocked) return
    for (const command of formatCellsCommands(initial, draft)) onCommand(command)
    onClose()
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog"
        role="dialog"
        aria-label={t('dlgFcTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgFcTitle')}</header>
        <nav className="dialog-tabs">
          {TABS.map((candidate) => (
            <button
              key={candidate}
              className={candidate === tab ? 'active' : ''}
              onClick={() => setTab(candidate)}
            >
              {t(TAB_LABELS[candidate])}
            </button>
          ))}
        </nav>
        <section className="dialog-body">
          {tab === 'Number' && (
            <div className="numfmt-tab">
              <div className="numfmt-cats" role="listbox" aria-label={t('dlgFcCategory')}>
                {NUMFMT_CATEGORIES.map((category) => (
                  <button
                    key={category}
                    role="option"
                    aria-selected={category === numOptions.category}
                    className={category === numOptions.category ? 'active' : ''}
                    onClick={() =>
                      updateNumfmt(
                        category === 'custom' ? { category, custom: draft.pattern } : { category },
                      )
                    }
                  >
                    {t(NUMFMT_CATEGORY_LABELS[category])}
                  </button>
                ))}
              </div>
              <div className="numfmt-body">
                <div className="numfmt-field numfmt-sample">
                  {t('dlgFcSample')}
                  <output>{numfmtPreview(draft.pattern, sample)}</output>
                </div>
                {DECIMAL_CATEGORIES.includes(numOptions.category) && (
                  <div className="numfmt-row">
                    <label className="numfmt-field">
                      {t('dlgFcDecimals')}
                      <input
                        type="number"
                        min={0}
                        max={30}
                        value={numOptions.decimals}
                        onChange={(e) =>
                          updateNumfmt({ decimals: clampDecimals(Number(e.target.value)) })
                        }
                      />
                    </label>
                    {(numOptions.category === 'currency' ||
                      numOptions.category === 'accounting') && (
                      <label className="numfmt-field">
                        {t('dlgFcCurrencySymbol')}
                        <Dropdown
                          ariaLabel={t('dlgFcCurrencySymbol')}
                          value={numOptions.symbol}
                          options={[
                            { value: '', label: t('dlgFcSymbolNone') },
                            ...CURRENCY_SYMBOLS.map((symbol) => ({
                              value: symbol,
                              label: symbol,
                            })),
                          ]}
                          onPick={(v) => updateNumfmt({ symbol: v })}
                        />
                      </label>
                    )}
                  </div>
                )}
                {numOptions.category === 'number' && (
                  <label className="dialog-check">
                    <input
                      type="checkbox"
                      checked={numOptions.thousands}
                      onChange={(e) => updateNumfmt({ thousands: e.target.checked })}
                    />
                    {t('dlgFcThousandsSep')}
                  </label>
                )}
                {(numOptions.category === 'number' || numOptions.category === 'currency') && (
                  <div className="numfmt-field">
                    {t('dlgFcNegNumbers')}
                    <div className="numfmt-list" role="listbox" aria-label={t('dlgFcNegNumbers')}>
                      {NEGATIVE_STYLES.map((style) => (
                        <button
                          key={style}
                          role="option"
                          aria-selected={style === numOptions.negative}
                          className={`${style === numOptions.negative ? 'active' : ''} ${
                            style.includes('red') ? 'red' : ''
                          }`.trim()}
                          onClick={() => updateNumfmt({ negative: style })}
                        >
                          {numfmtPreview(
                            numfmtPattern({ ...numOptions, negative: style }),
                            negativeSample,
                          )}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {(numOptions.category === 'date' || numOptions.category === 'time') && (
                  <div className="numfmt-field">
                    {t('dlgFcTypeLabel')}
                    <div className="numfmt-list" role="listbox" aria-label={t('dlgFcTypeLabel')}>
                      {(numOptions.category === 'date' ? datePatterns() : TIME_PATTERNS).map(
                        (candidate) => {
                          const key = numOptions.category === 'date' ? 'datePattern' : 'timePattern'
                          return (
                            <button
                              key={candidate}
                              role="option"
                              aria-selected={candidate === numOptions[key]}
                              className={candidate === numOptions[key] ? 'active' : ''}
                              onClick={() => updateNumfmt({ [key]: candidate })}
                            >
                              {numfmtPreview(candidate, sample)}
                            </button>
                          )
                        },
                      )}
                    </div>
                  </div>
                )}
                {numOptions.category === 'fraction' && (
                  <div className="numfmt-field">
                    {t('dlgFcTypeLabel')}
                    <div className="numfmt-list" role="listbox" aria-label={t('dlgFcTypeLabel')}>
                      {FRACTION_PATTERNS.map((candidate, index) => (
                        <button
                          key={candidate}
                          role="option"
                          aria-selected={candidate === numOptions.fractionPattern}
                          className={candidate === numOptions.fractionPattern ? 'active' : ''}
                          onClick={() => updateNumfmt({ fractionPattern: candidate })}
                        >
                          {t(FRACTION_LABELS[index] ?? 'dlgFcFracD1')}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {numOptions.category === 'special' && (
                  <>
                    <label className="numfmt-field">
                      {t('dlgFcSpLocale')}
                      <select
                        value={numOptions.specialLocale}
                        onChange={(event) => {
                          const locale =
                            SPECIAL_LOCALES.find((entry) => entry.id === event.target.value) ??
                            SPECIAL_LOCALES[0]!
                          updateNumfmt({
                            specialLocale: locale.id,
                            special: locale.formats[0]!.pattern,
                          })
                        }}
                      >
                        {SPECIAL_LOCALES.map((locale) => (
                          <option key={locale.id} value={locale.id}>
                            {locale.name}
                          </option>
                        ))}
                      </select>
                    </label>
                    <div className="numfmt-field">
                      {t('dlgFcTypeLabel')}
                      <div className="numfmt-list" role="listbox" aria-label={t('dlgFcTypeLabel')}>
                        {(
                          SPECIAL_LOCALES.find((entry) => entry.id === numOptions.specialLocale) ??
                          SPECIAL_LOCALES[0]!
                        ).formats.map((candidate) => (
                          <button
                            key={candidate.pattern}
                            role="option"
                            aria-selected={candidate.pattern === numOptions.special}
                            className={candidate.pattern === numOptions.special ? 'active' : ''}
                            onClick={() => updateNumfmt({ special: candidate.pattern })}
                          >
                            {t(candidate.label as StringKey)}
                          </button>
                        ))}
                      </div>
                    </div>
                  </>
                )}
                {numOptions.category === 'general' && (
                  <p className="numfmt-note">{t('dlgFcGeneralNote')}</p>
                )}
                {numOptions.category === 'text' && (
                  <p className="numfmt-note">{t('dlgFcTextNote')}</p>
                )}
                {numOptions.category === 'custom' && (
                  <label className="numfmt-field">
                    {t('dlgFcCustomCode')}
                    <input
                      value={draft.pattern}
                      placeholder="#,##0.00"
                      onChange={(e) => {
                        setNumOptions({ ...numOptions, custom: e.target.value })
                        set('pattern', e.target.value)
                      }}
                    />
                  </label>
                )}
              </div>
            </div>
          )}
          {tab === 'Alignment' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcHorizontal')}
                {selectionProtection?.alignment.horizontal === null && (
                  <small> · {t('dlgFcMixed')}</small>
                )}
                <Dropdown
                  ariaLabel={t('dlgFcHorizontal')}
                  value={draft.hAlign}
                  disabled={selectionProtection === null}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    ...H_ALIGNMENTS.map((h) => ({ value: h, label: t(H_ALIGN_LABELS[h]) })),
                  ]}
                  onPick={(v) => set('hAlign', v)}
                />
              </label>
              <label>
                {t('dlgFcVertical')}
                {selectionProtection?.alignment.vertical === null && (
                  <small> · {t('dlgFcMixed')}</small>
                )}
                <Dropdown
                  ariaLabel={t('dlgFcVertical')}
                  value={draft.vAlign}
                  disabled={selectionProtection === null}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    ...V_ALIGNMENTS.map((v) => ({ value: v, label: t(V_ALIGN_LABELS[v]) })),
                  ]}
                  onPick={(v) => set('vAlign', v)}
                />
              </label>
              <label>
                {t('dlgFcWrapText')}
                <Dropdown
                  ariaLabel={t('dlgFcWrapText')}
                  value={draft.wrapText}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    { value: 'on', label: t('dlgFcOn') },
                    { value: 'off', label: t('dlgFcOff') },
                  ]}
                  onPick={(v) => set('wrapText', v as FormatCellsDraft['wrapText'])}
                />
              </label>
              <label>
                {t('dlgFcOrientation')}
                {selectionProtection?.alignment.rotation === null && (
                  <small> · {t('dlgFcMixed')}</small>
                )}
                <Dropdown
                  ariaLabel={t('dlgFcOrientation')}
                  value={draft.rotation}
                  disabled={selectionProtection === null}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    ...(draft.rotation && !ROTATIONS.some((item) => item.value === draft.rotation)
                      ? [
                          {
                            value: draft.rotation,
                            label: `${ooxmlRotationDegrees(draft.rotation)}°`,
                          },
                        ]
                      : []),
                    ...ROTATIONS.map((r) => ({ value: r.value, label: t(r.labelKey) })),
                  ]}
                  onPick={(v) => set('rotation', v)}
                />
              </label>
              <label>
                {t('dlgFcDegrees')}
                <input
                  type="number"
                  disabled={selectionProtection === null}
                  min="-90"
                  max="90"
                  step="1"
                  value={ooxmlRotationDegrees(draft.rotation)}
                  placeholder={t('dlgFcUnchanged')}
                  onChange={(event) => set('rotation', degreesToOoxmlRotation(event.target.value))}
                />
              </label>
              <label>
                {t('dlgFcShrinkToFit')}
                {selectionProtection?.alignment.shrinkToFit === null && (
                  <small> · {t('dlgFcMixed')}</small>
                )}
                <Dropdown
                  ariaLabel={t('dlgFcShrinkToFit')}
                  value={draft.shrinkToFit}
                  disabled={selectionProtection === null}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    { value: 'on', label: t('dlgFcOn') },
                    { value: 'off', label: t('dlgFcOff') },
                  ]}
                  onPick={(value) => set('shrinkToFit', value as FormatCellsDraft['shrinkToFit'])}
                />
              </label>
              <label>
                {t('appMergeCells')}
                {selectionProtection?.alignment.merged === null && (
                  <small> · {t('dlgFcMixed')}</small>
                )}
                <Dropdown
                  ariaLabel={t('appMergeCells')}
                  value={draft.mergeCells}
                  disabled={
                    selectionProtection === null ||
                    (selectionProtection.range.startRow === selectionProtection.range.endRow &&
                      selectionProtection.range.startColumn ===
                        selectionProtection.range.endColumn &&
                      selectionProtection.alignment.merged === false)
                  }
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    { value: 'on', label: t('dlgFcOn') },
                    { value: 'off', label: t('dlgFcOff') },
                  ]}
                  onPick={(value) => set('mergeCells', value as FormatCellsDraft['mergeCells'])}
                />
              </label>
              {selectionProtection === null && (
                <p className="dialog-note" role="status">
                  {t('dlgFcAlignmentUnavailable')}
                </p>
              )}
              <p className="dialog-note">{t('dlgFcAlignmentPreviewNote')}</p>
              <label>
                {t('dlgFcIndent')}
                <input
                  type="number"
                  min="0"
                  max="250"
                  value={draft.indent}
                  placeholder={t('dlgFcUnchanged')}
                  onChange={(e) => set('indent', e.target.value)}
                />
              </label>
            </div>
          )}
          {tab === 'Font' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcFont')}
                {/* optgroups flattened in order: common families, then system ones
                    (deduped — the echoed family may appear in both groups) */}
                <Dropdown
                  ariaLabel={t('dlgFcFont')}
                  value={draft.family}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    ...[
                      ...fontGroups.common,
                      ...fontGroups.system.filter((f) => !fontGroups.common.includes(f)),
                    ].map((f) => ({ value: f, label: f })),
                  ]}
                  onPick={(v) => set('family', v)}
                />
              </label>
              <label>
                {t('dlgFcSize')}
                <Dropdown
                  ariaLabel={t('dlgFcSize')}
                  value={draft.size}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    ...sizeOptions.map((s) => ({ value: s, label: s })),
                  ]}
                  onPick={(v) => set('size', v)}
                />
              </label>
              {(
                [
                  ['dlgFcBold', 'bold'],
                  ['dlgFcItalic', 'italic'],
                  ['dlgFcStrikethrough', 'strike'],
                ] as const
              ).map(([labelKey, key]) => (
                <label key={key}>
                  {t(labelKey)}
                  <Dropdown
                    ariaLabel={t(labelKey)}
                    value={draft[key]}
                    options={[
                      { value: '', label: t('dlgFcUnchanged') },
                      { value: 'on', label: t('dlgFcOn') },
                      { value: 'off', label: t('dlgFcOff') },
                    ]}
                    onPick={(v) => set(key, v as FormatCellsDraft[typeof key])}
                  />
                </label>
              ))}
              <label>
                {t('dlgFcUnderline')}
                <Dropdown
                  ariaLabel={t('dlgFcUnderline')}
                  value={draft.underlineMode}
                  disabled={selectionProtection === null}
                  options={[
                    {
                      value: '',
                      label: t(
                        selectionProtection?.font?.underlineStyle === null
                          ? 'dlgFcMixed'
                          : 'dlgFcUnchanged',
                      ),
                    },
                    { value: 'none', label: t('dlgFcUnderlineNone') },
                    { value: 'single', label: t('dlgFcUnderlineSingle') },
                    { value: 'double', label: t('dlgFcUnderlineDouble') },
                    {
                      value: 'singleAccounting',
                      label: t('dlgFcUnderlineSingleAccounting'),
                      disabled: accountingUnavailable,
                    },
                    {
                      value: 'doubleAccounting',
                      label: t('dlgFcUnderlineDoubleAccounting'),
                      disabled: accountingUnavailable,
                    },
                  ]}
                  onPick={(value) =>
                    set('underlineMode', value as FormatCellsDraft['underlineMode'])
                  }
                />
              </label>
              {accountingUnavailable && selectionProtection !== null && (
                <p className="dialog-note" role="status">
                  {t('dlgFcAccountingUnavailable')}
                </p>
              )}
              <label>
                {t('dlgFcFontPosition')}
                <Dropdown
                  ariaLabel={t('dlgFcFontPosition')}
                  value={draft.fontVerticalAlign}
                  disabled={selectionProtection === null}
                  options={[
                    {
                      value: '',
                      label: t(
                        selectionProtection?.font?.verticalAlign === null
                          ? 'dlgFcMixed'
                          : 'dlgFcUnchanged',
                      ),
                    },
                    { value: 'baseline', label: t('dlgFcFontBaseline') },
                    { value: 'superscript', label: t('dlgFcFontSuperscript') },
                    { value: 'subscript', label: t('dlgFcFontSubscript') },
                  ]}
                  onPick={(value) =>
                    set('fontVerticalAlign', value as FormatCellsDraft['fontVerticalAlign'])
                  }
                />
              </label>
              <label>
                {t('dlgFcColor')}
                <ColorDropdown
                  label={t('dlgFcColor')}
                  value={draft.fontColor || '#000000'}
                  onPick={(hex) => {
                    if (hex) set('fontColor', hex)
                  }}
                />
              </label>
            </div>
          )}
          {tab === 'Border' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcBorderPresets')}
                <Dropdown
                  ariaLabel={t('dlgFcBorderPresets')}
                  value={draft.border}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    ...BORDER_PRESETS.map((preset) => ({
                      value: preset.value,
                      label: t(preset.labelKey),
                      disabled: preset.value === 'inside' && !hasHorizontalSeam && !hasVerticalSeam,
                    })),
                  ]}
                  disabled={!borderEditable}
                  onPick={(v) =>
                    setDraft((previous) => ({
                      ...previous,
                      border: v,
                      borderStyle: v === 'thick-outer' ? 'thick' : previous.borderStyle,
                      borderEdges: borderPresetEdges(previous.borderEdges, v),
                    }))
                  }
                />
              </label>
              <label>
                {t('dlgFcBorderStyle')}
                {selectionProtection?.border?.lineStyle === null && (
                  <small> · {t('dlgFcMixed')}</small>
                )}
                <Dropdown
                  ariaLabel={t('dlgFcBorderStyle')}
                  value={draft.borderStyle}
                  options={BORDER_LINE_STYLES.map((style) => ({
                    value: style.value,
                    label: t(style.labelKey),
                  }))}
                  onPick={(v) => set('borderStyle', v)}
                />
              </label>
              <label>
                {t('dlgFcColor')}
                {selectionProtection?.border?.lineColor === null && (
                  <small> · {t('dlgFcMixed')}</small>
                )}
                <ColorDropdown
                  label={t('dlgFcColor')}
                  value={draft.borderColor}
                  onPick={(hex) => {
                    if (hex) set('borderColor', hex)
                  }}
                />
              </label>
              {BORDER_CONTROLS.map((control) => {
                const key = {
                  left: 'dlgFcBorderLeft',
                  top: 'dlgFcBorderTop',
                  bottom: 'dlgFcBorderBottom',
                  right: 'dlgFcBorderRight',
                  insideHorizontal: 'dlgFcBorderInsideHorizontal',
                  insideVertical: 'dlgFcBorderInsideVertical',
                  diagonalUp: 'dlgFcBorderDiagonalUp',
                  diagonalDown: 'dlgFcBorderDiagonalDown',
                }[control] as StringKey
                const isInside = control === 'insideHorizontal' || control === 'insideVertical'
                const cannotApply =
                  !borderEditable ||
                  (isInside &&
                    (control === 'insideHorizontal' ? !hasHorizontalSeam : !hasVerticalSeam))
                return (
                  <label key={control}>
                    {t(key)}
                    <Dropdown
                      ariaLabel={t(key)}
                      value={draft.borderEdges[control]}
                      disabled={cannotApply}
                      options={[
                        {
                          value: '',
                          label: t(
                            selectionProtection?.border?.[control] === null
                              ? 'dlgFcMixed'
                              : 'dlgFcUnchanged',
                          ),
                        },
                        { value: 'on', label: t('dlgFcOn') },
                        { value: 'off', label: t('dlgFcOff') },
                      ]}
                      onPick={(value) =>
                        setDraft((previous) => ({
                          ...previous,
                          border: '',
                          borderEdges: { ...previous.borderEdges, [control]: value },
                        }))
                      }
                    />
                  </label>
                )
              })}
              {!borderEditable && (
                <p className="format-cells-note">{t('dlgFcBorderUnavailable')}</p>
              )}
            </div>
          )}
          {tab === 'Fill' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcBackground')}
                <ColorDropdown
                  label={t('dlgFcBackground')}
                  value={draft.fill || '#ffffff'}
                  disabled={draft.noFill}
                  onPick={(hex) => {
                    if (hex) set('fill', hex)
                  }}
                />
              </label>
              <label className="dialog-check">
                <input
                  type="checkbox"
                  checked={draft.noFill}
                  onChange={(e) => set('noFill', e.target.checked)}
                />
                {t('dlgFcNoFill')}
              </label>
              <label>
                {t('dlgFcPatternColor')}
                <ColorDropdown
                  label={t('dlgFcPatternColor')}
                  value={draft.patternColor || '#000000'}
                  disabled={draft.noFill}
                  onPick={(hex) => {
                    if (hex) set('patternColor', hex)
                  }}
                />
              </label>
              <div className="fill-pattern-field">
                {t('dlgFcPatternStyle')}
                <div
                  className="fill-pattern-grid"
                  role="listbox"
                  aria-label={t('dlgFcPatternStyle')}
                >
                  <button
                    type="button"
                    role="option"
                    aria-selected={draft.fillPattern === ''}
                    className={draft.fillPattern === '' ? 'active' : ''}
                    title={t('dlgFcPatternNone')}
                    disabled={draft.noFill}
                    onClick={() => set('fillPattern', '')}
                  >
                    {t('dlgFcPatternNone')}
                  </button>
                  {FILL_PATTERN_TYPES.map((pattern) => (
                    <button
                      type="button"
                      key={pattern}
                      role="option"
                      aria-selected={draft.fillPattern === pattern}
                      className={draft.fillPattern === pattern ? 'active' : ''}
                      title={t(PATTERN_LABELS[pattern])}
                      aria-label={t(PATTERN_LABELS[pattern])}
                      disabled={draft.noFill}
                      onClick={() => set('fillPattern', pattern)}
                    >
                      <i
                        style={patternPreviewCss({
                          pattern,
                          fg: draft.patternColor || '#000000',
                          bg: draft.fill || '#ffffff',
                        })}
                      />
                    </button>
                  ))}
                </div>
              </div>
              <div className="fill-pattern-sample" aria-label={t('dlgFcSample')}>
                <i
                  style={
                    draft.noFill
                      ? undefined
                      : draft.fillPattern
                        ? patternPreviewCss({
                            pattern: draft.fillPattern as CellFillPattern['pattern'],
                            fg: draft.patternColor || '#000000',
                            bg: draft.fill || '#ffffff',
                          })
                        : { backgroundColor: draft.fill || 'transparent' }
                  }
                />
              </div>
            </div>
          )}
          {tab === 'Protection' && (
            <div className="dialog-grid">
              <label>
                {t('dlgFcLocked')}
                {selectionProtection?.locked === null && <small> · {t('dlgFcMixed')}</small>}
                <Dropdown
                  ariaLabel={t('dlgFcLocked')}
                  value={draft.locked}
                  disabled={selectionProtection === null}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    { value: 'on', label: t('dlgFcOn') },
                    { value: 'off', label: t('dlgFcOff') },
                  ]}
                  onPick={(v) => set('locked', v as FormatCellsDraft['locked'])}
                />
              </label>
              <label>
                {t('dlgFcHidden')}
                {selectionProtection?.hidden === null && <small> · {t('dlgFcMixed')}</small>}
                <Dropdown
                  ariaLabel={t('dlgFcHidden')}
                  value={draft.hidden}
                  disabled={selectionProtection === null}
                  options={[
                    { value: '', label: t('dlgFcUnchanged') },
                    { value: 'on', label: t('dlgFcOn') },
                    { value: 'off', label: t('dlgFcOff') },
                  ]}
                  onPick={(v) => set('hidden', v as FormatCellsDraft['hidden'])}
                />
              </label>
              {selectionProtection === null && (
                <p className="dialog-note" role="status">
                  {t('dlgFcProtectionUnavailable')}
                </p>
              )}
              <p className="dialog-note">{t('dlgFcProtectionNote')}</p>
            </div>
          )}
        </section>
        <footer className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary-action"
            onClick={handleApply}
            disabled={accountingChangeBlocked}
          >
            {t('dlgOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
