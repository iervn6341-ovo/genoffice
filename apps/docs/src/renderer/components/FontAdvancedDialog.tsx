import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { Dropdown, isSymbolFontFamily, type DropdownOption } from '@genoffice/ui'
import { useI18n, type StringKey } from '../i18n/locale'
import { fontFamiliesFor } from '../font-list'
import { useSystemFontFamilies } from '../system-fonts'
import { cssFontFamily } from '../line-metrics'
import {
  applyFontChanges,
  commonFontValue,
  enabledStylisticSets,
  fontSelectionSamples,
  word2010PropertyPresent,
  word2010PropertyValue,
  type FontChanges,
  type FontSample,
} from '../editor/font-advanced'
import { useModalKeys } from './modal-keys'

const FONT_SIZES = [9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72]
const FONT_STYLES: Array<{ key: string; nameKey: StringKey }> = [
  { key: 'regular', nameKey: 'appFontRegular' },
  { key: 'italic', nameKey: 'appFontItalic' },
  { key: 'bold', nameKey: 'appFontBold' },
  { key: 'boldItalic', nameKey: 'appFontBoldItalic' },
]
const UNDERLINES: Array<{ key: string; label: StringKey }> = [
  { key: 'none', label: 'appFontNone' },
  { key: 'single', label: 'appFontUnderlineSingle' },
  { key: 'double', label: 'appFontUnderlineDouble' },
  { key: 'dotted', label: 'appFontUnderlineDotted' },
  { key: 'dash', label: 'appFontUnderlineDashed' },
  { key: 'wave', label: 'appFontUnderlineWavy' },
]
const EFFECTS: Array<{ key: NonNullable<FontChanges['textEffect']>; label: StringKey }> = [
  { key: 'none', label: 'appFontNone' },
  { key: 'outline', label: 'appFontEffectOutline' },
  { key: 'shadow', label: 'appFontEffectShadow' },
  { key: 'emboss', label: 'appFontEffectEmboss' },
  { key: 'imprint', label: 'appFontEffectImprint' },
]
const EMPHASIS: Array<{ key: NonNullable<FontChanges['em']>; label: StringKey }> = [
  { key: 'none', label: 'appFontNone' },
  { key: 'dot', label: 'appFontEmphasisDot' },
  { key: 'comma', label: 'appFontEmphasisComma' },
  { key: 'circle', label: 'appFontEmphasisCircle' },
  { key: 'underDot', label: 'appFontEmphasisUnderDot' },
]
const LIGATURES: Array<{ key: NonNullable<FontChanges['ligatures']>; label: StringKey }> = [
  { key: 'none', label: 'appFontNone' },
  { key: 'standard', label: 'appFontLigaturesStandard' },
  { key: 'standardContextual', label: 'appFontLigaturesStandardContextual' },
  { key: 'historicalDiscretionary', label: 'appFontLigaturesHistorical' },
  { key: 'all', label: 'appFontAll' },
]
const NUMBER_SPACING: Array<{ key: NonNullable<FontChanges['numberSpacing']>; label: StringKey }> =
  [
    { key: 'default', label: 'appFontDefault' },
    { key: 'proportional', label: 'appFontNumberProportional' },
    { key: 'tabular', label: 'appFontNumberTabular' },
  ]
const NUMBER_FORM: Array<{ key: NonNullable<FontChanges['numberForm']>; label: StringKey }> = [
  { key: 'default', label: 'appFontDefault' },
  { key: 'lining', label: 'appFontNumberLining' },
  { key: 'oldStyle', label: 'appFontNumberOldStyle' },
]

const styleKey = (bold: boolean, italic: boolean) =>
  bold ? (italic ? 'boldItalic' : 'bold') : italic ? 'italic' : 'regular'
const colorHex = (value: string) => (/^[0-9A-Fa-f]{6}$/.test(value) ? value : '000000')

function common<T>(samples: FontSample[], valueOf: (sample: FontSample) => T) {
  return commonFontValue(samples, valueOf)
}

function Check({
  label,
  value,
  mixed,
  onChange,
}: {
  label: string
  value: boolean
  mixed?: boolean
  onChange: (value: boolean) => void
}) {
  return (
    <label className="font-check">
      <input
        type="checkbox"
        checked={value}
        ref={(el) => {
          if (el) el.indeterminate = !!mixed
        }}
        onChange={(e) => onChange(e.target.checked)}
      />
      {label}
    </label>
  )
}

export function FontDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t, lang } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const fontFamilies = fontFamiliesFor(lang)
  const { families: systemFontFamilies, load: loadSystemFonts } = useSystemFontFamilies()
  useEffect(() => loadSystemFonts(), [loadSystemFonts])
  const [selection] = useState(() => ({
    from: editor.state.selection.from,
    to: editor.state.selection.to,
  }))
  const [samples] = useState(() => fontSelectionSamples(editor, selection.from, selection.to))
  const [tab, setTab] = useState<'font' | 'advanced'>('font')
  const [changes, setChanges] = useState<FontChanges>({})
  const change = <K extends keyof FontChanges>(key: K, value: FontChanges[K]) =>
    setChanges((prior) => ({ ...prior, [key]: value }))
  const latin = common(samples, (s) => String(s.attrs.fontAscii ?? ''))
  const eastAsian = common(samples, (s) => String(s.attrs.font ?? ''))
  const size = common(samples, (s) => Number(s.attrs.sizeHalfPoints ?? 0))
  const color = common(samples, (s) => String(s.attrs.color ?? ''))
  const style = common(samples, (s) => styleKey(s.marks.has('bold'), s.marks.has('italic')))
  const underline = common(samples, (s) =>
    String(s.attrs.underlineStyle ?? (s.marks.has('underline') ? 'single' : 'none')),
  )
  const underlineColor = common(samples, (s) => String(s.attrs.underlineColor ?? ''))
  const strike = common(samples, (s) => s.marks.has('strike'))
  const dstrike = common(samples, (s) => s.attrs.dstrike === true)
  const caps = common(samples, (s) => String(s.attrs.caps ?? 'none'))
  const vanish = common(samples, (s) => s.attrs.vanish === true)
  const vertAlign = common(samples, (s) => String(s.attrs.vertAlign ?? ''))
  const effect = common(samples, (s) => String(s.attrs.textEffect ?? 'none'))
  const emphasis = common(samples, (s) => String(s.attrs.em ?? 'none'))
  const scale = common(samples, (s) => Number(s.attrs.charScalePct ?? 100))
  const spacing = common(samples, (s) => Number(s.attrs.charSpacingTwips ?? 0))
  const position = common(samples, (s) => Number(s.attrs.positionHalfPoints ?? 0))
  const kern = common(samples, (s) => Number(s.attrs.kernHalfPoints ?? 0))
  const snap = common(
    samples,
    (s) => !/<w:snapToGrid\b[^>]*w:val="(?:0|false|off)"/.test(String(s.attrs.rawRPr ?? '')),
  )
  const rawValue = (s: FontSample, localName: string) =>
    word2010PropertyValue(
      String(s.attrs.rawRPr ?? ''),
      localName,
      (s.attrs.rawRPrW14Prefixes as string[] | null) ?? undefined,
    )
  const ligatures = common(samples, (s) => rawValue(s, 'ligatures') ?? 'none')
  const numberSpacing = common(samples, (s) => rawValue(s, 'numSpacing') ?? 'default')
  const numberForm = common(samples, (s) => rawValue(s, 'numForm') ?? 'default')
  const contextualAlternates = common(samples, (s) => {
    const raw = String(s.attrs.rawRPr ?? '')
    const prefixes = (s.attrs.rawRPrW14Prefixes as string[] | null) ?? undefined
    return (
      word2010PropertyPresent(raw, 'cntxtAlts', prefixes) &&
      !/^(?:0|false|off)$/.test(word2010PropertyValue(raw, 'cntxtAlts', prefixes) ?? '')
    )
  })
  const stylisticSet = common(samples, (s) => {
    const enabled = enabledStylisticSets(
      String(s.attrs.rawRPr ?? ''),
      (s.attrs.rawRPrW14Prefixes as string[] | null) ?? undefined,
    )
    return enabled === null || enabled.length > 1 ? -1 : (enabled[0] ?? 0)
  })
  const fontOptions = (current: string): DropdownOption[] => [
    { value: '', label: t('appFontMixedOrDefault') },
    ...fontFamilies.map((f) => ({
      value: f,
      label: f,
      render: <span style={{ fontFamily: cssFontFamily(f) }}>{f}</span>,
    })),
    ...systemFontFamilies.map((f) => ({
      value: f,
      label: f,
      render: (
        <span style={{ fontFamily: isSymbolFontFamily(f) ? undefined : cssFontFamily(f) }}>
          {f}
        </span>
      ),
    })),
    ...(current && !fontFamilies.includes(current) && !systemFontFamilies.includes(current)
      ? [{ value: current, label: current }]
      : []),
  ]
  const latinValue = changes.fontAscii ?? (latin.mixed ? '' : (latin.value ?? ''))
  const eastValue = changes.font ?? (eastAsian.mixed ? '' : (eastAsian.value ?? ''))
  const sizeValue = changes.sizeHalfPoints ?? (size.mixed ? 0 : (size.value ?? 0))
  const styleValue =
    changes.bold !== undefined || changes.italic !== undefined
      ? styleKey(
          changes.bold ?? style.value?.includes('bold') ?? false,
          changes.italic ?? style.value?.toLowerCase().includes('italic') ?? false,
        )
      : style.mixed
        ? ''
        : (style.value ?? '')
  const underlineValue =
    changes.underlineStyle ?? (underline.mixed ? '' : (underline.value ?? 'none'))
  const spacingValue = changes.charSpacingTwips ?? (spacing.mixed ? 0 : (spacing.value ?? 0))
  const positionValue = changes.positionHalfPoints ?? (position.mixed ? 0 : (position.value ?? 0))
  const kernValue = changes.kernHalfPoints ?? (kern.mixed ? 0 : (kern.value ?? 0))
  const point = (halfPoints: number) => Math.abs(halfPoints) / 2
  const numberChange = (value: string, scaleFactor: number, min: number, max: number) => {
    const n = Number(value)
    return Number.isFinite(n) ? Math.round(Math.max(min, Math.min(max, n)) * scaleFactor) : 0
  }
  const apply = () => {
    applyFontChanges(editor, selection.from, selection.to, changes)
    onClose()
  }
  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal font-advanced-dialog">
        <h2>{t('appFontDialogTitle')}</h2>
        <div className="font-dialog-tabs" role="tablist">
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'font'}
            onClick={() => setTab('font')}
          >
            {t('appFontTabFont')}
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={tab === 'advanced'}
            onClick={() => setTab('advanced')}
          >
            {t('appFontTabAdvanced')}
          </button>
        </div>
        {tab === 'font' ? (
          <>
            <div className="font-dialog-row">
              <label>
                {t('appFontLatin')}
                <Dropdown
                  value={latinValue}
                  ariaLabel={t('appFontLatin')}
                  options={fontOptions(latinValue)}
                  onPick={(v) => {
                    if (v) change('fontAscii', v)
                  }}
                />
              </label>
              <label>
                {t('appFontEastAsian')}
                <Dropdown
                  value={eastValue}
                  ariaLabel={t('appFontEastAsian')}
                  options={fontOptions(eastValue)}
                  onPick={(v) => {
                    if (v) change('font', v)
                  }}
                />
              </label>
            </div>
            <div className="font-dialog-row">
              <label>
                {t('appFontStyleLabel')}
                <Dropdown
                  value={styleValue}
                  ariaLabel={t('appFontStyleLabel')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    ...FONT_STYLES.map((s) => ({ value: s.key, label: t(s.nameKey) })),
                  ]}
                  onPick={(v) => {
                    if (v)
                      setChanges((old) => ({
                        ...old,
                        bold: v.includes('bold'),
                        italic: v.includes('Italic') || v === 'italic',
                      }))
                  }}
                />
              </label>
              <label>
                {t('appFontSizeLabel')}
                <Dropdown
                  value={sizeValue ? String(sizeValue / 2) : ''}
                  ariaLabel={t('appFontSizeLabel')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    ...[...new Set([...FONT_SIZES, ...(sizeValue ? [sizeValue / 2] : [])])]
                      .sort((a, b) => a - b)
                      .map((n) => ({ value: String(n), label: String(n) })),
                  ]}
                  onPick={(v) => {
                    if (v) change('sizeHalfPoints', Math.round(Number(v) * 2))
                  }}
                />
              </label>
              <label>
                {t('appFontColor')}
                <input
                  type="color"
                  value={`#${colorHex(changes.color ?? (color.mixed ? '000000' : color.value || '000000'))}`}
                  onChange={(e) => change('color', e.target.value.slice(1).toUpperCase())}
                />
              </label>
            </div>
            <div className="font-dialog-row">
              <label>
                {t('appFontUnderlineStyle')}
                <Dropdown
                  value={underlineValue}
                  ariaLabel={t('appFontUnderlineStyle')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    ...UNDERLINES.map((v) => ({ value: v.key, label: t(v.label) })),
                  ]}
                  onPick={(v) => {
                    if (v) change('underlineStyle', v)
                  }}
                />
              </label>
              <label>
                {t('appFontUnderlineColor')}
                <input
                  type="color"
                  value={`#${colorHex(changes.underlineColor ?? (underlineColor.mixed ? '000000' : underlineColor.value || '000000'))}`}
                  onChange={(e) => change('underlineColor', e.target.value.slice(1).toUpperCase())}
                />
              </label>
            </div>
            <div className="font-dialog-checks">
              <Check
                label={t('appStrikethrough')}
                value={changes.strike ?? strike.value ?? false}
                mixed={changes.strike === undefined && strike.mixed}
                onChange={(v) => change('strike', v)}
              />
              <Check
                label={t('appFontDoubleStrike')}
                value={changes.dstrike ?? dstrike.value ?? false}
                mixed={changes.dstrike === undefined && dstrike.mixed}
                onChange={(v) => change('dstrike', v)}
              />
              <Check
                label={t('appFontSmallCaps')}
                value={(changes.caps ?? caps.value) === 'small'}
                mixed={changes.caps === undefined && caps.mixed}
                onChange={(v) => change('caps', v ? 'small' : 'none')}
              />
              <Check
                label={t('appFontAllCaps')}
                value={(changes.caps ?? caps.value) === 'all'}
                mixed={changes.caps === undefined && caps.mixed}
                onChange={(v) => change('caps', v ? 'all' : 'none')}
              />
              <Check
                label={t('appFontHidden')}
                value={changes.vanish ?? vanish.value ?? false}
                mixed={changes.vanish === undefined && vanish.mixed}
                onChange={(v) => change('vanish', v)}
              />
              <Check
                label={t('appSuperscript')}
                value={(changes.vertAlign ?? vertAlign.value) === 'superscript'}
                mixed={changes.vertAlign === undefined && vertAlign.mixed}
                onChange={(v) => change('vertAlign', v ? 'superscript' : '')}
              />
              <Check
                label={t('appSubscript')}
                value={(changes.vertAlign ?? vertAlign.value) === 'subscript'}
                mixed={changes.vertAlign === undefined && vertAlign.mixed}
                onChange={(v) => change('vertAlign', v ? 'subscript' : '')}
              />
            </div>
            <div className="font-dialog-row">
              <label>
                {t('appFontTextEffect')}
                <Dropdown
                  value={changes.textEffect ?? (effect.mixed ? '' : (effect.value ?? 'none'))}
                  ariaLabel={t('appFontTextEffect')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    ...EFFECTS.map((e) => ({ value: e.key, label: t(e.label) })),
                  ]}
                  onPick={(v) => {
                    if (v) change('textEffect', v as FontChanges['textEffect'])
                  }}
                />
              </label>
              <label>
                {t('appFontEmphasis')}
                <Dropdown
                  value={changes.em ?? (emphasis.mixed ? '' : (emphasis.value ?? 'none'))}
                  ariaLabel={t('appFontEmphasis')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    ...EMPHASIS.map((e) => ({ value: e.key, label: t(e.label) })),
                  ]}
                  onPick={(v) => {
                    if (v) change('em', v as FontChanges['em'])
                  }}
                />
              </label>
            </div>
          </>
        ) : (
          <>
            <div className="font-dialog-row">
              <label>
                {t('appFontScale')}
                <input
                  type="number"
                  min="1"
                  max="600"
                  step="1"
                  value={changes.charScalePct ?? (scale.mixed ? '' : (scale.value ?? 100))}
                  onChange={(e) => change('charScalePct', numberChange(e.target.value, 1, 1, 600))}
                />
                %
              </label>
              <label>
                {t('appFontSpacing')}
                <Dropdown
                  value={
                    spacing.mixed && changes.charSpacingTwips === undefined
                      ? ''
                      : spacingValue > 0
                        ? 'expanded'
                        : spacingValue < 0
                          ? 'condensed'
                          : 'normal'
                  }
                  ariaLabel={t('appFontSpacing')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    { value: 'normal', label: t('appFontNormal') },
                    { value: 'expanded', label: t('appFontExpanded') },
                    { value: 'condensed', label: t('appFontCondensed') },
                  ]}
                  onPick={(v) => {
                    if (v)
                      change(
                        'charSpacingTwips',
                        v === 'normal'
                          ? 0
                          : (v === 'expanded' ? 1 : -1) * (Math.abs(spacingValue) || 20),
                      )
                  }}
                />
              </label>
              <label>
                {t('appFontByPt')}
                <input
                  type="number"
                  min="0"
                  max="100"
                  step="0.1"
                  value={
                    spacing.mixed && changes.charSpacingTwips === undefined
                      ? ''
                      : Math.abs(spacingValue) / 20
                  }
                  onChange={(e) =>
                    change(
                      'charSpacingTwips',
                      (spacingValue < 0 ? -1 : 1) * numberChange(e.target.value, 20, 0, 100),
                    )
                  }
                />
              </label>
            </div>
            <div className="font-dialog-row">
              <label>
                {t('appFontPosition')}
                <Dropdown
                  value={
                    position.mixed && changes.positionHalfPoints === undefined
                      ? ''
                      : positionValue > 0
                        ? 'raised'
                        : positionValue < 0
                          ? 'lowered'
                          : 'normal'
                  }
                  ariaLabel={t('appFontPosition')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    { value: 'normal', label: t('appFontNormal') },
                    { value: 'raised', label: t('appFontRaised') },
                    { value: 'lowered', label: t('appFontLowered') },
                  ]}
                  onPick={(v) => {
                    if (v)
                      change(
                        'positionHalfPoints',
                        v === 'normal'
                          ? 0
                          : (v === 'raised' ? 1 : -1) * (Math.abs(positionValue) || 2),
                      )
                  }}
                />
              </label>
              <label>
                {t('appFontByPt')}
                <input
                  type="number"
                  min="0"
                  max="100"
                  step="0.5"
                  value={
                    position.mixed && changes.positionHalfPoints === undefined
                      ? ''
                      : point(positionValue)
                  }
                  onChange={(e) =>
                    change(
                      'positionHalfPoints',
                      (positionValue < 0 ? -1 : 1) * numberChange(e.target.value, 2, 0, 100),
                    )
                  }
                />
              </label>
            </div>
            <div className="font-dialog-row">
              <Check
                label={t('appFontKerning')}
                value={
                  changes.kernHalfPoints !== undefined
                    ? changes.kernHalfPoints > 0
                    : (kern.value ?? 0) > 0
                }
                mixed={changes.kernHalfPoints === undefined && kern.mixed}
                onChange={(v) => change('kernHalfPoints', v ? (kernValue > 0 ? kernValue : 16) : 0)}
              />
              <label>
                {t('appFontAtPt')}
                <input
                  type="number"
                  min="0"
                  max="100"
                  step="0.5"
                  value={kern.mixed && changes.kernHalfPoints === undefined ? '' : point(kernValue)}
                  onChange={(e) =>
                    change('kernHalfPoints', numberChange(e.target.value, 2, 0, 100))
                  }
                />
              </label>
              <Check
                label={t('appFontSnapToGrid')}
                value={changes.snapToGrid ?? snap.value ?? true}
                mixed={changes.snapToGrid === undefined && snap.mixed}
                onChange={(v) => change('snapToGrid', v)}
              />
            </div>
            <div className="font-dialog-row">
              <label>
                {t('appFontLigatures')}
                <Dropdown
                  value={changes.ligatures ?? (ligatures.mixed ? '' : (ligatures.value ?? 'none'))}
                  ariaLabel={t('appFontLigatures')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    ...LIGATURES.map((v) => ({ value: v.key, label: t(v.label) })),
                  ]}
                  onPick={(v) => {
                    if (v) change('ligatures', v as FontChanges['ligatures'])
                  }}
                />
              </label>
              <label>
                {t('appFontNumberSpacing')}
                <Dropdown
                  value={
                    changes.numberSpacing ??
                    (numberSpacing.mixed ? '' : (numberSpacing.value ?? 'default'))
                  }
                  ariaLabel={t('appFontNumberSpacing')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    ...NUMBER_SPACING.map((v) => ({ value: v.key, label: t(v.label) })),
                  ]}
                  onPick={(v) => {
                    if (v) change('numberSpacing', v as FontChanges['numberSpacing'])
                  }}
                />
              </label>
              <label>
                {t('appFontNumberForm')}
                <Dropdown
                  value={
                    changes.numberForm ?? (numberForm.mixed ? '' : (numberForm.value ?? 'default'))
                  }
                  ariaLabel={t('appFontNumberForm')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    ...NUMBER_FORM.map((v) => ({ value: v.key, label: t(v.label) })),
                  ]}
                  onPick={(v) => {
                    if (v) change('numberForm', v as FontChanges['numberForm'])
                  }}
                />
              </label>
            </div>
            <div className="font-dialog-row">
              <Check
                label={t('appFontContextualAlternates')}
                value={changes.contextualAlternates ?? contextualAlternates.value ?? false}
                mixed={changes.contextualAlternates === undefined && contextualAlternates.mixed}
                onChange={(v) => change('contextualAlternates', v)}
              />
              <label>
                {t('appFontStylisticSets')}
                <Dropdown
                  value={
                    changes.stylisticSet !== undefined
                      ? String(changes.stylisticSet)
                      : stylisticSet.mixed || stylisticSet.value === -1
                        ? ''
                        : String(stylisticSet.value ?? 0)
                  }
                  ariaLabel={t('appFontStylisticSets')}
                  options={[
                    { value: '', label: t('appFontMixedOrDefault') },
                    { value: '0', label: t('appFontDefault') },
                    ...Array.from({ length: 20 }, (_, index) => ({
                      value: String(index + 1),
                      label: `${t('appFontStylisticSet')} ${index + 1}`,
                    })),
                  ]}
                  onPick={(value) => {
                    if (value) change('stylisticSet', Number(value))
                  }}
                />
              </label>
            </div>
          </>
        )}
        <div
          className="font-preview"
          style={{
            fontFamily: latinValue ? cssFontFamily(latinValue) : undefined,
            fontSize: `${Math.min(sizeValue ? sizeValue / 2 : 11, 28)}pt`,
            fontWeight: styleValue.includes('bold') ? 600 : 400,
            fontStyle: styleValue.toLowerCase().includes('italic') ? 'italic' : 'normal',
            color: `#${colorHex(changes.color ?? (color.mixed ? '000000' : color.value || '000000'))}`,
            textDecoration:
              underlineValue !== 'none' && underlineValue !== '' ? 'underline' : undefined,
            letterSpacing: `${spacingValue / 20}pt`,
            verticalAlign: `${positionValue / 2}pt`,
            fontFeatureSettings:
              (changes.stylisticSet ?? stylisticSet.value ?? 0) > 0
                ? `"ss${String(changes.stylisticSet ?? stylisticSet.value).padStart(2, '0')}" 1`
                : undefined,
          }}
        >
          {t('appFontPreviewSample')}
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={apply}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
