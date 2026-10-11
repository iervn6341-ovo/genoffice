import { useState } from 'react'
import type { LineNumbering, SectionSettings } from '@genoffice/docx-engine'
import { useI18n } from '../i18n/locale'
import { cmFromTwips, marginsFitPage, twipsFromCmInput } from './MarginDialog'
import { useModalKeys } from './modal-keys'

/**
 * Layout-tab dialogs Word opens from the Page Setup dropdowns: More Columns,
 * More Paper Sizes and Line Numbering Options. Each returns a full replacement
 * SectionSettings through onApply (the ribbon routes it to the cursor's section).
 */

const TWIPS_PER_CM = 1440 / 2.54
/** Word's limits for a custom page (1.0 in .. 22 in) */
const MIN_PAGE_TWIPS = 1440
const MAX_PAGE_TWIPS = 31680
/** Word's default column gap (0.5 in) */
const DEFAULT_COL_SPACE = 720
const MAX_COLUMNS = 12

/** Column widths in twips for a preset: Left = narrow first column, Right = narrow last (1 : 2) */
export function columnPresetWidths(
  kind: 'left' | 'right',
  textWidth: number,
  space: number,
): number[] {
  const total = Math.max(0, textWidth - space)
  const narrow = Math.round(total / 3)
  return kind === 'left' ? [narrow, total - narrow] : [total - narrow, narrow]
}

/** SectionSettings for n equal columns, or explicit widths when given (equal = no widths) */
export function withColumns(
  section: SectionSettings,
  columns: number,
  opts: { space?: number; widths?: number[]; sep?: boolean } = {},
): SectionSettings {
  const { colWidths: _drop, ...rest } = section
  const next: SectionSettings = {
    ...rest,
    columns,
    ...(columns > 1 ? { colSpace: opts.space ?? section.colSpace ?? DEFAULT_COL_SPACE } : {}),
  }
  if (opts.widths && columns > 1 && opts.widths.length === columns) next.colWidths = opts.widths
  // the line between the columns only exists for 2+ columns; keep an explicit choice either way
  if (opts.sep !== undefined) next.colSep = columns > 1 ? opts.sep : false
  else if (columns <= 1 && section.colSep) next.colSep = false
  return next
}

function Backdrop({
  onClose,
  children,
}: {
  onClose: () => void
  children: (keys: ReturnType<typeof useModalKeys>) => React.ReactNode
}) {
  const keys = useModalKeys(onClose)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      {children(keys)}
    </div>
  )
}

/* ---------------------------------------------------------------- Columns */

export function ColumnsDialog({
  section,
  onApply,
  onClose,
}: {
  section: SectionSettings
  onApply: (next: SectionSettings) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const textWidth = section.pageWidth - section.marginLeft - section.marginRight
  const [count, setCount] = useState(Math.max(1, section.columns))
  const [equal, setEqual] = useState(!section.colWidths)
  const [sep, setSep] = useState(!!section.colSep)
  const [spacing, setSpacing] = useState(String(cmFromTwips(section.colSpace ?? DEFAULT_COL_SPACE)))
  const [widths, setWidths] = useState<string[]>(() =>
    (section.colWidths ?? []).map((w) => String(cmFromTwips(w))),
  )

  const space = twipsFromCmInput(spacing, section.colSpace ?? DEFAULT_COL_SPACE)
  const maxSpace = count > 1 ? Math.floor((textWidth - 720 * count) / (count - 1)) : 0
  const spaceBad = count > 1 && (space < 0 || space > Math.max(0, maxSpace))

  /** the width fields in cm for the current count, padded from an equal split */
  const fieldWidths = (): string[] => {
    const each = Math.max(0, (textWidth - space * (count - 1)) / Math.max(1, count))
    return Array.from({ length: count }, (_x, i) => widths[i] ?? String(cmFromTwips(each)))
  }

  const preset = (kind: 'one' | 'two' | 'three' | 'left' | 'right') => {
    if (kind === 'one' || kind === 'two' || kind === 'three') {
      setCount(kind === 'one' ? 1 : kind === 'two' ? 2 : 3)
      setEqual(true)
      return
    }
    setCount(2)
    setEqual(false)
    setWidths(columnPresetWidths(kind, textWidth, space).map((w) => String(cmFromTwips(w))))
  }

  const submit = () => {
    if (spaceBad) return
    if (count <= 1) {
      onApply(withColumns(section, 1, { sep: false }))
      onClose()
      return
    }
    if (equal) {
      onApply(withColumns(section, count, { space, sep }))
    } else {
      // relative widths: scaled so the columns plus the gaps fill the text width exactly
      const raw = fieldWidths().map((w) => Math.max(0.5, Number(w) || 0))
      const available = Math.max(0, textWidth - space * (count - 1))
      const sum = raw.reduce((a, b) => a + b, 0)
      const scaled = raw.map((w) => Math.round((w / sum) * available))
      scaled[scaled.length - 1] += available - scaled.reduce((a, b) => a + b, 0)
      onApply(withColumns(section, count, { space, widths: scaled, sep }))
    }
    onClose()
  }

  const presetOf = (kind: Parameters<typeof preset>[0]) => (
    <button key={kind} type="button" className="btn-ghost" onClick={() => preset(kind)}>
      {t(
        kind === 'one'
          ? 'ribbonOneColumn'
          : kind === 'two'
            ? 'ribbonTwoColumns'
            : kind === 'three'
              ? 'ribbonThreeColumns'
              : kind === 'left'
                ? 'ribbonColumnsLeft'
                : 'ribbonColumnsRight',
      )}
    </button>
  )

  return (
    <Backdrop onClose={onClose}>
      {(keys) => (
        <div className="modal cols-dialog" onKeyDown={keys.onKeyDown}>
          <h2>{t('ribbonColumnsDialogTitle')}</h2>
          <div className="cols-presets">
            {(['one', 'two', 'three', 'left', 'right'] as const).map(presetOf)}
          </div>
          <div className="modal-row margin-row">
            <label>
              {t('ribbonColumnsCount')}
              <input
                type="number"
                min={1}
                max={MAX_COLUMNS}
                value={count}
                onChange={(e) =>
                  setCount(
                    Math.min(MAX_COLUMNS, Math.max(1, Math.round(Number(e.target.value)) || 1)),
                  )
                }
              />
            </label>
            <label>
              {t('ribbonColumnsSpacing')} (cm)
              <input
                type="number"
                min={0}
                step={0.1}
                disabled={count <= 1}
                value={spacing}
                onChange={(e) => setSpacing(e.target.value)}
              />
            </label>
          </div>
          <label className="font-check">
            <input
              type="checkbox"
              checked={equal}
              disabled={count <= 1}
              onChange={(e) => {
                setEqual(e.target.checked)
                if (!e.target.checked) setWidths(fieldWidths())
              }}
            />
            {t('ribbonColumnsEqual')}
          </label>
          {!equal && count > 1 && (
            <div className="cols-widths">
              {fieldWidths().map((w, i) => (
                <label key={i}>
                  {t('ribbonColumnsWidthN', { n: i + 1 })} (cm)
                  <input
                    type="number"
                    min={0.5}
                    step={0.1}
                    value={w}
                    onChange={(e) =>
                      setWidths(fieldWidths().map((x, j) => (j === i ? e.target.value : x)))
                    }
                  />
                </label>
              ))}
            </div>
          )}
          <label className="font-check">
            <input
              type="checkbox"
              checked={sep}
              disabled={count <= 1}
              onChange={(e) => setSep(e.target.checked)}
            />
            {t('ribbonColumnsLineBetween')}
          </label>
          {spaceBad && <div className="modal-error">{t('ribbonColumnsSpacingTooLarge')}</div>}
          <div className="modal-actions">
            <button className="btn-ghost" onClick={onClose}>
              {t('appCancel')}
            </button>
            <button className="btn-primary" disabled={spaceBad} onClick={submit}>
              {t('appOk')}
            </button>
          </div>
        </div>
      )}
    </Backdrop>
  )
}

/* -------------------------------------------------------------- Paper size */

export function PaperSizeDialog({
  section,
  onApply,
  onClose,
}: {
  section: SectionSettings
  onApply: (next: SectionSettings) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [width, setWidth] = useState(String(cmFromTwips(section.pageWidth)))
  const [height, setHeight] = useState(String(cmFromTwips(section.pageHeight)))

  const w = twipsFromCmInput(width, section.pageWidth)
  const h = twipsFromCmInput(height, section.pageHeight)
  const inRange = (v: number) => v >= MIN_PAGE_TWIPS && v <= MAX_PAGE_TWIPS
  const fits = marginsFitPage(
    {
      top: section.marginTop,
      right: section.marginRight,
      bottom: section.marginBottom,
      left: section.marginLeft,
    },
    w,
    h,
  )
  const invalid = !inRange(w) || !inRange(h) || !fits

  const submit = () => {
    if (invalid) return
    onApply({
      ...section,
      pageWidth: w,
      pageHeight: h,
      orientation: w > h ? 'landscape' : 'portrait',
    })
    onClose()
  }

  return (
    <Backdrop onClose={onClose}>
      {(keys) => (
        <div className="modal" onKeyDown={keys.onKeyDown}>
          <h2>{t('ribbonPaperDialogTitle')}</h2>
          <div className="modal-row margin-row">
            <label>
              {t('ribbonPaperWidth')} (cm)
              <input
                type="number"
                min={2.54}
                max={55.88}
                step={0.1}
                value={width}
                onChange={(e) => setWidth(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && submit()}
              />
            </label>
            <label>
              {t('ribbonPaperHeight')} (cm)
              <input
                type="number"
                min={2.54}
                max={55.88}
                step={0.1}
                value={height}
                onChange={(e) => setHeight(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && submit()}
              />
            </label>
          </div>
          {invalid && <div className="modal-error">{t('ribbonPaperInvalid')}</div>}
          <div className="modal-actions">
            <button className="btn-ghost" onClick={onClose}>
              {t('appCancel')}
            </button>
            <button className="btn-primary" disabled={invalid} onClick={submit}>
              {t('appOk')}
            </button>
          </div>
        </div>
      )}
    </Backdrop>
  )
}

/* ------------------------------------------------------------ Line numbers */

export const DEFAULT_LINE_NUMBERING: LineNumbering = {
  countBy: 1,
  start: 1,
  restart: 'newPage',
}

export function LineNumbersDialog({
  current,
  onApply,
  onClose,
}: {
  current: LineNumbering | null | undefined
  onApply: (next: LineNumbering) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const base = current ?? DEFAULT_LINE_NUMBERING
  const [start, setStart] = useState(String(base.start))
  const [countBy, setCountBy] = useState(String(base.countBy))
  const [distance, setDistance] = useState(
    base.distance !== undefined ? String(cmFromTwips(base.distance)) : '',
  )
  const [restart, setRestart] = useState<LineNumbering['restart']>(base.restart)

  const startN = Math.max(1, Math.round(Number(start)) || 1)
  const countN = Math.max(1, Math.round(Number(countBy)) || 1)

  const submit = () => {
    const distanceTwips =
      distance.trim() === ''
        ? undefined
        : Math.max(0, Math.round((Number(distance) || 0) * TWIPS_PER_CM))
    onApply({
      countBy: countN,
      start: startN,
      restart,
      ...(distanceTwips !== undefined ? { distance: distanceTwips } : {}),
    })
    onClose()
  }

  const radio = (value: LineNumbering['restart'], labelKey: Parameters<typeof t>[0]) => (
    <label className="font-check" key={value}>
      <input
        type="radio"
        name="ln-restart"
        checked={restart === value}
        onChange={() => setRestart(value)}
      />
      {t(labelKey)}
    </label>
  )

  return (
    <Backdrop onClose={onClose}>
      {(keys) => (
        <div className="modal" onKeyDown={keys.onKeyDown}>
          <h2>{t('ribbonLineNumbers')}</h2>
          <div className="modal-row margin-row">
            <label>
              {t('ribbonLineNumStart')}
              <input
                type="number"
                min={1}
                value={start}
                onChange={(e) => setStart(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && submit()}
              />
            </label>
            <label>
              {t('ribbonLineNumCountBy')}
              <input
                type="number"
                min={1}
                value={countBy}
                onChange={(e) => setCountBy(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && submit()}
              />
            </label>
          </div>
          <label>
            {t('ribbonLineNumDistance')} (cm)
            <input
              type="number"
              min={0}
              step={0.1}
              placeholder={String(cmFromTwips(360))}
              value={distance}
              onChange={(e) => setDistance(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && submit()}
            />
          </label>
          <div className="cols-presets">{t('ribbonLineNumNumbering')}</div>
          {radio('newPage', 'ribbonLineNumRestartPage')}
          {radio('newSection', 'ribbonLineNumRestartSection')}
          {radio('continuous', 'ribbonLineNumContinuous')}
          <div className="modal-actions">
            <button className="btn-ghost" onClick={onClose}>
              {t('appCancel')}
            </button>
            <button className="btn-primary" onClick={submit}>
              {t('appOk')}
            </button>
          </div>
        </div>
      )}
    </Backdrop>
  )
}
