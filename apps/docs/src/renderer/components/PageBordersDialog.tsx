import { useState } from 'react'
import type { SectionSettings } from '@genoffice/docx-engine'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/**
 * Design → Page Borders → Page Borders…: sides, line style, weight and color of the page box and
 * which pages it covers. Returns the section with `pageBorder` / `pageBorderProps` replaced; the
 * saver (applySectionSettings) writes w:pgBorders from those details.
 */
type Side = 'top' | 'right' | 'bottom' | 'left'
type Props = NonNullable<SectionSettings['pageBorderProps']>
type Setting = 'none' | 'box' | 'custom'

const SIDES: Array<[Side, StringKey]> = [
  ['top', 'ribbonBorderTop'],
  ['bottom', 'ribbonBorderBottom'],
  ['left', 'ribbonBorderLeft'],
  ['right', 'ribbonBorderRight'],
]
/** OOXML line styles the page renderer draws (pageBorderLineCss) */
const STYLES: Array<[string, StringKey]> = [
  ['single', 'ribbonBorderStyleSingle'],
  ['dotted', 'ribbonBorderStyleDotted'],
  ['dashed', 'ribbonBorderStyleDashed'],
  ['double', 'ribbonBorderStyleDouble'],
  ['triple', 'ribbonBorderStyleTriple'],
  ['thinThickSmallGap', 'ribbonBorderStyleThinThick'],
  ['wave', 'ribbonBorderStyleWave'],
]
const WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]
const DEFAULT_SPACE_PT = 24

const settingOf = (sides: Record<Side, boolean>): Setting => {
  const on = SIDES.filter(([s]) => sides[s]).length
  return on === 0 ? 'none' : on === 4 ? 'box' : 'custom'
}

export function PageBordersDialog({
  section,
  onApply,
  onClose,
}: {
  section: SectionSettings
  onApply: (next: SectionSettings) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const current = section.pageBorder ? section.pageBorderProps : undefined
  const first = current?.sides ? Object.values(current.sides).find(Boolean) : undefined

  const [sides, setSides] = useState<Record<Side, boolean>>(() => {
    if (!section.pageBorder) return { top: true, right: true, bottom: true, left: true }
    const s = current?.sides
    return s && Object.keys(s).length > 0
      ? { top: !!s.top, right: !!s.right, bottom: !!s.bottom, left: !!s.left }
      : { top: true, right: true, bottom: true, left: true }
  })
  const [none, setNone] = useState(!section.pageBorder)
  const [style, setStyle] = useState(first?.val && !first.art ? first.val : 'single')
  const [width, setWidth] = useState(first && !first.art ? first.widthPt : 0.5)
  const [auto, setAuto] = useState(!first?.color)
  const [color, setColor] = useState(`#${first?.color ?? '000000'}`)
  const [space, setSpace] = useState(String(first?.spacePt ?? DEFAULT_SPACE_PT))
  const [from, setFrom] = useState<'page' | 'text'>(current?.offsetFrom ?? 'page')
  const [display, setDisplay] = useState<'all' | 'firstPage' | 'notFirstPage'>(
    current?.display ?? 'all',
  )

  const setting: Setting = none ? 'none' : settingOf(sides)
  /** choosing a style / weight / color while the setting is None turns on the box, like picking Box first */
  const activate = () => {
    if (!none) return
    setNone(false)
    setSides({ top: true, right: true, bottom: true, left: true })
  }
  const spacePt = Math.min(31, Math.max(0, Math.round(Number(space) || 0)))

  const submit = () => {
    const { pageBorderProps: _old, ...rest } = section
    if (none || !SIDES.some(([s]) => sides[s])) {
      onApply({ ...rest, pageBorder: false })
      onClose()
      return
    }
    const line = {
      val: style,
      widthPt: width,
      spacePt,
      ...(auto ? {} : { color: color.replace('#', '').toUpperCase() }),
    }
    const out: NonNullable<Props['sides']> = {}
    for (const [s] of SIDES) if (sides[s]) out[s] = line
    onApply({
      ...rest,
      pageBorder: true,
      pageBorderProps: {
        offsetFrom: from,
        ...(display !== 'all' ? { display } : {}),
        spacePt,
        widthPt: width,
        ...(auto ? {} : { color: line.color }),
        sides: out,
      },
    })
    onClose()
  }

  const css = `${Math.max(1, Math.round((width * 96) / 72))}px ${
    style === 'dotted'
      ? 'dotted'
      : style === 'dashed'
        ? 'dashed'
        : style === 'single'
          ? 'solid'
          : 'double'
  } ${auto ? 'currentColor' : color}`

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal borders-dialog">
        <h2>{t('ribbonPageBordersDialogTitle')}</h2>
        <div className="borders-dialog-grid">
          <div>
            <div className="borders-dialog-head">{t('ribbonBorderSetting')}</div>
            <div className="cols-presets">
              {(
                [
                  ['none', 'ribbonBorderSettingNone'],
                  ['box', 'ribbonBorderSettingBox'],
                  ['custom', 'ribbonBorderSettingCustom'],
                ] as Array<[Setting, StringKey]>
              ).map(([id, key]) => (
                <button
                  key={id}
                  type="button"
                  aria-pressed={setting === id}
                  className={setting === id ? 'active' : ''}
                  onClick={() => {
                    setNone(id === 'none')
                    if (id === 'box') setSides({ top: true, right: true, bottom: true, left: true })
                  }}
                >
                  {t(key)}
                </button>
              ))}
            </div>
            {SIDES.map(([side, key]) => (
              <label className="font-check" key={side}>
                <input
                  type="checkbox"
                  checked={!none && sides[side]}
                  onChange={(e) => {
                    setNone(false)
                    setSides((s) => ({ ...s, [side]: e.target.checked }))
                  }}
                />
                {t(key)}
              </label>
            ))}
          </div>
          <div>
            <label>
              {t('ribbonBorderStyle')}
              <select
                value={style}
                onChange={(e) => {
                  activate()
                  setStyle(e.target.value)
                }}
              >
                {STYLES.map(([id, key]) => (
                  <option key={id} value={id}>
                    {t(key)}
                  </option>
                ))}
              </select>
            </label>
            <label>
              {t('ribbonBorderWidth')} (pt)
              <select
                value={String(width)}
                onChange={(e) => {
                  activate()
                  setWidth(Number(e.target.value))
                }}
              >
                {WIDTHS.map((w) => (
                  <option key={w} value={String(w)}>
                    {w}
                  </option>
                ))}
              </select>
            </label>
            <label>
              {t('ribbonBorderColor')}
              <input
                type="color"
                className="font-color-input"
                value={color}
                disabled={auto}
                onChange={(e) => {
                  activate()
                  setColor(e.target.value)
                }}
              />
            </label>
            <label className="font-check">
              <input
                type="checkbox"
                checked={auto}
                onChange={(e) => {
                  activate()
                  setAuto(e.target.checked)
                }}
              />
              {t('ribbonBorderAutoColor')}
            </label>
          </div>
        </div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonPageBorderApplyTo')}
            <select value={display} onChange={(e) => setDisplay(e.target.value as typeof display)}>
              <option value="all">{t('ribbonPageBorderAllPages')}</option>
              <option value="firstPage">{t('ribbonPageBorderFirstPage')}</option>
              <option value="notFirstPage">{t('ribbonPageBorderNotFirstPage')}</option>
            </select>
          </label>
          <label>
            {t('ribbonPageBorderMeasure')}
            <select value={from} onChange={(e) => setFrom(e.target.value as 'page' | 'text')}>
              <option value="page">{t('ribbonPageBorderFromPage')}</option>
              <option value="text">{t('ribbonPageBorderFromText')}</option>
            </select>
          </label>
          <label>
            {t('ribbonPageBorderDistance')} (pt)
            <input
              type="number"
              min={0}
              max={31}
              value={space}
              onChange={(e) => setSpace(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && submit()}
            />
          </label>
        </div>
        <div className="borders-dialog-preview" aria-hidden>
          <div
            style={{
              borderTop: !none && sides.top ? css : undefined,
              borderBottom: !none && sides.bottom ? css : undefined,
              borderLeft: !none && sides.left ? css : undefined,
              borderRight: !none && sides.right ? css : undefined,
            }}
          >
            <span />
            <span />
            <span />
          </div>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
