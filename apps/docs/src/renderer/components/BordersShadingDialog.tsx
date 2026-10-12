import { useState } from 'react'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  bordersOf,
  hasSide,
  parseBorderLines,
  toggleSide,
  WORD_BORDER_SPACE_PT,
  type BorderLines,
  type BorderSide,
} from '../editor/paragraph-borders'
import { useModalKeys } from './modal-keys'

/**
 * Home → Borders ▾ → Borders and Shading…: the paragraph's four sides plus the rule between
 * paragraphs, one line style / weight / color for the drawn sides, and the shading fill. Returns the
 * paragraph attributes it edits (borders string, per-side lines, shading).
 */
export interface BordersShadingResult {
  borders: string | null
  borderLines: string | null
  shadingFill: string | null
}

const STYLES: Array<[string, StringKey]> = [
  ['single', 'ribbonBorderStyleSingle'],
  ['dotted', 'ribbonBorderStyleDotted'],
  ['dashed', 'ribbonBorderStyleDashed'],
  ['double', 'ribbonBorderStyleDouble'],
]
/** Word's weights in points */
const WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]
const SIDES: Array<[BorderSide, StringKey]> = [
  ['t', 'ribbonBorderTop'],
  ['b', 'ribbonBorderBottom'],
  ['l', 'ribbonBorderLeft'],
  ['r', 'ribbonBorderRight'],
  ['h', 'ribbonBorderBetween'],
]
const SETTINGS: Array<['none' | 'box' | 'all' | 'custom', StringKey]> = [
  ['none', 'ribbonBorderSettingNone'],
  ['box', 'ribbonBorderSettingBox'],
  ['all', 'ribbonBorderSettingAll'],
  ['custom', 'ribbonBorderSettingCustom'],
]

const settingOf = (borders: string | null): 'none' | 'box' | 'all' | 'custom' => {
  if (!borders) return 'none'
  if (borders === 'tblr') return 'box'
  if (borders === 'tblrh') return 'all'
  return 'custom'
}

export function BordersShadingDialog({
  borders: initialBorders,
  borderLines: rawLines,
  shadingFill: initialFill,
  onApply,
  onClose,
}: {
  borders: string | null
  borderLines: unknown
  shadingFill: string | null
  onApply: (next: BordersShadingResult) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const lines: BorderLines = parseBorderLines(rawLines)
  const first = (initialBorders ?? '').split('').map((s) => lines[s as BorderSide])[0]

  const [borders, setBorders] = useState<string | null>(initialBorders)
  const [style, setStyle] = useState(first?.val ?? 'single')
  const [width, setWidth] = useState(first?.szPt ?? 0.5)
  const [auto, setAuto] = useState(!first?.color)
  const [color, setColor] = useState(`#${first?.color ?? '000000'}`)
  const [fill, setFill] = useState<string | null>(initialFill)

  const setting = settingOf(borders)
  const hex = color.replace('#', '').toUpperCase()

  const submit = () => {
    const drawn = bordersOf((borders ?? '').split(''))
    const line = {
      szPt: width,
      ...(style !== 'single' ? { val: style } : {}),
      ...(auto ? {} : { color: hex }),
    }
    const out: BorderLines = {}
    // a side that already had a gap keeps it; a new one gets Word's (1 pt above / below, 4 pt beside)
    for (const s of (drawn ?? '').split('') as BorderSide[]) {
      out[s] = { ...line, spacePt: lines[s]?.spacePt ?? WORD_BORDER_SPACE_PT[s] }
    }
    onApply({
      borders: drawn,
      borderLines: drawn ? JSON.stringify(out) : null,
      shadingFill: fill,
    })
    onClose()
  }

  const previewLine = `${Math.max(1, Math.round((width * 96) / 72))}px ${
    style === 'double'
      ? 'double'
      : style === 'dotted'
        ? 'dotted'
        : style === 'dashed'
          ? 'dashed'
          : 'solid'
  } ${auto ? 'currentColor' : color}`

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal borders-dialog">
        <h2>{t('ribbonBordersDialogTitle')}</h2>
        <div className="borders-dialog-grid">
          <div>
            <div className="borders-dialog-head">{t('ribbonBorderSetting')}</div>
            <div className="cols-presets">
              {SETTINGS.map(([id, key]) => (
                <button
                  key={id}
                  type="button"
                  aria-pressed={setting === id}
                  className={setting === id ? 'active' : ''}
                  onClick={() =>
                    setBorders(
                      id === 'none'
                        ? null
                        : id === 'box'
                          ? 'tblr'
                          : id === 'all'
                            ? 'tblrh'
                            : borders,
                    )
                  }
                >
                  {t(key)}
                </button>
              ))}
            </div>
            {SIDES.map(([side, key]) => (
              <label className="font-check" key={side}>
                <input
                  type="checkbox"
                  checked={hasSide(borders, side)}
                  onChange={() => setBorders((b) => toggleSide(b, side))}
                />
                {t(key)}
              </label>
            ))}
          </div>
          <div>
            <label>
              {t('ribbonBorderStyle')}
              <select value={style} onChange={(e) => setStyle(e.target.value)}>
                {STYLES.map(([id, key]) => (
                  <option key={id} value={id}>
                    {t(key)}
                  </option>
                ))}
              </select>
            </label>
            <label>
              {t('ribbonBorderWidth')} (pt)
              <select value={String(width)} onChange={(e) => setWidth(Number(e.target.value))}>
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
                onChange={(e) => setColor(e.target.value)}
              />
            </label>
            <label className="font-check">
              <input type="checkbox" checked={auto} onChange={(e) => setAuto(e.target.checked)} />
              {t('ribbonBorderAutoColor')}
            </label>
          </div>
        </div>
        <div className="borders-dialog-head">{t('ribbonParagraphShading')}</div>
        <div className="borders-dialog-fill">
          <input
            type="color"
            className="font-color-input"
            aria-label={t('ribbonParagraphShading')}
            value={`#${fill ?? 'FFFFFF'}`}
            disabled={fill === null}
            onChange={(e) => setFill(e.target.value.replace('#', '').toUpperCase())}
          />
          <label className="font-check">
            <input
              type="checkbox"
              checked={fill === null}
              onChange={(e) => setFill(e.target.checked ? null : 'FFFF99')}
            />
            {t('ribbonNoShading')}
          </label>
        </div>
        <div className="borders-dialog-preview" aria-hidden>
          <div
            style={{
              borderTop: hasSide(borders, 't') ? previewLine : undefined,
              borderBottom: hasSide(borders, 'b') ? previewLine : undefined,
              borderLeft: hasSide(borders, 'l') ? previewLine : undefined,
              borderRight: hasSide(borders, 'r') ? previewLine : undefined,
              background: fill ? `#${fill}` : undefined,
            }}
          >
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
