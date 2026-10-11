import { useState } from 'react'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

export interface PageMargins {
  top: number
  right: number
  bottom: number
  left: number
}

const TWIPS_PER_CM = 1440 / 2.54
/** Word rejects margins that leave less than about one inch of body */
const MIN_BODY_TWIPS = 1440

export const cmFromTwips = (twips: number): number => Math.round((twips / TWIPS_PER_CM) * 100) / 100

export const twipsFromCmInput = (value: string, originalTwips: number): number => {
  if (value === String(cmFromTwips(originalTwips))) return originalTwips
  return Math.round(Math.max(0, Number(value) || 0) * TWIPS_PER_CM)
}

export const marginsFitPage = (
  margins: PageMargins,
  pageWidth: number,
  pageHeight: number,
): boolean =>
  margins.left + margins.right <= pageWidth - MIN_BODY_TWIPS &&
  margins.top + margins.bottom <= pageHeight - MIN_BODY_TWIPS

type Side = 'top' | 'bottom' | 'left' | 'right'

/** what the dialog adds to the four margins: the binding gutter and "Multiple pages: Mirror margins" */
export interface MarginExtras {
  /** twips of binding gutter, already included in the returned left margin */
  gutter: number
  mirror: boolean
}

const SIDES: Array<[Side, StringKey]> = [
  ['top', 'ribbonMarginTop'],
  ['bottom', 'ribbonMarginBottom'],
  ['left', 'ribbonMarginLeft'],
  ['right', 'ribbonMarginRight'],
]

export function MarginDialog({
  margins,
  gutter,
  mirror,
  pageWidth,
  pageHeight,
  onApply,
  onClose,
}: {
  /** total margins; the left one includes the gutter */
  margins: PageMargins
  /** twips of binding gutter, or null when it sits at the top (not editable here) */
  gutter: number | null
  mirror: boolean
  pageWidth: number
  pageHeight: number
  onApply: (next: PageMargins, extras: MarginExtras) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const gutterTwips = gutter ?? 0
  const baseLeft = Math.max(0, margins.left - gutterTwips)
  const [values, setValues] = useState<Record<Side, string>>({
    top: String(cmFromTwips(margins.top)),
    bottom: String(cmFromTwips(margins.bottom)),
    left: String(cmFromTwips(baseLeft)),
    right: String(cmFromTwips(margins.right)),
  })
  const [gutterValue, setGutterValue] = useState(String(cmFromTwips(gutterTwips)))
  const [mirrored, setMirrored] = useState(mirror)
  const modalKeys = useModalKeys(onClose)

  const twips = (side: Side) =>
    twipsFromCmInput(values[side], side === 'left' ? baseLeft : margins[side])
  const gutterNext = gutter === null ? 0 : twipsFromCmInput(gutterValue, gutterTwips)
  const next: PageMargins = {
    top: twips('top'),
    right: twips('right'),
    bottom: twips('bottom'),
    left: twips('left') + gutterNext,
  }
  const tooLarge = !marginsFitPage(next, pageWidth, pageHeight)

  const submit = () => {
    if (tooLarge) return
    onApply(next, { gutter: gutterNext, mirror: mirrored })
    onClose()
  }

  const labelKeyOf = (side: Side, labelKey: StringKey): StringKey =>
    mirrored && side === 'left'
      ? 'ribbonMarginInside'
      : mirrored && side === 'right'
        ? 'ribbonMarginOutside'
        : labelKey

  const field = ([side, labelKey]: [Side, StringKey]) => (
    <label key={side}>
      {t(labelKeyOf(side, labelKey))} (cm)
      <input
        type="number"
        min={0}
        step={0.1}
        value={values[side]}
        onChange={(e) => setValues((v) => ({ ...v, [side]: e.target.value }))}
        onKeyDown={(e) => {
          if (e.key === 'Enter') submit()
        }}
      />
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{t('ribbonMarginDialogTitle')}</h2>
        <div className="modal-row margin-row">{SIDES.slice(0, 2).map(field)}</div>
        <div className="modal-row margin-row">{SIDES.slice(2).map(field)}</div>
        <div className="modal-row margin-row">
          <label>
            {t('ribbonMarginGutter')} (cm)
            <input
              type="number"
              min={0}
              step={0.1}
              disabled={gutter === null}
              value={gutterValue}
              onChange={(e) => setGutterValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit()
              }}
            />
          </label>
          <label>
            {t('ribbonMarginMultiPages')}
            <select
              value={mirrored ? 'mirror' : 'normal'}
              onChange={(e) => setMirrored(e.target.value === 'mirror')}
            >
              <option value="normal">{t('ribbonMarginMultiNormal')}</option>
              <option value="mirror">{t('ribbonMarginMultiMirror')}</option>
            </select>
          </label>
        </div>
        {tooLarge && <div className="modal-error">{t('ribbonMarginTooLarge')}</div>}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" disabled={tooLarge} onClick={submit}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
