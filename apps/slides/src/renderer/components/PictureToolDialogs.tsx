/** Alt Text and Picture Quality dialogs for the Picture Format tab. Reuses SettingsModal's modal styles. */
import { useEffect, useState } from 'react'
import { useI18n } from '../i18n/locale'

export function AltTextDialog({
  initialText,
  initialDecorative,
  onSave,
  onClose,
}: {
  initialText: string
  initialDecorative: boolean
  onSave: (text: string, decorative: boolean) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [text, setText] = useState(initialText)
  const [decorative, setDecorative] = useState(initialDecorative)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])
  return (
    <div className="modal-backdrop" onClick={onClose} data-testid="alt-text-dialog">
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonAltText')}</h2>
        <label>
          {t('ribbonAltTextLabel')}
          <textarea
            value={decorative ? '' : text}
            disabled={decorative}
            rows={4}
            autoFocus
            data-testid="alt-text-input"
            onChange={(e) => setText(e.target.value)}
          />
        </label>
        <label className="fp-checkrow">
          <input
            type="checkbox"
            checked={decorative}
            data-testid="alt-text-decorative"
            onChange={(e) => setDecorative(e.target.checked)}
          />
          <span>{t('ribbonAltTextDecorative')}</span>
        </label>
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button
            className="primary"
            data-testid="alt-text-save"
            onClick={() => onSave(text.trim(), decorative)}
          >
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

export const PICTURE_QUALITY_PPI = [220, 150, 96] as const

export function PictureQualityDialog({
  hasCrop,
  onApply,
  onClose,
}: {
  hasCrop: boolean
  onApply: (ppi: number, deleteCropped: boolean) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [ppi, setPpi] = useState<number>(150)
  const [deleteCropped, setDeleteCropped] = useState(true)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])
  const labels: Record<number, string> = {
    220: t('ribbonPicQualityPrint'),
    150: t('ribbonPicQualityWeb'),
    96: t('ribbonPicQualityEmail'),
  }
  return (
    <div className="modal-backdrop" onClick={onClose} data-testid="picture-quality-dialog">
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <h2>{t('ribbonPictureQuality')}</h2>
        {PICTURE_QUALITY_PPI.map((value) => (
          <label key={value} className="fp-checkrow">
            <input
              type="radio"
              name="picture-ppi"
              checked={ppi === value}
              data-testid={`picture-quality-${value}`}
              onChange={() => setPpi(value)}
            />
            <span>{labels[value]}</span>
          </label>
        ))}
        <label className="fp-checkrow">
          <input
            type="checkbox"
            checked={deleteCropped && hasCrop}
            disabled={!hasCrop}
            onChange={(e) => setDeleteCropped(e.target.checked)}
          />
          <span>{t('ribbonPicQualityDeleteCropped')}</span>
        </label>
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
          <button
            className="primary"
            data-testid="picture-quality-apply"
            onClick={() => onApply(ppi, deleteCropped && hasCrop)}
          >
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
