import React, { useState } from 'react'
import { useI18n } from '../i18n/locale'
import { preselectedScreen, type ShowScreen } from '../show-screen'

/**
 * Shown before a show starts when several screens could take it (a laptop plus two or more
 * external ones): the user picks where the audience sees the slides. Presenter view stays on the
 * screen it opens on.
 */
export function ShowScreenDialog({
  candidates,
  presenterLabel,
  remembered,
  onStart,
  onClose,
}: {
  candidates: ShowScreen[]
  /** the screen presenter view stays on */
  presenterLabel: string
  remembered: string | null
  onStart: (screen: ShowScreen) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [pickedId, setPickedId] = useState(() => preselectedScreen(candidates, remembered)?.id)
  const picked = candidates.find((s) => s.id === pickedId)
  const start = () => picked && onStart(picked)
  return (
    <div
      className="modal-backdrop"
      onClick={onClose}
      onKeyDown={(e) => {
        if (e.key === 'Escape') {
          e.stopPropagation()
          onClose()
        } else if (e.key === 'Enter') {
          e.preventDefault()
          start()
        }
      }}
    >
      <div
        className="modal ssd"
        role="dialog"
        aria-label={t('paneShowScreenTitle')}
        onClick={(e) => e.stopPropagation()}
      >
        <h2>{t('paneShowScreenTitle')}</h2>
        <p className="ssd-note">{t('paneShowScreenNote', { name: presenterLabel })}</p>
        <div className="ssd-list" role="radiogroup">
          {candidates.map((s) => (
            <button
              key={s.id}
              role="radio"
              aria-checked={s.id === pickedId}
              className={`ssd-item${s.id === pickedId ? ' active' : ''}`}
              onClick={() => setPickedId(s.id)}
              onDoubleClick={() => onStart(s)}
              autoFocus={s.id === pickedId}
            >
              <span className="ssd-name">{s.label}</span>
              <span className="ssd-size">
                {s.width} × {s.height}
              </span>
            </button>
          ))}
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('paneCancel')}</button>
          <button className="primary" disabled={!picked} onClick={start}>
            {t('paneShowScreenStart')}
          </button>
        </div>
      </div>
    </div>
  )
}
