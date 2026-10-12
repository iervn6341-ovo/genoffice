import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { parseAreaText } from './table-design'

/// Excel's Resize Table dialog: the table's range as text, to be changed. OK hands a range that reads
/// as one to the command (which checks it against the table); anything else stays open with a hint.
export function ResizeTableDialog({
  range,
  onApply,
  onClose,
}: {
  readonly range: string
  readonly onApply: (range: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [text, setText] = useState(range)
  const [error, setError] = useState(false)
  const apply = (): void => {
    if (!parseAreaText(text)) {
      setError(true)
      return
    }
    onClose()
    onApply(text.trim())
  }
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog resize-table-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('dlgResizeTableTitle')}
        onClick={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault()
            event.stopPropagation()
            onClose()
          } else if (event.key === 'Enter' && (event.target as HTMLElement).tagName !== 'BUTTON') {
            event.preventDefault()
            apply()
          }
        }}
      >
        <header>{t('dlgResizeTableTitle')}</header>
        <section className="dialog-body">
          <label className="resize-table-field">
            {t('dlgResizeTableRange')}
            <input
              type="text"
              autoFocus
              spellCheck={false}
              value={text}
              aria-invalid={error}
              onChange={(event) => {
                setText(event.target.value)
                setError(false)
              }}
            />
          </label>
          <p
            className={error ? 'dialog-note dialog-error' : 'dialog-note'}
            role={error ? 'alert' : undefined}
          >
            {error ? t('appTableResizeBadRange') : t('dlgResizeTableNote')}
          </p>
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" onClick={apply}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
