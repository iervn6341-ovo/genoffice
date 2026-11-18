import React, { useEffect } from 'react'
import type { ConflictInfo } from '../../shared/ipc'
import { useI18n } from '../i18n/locale'

export interface ConflictDialogProps {
  conflict: ConflictInfo
  onReload: () => void | Promise<void>
  onOverwrite: () => void | Promise<void>
  onSaveCopy: () => void | Promise<void>
  onClose: () => void
}

function formatBytes(bytes?: number): string {
  if (!bytes || bytes <= 0) return ''
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

export function ConflictDialog({
  conflict,
  onReload,
  onOverwrite,
  onSaveCopy,
  onClose,
}: ConflictDialogProps) {
  const { t } = useI18n()

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault()
        onClose()
      }
    }
    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [onClose])

  const formattedMtime = conflict.mtimeMs ? new Date(conflict.mtimeMs).toLocaleString() : null
  const formattedSize = formatBytes(conflict.size)

  return (
    <div
      className="modal-backdrop"
      onClick={onClose}
      role="dialog"
      aria-modal="true"
      data-testid="conflict-dialog-backdrop"
    >
      <div
        className="modal conflict-dialog"
        onClick={(e) => e.stopPropagation()}
        data-testid="conflict-dialog"
      >
        <div className="conflict-dialog-header">
          <div className="conflict-dialog-icon" aria-hidden="true">
            <svg
              width="24"
              height="24"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
            >
              <path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z" />
              <line x1="12" y1="9" x2="12" y2="13" />
              <line x1="12" y1="17" x2="12.01" y2="17" />
            </svg>
          </div>
          <div className="conflict-dialog-titles">
            <h2>{t('conflictTitle')}</h2>
            <p>{t('conflictSubtitle')}</p>
          </div>
        </div>

        <div className="conflict-cards">
          {/* Card 1: Local in-memory edits */}
          <div className="conflict-card conflict-card-local">
            <div className="conflict-card-header">
              <span className="conflict-card-title">{t('conflictLocalTitle')}</span>
              <span className="conflict-badge local">{t('conflictLocalBadge')}</span>
            </div>
            <p className="conflict-card-desc">{t('conflictLocalDesc')}</p>
            <div className="conflict-card-actions">
              <button
                type="button"
                className="btn btn-primary"
                onClick={() => void onOverwrite()}
                data-testid="conflict-btn-overwrite"
              >
                {t('conflictLocalOverwrite')}
              </button>
              <button
                type="button"
                className="btn btn-secondary"
                onClick={() => void onSaveCopy()}
                data-testid="conflict-btn-save-copy"
              >
                {t('conflictLocalSaveCopy')}
              </button>
            </div>
          </div>

          {/* Card 2: Disk external version */}
          <div className="conflict-card conflict-card-disk">
            <div className="conflict-card-header">
              <span className="conflict-card-title">{t('conflictDiskTitle')}</span>
              <span className="conflict-badge disk">{t('conflictDiskBadge')}</span>
            </div>
            <p className="conflict-card-desc">
              {conflict.missing ? t('conflictDiskMissing') : t('conflictDiskDesc')}
            </p>
            <div className="conflict-meta">
              <div className="conflict-meta-row">
                <span>{conflict.fileName}</span>
                {formattedSize && <span>{formattedSize}</span>}
              </div>
              {formattedMtime && (
                <div className="conflict-meta-row" style={{ marginTop: '4px' }}>
                  <span>{formattedMtime}</span>
                </div>
              )}
            </div>
            <div className="conflict-card-actions">
              {!conflict.missing ? (
                <button
                  type="button"
                  className="btn btn-primary"
                  onClick={() => void onReload()}
                  data-testid="conflict-btn-reload"
                >
                  {t('conflictDiskReload')}
                </button>
              ) : (
                <button type="button" className="btn btn-secondary" disabled>
                  {t('conflictDiskMissing')}
                </button>
              )}
            </div>
          </div>
        </div>

        <div className="conflict-footer">
          <button
            type="button"
            className="btn btn-secondary"
            onClick={onClose}
            data-testid="conflict-btn-later"
          >
            {t('conflictDecideLater')}
          </button>
        </div>
      </div>
    </div>
  )
}
