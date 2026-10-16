import { useEffect, useRef, useState } from 'react'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

/**
 * In-app replacement for window.prompt, which Electron renderers do not
 * support. Single-line by default; multiline renders a textarea (notes).
 */
export function PromptModal({
  title,
  placeholder,
  initial = '',
  multiline = false,
  allowEmpty = false,
  secondary,
  onSubmit,
  onClose,
}: {
  title: string
  placeholder?: string
  initial?: string
  multiline?: boolean
  /** submit button stays enabled for empty input (e.g. clearing a value) */
  allowEmpty?: boolean
  secondary?: { label: string; initial?: string; hint?: string }
  onSubmit: (value: string, secondary?: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [value, setValue] = useState(initial)
  const [secondaryValue, setSecondaryValue] = useState(secondary?.initial ?? '')
  const inputRef = useRef<HTMLInputElement>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)
  const modalKeys = useModalKeys(onClose)

  useEffect(() => {
    const el = multiline ? textareaRef.current : inputRef.current
    el?.focus()
    el?.select()
  }, [multiline])

  const submit = () => {
    if ((!allowEmpty && !value.trim()) || !validSecondary) return
    onSubmit(value.trim(), secondary ? secondaryValue.trim() : undefined)
    onClose()
  }
  const validSecondary =
    Array.from(secondaryValue).length <= 8 &&
    Array.from(secondaryValue).every((char) => {
      const cp = char.codePointAt(0) ?? 0
      return cp >= 0x20 && cp !== 0x7f && !(cp >= 0xd800 && cp <= 0xdfff)
    })

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{title}</h2>
        <label>
          {multiline ? (
            <textarea
              ref={textareaRef}
              className="prompt-textarea"
              rows={4}
              value={value}
              placeholder={placeholder}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit()
              }}
            />
          ) : (
            <input
              ref={inputRef}
              value={value}
              placeholder={placeholder}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') submit()
              }}
            />
          )}
        </label>
        {secondary && (
          <label className="prompt-secondary">
            {secondary.label}
            <input
              aria-label={secondary.label}
              value={secondaryValue}
              maxLength={16}
              onChange={(event) => setSecondaryValue(event.target.value)}
            />
            {secondary.hint && <small>{secondary.hint}</small>}
            <span className="note-mark-presets">
              {(['♠', '♥'] as const).map((symbol) => (
                <button
                  key={symbol}
                  type="button"
                  aria-label={symbol}
                  onClick={() => setSecondaryValue(symbol)}
                >
                  {symbol}
                </button>
              ))}
            </span>
          </label>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button
            className="btn-primary"
            disabled={(!allowEmpty && !value.trim()) || !validSecondary}
            onClick={submit}
          >
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
