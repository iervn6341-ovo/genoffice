import { useState } from 'react'

import { useI18n, type StringKey } from './i18n/locale'
import {
  DEFAULT_PASTE_SPECIAL,
  type PasteOperation,
  type PasteSpecialOptions,
  type PasteWhat,
} from './paste-special'

const WHATS: readonly { value: PasteWhat; label: StringKey }[] = [
  { value: 'all', label: 'dlgPasteAll' },
  { value: 'formulas', label: 'dlgPasteFormulas' },
  { value: 'values', label: 'dlgPasteValues' },
  { value: 'formats', label: 'dlgPasteFormats' },
  { value: 'comments', label: 'dlgPasteComments' },
  { value: 'validation', label: 'dlgPasteValidation' },
  { value: 'allUsingSourceTheme', label: 'dlgPasteAllSourceTheme' },
  { value: 'allExceptBorders', label: 'dlgPasteAllExceptBorders' },
  { value: 'columnWidths', label: 'dlgPasteColumnWidths' },
  { value: 'formulasAndNumberFormats', label: 'dlgPasteFormulasNumFmt' },
  { value: 'valuesAndNumberFormats', label: 'dlgPasteValuesNumFmt' },
  { value: 'allMergingConditionalFormats', label: 'dlgPasteAllMergeCf' },
]

const OPERATIONS: readonly { value: PasteOperation; label: StringKey }[] = [
  { value: 'none', label: 'dlgOpNone' },
  { value: 'add', label: 'dlgOpAdd' },
  { value: 'subtract', label: 'dlgOpSubtract' },
  { value: 'multiply', label: 'dlgOpMultiply' },
  { value: 'divide', label: 'dlgOpDivide' },
]

/// Excel's Paste Special dialog: what to paste, an operation against the cells underneath, Skip
/// blanks and Transpose. Column widths takes no operation, Skip blanks or Transpose (nothing to
/// combine), and the operation applies only to values — Formats ignores it, as Excel does.
export function PasteSpecialDialog({
  onApply,
  onClose,
}: {
  readonly onApply: (options: PasteSpecialOptions) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [options, setOptions] = useState<PasteSpecialOptions>(DEFAULT_PASTE_SPECIAL)
  const widthsOnly = options.what === 'columnWidths'
  const noOperation = widthsOnly || options.what === 'comments' || options.what === 'validation'
  const apply = (): void => {
    onClose()
    onApply(
      widthsOnly
        ? { ...options, operation: 'none', skipBlanks: false, transpose: false }
        : noOperation
          ? { ...options, operation: 'none', skipBlanks: false }
          : options,
    )
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog paste-special-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('dlgPasteSpecialTitle')}
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
        <header>{t('dlgPasteSpecialTitle')}</header>
        <section className="dialog-body">
          <div className="paste-special-groups">
            <fieldset role="radiogroup" aria-label={t('dlgPasteSpecialPaste')}>
              <legend>{t('dlgPasteSpecialPaste')}</legend>
              {WHATS.map((option) => (
                <label key={option.value} className="dialog-radio">
                  <input
                    type="radio"
                    name="paste-what"
                    value={option.value}
                    checked={options.what === option.value}
                    autoFocus={option.value === 'all'}
                    onChange={() => setOptions({ ...options, what: option.value })}
                  />
                  {t(option.label)}
                </label>
              ))}
            </fieldset>
            <fieldset
              role="radiogroup"
              aria-label={t('dlgPasteSpecialOperation')}
              disabled={noOperation}
            >
              <legend>{t('dlgPasteSpecialOperation')}</legend>
              {OPERATIONS.map((option) => (
                <label key={option.value} className="dialog-radio">
                  <input
                    type="radio"
                    name="paste-operation"
                    value={option.value}
                    checked={options.operation === option.value}
                    onChange={() => setOptions({ ...options, operation: option.value })}
                  />
                  {t(option.label)}
                </label>
              ))}
            </fieldset>
          </div>
          <div className="paste-special-checks">
            <label className="dialog-check">
              <input
                type="checkbox"
                checked={options.skipBlanks && !noOperation}
                disabled={noOperation}
                onChange={(event) => setOptions({ ...options, skipBlanks: event.target.checked })}
              />
              {t('dlgPasteSkipBlanks')}
            </label>
            <label className="dialog-check">
              <input
                type="checkbox"
                checked={options.transpose && !widthsOnly}
                disabled={widthsOnly}
                onChange={(event) => setOptions({ ...options, transpose: event.target.checked })}
              />
              {t('dlgPasteTranspose')}
            </label>
          </div>
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
