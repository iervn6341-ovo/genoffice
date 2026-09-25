import { useState } from 'react'

import { cellShiftOptions, type CellShiftMode, type CellShiftOption } from './cell-shift'
import { useI18n, type StringKey } from './i18n/locale'

const OPTION_LABELS: Record<CellShiftOption, StringKey> = {
  'shift-right': 'dlgCellsShiftRight',
  'shift-down': 'dlgCellsShiftDown',
  'shift-left': 'dlgCellsShiftLeft',
  'shift-up': 'dlgCellsShiftUp',
  row: 'dlgCellsEntireRow',
  column: 'dlgCellsEntireColumn',
}

/// Excel's Insert (⌃⇧= / ⌘⇧+) and Delete (⌃- / ⌘-) cells dialog: four radios,
/// OK applies the choice to the selection, Esc or Cancel closes.
export function InsertCellsDialog({
  mode,
  initial,
  onApply,
  onClose,
}: {
  readonly mode: CellShiftMode
  readonly initial: CellShiftOption
  readonly onApply: (option: CellShiftOption) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [option, setOption] = useState<CellShiftOption>(initial)
  const title = t(mode === 'insert' ? 'dlgInsertCellsTitle' : 'dlgDeleteCellsTitle')
  const apply = (): void => {
    onClose()
    onApply(option)
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog insert-cells-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={title}
        onClick={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault()
            event.stopPropagation()
            onClose()
          } else if (event.key === 'Enter') {
            event.preventDefault()
            apply()
          }
        }}
      >
        <header>{title}</header>
        <section className="dialog-body">
          <fieldset className="dialog-span" role="radiogroup" aria-label={title}>
            {cellShiftOptions(mode).map((value) => (
              <label key={value} className="dialog-radio">
                <input
                  type="radio"
                  name="cell-shift"
                  value={value}
                  checked={option === value}
                  autoFocus={option === value}
                  onChange={() => setOption(value)}
                />
                {t(OPTION_LABELS[value])}
              </label>
            ))}
          </fieldset>
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
