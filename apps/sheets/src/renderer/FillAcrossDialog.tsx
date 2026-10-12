import { useState } from 'react'

import type { FillAcrossParams, FillAcrossWhat } from './fill-across'
import { useI18n, type StringKey } from './i18n/locale'
import { useDialogEscape } from './use-dialog-escape'

const WHATS: readonly { value: FillAcrossWhat; label: StringKey }[] = [
  { value: 'all', label: 'dlgAcrossAll' },
  { value: 'contents', label: 'dlgAcrossContents' },
  { value: 'formats', label: 'dlgAcrossFormats' },
]

/// Excel's Fill Across Worksheets: which sheets, and what to fill (All, Contents or Formats). Excel
/// fills the grouped tabs; the tabs are not grouped here, so the sheets are ticked in the dialog.
export function FillAcrossDialog({
  sheets,
  onApply,
  onClose,
}: {
  readonly sheets: readonly { id: string; name: string }[]
  readonly onApply: (params: FillAcrossParams) => void
  readonly onClose: () => void
}): React.JSX.Element {
  useDialogEscape(onClose)
  const { t } = useI18n()
  const [picked, setPicked] = useState<ReadonlySet<string>>(new Set())
  const [what, setWhat] = useState<FillAcrossWhat>('all')
  const apply = (): void => {
    if (picked.size === 0) return
    onClose()
    onApply({ sheetIds: [...picked], what })
  }
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog fill-across-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('dlgAcrossTitle')}
        onClick={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault()
            event.stopPropagation()
            onClose()
          }
        }}
      >
        <header>{t('dlgAcrossTitle')}</header>
        <section className="dialog-body">
          <div className="fill-across-groups">
            <fieldset aria-label={t('dlgAcrossSheets')}>
              <legend>{t('dlgAcrossSheets')}</legend>
              {sheets.length === 0 && <p className="dialog-note">{t('appFillAcrossNoSheets')}</p>}
              {sheets.map((sheet) => (
                <label key={sheet.id} className="dialog-check">
                  <input
                    type="checkbox"
                    checked={picked.has(sheet.id)}
                    onChange={(event) => {
                      const next = new Set(picked)
                      if (event.target.checked) next.add(sheet.id)
                      else next.delete(sheet.id)
                      setPicked(next)
                    }}
                  />
                  {sheet.name}
                </label>
              ))}
            </fieldset>
            <fieldset role="radiogroup" aria-label={t('dlgAcrossFill')}>
              <legend>{t('dlgAcrossFill')}</legend>
              {WHATS.map((option) => (
                <label key={option.value} className="dialog-radio">
                  <input
                    type="radio"
                    name="across-what"
                    value={option.value}
                    checked={what === option.value}
                    onChange={() => setWhat(option.value)}
                  />
                  {t(option.label)}
                </label>
              ))}
            </fieldset>
          </div>
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" onClick={apply} disabled={picked.size === 0}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
