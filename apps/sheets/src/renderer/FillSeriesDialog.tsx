import { useState } from 'react'

import type { FillSeriesParams } from './fill-actions'
import type { SeriesAxis, SeriesDateUnit } from './fill-series'
import { useI18n, type StringKey } from './i18n/locale'
import { useDialogEscape } from './use-dialog-escape'

export interface FillSeriesSeed {
  readonly axis: SeriesAxis
  /// the first cell holds a date: the dialog opens on Date, like Excel
  readonly startIsDate: boolean
}

const TYPES: readonly { value: FillSeriesParams['type']; label: StringKey }[] = [
  { value: 'linear', label: 'dlgSeriesLinear' },
  { value: 'growth', label: 'dlgSeriesGrowth' },
  { value: 'date', label: 'dlgSeriesDate' },
  { value: 'autofill', label: 'dlgSeriesAutoFill' },
]

const UNITS: readonly { value: SeriesDateUnit; label: StringKey }[] = [
  { value: 'day', label: 'dlgSeriesDay' },
  { value: 'weekday', label: 'dlgSeriesWeekday' },
  { value: 'month', label: 'dlgSeriesMonth' },
  { value: 'year', label: 'dlgSeriesYear' },
]

/// Excel's Fill → Series… dialog. Step and Stop are read as typed; an empty Stop means none. OK hands
/// the choice to the ribbon command, which plans and writes the series (fill-series.ts).
export function FillSeriesDialog({
  seed,
  onApply,
  onClose,
}: {
  readonly seed: FillSeriesSeed
  readonly onApply: (params: FillSeriesParams) => void
  readonly onClose: () => void
}): React.JSX.Element {
  useDialogEscape(onClose)
  const { t } = useI18n()
  const [axis, setAxis] = useState<SeriesAxis>(seed.axis)
  const [type, setType] = useState<FillSeriesParams['type']>(seed.startIsDate ? 'date' : 'linear')
  const [dateUnit, setDateUnit] = useState<SeriesDateUnit>('day')
  const [trend, setTrend] = useState(false)
  const [step, setStep] = useState('1')
  const [stop, setStop] = useState('')
  const [error, setError] = useState<string | null>(null)
  const canTrend = type === 'linear' || type === 'growth'
  // a Trend ticked under Linear must not lock Step / Stop once the type is Date
  const stepless = type === 'autofill' || (canTrend && trend)

  const apply = (): void => {
    const stepValue = Number(step.trim().replace(',', '.'))
    const stopText = stop.trim().replace(',', '.')
    const stopValue = stopText === '' ? null : Number(stopText)
    if (
      (!stepless && (step.trim() === '' || !Number.isFinite(stepValue))) ||
      (stopValue !== null && !Number.isFinite(stopValue))
    ) {
      setError(t('dlgSeriesInvalid'))
      return
    }
    onClose()
    onApply({
      axis,
      type,
      dateUnit,
      step: stepless ? 1 : stepValue,
      stop: stepless ? null : stopValue,
      trend: canTrend && trend,
    })
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog fill-series-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('dlgSeriesTitle')}
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
        <header>{t('dlgSeriesTitle')}</header>
        <section className="dialog-body">
          <div className="fill-series-groups">
            <fieldset role="radiogroup" aria-label={t('dlgSeriesIn')}>
              <legend>{t('dlgSeriesIn')}</legend>
              {(['rows', 'columns'] as const).map((value) => (
                <label key={value} className="dialog-radio">
                  <input
                    type="radio"
                    name="series-axis"
                    value={value}
                    checked={axis === value}
                    onChange={() => setAxis(value)}
                  />
                  {t(value === 'rows' ? 'dlgSeriesRows' : 'dlgSeriesColumns')}
                </label>
              ))}
            </fieldset>
            <fieldset role="radiogroup" aria-label={t('dlgSeriesType')}>
              <legend>{t('dlgSeriesType')}</legend>
              {TYPES.map((option) => (
                <label key={option.value} className="dialog-radio">
                  <input
                    type="radio"
                    name="series-type"
                    value={option.value}
                    checked={type === option.value}
                    onChange={() => setType(option.value)}
                  />
                  {t(option.label)}
                </label>
              ))}
            </fieldset>
            <fieldset
              role="radiogroup"
              aria-label={t('dlgSeriesDateUnit')}
              disabled={type !== 'date'}
            >
              <legend>{t('dlgSeriesDateUnit')}</legend>
              {UNITS.map((option) => (
                <label key={option.value} className="dialog-radio">
                  <input
                    type="radio"
                    name="series-unit"
                    value={option.value}
                    checked={dateUnit === option.value}
                    onChange={() => setDateUnit(option.value)}
                  />
                  {t(option.label)}
                </label>
              ))}
            </fieldset>
          </div>
          <label className="dialog-check fill-series-trend">
            <input
              type="checkbox"
              checked={canTrend && trend}
              disabled={!canTrend}
              onChange={(event) => setTrend(event.target.checked)}
            />
            {t('dlgSeriesTrend')}
          </label>
          <div className="fill-series-values">
            <label>
              {t('dlgSeriesStep')}
              <input
                value={step}
                disabled={stepless}
                inputMode="decimal"
                aria-label={t('dlgSeriesStep')}
                onChange={(event) => {
                  setStep(event.target.value)
                  setError(null)
                }}
              />
            </label>
            <label>
              {t('dlgSeriesStop')}
              <input
                value={stop}
                disabled={stepless}
                inputMode="decimal"
                aria-label={t('dlgSeriesStop')}
                onChange={(event) => {
                  setStop(event.target.value)
                  setError(null)
                }}
              />
            </label>
          </div>
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
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
