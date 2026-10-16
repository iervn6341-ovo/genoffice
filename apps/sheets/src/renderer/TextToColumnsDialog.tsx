import { useMemo, useRef, useState } from 'react'

import { useI18n } from './i18n/locale'
import {
  DEFAULT_NUMBER_SETTINGS,
  parseTextToColumnsAddress,
  planTextToColumns,
  splitDelimitedCell,
  splitFixedWidthCell,
  textToColumnsAddress,
  textToColumnsDateOrder,
  TEXT_TO_COLUMNS_DATE_ORDERS,
  validNumberSettings,
  type TextToColumnsDateOrder,
  type TextToColumnsError,
  type TextToColumnsFormat,
  type TextToColumnsNumberSettings,
  type TextToColumnsOptions,
  type TextToColumnsSeed,
} from './text-to-columns'
import { useDialogEscape } from './use-dialog-escape'

export interface TextToColumnsChoice {
  readonly options: TextToColumnsOptions
  readonly destination: string
  readonly formats: readonly TextToColumnsFormat[]
}

const INITIAL_OPTIONS: TextToColumnsOptions = {
  tab: false,
  semicolon: false,
  comma: true,
  space: false,
  other: '',
  collapse: false,
  qualifier: 'double',
}

const MAX_FIXED_RULER = 128

function formatKind(format: TextToColumnsFormat): 'general' | 'text' | 'date' | 'skip' {
  return textToColumnsDateOrder(format) ? 'date' : (format as 'general' | 'text' | 'skip')
}

function dateOrderOf(format: TextToColumnsFormat | undefined): TextToColumnsDateOrder {
  return (format ? textToColumnsDateOrder(format) : null) ?? 'mdy'
}

export function TextToColumnsDialog({
  seed,
  onApply,
  onClose,
}: {
  readonly seed: TextToColumnsSeed
  /** `overwrite` is true once the user accepted Excel's "replace existing data?" prompt. */
  readonly onApply: (
    choice: TextToColumnsChoice,
    overwrite: boolean,
  ) => TextToColumnsError | 'sourceChanged' | 'confirmOverwrite' | null
  readonly onClose: () => void
}): React.JSX.Element {
  useDialogEscape(onClose)
  const { t } = useI18n()
  const [step, setStep] = useState<1 | 2 | 3>(1)
  const [mode, setMode] = useState<'delimited' | 'fixed'>('delimited')
  const [options, setOptions] = useState<TextToColumnsOptions>(INITIAL_OPTIONS)
  const [breaks, setBreaks] = useState<number[]>([])
  const [dragTarget, setDragTarget] = useState<number | null>(null)
  const draggingBreak = useRef<number | null>(null)
  const suppressClick = useRef(false)
  const rulerRef = useRef<HTMLDivElement>(null)
  const [destination, setDestination] = useState(
    textToColumnsAddress(seed.rows[0]?.row ?? 0, seed.sourceColumn),
  )
  const [formats, setFormats] = useState<TextToColumnsFormat[]>([])
  const [selectedColumn, setSelectedColumn] = useState(0)
  const [error, setError] = useState<TextToColumnsError | 'sourceChanged' | null>(null)
  const [numbers, setNumbers] = useState<TextToColumnsNumberSettings>(DEFAULT_NUMBER_SETTINGS)
  const [advancedOpen, setAdvancedOpen] = useState(false)
  const [confirmingOverwrite, setConfirmingOverwrite] = useState(false)
  const justAdded = useRef<number | null>(null)
  const effectiveOptions = useMemo<TextToColumnsOptions>(
    () =>
      mode === 'fixed' ? { ...options, fixedWidth: { breaks }, numbers } : { ...options, numbers },
    [breaks, mode, numbers, options],
  )
  const rulerWidth = Math.min(
    MAX_FIXED_RULER,
    Math.max(1, ...seed.rows.slice(0, 12).map((source) => [...source.text].length)),
  )
  const shownBreaks =
    dragTarget === null || draggingBreak.current === null
      ? breaks
      : [...breaks.filter((position) => position !== draggingBreak.current), dragTarget]
  const breakSet = new Set(shownBreaks)
  const parsed = useMemo(() => {
    const rows: string[][] = []
    let width = 0
    for (const source of seed.rows.slice(0, 12)) {
      const result = effectiveOptions.fixedWidth
        ? splitFixedWidthCell(source.text, effectiveOptions.fixedWidth.breaks)
        : splitDelimitedCell(source.text, effectiveOptions)
      if (result.error) return { rows: [], width: 0, error: result.error }
      rows.push(result.fields)
      width = Math.max(width, result.fields.length)
    }
    return { rows, width, error: null }
  }, [effectiveOptions, seed.rows])
  const checked = useMemo(() => {
    const address = parseTextToColumnsAddress(destination)
    if (!address) return { error: 'destination' as const }
    return planTextToColumns(seed, effectiveOptions, address, formats, () => 'free')
  }, [destination, effectiveOptions, formats, seed])
  const problem = error ?? (step > 1 ? (parsed.error ?? checked.error) : null)
  const errorLabel = (code: TextToColumnsError | 'sourceChanged'): string => {
    if (code === 'occupied') return t('dlgTtcOccupied')
    if (code === 'sourceChanged') return t('dlgTtcChanged')
    if (code === 'qualifier') return t('dlgTtcQualifierError')
    if (code === 'delimiter') return t('dlgTtcDelimiterError')
    if (code === 'breaks') return t('dlgTtcBreaksError')
    if (code === 'size') return t('dlgTtcSizeError')
    if (code === 'numbers') return t('dlgTtcNumbersError')
    if (code === 'source') return t('dlgTtcSourceError')
    return t('dlgTtcDestinationError')
  }
  const next = (): void => {
    setError(null)
    if (step === 1) setStep(2)
    else if (step === 2) {
      if (!parsed.error) setStep(3)
    } else {
      if (checked.error) {
        setError(checked.error)
        return
      }
      const result = onApply({ options: effectiveOptions, destination, formats }, false)
      if (result === 'confirmOverwrite') setConfirmingOverwrite(true)
      else if (result) setError(result)
      else onClose()
    }
  }
  const setOption = <Key extends keyof TextToColumnsOptions>(
    key: Key,
    value: TextToColumnsOptions[Key],
  ): void => {
    setError(null)
    setOptions((current) => ({ ...current, [key]: value }))
  }
  const confirmOverwrite = (): void => {
    const result = onApply({ options: effectiveOptions, destination, formats }, true)
    setConfirmingOverwrite(false)
    if (result) setError(result === 'confirmOverwrite' ? 'occupied' : result)
    else onClose()
  }
  const addBreak = (position: number): void => {
    setError(null)
    setBreaks((current) =>
      current.includes(position)
        ? current
        : [...current, position].sort((left, right) => left - right),
    )
  }
  const removeBreak = (position: number): void => {
    setError(null)
    setBreaks((current) => current.filter((candidate) => candidate !== position))
  }
  const setFormat = (column: number, format: TextToColumnsFormat): void =>
    setFormats((current) => {
      const next = [...current]
      next[column] = format
      return next
    })
  const positionAt = (clientX: number): number => {
    const first = rulerRef.current?.querySelector('button')?.getBoundingClientRect()
    if (!first) return 1
    return Math.max(1, Math.min(rulerWidth, Math.floor((clientX - first.left) / first.width) + 1))
  }
  const fixedRuler = (
    <div className="text-columns-fixed" role="group" aria-label={t('dlgTtcFixedWidth')}>
      <p className="dialog-note">{t('dlgTtcBreaksHint')}</p>
      <div className="text-columns-fixed-scroll">
        <div className="text-columns-fixed-ruler" ref={rulerRef}>
          {Array.from({ length: rulerWidth }, (_, index) => {
            const position = index + 1
            const active = breakSet.has(position)
            return (
              <button
                key={position}
                type="button"
                className={active ? 'active' : ''}
                data-position={position}
                aria-pressed={active}
                aria-label={`${t(active ? 'dlgTtcRemoveBreak' : 'dlgTtcAddBreak')} ${position}`}
                title={`${t(active ? 'dlgTtcRemoveBreak' : 'dlgTtcAddBreak')} ${position}`}
                onClick={(event) => {
                  if (suppressClick.current) {
                    suppressClick.current = false
                    return
                  }
                  if (!breaks.includes(position)) {
                    // Excel: a click on empty ruler adds a break line
                    justAdded.current = position
                    window.setTimeout(() => {
                      if (justAdded.current === position) justAdded.current = null
                    }, 600)
                    addBreak(position)
                  } else if (event.detail === 0) removeBreak(position) // keyboard Enter / Space
                }}
                onDoubleClick={() => {
                  // Excel: double-click removes an existing break line (but not the one
                  // that the first click of this same gesture just created)
                  if (justAdded.current === position) {
                    justAdded.current = null
                    return
                  }
                  if (breaks.includes(position)) removeBreak(position)
                }}
                onPointerDown={(event) => {
                  if (!breaks.includes(position)) return
                  draggingBreak.current = position
                  event.currentTarget.setPointerCapture(event.pointerId)
                }}
                onPointerMove={(event) => {
                  if (draggingBreak.current !== null) setDragTarget(positionAt(event.clientX))
                }}
                onPointerUp={(event) => {
                  const from = draggingBreak.current
                  if (from === null) return
                  const to = positionAt(event.clientX)
                  draggingBreak.current = null
                  setDragTarget(null)
                  if (to !== from) {
                    if (!breaks.includes(to)) {
                      setBreaks((current) =>
                        [...current.filter((candidate) => candidate !== from), to].sort(
                          (left, right) => left - right,
                        ),
                      )
                    }
                    suppressClick.current = true
                    window.setTimeout(() => {
                      suppressClick.current = false
                    }, 0)
                  }
                }}
                onPointerCancel={() => {
                  draggingBreak.current = null
                  setDragTarget(null)
                }}
              >
                {active ? '│' : position % 5 === 0 ? '·' : ' '}
              </button>
            )
          })}
        </div>
        {seed.rows.slice(0, 8).map((source) => {
          const characters = [...source.text]
          return (
            <div className="text-columns-fixed-line" key={source.row}>
              {Array.from({ length: rulerWidth }, (_, index) => (
                <span className={breakSet.has(index + 1) ? 'break' : ''} key={index}>
                  {characters[index] ?? ' '}
                </span>
              ))}
            </div>
          )
        })}
      </div>
    </div>
  )
  const formatLabel = (format: TextToColumnsFormat): string => {
    const order = textToColumnsDateOrder(format)
    if (order) return `${t('dlgTtcDate')} ${order.toUpperCase()}`
    if (format === 'text') return t('dlgFcNumText')
    if (format === 'skip') return t('dlgTtcSkip')
    return t('dlgFcNumGeneral')
  }
  const preview = (
    <div className="text-columns-preview" role="region" aria-label={t('dlgEquationPreview')}>
      <table>
        <thead>
          <tr>
            {Array.from({ length: Math.min(parsed.width, 16) }, (_, index) => (
              <th key={index} className={formats[index] === 'skip' ? 'skipped' : ''}>
                <button
                  type="button"
                  className={selectedColumn === index ? 'selected' : ''}
                  aria-label={`${t('dlgTtcColumnFormat')} ${index + 1}`}
                  onClick={() => setSelectedColumn(index)}
                >
                  {index + 1} · {formatLabel(formats[index] ?? 'general')}
                </button>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {parsed.rows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {Array.from({ length: Math.min(parsed.width, 16) }, (_, column) => (
                <td key={column} className={formats[column] === 'skip' ? 'skipped' : ''}>
                  {row[column] ?? ''}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog text-columns-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('appTextToColumns')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>
          {t('appTextToColumns')} · {step} / 3
        </header>
        <div className="dialog-body">
          {step === 1 && (
            <>
              <label className="dialog-radio">
                <input
                  type="radio"
                  name="text-to-columns-mode"
                  checked={mode === 'delimited'}
                  onChange={() => {
                    setMode('delimited')
                    setFormats([])
                    setSelectedColumn(0)
                    setError(null)
                  }}
                />
                {t('dlgTtcDelimited')}
              </label>
              <label className="dialog-radio">
                <input
                  type="radio"
                  name="text-to-columns-mode"
                  checked={mode === 'fixed'}
                  onChange={() => {
                    setMode('fixed')
                    setFormats([])
                    setSelectedColumn(0)
                    setError(null)
                  }}
                />
                {t('dlgTtcFixedWidth')}
              </label>
              <p className="dialog-note">{t('dlgEquationPreview')}</p>
              <div className="text-columns-source-preview">
                {seed.rows.slice(0, 12).map((row) => (
                  <div key={row.row}>{row.text}</div>
                ))}
              </div>
            </>
          )}
          {step === 2 && (
            <>
              {mode === 'fixed' ? (
                fixedRuler
              ) : (
                <>
                  <fieldset>
                    <legend>{t('dlgTtcDelimited')}</legend>
                    {(
                      [
                        ['tab', 'appSplitByTab'],
                        ['semicolon', 'appSplitBySemicolon'],
                        ['comma', 'appSplitByComma'],
                        ['space', 'appSplitBySpace'],
                      ] as const
                    ).map(([key, label]) => (
                      <label className="dialog-check" key={key}>
                        <input
                          type="checkbox"
                          checked={options[key]}
                          onChange={(event) => setOption(key, event.target.checked)}
                        />
                        {t(label)}
                      </label>
                    ))}
                    <label className="dialog-check">
                      {t('dlgFnCatOther')}
                      <input
                        aria-label={t('dlgFnCatOther')}
                        value={options.other}
                        maxLength={2}
                        onChange={(event) => setOption('other', event.target.value)}
                      />
                    </label>
                  </fieldset>
                  <label className="dialog-check">
                    <input
                      type="checkbox"
                      checked={options.collapse}
                      onChange={(event) => setOption('collapse', event.target.checked)}
                    />
                    {t('dlgTtcConsecutive')}
                  </label>
                  <label>
                    {t('dlgTtcQualifier')}
                    <select
                      aria-label={t('dlgTtcQualifier')}
                      value={options.qualifier}
                      onChange={(event) =>
                        setOption(
                          'qualifier',
                          event.target.value as TextToColumnsOptions['qualifier'],
                        )
                      }
                    >
                      <option value="double">&quot;</option>
                      <option value="none">{t('dlgFcUnderlineNone')}</option>
                    </select>
                  </label>
                </>
              )}
              {preview}
            </>
          )}
          {step === 3 && (
            <>
              <label>
                {t('dlgPivotTargetCell')}
                <input
                  aria-label={t('dlgPivotTargetCell')}
                  value={destination}
                  onChange={(event) => {
                    setError(null)
                    setDestination(event.target.value)
                  }}
                />
              </label>
              <fieldset className="text-columns-format">
                <legend>
                  {t('dlgTtcColumnFormat')} {selectedColumn + 1}
                </legend>
                {(
                  [
                    ['general', 'dlgFcNumGeneral'],
                    ['text', 'dlgFcNumText'],
                    ['date', 'dlgTtcDate'],
                    ['skip', 'dlgTtcSkip'],
                  ] as const
                ).map(([kind, label]) => (
                  <label className="dialog-radio" key={kind}>
                    <input
                      type="radio"
                      name="text-to-columns-format"
                      checked={formatKind(formats[selectedColumn] ?? 'general') === kind}
                      onChange={() =>
                        setFormat(
                          selectedColumn,
                          kind === 'date'
                            ? `date-${dateOrderOf(formats[selectedColumn])}`
                            : (kind as TextToColumnsFormat),
                        )
                      }
                    />
                    {t(label)}
                  </label>
                ))}
                <select
                  aria-label={t('dlgTtcDateOrder')}
                  disabled={formatKind(formats[selectedColumn] ?? 'general') !== 'date'}
                  value={dateOrderOf(formats[selectedColumn])}
                  onChange={(event) =>
                    setFormat(
                      selectedColumn,
                      `date-${event.target.value as TextToColumnsDateOrder}`,
                    )
                  }
                >
                  {TEXT_TO_COLUMNS_DATE_ORDERS.map((order) => (
                    <option key={order} value={order}>
                      {order.toUpperCase()}
                    </option>
                  ))}
                </select>
              </fieldset>
              <div>
                <button
                  type="button"
                  aria-expanded={advancedOpen}
                  onClick={() => setAdvancedOpen((open) => !open)}
                >
                  {t('dlgTtcAdvanced')}
                </button>
              </div>
              {advancedOpen && (
                <fieldset className="text-columns-advanced">
                  <legend>{t('dlgTtcAdvanced')}</legend>
                  <label>
                    {t('dlgTtcDecimalSeparator')}
                    <input
                      aria-label={t('dlgTtcDecimalSeparator')}
                      value={numbers.decimal}
                      maxLength={1}
                      onChange={(event) => {
                        setError(null)
                        setNumbers((current) => ({ ...current, decimal: event.target.value }))
                      }}
                    />
                  </label>
                  <label>
                    {t('dlgTtcThousandsSeparator')}
                    <input
                      aria-label={t('dlgTtcThousandsSeparator')}
                      value={numbers.thousands}
                      maxLength={1}
                      onChange={(event) => {
                        setError(null)
                        setNumbers((current) => ({ ...current, thousands: event.target.value }))
                      }}
                    />
                  </label>
                  <label className="dialog-check">
                    <input
                      type="checkbox"
                      checked={numbers.trailingMinus}
                      onChange={(event) =>
                        setNumbers((current) => ({
                          ...current,
                          trailingMinus: event.target.checked,
                        }))
                      }
                    />
                    {t('dlgTtcTrailingMinus')}
                  </label>
                  {!validNumberSettings(numbers) && <p role="alert">{t('dlgTtcNumbersError')}</p>}
                </fieldset>
              )}
              {confirmingOverwrite && (
                <div className="text-columns-confirm" role="alertdialog" aria-live="assertive">
                  <p>{t('dlgTtcOverwritePrompt')}</p>
                  <button type="button" className="primary" onClick={confirmOverwrite}>
                    {t('appOk')}
                  </button>
                  <button type="button" onClick={() => setConfirmingOverwrite(false)}>
                    {t('appCancel')}
                  </button>
                </div>
              )}
              {preview}
            </>
          )}
          {problem && <p role="alert">{errorLabel(problem)}</p>}
        </div>
        <footer className="dialog-actions">
          <button type="button" onClick={onClose}>
            {t('appCancel')}
          </button>
          {step > 1 && (
            <button type="button" onClick={() => setStep(step === 3 ? 2 : 1)}>
              {t('appPrevious')}
            </button>
          )}
          <button
            type="button"
            className="primary"
            disabled={
              step === 2 ? Boolean(parsed.error) : step === 3 ? Boolean(checked.error) : false
            }
            onClick={next}
          >
            {step === 3 ? t('dlgTtcFinish') : t('appNext')}
          </button>
        </footer>
      </div>
    </div>
  )
}
