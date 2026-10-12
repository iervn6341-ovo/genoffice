import { useEffect, useMemo, useState } from 'react'

import {
  CUSTOM_ELEMENT_KEYS,
  EDGES_OF,
  customStyleDef,
  nextStyleName,
  styleNameProblem,
  type CustomElement,
  type CustomElementKey,
  type CustomTableStyle,
  type EdgeKey,
} from './custom-table-styles'
import { useI18n, type StringKey } from './i18n/locale'
import { previewCellStyle } from './table-style-gallery'
import { setCustomTableStyleDef, tableCellLook, type ThemeColors } from './table-styles'
import type { BorderLine } from './table-style-defs'
import { useDialogEscape } from './use-dialog-escape'

const ELEMENT_LABEL: Record<CustomElementKey, StringKey> = {
  wholeTable: 'dlgStyleElWholeTable',
  headerRow: 'dlgStyleElHeaderRow',
  totalRow: 'dlgStyleElTotalRow',
  firstColumn: 'dlgStyleElFirstColumn',
  lastColumn: 'dlgStyleElLastColumn',
  firstRowStripe: 'dlgStyleElFirstRowStripe',
  secondRowStripe: 'dlgStyleElSecondRowStripe',
  firstColumnStripe: 'dlgStyleElFirstColumnStripe',
  secondColumnStripe: 'dlgStyleElSecondColumnStripe',
}

const EDGE_LABEL: Record<EdgeKey, StringKey> = {
  top: 'dlgStyleEdgeTop',
  bottom: 'dlgStyleEdgeBottom',
  left: 'dlgStyleEdgeLeft',
  right: 'dlgStyleEdgeRight',
  insideH: 'dlgStyleEdgeInsideH',
  insideV: 'dlgStyleEdgeInsideV',
}

const LINES: readonly { value: BorderLine | 'none'; label: StringKey }[] = [
  { value: 'none', label: 'dlgStyleLineNone' },
  { value: 'thin', label: 'dlgStyleLineThin' },
  { value: 'medium', label: 'dlgStyleLineMedium' },
  { value: 'thick', label: 'dlgStyleLineThick' },
  { value: 'double', label: 'dlgStyleLineDouble' },
]

/// the preview draws every element: header and total rows, both stripes, first and last column
const PREVIEW_OPTIONS = {
  headerRow: true,
  totalRow: true,
  bandedRows: true,
  bandedColumns: true,
  firstColumn: true,
  lastColumn: true,
} as const
const PREVIEW_ROWS = 6
const PREVIEW_COLUMNS = 5
const PREVIEW_NAME = '\u0000new-table-style-preview'

const MAX_EDGES = 6

/// Excel's New Table Style dialog: a name and, for each table element, its fill, font and borders —
/// with a preview of the table they make. OK hands the style to the ribbon command (which keeps it
/// in the workbook and uses it); the name must not be one Excel or this workbook already has.
export function NewTableStyleDialog({
  taken,
  colors,
  onApply,
  onClose,
}: {
  readonly taken: readonly string[]
  readonly colors: ThemeColors
  readonly onApply: (style: CustomTableStyle) => void
  readonly onClose: () => void
}): React.JSX.Element {
  useDialogEscape(onClose)
  const { t } = useI18n()
  const [name, setName] = useState(() => nextStyleName(t('dlgNewStyleDefaultName'), taken))
  const [elements, setElements] = useState<Partial<Record<CustomElementKey, CustomElement>>>({
    headerRow: { fill: '#4472C4', font: '#FFFFFF', bold: true },
    firstRowStripe: { fill: '#D9E1F2' },
  })
  const [selected, setSelected] = useState<CustomElementKey>('headerRow')
  const [problem, setProblem] = useState<string | null>(null)

  const current: CustomElement = elements[selected] ?? {}
  const update = (patch: Partial<CustomElement>): void => {
    setElements((all) => ({ ...all, [selected]: { ...all[selected], ...patch } }))
    setProblem(null)
  }
  const edges = EDGES_OF[selected]
  const line = current.line ?? 'none'

  // the draft is registered under a name of its own while the dialog is open, so the preview draws it
  const draft: CustomTableStyle = useMemo(() => ({ name: PREVIEW_NAME, elements }), [elements])
  useEffect(() => {
    setCustomTableStyleDef(PREVIEW_NAME, customStyleDef(draft))
  }, [draft])
  useEffect(() => () => setCustomTableStyleDef(PREVIEW_NAME, null), [])
  const preview = useMemo(() => {
    const cells: React.CSSProperties[] = []
    for (let row = 0; row < PREVIEW_ROWS; row += 1) {
      for (let col = 0; col < PREVIEW_COLUMNS; col += 1) {
        const look = tableCellLook(
          PREVIEW_NAME,
          { rows: PREVIEW_ROWS, columns: PREVIEW_COLUMNS, row, col },
          colors,
          PREVIEW_OPTIONS,
        )
        cells.push({
          ...previewCellStyle(look),
          ...(look.color ? { color: look.color } : {}),
          ...(look.bold ? { fontWeight: 700 } : {}),
        })
      }
    }
    return cells
    // eslint-disable-next-line react-hooks/exhaustive-deps -- the draft's definition is what changes
  }, [draft, colors])

  const apply = (): void => {
    const trimmed = name.trim()
    const issue = styleNameProblem(trimmed, taken)
    if (issue) {
      setProblem(
        t(
          issue === 'empty'
            ? 'appTableStyleNameEmpty'
            : issue === 'invalid'
              ? 'appTableStyleNameInvalid'
              : issue === 'builtin'
                ? 'appTableStyleNameBuiltin'
                : 'appTableStyleNameTaken',
        ),
      )
      return
    }
    onClose()
    onApply({ name: trimmed, elements })
  }

  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog new-table-style-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={t('dlgNewStyleTitle')}
        onClick={(event) => event.stopPropagation()}
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault()
            event.stopPropagation()
            onClose()
          }
        }}
      >
        <header>{t('dlgNewStyleTitle')}</header>
        <section className="dialog-body">
          <label className="new-style-name">
            {t('dlgNewStyleName')}
            <input
              type="text"
              value={name}
              spellCheck={false}
              aria-invalid={problem !== null}
              onChange={(event) => {
                setName(event.target.value)
                setProblem(null)
              }}
            />
          </label>
          <div className="new-style-body">
            <div className="new-style-elements" role="listbox" aria-label={t('dlgNewStyleElement')}>
              {CUSTOM_ELEMENT_KEYS.map((key) => (
                <button
                  type="button"
                  key={key}
                  role="option"
                  aria-selected={key === selected}
                  className={key === selected ? 'selected' : ''}
                  onClick={() => setSelected(key)}
                >
                  {t(ELEMENT_LABEL[key])}
                </button>
              ))}
            </div>
            <div className="new-style-format">
              <div className="new-style-row">
                <label className="dialog-check">
                  <input
                    type="checkbox"
                    checked={current.fill !== undefined}
                    onChange={(event) =>
                      update({ fill: event.target.checked ? '#D9E1F2' : undefined })
                    }
                  />
                  {t('dlgStyleFill')}
                </label>
                <input
                  type="color"
                  aria-label={t('dlgStylePickColor', { what: t('dlgStyleFill') })}
                  disabled={current.fill === undefined}
                  value={current.fill ?? '#D9E1F2'}
                  onChange={(event) => update({ fill: event.target.value })}
                />
              </div>
              <div className="new-style-row">
                <label className="dialog-check">
                  <input
                    type="checkbox"
                    checked={current.font !== undefined}
                    onChange={(event) =>
                      update({ font: event.target.checked ? '#000000' : undefined })
                    }
                  />
                  {t('dlgStyleFont')}
                </label>
                <input
                  type="color"
                  aria-label={t('dlgStylePickColor', { what: t('dlgStyleFont') })}
                  disabled={current.font === undefined}
                  value={current.font ?? '#000000'}
                  onChange={(event) => update({ font: event.target.value })}
                />
              </div>
              <label className="dialog-check">
                <input
                  type="checkbox"
                  checked={current.bold === true}
                  onChange={(event) => update({ bold: event.target.checked || undefined })}
                />
                {t('dlgStyleBold')}
              </label>
              <div className="new-style-row">
                <label>
                  {t('dlgStyleLine')}
                  <select
                    value={line}
                    aria-label={t('dlgStyleLine')}
                    onChange={(event) => {
                      const next = event.target.value as BorderLine | 'none'
                      update(
                        next === 'none'
                          ? { line: undefined, edges: undefined }
                          : {
                              line: next,
                              edges: current.edges?.length ? current.edges : [edges[0]!],
                            },
                      )
                    }}
                  >
                    {LINES.map((option) => (
                      <option key={option.value} value={option.value}>
                        {t(option.label)}
                      </option>
                    ))}
                  </select>
                </label>
                <input
                  type="color"
                  aria-label={t('dlgStylePickColor', { what: t('dlgStyleLine') })}
                  disabled={line === 'none'}
                  value={current.lineColor ?? '#000000'}
                  onChange={(event) => update({ lineColor: event.target.value })}
                />
              </div>
              <div className="new-style-edges" role="group" aria-label={t('dlgStyleLine')}>
                {edges.slice(0, MAX_EDGES).map((edge) => (
                  <label key={edge} className="dialog-check">
                    <input
                      type="checkbox"
                      disabled={line === 'none'}
                      checked={current.edges?.includes(edge) === true}
                      onChange={(event) => {
                        const on = new Set(current.edges ?? [])
                        if (event.target.checked) on.add(edge)
                        else on.delete(edge)
                        update({ edges: [...on] })
                      }}
                    />
                    {t(EDGE_LABEL[edge])}
                  </label>
                ))}
              </div>
            </div>
          </div>
          <div className="new-style-preview-box">
            <span>{t('dlgNewStylePreview')}</span>
            <div
              className="new-style-preview"
              aria-hidden="true"
              style={{ gridTemplateColumns: `repeat(${PREVIEW_COLUMNS}, 1fr)` }}
            >
              {preview.map((cell, index) => (
                <i key={index} style={cell}>
                  {index < PREVIEW_COLUMNS ? 'Aa' : ''}
                </i>
              ))}
            </div>
          </div>
          {problem && (
            <p className="dialog-note dialog-error" role="alert">
              {problem}
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
