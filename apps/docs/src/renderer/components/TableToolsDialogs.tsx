import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import {
  canSortTable,
  convertTableToText,
  convertTextToTable,
  sortTableRows,
  type SortKey,
  type SortKind,
  type TextSeparator,
} from '../editor/table-tools'
import {
  caretGridPosition,
  defaultFormulaFor,
  formulaResultText,
  insertTableFormula,
  tableFormulaGrid,
  updateTableFormulas,
} from '../editor/table-formula'

type ModalProps = { editor: Editor; onClose: () => void }

const FORMULA_PICTURES = ['', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', '$#,##0.00']

function Shell({
  title,
  onClose,
  children,
  actions,
}: {
  title: string
  onClose: () => void
  children: React.ReactNode
  actions: React.ReactNode
}) {
  const keys = useModalKeys(onClose)
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(event) => event.target === event.currentTarget && onClose()}
    >
      <div className="modal gs-form table-tools-dialog" role="dialog" aria-label={title}>
        <h2>{title}</h2>
        {children}
        <div className="modal-actions">{actions}</div>
      </div>
    </div>
  )
}

/** Table Layout > Sort: up to three keys, type, direction and header row. */
export function TableSortDialog({ editor, onClose }: ModalProps) {
  const { t } = useI18n()
  const columns = useMemo(() => {
    let count = 0
    editor.state.doc.descendants((node) => {
      if (node.type.name === 'docTable' && count === 0) {
        count = node.firstChild?.childCount ?? 0
        return false
      }
      return true
    })
    return count
  }, [editor])
  const sortable = canSortTable(editor.state)
  const [header, setHeader] = useState(true)
  const [keys, setKeys] = useState<
    Array<{ column: number | null; kind: SortKind; descending: boolean }>
  >([
    { column: 0, kind: 'text', descending: false },
    { column: null, kind: 'text', descending: false },
    { column: null, kind: 'text', descending: false },
  ])
  const update = (index: number, patch: Partial<(typeof keys)[number]>) =>
    setKeys((previous) => previous.map((key, at) => (at === index ? { ...key, ...patch } : key)))
  const apply = () => {
    const chosen: SortKey[] = keys.flatMap((key) =>
      key.column === null
        ? []
        : [{ column: key.column, kind: key.kind, descending: key.descending }],
    )
    if (chosen.length === 0) return
    sortTableRows({ keys: chosen, header })(editor.state, editor.view.dispatch)
    onClose()
  }
  return (
    <Shell
      title={t('tableSortTitle')}
      onClose={onClose}
      actions={
        <>
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!sortable} onClick={apply}>
            {t('ribbonOk')}
          </button>
        </>
      }
    >
      {!sortable && <p className="modal-desc">{t('tableSortMerged')}</p>}
      {keys.map((key, index) => (
        <fieldset key={index} className="table-sort-key" disabled={!sortable}>
          <label>
            {index === 0 ? t('tableSortBy') : t('tableSortThenBy')}
            <select
              value={key.column ?? ''}
              onChange={(event) =>
                update(index, {
                  column: event.target.value === '' ? null : Number(event.target.value),
                })
              }
            >
              {index > 0 && <option value="">{t('tableSortNone')}</option>}
              {Array.from({ length: columns }, (_, column) => (
                <option key={column} value={column}>
                  {t('tableSortColumn', { n: column + 1 })}
                </option>
              ))}
            </select>
          </label>
          <label>
            {t('tableSortType')}
            <select
              value={key.kind}
              disabled={key.column === null}
              onChange={(event) => update(index, { kind: event.target.value as SortKind })}
            >
              <option value="text">{t('tableSortText')}</option>
              <option value="number">{t('tableSortNumber')}</option>
              <option value="date">{t('tableSortDate')}</option>
            </select>
          </label>
          <div className="table-tools-radios">
            <label>
              <input
                type="radio"
                name={`sort-dir-${index}`}
                disabled={key.column === null}
                checked={!key.descending}
                onChange={() => update(index, { descending: false })}
              />
              {t('tableSortAscending')}
            </label>
            <label>
              <input
                type="radio"
                name={`sort-dir-${index}`}
                disabled={key.column === null}
                checked={key.descending}
                onChange={() => update(index, { descending: true })}
              />
              {t('tableSortDescending')}
            </label>
          </div>
        </fieldset>
      ))}
      <label className="table-tools-check">
        <input
          type="checkbox"
          checked={header}
          onChange={(event) => setHeader(event.target.checked)}
        />
        {t('tableSortHeader')}
      </label>
    </Shell>
  )
}

/** Table Layout > Formula. */
export function TableFormulaDialog({ editor, onClose }: ModalProps) {
  const { t } = useI18n()
  const here = caretGridPosition(editor.state)
  const initial = here
    ? defaultFormulaFor(tableFormulaGrid(here.table), here.row, here.column)
    : '=SUM(ABOVE)'
  const [formula, setFormula] = useState(initial)
  const [picture, setPicture] = useState('')
  let preview = ''
  let error = ''
  if (here) {
    try {
      preview = formulaResultText(
        formula,
        picture || null,
        tableFormulaGrid(here.table),
        here.row,
        here.column,
      )
    } catch (caught) {
      error = caught instanceof Error ? caught.message : String(caught)
    }
  }
  const apply = () => {
    if (error) return
    insertTableFormula(formula, picture || null)(editor.state, editor.view.dispatch)
    onClose()
  }
  return (
    <Shell
      title={t('tableFormulaTitle')}
      onClose={onClose}
      actions={
        <>
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!here || !!error} onClick={apply}>
            {t('ribbonOk')}
          </button>
        </>
      }
    >
      <label>
        {t('tableFormulaLabel')}
        <input
          value={formula}
          spellCheck={false}
          onChange={(event) => setFormula(event.target.value)}
          onKeyDown={(event) => event.key === 'Enter' && apply()}
        />
      </label>
      <label>
        {t('tableFormulaNumberFormat')}
        <input
          list="table-formula-pictures"
          value={picture}
          spellCheck={false}
          onChange={(event) => setPicture(event.target.value)}
        />
        <datalist id="table-formula-pictures">
          {FORMULA_PICTURES.filter(Boolean).map((value) => (
            <option key={value} value={value} />
          ))}
        </datalist>
      </label>
      <p className="modal-desc" role="status">
        {error ? t('tableFormulaInvalid', { reason: error }) : `= ${preview}`}
      </p>
      <p className="modal-desc">{t('tableFormulaHelp')}</p>
      <div className="table-tools-inline-actions">
        <button
          className="btn-ghost"
          onClick={() => {
            updateTableFormulas()(editor.state, editor.view.dispatch)
            onClose()
          }}
        >
          {t('ribbonTableUpdateFormulas')}
        </button>
      </div>
    </Shell>
  )
}

type SeparatorChoice = 'paragraph' | 'tab' | 'comma' | 'other'

function SeparatorPicker({
  value,
  onChange,
  other,
  onOther,
}: {
  value: SeparatorChoice
  onChange: (value: SeparatorChoice) => void
  other: string
  onOther: (value: string) => void
}) {
  const { t } = useI18n()
  const options: Array<[SeparatorChoice, string]> = [
    ['paragraph', t('tableSepParagraph')],
    ['tab', t('tableSepTab')],
    ['comma', t('tableSepComma')],
    ['other', t('tableSepOther')],
  ]
  return (
    <div className="table-tools-radios" role="radiogroup" aria-label={t('tableSeparateTextWith')}>
      {options.map(([choice, label]) => (
        <label key={choice}>
          <input
            type="radio"
            name="table-text-separator"
            checked={value === choice}
            onChange={() => onChange(choice)}
          />
          {label}
          {choice === 'other' && (
            <input
              className="table-tools-other"
              maxLength={4}
              value={other}
              onFocus={() => onChange('other')}
              onChange={(event) => onOther(event.target.value)}
            />
          )}
        </label>
      ))}
    </div>
  )
}

const separatorOf = (choice: SeparatorChoice, other: string): TextSeparator | null =>
  choice === 'other' ? (other ? { custom: other } : null) : choice

/** Table Layout > Convert to Text. */
export function TableToText({ editor, onClose }: ModalProps) {
  const { t } = useI18n()
  const [choice, setChoice] = useState<SeparatorChoice>('tab')
  const [other, setOther] = useState('')
  const [failed, setFailed] = useState(false)
  const apply = () => {
    const separator = separatorOf(choice, other)
    if (!separator) return
    const done = convertTableToText(separator)(editor.state, editor.view.dispatch)
    if (done) onClose()
    else setFailed(true)
  }
  return (
    <Shell
      title={t('tableConvertToTextTitle')}
      onClose={onClose}
      actions={
        <>
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={apply}>
            {t('ribbonOk')}
          </button>
        </>
      }
    >
      <p className="modal-desc">{t('tableSeparateTextWith')}</p>
      <SeparatorPicker value={choice} onChange={setChoice} other={other} onOther={setOther} />
      {failed && (
        <p className="modal-desc" role="alert">
          {t('tableConvertNested')}
        </p>
      )}
    </Shell>
  )
}

/** Insert > Table > Convert Text to Table. */
export function TextToTableDialog({ editor, onClose }: ModalProps) {
  const { t } = useI18n()
  const [choice, setChoice] = useState<SeparatorChoice>('tab')
  const [other, setOther] = useState('')
  const [columns, setColumns] = useState(2)
  const [failed, setFailed] = useState(false)
  const apply = () => {
    const separator = separatorOf(choice, other)
    if (!separator) return
    const done = convertTextToTable({
      separator,
      columns: choice === 'paragraph' ? columns : undefined,
    })(editor.state, editor.view.dispatch)
    if (done) onClose()
    else setFailed(true)
  }
  return (
    <Shell
      title={t('tableConvertToTableTitle')}
      onClose={onClose}
      actions={
        <>
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={apply}>
            {t('ribbonOk')}
          </button>
        </>
      }
    >
      <p className="modal-desc">{t('tableSeparateTextWith')}</p>
      <SeparatorPicker value={choice} onChange={setChoice} other={other} onOther={setOther} />
      {choice === 'paragraph' && (
        <label>
          {t('tableColumnsCount')}
          <input
            type="number"
            min={1}
            max={63}
            value={columns}
            onChange={(event) =>
              setColumns(Math.max(1, Math.min(63, Number(event.target.value) || 1)))
            }
          />
        </label>
      )}
      {failed && (
        <p className="modal-desc" role="alert">
          {t('tableConvertSelectParas')}
        </p>
      )}
    </Shell>
  )
}
