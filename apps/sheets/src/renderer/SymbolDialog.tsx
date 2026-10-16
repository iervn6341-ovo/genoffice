import { useRef, useState } from 'react'
import { SymbolPicker } from '@genoffice/ui'
import { useI18n } from './i18n/locale'
import {
  formulaSymbolSourceProblem,
  insertIntoFormulaSource,
  type FormulaSymbolDraft,
} from './symbol-formula'
import { useDialogEscape } from './use-dialog-escape'

/// Appends a picked symbol to the active cell's text. Numbers and booleans
/// become text, matching Excel's Symbol insertion behavior.
export function appendSymbol(
  existing: string | number | boolean | null | undefined,
  char: string,
): string {
  if (existing === null || existing === undefined || existing === '') return char
  return `${String(existing)}${char}`
}

export function SymbolDialog({
  onInsert,
  onClose,
  formulaDraft,
  onApplyFormula,
}: {
  readonly onInsert: (char: string, font: string) => void
  readonly onClose: () => void
  readonly formulaDraft?: FormulaSymbolDraft | null
  readonly onApplyFormula?: (source: string) => string | null
}): React.JSX.Element {
  useDialogEscape(onClose)
  const { t } = useI18n()
  const [formulaSource, setFormulaSource] = useState(formulaDraft?.source ?? '')
  const [formulaError, setFormulaError] = useState<string | null>(null)
  const formulaRef = useRef<HTMLTextAreaElement>(null)
  const hasFormulaFocusRef = useRef(false)
  const selectionRef = useRef({
    start: formulaDraft?.selectionStart ?? 0,
    end: formulaDraft?.selectionEnd ?? 0,
  })

  function pick(char: string, font: string): void {
    if (!formulaDraft) {
      onInsert(char, font)
      return
    }
    const field = formulaRef.current
    const selection =
      field && hasFormulaFocusRef.current
        ? { start: field.selectionStart, end: field.selectionEnd }
        : selectionRef.current
    const inserted = insertIntoFormulaSource(formulaSource, selection.start, selection.end, char)
    setFormulaSource(inserted.source)
    setFormulaError(null)
    selectionRef.current = { start: inserted.caret, end: inserted.caret }
    requestAnimationFrame(() => {
      field?.focus()
      field?.setSelectionRange(inserted.caret, inserted.caret)
    })
  }

  function applyFormula(): void {
    const problem = formulaSymbolSourceProblem(formulaSource)
    if (problem) {
      setFormulaError(
        t(
          problem === 'start'
            ? 'appFormulaStartsEquals'
            : problem === 'parentheses'
              ? 'appUnbalancedParens'
              : 'appSymbolInsertFailed',
        ),
      )
      return
    }
    const error = onApplyFormula?.(formulaSource)
    if (error) setFormulaError(error)
    else onClose()
  }
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog symbol-dialog"
        role="dialog"
        aria-label={t('dlgSymbolTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgSymbolTitle')}</header>
        <section className="dialog-body">
          {formulaDraft && (
            <label className="symbol-formula-source">
              {t('dlgFnFormula')}
              <textarea
                ref={formulaRef}
                value={formulaSource}
                onChange={(event) => {
                  setFormulaSource(event.target.value)
                  setFormulaError(null)
                }}
                onSelect={(event) => {
                  selectionRef.current = {
                    start: event.currentTarget.selectionStart,
                    end: event.currentTarget.selectionEnd,
                  }
                }}
                onFocus={(event) => {
                  hasFormulaFocusRef.current = true
                  event.currentTarget.setSelectionRange(
                    selectionRef.current.start,
                    selectionRef.current.end,
                  )
                }}
                aria-label={t('dlgFnFormula')}
                spellCheck={false}
              />
            </label>
          )}
          <SymbolPicker
            onPick={pick}
            labels={{
              font: t('dlgSymbolFont'),
              subset: t('dlgSymbolSubset'),
              search: t('dlgSymbolSearch'),
              codePoint: t('dlgSymbolCodePoint'),
              recent: t('dlgSymbolRecent'),
              insert: t('dlgSymbolInsert'),
              empty: t('dlgSymbolEmpty'),
              invalidCodePoint: t('dlgSymbolInvalidCodePoint'),
              subsetNames: {
                latin: t('dlgSymbolSubsetLatin'),
                greek: t('dlgSymbolCatGreek'),
                arrows: t('dlgSymbolCatArrows'),
                math: t('dlgSymbolCatMath'),
                currency: t('dlgSymbolCatCurrency'),
                misc: t('dlgSymbolCatMisc'),
              },
            }}
          />
          {formulaError && (
            <p className="dialog-note" role="alert">
              {formulaError}
            </p>
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t(formulaDraft ? 'dlgCancel' : 'dlgClose')}
          </button>
          {formulaDraft && <button onClick={applyFormula}>{t('dlgApply')}</button>}
        </div>
      </div>
    </div>
  )
}
