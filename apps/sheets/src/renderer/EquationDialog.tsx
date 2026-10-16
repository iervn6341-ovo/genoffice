import { useMemo, useState } from 'react'

import { latexToOmml, ommlToMathML } from '@genoffice/office-math'
import { MathWorkbench, mathPreview } from '@genoffice/ui'
import { useI18n } from './i18n/locale'
import { useDialogEscape } from './use-dialog-escape'

export interface EquationDraft {
  readonly latex: string
  readonly omml: string
}

/// The dialog edits linear source, while the workbook drawing stores native
/// OMML in the a14:m choice. External OMML outside our reversible subset is
/// previewed but cannot be silently replaced with stale source text.
export function EquationDialog({
  initialLatex,
  originalOmml,
  onSave,
  onClose,
}: {
  readonly initialLatex?: string | null | undefined
  readonly originalOmml?: string | undefined
  readonly onSave: (equation: EquationDraft) => void
  readonly onClose: () => void
}): React.JSX.Element {
  useDialogEscape(onClose)
  const { t } = useI18n()
  const [latex, setLatex] = useState(initialLatex ?? '')
  const unsupported = initialLatex === null && originalOmml !== undefined
  const preview = useMemo(() => mathPreview(latex), [latex])
  const originalPreview = useMemo(
    () => (originalOmml ? ommlToMathML(`<m:oMath>${originalOmml}</m:oMath>`) : ''),
    [originalOmml],
  )
  const save = (): void => {
    if (unsupported || !preview || !('mathml' in preview)) return
    try {
      onSave({ latex, omml: latexToOmml(latex) })
      onClose()
    } catch {
      // The shared workbench keeps the invalid source and shows its error.
    }
  }
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog equation-dialog"
        role="dialog"
        aria-label={t('dlgEquationTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgEquationTitle')}</header>
        <section className="dialog-body">
          {unsupported ? (
            <>
              <p className="dialog-note">{t('dlgEquationUnsupportedEdit')}</p>
              <div
                className="go-math-preview"
                dangerouslySetInnerHTML={{ __html: originalPreview }}
              />
            </>
          ) : (
            <MathWorkbench
              latex={latex}
              onLatexChange={setLatex}
              autoFocus
              labels={{
                input: t('dlgEquationInput'),
                structures: t('dlgEquationStructures'),
                symbols: t('dlgSymbolTitle'),
                symbolSearch: t('dlgSymbolSearch'),
                symbolSubsetNames: {
                  math: t('dlgSymbolCatMath'),
                  greek: t('dlgSymbolCatGreek'),
                  arrows: t('dlgSymbolCatArrows'),
                },
                professional: t('dlgEquationProfessional'),
                linear: t('dlgEquationLinear'),
                preview: t('dlgEquationPreview'),
                empty: t('dlgEquationHint'),
                categoryNames: {
                  fraction: t('dlgEquationCatFraction'),
                  script: t('dlgEquationCatScript'),
                  radical: t('dlgEquationCatRadical'),
                  integral: t('dlgEquationCatIntegral'),
                  largeOperator: t('dlgEquationCatLargeOperator'),
                  delimiter: t('dlgEquationCatDelimiter'),
                  function: t('dlgEquationCatFunction'),
                  accent: t('dlgEquationCatAccent'),
                  limitLog: t('dlgEquationCatLimitLog'),
                  operator: t('dlgEquationCatOperator'),
                  matrix: t('dlgEquationCatMatrix'),
                },
              }}
            />
          )}
        </section>
        <div className="dialog-actions">
          <button className="secondary" onClick={onClose}>
            {t('dlgCancel')}
          </button>
          <button
            className="primary"
            disabled={unsupported || !preview || !('mathml' in preview)}
            onClick={save}
          >
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
