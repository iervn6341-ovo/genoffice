/**
 * Contextual "Equation" tab: shown while a native OMML equation shape is
 * selected. Structures and symbols append to the equation's source and commit
 * through `replaceEquation`, so the saved data stays native OMML.
 */
import { useState } from 'react'
import { insertMathTemplate, mathPreview } from '@genoffice/ui'
import { EQUATION_STRUCTURES, SYMBOL_SUBSETS } from '@genoffice/office-math'
import { useI18n, type StringKey } from '../i18n/locale'
import { IconEquation } from './icons'
import { BIG, Group, RbCaret } from './ribbon-shared'

const CATEGORY_LABEL: Record<string, StringKey> = {
  fraction: 'ribbonMathCatFraction',
  script: 'ribbonMathCatScript',
  radical: 'ribbonMathCatRadical',
  integral: 'ribbonMathCatIntegral',
  largeOperator: 'ribbonMathCatLargeOperator',
  delimiter: 'ribbonMathCatDelimiter',
  function: 'ribbonMathCatFunction',
  accent: 'ribbonMathCatAccent',
  limitLog: 'ribbonMathCatLimitLog',
  operator: 'ribbonMathCatOperator',
  matrix: 'ribbonMathCatMatrix',
}

const SYMBOL_SUBSET_LABEL: Record<string, StringKey> = {
  math: 'ribbonSymbolMath',
  greek: 'ribbonSymbolGreek',
  arrows: 'ribbonSymbolArrows',
}

const MATH_SUBSETS = SYMBOL_SUBSETS.filter((s) => s.id in SYMBOL_SUBSET_LABEL)

/** New LaTeX source after appending a structure template or a symbol to the current source. */
export function appendToEquationSource(source: string, addition: string): string {
  const end = source.length
  return insertMathTemplate(source, end, end, addition).latex
}

export function RibbonEquationTab({
  latex,
  editable,
  onEdit,
  onApply,
}: {
  latex: string
  /** false for imported OMML the strict converter cannot round-trip (kept read-only) */
  editable: boolean
  onEdit: () => void
  onApply: (latex: string) => void
}) {
  const { t } = useI18n()
  const [open, setOpen] = useState<string | null>(null)
  const apply = (addition: string) => {
    setOpen(null)
    onApply(appendToEquationSource(latex, addition))
  }
  const toggle = (id: string) => setOpen((cur) => (cur === id ? null : id))

  return (
    <>
      <Group label={t('ribbonMathTools')}>
        <button
          className="rb-big"
          onClick={onEdit}
          data-testid="equation-tab-edit"
          data-tip={t('ribbonEquationEdit')}
        >
          <span className="rb-big-icon">
            <IconEquation size={BIG} />
          </span>
          <span>{t('ribbonEquationEdit')}</span>
        </button>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonSymbol')}>
        <div className="rb-drop-wrap">
          <button
            className={`rb-big ${open === 'symbols' ? 'active' : ''}`}
            disabled={!editable}
            data-testid="equation-tab-symbols"
            onMouseDown={(e) => e.stopPropagation()}
            onClick={() => toggle('symbols')}
          >
            <span className="rb-big-icon">
              <span style={{ fontSize: 22 }}>π</span>
              <RbCaret />
            </span>
            <span>{t('ribbonSymbol')}</span>
          </button>
          {open === 'symbols' && (
            <div
              className="rb-drop"
              style={{ maxWidth: 360 }}
              onMouseDown={(e) => e.stopPropagation()}
            >
              {MATH_SUBSETS.map((subset) => (
                <div key={subset.id}>
                  <div className="rb-eq-drop-label">{t(SYMBOL_SUBSET_LABEL[subset.id]!)}</div>
                  <div
                    className="go-symbol-grid"
                    style={{ gridTemplateColumns: 'repeat(12, 1fr)' }}
                  >
                    {subset.symbols.slice(0, 48).map((char) => (
                      <button
                        key={char}
                        type="button"
                        data-testid="equation-tab-symbol"
                        onClick={() => apply(char)}
                      >
                        {char}
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </Group>
      <div className="ribbon-sep" />
      <Group label={t('ribbonMathStructures')}>
        <div className="rb-eq-structures">
          {EQUATION_STRUCTURES.map((cat) => (
            <div className="rb-drop-wrap" key={cat.id}>
              <button
                className={`rb-small ${open === cat.id ? 'active' : ''}`}
                disabled={!editable}
                data-testid={`equation-tab-${cat.id}`}
                onMouseDown={(e) => e.stopPropagation()}
                onClick={() => toggle(cat.id)}
              >
                <span>{t(CATEGORY_LABEL[cat.labelKey] ?? CATEGORY_LABEL[cat.id]!)}</span>
                <RbCaret />
              </button>
              {open === cat.id && (
                <div
                  className="rb-drop go-math-templates"
                  style={{ flexWrap: 'wrap', maxWidth: 360, maxHeight: 360, overflowY: 'auto' }}
                  onMouseDown={(e) => e.stopPropagation()}
                >
                  {cat.templates.map((template) => {
                    const sample = template.latex.replace('{{selection}}', 'x')
                    const result = mathPreview(sample)
                    if (!result || 'error' in result) return null
                    return (
                      <button
                        key={template.id}
                        type="button"
                        className="go-math-template"
                        title={sample}
                        aria-label={sample}
                        data-testid="equation-tab-template"
                        onClick={() => apply(template.latex)}
                      >
                        <span dangerouslySetInnerHTML={{ __html: result.mathml }} />
                      </button>
                    )
                  })}
                </div>
              )}
            </div>
          ))}
        </div>
      </Group>
    </>
  )
}
