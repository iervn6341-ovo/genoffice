import { useMemo, useState } from 'react'
import type { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { ommlToLatex } from '@genoffice/docx-engine'
import { MathWorkbench, mathPreview, type MathWorkbenchLabels } from '@genoffice/ui'
import { useI18n, type StringKey } from '../i18n/locale'
import { equationBlockJson, inlineEquationNodeJson } from '../editor/equation'
import { useModalKeys } from './modal-keys'

/** Built-in equation library (the common preset set for the "Insert Equation" dropdown) */
export const BUILTIN_EQUATIONS: Array<{ name: string; nameKey: StringKey; latex: string }> = [
  { name: 'Area of Circle', nameKey: 'appEqCircleArea', latex: 'A = \\pi r^2' },
  {
    name: 'Binomial Theorem',
    nameKey: 'appEqBinomial',
    latex: '(x+a)^n = \\sum_{k=0}^{n} \\binom{n}{k} x^k a^{n-k}',
  },
  {
    name: 'Expansion of a Sum',
    nameKey: 'appEqSumExpansion',
    latex: '(1+x)^n = 1 + \\frac{nx}{1!} + \\frac{n(n-1)x^2}{2!} + \\cdots',
  },
  {
    name: 'Fourier Series',
    nameKey: 'appEqFourier',
    latex:
      'f(x) = a_0 + \\sum_{n=1}^{\\infty} \\left( a_n \\cos \\frac{n\\pi x}{L} + b_n \\sin \\frac{n\\pi x}{L} \\right)',
  },
  { name: 'Pythagorean Theorem', nameKey: 'appEqPythagorean', latex: 'a^2 + b^2 = c^2' },
  {
    name: 'Quadratic Formula',
    nameKey: 'appEqQuadratic',
    latex: 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}',
  },
  {
    name: 'Taylor Expansion',
    nameKey: 'appEqTaylor',
    latex:
      'e^x = 1 + \\frac{x}{1!} + \\frac{x^2}{2!} + \\frac{x^3}{3!} + \\cdots, -\\infty < x < \\infty',
  },
  {
    name: 'Trig Identity',
    nameKey: 'appEqTrigIdentity',
    latex:
      '\\sin \\alpha \\pm \\sin \\beta = 2 \\sin \\frac{1}{2} (\\alpha \\pm \\beta) \\cos \\frac{1}{2} (\\alpha \\mp \\beta)',
  },
]

/** Build an equation block from LaTeX and insert it (protected block, genXml goes through the save path) */
export function insertEquationFromLatex(editor: Editor, latex: string): void {
  editor.chain().focus().insertContent(equationBlockJson(latex)).run()
}

/** Insert an inline equation at the cursor that flows with the text */
export function insertInlineEquationFromLatex(editor: Editor, latex: string): void {
  editor.chain().focus().insertContent(inlineEquationNodeJson(latex)).run()
}

/** Equation re-edit target (inline = inline node; block = standalone protected equation block) */
export interface MathEditTarget {
  pos: number
  latex: string
  kind: 'inline' | 'block'
}

export function currentEquationSource(editor: Editor, target: MathEditTarget): string | null {
  const node = editor.state.doc.nodeAt(target.pos)
  if (target.kind === 'inline' && node?.type.name === 'docInlineMath') {
    return ommlToLatex(String(node.attrs.omml ?? ''))
  }
  if (target.kind === 'block' && node?.type.name === 'docProtected') {
    const formula = node.attrs.formulaDisplay as { omml?: string } | null
    return formula?.omml ? ommlToLatex(formula.omml) : null
  }
  return null
}

/** Locate the equation selected in the document before ribbon focus moves away. */
export function selectedEquationTarget(editor: Editor): MathEditTarget | null {
  const selection = editor.state.selection
  const candidates: Array<{ pos: number; kind: MathEditTarget['kind'] }> = []
  if (selection instanceof NodeSelection && selection.node.type.name === 'docProtected') {
    candidates.push({ pos: selection.from, kind: 'block' })
  }
  if (selection.$from.nodeAfter?.type.name === 'docInlineMath') {
    candidates.push({ pos: selection.from, kind: 'inline' })
  }
  if (selection.$from.nodeBefore?.type.name === 'docInlineMath') {
    candidates.push({ pos: selection.from - selection.$from.nodeBefore.nodeSize, kind: 'inline' })
  }
  for (const candidate of candidates) {
    const target = { ...candidate, latex: '' }
    const latex = currentEquationSource(editor, target)
    if (latex !== null) return { ...target, latex }
  }
  return null
}

/** Replace a reversible equation; leave unsupported original OMML intact. */
export function updateEquationFromLatex(
  editor: Editor,
  target: MathEditTarget,
  latex: string,
): boolean {
  if (currentEquationSource(editor, target) === null) return false
  const node = editor.state.doc.nodeAt(target.pos)
  if (target.kind === 'inline' && node?.type.name === 'docInlineMath') {
    const next = inlineEquationNodeJson(latex)
    editor.view.dispatch(editor.state.tr.setNodeMarkup(target.pos, undefined, next.attrs))
    return true
  }
  if (target.kind === 'block' && node?.type.name === 'docProtected') {
    // A full re-edit generates new OMML and a fresh block anchor.
    const next = editor.schema.nodeFromJSON(equationBlockJson(latex))
    editor.view.dispatch(editor.state.tr.replaceWith(target.pos, target.pos + node.nodeSize, next))
    return true
  }
  return false
}

/** Insert Equation dialog: LaTeX input + live MathML preview, shared by the ribbon and the native menu */
export function EquationModal({
  editor,
  onClose,
  editTarget,
}: {
  editor: Editor
  onClose: () => void
  /** When provided, "re-edit equation" mode: update that node instead of inserting */
  editTarget?: MathEditTarget
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [latex, setLatex] = useState(() =>
    editTarget ? (currentEquationSource(editor, editTarget) ?? '') : '',
  )
  const [inline, setInline] = useState(editTarget?.kind === 'inline')
  const preview = useMemo(() => mathPreview(latex), [latex])
  const canInsert =
    preview !== null &&
    'mathml' in preview &&
    (!editTarget || currentEquationSource(editor, editTarget) !== null)
  const labels: MathWorkbenchLabels = {
    input: t('appMathSource'),
    structures: t('appMathStructures'),
    professional: t('appMathProfessional'),
    linear: t('appMathLinear'),
    preview: t('appMathPreview'),
    empty: t('appLatexPreviewHint'),
    symbols: t('appSymbolMath'),
    inputFormats: { latex: 'LaTeX', unicodeMath: 'UnicodeMath' },
    symbolSearch: t('appSymbolSearch'),
    symbolSubsetNames: {
      math: t('appSymbolMath'),
      greek: t('appSymbolGreek'),
      arrows: t('appSymbolArrows'),
    },
    categoryNames: {
      fraction: t('appMathFraction'),
      script: t('appMathScript'),
      radical: t('appMathRadical'),
      integral: t('appMathIntegral'),
      largeOperator: t('appMathLargeOperator'),
      delimiter: t('appMathDelimiter'),
      function: t('appMathFunction'),
      accent: t('appMathAccent'),
      limitLog: t('appMathLimitLog'),
      operator: t('appMathOperator'),
      matrix: t('appMathMatrix'),
    },
  }

  const insert = () => {
    if (!canInsert || !editor.isEditable) return
    if (editTarget) {
      if (!updateEquationFromLatex(editor, editTarget, latex)) return
    } else if (inline) {
      insertInlineEquationFromLatex(editor, latex)
    } else {
      insertEquationFromLatex(editor, latex)
    }
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        className="modal equation-modal"
        onKeyDown={(event) => {
          if ((event.metaKey || event.ctrlKey) && event.key === 'Enter') insert()
        }}
      >
        <h2>{t('appInsertEquation')}</h2>
        <MathWorkbench latex={latex} onLatexChange={setLatex} labels={labels} autoFocus />
        {!editTarget && (
          <label className="equation-inline-toggle">
            <input type="checkbox" checked={inline} onChange={(e) => setInline(e.target.checked)} />
            {t('appInlineEquationToggle')}
          </label>
        )}
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" disabled={!canInsert} onClick={insert}>
            {editTarget ? t('appUpdate') : t('appInsert')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** Built-in equation gallery (ribbon Equation ▾ dropdown), with MathML thumbnail previews */
export function EquationGallery({
  editor,
  onPick,
  onCustom,
}: {
  editor: Editor
  onPick: () => void
  onCustom: () => void
}) {
  const { t } = useI18n()
  return (
    <div data-rb-panel="" className="equation-gallery">
      {BUILTIN_EQUATIONS.map((eq) => {
        const preview = mathPreview(eq.latex)
        return (
          <button
            key={eq.nameKey}
            className="equation-gallery-item"
            data-tip={t(eq.nameKey)}
            onClick={() => {
              insertEquationFromLatex(editor, eq.latex)
              onPick()
            }}
          >
            <span className="equation-gallery-name">{t(eq.nameKey)}</span>
            {preview !== null && 'mathml' in preview && (
              <span
                className="equation-gallery-preview"
                dangerouslySetInnerHTML={{ __html: preview.mathml }}
              />
            )}
          </button>
        )
      })}
      <button className="equation-gallery-custom" onClick={onCustom}>
        {t('appInsertNewEquation')}
      </button>
    </div>
  )
}
