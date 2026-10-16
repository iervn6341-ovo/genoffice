import type { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import {
  ommlToLatex,
  ommlToUnicodeMath,
  unicodeMathToOmml,
  latexToOmml,
} from '@genoffice/office-math'
import {
  equationBlockFromInner,
  equationBlockJson,
  inlineEquationNodeFromInner,
  inlineEquationNodeJson,
} from './equation'

/** The equation node the Equation contextual tab acts on (a node selection only, like Word's caret-in-equation). */
export interface SelectedEquation {
  kind: 'inline' | 'block'
  pos: number
  /** full <m:oMath> fragment */
  omml: string
}

export function selectedEquation(editor: Editor): SelectedEquation | null {
  const selection = editor.state.selection
  if (!(selection instanceof NodeSelection)) return null
  const node = selection.node
  if (node.type.name === 'docInlineMath') {
    const omml = String(node.attrs.omml ?? '')
    return omml ? { kind: 'inline', pos: selection.from, omml } : null
  }
  if (node.type.name === 'docProtected') {
    const formula = node.attrs.formulaDisplay as { omml?: string } | null | undefined
    return formula?.omml ? { kind: 'block', pos: selection.from, omml: formula.omml } : null
  }
  return null
}

/** LaTeX source when the structure is fully expressible; null keeps the Office content untouched. */
export const equationSource = (equation: SelectedEquation): string | null =>
  ommlToLatex(equation.omml)

const LINEAR_MARKS = /[\^_/√∛∜▒■█▭┬┴¦⊘∕]/

/** innerText of an equation that is a single plain run (Word's "linear" display), else null */
export function linearTextOf(omml: string): string | null {
  const body = omml.replace(/^<m:oMath[^>]*>/, '').replace(/<\/m:oMath>$/, '')
  const m =
    /^<m:r>(?:<m:rPr>(?:(?!<\/m:rPr>)[\s\S])*<\/m:rPr>)?(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<m:t[^>]*>([\s\S]*?)<\/m:t><\/m:r>$/.exec(
      body,
    )
  if (!m) return null
  return m[1]!
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&')
}

export const isLinearEquation = (equation: SelectedEquation): boolean => {
  const text = linearTextOf(equation.omml)
  return text !== null && LINEAR_MARKS.test(text)
}

function replaceEquation(
  editor: Editor,
  equation: SelectedEquation,
  inner: string,
  latex: string,
): boolean {
  const node = editor.state.doc.nodeAt(equation.pos)
  if (!node) return false
  if (equation.kind === 'inline') {
    const next = inlineEquationNodeFromInner(inner, latex)
    editor.view.dispatch(editor.state.tr.setNodeMarkup(equation.pos, undefined, next.attrs))
    return true
  }
  const next = editor.schema.nodeFromJSON(equationBlockFromInner(inner, latex))
  const tr = editor.state.tr.replaceWith(equation.pos, equation.pos + node.nodeSize, next)
  editor.view.dispatch(tr)
  return true
}

const escapeText = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')

/** Professional -> Linear: the equation becomes one run of UnicodeMath text (valid native OMML Word re-builds). */
export function convertEquationToLinear(editor: Editor, equation: SelectedEquation): boolean {
  const text = ommlToUnicodeMath(equation.omml)
  if (text === null || text === '') return false
  const inner = `<m:r><m:t xml:space="preserve">${escapeText(text)}</m:t></m:r>`
  return replaceEquation(editor, equation, inner, '')
}

/** Linear -> Professional: parse the UnicodeMath run into structures. */
export function convertEquationToProfessional(editor: Editor, equation: SelectedEquation): boolean {
  const text = linearTextOf(equation.omml)
  if (text === null) return false
  let inner: string
  try {
    inner = unicodeMathToOmml(text)
  } catch {
    return false
  }
  return replaceEquation(editor, equation, inner, ommlToLatex(`<m:oMath>${inner}</m:oMath>`) ?? '')
}

/** Append a gallery template or symbol (with the `{{selection}}` slot filled by `x`) to the equation's source. */
export function appendToEquation(
  editor: Editor,
  equation: SelectedEquation,
  template: string,
): boolean {
  const source = equationSource(equation)
  if (source === null) return false
  const piece = template.replace('{{selection}}', 'x')
  const latex = `${source}${source && !/\s$/.test(source) ? ' ' : ''}${piece}`.trim()
  try {
    latexToOmml(latex)
  } catch {
    return false
  }
  const node = editor.state.doc.nodeAt(equation.pos)
  if (!node) return false
  if (equation.kind === 'inline') {
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(equation.pos, undefined, inlineEquationNodeJson(latex).attrs),
    )
  } else {
    const next = editor.schema.nodeFromJSON(equationBlockJson(latex))
    editor.view.dispatch(
      editor.state.tr.replaceWith(equation.pos, equation.pos + node.nodeSize, next),
    )
  }
  return true
}
