import {
  latexToOmml,
  mathParagraphXml,
  mathTokensOf,
  ommlToLatex,
  ommlToMathML,
} from '@genoffice/docx-engine'
import { t } from '../i18n/locale'
import type { PmNode } from './convert'

/**
 * Protected display-equation block built from LaTeX (throws on syntax outside
 * the supported subset). Shared by the insert dialog/gallery and the AI
 * <formula> tag; genXml carries the OMML through the save path.
 */
/** MathML for inline flow (ommlToMathML emits display="block" for equations) */
export function inlineMathML(omml: string): string {
  return ommlToMathML(omml).replace(/ display="block"/g, ' display="inline"')
}

/**
 * Atomic inline formula node built from LaTeX (throws on unsupported syntax).
 * Flows with paragraph text; saves as a Run.math (<m:oMath> emitted verbatim).
 */
export function inlineEquationNodeJson(latex: string): PmNode {
  return inlineEquationNodeFromInner(latexToOmml(latex), latex.trim())
}

/** Inline equation node from the children of <m:oMath>; `latex` is the editable source ('' when none). */
export function inlineEquationNodeFromInner(inner: string, latex: string): PmNode {
  const omml = `<m:oMath>${inner}</m:oMath>`
  return {
    type: 'docInlineMath',
    attrs: {
      omml,
      mathml: inlineMathML(omml),
      // recovered LaTeX makes parsed formulas re-editable; null = atom only
      latex: latex || (ommlToLatex(omml) ?? ''),
      text: mathTokensOf(omml).join(''),
    },
  }
}

export function equationBlockJson(latex: string): PmNode {
  return equationBlockFromInner(latexToOmml(latex), latex.trim())
}

/** Display equation block from the children of <m:oMath>. */
export function equationBlockFromInner(inner: string, latex: string): PmNode {
  const omml = inner
  return {
    type: 'docProtected',
    attrs: {
      docxIndex: null,
      blockType: 'passthrough',
      label: t('editorEquation'),
      previewText: latex,
      genXml: mathParagraphXml(omml),
      formulaDisplay: {
        tokens: mathTokensOf(omml),
        mathml: ommlToMathML(`<m:oMath>${omml}</m:oMath>`),
        omml: `<m:oMath>${omml}</m:oMath>`,
        latex,
      },
    },
  }
}
