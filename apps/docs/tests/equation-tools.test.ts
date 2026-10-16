import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  insertEquationFromLatex,
  insertInlineEquationFromLatex,
} from '../src/renderer/components/EquationModal'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  appendToEquation,
  convertEquationToLinear,
  convertEquationToProfessional,
  equationSource,
  isLinearEquation,
  selectedEquation,
} from '../src/renderer/editor/equation-tools'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function open() {
  const parsed = await parseDocx(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>hello</w:t></w:r></w:p>' }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

function selectNodeOfType(editor: Editor, type: string) {
  let found = -1
  editor.state.doc.descendants((node, pos) => {
    if (found < 0 && node.type.name === type) found = pos
  })
  editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, found)))
}

describe('Equation Tools', () => {
  it('only acts on a selected equation node', async () => {
    const { editor } = await open()
    expect(selectedEquation(editor)).toBeNull()
    insertInlineEquationFromLatex(editor, '\\frac{a+b}{c}')
    expect(selectedEquation(editor)).toBeNull()
    selectNodeOfType(editor, 'docInlineMath')
    const eq = selectedEquation(editor)!
    expect(eq.kind).toBe('inline')
    expect(equationSource(eq)).toBe('\\frac{a+b}{c}')
    editor.destroy()
  })

  it('converts an inline equation to linear text and back to the same structure', async () => {
    const { editor } = await open()
    insertInlineEquationFromLatex(editor, '\\frac{a+b}{c}')
    selectNodeOfType(editor, 'docInlineMath')
    expect(convertEquationToLinear(editor, selectedEquation(editor)!)).toBe(true)
    selectNodeOfType(editor, 'docInlineMath')
    const linear = selectedEquation(editor)!
    expect(isLinearEquation(linear)).toBe(true)
    expect(linear.omml).toContain('(a+b)/c')
    expect(convertEquationToProfessional(editor, linear)).toBe(true)
    selectNodeOfType(editor, 'docInlineMath')
    expect(equationSource(selectedEquation(editor)!)).toBe('\\frac{a+b}{c}')
    editor.destroy()
  })

  it('keeps a linear display equation across save and reopen', async () => {
    const { editor, parsed } = await open()
    insertEquationFromLatex(editor, '{x}_{i}^{2}')
    selectNodeOfType(editor, 'docProtected')
    expect(convertEquationToLinear(editor, selectedEquation(editor)!)).toBe(true)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const formula = reparsed.blocks.find((b) => b.formulaDisplay)!
    expect(formula.formulaDisplay!.tokens.join('')).toBe('x_i^2')
    editor.destroy()
  })

  it('refuses to convert structures the linear form cannot express and leaves the equation alone', async () => {
    const { editor } = await open()
    insertInlineEquationFromLatex(editor, '\\mathbf{v}')
    selectNodeOfType(editor, 'docInlineMath')
    const before = editor.state.doc.toJSON()
    expect(convertEquationToLinear(editor, selectedEquation(editor)!)).toBe(false)
    expect(editor.state.doc.toJSON()).toEqual(before)
    editor.destroy()
  })

  it('appends gallery templates to a reversible equation', async () => {
    const { editor } = await open()
    insertInlineEquationFromLatex(editor, 'a+')
    selectNodeOfType(editor, 'docInlineMath')
    expect(appendToEquation(editor, selectedEquation(editor)!, '\\sqrt{{{selection}}}')).toBe(true)
    selectNodeOfType(editor, 'docInlineMath')
    expect(equationSource(selectedEquation(editor)!)).toBe('a+ \\sqrt{x}')
    editor.destroy()
  })

  it('does not touch an imported equation the source cannot express', async () => {
    const { editor } = await open()
    editor.commands.insertContent({
      type: 'docInlineMath',
      attrs: {
        omml: '<m:oMath><m:box><m:e><m:r><m:t>x</m:t></m:r></m:e></m:box></m:oMath>',
        mathml: '',
        latex: '',
        text: 'x',
      },
    })
    selectNodeOfType(editor, 'docInlineMath')
    const eq = selectedEquation(editor)!
    expect(equationSource(eq)).toBeNull()
    expect(appendToEquation(editor, eq, '\\sqrt{{{selection}}}')).toBe(false)
    expect(convertEquationToLinear(editor, eq)).toBe(false)
    editor.destroy()
  })
})
