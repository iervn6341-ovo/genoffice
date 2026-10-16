import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  BUILTIN_EQUATIONS,
  currentEquationSource,
  insertEquationFromLatex,
  insertInlineEquationFromLatex,
  selectedEquationTarget,
  updateEquationFromLatex,
} from '../src/renderer/components/EquationModal'
import { insertSymbolAtSelection } from '../src/renderer/components/SymbolDialog'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

async function openDoc(): Promise<{
  editor: Editor
  parsed: Awaited<ReturnType<typeof parseDocx>>
}> {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>hello</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { editor, parsed }
}

describe('equation insertion', () => {
  it('inserts a display equation that renders native MathML in the editor', async () => {
    const { editor } = await openDoc()
    insertEquationFromLatex(editor, 'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}')

    const host = editor.view.dom.querySelector('.doc-formula-math')
    expect(host).not.toBeNull()
    expect(host!.innerHTML).toContain('<mfrac>')
    expect(host!.innerHTML).toContain('<msqrt>')
    // the editable token strip exists alongside as the edit-mode fallback
    expect(editor.view.dom.querySelectorAll('.doc-formula-token').length).toBeGreaterThan(0)
    editor.destroy()
  })

  it('saves the inserted equation and reparses it as a formula block', async () => {
    const { editor, parsed } = await openDoc()
    insertEquationFromLatex(editor, 'a^2 + b^2 = c^2')

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const bytes = await saveDocx(parsed, plan.saveBlocks)
    const reparsed = await parseDocx(bytes)

    const formulaBlock = reparsed.blocks.find((b) => b.formulaDisplay)
    expect(formulaBlock).toBeDefined()
    expect(formulaBlock!.formulaDisplay!.mathml).toContain('<msup>')
    expect(formulaBlock!.formulaDisplay!.tokens.join('')).toContain('=')
    editor.destroy()
  })

  it('every builtin gallery equation converts without error', async () => {
    const { editor } = await openDoc()
    for (const eq of BUILTIN_EQUATIONS) {
      expect(() => insertEquationFromLatex(editor, eq.latex), eq.name).not.toThrow()
    }
    expect(editor.view.dom.querySelectorAll('.doc-formula-math').length).toBe(
      BUILTIN_EQUATIONS.length,
    )
    editor.destroy()
  })

  it('re-edits a saved display equation using the reopened OMML, then saves the change', async () => {
    const { editor, parsed } = await openDoc()
    insertEquationFromLatex(editor, '\\frac{a}{b}')
    const first = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, first.saveBlocks))
    editor.destroy()

    const nextEditor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reopened.blocks) as never,
    })
    let pos = -1
    nextEditor.state.doc.descendants((node, at) => {
      if (node.type.name === 'docProtected' && node.attrs.formulaDisplay) pos = at
    })
    expect(pos).toBeGreaterThanOrEqual(0)
    const target = { pos, latex: 'stale metadata', kind: 'block' as const }
    expect(currentEquationSource(nextEditor, target)).toContain('\\frac')
    nextEditor.view.dispatch(
      nextEditor.state.tr.setSelection(NodeSelection.create(nextEditor.state.doc, pos)),
    )
    expect(selectedEquationTarget(nextEditor)?.kind).toBe('block')
    expect(updateEquationFromLatex(nextEditor, target, '\\sqrt{x}')).toBe(true)
    const second = pmDocToSavePlan(nextEditor.getJSON() as PmNode, reopened.blocks)
    const final = await parseDocx(await saveDocx(reopened, second.saveBlocks))
    expect(final.blocks.find((block) => block.formulaDisplay)?.formulaDisplay?.mathml).toContain(
      '<msqrt>',
    )
    nextEditor.destroy()
  })

  it('re-edits a saved inline equation and keeps surrounding text', async () => {
    const { editor, parsed } = await openDoc()
    editor.commands.setTextSelection(3)
    insertInlineEquationFromLatex(editor, 'x^2')
    const first = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reopened = await parseDocx(await saveDocx(parsed, first.saveBlocks))
    editor.destroy()

    const nextEditor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(reopened.blocks) as never,
    })
    let pos = -1
    nextEditor.state.doc.descendants((node, at) => {
      if (node.type.name === 'docInlineMath') pos = at
    })
    expect(pos).toBeGreaterThanOrEqual(0)
    const target = { pos, latex: '', kind: 'inline' as const }
    expect(currentEquationSource(nextEditor, target)).toContain('^')
    nextEditor.commands.setTextSelection(pos)
    expect(selectedEquationTarget(nextEditor)?.kind).toBe('inline')
    expect(updateEquationFromLatex(nextEditor, target, '\\frac{a}{b}')).toBe(true)
    const second = pmDocToSavePlan(nextEditor.getJSON() as PmNode, reopened.blocks)
    const final = await parseDocx(await saveDocx(reopened, second.saveBlocks))
    const runs = final.blocks[0].runs ?? []
    expect(
      runs
        .filter((run) => !run.math)
        .map((run) => run.text)
        .join(''),
    ).toBe('hello')
    expect(runs.find((run) => run.math)?.math?.omml).toContain('<m:f>')
    nextEditor.destroy()
  })

  it('inserts a supplementary-plane symbol with its chosen font at the captured cursor', async () => {
    const { editor, parsed } = await openDoc()
    insertSymbolAtSelection(editor, { from: 3, to: 3 }, '𝄞', 'Segoe UI Symbol')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reparsed.blocks[0].runs?.map((run) => run.text).join('')).toBe('he𝄞llo')
    expect(reparsed.blocks[0].runs?.find((run) => run.text === '𝄞')?.fontAscii).toBe(
      'Segoe UI Symbol',
    )
    editor.destroy()
  })

  it('keeps bold at the insertion cursor while applying the chosen symbol font', async () => {
    const { editor, parsed } = await openDoc()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.setMark('bold')
    editor.commands.setMark('docTextStyle', { sizeHalfPoints: 28, color: 'C00000' })
    editor.commands.setTextSelection(3)

    insertSymbolAtSelection(editor, { from: 3, to: 3 }, '𝄞', 'Segoe UI Symbol')

    const symbolNode = (editor.getJSON().content?.[0]?.content ?? []).find(
      (node) => 'text' in node && typeof node.text === 'string' && node.text.includes('𝄞'),
    )
    expect(symbolNode?.marks?.some((mark) => mark.type === 'bold')).toBe(true)
    expect(symbolNode?.marks?.find((mark) => mark.type === 'docTextStyle')?.attrs).toMatchObject({
      fontAscii: 'Segoe UI Symbol',
      sizeHalfPoints: 28,
      color: 'C00000',
    })

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    const symbolRun = reparsed.blocks[0].runs?.find((run) => run.text === '𝄞')
    expect(symbolRun?.bold).toBe(true)
    expect(symbolRun?.fontAscii).toBe('Segoe UI Symbol')
    expect(symbolRun?.sizeHalfPoints).toBe(28)
    expect(symbolRun?.color).toBe('C00000')
    editor.destroy()
  })

  it('keeps unsupported imported OMML unavailable for full re-edit and unchanged on save', async () => {
    const original = '<m:oMath><m:phant><m:e><m:r><m:t>x</m:t></m:r></m:e></m:phant></m:oMath>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: `<w:p>${original}</w:p>` }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const target = { pos: 0, latex: 'stale', kind: 'block' as const }
    expect(currentEquationSource(editor, target)).toBeNull()
    expect(updateEquationFromLatex(editor, target, 'y')).toBe(false)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    const reparsed = await parseDocx(await saveDocx(parsed, plan.saveBlocks))
    expect(reparsed.internal.documentXml).toContain(original)
    editor.destroy()
  })
})
