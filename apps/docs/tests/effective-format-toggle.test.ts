/**
 * C1: bold/italic inherited from a paragraph style must be reachable-off (Word
 * writes w:b w:val="0"), the ribbon must report the EFFECTIVE value, and Format
 * Painter from Normal text onto a heading must un-bold it.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildBlankDocx, parseDocx, saveDocx, type ParsedDocFull } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import {
  effectiveFlag,
  isFlagActive,
  styleContextOf,
  toggleEffectiveFlag,
} from '../src/renderer/editor/effective-format'
import { undo, redo } from '@tiptap/pm/history'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

async function headingEditor(): Promise<{ editor: Editor; parsed: ParsedDocFull }> {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = (await parseDocx(await buildBlankDocx())) as ParsedDocFull
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editors.push(editor)
  editor.storage.listNumbering.styles = parsed.styles
  editor.storage.listNumbering.docDefaults = parsed.docDefaults
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  const heading = editor.schema.nodeFromJSON({
    type: 'docHeading',
    attrs: { level: 1, styleId: 'Heading1' },
    content: [{ type: 'text', text: 'Title' }],
  })
  const body = editor.schema.nodeFromJSON({
    type: 'docParagraph',
    content: [{ type: 'text', text: 'Body' }],
  })
  editor.view.dispatch(
    editor.state.tr.replaceWith(0, editor.state.doc.content.size, [heading, body]),
  )
  return { editor, parsed }
}

const headingText = (editor: Editor) => editor.state.doc.firstChild!.firstChild!

async function reopenHeadingRuns(editor: Editor, parsed: ParsedDocFull) {
  const bytes = await saveDocx(
    parsed,
    pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
  )
  const reparsed = await parseDocx(bytes)
  const heading = reparsed.blocks.find((b) => b.type === 'heading')!
  return { reparsed, runs: heading.runs ?? [] }
}

describe('effective bold/italic (C1)', () => {
  it('reports bold inherited from Heading 1 as active', async () => {
    const { editor } = await headingEditor()
    expect(editor.storage.listNumbering.styles?.get('Heading1')?.display?.bold).toBe(true)
    editor.commands.setTextSelection({ from: 1, to: 6 })
    expect(editor.isActive('bold')).toBe(false) // the old ribbon source
    expect(isFlagActive(editor.state, 'bold', styleContextOf(editor))).toBe(true)
    editor.commands.setTextSelection(3)
    expect(isFlagActive(editor.state, 'bold', styleContextOf(editor))).toBe(true)
  })

  it('toggling off writes an explicit off-switch; toggling again restores bold', async () => {
    const { editor } = await headingEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    toggleEffectiveFlag(editor, 'bold')
    const ctx = styleContextOf(editor)
    const run = headingText(editor)
    expect(run.marks.some((m) => m.type.name === 'bold')).toBe(false)
    expect(run.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.boldOff).toBe(true)
    expect(isFlagActive(editor.state, 'bold', ctx)).toBe(false)

    toggleEffectiveFlag(editor, 'bold')
    expect(isFlagActive(editor.state, 'bold', ctx)).toBe(true)
    expect(
      effectiveFlag('bold', headingText(editor).marks, editor.state.doc.firstChild!, ctx),
    ).toBe(true)
  })

  it('toggling in a body paragraph keeps plain toggleMark behavior (no off-switch)', async () => {
    const { editor } = await headingEditor()
    const bodyStart = editor.state.doc.firstChild!.nodeSize + 1
    editor.commands.setTextSelection({ from: bodyStart, to: bodyStart + 4 })
    toggleEffectiveFlag(editor, 'bold')
    toggleEffectiveFlag(editor, 'bold')
    const run = editor.state.doc.child(1).firstChild!
    expect(run.marks).toHaveLength(0)
  })

  it('caret toggle stores the off-switch for typed text', async () => {
    const { editor } = await headingEditor()
    editor.commands.setTextSelection(6)
    toggleEffectiveFlag(editor, 'bold')
    editor.view.dispatch(editor.state.tr.insertText('X'))
    const typed = editor.state.doc.firstChild!.lastChild!
    expect(typed.text).toBe('X')
    expect(typed.marks.find((m) => m.type.name === 'docTextStyle')?.attrs.boldOff).toBe(true)
  })

  it('undo/redo restores and reapplies the un-bold as one step', async () => {
    const { editor } = await headingEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    const ctx = styleContextOf(editor)
    toggleEffectiveFlag(editor, 'bold')
    expect(isFlagActive(editor.state, 'bold', ctx)).toBe(false)
    undo(editor.state, editor.view.dispatch)
    editor.commands.setTextSelection({ from: 1, to: 6 })
    expect(isFlagActive(editor.state, 'bold', ctx)).toBe(true)
    redo(editor.state, editor.view.dispatch)
    editor.commands.setTextSelection({ from: 1, to: 6 })
    expect(isFlagActive(editor.state, 'bold', ctx)).toBe(false)
  })

  it('DOCX round-trip writes w:b w:val="0" and reopens un-bold', async () => {
    const { editor, parsed } = await headingEditor()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    toggleEffectiveFlag(editor, 'bold')
    toggleEffectiveFlag(editor, 'italic') // Heading 1 is not italic: italic turns ON
    const { runs } = await reopenHeadingRuns(editor, parsed)
    expect(runs.map((r) => r.text).join('')).toBe('Title')
    expect(runs[0].bold).toBe(false)
    expect(runs[0].italic).toBe(true)
    expect(runs[0].rawRPr).toMatch(/<w:b w:val="0"\/>/)

    // reopen → editor carries boldOff → effective off; saving again is stable
    const { editorExtensions } = await import('../src/renderer/editor/extensions')
    const reparsed = (await parseDocx(
      await saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks),
    )) as ParsedDocFull
    const again = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editors.push(again)
    again.storage.listNumbering.styles = reparsed.styles
    again.storage.listNumbering.docDefaults = reparsed.docDefaults
    again.commands.setContent(blocksToPmDoc(reparsed.blocks) as never)
    const hIdx = reparsed.blocks.findIndex((b) => b.type === 'heading')
    let pos = 0
    for (let i = 0; i < hIdx; i++) pos += again.state.doc.child(i).nodeSize
    again.commands.setTextSelection({ from: pos + 1, to: pos + 6 })
    expect(isFlagActive(again.state, 'bold', styleContextOf(again))).toBe(false)
    // continue editing: bold it back on → saved without the off-switch
    toggleEffectiveFlag(again, 'bold')
    const { runs: runs2 } = await reopenHeadingRuns(again, reparsed)
    expect(runs2[0].bold).toBe(true)
  })

  it('Format Painter semantics: a non-bold source makes heading text effectively non-bold', async () => {
    const { editor } = await headingEditor()
    const ctx = styleContextOf(editor)
    const bodyPara = editor.state.doc.child(1)
    // the painter's pickup records boldOff when the source's effective bold is off
    expect(effectiveFlag('bold', bodyPara.firstChild!.marks, bodyPara, ctx)).toBe(false)
    const style = editor.schema.marks.docTextStyle.create({ boldOff: true })
    editor.view.dispatch(editor.state.tr.addMark(1, 6, style))
    editor.commands.setTextSelection({ from: 1, to: 6 })
    expect(isFlagActive(editor.state, 'bold', ctx)).toBe(false)
  })

  it('a gallery heading (level only, styleId null) inherits its level style bold', async () => {
    const { editor } = await headingEditor()
    const tr = editor.state.tr.setNodeMarkup(0, undefined, {
      ...editor.state.doc.firstChild!.attrs,
      styleId: null,
      level: 1,
    })
    editor.view.dispatch(tr)
    editor.commands.setTextSelection({ from: 1, to: 6 })
    const ctx = styleContextOf(editor)
    expect(isFlagActive(editor.state, 'bold', ctx)).toBe(true)
    toggleEffectiveFlag(editor, 'bold')
    expect(isFlagActive(editor.state, 'bold', ctx)).toBe(false)
    expect(
      headingText(editor).marks.find((m) => m.type.name === 'docTextStyle')?.attrs.boldOff,
    ).toBe(true)
    // a plain paragraph with no styleId resolves to Normal (not bold)
    const body = editor.state.doc.child(1)
    expect(effectiveFlag('bold', body.firstChild!.marks, body, ctx)).toBe(false)
  })
})
