import { Editor } from '@tiptap/core'
import { afterEach, describe, expect, it } from 'vitest'
import { TextSelection } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  bodyDirtyAgainstSaved,
  carrySavedTocHistorySource,
  forceBodyDirty,
  persistedBodyChanged,
  rememberSavedBody,
} from '../src/renderer/doc-dirty'

const editors: Editor[] = []
afterEach(() => {
  editors.splice(0).forEach((editor) => editor.destroy())
})
function editor() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'Read me', marks: [{ type: 'bold' }] }],
        },
        { type: 'docParagraph' },
        {
          type: 'docParagraph',
          content: [{ type: 'docNoteRef', attrs: { id: '8', kind: 'footnote', num: 8 } }],
        },
      ],
    },
  })
  editors.push(editor)
  return editor
}

describe('persisted body dirtiness', () => {
  it('ignores automatic TOC page labels and keeps explicit field updates undoable to clean', () => {
    const e = editor()
    e.commands.setContent({
      type: 'doc',
      content: [
        {
          type: 'docProtected',
          attrs: {
            docxIndex: 0,
            fieldDisplay: { kind: 'tocLine', left: 'Chapter', right: '9', anchor: '_Toc1' },
          },
        },
      ],
    })
    rememberSavedBody(e)
    const original = e.state.doc.firstChild!
    e.view.dispatch(
      e.state.tr
        .setNodeMarkup(0, undefined, {
          ...original.attrs,
          fieldDisplay: { ...original.attrs.fieldDisplay, right: '2' },
        })
        .setMeta('addToHistory', false),
    )
    expect(bodyDirtyAgainstSaved(e)).toBe(false)
    const source = e.state.doc.firstChild!
    carrySavedTocHistorySource(e, source, '<w:p>original field</w:p>')
    e.view.dispatch(
      e.state.tr
        .setNodeMarkup(0, undefined, {
          ...source.attrs,
          genXml: '<w:p>original field</w:p>',
        })
        .setMeta('addToHistory', false),
    )
    expect(bodyDirtyAgainstSaved(e)).toBe(false)
    e.view.dispatch(closeHistory(e.state.tr))
    e.view.dispatch(
      e.state.tr.setNodeMarkup(0, undefined, {
        ...e.state.doc.firstChild!.attrs,
        genXml: '<w:p>updated field</w:p>',
      }),
    )
    expect(bodyDirtyAgainstSaved(e)).toBe(true)
    e.commands.undo()
    expect(bodyDirtyAgainstSaved(e)).toBe(false)
    e.commands.redo()
    expect(bodyDirtyAgainstSaved(e)).toBe(true)
  })

  it('tracks Undo and Redo against the last successful save, including Undo across that save', () => {
    const e = editor()
    rememberSavedBody(e)
    e.commands.insertContent('Edited')
    expect(bodyDirtyAgainstSaved(e)).toBe(true)
    e.commands.undo()
    expect(bodyDirtyAgainstSaved(e)).toBe(false)
    e.commands.redo()
    expect(bodyDirtyAgainstSaved(e)).toBe(true)
    rememberSavedBody(e)
    expect(bodyDirtyAgainstSaved(e)).toBe(false)
    e.commands.undo()
    expect(bodyDirtyAgainstSaved(e)).toBe(true)
    e.commands.redo()
    expect(bodyDirtyAgainstSaved(e)).toBe(false)
    forceBodyDirty(e)
    expect(bodyDirtyAgainstSaved(e)).toBe(true)
    rememberSavedBody(e)
    expect(bodyDirtyAgainstSaved(e)).toBe(false)
  })
  it('ignores selection and session-only caret format memory, while typing is still an edit', () => {
    const e = editor()
    const before = e.state.doc
    const pos = before.child(0).nodeSize
    e.view.dispatch(
      e.state.tr
        .setSelection(TextSelection.create(e.state.doc, pos + 1))
        .setStoredMarks([e.schema.marks.bold.create()]),
    )
    expect(e.state.doc.child(1).attrs.caretMarks).not.toBeNull()
    expect(e.state.doc.eq(before)).toBe(false)
    expect(persistedBodyChanged(before, e.state.doc)).toBe(false)
    e.commands.insertContent('new text')
    expect(persistedBodyChanged(before, e.state.doc)).toBe(true)
  })

  it('ignores derived note numbering but preserves real note IDs and paragraph formatting changes', () => {
    const e = editor()
    const before = e.state.doc
    let ref = 0
    before.descendants((node, pos) => {
      if (node.type.name === 'docNoteRef') ref = pos
    })
    const attrs = before.nodeAt(ref)!.attrs
    e.view.dispatch(
      e.state.tr.setNodeMarkup(ref, undefined, { ...attrs, num: 1 }).setMeta('addToHistory', false),
    )
    expect(persistedBodyChanged(before, e.state.doc)).toBe(false)
    e.view.dispatch(e.state.tr.setNodeMarkup(ref, undefined, { ...attrs, id: '9', num: 1 }))
    expect(persistedBodyChanged(before, e.state.doc)).toBe(true)
    const clean = editor()
    const original = clean.state.doc
    clean.view.dispatch(
      clean.state.tr.setNodeMarkup(0, undefined, {
        ...original.firstChild!.attrs,
        align: 'center',
      }),
    )
    expect(persistedBodyChanged(original, clean.state.doc)).toBe(true)
  })

  it('does not suppress real changes just because their transactions are excluded from history', () => {
    const e = editor()
    const before = e.state.doc
    e.view.dispatch(e.state.tr.insertText('Author edit', 1).setMeta('addToHistory', false))
    expect(persistedBodyChanged(before, e.state.doc)).toBe(true)
  })
})
