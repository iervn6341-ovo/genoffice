import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { nextNoteReference } from '../src/renderer/editor/note-navigation'

describe('note reference navigation', () => {
  it('moves through editable references by kind and wraps without changing the document', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [
              { type: 'text', text: 'A' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '7' } },
              { type: 'text', text: 'B' },
              { type: 'docNoteRef', attrs: { kind: 'endnote', id: '1' } },
              { type: 'text', text: 'C' },
              { type: 'docNoteRef', attrs: { kind: 'footnote', id: '9' } },
            ],
          },
        ],
      },
    })
    try {
      const doc = editor.state.doc
      const before = doc.toJSON()
      const first = nextNoteReference(doc, 0, 'footnote')
      expect(first).toMatchObject({ id: '7' })
      const second = nextNoteReference(doc, first!.pos, 'footnote')
      expect(second).toMatchObject({ id: '9' })
      expect(nextNoteReference(doc, second!.pos, 'footnote')).toEqual(first)
      expect(nextNoteReference(doc, 0, 'endnote')).toMatchObject({ id: '1' })
      expect(editor.state.doc.toJSON()).toEqual(before)
    } finally {
      editor.destroy()
    }
  })

  it('returns null when the chosen note kind has no editable reference', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: { type: 'doc', content: [{ type: 'docParagraph' }] },
    })
    try {
      expect(nextNoteReference(editor.state.doc, 0, 'footnote')).toBeNull()
      expect(nextNoteReference(editor.state.doc, 0, 'endnote')).toBeNull()
    } finally {
      editor.destroy()
    }
  })
})
