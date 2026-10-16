import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { beforeEach, describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { writeUserAutoCorrect } from '../src/renderer/editor/autocorrect'

function typed(text: string, trigger: string): string {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', content: [{ type: 'text', text }] }],
    },
  })
  const end = editor.state.doc.content.size - 1
  editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, end)))
  const handled = editor.view.someProp('handleTextInput', (f) =>
    f(editor.view, end, end, trigger, () => editor.state.tr),
  )
  if (!handled) editor.view.dispatch(editor.state.tr.insertText(trigger))
  const out = editor.state.doc.textContent
  editor.destroy()
  return out
}

describe('AutoCorrect', () => {
  beforeEach(() => localStorage.clear())
  it('replaces the default entries when a space follows', () => {
    expect(typed('(c)', ' ')).toBe('© ')
    expect(typed('Fee (tm)', ' ')).toBe('Fee ™ ')
    expect(typed('wait ...', ' ')).toBe('wait … ')
    expect(typed('a --> b -->', ' ')).toBe('a --> b → ')
  })
  it('leaves other text alone and keeps the typed trigger', () => {
    expect(typed('abc(c)', ' ')).toBe('abc(c) ')
    expect(typed('(c)', 'x')).toBe('(c)x')
  })
  it('uses the user entries stored by the Symbol dialog', () => {
    writeUserAutoCorrect({ ':ok': '✔' })
    expect(typed('done :ok', ' ')).toBe('done ✔ ')
  })
})
