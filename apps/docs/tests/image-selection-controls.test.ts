import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { afterEach, describe, expect, it } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'

const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

let editor: Editor | null = null
afterEach(() => {
  editor?.destroy()
  editor = null
})

describe('image selection controls (8-point resize and rotate handle)', () => {
  it('renders 8 resize handles and a rotation handle on protected image blocks', () => {
    const el = document.createElement('div')
    document.body.appendChild(el)
    editor = new Editor({
      element: el,
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: {
              blockType: 'image',
              label: 'Picture 1',
              imageDataUrl: PNG,
              imageWidthPx: 120,
              imageHeightPx: 80,
            },
          },
        ],
      },
    })

    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)),
    )
    editor.view.focus()

    const wrapper = el.querySelector('.doc-protected[data-doc-protected="image"]')
    expect(wrapper).not.toBeNull()

    const rotateHandle = wrapper!.querySelector('.img-rotate-handle')
    expect(rotateHandle).not.toBeNull()

    const rotateLine = wrapper!.querySelector('.img-rotate-handle-line')
    expect(rotateLine).not.toBeNull()

    const resizeHandles = wrapper!.querySelectorAll('.img-resize-handle')
    expect(resizeHandles.length).toBe(8)

    const expectedDirections = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
    const foundDirections = Array.from(resizeHandles).map(
      (h) => (h as HTMLElement).dataset.resizeHandle,
    )
    expect(foundDirections).toEqual(expectedDirections)
  })
})
