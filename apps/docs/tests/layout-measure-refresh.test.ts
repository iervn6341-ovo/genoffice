import { Editor } from '@tiptap/core'
import { describe, expect, it, vi } from 'vitest'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { refreshLayoutMeasurement } from '../src/renderer/editor/layout-measure-refresh'

describe('document settings live layout', () => {
  it('remeasures CJK punctuation and default tabs without changing content, selection, focus, or undo history', () => {
    const host = document.createElement('div')
    document.body.append(host)
    const editor = new Editor({
      element: host,
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docParagraph', content: [{ type: 'text', text: '日本。' }] },
          { type: 'docParagraph', content: [{ type: 'text', text: 'Before\tAfter' }] },
        ],
      },
    })
    const originalDoc = editor.state.doc
    const originalSelection = editor.state.selection
    const originalFocus = document.activeElement
    const offsetWidth = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'offsetWidth')
    Object.defineProperty(HTMLElement.prototype, 'offsetWidth', {
      configurable: true,
      get: () => 500,
    })
    const measured = vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect')
    try {
      editor.storage.cjkPunctShrink.enabled = true
      refreshLayoutMeasurement(editor, 'cjkPunctShrink')
      expect(
        measured.mock.instances.some((el) => (el as HTMLElement).textContent === '日本。'),
      ).toBe(true)
      measured.mockClear()

      editor.storage.tabStops.defaultTabStopTwips = 480
      refreshLayoutMeasurement(editor, 'tabStops')
      expect(
        measured.mock.instances.some((el) => (el as HTMLElement).textContent === 'Before\tAfter'),
      ).toBe(true)
      expect(editor.state.doc).toBe(originalDoc)
      expect(editor.state.selection.eq(originalSelection)).toBe(true)
      expect(document.activeElement).toBe(originalFocus)
      expect(editor.can().undo()).toBe(false)
    } finally {
      measured.mockRestore()
      if (offsetWidth) Object.defineProperty(HTMLElement.prototype, 'offsetWidth', offsetWidth)
      editor.destroy()
      host.remove()
    }
  })
})
