/**
 * A double-click in the empty space right of a paragraph's last line: Chromium selects
 * the paragraph break there, so the next keystroke replaced it and merged the paragraph
 * with the one below. Word never swallows the break (its Click-and-Type puts the caret
 * on that line), so such a double-click places the caret at the end of the paragraph
 * instead of letting the browser pick a "word".
 */
import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'

export const DoubleClickLineEnd = Extension.create({
  name: 'doubleClickLineEnd',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('doubleClickLineEnd'),
        props: {
          handleDoubleClick(view, pos, event) {
            const $pos = view.state.doc.resolve(pos)
            if (!$pos.parent.isTextblock || $pos.parentOffset !== $pos.parent.content.size) {
              return false
            }
            // only a click beyond the text: on the last word itself stays a word selection
            if (event.clientX <= view.coordsAtPos(pos).right + 2) return false
            view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)))
            return true
          },
        },
      }),
    ]
  },
})
