/**
 * Word's AutoFormat As You Type for lists: typing "* ", "- ", "1. ", "1) ", "a. ", "a) ",
 * "A. " or "A) " at the very start of a plain paragraph turns it into a bulleted / numbered
 * list item and drops the typed marker. The conversion is its own undo step, so ⌘Z right
 * after brings the marker back as literal text — Word's AutoCorrect undo.
 */
import { Extension } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import { Plugin, PluginKey } from '@tiptap/pm/state'

export type AutoListSpec =
  | { kind: 'bullet'; glyph: string }
  | { kind: 'ordered'; numFmt: 'decimal' | 'lowerLetter' | 'upperLetter'; pattern: '%1.' | '%1)' }

/** The list a typed paragraph-start marker asks for; null = not a list marker */
export function autoListSpec(marker: string): AutoListSpec | null {
  if (marker === '*' || marker === '•') return { kind: 'bullet', glyph: '•' }
  if (marker === '-') return { kind: 'bullet', glyph: '–' }
  const m = /^([1aA])([.)])$/.exec(marker)
  if (!m) return null
  const numFmt = m[1] === '1' ? 'decimal' : m[1] === 'a' ? 'lowerLetter' : 'upperLetter'
  return { kind: 'ordered', numFmt, pattern: m[2] === '.' ? '%1.' : '%1)' }
}

export interface ListAutoFormatStorage {
  /** Set by the app: the numId of the list the spec describes (null leaves the text alone) */
  resolve: ((spec: AutoListSpec) => string | null) | null
}

declare module '@tiptap/core' {
  interface Storage {
    listAutoFormat: ListAutoFormatStorage
  }
}

export const ListAutoFormat = Extension.create<Record<string, never>, ListAutoFormatStorage>({
  name: 'listAutoFormat',
  addStorage() {
    return { resolve: null }
  },
  addProseMirrorPlugins() {
    const editor = this.editor
    const storage = this.storage
    return [
      new Plugin({
        key: new PluginKey('listAutoFormat'),
        props: {
          handleTextInput(view, from, to, text) {
            if (text !== ' ' || from !== to || !storage.resolve || !editor.isEditable) return false
            const $from = view.state.doc.resolve(from)
            const para = $from.parent
            // plain paragraphs only: headings keep their style, list items are lists already
            if (para.type.name !== 'docParagraph') return false
            // the marker must be the whole paragraph so far (an inline object reads as U+FFFC)
            const marker = para.textBetween(0, $from.parentOffset, undefined, '￼')
            const spec = autoListSpec(marker)
            if (!spec) return false
            // the space lands as typed, in the typing undo step...
            view.dispatch(view.state.tr.insertText(' ', from, to))
            const numId = storage.resolve(spec)
            if (!numId) return true
            // ...and the conversion is a step of its own (closeHistory), like the Bullets button
            const start = $from.start()
            editor
              .chain()
              .command(({ tr }) => {
                tr.delete(start, start + marker.length + 1)
                closeHistory(tr)
                return true
              })
              .setNode('docListItem', { kind: spec.kind, numId, ilvl: 0 })
              .run()
            return true
          },
        },
      }),
    ]
  },
})
