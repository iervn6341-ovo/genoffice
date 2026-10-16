import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { autoCorrectReplacement } from '@genoffice/office-math'

/**
 * Replace-as-you-type: Word's default AutoCorrect entries ((c), (r), (tm), "...", arrows, fractions)
 * plus the user's own entries from the Symbol dialog. Fires when a space or punctuation is typed
 * right after a matching text that starts the line or follows whitespace; Undo restores the typed text.
 */

const STORAGE_KEY = 'genoffice.docs.autocorrect'
const TRIGGERS = new Set([' ', '.', ',', ';', ':', '!', '?', ')', '\n'])

export function readUserAutoCorrect(): Record<string, string> {
  try {
    const value = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? '{}') as unknown
    if (!value || typeof value !== 'object') return {}
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>).filter(
        (entry): entry is [string, string] => typeof entry[1] === 'string' && entry[0] !== '',
      ),
    )
  } catch {
    return {}
  }
}

export function writeUserAutoCorrect(entries: Record<string, string>): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(entries))
  } catch {
    /* storage is optional */
  }
}

const key = new PluginKey('docsAutoCorrect')

export const AutoCorrectExtension = Extension.create({
  name: 'docsAutoCorrect',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key,
        props: {
          handleTextInput(view, from, to, text) {
            if (text.length !== 1 || !TRIGGERS.has(text) || from !== to) return false
            const $from = view.state.doc.resolve(from)
            if (!$from.parent.isTextblock || $from.parent.type.spec.code) return false
            const start = Math.max($from.start(), from - 24)
            const before = view.state.doc.textBetween(start, from, undefined, '￼')
            const match = autoCorrectReplacement(before, readUserAutoCorrect())
            if (!match || match.length > from - start) return false
            const tr = view.state.tr.insertText(match.with + text, from - match.length, to)
            view.dispatch(tr.scrollIntoView())
            return true
          },
        },
      }),
    ]
  },
})
