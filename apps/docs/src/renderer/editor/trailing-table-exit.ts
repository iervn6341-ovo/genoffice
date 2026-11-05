import { Extension } from '@tiptap/core'
import { GapCursor } from '@tiptap/pm/gapcursor'
import { Plugin, PluginKey, TextSelection } from '@tiptap/pm/state'
import { TRACK_IGNORE } from './revisions'

const key = new PluginKey('trailingTableExit')

export const TrailingTableExitExtension = Extension.create({
  name: 'trailingTableExit',

  addProseMirrorPlugins() {
    const editor = this.editor
    return [
      new Plugin({
        key,
        appendTransaction(transactions, oldState, newState) {
          const editsLanding = transactions.some((transaction) => {
            if (!transaction.docChanged || transaction.getMeta('appendedTransaction')) return false
            const last = transaction.before.lastChild
            if (!last?.attrs.navigationOnly) return false
            const start = transaction.before.content.size - last.nodeSize
            let touched = false
            for (const step of transaction.steps)
              step.getMap().forEach((from, to) => {
                if (from <= transaction.before.content.size && to >= start) touched = true
              })
            return touched
          })
          if (editsLanding) {
            const authored = newState.tr
            newState.doc.forEach((node, pos) => {
              if (node.attrs.navigationOnly)
                authored.setNodeMarkup(pos, undefined, { ...node.attrs, navigationOnly: false })
            })
            if (authored.docChanged) return authored.setMeta('addToHistory', false)
          }
          if (
            !editor.view.editable ||
            oldState.doc !== newState.doc ||
            !(newState.selection instanceof GapCursor) ||
            newState.selection.from !== newState.doc.content.size ||
            newState.doc.lastChild?.type.name !== 'docTable' ||
            !transactions.some((transaction) => transaction.selectionSet && !transaction.docChanged)
          ) {
            return null
          }

          const paragraph = newState.schema.nodes.docParagraph?.createAndFill({
            navigationOnly: true,
          })
          if (!paragraph) return null
          const pos = newState.doc.content.size
          const transaction = newState.tr.insert(pos, paragraph)
          // leaving the table is navigation, not authoring: no revision mark
          return transaction
            .setMeta(TRACK_IGNORE, true)
            .setSelection(TextSelection.create(transaction.doc, pos + 1))
            .scrollIntoView()
        },
      }),
    ]
  },
})
