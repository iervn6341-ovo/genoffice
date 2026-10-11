/**
 * Markdown shortcuts while typing, each with a small hint after the text and its own undo step
 * (⌘Z right after brings the typed characters back as literal text):
 *
 *  - a plain paragraph that starts with "#" … "######" shows "↵ Heading n"; Enter turns it into
 *    that heading and drops the typed marker
 *  - two or more consecutive paragraphs written as table rows ("| a | b |", usually with a
 *    "| -- | -- |" line under the header) show "↵ table"; Enter turns the whole block into a table
 */
import { Extension } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import { DOMParser as PmDOMParser } from '@tiptap/pm/model'
import { Plugin, PluginKey, TextSelection, type EditorState } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

import { t } from '../i18n/locale'
import { prepareMarkdownBlocks, separateTables } from './markdown-blocks'
import { isPipeRow, isPipeSeparator, markdownPasteHtml, normalizeTableRow } from './markdown-paste'

export interface MarkdownHeading {
  level: number
  /** characters of the typed marker, including the spaces after it */
  markerLength: number
}

/** The heading a paragraph's typed text asks for ("##", "## Title"); null = not a marker */
export function markdownHeading(text: string): MarkdownHeading | null {
  const m = /^(#{1,6})(?:[ \t]+|$)/.exec(text)
  return m ? { level: m[1]!.length, markerLength: m[0].length } : null
}

/** What the typed table rows turn into: the Markdown to convert and where the rows sit. */
export interface TypedTable {
  /** document range of the row paragraphs */
  from: number
  to: number
  /** the rows as Markdown; a trailing header/separator-only block gets one empty body row */
  markdown: string
  /** the block ended at its separator line: the caret goes into the added body row */
  needsBodyRow: boolean
}

const MAX_TABLE_ROWS = 500

interface RowSource {
  childCount: number
  child(i: number): {
    type: { name: string }
    content: { size: number }
    textBetween(from: number, to: number, blockSeparator?: string, leafText?: string): string
    nodeSize: number
  }
}

/**
 * Rows of consecutive table-row paragraphs around `index`, or null when fewer than two rows.
 * A paragraph may hold several rows separated by line breaks (Shift+Enter): each line is a row.
 */
export function typedTableAt(doc: RowSource, index: number): TypedTable | null {
  // the rows of one paragraph, or null when any of its lines is not a table row
  const rowsAt = (i: number): string[] | null => {
    const node = doc.child(i)
    if (node.type.name !== 'docParagraph') return null
    const lines = node
      .textBetween(0, node.content.size, '', '\n')
      .split('\n')
      .map(normalizeTableRow)
    return lines.length > 0 && lines.every((line) => isPipeRow(line)) ? lines : null
  }
  if (rowsAt(index) === null) return null
  let first = index
  let last = index
  while (first > 0 && index - first < MAX_TABLE_ROWS && rowsAt(first - 1) !== null) first -= 1
  while (last + 1 < doc.childCount && last - index < MAX_TABLE_ROWS && rowsAt(last + 1) !== null) {
    last += 1
  }
  let from = 0
  for (let i = 0; i < first; i += 1) from += doc.child(i).nodeSize
  let to = from
  const lines: string[] = []
  for (let i = first; i <= last; i += 1) {
    to += doc.child(i).nodeSize
    lines.push(...rowsAt(i)!)
  }
  if (lines.length < 2) return null
  const needsBodyRow = lines.length === 2 && isPipeSeparator(lines[1]!)
  if (needsBodyRow) {
    const cells = Math.max(
      1,
      lines[0]!
        .trim()
        .replace(/^\||\|$/g, '')
        .split('|').length,
    )
    lines.push(`|${'  |'.repeat(cells)}`)
  }
  return { from, to, markdown: lines.join('\n'), needsBodyRow }
}

/** The block whose hint the user dismissed with Esc: kept (mapped through edits) while the caret
 *  stays inside it, forgotten once the caret leaves — moving back in re-arms the hint. */
type Dismissed = { from: number; to: number } | null

const KEY = new PluginKey<Dismissed>('markdownShortcut')

export const MarkdownShortcut = Extension.create({
  name: 'markdownShortcut',
  // ahead of the Enter keymaps (list / heading / next-style rules)
  priority: 200,
  addProseMirrorPlugins() {
    const editor = this.editor
    const typedHeading = (state: EditorState) => {
      const { selection } = state
      if (!(selection instanceof TextSelection) || !selection.empty) return null
      const { $from } = selection
      if ($from.parent.type.name !== 'docParagraph') return null
      // runs on every state update: a paragraph that does not open with "#" costs one comparison,
      // and one that does is only read as far as a marker can reach
      const size = $from.parent.content.size
      if (size === 0 || $from.parent.textBetween(0, 1) !== '#') return null
      const found = markdownHeading($from.parent.textBetween(0, Math.min(size, 64)))
      return found ? { ...found, start: $from.start(), end: $from.end() } : null
    }
    const typedTable = (state: EditorState) => {
      const { selection } = state
      if (!(selection instanceof TextSelection) || !selection.empty) return null
      const { $from } = selection
      if ($from.depth !== 1 || $from.parent.type.name !== 'docParagraph') return null
      // cheap gate for every state update: a table row opens with "|" (after up to 3 spaces)
      if (!$from.parent.textBetween(0, Math.min($from.parent.content.size, 4)).match(/[|｜]/)) {
        return null
      }
      const found = typedTableAt(state.doc, $from.index(0))
      return found ? { ...found, end: $from.end() } : null
    }
    const hint = (pos: number, label: string, key: string) =>
      Decoration.widget(
        pos,
        () => {
          const root = document.createElement('span')
          root.className = 'md-shortcut-hint'
          root.contentEditable = 'false'
          const text = document.createElement('span')
          text.className = 'md-shortcut-hint-label'
          text.textContent = `↵ ${label}`
          root.append(text)
          return root
        },
        { side: 1, key, ignoreSelection: true },
      )
    return [
      new Plugin({
        key: KEY,
        state: {
          init: (): Dismissed => null,
          apply(tr, value, _old, next): Dismissed {
            const meta = tr.getMeta(KEY) as { dismiss: Dismissed } | undefined
            if (meta) return meta.dismiss
            if (!value) return null
            const from = tr.mapping.map(value.from, -1)
            const to = tr.mapping.map(value.to, 1)
            const caret = next.selection.from
            return caret >= from && caret <= to ? { from, to } : null
          },
        },
        props: {
          decorations(state) {
            if (!editor.isEditable || KEY.getState(state)) return null
            const heading = typedHeading(state)
            if (heading) {
              return DecorationSet.create(state.doc, [
                hint(
                  heading.end,
                  t('editorMdHeadingHint', { n: heading.level }),
                  `md-hint-${heading.level}`,
                ),
              ])
            }
            const table = typedTable(state)
            if (table) {
              return DecorationSet.create(state.doc, [
                hint(table.end, t('editorMdTableHint'), 'md-hint-table'),
              ])
            }
            return null
          },
          handleKeyDown(view, event) {
            // Esc pauses the hint (and the Enter conversion) for the block the caret is in
            if (event.key === 'Escape' && editor.isEditable && !KEY.getState(view.state)) {
              const heading = typedHeading(view.state)
              const table = heading ? null : typedTable(view.state)
              const range = heading
                ? {
                    from: view.state.selection.$from.before(1),
                    to: view.state.selection.$from.after(1),
                  }
                : table
                  ? { from: table.from, to: table.to }
                  : null
              if (!range) return false
              view.dispatch(view.state.tr.setMeta(KEY, { dismiss: range }))
              return true
            }
            if (KEY.getState(view.state)) {
              // Enter in a paused block is an ordinary Enter — and ends the pause, so coming back
              // to the block later offers the hint again
              if (event.key === 'Enter')
                view.dispatch(view.state.tr.setMeta(KEY, { dismiss: null }))
              return false
            }
            if (event.key !== 'Enter' || event.shiftKey || event.altKey || event.ctrlKey)
              return false
            if (event.metaKey || event.isComposing || event.keyCode === 229) return false
            if (!editor.isEditable) return false
            const heading = typedHeading(view.state)
            if (heading) {
              event.preventDefault()
              editor
                .chain()
                .command(({ tr }) => {
                  tr.delete(heading.start, heading.start + heading.markerLength)
                  closeHistory(tr)
                  return true
                })
                .setNode('docHeading', {
                  level: heading.level,
                  styleId: null,
                  outlineOnly: null,
                })
                .run()
              return true
            }
            const table = typedTable(view.state)
            if (!table) return false
            const html = markdownPasteHtml(table.markdown)
            if (!html || !html.includes('<table')) return false
            const schema = view.state.schema
            const dom = new window.DOMParser().parseFromString(html, 'text/html')
            const parsed = PmDOMParser.fromSchema(schema).parse(dom.body).content
            if (parsed.childCount === 0) return false
            event.preventDefault()
            // its own undo step: the rows come back as text
            view.dispatch(closeHistory(view.state.tr))
            const tr = view.state.tr
            const $from = tr.doc.resolve(table.from)
            const prev = $from.index(0) > 0 ? tr.doc.child($from.index(0) - 1) : null
            const $to = tr.doc.resolve(table.to)
            const next = $to.index(0) < tr.doc.childCount ? tr.doc.child($to.index(0)) : null
            const blocks = separateTables(
              prepareMarkdownBlocks(parsed, schema, true),
              prev,
              next,
              schema,
            )
            tr.replaceWith(table.from, table.to, blocks)
            // caret: the new body row of a header-only block, else the end of the last cell
            let tablePos = -1
            let offset = table.from
            blocks.forEach((node) => {
              if (tablePos < 0 && node.type.name === 'docTable') tablePos = offset
              offset += node.nodeSize
            })
            const tableNode = tablePos >= 0 ? tr.doc.nodeAt(tablePos) : null
            if (tableNode) {
              const intoBodyRow = table.needsBodyRow && tableNode.childCount > 1
              const target = intoBodyRow
                ? tablePos + 1 + tableNode.child(0).nodeSize + 1
                : tablePos + tableNode.nodeSize - 2
              tr.setSelection(
                TextSelection.near(
                  tr.doc.resolve(Math.min(target, tr.doc.content.size)),
                  intoBodyRow ? 1 : -1,
                ),
              )
            }
            view.dispatch(tr.scrollIntoView())
            return true
          },
        },
      }),
    ]
  },
})
