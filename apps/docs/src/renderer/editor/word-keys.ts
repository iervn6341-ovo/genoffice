/**
 * Word-for-Mac editing keys that Chromium / ProseMirror get differently. Every rule here
 * was observed in the real app (Word for Mac, zh-TW) on a probe document, key by key,
 * and replayed against GenOffice — see e2e/word-parity.spec.ts for the golden cases.
 *
 *  - Enter at the END of a paragraph continues with the style's w:next (Heading 1 -> Normal);
 *    Enter at the START of a heading keeps the heading on both halves.
 *  - Backspace / Delete over a selection that starts MID-paragraph and spans paragraphs
 *    deletes the text but keeps the paragraph boundary (Word never merges them); a selection
 *    that starts at a paragraph start keeps the LAST paragraph's formatting.
 *  - Option+Right / Option+Left move by Word's word units (a run of letters/digits, of
 *    punctuation or of ideographs, plus the spaces after it); Option+Up/Down and Cmd+Up/Down
 *    move by paragraph (start of paragraph / start of the next one); Shift+End and
 *    Cmd+Shift+Right pick up the paragraph mark at a paragraph's last line.
 */
import { Extension, getSplittedAttributes } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import type { StyleInfo } from '@genoffice/docx-engine'

const IS_MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform)

/** Word's character classes for word movement: letters/digits, ideographs, punctuation, spaces */
type CharClass = 'w' | 'c' | 'p' | 's'

const IDEOGRAPH = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u

function classOf(ch: string): CharClass {
  if (/\s/u.test(ch)) return 's'
  if (/[\p{L}\p{N}\p{M}]/u.test(ch)) return IDEOGRAPH.test(ch) ? 'c' : 'w'
  return 'p'
}

/**
 * Word's word units of a paragraph's text: a maximal run of one character class
 * (letters/digits, ideographs or punctuation) together with the spaces that follow it.
 * Option+Right lands on the end of a unit, Option+Left on the start of one.
 */
export function wordUnits(text: string): Array<[number, number]> {
  const units: Array<[number, number]> = []
  const at = (i: number): string => String.fromCodePoint(text.codePointAt(i) ?? 0)
  let i = 0
  while (i < text.length) {
    const start = i
    const first = at(i)
    const cls = classOf(first)
    let j = i + first.length
    if (cls !== 's') {
      while (j < text.length && classOf(at(j)) === cls) j += at(j).length
    }
    while (j < text.length && classOf(at(j)) === 's') j += at(j).length
    units.push([start, j])
    i = j
  }
  return units
}

/** end of the first word unit that ends after `offset` (the paragraph end when none does) */
export function wordRight(text: string, offset: number): number {
  for (const [, end] of wordUnits(text)) if (end > offset) return end
  return text.length
}

/** start of the last word unit that starts before `offset` (0 when none does) */
export function wordLeft(text: string, offset: number): number {
  let found = 0
  for (const [start] of wordUnits(text)) {
    if (start < offset) found = start
    else break
  }
  return found
}

/** a textblock's characters, one UTF-16 unit per document position (atoms become U+FFFC) */
const blockText = (node: PmNode): string => node.textBetween(0, node.content.size, undefined, '￼')

interface StyleStorage {
  styles?: Map<string, StyleInfo>
}

/** the open document's style table, as the list numbering extension keeps it */
function styleTable(storage: unknown): Map<string, StyleInfo> | undefined {
  return (storage as { listNumbering?: StyleStorage } | undefined)?.listNumbering?.styles
}

/** the sibling textblock `step` places away in the same parent, or null */
function siblingTextblock($pos: ReturnType<EditorView['state']['doc']['resolve']>, step: 1 | -1) {
  const depth = $pos.depth - 1
  if (depth < 0) return null
  const parent = $pos.node(depth)
  const index = $pos.index(depth) + step
  if (index < 0 || index >= parent.childCount) return null
  const node = parent.child(index)
  if (!node.isTextblock) return null
  const before = $pos.posAtIndex(index, depth)
  return { node, start: before + 1, end: before + 1 + node.content.size }
}

export const WordEnterStyle = Extension.create({
  name: 'wordEnterStyle',
  priority: 150,
  addKeyboardShortcuts() {
    return {
      Enter: () => {
        const { state } = this.editor
        const { selection } = state
        if (!selection.empty || !this.editor.isEditable) return false
        const { $from } = selection
        const node = $from.parent
        if (!node.isTextblock || node.type.name === 'docListItem') return false
        // a break paragraph has its own split rules (WordEditorShortcuts)
        if (node.attrs.pageBreakBefore === true && node.content.size > 0) return false
        const atEnd = $from.parentOffset === node.content.size
        const atStart = $from.parentOffset === 0
        const styles = styleTable(this.editor.storage)
        const styleId = typeof node.attrs.styleId === 'string' ? node.attrs.styleId : null
        const next = styleId ? styles?.get(styleId)?.next : undefined
        const nextInfo = next ? styles?.get(next) : undefined
        if (atEnd && next && next !== styleId) {
          // Enter at the end continues with the style's "style for following paragraph"
          return this.editor
            .chain()
            .splitBlock()
            .command(({ tr, dispatch }) => {
              const $c = tr.selection.$from
              const created = $c.parent
              const heading = nextInfo?.headingLevel
              const type = this.editor.schema.nodes[heading ? 'docHeading' : 'docParagraph']
              if (!type || !created.isTextblock) return false
              if (dispatch) {
                tr.setNodeMarkup($c.before($c.depth), type, {
                  ...created.attrs,
                  styleId: nextInfo?.isDefault ? null : next,
                  ...(heading ? { level: heading, outlineOnly: null } : {}),
                })
              }
              return true
            })
            .run()
        }
        if (atStart && !atEnd && node.type.name === 'docHeading') {
          // Enter at the start of a heading leaves an empty HEADING above (TipTap would turn
          // that first half into a plain paragraph)
          const attrs = getSplittedAttributes(
            this.editor.extensionManager.attributes,
            node.type.name,
            node.attrs,
          )
          return this.editor
            .chain()
            .command(({ tr, dispatch }) => {
              if (dispatch) {
                tr.split(tr.selection.from, 1, [{ type: node.type, attrs }])
                dispatch(tr.scrollIntoView())
              }
              return true
            })
            .run()
        }
        return false
      },
    }
  },
})

/**
 * Deleting a selection that spans paragraphs (Backspace / Delete). Word keeps the paragraph
 * boundary when the selection starts inside a paragraph's text, and takes the LAST
 * paragraph's formatting when it starts at a paragraph's start.
 */
export const WordSelectionDelete = Extension.create({
  name: 'wordSelectionDelete',
  priority: 150,
  addKeyboardShortcuts() {
    const run = (): boolean => {
      const { state, view } = this.editor
      const sel = state.selection
      if (!this.editor.isEditable || !(sel instanceof TextSelection) || sel.empty) return false
      const { $from, $to } = sel
      if ($from.sameParent($to)) return false
      if (!$from.parent.isTextblock || !$to.parent.isTextblock) return false
      if ($from.depth !== $to.depth) return false
      const depth = $from.depth - 1
      if ($from.node(depth) !== $to.node(depth)) return false
      // tracked deletions have their own recorder
      if ((this.editor.storage as { trackChanges?: { enabled?: boolean } }).trackChanges?.enabled)
        return false
      const container = $from.node(depth)
      const first = $from.index(depth)
      const last = $to.index(depth)
      for (let k = first + 1; k < last; k++) if (!container.child(k).isTextblock) return false
      const a = $from.parent
      const startOffset = $from.parentOffset
      if (startOffset === a.content.size) return false // just the paragraph mark: a plain merge
      const tr = state.tr.setMeta('wordSurvivingPara', true)
      if (startOffset === 0) {
        // the first paragraph and everything between go whole; the last one keeps its own
        // mark and formatting with whatever follows the selection
        if ($to.parentOffset > 0) tr.delete($to.start(), $to.pos)
        tr.delete($from.before($from.depth), $to.before($to.depth))
      } else {
        // text only: head of the last paragraph, the paragraphs in between, tail of the first
        if ($to.parentOffset > 0) tr.delete($to.start(), $to.pos)
        if (last - first > 1) tr.delete($from.after(), $to.before())
        tr.delete($from.pos, $from.end())
      }
      const caret = startOffset === 0 ? $from.before($from.depth) + 1 : $from.pos
      tr.setSelection(TextSelection.create(tr.doc, Math.min(caret, tr.doc.content.size)))
      view.dispatch(tr.scrollIntoView())
      return true
    }
    return { Backspace: run, Delete: run }
  },
})

/**
 * Option+Backspace: the word unit before the caret, plus the space in front of it when what
 * follows is closing punctuation (Word's smart cut: "Hello, world|!" → "Hello,!").
 */
export function wordDeleteRange(text: string, offset: number): [number, number] {
  let from = wordLeft(text, offset)
  const after = text[offset]
  if (from > 0 && text[from - 1] === ' ' && after !== undefined && /[.,;:!?)\]}]/.test(after)) {
    from -= 1
  }
  return [from, offset]
}

/** Option / Cmd arrow movement and the selection picks Word makes at a paragraph end */
export const WordNavigation = Extension.create({
  name: 'wordNavigation',
  priority: 150,
  addKeyboardShortcuts() {
    const place = (anchor: number, head: number, extend: boolean): boolean => {
      const { state, view } = this.editor
      const size = state.doc.content.size
      const to = Math.max(0, Math.min(head, size))
      const from = extend ? anchor : to
      view.dispatch(
        state.tr.setSelection(TextSelection.create(state.doc, from, to)).scrollIntoView(),
      )
      return true
    }
    const word = (dir: 1 | -1, extend: boolean) => (): boolean => {
      const { selection } = this.editor.state
      if (!(selection instanceof TextSelection)) return false
      const { $head } = selection
      if (!$head.parent.isTextblock) return false
      const text = blockText($head.parent)
      const offset = $head.parentOffset
      if (dir === 1) {
        if (offset >= text.length) {
          // past the paragraph end: the start of the next paragraph
          const next = siblingTextblock($head, 1)
          return next ? place(selection.anchor, next.start, extend) : false
        }
        return place(selection.anchor, $head.start() + wordRight(text, offset), extend)
      }
      if (offset === 0) {
        const prev = siblingTextblock($head, -1)
        return prev ? place(selection.anchor, prev.end, extend) : false
      }
      return place(selection.anchor, $head.start() + wordLeft(text, offset), extend)
    }
    const paragraph = (dir: 1 | -1, extend: boolean) => (): boolean => {
      const { selection } = this.editor.state
      if (!(selection instanceof TextSelection)) return false
      const { $head } = selection
      if (!$head.parent.isTextblock) return false
      if (dir === -1) {
        if ($head.parentOffset > 0) return place(selection.anchor, $head.start(), extend)
        const prev = siblingTextblock($head, -1)
        return prev ? place(selection.anchor, prev.start, extend) : false
      }
      const next = siblingTextblock($head, 1)
      return next
        ? place(selection.anchor, next.start, extend)
        : place(selection.anchor, $head.end(), extend)
    }
    // Shift+End / Cmd+Shift+Right: the browser stops before the paragraph mark; Word selects it too
    const lineEndWithMark = (): boolean => {
      const { view } = this.editor
      const dom = view.dom.ownerDocument.getSelection()
      const { selection } = view.state
      if (!dom || !(selection instanceof TextSelection)) return false
      const anchor = selection.anchor
      const modify = (dom as Selection & { modify?: (a: string, d: string, g: string) => void })
        .modify
      if (!modify) return false
      modify.call(dom, 'extend', 'forward', 'lineboundary')
      if (!dom.focusNode) return false
      let head: number
      try {
        head = view.posAtDOM(dom.focusNode, dom.focusOffset)
      } catch {
        return false
      }
      const $h = view.state.doc.resolve(head)
      if ($h.parent.isTextblock && $h.parentOffset === $h.parent.content.size) {
        const next = siblingTextblock($h, 1)
        if (next) head = next.start
      }
      view.dispatch(
        view.state.tr
          .setSelection(TextSelection.create(view.state.doc, anchor, head))
          .scrollIntoView(),
      )
      return true
    }
    const deleteWordBack = (): boolean => {
      const { state, view } = this.editor
      const { selection } = state
      if (!(selection instanceof TextSelection) || !selection.empty) return false
      const { $head } = selection
      if (!$head.parent.isTextblock || $head.parentOffset === 0) return false
      const [from, to] = wordDeleteRange(blockText($head.parent), $head.parentOffset)
      const start = $head.start()
      view.dispatch(state.tr.delete(start + from, start + to).scrollIntoView())
      return true
    }
    // Ctrl+Up/Down is Word's paragraph movement on Windows too; the Option and Shift+End rules
    // are the Mac keymap (Windows keeps the browser's Ctrl+Arrow word movement)
    const everywhere = {
      'Mod-ArrowUp': paragraph(-1, false),
      'Mod-ArrowDown': paragraph(1, false),
      'Mod-Shift-ArrowUp': paragraph(-1, true),
      'Mod-Shift-ArrowDown': paragraph(1, true),
    }
    if (!IS_MAC) return everywhere
    return {
      ...everywhere,
      'Alt-Backspace': deleteWordBack,
      'Alt-ArrowRight': word(1, false),
      'Alt-ArrowLeft': word(-1, false),
      'Alt-Shift-ArrowRight': word(1, true),
      'Alt-Shift-ArrowLeft': word(-1, true),
      'Alt-ArrowUp': paragraph(-1, false),
      'Alt-ArrowDown': paragraph(1, false),
      'Alt-Shift-ArrowUp': paragraph(-1, true),
      'Alt-Shift-ArrowDown': paragraph(1, true),
      'Shift-End': lineEndWithMark,
      'Mod-Shift-ArrowRight': lineEndWithMark,
    }
  },
})
