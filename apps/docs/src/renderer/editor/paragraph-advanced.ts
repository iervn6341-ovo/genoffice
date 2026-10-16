import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { closeHistory } from '@tiptap/pm/history'
import { alignAttrFor, effectiveBidi } from './direction'
import { LAYOUT_ATTR, RECORD_META, type LayoutSnapshot } from './layout-history'

const paragraphTypes = new Set(['docParagraph', 'docHeading', 'docListItem'])

export interface ParagraphSample {
  pos: number
  node: PmNode
}

export function selectedParagraphs(editor: Editor, from: number, to: number): ParagraphSample[] {
  const doc = editor.state.doc
  const found = new Map<number, PmNode>()
  if (from === to) {
    const resolved = doc.resolve(Math.max(0, Math.min(from, doc.content.size)))
    for (let depth = resolved.depth; depth > 0; depth--) {
      const node = resolved.node(depth)
      if (!paragraphTypes.has(node.type.name)) continue
      found.set(resolved.before(depth), node)
      break
    }
  } else {
    doc.nodesBetween(from, to, (node, pos) => {
      if (paragraphTypes.has(node.type.name)) found.set(pos, node)
    })
  }
  return [...found].map(([pos, node]) => ({ pos, node }))
}

export function commonParagraphValue<T>(
  samples: ParagraphSample[],
  valueOf: (sample: ParagraphSample) => T,
): { value: T | undefined; mixed: boolean } {
  if (!samples.length) return { value: undefined, mixed: false }
  const value = valueOf(samples[0])
  return { value, mixed: samples.slice(1).some((sample) => valueOf(sample) !== value) }
}

/** Apply every touched dialog control in one history step while preserving inline content. */
export function applyParagraphChanges(
  editor: Editor,
  from: number,
  to: number,
  attrs: Record<string, unknown>,
  visualAlign?: 'left' | 'center' | 'right' | 'justify',
  layout?: { before: LayoutSnapshot; after: LayoutSnapshot },
): boolean {
  if (
    !editor.isEditable ||
    (Object.keys(attrs).length === 0 && visualAlign === undefined && !layout)
  )
    return false
  const { view } = editor
  if (layout && editor.state.doc.attrs[LAYOUT_ATTR] == null) {
    view.dispatch(
      editor.state.tr
        .setDocAttribute(LAYOUT_ATTR, layout.before)
        .setMeta('addToHistory', false)
        .setMeta(RECORD_META, true),
    )
  }
  const { state } = editor
  const tr = closeHistory(state.tr)
  for (const { pos, node } of selectedParagraphs(editor, from, to)) {
    const next = { ...node.attrs, ...attrs }
    if (visualAlign !== undefined) next.align = alignAttrFor(visualAlign, effectiveBidi(node.attrs))
    if ('spaceBefore' in attrs && !('spaceBeforeAuto' in attrs)) next.spaceBeforeAuto = false
    if ('spaceAfter' in attrs && !('spaceAfterAuto' in attrs)) next.spaceAfterAuto = false
    if (Object.keys(next).some((key) => next[key] !== node.attrs[key]))
      tr.setNodeMarkup(pos, undefined, next)
  }
  if (layout) tr.setDocAttribute(LAYOUT_ATTR, layout.after).setMeta(RECORD_META, true)
  if (!tr.docChanged) return false
  view.dispatch(tr.scrollIntoView())
  view.focus()
  return true
}
