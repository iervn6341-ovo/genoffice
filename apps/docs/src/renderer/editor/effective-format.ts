import type { Editor } from '@tiptap/core'
import type { Mark as PmMark, Node as PmNode } from '@tiptap/pm/model'
import type { EditorState, Transaction } from '@tiptap/pm/state'
import type { DocDefaults, StyleInfo } from '@genoffice/docx-engine'

/**
 * Word's Bold/Italic reflect the EFFECTIVE format — direct run formatting over
 * the character style over the paragraph style over docDefaults — and turning
 * one off under a bold style writes an explicit `w:b w:val="0"` (the
 * docTextStyle boldOff/italicOff attr here). A plain toggleMark only knows the
 * direct mark, so Heading 1 text showed Bold inactive and could never be
 * un-bolded.
 */
export type ToggleFlag = 'bold' | 'italic'

const OFF_ATTR: Record<ToggleFlag, 'boldOff' | 'italicOff'> = {
  bold: 'boldOff',
  italic: 'italicOff',
}

export interface StyleContext {
  styles?: Map<string, StyleInfo>
  docDefaults?: DocDefaults
}

export function styleContextOf(editor: Editor): StyleContext {
  const storage = (editor.storage as { listNumbering?: StyleContext }).listNumbering
  return { styles: storage?.styles, docDefaults: storage?.docDefaults }
}

/** The flag the styles give a run: character style → paragraph style → docDefaults */
export function inheritedFlag(
  flag: ToggleFlag,
  charStyleId: unknown,
  paraStyleId: unknown,
  ctx: StyleContext,
): boolean {
  const display = (id: unknown) =>
    typeof id === 'string' && id ? ctx.styles?.get(id)?.display : undefined
  return !!(display(charStyleId)?.[flag] ?? display(paraStyleId)?.[flag] ?? ctx.docDefaults?.[flag])
}

/** Effective flag for a run with these marks inside `parent` */
export function effectiveFlag(
  flag: ToggleFlag,
  marks: readonly PmMark[],
  parent: PmNode,
  ctx: StyleContext,
): boolean {
  if (marks.some((m) => m.type.name === flag)) return true
  const style = marks.find((m) => m.type.name === 'docTextStyle')
  if (style?.attrs[OFF_ATTR[flag]] === true) return false
  return inheritedFlag(flag, style?.attrs.styleId, parent.attrs.styleId, ctx)
}

/** Is the flag effectively on across the selection (every text run; the caret's marks when empty)? */
export function isFlagActive(state: EditorState, flag: ToggleFlag, ctx: StyleContext): boolean {
  const { selection } = state
  if (selection.empty) {
    const { $from } = selection
    return effectiveFlag(flag, state.storedMarks ?? $from.marks(), $from.parent, ctx)
  }
  let seen = false
  let all = true
  for (const range of selection.ranges) {
    state.doc.nodesBetween(range.$from.pos, range.$to.pos, (node, _pos, parent) => {
      if (!all) return false
      if (!node.isText || !parent) return true
      seen = true
      if (!effectiveFlag(flag, node.marks, parent, ctx)) all = false
      return false
    })
  }
  return seen && all
}

/** Marks with the flag switched to `on`: the direct mark is added/removed, and an
 *  explicit off-switch is written only where a style would otherwise paint it. */
function marksWithFlag(
  state: Pick<EditorState, 'schema'>,
  flag: ToggleFlag,
  on: boolean,
  marks: readonly PmMark[],
  parent: PmNode,
  ctx: StyleContext,
): readonly PmMark[] {
  const markType = state.schema.marks[flag]
  const styleType = state.schema.marks.docTextStyle
  let out: readonly PmMark[] = marks.filter((m) => m.type !== markType)
  if (on) return markType.create().addToSet(out)
  const style = out.find((m) => m.type === styleType)
  const off = OFF_ATTR[flag]
  const needsOff = inheritedFlag(flag, style?.attrs.styleId, parent.attrs.styleId, ctx)
  if (!styleType) return out
  if (needsOff) {
    out = styleType.create({ ...(style?.attrs ?? {}), [off]: true }).addToSet(out)
  } else if (style && style.attrs[off] != null) {
    out = styleType.create({ ...style.attrs, [off]: null }).addToSet(out)
  }
  return out
}

/** Set the flag's effective value over the transaction's selection (one undo step with
 *  whatever else `tr` carries, e.g. the Font dialog's other changes) */
export function applyFlag(
  tr: Transaction,
  flag: ToggleFlag,
  on: boolean,
  ctx: StyleContext,
): Transaction {
  const { selection, doc } = tr
  const schemaState = { schema: doc.type.schema }
  if (selection.empty) {
    const { $from } = selection
    return tr.setStoredMarks(
      marksWithFlag(schemaState, flag, on, tr.storedMarks ?? $from.marks(), $from.parent, ctx),
    )
  }
  for (const range of selection.ranges) {
    const start = range.$from.pos
    const end = range.$to.pos
    doc.nodesBetween(start, end, (node, pos, parent) => {
      if (!node.isText || !parent) return true
      const from = Math.max(pos, start)
      const to = Math.min(pos + node.nodeSize, end)
      const next = marksWithFlag(schemaState, flag, on, node.marks, parent, ctx)
      for (const m of node.marks) if (!m.isInSet(next)) tr.removeMark(from, to, m)
      for (const m of next) if (!m.isInSet(node.marks)) tr.addMark(from, to, m)
      return false
    })
  }
  return tr
}

export function setFlagTr(
  state: EditorState,
  flag: ToggleFlag,
  on: boolean,
  ctx: StyleContext,
): Transaction {
  return applyFlag(state.tr, flag, on, ctx)
}

/** Word's B / I button and ⌘B / ⌘I: flip the effective value */
export function toggleEffectiveFlag(editor: Editor, flag: ToggleFlag): boolean {
  if (!editor.isEditable) return false
  const ctx = styleContextOf(editor)
  const on = !isFlagActive(editor.state, flag, ctx)
  editor.view.dispatch(setFlagTr(editor.state, flag, on, ctx).scrollIntoView())
  return true
}
