import type { Editor } from '@tiptap/core'
import { Mark } from '@tiptap/pm/model'

type SizeStep = (size: number) => number

/**
 * Re-size every inline run in `ranges` by its own size (Word grows 10pt + 20pt text to 10.5pt +
 * 22pt, it does not flatten them) in ONE transaction. Runs without an explicit size start from
 * `fallbackSize`. Returns false when there was nothing to write.
 */
export function stepRunSizes(editor: Editor, fallbackSize: number, step: SizeStep): boolean {
  const { state } = editor
  const markType = state.schema.marks.docTextStyle
  if (!markType) return false
  const tr = state.tr
  let touched = false
  for (const range of state.selection.ranges) {
    const from = range.$from.pos
    const to = range.$to.pos
    if (from === to) continue
    state.doc.nodesBetween(from, to, (node, pos, parent) => {
      if (!node.isInline || !parent?.type.allowsMarkType(markType)) return
      const start = Math.max(from, pos)
      const end = Math.min(to, pos + node.nodeSize)
      if (start >= end) return
      const existing = node.marks.find((m) => m.type === markType)
      const half = existing?.attrs.sizeHalfPoints as number | null | undefined
      const base = half != null ? half / 2 : fallbackSize
      const next = Math.round(step(base) * 2)
      tr.addMark(start, end, markType.create({ ...existing?.attrs, sizeHalfPoints: next }))
      touched = true
    })
  }
  if (!touched) return false
  editor.view.dispatch(tr)
  return true
}

/** Coalesce a burst within a frame, without waiting for the user to stop clicking. */
export function createFontSizeStepper() {
  let pending: {
    editor: Editor
    state: Editor['state']
    fallbackSize: number
    steps: SizeStep[]
  } | null = null
  let frame: number | null = null

  const current = () => {
    if (!pending || pending.editor.isDestroyed || !pending.editor.isEditable) return false
    const state = pending.editor.state
    return (
      state.doc === pending.state.doc &&
      state.selection.eq(pending.state.selection) &&
      Mark.sameSet(state.storedMarks ?? [], pending.state.storedMarks ?? [])
    )
  }
  const cancel = () => {
    if (frame !== null) cancelAnimationFrame(frame)
    frame = null
    pending = null
  }
  const fold = (steps: SizeStep[], size: number) => steps.reduce((acc, f) => f(acc), size)
  return {
    cancel,
    step(editor: Editor, fallbackSize: number, nextSize: SizeStep) {
      if (editor.isDestroyed || !editor.isEditable) return
      if (!current() || pending?.editor !== editor) cancel()
      // clicks inside one frame compose; each run replays the whole chain from its own size
      if (pending) pending.steps.push(nextSize)
      else pending = { editor, state: editor.state, fallbackSize, steps: [nextSize] }
      if (frame !== null) return
      frame = requestAnimationFrame(() => {
        const valid = current()
        const apply = pending
        frame = null
        pending = null
        if (!valid || !apply) return
        const chain = (size: number) => fold(apply.steps, size)
        if (!apply.editor.state.selection.empty) {
          // every run walks the ladder from its own size, written as one undo step
          stepRunSizes(apply.editor, apply.fallbackSize, chain)
          return
        }
        // The ribbon's React snapshot can lag a just-applied mark or a caret change.
        const explicit = apply.editor.getAttributes('docTextStyle').sizeHalfPoints as
          number | null | undefined
        const base = explicit != null ? explicit / 2 : apply.fallbackSize
        // Do not steal focus if a different control was used before the frame ran.
        apply.editor
          .chain()
          .setMark('docTextStyle', { sizeHalfPoints: Math.round(chain(base) * 2) })
          .run()
      })
    },
  }
}
