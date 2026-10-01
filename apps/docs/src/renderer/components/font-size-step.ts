import type { Editor } from '@tiptap/core'
import { Mark } from '@tiptap/pm/model'

/** Coalesce a burst within a frame, without waiting for the user to stop clicking. */
export function createFontSizeStepper() {
  let pending: {
    editor: Editor
    state: Editor['state']
    size: number
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
  return {
    cancel,
    step(editor: Editor, fallbackSize: number, nextSize: (size: number) => number) {
      if (editor.isDestroyed || !editor.isEditable) return
      if (!current() || pending?.editor !== editor) cancel()
      // The ribbon's React snapshot can lag a just-applied mark or a caret change.
      const explicit = editor.getAttributes('docTextStyle').sizeHalfPoints as number | null
      const base = pending?.size ?? (explicit != null ? explicit / 2 : fallbackSize)
      pending = { editor, state: editor.state, size: nextSize(base) }
      if (frame !== null) return
      frame = requestAnimationFrame(() => {
        const valid = current()
        const apply = pending
        frame = null
        pending = null
        if (!valid || !apply) return
        // Do not steal focus if a different control was used before the frame ran.
        apply.editor
          .chain()
          .setMark('docTextStyle', { sizeHalfPoints: Math.round(apply.size * 2) })
          .run()
      })
    },
  }
}
