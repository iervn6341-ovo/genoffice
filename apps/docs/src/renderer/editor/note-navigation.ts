import type { Node as PmNode } from '@tiptap/pm/model'

export type NoteKind = 'footnote' | 'endnote'

export interface NoteReferenceTarget {
  pos: number
  id: string
}

/** Find the next editable native note reference, wrapping at the document end. */
export function nextNoteReference(
  doc: PmNode,
  after: number,
  kind: NoteKind,
): NoteReferenceTarget | null {
  let first: NoteReferenceTarget | null = null
  let next: NoteReferenceTarget | null = null
  doc.descendants((node, pos) => {
    if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return true
    const target = { pos, id: String(node.attrs.id) }
    first ??= target
    if (pos > after && next === null) next = target
    return false
  })
  return next ?? first
}
