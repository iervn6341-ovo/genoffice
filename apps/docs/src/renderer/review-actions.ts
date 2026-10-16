/**
 * Review-tab actions: footnotes/endnotes, comments, revisions, ink
 * annotations and compare (document protection lives in ProtectDialog +
 * App.applyProtectDialog). Extracted from App.tsx; the App component passes a
 * ReviewContext built fresh per call so state never goes stale.
 */
import type { Editor } from '@tiptap/core'
import type { Transaction } from '@tiptap/pm/state'
import {
  nextNoteId,
  parseDocx,
  wordSymbolForMark,
  type CommentInfo,
  type NoteInfo,
} from '@genoffice/docx-engine'
import type { Dispatch, SetStateAction } from 'react'
import type { DocState } from './doc-state'
import {
  addCommentToRange,
  addCommentToSelection,
  addReplyToCommentRange,
  nextCommentId,
  removeCommentFromDoc,
  wordRangeAtCaret,
} from './editor/comments'
import { blockTexts, compareParagraphs, type CompareEntry } from './editor/compare'
import { pendingCommentPluginKey } from './editor/extensions'
import { protectedNoteMarkBlock } from './ai/note-ops'
import type { InkAnnotation } from './editor/ink'
import {
  acceptAllRevisions,
  acceptCurrentRevision,
  rejectAllRevisions,
  rejectCurrentRevision,
} from './editor/revisions'
import { t } from './i18n/locale'
import { currentNoteNumFmt } from './note-format'
import { planNoteConversion } from './editor/note-conversion'

/** open the note text dialog; id set = editing an existing note */
export interface NotePrompt {
  kind: 'footnote' | 'endnote'
  id?: string
}

/** The App state the review actions need; built fresh per call. */
export interface ReviewContext {
  editor: Editor | null
  doc: DocState | null
  dirtyRef: { current: boolean }
  setStatus: (status: string) => void
  notePrompt: NotePrompt | null
  setNotePrompt: (value: NotePrompt | null) => void
  footnotes: NoteInfo[]
  endnotes: NoteInfo[]
  setFootnotes: Dispatch<SetStateAction<NoteInfo[]>>
  setEndnotes: Dispatch<SetStateAction<NoteInfo[]>>
  setNotesDirty: (dirty: boolean) => void
  recordNoteChange?: (
    kind: 'footnote' | 'endnote',
    notes: NoteInfo[],
    markerEdit?: (tr: Transaction) => Transaction,
  ) => void
  recordNoteConversion?: (
    before: { footnotes: NoteInfo[]; endnotes: NoteInfo[] },
    after: { footnotes: NoteInfo[]; endnotes: NoteInfo[] },
    markerEdit: (tr: Transaction) => Transaction,
  ) => void
  normalizeNoteMarkers?: (tr: Transaction, kind: 'footnote' | 'endnote') => Transaction
  comments: CommentInfo[]
  setComments: Dispatch<SetStateAction<CommentInfo[]>>
  setCommentsDirty: (dirty: boolean) => void
  setCommentComposing: (composing: boolean) => void
  setShowComments: (show: boolean) => void
  setInkAnnotations: Dispatch<SetStateAction<InkAnnotation[]>>
  setInksDirty: (dirty: boolean) => void
  setCompareResult: (value: { otherName: string; entries: CompareEntry[] } | null) => void
}

// ---- References: footnotes / endnotes ----

/** dialog submit: create a new note (+ caret marker) or update an existing one */
export function submitNote(ctx: ReviewContext, text: string, customMark?: string): void {
  if (!ctx.notePrompt || !ctx.editor) return
  const { kind, id } = ctx.notePrompt
  const list = kind === 'footnote' ? ctx.footnotes : ctx.endnotes
  const setList = kind === 'footnote' ? ctx.setFootnotes : ctx.setEndnotes
  const mark = customMark?.trim() || undefined
  if (id !== undefined) {
    const original = list.find((note) => note.id === id)
    if (!original) return
    const requestedMark = customMark === undefined ? original.customMark : mark
    const changingMark = requestedMark !== original.customMark
    const requestedSymbol = requestedMark
      ? changingMark
        ? wordSymbolForMark(requestedMark)
        : original.customMarkSymbol
      : undefined
    if (changingMark && original.protectedContent) {
      ctx.setStatus(t('appNoteProtectedError'))
      return
    }
    let hasEditableReference = false
    ctx.editor.state.doc.descendants((node) => {
      if (
        node.type.name === 'docNoteRef' &&
        node.attrs.kind === kind &&
        String(node.attrs.id) === id
      )
        hasEditableReference = true
    })
    if (changingMark && !hasEditableReference) {
      ctx.setStatus(t('appNoteReferenceProtectedError'))
      return
    }
    const next = list.map((note) => {
      if (note.id !== id) return note
      const { customMark: _oldMark, customMarkSymbol: _oldSymbol, ...rest } = note
      return {
        ...rest,
        text,
        ...(requestedMark ? { customMark: requestedMark } : {}),
        ...(requestedSymbol ? { customMarkSymbol: requestedSymbol } : {}),
      }
    })
    const updateMark = changingMark
      ? (tr: Transaction) => {
          tr.doc.descendants((node, pos) => {
            if (
              node.type.name !== 'docNoteRef' ||
              node.attrs.kind !== kind ||
              String(node.attrs.id) !== id
            )
              return true
            tr.setNodeMarkup(pos, undefined, {
              ...node.attrs,
              customMark: requestedMark ?? null,
              customMarkSymbol: requestedSymbol ?? null,
            })
            return false
          })
          return ctx.normalizeNoteMarkers?.(tr, kind) ?? tr
        }
      : undefined
    if (ctx.recordNoteChange) ctx.recordNoteChange(kind, next, updateMark)
    else {
      setList(next)
      if (updateMark) ctx.editor.view.dispatch(updateMark(ctx.editor.state.tr))
    }
  } else {
    const newId = nextNoteId(list)
    const symbol = mark ? wordSymbolForMark(mark) : undefined
    const next = [
      ...list,
      {
        id: newId,
        text,
        ...(mark ? { customMark: mark } : {}),
        ...(symbol ? { customMarkSymbol: symbol } : {}),
      },
    ]
    if (ctx.recordNoteChange) {
      const marker = ctx.editor.schema.nodes.docNoteRef.create({
        kind,
        id: newId,
        num: list.length + 1,
        fmt: currentNoteNumFmt(kind),
        customMark: mark ?? null,
        customMarkSymbol: symbol ?? null,
      })
      ctx.recordNoteChange(kind, next, (tr) => {
        tr.replaceSelectionWith(marker)
        return ctx.normalizeNoteMarkers?.(tr, kind) ?? tr
      })
      ctx.editor.view.focus()
    } else {
      setList(next)
      ctx.editor
        .chain()
        .focus()
        .insertContent({
          type: 'docNoteRef',
          attrs: {
            kind,
            id: newId,
            num: list.length + 1,
            fmt: currentNoteNumFmt(kind),
            customMark: mark ?? null,
            customMarkSymbol: symbol ?? null,
          },
        } as never)
        .run()
    }
  }
  ctx.setNotesDirty(true)
}

/** delete a note, remove its in-text marker, renumber the remaining markers */
export function deleteNote(ctx: ReviewContext, kind: 'footnote' | 'endnote', id: string): void {
  const list = kind === 'footnote' ? ctx.footnotes : ctx.endnotes
  if (!list.some((note) => note.id === id)) return
  const next = list.filter((n) => n.id !== id)
  if (!ctx.editor) return
  const editor = ctx.editor
  const protectedBlock = protectedNoteMarkBlock(
    editor.state.doc,
    (block) => {
      if (typeof block.attrs.genXml === 'string' && block.attrs.genXml) return block.attrs.genXml
      const index = block.attrs.docxIndex
      if (typeof index !== 'number') return ''
      return ctx.doc?.parsed.blocks.find((source) => source.docxIndex === index)?.originalXml ?? ''
    },
    kind,
    id,
  )
  const removals: Array<{ pos: number; size: number }> = []
  const renumbers: Array<{ pos: number; attrs: Record<string, unknown> }> = []
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return
    if (String(node.attrs.id) === id) {
      removals.push({ pos, size: node.nodeSize })
    } else {
      const num = next.findIndex((n) => n.id === String(node.attrs.id)) + 1
      if (num > 0 && num !== node.attrs.num) {
        renumbers.push({ pos, attrs: { ...node.attrs, num } })
      }
    }
  })
  // A protected original block can still contain a native reference in its
  // XML. Deleting only the note part would leave that reference dangling.
  if (protectedBlock !== null || removals.length === 0) {
    ctx.setStatus(t('appNoteReferenceProtectedError'))
    return
  }
  const editMarkers = (tr: Transaction) => {
    for (const { pos, attrs } of renumbers) tr.setNodeMarkup(pos, undefined, attrs)
    for (const { pos, size } of removals.reverse()) tr.delete(pos, pos + size)
    return tr
  }
  if (ctx.recordNoteChange) ctx.recordNoteChange(kind, next, editMarkers)
  else {
    ;(kind === 'footnote' ? ctx.setFootnotes : ctx.setEndnotes)(next)
    const tr = editMarkers(editor.state.tr)
    if (tr.docChanged) editor.view.dispatch(tr)
  }
  ctx.setNotesDirty(true)
}

/** Convert one editable note and its body reference in the same undo step. */
export function convertNote(ctx: ReviewContext, kind: 'footnote' | 'endnote', id: string): void {
  const editor = ctx.editor
  const parsed = ctx.doc?.parsed
  if (!editor?.isEditable || !parsed || !ctx.recordNoteConversion) return
  const affected = [
    { kind, id },
    ...(kind === 'footnote' ? ctx.endnotes : ctx.footnotes).map((note) => ({
      kind: kind === 'footnote' ? ('endnote' as const) : ('footnote' as const),
      id: note.id,
    })),
  ]
  const rawBlock = (block: { attrs: Record<string, unknown> }) => {
    if (typeof block.attrs.genXml === 'string' && block.attrs.genXml) return block.attrs.genXml
    const index = block.attrs.docxIndex
    if (typeof index !== 'number') return ''
    return parsed.blocks.find((source) => source.docxIndex === index)?.originalXml ?? ''
  }
  if (
    affected.some(
      (note) => protectedNoteMarkBlock(editor.state.doc, rawBlock, note.kind, note.id) !== null,
    )
  ) {
    ctx.setStatus(t('appNoteReferenceProtectedError'))
    return
  }
  const plan = planNoteConversion(editor.state.doc, parsed, ctx.footnotes, ctx.endnotes, kind, id)
  if (!plan) {
    ctx.setStatus(t('appNoteConversionUnsupported'))
    return
  }
  ctx.recordNoteConversion(plan.before, plan.after, (tr) => {
    for (const edit of plan.markers) {
      const node = tr.doc.nodeAt(edit.pos)
      if (node?.type.name !== 'docNoteRef') continue
      tr.setNodeMarkup(edit.pos, undefined, {
        ...node.attrs,
        kind: edit.kind,
        id: edit.id,
        rawRPr: edit.rawRPr,
      })
    }
    tr = ctx.normalizeNoteMarkers?.(tr, 'footnote') ?? tr
    return ctx.normalizeNoteMarkers?.(tr, 'endnote') ?? tr
  })
  ctx.setNotesDirty(true)
}

// ---- Review: comments / revisions / compare / protection ----

/** keeps the picked range highlighted while the composer holds focus */
export function setPendingCommentRange(
  ctx: ReviewContext,
  range: { from: number; to: number } | null,
): void {
  if (!ctx.editor) return
  ctx.editor.view.dispatch(ctx.editor.state.tr.setMeta(pendingCommentPluginKey, range))
}

export function cancelNewComment(ctx: ReviewContext): void {
  ctx.setCommentComposing(false)
  setPendingCommentRange(ctx, null)
}

/** New comment: open the pane with the composer; the mark is applied on submit */
export function startNewComment(ctx: ReviewContext): void {
  const editor = ctx.editor
  if (!editor) return
  if (editor.state.selection.empty) {
    // Word anchors on the word under a collapsed caret rather than refusing
    const word = wordRangeAtCaret(editor)
    if (!word) {
      ctx.setStatus(t('appSelectTextToComment'))
      return
    }
    editor.commands.setTextSelection(word)
  }
  const { from, to } = editor.state.selection
  setPendingCommentRange(ctx, { from, to })
  ctx.setShowComments(true)
  ctx.setCommentComposing(true)
}

export function submitNewComment(ctx: ReviewContext, text: string): void {
  if (!ctx.editor) return
  setPendingCommentRange(ctx, null)
  const id = nextCommentId(ctx.comments)
  if (!addCommentToSelection(ctx.editor, id)) {
    ctx.setStatus(t('appCommentSelectionLost'))
    ctx.setCommentComposing(false)
    return
  }
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  ctx.setComments((prev) => [...prev, { id, author: 'User', date: now, text }])
  ctx.setCommentsDirty(true)
  ctx.setCommentComposing(false)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentAdded'))
}

/** New thread on an explicit range (AI add_comment); the new id, null when the range holds no text */
export function addCommentAt(
  ctx: ReviewContext,
  range: { from: number; to: number },
  text: string,
  author: string,
  initials?: string,
): string | null {
  if (!ctx.editor) return null
  const id = nextCommentId(ctx.comments)
  if (!addCommentToRange(ctx.editor, range.from, range.to, id)) return null
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  ctx.setComments((prev) => [
    ...prev,
    { id, author, date: now, text, ...(initials ? { initials } : {}) },
  ])
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentAdded'))
  return id
}

/** Reply to a comment: the new entry carries parentId; the anchor shares the parent comment's range */
export function replyToComment(
  ctx: ReviewContext,
  parentId: string,
  text: string,
  author = 'User',
): boolean {
  if (!ctx.editor) return false
  const id = nextCommentId(ctx.comments)
  if (!addReplyToCommentRange(ctx.editor, parentId, id)) {
    ctx.setStatus(t('appCommentAnchorGone'))
    return false
  }
  const now = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z')
  ctx.setComments((prev) => [...prev, { id, author, date: now, text, parentId }])
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentReplied'))
  return true
}

/** Word: comment text edits in place; the author, date and anchor stay */
export function editComment(ctx: ReviewContext, id: string, text: string): void {
  ctx.setComments((prev) => prev.map((c) => (c.id === id ? { ...c, text } : c)))
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appCommentEdited'))
}

/** Resolve/reopen: the whole thread (parent + replies) gets done set together */
export function resolveComment(ctx: ReviewContext, id: string, done: boolean): void {
  ctx.setComments((prev) =>
    prev.map((c) => (c.id === id || c.parentId === id ? { ...c, done } : c)),
  )
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(done ? t('appCommentResolvedMsg') : t('appCommentReopenedMsg'))
}

export function deleteComment(ctx: ReviewContext, id: string): void {
  if (!ctx.editor) return
  // cascade: deleting a parent comment also deletes its replies (including their anchor marks)
  const victims = [id, ...ctx.comments.filter((c) => c.parentId === id).map((c) => c.id)]
  for (const v of victims) removeCommentFromDoc(ctx.editor, v)
  ctx.setComments((prev) => prev.filter((c) => !victims.includes(c.id)))
  ctx.setCommentsDirty(true)
  ctx.dirtyRef.current = true
}

export function handleRevision(
  ctx: ReviewContext,
  action: 'accept' | 'reject',
  all: boolean,
): void {
  if (!ctx.editor) return
  if (all) {
    if (action === 'accept') acceptAllRevisions(ctx.editor)
    else rejectAllRevisions(ctx.editor)
    ctx.setStatus(action === 'accept' ? t('appAllRevisionsAccepted') : t('appAllRevisionsRejected'))
  } else {
    const ok =
      action === 'accept' ? acceptCurrentRevision(ctx.editor) : rejectCurrentRevision(ctx.editor)
    if (!ok) ctx.setStatus(t('appNoRevisionsToHandle'))
  }
  ctx.dirtyRef.current = true
}

// ---- Draw: overlay annotations ----

export function addInk(ctx: ReviewContext, annotation: InkAnnotation): void {
  ctx.setInkAnnotations((prev) => [...prev, annotation])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
}

export function removeInks(ctx: ReviewContext, ids: string[]): void {
  const gone = new Set(ids)
  ctx.setInkAnnotations((prev) => prev.filter((a) => !gone.has(a.id)))
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
}

export function clearInks(ctx: ReviewContext): void {
  ctx.setInkAnnotations([])
  ctx.setInksDirty(true)
  ctx.dirtyRef.current = true
  ctx.setStatus(t('appInksCleared'))
}

/** Compare: pick a second .docx and diff it against the open document */
export async function compareWithFile(ctx: ReviewContext): Promise<void> {
  if (!ctx.doc) return
  const other = await window.desktop.openDocx()
  if (!other) return
  // password-protected comparison target: not wired through the decrypt prompt (yet)
  if ('needsPassword' in other) {
    ctx.setStatus(t('appCompareFailed', { error: t('appDocPwdTitle') }))
    return
  }
  try {
    const otherParsed = await parseDocx(new Uint8Array(other.data))
    const entries = compareParagraphs(
      blockTexts(ctx.doc.parsed.blocks),
      blockTexts(otherParsed.blocks),
    )
    ctx.setCompareResult({ otherName: other.name, entries })
  } catch (err) {
    ctx.setStatus(t('appCompareFailed', { error: String(err) }))
  }
}
