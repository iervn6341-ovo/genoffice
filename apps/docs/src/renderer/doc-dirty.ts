/**
 * Composite "has unsaved changes" check shared by the close guard, the autosave
 * tick and the crash-recovery push. Only persisted state counts — transient UI
 * state (AI highlights, selection, view modes) must never appear here.
 */
import type { HeaderFooter, SectionInfo, StyleUpsert } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'

import type { PendingNumbering } from './doc-state'
import type { Editor } from '@tiptap/core'

const savedBodies = new WeakMap<Editor, { doc: PmNode; forced: boolean }>()

/** The checkpoint is advanced only by open/new or a successful save, never by Undo. */
export function rememberSavedBody(editor: Editor, doc = editor.state.doc, forced = false): void {
  savedBodies.set(editor, { doc, forced })
}

export function forceBodyDirty(editor: Editor): void {
  const saved = savedBodies.get(editor)
  savedBodies.set(editor, { doc: saved?.doc ?? editor.state.doc, forced: true })
}

export function bodyDirtyAgainstSaved(editor: Editor): boolean | null {
  const saved = savedBodies.get(editor)
  return saved ? saved.forced || persistedBodyChanged(saved.doc, editor.state.doc) : null
}

/** Preserve the original XML used by TOC history without turning that bookkeeping into an edit. */
export function carrySavedTocHistorySource(editor: Editor, source: PmNode, xml: string): void {
  const saved = savedBodies.get(editor)
  const index = source.attrs.docxIndex
  if (!saved || index === null || index === undefined) return
  for (let i = 0; i < saved.doc.childCount; i++) {
    const node = saved.doc.child(i)
    if (node.attrs.docxIndex !== index || persistedBodyChanged(node, source)) continue
    const next = node.type.create({ ...node.attrs, genXml: xml }, node.content, node.marks)
    saved.doc = saved.doc.copy(saved.doc.content.replaceChild(i, next))
    return
  }
}

/** Empty cursor landing below a final table, until text or paragraph formatting is authored. */
export function isNavigationPlaceholder(node: PmNode): boolean {
  if (node.type.name !== 'docParagraph' || !node.attrs.navigationOnly || node.content.size > 0)
    return false
  return Object.keys(node.attrs).every(
    (key) =>
      key === 'navigationOnly' ||
      key === 'caretMarks' ||
      node.attrs[key] === node.type.spec.attrs?.[key]?.default,
  )
}

/** Compare persisted content, sharing unchanged PM subtrees instead of serializing the whole document. */
export function persistedBodyChanged(before: PmNode, after: PmNode): boolean {
  const equal = (a: PmNode, b: PmNode): boolean => {
    if (a === b) return true
    if (
      a.type !== b.type ||
      a.text !== b.text ||
      a.marks.length !== b.marks.length ||
      a.marks.some((mark, index) => !mark.eq(b.marks[index]!))
    )
      return false
    for (const key of new Set([...Object.keys(a.attrs), ...Object.keys(b.attrs)])) {
      // Caret format memory and the temporary landing flag never reach DOCX.
      // A note reference persists its ID; its display number comes from pagination.
      if (
        key === 'caretMarks' ||
        key === 'navigationOnly' ||
        // Page layout has its own value/dirty checkpoint outside the PM body.
        (a.type.name === 'doc' && key === 'layout') ||
        (a.type.name === 'docNoteRef' && key === 'num')
      )
        continue
      if (a.attrs[key] === b.attrs[key]) continue
      if (
        key === 'fieldDisplay' &&
        a.attrs[key]?.kind === 'tocLine' &&
        b.attrs[key]?.kind === 'tocLine'
      ) {
        const { right: _aPage, ...aField } = a.attrs[key]
        const { right: _bPage, ...bField } = b.attrs[key]
        if (JSON.stringify(aField) === JSON.stringify(bField)) continue
      }
      if (
        typeof a.attrs[key] !== 'object' ||
        typeof b.attrs[key] !== 'object' ||
        JSON.stringify(a.attrs[key]) !== JSON.stringify(b.attrs[key])
      )
        return false
    }
    let i = 0
    let j = 0
    while (i < a.childCount || j < b.childCount) {
      while (i < a.childCount && isNavigationPlaceholder(a.child(i))) i++
      while (j < b.childCount && isNavigationPlaceholder(b.child(j))) j++
      if (i === a.childCount || j === b.childCount) return i === a.childCount && j === b.childCount
      if (!equal(a.child(i++), b.child(j++))) return false
    }
    return true
  }
  return !equal(before, after)
}

export interface DocDirtyState {
  dirtyRef: { current: boolean }
  sectionDirty: boolean
  sectionsDirty: readonly number[]
  trailingStartType: unknown
  pageColorDirty: boolean
  headerDirty: boolean
  footerDirty: boolean
  hfVariantsDirty: readonly unknown[]
  sectionHfEdits: Record<string, unknown>
  sectionHfLinkEdits: Record<string, unknown>
  sectionHfRefEdits: Record<string, unknown>
  pgNumEdit: unknown
  pgNumDirtySections: readonly number[]
  numberingDirty: boolean
  styleUpserts: Record<string, unknown>
  titlePgDirty: boolean
  evenOddHfDirty: boolean
  watermarkDirty: boolean
  inksDirty: boolean
  notesDirty: boolean
  sourcesDirty: boolean
  zoteroDocumentDataDirty: boolean
  themeFontsDirty: boolean
  themeColorsDirty: boolean
  docFlagsDirty?: boolean
  commentsDirty: boolean
  protectionDirty: boolean
  writeProtectionDirty: boolean
  removePersonalInfoDirty: boolean
}

export function isDocDirty(s: DocDirtyState): boolean {
  return (
    s.dirtyRef.current ||
    s.sectionDirty ||
    s.sectionsDirty.length > 0 ||
    s.trailingStartType !== null ||
    s.pageColorDirty ||
    s.headerDirty ||
    s.footerDirty ||
    s.hfVariantsDirty.length > 0 ||
    Object.keys(s.sectionHfEdits).length > 0 ||
    Object.keys(s.sectionHfLinkEdits).length > 0 ||
    Object.keys(s.sectionHfRefEdits).length > 0 ||
    s.pgNumEdit !== null ||
    s.pgNumDirtySections.length > 0 ||
    s.numberingDirty ||
    Object.keys(s.styleUpserts).length > 0 ||
    s.titlePgDirty ||
    s.evenOddHfDirty ||
    s.watermarkDirty ||
    s.inksDirty ||
    s.notesDirty ||
    s.sourcesDirty ||
    s.zoteroDocumentDataDirty ||
    s.themeFontsDirty ||
    s.themeColorsDirty ||
    s.docFlagsDirty === true ||
    s.commentsDirty ||
    s.protectionDirty ||
    s.writeProtectionDirty ||
    s.removePersonalInfoDirty
  )
}

/** The edit-tracking setters resetCrossDocEditState clears (structural subset
 *  of FileActionContext, so the helper stays unit-testable without the editor). */
export interface CrossDocEditStateSink {
  setSectionsDirty: (value: number[]) => void
  setTrailingStartType: (value: SectionInfo['startType'] | null) => void
  setSectionHfEdits: (value: Record<string, HeaderFooter>) => void
  setSectionHfLinkEdits: (value: Record<string, true>) => void
  setSectionHfRefEdits: (value: Record<string, string>) => void
  setPgNumEdit: (
    value: {
      fmt?: string
      start?: number
      chapStyle?: number | null
      chapSep?: string | null
    } | null,
  ) => void
  setPgNumDirtySections: (value: number[]) => void
  setPendingNumbering: (value: PendingNumbering) => void
  setStyleUpserts: (value: Record<string, StyleUpsert>) => void
}

/**
 * Clear the section/numbering/style edit-tracking states. The save path calls
 * this once the bytes land; a document swap (open/new) must call it too, or
 * doc A's edits leak into pristine doc B — tripping the close guard and
 * mis-applying section indices, numbering restarts and style upserts on B's
 * next save. Single source of truth so the call sites cannot drift apart.
 */
export function resetCrossDocEditState(sink: CrossDocEditStateSink): void {
  sink.setSectionsDirty([])
  sink.setTrailingStartType(null)
  sink.setSectionHfEdits({})
  sink.setSectionHfLinkEdits({})
  sink.setSectionHfRefEdits({})
  sink.setPgNumEdit(null)
  sink.setPgNumDirtySections([])
  sink.setPendingNumbering({ newDefs: [], restartNums: [] })
  sink.setStyleUpserts({})
}
