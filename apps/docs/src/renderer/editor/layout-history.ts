import type { Editor } from '@tiptap/core'
import { closeHistory } from '@tiptap/pm/history'
import type { Transaction } from '@tiptap/pm/state'
import type { SectionInfo, SectionSettings } from '@genoffice/docx-engine'

/**
 * Page layout (margins, orientation, size, different-first-page, odd/even
 * headers, page numbering) lives in App state, outside ProseMirror — so ⌘Z
 * after "type a word, change the margins" undid the word (C7).
 *
 * Each layout change now also writes a snapshot of the whole layout state into
 * the doc node's `layout` attribute (a DocAttrStep). That step sits on the same
 * history stack as text edits; when undo/redo moves the attribute, the App
 * restores its state from the snapshot. The attribute is display state only:
 * the save path reads App state, never this attribute.
 */

export interface SectionLayout {
  readonly settings: SectionSettings
  readonly titlePg?: boolean | undefined
  readonly pageNumberFmt?: string | undefined
  readonly pageNumberStart?: number | undefined
}

export interface LayoutSnapshot {
  readonly section: SectionSettings | null
  readonly sectionDirty: boolean
  readonly sections: readonly SectionLayout[]
  readonly sectionsDirty: readonly number[]
  readonly titlePg: boolean
  readonly titlePgDirty: boolean
  readonly evenOddHf: boolean
  readonly evenOddHfDirty: boolean
  readonly pgNumEdit: { readonly fmt?: string; readonly start?: number } | null
  readonly pgNumDirtySections: readonly number[]
}

export const LAYOUT_ATTR = 'layout'
const RECORD_META = 'layoutHistoryRecord'

export function sectionLayouts(sections: readonly SectionInfo[]): SectionLayout[] {
  return sections.map((s) => ({
    settings: s.settings,
    titlePg: s.titlePg,
    pageNumberFmt: s.pageNumberFmt,
    pageNumberStart: s.pageNumberStart,
  }))
}

/** merge per-section layout back into the live sections (by index) */
export function applySectionLayouts(
  sections: readonly SectionInfo[],
  layouts: readonly SectionLayout[],
): SectionInfo[] {
  return sections.map((s, i) => {
    const l = layouts[i]
    if (!l) return s
    return {
      ...s,
      settings: l.settings,
      titlePg: l.titlePg ?? s.titlePg,
      pageNumberFmt: l.pageNumberFmt,
      pageNumberStart: l.pageNumberStart,
    }
  })
}

/**
 * Record a layout change as one undo step. The first change of a session also
 * pins the pre-change snapshot (outside history), so undo has a state to
 * return to rather than `null`.
 */
export function recordLayoutChange(
  editor: Editor,
  before: LayoutSnapshot,
  after: LayoutSnapshot,
): void {
  if (editor.isDestroyed) return
  const { state } = editor
  if (state.doc.type.spec.attrs?.[LAYOUT_ATTR] === undefined) return
  if (state.doc.attrs[LAYOUT_ATTR] == null) {
    editor.view.dispatch(
      state.tr
        .setDocAttribute(LAYOUT_ATTR, before)
        .setMeta('addToHistory', false)
        .setMeta(RECORD_META, true),
    )
  }
  editor.view.dispatch(
    closeHistory(editor.state.tr).setDocAttribute(LAYOUT_ATTR, after).setMeta(RECORD_META, true),
  )
}

/**
 * The snapshot to restore when a transaction moved the layout attribute for a
 * reason other than recordLayoutChange (undo / redo); null otherwise.
 */
export function restoredLayout(tr: Transaction, previous: unknown): LayoutSnapshot | null {
  if (tr.getMeta(RECORD_META)) return null
  const next = tr.doc.attrs[LAYOUT_ATTR] as LayoutSnapshot | null | undefined
  if (!next || next === previous) return null
  return next
}

const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const union = (a: readonly number[], b: readonly number[]) => [...new Set([...a, ...b])]

/**
 * A restored snapshot's dirty flags describe the moment it was recorded; a
 * save since then cleared the live flags. Undoing past a save must still
 * write the reverted layout, so a flag stays set when either side has it or
 * the value actually changes.
 */
export function withRestoreDirty(restored: LayoutSnapshot, live: LayoutSnapshot): LayoutSnapshot {
  const changedSections = restored.sections.flatMap((s, i) =>
    same(s, live.sections[i]) ? [] : [i],
  )
  const last = Math.max(restored.sections.length - 1, 0)
  return {
    ...restored,
    sectionDirty:
      restored.sectionDirty || live.sectionDirty || !same(restored.section, live.section),
    sectionsDirty: union(
      union(restored.sectionsDirty, live.sectionsDirty),
      changedSections.filter((i) => i !== last),
    ),
    titlePgDirty: restored.titlePgDirty || live.titlePgDirty || restored.titlePg !== live.titlePg,
    evenOddHfDirty:
      restored.evenOddHfDirty || live.evenOddHfDirty || restored.evenOddHf !== live.evenOddHf,
    pgNumDirtySections: union(restored.pgNumDirtySections, live.pgNumDirtySections),
  }
}
