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

/**
 * Design-tab and header/footer state that Word also puts on the undo stack: page color, watermark,
 * theme fonts / colors and the header / footer content (page numbers insert into the latter).
 * Values are held as-is (App owns their types); `undefined` on a restored older snapshot = leave alone.
 */
export interface DocFlags {
  /** settings.xml w:autoHyphenation */
  readonly autoHyphenation: boolean
  /** settings.xml w:mirrorMargins */
  readonly mirrorMargins: boolean
  /** Design → Paragraph Spacing preset id (rewrites the Normal style's spacing); null = the file's own */
  readonly paragraphSpacing: string | null
}

export interface DesignSnapshot {
  readonly docFlags?: DocFlags | undefined
  readonly docFlagsDirty?: boolean | undefined
  readonly pageColor: string | null
  readonly pageColorDirty: boolean
  readonly watermark: string | null
  readonly watermarkDirty: boolean
  readonly watermarkStyle: unknown
  readonly watermarkPicture: unknown
  readonly themeFonts: unknown
  readonly themeFontsDirty: boolean
  readonly themeColors: unknown
  readonly themeColorsDirty: boolean
  readonly header: unknown
  readonly headerDirty: boolean
  readonly footer: unknown
  readonly footerDirty: boolean
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
  readonly design?: DesignSnapshot | undefined
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
    ...(restored.design && live.design
      ? {
          design: {
            ...restored.design,
            docFlagsDirty:
              restored.design.docFlagsDirty ||
              live.design.docFlagsDirty ||
              !same(restored.design.docFlags, live.design.docFlags),
            pageColorDirty:
              restored.design.pageColorDirty ||
              live.design.pageColorDirty ||
              restored.design.pageColor !== live.design.pageColor,
            watermarkDirty:
              restored.design.watermarkDirty ||
              live.design.watermarkDirty ||
              restored.design.watermark !== live.design.watermark ||
              !same(restored.design.watermarkStyle, live.design.watermarkStyle) ||
              !same(restored.design.watermarkPicture, live.design.watermarkPicture),
            themeFontsDirty:
              restored.design.themeFontsDirty ||
              live.design.themeFontsDirty ||
              !same(restored.design.themeFonts, live.design.themeFonts),
            themeColorsDirty:
              restored.design.themeColorsDirty ||
              live.design.themeColorsDirty ||
              !same(restored.design.themeColors, live.design.themeColors),
            headerDirty:
              restored.design.headerDirty ||
              live.design.headerDirty ||
              !same(restored.design.header, live.design.header),
            footerDirty:
              restored.design.footerDirty ||
              live.design.footerDirty ||
              !same(restored.design.footer, live.design.footer),
          },
        }
      : {}),
  }
}
