import { readSections, type NoteProps, type ParsedDocFull } from '@genoffice/docx-engine'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { SectionInfo } from '@genoffice/docx-engine'

export type NoteKind = 'footnote' | 'endnote'
export type NoteSettingsChange = Partial<NoteProps>

export interface NoteSettingsState {
  settings: Partial<Record<NoteKind, NoteProps>>
  sections: Array<{
    lastBlockIndex: number
    footnote?: NoteProps
    endnote?: NoteProps
  }>
}

type NotePatch = { [K in keyof NoteProps]?: NoteProps[K] | null }

export function noteSettingsFromParsed(parsed: ParsedDocFull): NoteSettingsState {
  return {
    settings: {
      ...(parsed.footnoteSettings ? { footnote: parsed.footnoteSettings } : {}),
      ...(parsed.endnoteSettings ? { endnote: parsed.endnoteSettings } : {}),
    },
    sections: readSections(parsed).map((section) => ({
      lastBlockIndex: section.lastBlockIndex,
      ...(section.settings.footnotePr ? { footnote: section.settings.footnotePr } : {}),
      ...(section.settings.endnotePr ? { endnote: section.settings.endnotePr } : {}),
    })),
  }
}

/**
 * Word for Mac 16.113 ignores a sectPr-level `w:pos` on `w:endnotePr` (verified with
 * section-end probes): the endnote position is a document setting only. Section-level
 * values are kept in the file untouched but never take part in the effective value.
 */
function mergeNoteProps(
  state: NoteSettingsState,
  kind: NoteKind,
  local: NoteProps | undefined,
): NoteProps {
  const merged = { ...state.settings[kind], ...local }
  if (kind === 'endnote') {
    const documentPos = state.settings.endnote?.pos
    if (documentPos === undefined) delete merged.pos
    else merged.pos = documentPos
  }
  return merged
}

export function effectiveNoteSettings(state: NoteSettingsState, kind: NoteKind): NoteProps {
  return mergeNoteProps(state, kind, state.sections.at(-1)?.[kind])
}

export function sectionNoteSettings(
  state: NoteSettingsState,
  kind: NoteKind,
  sectionIndex: number,
): NoteProps {
  return mergeNoteProps(state, kind, state.sections[sectionIndex]?.[kind])
}

/** Whole-document change touches only chosen fields, including local overrides that would shadow settings.xml. */
export function applyWholeDocumentNoteSettings(
  state: NoteSettingsState,
  kind: NoteKind,
  change: NoteSettingsChange,
): NoteSettingsState {
  return {
    settings: {
      ...state.settings,
      [kind]: { ...state.settings[kind], ...change },
    },
    sections: state.sections.map((section) =>
      section[kind] ? { ...section, [kind]: { ...section[kind], ...change } } : section,
    ),
  }
}

export function applySectionNoteSettings(
  state: NoteSettingsState,
  kind: NoteKind,
  sectionIndex: number,
  change: NoteSettingsChange,
): NoteSettingsState {
  // the endnote position is document-wide in Word; a section-scoped request never writes it
  const applied: NoteSettingsChange = { ...change }
  if (kind === 'endnote') delete applied.pos
  if (Object.keys(applied).length === 0) return state
  return {
    ...state,
    sections: state.sections.map((section, index) =>
      index === sectionIndex ? { ...section, [kind]: { ...section[kind], ...applied } } : section,
    ),
  }
}

/** The caret's section, including unsaved paragraphs immediately after a section break. */
export function noteSectionAtPosition(
  doc: PmNode,
  position: number,
  sections: readonly SectionInfo[],
): number | null {
  if (sections.length === 0) return null
  const target = doc.resolve(position).index(0)
  let section = 0
  doc.forEach((block, _offset, index) => {
    if (index > target) return
    const docxIndex = block.attrs.docxIndex
    if (typeof docxIndex === 'number') {
      while (section < sections.length - 1 && docxIndex > sections[section].lastBlockIndex)
        section++
      if (
        index < target &&
        section < sections.length - 1 &&
        docxIndex === sections[section].lastBlockIndex
      )
        section++
    }
  })
  return section
}

function patch(
  current: NoteProps | undefined,
  baseline: NoteProps | undefined,
): NotePatch | undefined {
  const out: NotePatch = {}
  for (const key of ['pos', 'numFmt', 'numStart', 'numRestart'] as const) {
    if (current?.[key] !== baseline?.[key]) {
      // `null` removes a previously explicit value on cross-save Undo.
      ;(out as Record<string, unknown>)[key] = current?.[key] ?? null
    }
  }
  return Object.keys(out).length ? out : undefined
}

/** Compare a restored snapshot to the latest saved baseline, not to the original open. */
export function noteSettingsSaveEdits(
  current: NoteSettingsState,
  baseline: NoteSettingsState,
): {
  noteProps?: { footnote?: NotePatch; endnote?: NotePatch }
  sectionNoteProps?: Array<{ lastBlockIndex: number; footnote?: NotePatch; endnote?: NotePatch }>
} {
  const footnote = patch(current.settings.footnote, baseline.settings.footnote)
  const endnote = patch(current.settings.endnote, baseline.settings.endnote)
  const sectionNoteProps = baseline.sections.flatMap((section, index) => {
    const restored = current.sections[index]
    const footnote = patch(restored?.footnote, section.footnote)
    const endnote = patch(restored?.endnote, section.endnote)
    return footnote || endnote
      ? [
          {
            lastBlockIndex: section.lastBlockIndex,
            ...(footnote ? { footnote } : {}),
            ...(endnote ? { endnote } : {}),
          },
        ]
      : []
  })
  return {
    ...(footnote || endnote
      ? { noteProps: { ...(footnote ? { footnote } : {}), ...(endnote ? { endnote } : {}) } }
      : {}),
    ...(sectionNoteProps.length ? { sectionNoteProps } : {}),
  }
}

export interface NoteReferenceDisplay {
  no: number
  fmt: string
}

/** Re-number footnotes from their measured reference pages without changing native note ids. */
export function numberFootnotesByPage<T extends { id: string; no: number; customMark?: string }>(
  pages: readonly (readonly T[])[],
  state: NoteSettingsState,
  sectionById: Readonly<Record<string, number>>,
): T[][] {
  return pages.map((page) => {
    let previousSection = -1
    let next = 1
    return page.map((item) => {
      const section = sectionById[item.id]
      if (section === undefined) return { ...item }
      const props = sectionNoteSettings(state, 'footnote', section)
      if (props.numRestart !== 'eachPage') return { ...item }
      if (section !== previousSection) next = props.numStart ?? 1
      previousSection = section
      if (item.customMark) return { ...item }
      return { ...item, no: next++ }
    })
  })
}

/** Live body-order numbers and formats, resolved from the reference's own section. */
export function noteReferencesInEditor(
  doc: PmNode,
  sections: readonly SectionInfo[],
  kind: NoteKind,
  state: NoteSettingsState,
): Record<string, NoteReferenceDisplay> {
  const out: Record<string, NoteReferenceDisplay> = {}
  let props = sectionNoteSettings(state, kind, 0)
  let next = props.numStart ?? 1
  let section = 0
  doc.forEach((block) => {
    const docxIndex = block.attrs.docxIndex
    if (typeof docxIndex === 'number') {
      while (section < sections.length - 1 && docxIndex > sections[section].lastBlockIndex) {
        section++
        const following = sectionNoteSettings(state, kind, section)
        if (following.numRestart === 'eachSect' || following.numStart !== props.numStart)
          next = following.numStart ?? 1
        props = following
      }
    }
    block.descendants((node) => {
      if (node.type.name !== 'docNoteRef' || node.attrs.kind !== kind) return true
      const id = String(node.attrs.id)
      if (out[id] === undefined)
        out[id] = {
          no: node.attrs.customMark ? next : next++,
          fmt: props.numFmt ?? (kind === 'footnote' ? 'decimal' : 'lowerRoman'),
        }
      return false
    })
    // A newly inserted paragraph has no original docxIndex. Once the
    // section-break paragraph itself has been visited, such a paragraph
    // belongs to the following section even before it gets saved/reparsed.
    if (
      typeof docxIndex === 'number' &&
      section < sections.length - 1 &&
      docxIndex === sections[section].lastBlockIndex
    ) {
      section++
      const following = sectionNoteSettings(state, kind, section)
      if (following.numRestart === 'eachSect' || following.numStart !== props.numStart)
        next = following.numStart ?? 1
      props = following
    }
  })
  return out
}
