import { Extension } from '@tiptap/core'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import { Decoration, DecorationSet } from '@tiptap/pm/view'
import type { EditorView } from '@tiptap/pm/view'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { SectionInfo } from '@genoffice/docx-engine'
import type { NoteSettingsState } from './note-settings'

/**
 * Section-end endnotes (w:endnotePr/w:pos = sectEnd).
 *
 * Word ground truth (Word for Mac 16.113, docs/qa/office-mac-field-evidence-2026-10-13/
 * docs-native-notes/section-end-endnotes/):
 *  - only the document-level w:endnotePr in settings.xml selects the position; a
 *    sectPr-level w:pos on w:endnotePr is ignored by Word (all notes stay at the end);
 *  - with sectEnd every section's notes follow that section's last body paragraph with
 *    the same offset as the document-end area, flow onto later pages when they do not
 *    fit, and the next section starts after them (next page for a nextPage break, on
 *    the same page right after the notes for a continuous break);
 *  - numbering stays continuous across sections unless a restart rule says otherwise.
 *
 * The earlier sections' areas are real in-flow widget rows (one widget per entry) so
 * the page slicer measures them like body blocks; the last section keeps the existing
 * document-end area.
 */

export type EndnotePlacement = 'docEnd' | 'sectEnd'

/** Word only honours the settings.xml position; sectPr-level overrides never move notes. */
export function endnotePlacement(state: NoteSettingsState | null | undefined): EndnotePlacement {
  return state?.settings.endnote?.pos === 'sectEnd' ? 'sectEnd' : 'docEnd'
}

export interface EndnoteSectionMap {
  /** endnote id -> section index of its first reference (body order) */
  sectionOf: Record<string, number>
  /** per section (except the last): doc position right after the section's closing block, or null when the break block is gone */
  breakPos: Array<number | null>
}

/** Walk the live ProseMirror document once; unsaved paragraphs after a break join the next section. */
export function endnoteSectionMap(
  doc: PmNode,
  sections: readonly SectionInfo[],
): EndnoteSectionMap {
  const last = Math.max(0, sections.length - 1)
  const sectionOf: Record<string, number> = {}
  const breakPos: Array<number | null> = Array.from({ length: last }, () => null)
  let section = 0
  doc.forEach((block, offset) => {
    const docxIndex = block.attrs.docxIndex
    if (typeof docxIndex === 'number') {
      while (section < last && docxIndex > sections[section].lastBlockIndex) section++
    }
    block.descendants((node) => {
      if (node.type.name !== 'docNoteRef' || node.attrs.kind !== 'endnote') return true
      const id = String(node.attrs.id)
      if (sectionOf[id] === undefined) sectionOf[id] = section
      return false
    })
    if (
      typeof docxIndex === 'number' &&
      section < last &&
      docxIndex === sections[section].lastBlockIndex
    ) {
      breakPos[section] = offset + block.nodeSize
      section++
    }
  })
  return { sectionOf, breakPos }
}

export interface EndnoteAreaGroup<T> {
  section: number
  pos: number
  items: T[]
}

/**
 * Split endnote entries into per-section areas plus the document-end tail. A section whose
 * break block was deleted (no position) merges into the next section that still ends at a
 * break; entries with no body reference stay in the tail (Word keeps orphans at the end).
 */
export function groupEndnotesBySection<T extends { id: string }>(
  items: readonly T[],
  map: EndnoteSectionMap,
): { areas: Array<EndnoteAreaGroup<T>>; tail: T[] } {
  const count = map.breakPos.length
  const owner = (section: number): number | null => {
    for (let s = section; s < count; s++) if (map.breakPos[s] !== null) return s
    return null
  }
  const bySection = new Map<number, T[]>()
  const tail: T[] = []
  for (const item of items) {
    const section = map.sectionOf[item.id]
    const target = section === undefined ? null : owner(section)
    if (target === null) tail.push(item)
    else bySection.set(target, [...(bySection.get(target) ?? []), item])
  }
  const areas = [...bySection.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([section, group]) => ({ section, pos: map.breakPos[section]!, items: group }))
  return { areas, tail }
}

/* ---------------------------------------------------------------- widgets */

export interface SectionEndnoteRow {
  id: string
  marker: string
  noRefMark?: boolean
  text: string
  /** full row height in px, separator included for the first row */
  height: number
  withSeparator: boolean
  lineHeightPx?: number
  fontSizePt?: number
  fontFamily?: string
}

export interface SectionEndnoteArea {
  /** stable id (section index) */
  key: string
  pos: number
  rows: SectionEndnoteRow[]
}

export interface SectionEndnoteHandlers {
  onEdit: (id: string) => void
  onDelete: (id: string) => void
  onConvert: (id: string) => void
}

export interface SectionEndnoteLabels {
  edit: string
  delete: string
  convert: string
}

/** Row widgets sort before every page gap (side -1) at the same position. */
const ROW_SIDE_BASE = -100000
export const sectionEndnoteRowSide = (row: number): number => ROW_SIDE_BASE + row * 2
/** side that places a page gap right before row `row` of an area */
export const sectionEndnoteGapSide = (row: number): number => ROW_SIDE_BASE + row * 2 - 1

interface AreaState {
  sig: string
  set: DecorationSet
}

const areaPluginKey = new PluginKey<AreaState>('sectionEndnoteAreas')
const noop = () => {}
const handlersRef: { current: SectionEndnoteHandlers } = {
  current: { onEdit: noop, onDelete: noop, onConvert: noop },
}
let labelsRef: SectionEndnoteLabels = { edit: '', delete: '', convert: '' }

export const SectionEndnoteAreasExtension = Extension.create({
  name: 'sectionEndnoteAreas',

  addProseMirrorPlugins() {
    return [
      new Plugin<AreaState>({
        key: areaPluginKey,
        state: {
          init: () => ({ sig: '', set: DecorationSet.empty }),
          apply(tr, value) {
            const next = tr.getMeta(areaPluginKey) as AreaState | undefined
            if (next) return next
            return tr.docChanged
              ? { sig: value.sig, set: value.set.map(tr.mapping, tr.doc) }
              : value
          },
        },
        props: {
          decorations: (state) => areaPluginKey.getState(state)?.set,
        },
      }),
    ]
  },
})

function buildRow(area: SectionEndnoteArea, index: number): HTMLElement {
  const row = area.rows[index]
  const el = document.createElement('div')
  el.className = 'page-section-endnote-row'
  el.contentEditable = 'false'
  el.dataset.areaKey = area.key
  el.dataset.row = String(index)
  el.dataset.noteId = row.id
  el.style.height = `${row.height}px`
  if (row.lineHeightPx) el.style.lineHeight = `${row.lineHeightPx}px`
  if (row.fontSizePt) el.style.fontSize = `${row.fontSizePt}pt`
  if (row.fontFamily) el.style.fontFamily = row.fontFamily
  if (row.withSeparator) {
    const sep = document.createElement('div')
    sep.className = 'page-section-endnote-sep'
    el.appendChild(sep)
  }
  const line = document.createElement('div')
  line.className = 'page-note page-section-endnote-line'
  if (!row.noRefMark) {
    const sup = document.createElement('sup')
    sup.textContent = row.marker
    line.appendChild(sup)
  }
  const text = document.createElement('span')
  text.className = 'page-note-text'
  text.textContent = row.text
  line.appendChild(text)
  const button = (glyph: string, label: string, run: (id: string) => void) => {
    const b = document.createElement('button')
    b.type = 'button'
    b.className = 'page-note-btn page-section-endnote-btn'
    b.textContent = glyph
    b.dataset.tip = label
    b.setAttribute('aria-label', label)
    b.addEventListener('mousedown', (event) => event.preventDefault())
    b.addEventListener('click', () => run(row.id))
    line.appendChild(b)
  }
  button('✎', labelsRef.edit, (id) => handlersRef.current.onEdit(id))
  button('×', labelsRef.delete, (id) => handlersRef.current.onDelete(id))
  button('⇄', labelsRef.convert, (id) => handlersRef.current.onConvert(id))
  el.appendChild(line)
  return el
}

/** Replace the section-end areas; returns true when a transaction was dispatched. */
export function setSectionEndnoteAreas(
  view: EditorView,
  areas: readonly SectionEndnoteArea[],
  handlers: SectionEndnoteHandlers,
  labels: SectionEndnoteLabels,
): boolean {
  handlersRef.current = handlers
  labelsRef = labels
  const prev = areaPluginKey.getState(view.state)
  if (!prev) return false
  const valid = areas.filter((a) => a.rows.length > 0 && a.pos <= view.state.doc.content.size)
  const sig = JSON.stringify([labels, valid])
  // an edit above shifts the mapped decoration positions: they are part of the identity
  const mapped = prev.set
    .find()
    .map((d) => d.from)
    .join(',')
  const wanted = valid.flatMap((a) => a.rows.map(() => a.pos)).join(',')
  if (prev.sig === sig && mapped === wanted) return false
  const decos: Decoration[] = []
  for (const area of valid) {
    area.rows.forEach((row, index) => {
      decos.push(
        Decoration.widget(area.pos, () => buildRow(area, index), {
          side: sectionEndnoteRowSide(index),
          key: `section-endnote-${area.key}-${index}-${row.id}-${Math.round(row.height)}-${row.marker}-${row.text.length}`,
          areaKey: area.key,
        }),
      )
    })
  }
  const next: AreaState = { sig, set: DecorationSet.create(view.state.doc, decos) }
  view.dispatch(view.state.tr.setMeta(areaPluginKey, next).setMeta('addToHistory', false))
  return true
}

/** Current doc position of an area (decorations map along edits), for page-gap anchoring. */
export function sectionEndnoteAreaPos(view: EditorView, key: string): number | null {
  const state = areaPluginKey.getState(view.state)
  const found = state?.set.find(undefined, undefined, (spec) => spec.areaKey === key)
  return found && found.length > 0 ? found[0].from : null
}
