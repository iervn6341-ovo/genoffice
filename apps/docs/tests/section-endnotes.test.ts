import { Editor, type JSONContent } from '@tiptap/core'
import { describe, expect, it, vi } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  endnotePlacement,
  endnoteSectionMap,
  groupEndnotesBySection,
  sectionEndnoteAreaPos,
  setSectionEndnoteAreas,
} from '../src/renderer/editor/section-endnotes'
import {
  applySectionNoteSettings,
  effectiveNoteSettings,
  sectionNoteSettings,
  type NoteSettingsState,
} from '../src/renderer/editor/note-settings'

const para = (
  docxIndex: number,
  content: JSONContent[] = [{ type: 'text', text: `p${docxIndex}` }],
) => ({
  type: 'docParagraph',
  attrs: { docxIndex },
  content,
})
const ref = (id: string): JSONContent => ({ type: 'docNoteRef', attrs: { kind: 'endnote', id } })
const sections = [
  { lastBlockIndex: 1 },
  { lastBlockIndex: 3 },
  { lastBlockIndex: 5 },
] as SectionInfo[]

function build(content: JSONContent[]) {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content },
  })
}

describe('section-end endnote grouping', () => {
  it('assigns endnotes to the section of their first reference and records break positions', () => {
    const editor = build([
      para(0, [{ type: 'text', text: 'a' }, ref('1')]),
      para(1),
      para(2, [ref('2')]),
      para(3),
      para(4, [ref('3')]),
    ])
    try {
      const map = endnoteSectionMap(editor.state.doc, sections)
      expect(map.sectionOf).toEqual({ '1': 0, '2': 1, '3': 2 })
      const first = editor.state.doc.child(0).nodeSize + editor.state.doc.child(1).nodeSize
      expect(map.breakPos[0]).toBe(first)
      expect(map.breakPos[1]).toBe(
        first + editor.state.doc.child(2).nodeSize + editor.state.doc.child(3).nodeSize,
      )
      const grouped = groupEndnotesBySection(
        [{ id: '1' }, { id: '2' }, { id: '3' }, { id: '9' }],
        map,
      )
      expect(grouped.areas.map((a) => [a.section, a.items.map((i) => i.id)])).toEqual([
        [0, ['1']],
        [1, ['2']],
      ])
      // section 2 is the last one: its notes and the orphan stay in the document-end tail
      expect(grouped.tail.map((i) => i.id)).toEqual(['3', '9'])
    } finally {
      editor.destroy()
    }
  })

  it('merges a section whose break block was deleted into the next break', () => {
    const editor = build([para(0, [ref('1')]), para(2, [ref('2')]), para(3), para(4)])
    try {
      const map = endnoteSectionMap(editor.state.doc, sections)
      expect(map.breakPos[0]).toBeNull()
      expect(map.breakPos[1]).not.toBeNull()
      const grouped = groupEndnotesBySection([{ id: '1' }, { id: '2' }], map)
      expect(grouped.areas).toHaveLength(1)
      expect(grouped.areas[0].items.map((i) => i.id)).toEqual(['1', '2'])
      expect(grouped.tail).toEqual([])
    } finally {
      editor.destroy()
    }
  })
})

describe('section-end endnote placement setting', () => {
  const base: NoteSettingsState = {
    settings: { endnote: { pos: 'sectEnd', numFmt: 'lowerRoman' } },
    sections: [{ lastBlockIndex: 1, endnote: { pos: 'docEnd' } }, { lastBlockIndex: 3 }],
  }
  it('uses only the document-level position (Word ignores sectPr-level w:pos)', () => {
    expect(endnotePlacement(base)).toBe('sectEnd')
    expect(effectiveNoteSettings(base, 'endnote').pos).toBe('sectEnd')
    expect(sectionNoteSettings(base, 'endnote', 0).pos).toBe('sectEnd')
    expect(endnotePlacement({ settings: {}, sections: [] })).toBe('docEnd')
    expect(endnotePlacement(null)).toBe('docEnd')
  })
  it('never writes the endnote position at section scope', () => {
    expect(applySectionNoteSettings(base, 'endnote', 1, { pos: 'docEnd' })).toBe(base)
    const next = applySectionNoteSettings(base, 'endnote', 1, { pos: 'docEnd', numFmt: 'decimal' })
    expect(next.sections[1].endnote).toEqual({ numFmt: 'decimal' })
  })
})

describe('section-end endnote area widgets', () => {
  it('renders one in-flow row per entry before the next section and tracks edits', () => {
    const editor = build([para(0, [ref('1')]), para(1), para(2), para(3), para(4)])
    try {
      const handlers = { onEdit: vi.fn(), onDelete: vi.fn(), onConvert: vi.fn() }
      const labels = { edit: 'Edit', delete: 'Delete', convert: 'Convert' }
      const pos = editor.state.doc.child(0).nodeSize + editor.state.doc.child(1).nodeSize
      const areas = [
        {
          key: '0',
          pos,
          rows: [
            { id: '1', marker: 'i', text: 'Alpha note', height: 40, withSeparator: true },
            { id: '2', marker: 'ii', text: 'Second', height: 24, withSeparator: false },
          ],
        },
      ]
      expect(setSectionEndnoteAreas(editor.view, areas, handlers, labels)).toBe(true)
      // identical input is a no-op (no transaction churn)
      expect(setSectionEndnoteAreas(editor.view, areas, handlers, labels)).toBe(false)
      const rows = editor.view.dom.querySelectorAll<HTMLElement>('.page-section-endnote-row')
      expect(rows).toHaveLength(2)
      expect(rows[0].style.height).toBe('40px')
      expect(rows[0].querySelector('.page-section-endnote-sep')).not.toBeNull()
      expect(rows[1].querySelector('.page-section-endnote-sep')).toBeNull()
      expect(rows[1].textContent).toContain('ii')
      rows[0].querySelectorAll<HTMLButtonElement>('button')[0].click()
      expect(handlers.onEdit).toHaveBeenCalledWith('1')
      // widgets sit before the first block of the next section, in document order
      const children = Array.from(editor.view.dom.children)
      const firstRow = children.findIndex((el) => el.classList.contains('page-section-endnote-row'))
      expect(children[firstRow - 1].textContent).toContain('p1')
      expect(children[firstRow + 2].textContent).toContain('p2')
      // edit above maps the anchor along; saved document content is untouched
      editor.commands.insertContentAt(1, 'zz')
      expect(sectionEndnoteAreaPos(editor.view, '0')).toBe(pos + 2)
      expect(editor.getJSON().content).toHaveLength(5)
      expect(JSON.stringify(editor.getJSON())).not.toContain('Alpha note')
    } finally {
      editor.destroy()
    }
  })
})
