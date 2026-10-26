import { Editor } from '@tiptap/core'
import { Slice } from '@tiptap/pm/model'
import { TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { tableDragGuard } from '../src/renderer/editor/table-drag'

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const BODY =
  '<w:p><w:r><w:t>outside paragraph</w:t></w:r></w:p>' +
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('alpha one')}${cell('bravo')}</w:tr><w:tr>${cell('charlie')}${cell('delta')}</w:tr></w:tbl><w:p/>`

let editor: Editor | null = null
afterEach(() => {
  editor?.destroy()
  editor = null
})

async function open() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return editor
}

function cells(ed: Editor): number[] {
  const out: number[] = []
  ed.state.doc.descendants((n, pos) => {
    if (n.type.name === 'docTableCell' || n.type.name === 'docTableHeader') out.push(pos)
  })
  return out
}

/** A view stand-in whose posAtCoords always resolves to `dropPos` */
const viewAt = (ed: Editor, dropPos: number) => ({
  state: ed.state,
  posAtCoords: () => ({ pos: dropPos, inside: -1 }),
})
const ev = () =>
  ({ clientX: 1, clientY: 1, preventDefault: vi.fn() }) as unknown as DragEvent & {
    preventDefault: ReturnType<typeof vi.fn>
  }

describe('docs table drag guard', () => {
  it('cancels the native drag of a cell-range selection', async () => {
    const ed = await open()
    const [a, , , d] = cells(ed)
    ed.view.dispatch(ed.state.tr.setSelection(CellSelection.create(ed.state.doc, a, d)))
    const e = ev()
    expect(tableDragGuard.dragstart(viewAt(ed, 1), e)).toBe(true)
    expect(e.preventDefault).toHaveBeenCalled()
  })

  it('lets a plain text selection start its drag', async () => {
    const ed = await open()
    const [a] = cells(ed)
    ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, a + 2, a + 5)))
    const e = ev()
    expect(tableDragGuard.dragstart(viewAt(ed, 1), e)).toBe(false)
    expect(e.preventDefault).not.toHaveBeenCalled()
  })

  it('swallows a cross-cell text move and leaves the document untouched', async () => {
    const ed = await open()
    const [a, b] = cells(ed)
    ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, a + 2, a + 7)))
    const before = ed.state.doc.toJSON()
    const slice = ed.state.selection.content()
    const e = ev()
    expect(tableDragGuard.drop(viewAt(ed, b + 3), e, slice, true)).toBe(true)
    expect(e.preventDefault).toHaveBeenCalled()
    expect(ed.state.doc.toJSON()).toEqual(before)
  })

  it('still allows moving text within the same cell', async () => {
    const ed = await open()
    const [a] = cells(ed)
    ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, a + 2, a + 5)))
    const slice = ed.state.selection.content()
    expect(tableDragGuard.drop(viewAt(ed, a + 9), ev(), slice, true)).toBe(false)
  })

  it('swallows a row/cell slice dropped into a table, even when copied (not moved)', async () => {
    const ed = await open()
    const [a, , , d] = cells(ed)
    ed.view.dispatch(ed.state.tr.setSelection(CellSelection.create(ed.state.doc, a, d)))
    const slice = ed.state.selection.content()
    expect(slice.content.childCount).toBeGreaterThan(0)
    expect(tableDragGuard.drop(viewAt(ed, cells(ed)[2] + 3), ev(), slice, false)).toBe(true)
  })

  it('keeps ordinary paragraph drags (outside a table) and drops onto a table from a paragraph', async () => {
    const ed = await open()
    // text in the paragraph before the table, dropped back into that paragraph
    ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, 2, 6)))
    const slice = ed.state.selection.content()
    expect(tableDragGuard.drop(viewAt(ed, 10), ev(), slice, true)).toBe(false)
    // moving paragraph text into a cell is a deliberate Word move and stays allowed
    const [, b] = cells(ed)
    expect(tableDragGuard.drop(viewAt(ed, b + 3), ev(), slice, true)).toBe(false)
  })

  it('ignores a drop with no resolvable position', async () => {
    const ed = await open()
    const e = ev()
    const v = { state: ed.state, posAtCoords: () => null }
    expect(tableDragGuard.drop(v, e, Slice.empty, true)).toBe(false)
  })
})
