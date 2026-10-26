import { afterAll, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { tableDragGuard } from '../src/renderer/editor/tableDragGuard'

const editors: Editor[] = []
afterAll(() => editors.forEach((e) => e.destroy()))

function open(): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: 'intro text\n\n| A | B |\n| --- | --- |\n| alpha one | bravo |\n| charlie | delta |\n',
    contentType: 'markdown',
  })
  editors.push(editor)
  return editor
}

function cells(ed: Editor): number[] {
  const out: number[] = []
  ed.state.doc.descendants((n, pos) => {
    if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') out.push(pos)
  })
  return out
}
const viewAt = (ed: Editor, pos: number) => ({
  state: ed.state,
  posAtCoords: () => ({ pos, inside: -1 }),
})
const ev = () => ({ clientX: 1, clientY: 1, preventDefault: vi.fn() }) as unknown as DragEvent

describe('markdown table drag guard', () => {
  it('is installed as an editor plugin', () => {
    const ed = open()
    const names = ed.extensionManager.extensions.map((x) => x.name)
    expect(names).toContain('tableDragGuard')
    expect(cells(ed).length).toBe(6)
  })

  it('cancels the drag of a cell-range selection', () => {
    const ed = open()
    const c = cells(ed)
    ed.view.dispatch(ed.state.tr.setSelection(CellSelection.create(ed.state.doc, c[2], c[5])))
    const e = ev()
    expect(tableDragGuard.dragstart(viewAt(ed, 1), e)).toBe(true)
    expect(e.preventDefault).toHaveBeenCalled()
  })

  it('swallows a cross-cell move but allows a same-cell move', () => {
    const ed = open()
    const c = cells(ed)
    ed.view.dispatch(
      ed.state.tr.setSelection(TextSelection.create(ed.state.doc, c[2] + 2, c[2] + 6)),
    )
    const slice = ed.state.selection.content()
    expect(tableDragGuard.drop(viewAt(ed, c[3] + 3), ev(), slice, true)).toBe(true)
    expect(tableDragGuard.drop(viewAt(ed, c[2] + 8), ev(), slice, true)).toBe(false)
  })

  it('swallows a row/cell slice dropped into the table', () => {
    const ed = open()
    const c = cells(ed)
    ed.view.dispatch(ed.state.tr.setSelection(CellSelection.create(ed.state.doc, c[2], c[3])))
    const slice = ed.state.selection.content()
    expect(tableDragGuard.drop(viewAt(ed, c[5] + 3), ev(), slice, false)).toBe(true)
  })

  it('leaves paragraph drags alone', () => {
    const ed = open()
    ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, 2, 6)))
    const slice = ed.state.selection.content()
    expect(tableDragGuard.drop(viewAt(ed, 8), ev(), slice, true)).toBe(false)
  })
})
