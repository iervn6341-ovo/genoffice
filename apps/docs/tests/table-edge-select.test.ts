import { Editor } from '@tiptap/core'
import { CellSelection } from '@tiptap/pm/tables'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
const row = (...c: string[]) => `<w:tr>${c.join('')}</w:tr>`
const BODY =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  row(cell('a1'), cell('b1'), cell('c1')) +
  row(cell('a2'), cell('b2'), cell('c2')) +
  row(cell('a3'), cell('b3'), cell('c3')) +
  '</w:tbl><w:p/>'

let editor: Editor | null = null
afterEach(() => {
  window.dispatchEvent(new MouseEvent('mouseup'))
  editor?.destroy()
})
// jsdom has no hit testing; other mousedown plugins probe for glyphs under the pointer
;(document as unknown as { caretRangeFromPoint: () => null }).caretRangeFromPoint = () => null
;(document as unknown as { elementsFromPoint: () => [] }).elementsFromPoint = () => []
;(document as unknown as { elementFromPoint: () => null }).elementFromPoint = () => null

const rect = (left: number, top: number, width: number, height: number) =>
  ({
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
    x: left,
    y: top,
  }) as DOMRect

/** lay the table out on a 3x3 grid: columns at x=100,200,300 (100 wide), rows at y=50,80,110 (30 tall) */
async function open() {
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const table = editor.view.dom.querySelector('table.doc-table') as HTMLTableElement
  table.getBoundingClientRect = () => rect(100, 50, 300, 90)
  Array.from(table.rows).forEach((tr, r) => {
    tr.getBoundingClientRect = () => rect(100, 50 + r * 30, 300, 30)
    Array.from(tr.cells).forEach((td, c) => {
      td.getBoundingClientRect = () => rect(100 + c * 100, 50 + r * 30, 100, 30)
    })
  })
  return { editor, table }
}

const press = (ed: Editor, x: number, y: number) =>
  ed.view.dom.dispatchEvent(
    new MouseEvent('mousedown', {
      bubbles: true,
      cancelable: true,
      clientX: x,
      clientY: y,
      button: 0,
    }),
  )

const selectedTexts = (ed: Editor): string[] => {
  const sel = ed.state.selection as CellSelection
  const out: string[] = []
  sel.forEachCell((node) => out.push(node.textContent))
  return out
}

describe('table edge selectors', () => {
  it('a click left of a row selects the whole row', async () => {
    const { editor: ed } = await open()
    press(ed, 90, 95) // left of row 2, inside the 16px band
    expect(ed.state.selection).toBeInstanceOf(CellSelection)
    expect(selectedTexts(ed)).toEqual(['a2', 'b2', 'c2'])
  })

  it('a click above a column selects the whole column', async () => {
    const { editor: ed } = await open()
    press(ed, 250, 40)
    expect(ed.state.selection).toBeInstanceOf(CellSelection)
    expect(selectedTexts(ed)).toEqual(['b1', 'b2', 'b3'])
  })

  it('dragging down along the left edge extends to more rows', async () => {
    const { editor: ed } = await open()
    press(ed, 90, 65)
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 90, clientY: 125 }))
    window.dispatchEvent(new MouseEvent('mouseup'))
    expect(selectedTexts(ed)).toEqual(['a1', 'b1', 'c1', 'a2', 'b2', 'c2', 'a3', 'b3', 'c3'])
  })

  it('a click inside the table or far outside the band is left to ProseMirror', async () => {
    const { editor: ed } = await open()
    const before = ed.state.selection
    press(ed, 150, 95) // inside the table
    press(ed, 40, 95) // well outside the 16px band
    expect(ed.state.selection.eq(before)).toBe(true)
  })

  it('the selected row is a normal CellSelection: Insert Below adds one row per selected row', async () => {
    const { editor: ed } = await open()
    press(ed, 90, 65)
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 90, clientY: 95 }))
    window.dispatchEvent(new MouseEvent('mouseup'))
    const { addRowAfter } = await import('@tiptap/pm/tables')
    const { insertPerSelection } = await import('../src/renderer/editor/table-tools')
    insertPerSelection(addRowAfter, 'row')(ed.state, ed.view.dispatch)
    let rows = 0
    ed.state.doc.descendants((n) => {
      if (n.type.name === 'docTableRow') rows++
    })
    expect(rows).toBe(5)
  })
})

describe('right press inside a cell range', () => {
  it('is swallowed so the native caret cannot collapse the range', async () => {
    const { editor: ed } = await open()
    press(ed, 90, 65)
    window.dispatchEvent(new MouseEvent('mousemove', { clientX: 90, clientY: 95 }))
    window.dispatchEvent(new MouseEvent('mouseup'))
    const before = ed.state.selection
    // a position inside the selected cells (what posAtCoords would report)
    const inside = (before as CellSelection).ranges[0].$from.pos + 1
    ed.view.posAtCoords = () => ({ pos: inside, inside: -1 })
    const ev = new MouseEvent('mousedown', {
      bubbles: true,
      cancelable: true,
      button: 2,
      clientX: 150,
      clientY: 65,
    })
    ed.view.dom.dispatchEvent(ev)
    expect(ev.defaultPrevented).toBe(true)
    expect(ed.state.selection.eq(before)).toBe(true)
    // outside the range the press is left alone
    ed.view.posAtCoords = () => ({ pos: 1, inside: -1 })
    const ev2 = new MouseEvent('mousedown', {
      bubbles: true,
      cancelable: true,
      button: 2,
      clientX: 150,
      clientY: 125,
    })
    ed.view.dom.dispatchEvent(ev2)
    expect(ev2.defaultPrevented).toBe(false)
  })
})
