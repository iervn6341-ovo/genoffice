import { Editor } from '@tiptap/core'
import { CellSelection } from '@tiptap/pm/tables'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { runUiOps } from '../src/renderer/ai/ops'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { setParagraphDirection, setSelectionAlign } from '../src/renderer/editor/direction'
import { editorExtensions } from '../src/renderer/editor/extensions'

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`
const row = (...c: string[]) => `<w:tr>${c.join('')}</w:tr>`
const BODY =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  row(cell('a1'), cell('b1'), cell('c1')) +
  row(cell('a2'), cell('b2'), cell('c2')) +
  row(cell('a3'), cell('b3'), cell('c3')) +
  '</w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p>'

let editor: Editor | null = null
afterEach(() => {
  editor?.destroy()
  editor = null
})

async function open(): Promise<Editor> {
  const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
  editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return editor
}

/** paragraph attrs by cell text, in document order */
function paras(ed: Editor, key: string): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  ed.state.doc.descendants((n) => {
    if (n.type.name === 'docParagraph' && n.textContent) out[n.textContent] = n.attrs[key]
  })
  return out
}

function selectRange(ed: Editor, anchorText: string, headText: string) {
  const pos: Record<string, number> = {}
  ed.state.doc.descendants((n, p) => {
    if (n.type.name === 'docParagraph' && n.textContent) pos[n.textContent] = p
  })
  // each paragraph is the first child of its cell, so the cell opens one position earlier
  ed.view.dispatch(
    ed.state.tr.setSelection(
      CellSelection.create(ed.state.doc, pos[anchorText] - 1, pos[headText] - 1),
    ),
  )
}

describe('paragraph formatting reaches every cell of a CellSelection', () => {
  it('Align via the ribbon helper centers all four cells of a 2x2 range', async () => {
    const ed = await open()
    selectRange(ed, 'a1', 'b2')
    expect(ed.state.selection).toBeInstanceOf(CellSelection)
    expect(setSelectionAlign(ed, 'center')).toBe(true)
    const a = paras(ed, 'align')
    expect([a.a1, a.b1, a.a2, a.b2]).toEqual(['center', 'center', 'center', 'center'])
    expect(a.c1).not.toBe('center')
    expect(a.a3).not.toBe('center')
  })

  it('works regardless of drag direction (head cell first in ranges)', async () => {
    const ed = await open()
    selectRange(ed, 'b2', 'a1')
    setSelectionAlign(ed, 'right')
    const a = paras(ed, 'align')
    expect([a.a1, a.b1, a.a2, a.b2]).toEqual(['right', 'right', 'right', 'right'])
  })

  it('line spacing (setParagraphAttrs op) covers the whole range, in one undo step', async () => {
    const ed = await open()
    selectRange(ed, 'a1', 'b2')
    runUiOps(ed, [
      { op: 'setParagraphAttrs', target: { scope: 'selection' }, attrs: { lineSpacing: 2 } },
    ])
    const ls = paras(ed, 'lineSpacing')
    expect([ls.a1, ls.b1, ls.a2, ls.b2]).toEqual([2, 2, 2, 2])
    expect(ls.c1).not.toBe(2)
    expect(ls.a3).not.toBe(2)
    ed.commands.undo()
    expect(paras(ed, 'lineSpacing').a1).not.toBe(2)
    expect(paras(ed, 'lineSpacing').b2).not.toBe(2)
  })

  it('direction change also reaches the range', async () => {
    const ed = await open()
    selectRange(ed, 'a1', 'b2')
    setParagraphDirection(ed, 'rtl')
    const b = paras(ed, 'bidi')
    expect([b.a1, b.b1, b.a2, b.b2]).toEqual([true, true, true, true])
  })

  it('a caret in one cell still formats only that cell', async () => {
    const ed = await open()
    let target = 0
    ed.state.doc.descendants((n, p) => {
      if (n.type.name === 'docParagraph' && n.textContent === 'c2') target = p + 1
    })
    ed.commands.setTextSelection(target)
    setSelectionAlign(ed, 'center')
    const a = paras(ed, 'align')
    expect(a.c2).toBe('center')
    expect(a.b2).not.toBe('center')
  })
})
