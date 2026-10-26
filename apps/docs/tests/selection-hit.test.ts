import { Editor } from '@tiptap/core'
import { CellSelection } from '@tiptap/pm/tables'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { posInsideSelection } from '../src/renderer/editor/selection-hit'

const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
const BODY =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('b1')}</w:tr><w:tr>${cell('a2')}${cell('b2')}</w:tr></w:tbl><w:p/>`

let editor: Editor | null = null
afterEach(() => editor?.destroy())

describe('posInsideSelection', () => {
  it('counts every cell of a CellSelection as inside, and nothing outside it', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    const start: Record<string, number> = {}
    editor.state.doc.descendants((n, p) => {
      if (n.type.name === 'docParagraph' && n.textContent) start[n.textContent] = p + 1
    })
    // select a1..b1 (first row) with b1 as head: selection.from/to cover only b1
    editor.view.dispatch(
      editor.state.tr.setSelection(
        CellSelection.create(editor.state.doc, start.a1 - 2, start.b1 - 2),
      ),
    )
    const sel = editor.state.selection
    expect(posInsideSelection(sel, start.b1 + 1)).toBe(true)
    expect(posInsideSelection(sel, start.a1 + 1)).toBe(true) // outside from/to, inside a range
    expect(sel.from <= start.a1 + 1 && start.a1 + 1 <= sel.to).toBe(false)
    expect(posInsideSelection(sel, start.a2 + 1)).toBe(false)
    expect(posInsideSelection(sel, start.b2 + 1)).toBe(false)
  })
})
