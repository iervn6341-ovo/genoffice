import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { pictureInsertPosOutsideTable } from '../src/renderer/components/ribbon-tabs'
import { insertImageFromDataUrl } from '../src/renderer/components/ribbon-tabs'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
const BODY =
  '<w:p><w:r><w:t>before</w:t></w:r></w:p>' +
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('b1')}</w:tr><w:tr>${cell('a2')}${cell('b2')}</w:tr></w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p>`
// 1x1 transparent PNG
const PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

let editor: Editor | null = null
afterEach(() => {
  editor?.destroy()
  editor = null
  vi.unstubAllGlobals()
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
const topNames = (ed: Editor) =>
  ed.state.doc.content.content.map(
    (n) => n.type.name + (n.attrs.blockType === 'image' ? ':image' : ''),
  )

function caretIn(ed: Editor, text: string) {
  let at = -1
  ed.state.doc.descendants((n, p) => {
    if (n.type.name === 'docParagraph' && n.textContent === text) at = p + 1
  })
  ed.view.dispatch(ed.state.tr.setSelection(TextSelection.create(ed.state.doc, at)))
}

describe('picture insert never splits a table', () => {
  it('reports the position right after the outermost table when the caret is in a cell', async () => {
    const ed = await open()
    caretIn(ed, 'a2')
    const pos = pictureInsertPosOutsideTable(ed.state.selection.$from)
    expect(pos).not.toBeNull()
    expect(ed.state.doc.resolve(pos!).nodeBefore?.type.name).toBe('docTable')
    caretIn(ed, 'before')
    expect(pictureInsertPosOutsideTable(ed.state.selection.$from)).toBeNull()
  })

  it('keeps ONE table and puts the picture after it', async () => {
    const ed = await open()
    // jsdom never decodes images: report a natural size
    vi.stubGlobal(
      'Image',
      class {
        naturalWidth = 40
        naturalHeight = 20
        onload: (() => void) | null = null
        set src(_v: string) {
          queueMicrotask(() => this.onload?.())
        }
      },
    )
    caretIn(ed, 'a2')
    expect(await insertImageFromDataUrl(ed, PNG)).toBe(true)
    const names = topNames(ed)
    expect(names.filter((n) => n === 'docTable')).toHaveLength(1)
    const t = names.indexOf('docTable')
    expect(names[t + 1]).toBe('docProtected:image')
    // the caret ends up in a text position after the picture, not in the table
    expect(ed.state.selection.$from.parent.isTextblock).toBe(true)
  })
})
