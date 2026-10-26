import { Editor } from '@tiptap/core'
import { parseDocx } from '@genoffice/docx-engine'
import { afterEach, describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const cell = (t: string) => `<w:tc><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
const BODY =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell('a')}</w:tr></w:tbl><w:p/>`

let editor: Editor | null = null
afterEach(() => editor?.destroy())

const rect = (left: number, top: number, width = 100, height = 50) =>
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

describe('table move handle placement', () => {
  it('measures against the CURRENT page container, not the detached one it was created in', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: BODY }))
    const detached = document.createElement('div')
    editor = new Editor({
      element: detached,
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    // what EditorContent does on mount: move everything into the real page container
    const page = document.createElement('div')
    document.body.append(page)
    page.append(...Array.from(detached.childNodes))
    expect(editor.view.dom.parentElement).toBe(page)

    const table = editor.view.dom.querySelector('table.doc-table') as HTMLElement
    expect(table).not.toBeNull()
    // the page container sits at (415,116); the table at (511,241) in the viewport
    page.getBoundingClientRect = () => rect(415, 116, 800, 900)
    detached.getBoundingClientRect = () => rect(0, 0, 0, 0)
    table.getBoundingClientRect = () => rect(511, 241)
    table.dispatchEvent(new MouseEvent('mousemove', { bubbles: true }))

    const handle = editor.view.dom.parentElement!.querySelector('.doc-table-handle') as HTMLElement
    expect(handle.style.display).toBe('flex')
    expect(handle.parentElement).toBe(page)
    // table top-left relative to the page, minus the 18px outside offset
    expect(handle.style.left).toBe(`${511 - 415 - 18}px`)
    expect(handle.style.top).toBe(`${241 - 116 - 18}px`)
  })
})
