import { afterEach, describe, expect, it, vi } from 'vitest'
import type { EditorView } from '@tiptap/pm/view'
import { wireProtectedInteractionMode } from '../src/renderer/editor/extensions'

afterEach(() => {
  document.body.innerHTML = ''
})

function setup(singleClickEdits = false) {
  const dom = document.createElement('div')
  dom.innerHTML = '<span class="doc-move-handle"></span><div class="doc-textbox">hello</div>'
  const ribbon = document.createElement('div')
  ribbon.setAttribute('data-keep-edit', '')
  ribbon.innerHTML = '<button id="bold">B</button><input id="size" />'
  const page = document.createElement('p')
  page.id = 'page'
  document.body.append(dom, ribbon, page)
  const editor = { setEditable: vi.fn(), commit: vi.fn() }
  const cleanup = wireProtectedInteractionMode(
    dom,
    () => undefined,
    {} as unknown as EditorView,
    editor,
    () => singleClickEdits,
  )
  const press = (el: Element) =>
    el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))
  const startEditing = () =>
    dom
      .querySelector('.doc-textbox')!
      .dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }))
  return { dom, editor, cleanup, press, startEditing }
}

describe('text box editing survives ribbon presses (Word keeps the edit and selection)', () => {
  it('a ribbon button or field press keeps the box in edit mode and does not commit', () => {
    const { dom, editor, press, startEditing, cleanup } = setup()
    startEditing()
    expect(dom.classList.contains('doc-content-editing')).toBe(true)
    editor.commit.mockClear()
    press(document.getElementById('bold')!)
    press(document.getElementById('size')!)
    expect(dom.classList.contains('doc-content-editing')).toBe(true)
    expect(editor.commit).not.toHaveBeenCalled()
    cleanup?.()
  })

  it('a press on the page outside every zone still ends the edit and commits', () => {
    const { dom, editor, press, startEditing, cleanup } = setup()
    startEditing()
    editor.commit.mockClear()
    press(document.getElementById('page')!)
    expect(dom.classList.contains('doc-content-editing')).toBe(false)
    expect(editor.commit).toHaveBeenCalledTimes(1)
    cleanup?.()
  })
})

describe('single click in a plain text box', () => {
  it('enters edit mode on the first press when the box allows it', () => {
    const { dom, press, cleanup } = setup(true)
    press(dom.querySelector('.doc-textbox')!)
    expect(dom.classList.contains('doc-content-editing')).toBe(true)
    cleanup?.()
  })
  it('only selects the object for shapes that keep press-to-move', () => {
    const { dom, press, cleanup } = setup(false)
    press(dom.querySelector('.doc-textbox')!)
    expect(dom.classList.contains('doc-content-editing')).toBe(false)
    cleanup?.()
  })
})

describe('typing at a selected text box', () => {
  it('textboxTypeThroughPlugin enters the box instead of replacing the object', async () => {
    const { Editor } = await import('@tiptap/core')
    const { NodeSelection } = await import('@tiptap/pm/state')
    const { editorExtensions, textboxTypeThroughPlugin } =
      await import('../src/renderer/editor/extensions')
    const { blocksToPmDoc } = await import('../src/renderer/editor/convert')
    const { parseDocx } = await import('@genoffice/docx-engine')
    const { buildDocx } = await import('../../../packages/docx-engine/tests/helpers/build-docx')
    const txbx = '<w:txbxContent><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:txbxContent>'
    const body =
      '<w:p><w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:extent cx="1800000" cy="1080000"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="1080000"/></a:xfrm></wps:spPr>' +
      `<wps:txbx>${txbx}</wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
      '<mc:Fallback><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" id="b1" style="position:absolute">' +
      `<v:textbox>${txbx}</v:textbox></v:rect></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>`
    const parsed = await parseDocx(await buildDocx({ bodyXml: body }))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
    })
    editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    let pos = -1
    editor.state.doc.descendants((n, p) => {
      if (pos === -1 && n.type.name === 'docProtected' && n.attrs.textboxes) pos = p
      return pos === -1
    })
    expect(pos).toBeGreaterThanOrEqual(0)
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    const exec = vi.fn(() => true)
    ;(document as unknown as { execCommand: unknown }).execCommand = exec
    const plugin = textboxTypeThroughPlugin()
    const handled = plugin.props.handleTextInput!.call(
      plugin,
      editor.view,
      pos,
      pos + 1,
      'Z',
      () => editor.state.tr,
    )
    expect(handled).toBe(true)
    expect(exec).toHaveBeenCalledWith('insertText', false, 'Z')
    // the object is still there and now in edit mode
    expect(editor.state.doc.nodeAt(pos)?.type.name).toBe('docProtected')
    const dom = editor.view.nodeDOM(pos) as HTMLElement
    expect(dom.classList.contains('doc-content-editing')).toBe(true)
    editor.destroy()
  })
})
