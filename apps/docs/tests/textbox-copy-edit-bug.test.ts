import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { NodeSelection } from '@tiptap/pm/state'
import { buildTextboxParagraphXml, parseDocx, type TextboxDisplay } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { insertTextboxAt } from '../src/renderer/components/ribbon-tabs'
import { createClipHost } from '../src/renderer/editor/clipboard-host'

const WIDTH_EMU = 1800000
const HEIGHT_EMU = 1080000

function makeEmptyTextbox(): TextboxDisplay {
  return {
    fill: 'FFFFFF',
    borderColor: '000000',
    widthPx: Math.round(WIDTH_EMU / 9525),
    heightPx: Math.round(HEIGHT_EMU / 9525),
    paras: [{ runs: [{ text: '' }] }],
  }
}

async function openBlankDoc() {
  const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>' })
  const parsed = await parseDocx(source)
  const container = document.createElement('div')
  document.body.appendChild(container)
  const editor = new Editor({
    element: container,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const store = editor.storage.clipboardTransfer
  if (store) {
    store.host = createClipHost({
      parsed: () => parsed,
      addStyleUpsert: () => {},
      styleSheetChanged: () => {},
      createListDef: () => '1',
      addNote: () => null,
      normalizeNotes: (tr) => tr,
      numbering: () => new Map(),
    })
  }
  return { editor, parsed, container }
}

describe('textbox copy and edit bugs', () => {
  it('reproduces copy-paste of a textbox', async () => {
    const { editor, container } = await openBlankDoc()
    insertTextboxAt(editor)

    // Find docProtected node
    let pos = -1
    editor.state.doc.descendants((node, p) => {
      if (pos !== -1) return false
      if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = p
      return pos === -1
    })
    expect(pos).toBeGreaterThanOrEqual(0)

    // Select the node
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    expect(editor.state.selection instanceof NodeSelection).toBe(true)

    // Check if the DOM has ProseMirror-selectednode
    const selectedEl = container.querySelector('.ProseMirror-selectednode')
    expect(selectedEl).not.toBeNull()

    // Simulate Command+C (copy)
    const clipboardDataStore: Record<string, string> = {}
    const clipboardData = {
      setData: (type: string, val: string) => {
        clipboardDataStore[type] = val
      },
      getData: (type: string) => clipboardDataStore[type] || '',
      clearData: () => {
        Object.keys(clipboardDataStore).forEach((k) => delete clipboardDataStore[k])
      },
    }
    const copyEvent = new Event('copy', { bubbles: true, cancelable: true }) as any
    copyEvent.clipboardData = clipboardData

    // Check handleDOMEvents copy
    const handled = editor.view.someProp('handleDOMEvents', (handlers) => {
      return handlers.copy?.(editor.view, copyEvent)
    })
    console.log('Handled copy:', handled)
    console.log('Clipboard store after copy:', clipboardDataStore)

    // Now test PASTE!
    const pasteEvent = new Event('paste', { bubbles: true, cancelable: true }) as any
    pasteEvent.clipboardData = clipboardData
    pasteEvent.clipboardData.items = []

    // Keep NodeSelection on the textbox to test pasting directly onto the selected object
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, pos)))
    expect(editor.state.selection instanceof NodeSelection).toBe(true)
    const initialChildCount = editor.state.doc.childCount

    // Test parseClipboardDom
    const host = (editor.storage.clipboardTransfer as any).host
    const dom = new DOMParser().parseFromString(clipboardDataStore['text/html'], 'text/html').body

    const slice = editor.view.someProp('clipboardParser', (parser) => {
      return (parser as any).parseSlice(dom, {})
    })
    expect(slice).not.toBeNull()

    // Test handlePaste with NodeSelection active
    const pasteHandled = editor.view.someProp('handlePaste', (hp) => {
      return hp(editor.view, pasteEvent, slice!)
    })
    expect(pasteHandled).toBe(true)
    expect(editor.state.doc.childCount).toBe(initialChildCount + 1)
    expect(editor.state.selection instanceof NodeSelection).toBe(true)

    editor.destroy()
    container.remove()
  })

  it('reproduces text reverting after editing in a textbox', async () => {
    const { editor, container } = await openBlankDoc()
    insertTextboxAt(editor)

    let pos = -1
    editor.state.doc.descendants((node, p) => {
      if (pos !== -1) return false
      if (node.type.name === 'docProtected' && node.attrs.textboxes) pos = p
      return pos === -1
    })

    // Find the sub-editor DOM element
    const boxEl = container.querySelector('.doc-textbox') as HTMLElement
    expect(boxEl).not.toBeNull()

    // Double click to enter edit mode
    boxEl.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }))

    // The sub-editor inside
    const subEditorEl = boxEl.querySelector('.doc-textbox-editor') as HTMLElement
    expect(subEditorEl).not.toBeNull()

    // Let's find the TipTap sub-editor instance from active editor or editors array
    // Simulate user typing in sub-editor
    // In TipTap/ProseMirror, typing in the sub-editor dispatches a transaction to the sub-editor
    // Let's find how the sub-editor is accessed
    const subEditorField = boxEl.querySelector('[contenteditable="true"]') as HTMLElement
    expect(subEditorField).not.toBeNull()

    // Let's inspect active sub-editor
    const { getActiveSubEditor } = await import('../src/renderer/editor/active-editor')
    const sub = getActiveSubEditor()
    console.log('Active sub editor found?', !!sub)
    if (sub) {
      sub.commands.setContent({
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            content: [{ type: 'text', text: 'Hello World New Text' }],
          },
        ],
      })
      console.log('Sub editor text after edit:', sub.getText())

      // Now simulate clicking outside or blurring
      // 1. focusout event on dom
      const protectedDom = container.querySelector('.doc-protected') as HTMLElement
      protectedDom.dispatchEvent(
        new FocusEvent('focusout', { bubbles: true, relatedTarget: document.body }),
      )

      // 2. document mousedown outside
      document.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }))

      // Let's check node in main editor state
      const updatedNode = editor.state.doc.nodeAt(pos)!
      const boxes = updatedNode.attrs.textboxes as TextboxDisplay[]
      console.log('Main doc node textboxes paras:', JSON.stringify(boxes[0].paras))

      // Check sub editor DOM text
      console.log('Sub editor DOM text after blur:', subEditorEl.textContent)

      // What if the main editor has a transaction (e.g. selection change or typing elsewhere)?
      editor.commands.focus('start')
      console.log('Sub editor DOM text after main editor transaction:', subEditorEl.textContent)
      console.log('Sub editor getText():', sub.getText())
    }

    editor.destroy()
    container.remove()
  })
})
