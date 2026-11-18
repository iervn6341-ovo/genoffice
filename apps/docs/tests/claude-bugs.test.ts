import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx, TINY_PNG_BASE64 } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

const INLINE_DRAWING_XML =
  '<w:drawing>' +
  '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="914400" cy="914400"/>' +
  '<wp:docPr id="1" name="Picture 1"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:nvPicPr><pic:cNvPr id="1" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill>' +
  '<a:blip r:embed="rId1"/>' +
  '<a:stretch><a:fillRect/></a:stretch>' +
  '</pic:blipFill>' +
  '<pic:spPr>' +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '</pic:spPr>' +
  '</pic:pic>' +
  '</a:graphicData>' +
  '</a:graphic>' +
  '</wp:inline>' +
  '</w:drawing>'

describe('Verification of B2 and B3 in Electron Docs', () => {
  it('B2: typing text at the start of a paragraph with an inline image preserves the image on save', async () => {
    const source = await buildDocx({
      bodyXml: `<w:p><w:r>${INLINE_DRAWING_XML}</w:r></w:p>`,
      withImage: true,
      binaryParts: [
        {
          path: 'word/media/image1.png',
          base64: TINY_PNG_BASE64,
          extension: 'png',
          contentType: 'image/png',
        },
      ],
      extraRels:
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>',
    })

    const parsed = await parseDocx(source)
    const host = document.createElement('div')
    document.body.appendChild(host)

    const editor = new Editor({
      element: host,
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })

    // Insert text at position 1 (start of paragraph before image)
    editor.commands.insertContentAt(1, 'Hello before image: ')

    // Save
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const savedBytes = await saveDocx(parsed, plan.saveBlocks)

    // Reopen and verify
    const reparsed = await parseDocx(savedBytes)
    const imageBlock = reparsed.blocks.find((b) => b.type === 'image')
    const textBlock = reparsed.blocks.find((b) => b.type === 'paragraph')

    // Image is preserved!
    expect(imageBlock).toBeDefined()
    // Text is preserved!
    expect(textBlock).toBeDefined()
    expect(textBlock?.runs?.[0]?.text).toContain('Hello before image: ')

    editor.destroy()
    host.remove()
  })

  it('B3: save -> undo -> save preserves the undone state in the saved file', async () => {
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Original Text</w:t></w:r></w:p>',
    })

    let parsed = await parseDocx(source)
    const host = document.createElement('div')
    document.body.appendChild(host)

    const editor = new Editor({
      element: host,
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })

    // Modify text: select all, replace with "Modified Text"
    editor.chain().selectAll().insertContent('Modified Text').run()
    expect(editor.getText()).toBe('Modified Text')

    // First save
    let plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    let savedBytes = await saveDocx(parsed, plan.saveBlocks)

    // Verify first save has "Modified Text"
    let reparsed = await parseDocx(savedBytes)
    expect(reparsed.blocks[0].runs![0].text).toBe('Modified Text')

    // In the real application, after save, parsed is updated to reparsed (or rebased)
    parsed = reparsed

    // Undo in editor
    editor.commands.undo()
    expect(editor.getText()).toBe('Original Text')

    // Second save after undo
    plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    savedBytes = await saveDocx(parsed, plan.saveBlocks)

    // Reopen and verify it contains "Original Text", NOT "Modified Text"
    reparsed = await parseDocx(savedBytes)
    expect(reparsed.blocks[0].runs![0].text).toBe('Original Text')

    editor.destroy()
    host.remove()
  })
})
