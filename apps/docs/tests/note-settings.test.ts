import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, saveDocx, type SectionInfo } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  applySectionNoteSettings,
  applyWholeDocumentNoteSettings,
  effectiveNoteSettings,
  noteSettingsFromParsed,
  noteSettingsSaveEdits,
  noteReferencesInEditor,
  noteSectionAtPosition,
  numberFootnotesByPage,
} from '../src/renderer/editor/note-settings'
import {
  blocksToPmDoc,
  pmDocOptions,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { sameBodyAfterManagedRPrRebase } from '../src/renderer/editor/save-rebase'

describe('whole-document note settings history', () => {
  it('keeps Undo history when reopening only changes a per-page reference display number', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:r><w:t>first</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>' +
          '<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>second</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>',
        sectPrExtra:
          '<w:footnotePr><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr>',
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
        extraParts: [
          {
            path: 'word/footnotes.xml',
            xml: '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:r><w:t>one</w:t></w:r></w:p></w:footnote><w:footnote w:id="3"><w:p><w:r><w:t>two</w:t></w:r></w:p></w:footnote></w:footnotes>',
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
          },
        ],
      }),
    )
    const live = blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed))
    const refs = live
      .content!.flatMap((block) => block.content ?? [])
      .filter((node) => node.type === 'docNoteRef')
    refs[1]!.attrs = { ...refs[1]!.attrs, num: 3 }
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: live as never,
    })
    try {
      const bytes = await saveDocx(
        parsed,
        parsed.blocks
          .filter((block) => !block.hidden && block.docxIndex !== null)
          .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! })),
      )
      const reopened = await parseDocx(bytes)
      const reopenedPm = blocksToPmDoc(
        reopened.blocks,
        readSections(reopened),
        pmDocOptions(reopened),
      )
      expect(
        reopenedPm
          .content!.flatMap((block) => block.content ?? [])
          .filter((node) => node.type === 'docNoteRef')
          .map((node) => node.attrs?.num),
      ).toEqual([3, 4])
      expect(
        sameBodyAfterManagedRPrRebase(
          editor.state.doc,
          editor.schema.nodeFromJSON({ ...reopenedPm, attrs: editor.state.doc.attrs }),
        ),
      ).toBe(true)
      const changedId = structuredClone(reopenedPm)
      changedId.content![1]!.content!.find((node) => node.type === 'docNoteRef')!.attrs!.id = '99'
      expect(
        sameBodyAfterManagedRPrRebase(
          editor.state.doc,
          editor.schema.nodeFromJSON({ ...changedId, attrs: editor.state.doc.attrs }),
        ),
      ).toBe(false)
    } finally {
      editor.destroy()
    }
  })

  it('patches only the caret section and numbers each measured page from its own start', () => {
    const state = {
      settings: { footnote: { numFmt: 'decimal', numStart: 1 } },
      sections: [
        { lastBlockIndex: 0, footnote: { numFmt: 'upperRoman', numStart: 3 } },
        { lastBlockIndex: 1, footnote: { numFmt: 'lowerLetter', numStart: 8 } },
      ],
    }
    const scoped = applySectionNoteSettings(state, 'footnote', 1, {
      numFmt: 'lowerRoman',
      numRestart: 'eachPage',
    })
    expect(scoped.settings).toEqual(state.settings)
    expect(scoped.sections[0]).toEqual(state.sections[0])
    expect(noteSettingsSaveEdits(scoped, state).sectionNoteProps).toEqual([
      {
        lastBlockIndex: 1,
        footnote: { numFmt: 'lowerRoman', numRestart: 'eachPage' },
      },
    ])
    const pages = [
      [
        { id: '2', no: 3 },
        { id: '3', no: 4 },
      ],
      [{ id: '4', no: 5 }],
      [
        { id: '5', no: 8 },
        { id: '6', no: 9 },
      ],
      [{ id: '7', no: 10 }],
    ]
    const both = applySectionNoteSettings(scoped, 'footnote', 0, { numRestart: 'eachPage' })
    expect(
      numberFootnotesByPage(pages, both, {
        '2': 0,
        '3': 0,
        '4': 0,
        '5': 1,
        '6': 1,
        '7': 1,
      }).map((page) => page.map((item) => item.no)),
    ).toEqual([[3, 4], [3], [8, 9], [8]])
  })

  it('keeps section anchors when preview numbers or section-break text change across saves', async () => {
    const original = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachSect"/><w:footnote w:id="0"/></w:footnotePr></w:sectPr></w:pPr><w:r><w:t>FIRST</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>' +
          '<w:p><w:r><w:t>SECOND</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>',
        sectPrExtra:
          '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="8"/><w:numRestart w:val="eachSect"/><w:footnote w:id="1"/></w:footnotePr>',
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
        extraParts: [
          {
            path: 'word/footnotes.xml',
            xml: '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:r><w:t>one</w:t></w:r></w:p></w:footnote><w:footnote w:id="3"><w:p><w:r><w:t>two</w:t></w:r></w:p></w:footnote></w:footnotes>',
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
          },
        ],
      }),
    )
    const before = noteSettingsFromParsed(original)
    const changed = applyWholeDocumentNoteSettings(before, 'footnote', {
      numFmt: 'decimal',
      numStart: 4,
    })
    const preview = blocksToPmDoc(original.blocks, readSections(original), pmDocOptions(original))
    for (const block of preview.content ?? []) {
      for (const node of block.content ?? []) {
        if (node.type === 'docNoteRef') node.attrs = { ...node.attrs, num: 4, fmt: 'decimal' }
      }
    }
    const previewPlan = pmDocToSavePlan(preview, original.blocks)
    expect(previewPlan.saveBlocks.slice(0, 2)).toEqual([
      { kind: 'original', docxIndex: 0 },
      { kind: 'original', docxIndex: 1 },
    ])

    const save = async (source: typeof original, doc: PmNode, desired: typeof before) =>
      parseDocx(
        await saveDocx(source, pmDocToSavePlan(doc, source.blocks).saveBlocks, {
          ...noteSettingsSaveEdits(desired, noteSettingsFromParsed(source)),
        }),
      )
    const first = await save(original, preview, changed)
    expect(readSections(first).map((section) => section.settings.footnotePr?.numFmt)).toEqual([
      'decimal',
      'decimal',
    ])

    const textEdited = structuredClone(preview)
    textEdited.content![0]!.content!.find((node) => node.type === 'text')!.text = 'FIRST EDIT'
    const textPlan = pmDocToSavePlan(textEdited, original.blocks)
    expect(textPlan.saveBlocks[0]).toMatchObject({ kind: 'generated', docxIndex: 0 })
    expect(
      textPlan.saveBlocks.filter((block) => 'docxIndex' in block && block.docxIndex === 0),
    ).toHaveLength(1)
    const editedBytes = await saveDocx(original, textPlan.saveBlocks, {
      ...noteSettingsSaveEdits(changed, before),
    })
    const edited = await parseDocx(editedBytes)
    expect(readSections(edited).map((section) => section.settings.footnotePr?.numFmt)).toEqual([
      'decimal',
      'decimal',
    ])
    const editedZip = await JSZip.loadAsync(editedBytes)
    const editedXml = await editedZip.file('word/document.xml')!.async('string')
    expect(editedXml).toContain('FIRST EDIT')
    expect(editedXml).toContain('<w:footnoteReference w:id="2"/>')
    expect(editedXml.match(/<w:numFmt w:val="decimal"\/>/g)).toHaveLength(2)

    const restored = blocksToPmDoc(original.blocks, readSections(original), pmDocOptions(original))
    const undone = await save(edited, restored, before)
    expect(readSections(undone).map((section) => section.settings.footnotePr?.numFmt)).toEqual([
      'upperRoman',
      'lowerLetter',
    ])
    const redone = await save(undone, textEdited, changed)
    expect(readSections(redone).map((section) => section.settings.footnotePr?.numFmt)).toEqual([
      'decimal',
      'decimal',
    ])
  })

  it('renumbers live editable references by section from the chosen starting number', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [{ type: 'docNoteRef', attrs: { kind: 'footnote', id: '2', num: 1 } }],
          },
          {
            type: 'docParagraph',
            attrs: { docxIndex: 1 },
            content: [{ type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 2 } }],
          },
        ],
      } as never,
    })
    try {
      const sections = [{ lastBlockIndex: 0 }, { lastBlockIndex: 1 }] as SectionInfo[]
      const state = (numRestart: 'eachSect' | 'continuous') => ({
        settings: { footnote: { numStart: 4, numRestart } },
        sections: [{ lastBlockIndex: 0 }, { lastBlockIndex: 1 }],
      })
      expect(
        noteReferencesInEditor(editor.state.doc, sections, 'footnote', state('eachSect')),
      ).toEqual({
        '2': { no: 4, fmt: 'decimal' },
        '3': { no: 4, fmt: 'decimal' },
      })
      expect(
        noteReferencesInEditor(editor.state.doc, sections, 'footnote', state('continuous')),
      ).toEqual({
        '2': { no: 4, fmt: 'decimal' },
        '3': { no: 5, fmt: 'decimal' },
      })
    } finally {
      editor.destroy()
    }
  })

  it('uses the following section for a new paragraph inserted after its section break', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [{ type: 'docNoteRef', attrs: { kind: 'footnote', id: '2', num: 3 } }],
          },
          {
            type: 'docParagraph',
            content: [{ type: 'docNoteRef', attrs: { kind: 'footnote', id: '4', num: 8 } }],
          },
          {
            type: 'docParagraph',
            attrs: { docxIndex: 1 },
            content: [{ type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 9 } }],
          },
        ],
      } as never,
    })
    try {
      expect(
        noteSectionAtPosition(editor.state.doc, editor.state.doc.child(0).nodeSize + 1, [
          { lastBlockIndex: 0 },
          { lastBlockIndex: 1 },
        ] as SectionInfo[]),
      ).toBe(1)
      expect(
        noteReferencesInEditor(
          editor.state.doc,
          [{ lastBlockIndex: 0 }, { lastBlockIndex: 1 }] as SectionInfo[],
          'footnote',
          {
            settings: {},
            sections: [
              { lastBlockIndex: 0, footnote: { numFmt: 'upperRoman', numStart: 3 } },
              {
                lastBlockIndex: 1,
                footnote: { numFmt: 'lowerLetter', numStart: 8, numRestart: 'eachSect' },
              },
            ],
          },
        ),
      ).toEqual({
        '2': { no: 3, fmt: 'upperRoman' },
        '4': { no: 8, fmt: 'lowerLetter' },
        '3': { no: 9, fmt: 'lowerLetter' },
      })
    } finally {
      editor.destroy()
    }
  })

  it('uses each section format/start for both the editable and imported reference markers', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachSect"/></w:footnotePr></w:sectPr></w:pPr><w:r><w:t>FIRST</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>' +
          '<w:p><w:r><w:t>SECOND</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>',
        sectPrExtra:
          '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="8"/><w:numRestart w:val="eachSect"/></w:footnotePr>',
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
        extraParts: [
          {
            path: 'word/footnotes.xml',
            xml: '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:r><w:t>one</w:t></w:r></w:p></w:footnote><w:footnote w:id="3"><w:p><w:r><w:t>two</w:t></w:r></w:p></w:footnote></w:footnotes>',
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
          },
        ],
      }),
    )
    const sections = readSections(parsed)
    const json = blocksToPmDoc(parsed.blocks, sections, pmDocOptions(parsed))
    const refs = json
      .content!.flatMap((block) => block.content ?? [])
      .filter((node) => node.type === 'docNoteRef')
    expect(refs.map((node) => [node.attrs?.num, node.attrs?.fmt])).toEqual([
      [3, 'upperRoman'],
      [8, 'lowerLetter'],
    ])
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: json as never,
    })
    try {
      expect(
        noteReferencesInEditor(
          editor.state.doc,
          sections,
          'footnote',
          noteSettingsFromParsed(parsed),
        ),
      ).toEqual({
        '2': { no: 3, fmt: 'upperRoman' },
        '3': { no: 8, fmt: 'lowerLetter' },
      })
    } finally {
      editor.destroy()
    }
  })

  it('restores distinct section overrides across Save, Undo+Save and Redo+Save', async () => {
    const original = await parseDocx(
      await buildDocx({
        bodyXml:
          '<w:p><w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="upperRoman"/><w:footnote w:id="0"/></w:footnotePr></w:sectPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t>second</w:t></w:r></w:p>',
        sectPrExtra:
          '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:footnote w:id="1"/></w:footnotePr>',
        extraParts: [
          {
            path: 'word/settings.xml',
            xml:
              '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
              '<w:footnotePr><w:numFmt w:val="lowerRoman"/><w:footnote w:id="-1"/></w:footnotePr>' +
              '</w:settings>',
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
          },
        ],
      }),
    )
    const before = noteSettingsFromParsed(original)
    const changed = applyWholeDocumentNoteSettings(before, 'footnote', {
      pos: 'beneathText',
      numFmt: 'decimal',
      numStart: 4,
      numRestart: 'eachSect',
    })
    const save = async (source: typeof original, desired: typeof before) => {
      const blocks = source.blocks
        .filter((block) => !block.hidden && block.docxIndex !== null)
        .map((block) => ({ kind: 'original' as const, docxIndex: block.docxIndex! }))
      return saveDocx(
        source,
        blocks,
        noteSettingsSaveEdits(desired, noteSettingsFromParsed(source)),
      )
    }
    const firstBytes = await save(original, changed)
    const first = await parseDocx(firstBytes)
    expect(effectiveNoteSettings(noteSettingsFromParsed(first), 'footnote')).toEqual({
      pos: 'beneathText',
      numFmt: 'decimal',
      numStart: 4,
      numRestart: 'eachSect',
    })
    const undoBytes = await save(first, before)
    const undone = await parseDocx(undoBytes)
    const undoXml = await (
      await JSZip.loadAsync(undoBytes)
    )
      .file('word/document.xml')!
      .async('string')
    expect(undone.footnoteSettings?.numFmt).toBe('lowerRoman')
    expect(undoXml).toContain('<w:numFmt w:val="upperRoman"/>')
    expect(undoXml).toContain('<w:numFmt w:val="lowerLetter"/>')
    expect(undoXml).toContain('<w:footnote w:id="0"/>')
    expect(undoXml).toContain('<w:footnote w:id="1"/>')
    const redo = await parseDocx(await save(undone, changed))
    expect(redo.footnoteSettings?.numFmt).toBe('decimal')
    expect(redo.footnoteProps?.numStart).toBe(4)
  })
})
