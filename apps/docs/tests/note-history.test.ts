import { Editor } from '@tiptap/core'
import type { Transaction } from '@tiptap/pm/state'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type NoteInfo } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  LAYOUT_ATTR,
  recordLayoutChange,
  type LayoutSnapshot,
} from '../src/renderer/editor/layout-history'
import { sameBodyAfterManagedRPrRebase } from '../src/renderer/editor/save-rebase'
import { t } from '../src/renderer/i18n/locale'
import { deleteNote, submitNote, type ReviewContext } from '../src/renderer/review-actions'

describe('note history', () => {
  it('inserts a native footnote reference and body in one undoable edit across save', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    try {
      let notes: NoteInfo[] = []
      const base = { notes: { footnotes: [], endnotes: [] } } as unknown as LayoutSnapshot
      const ctx = {
        editor,
        notePrompt: { kind: 'footnote' },
        footnotes: notes,
        endnotes: [],
        setFootnotes: (value: NoteInfo[]) => {
          notes = value
        },
        setNotesDirty: () => undefined,
        recordNoteChange: (
          _kind: 'footnote' | 'endnote',
          value: NoteInfo[],
          markerEdit?: (tr: Transaction) => Transaction,
        ) => {
          recordLayoutChange(
            editor,
            base,
            { ...base, notes: { footnotes: value, endnotes: [] } },
            markerEdit,
          )
          notes = value
        },
      } as unknown as ReviewContext
      submitNote(ctx, 'New native note')
      expect(notes).toEqual([{ id: '1', text: 'New native note' }])
      expect(editor.state.doc.textBetween(0, editor.state.doc.content.size)).toContain('Body')
      expect(
        editor.state.doc
          .toJSON()
          .content?.[0]?.content?.some((n: { type: string }) => n.type === 'docNoteRef'),
      ).toBe(true)
      const save = (source: typeof parsed, entries: NoteInfo[]) =>
        saveDocx(source, pmDocToSavePlan(editor.getJSON() as PmNode, source.blocks).saveBlocks, {
          footnotes: entries,
        })
      const first = await parseDocx(await save(parsed, notes))
      expect(first.footnotes).toEqual([{ id: '1', text: 'New native note' }])
      expect(first.blocks[0].type).toBe('paragraph')
      expect(
        sameBodyAfterManagedRPrRebase(
          editor.state.doc,
          editor.schema.nodeFromJSON({
            ...blocksToPmDoc(first.blocks),
            attrs: { [LAYOUT_ATTR]: editor.state.doc.attrs[LAYOUT_ATTR] },
          }),
        ),
      ).toBe(true)
      expect(editor.commands.undo()).toBe(true)
      expect(
        editor.state.doc
          .toJSON()
          .content?.[0]?.content?.some((n: { type: string }) => n.type === 'docNoteRef'),
      ).toBe(false)
      const undoneNotes = editor.state.doc.attrs.layout?.notes?.footnotes ?? []
      const undone = await parseDocx(await save(first, undoneNotes))
      expect(undone.footnotes).toEqual([])
      expect(editor.commands.redo()).toBe(true)
      const redoneNotes = editor.state.doc.attrs.layout?.notes?.footnotes ?? []
      const redone = await parseDocx(await save(undone, redoneNotes))
      expect(redone.footnotes).toEqual([{ id: '1', text: 'New native note' }])
    } finally {
      editor.destroy()
    }
  })

  it('edits and deletes an imported endnote with undo and redo on the native part', async () => {
    const originalEntry =
      '<w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r>' +
      '<w:r><w:t>Original note</w:t></w:r></w:p></w:endnote>'
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p>',
        extraRels:
          '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>',
        extraParts: [
          {
            path: 'word/endnotes.xml',
            xml:
              '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
              originalEntry +
              '</w:endnotes>',
            contentType:
              'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
          },
        ],
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    try {
      let notes = parsed.endnotes
      const base = {
        notes: { footnotes: [], endnotes: parsed.endnotes },
      } as unknown as LayoutSnapshot
      const ctx = {
        editor,
        notePrompt: { kind: 'endnote', id: '2' },
        footnotes: [],
        get endnotes() {
          return notes
        },
        setEndnotes: (value: NoteInfo[]) => {
          notes = value
        },
        setNotesDirty: () => undefined,
        recordNoteChange: (
          _kind: 'footnote' | 'endnote',
          value: NoteInfo[],
          markerEdit?: (tr: Transaction) => Transaction,
        ) => {
          const before = (editor.state.doc.attrs[LAYOUT_ATTR] ?? base) as LayoutSnapshot
          recordLayoutChange(
            editor,
            before,
            { ...before, notes: { footnotes: [], endnotes: value } },
            markerEdit,
          )
          notes = value
        },
      } as unknown as ReviewContext
      const entries = () => editor.state.doc.attrs[LAYOUT_ATTR]?.notes?.endnotes ?? parsed.endnotes
      const save = (source: typeof parsed) =>
        saveDocx(source, pmDocToSavePlan(editor.getJSON() as PmNode, source.blocks).saveBlocks, {
          endnotes: entries(),
        })
      submitNote(ctx, 'Edited note')
      let saved = await parseDocx(await save(parsed))
      expect(saved.endnotes.map(({ text }) => text)).toEqual(['Edited note'])
      expect(editor.commands.undo()).toBe(true)
      saved = await parseDocx(await save(saved))
      expect(saved.endnotes.map(({ text }) => text)).toEqual(['Original note'])
      expect(editor.commands.redo()).toBe(true)
      saved = await parseDocx(await save(saved))
      expect(saved.endnotes.map(({ text }) => text)).toEqual(['Edited note'])

      deleteNote(ctx, 'endnote', '2')
      saved = await parseDocx(await save(saved))
      expect(saved.endnotes).toEqual([])
      expect(
        editor.state.doc
          .toJSON()
          .content?.[0]?.content?.some((n: { type: string }) => n.type === 'docNoteRef'),
      ).toBe(false)
      expect(editor.commands.undo()).toBe(true)
      saved = await parseDocx(await save(saved))
      expect(saved.endnotes.map(({ text }) => text)).toEqual(['Edited note'])
      expect(editor.commands.redo()).toBe(true)
      saved = await parseDocx(await save(saved))
      expect(saved.endnotes).toEqual([])
    } finally {
      editor.destroy()
    }
  })

  it('refuses deletion when the only reference is protected or missing', () => {
    const makeEditor = (protectedXml?: string, fromParsed = false) =>
      new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: {
          type: 'doc',
          content: protectedXml
            ? [
                {
                  type: 'docProtected',
                  attrs: {
                    ...(fromParsed ? { docxIndex: 0 } : { genXml: protectedXml }),
                    blockType: 'passthrough',
                  },
                },
              ]
            : [{ type: 'paragraph', content: [{ type: 'text', text: 'No reference' }] }],
        } as never,
      })
    for (const { xml, fromParsed } of [
      { xml: '<w:p><w:r><w:footnoteReference w:id="7"/></w:r></w:p>' },
      { xml: "<w:p><w:r><w:footnoteReference w:id='7'/></w:r></w:p>" },
      { xml: '<w:p><w:r><w:footnoteReference w:id="7"/></w:r></w:p>', fromParsed: true },
      { xml: undefined },
    ]) {
      const editor = makeEditor(xml, fromParsed)
      try {
        let changed = false
        let status = ''
        const ctx = {
          editor,
          doc: fromParsed ? { parsed: { blocks: [{ docxIndex: 0, originalXml: xml }] } } : null,
          footnotes: [{ id: '7', text: 'Keep me' }],
          endnotes: [],
          setFootnotes: () => {
            changed = true
          },
          setNotesDirty: () => {
            changed = true
          },
          recordNoteChange: () => {
            changed = true
          },
          setStatus: (value: string) => {
            status = value
          },
        } as unknown as ReviewContext
        deleteNote(ctx, 'footnote', '7')
        expect(changed).toBe(false)
        expect(status).toBe(t('appNoteReferenceProtectedError'))
        expect(editor.state.doc.toJSON().content?.length).toBe(1)
      } finally {
        editor.destroy()
      }
    }
  })
})
