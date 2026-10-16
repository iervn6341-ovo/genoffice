import { Editor } from '@tiptap/core'
import type { Transaction } from '@tiptap/pm/state'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type NoteInfo } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  LAYOUT_ATTR,
  RECORD_META,
  recordLayoutChange,
  type LayoutSnapshot,
} from '../src/renderer/editor/layout-history'
import { convertNote, type ReviewContext } from '../src/renderer/review-actions'

const fixture = resolve(
  __dirname,
  '../../../docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/conversion/word-mac-conversion-before.docx',
)

async function load() {
  return parseDocx(new Uint8Array(await readFile(fixture)))
}

function setup(source: Awaited<ReturnType<typeof parseDocx>>) {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(source.blocks) as never,
  })
  let footnotes = source.footnotes
  let endnotes = source.endnotes
  let status = ''
  const baseline = {
    notes: { footnotes, endnotes },
  } as unknown as LayoutSnapshot
  const ctx = {
    editor,
    doc: { parsed: source },
    get footnotes() {
      return footnotes
    },
    get endnotes() {
      return endnotes
    },
    setStatus: (value: string) => {
      status = value
    },
    setNotesDirty: () => undefined,
    recordNoteConversion: (
      beforeNotes: { footnotes: NoteInfo[]; endnotes: NoteInfo[] },
      afterNotes: { footnotes: NoteInfo[]; endnotes: NoteInfo[] },
      markerEdit: (tr: Transaction) => Transaction,
    ) => {
      const current = (editor.state.doc.attrs[LAYOUT_ATTR] ?? baseline) as LayoutSnapshot
      const before = { ...current, notes: beforeNotes }
      const after = { ...current, notes: afterNotes }
      editor.view.dispatch(
        editor.state.tr
          .setDocAttribute(LAYOUT_ATTR, before)
          .setMeta('addToHistory', false)
          .setMeta(RECORD_META, true),
      )
      recordLayoutChange(editor, before, after, markerEdit)
      footnotes = afterNotes.footnotes
      endnotes = afterNotes.endnotes
    },
  } as unknown as ReviewContext
  const notes = () =>
    (editor.state.doc.attrs[LAYOUT_ATTR] as LayoutSnapshot | null)?.notes ?? baseline.notes!
  const save = async (parsed: Awaited<ReturnType<typeof parseDocx>>) =>
    parseDocx(
      await saveDocx(
        parsed,
        pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks,
        {
          footnotes: [...notes().footnotes],
          endnotes: [...notes().endnotes],
        },
      ),
    )
  return { editor, ctx, notes, save, status: () => status }
}

describe('Word Mac one-note conversion history', () => {
  for (const kind of ['footnote', 'endnote'] as const) {
    it(`converts a ${kind} on colliding ID and restores both native parts across save, undo and redo`, async () => {
      const source = await load()
      const { editor, ctx, notes, save, status } = setup(source)
      try {
        convertNote(ctx, kind, '1')
        expect(status()).toBe('')
        const destination = kind === 'footnote' ? 'endnotes' : 'footnotes'
        expect(notes()[destination].map((note) => [note.id, note.customMark])).toEqual(
          kind === 'footnote'
            ? [
                ['1', '♠'],
                ['2', '♥'],
              ]
            : [
                ['1', '♠'],
                ['2', '♥'],
              ],
        )
        let saved = await save(source)
        expect(saved[destination].map((note) => [note.id, note.customMark])).toEqual([
          ['1', '♠'],
          ['2', '♥'],
        ])
        const movedId = kind === 'footnote' ? '1' : '2'
        const movedEntry =
          saved.internal.noteEntryXml?.[kind === 'footnote' ? 'endnote' : 'footnote'][movedId]
        expect(movedEntry).toContain(`<w:pStyle w:val="${kind === 'footnote' ? 'af1' : 'ae'}"/>`)
        expect(movedEntry).toContain(`w:char="${kind === 'footnote' ? 'F0AA' : 'F0A9'}"`)
        expect(saved.internal.documentXml).toContain('w:customMarkFollows="1"')
        const targetKind = kind === 'footnote' ? 'endnote' : 'footnote'
        const targetStyle = kind === 'footnote' ? 'af3' : 'af0'
        expect(
          saved.internal.documentXml.match(
            new RegExp(
              `<w:rPr><w:rStyle w:val="${targetStyle}"\\/></w:rPr><w:${targetKind}Reference`,
              'g',
            ),
          ),
        ).toHaveLength(2)
        expect(editor.commands.undo()).toBe(true)
        saved = await save(saved)
        expect(saved.footnotes.map((note) => note.customMark)).toEqual(['♠'])
        expect(saved.endnotes.map((note) => note.customMark)).toEqual(['♥'])
        expect(saved.internal.noteEntryXml?.footnote['1']).toBe(
          source.internal.noteEntryXml?.footnote['1'],
        )
        expect(saved.internal.noteEntryXml?.endnote['1']).toBe(
          source.internal.noteEntryXml?.endnote['1'],
        )
        expect(saved.internal.documentXml).toContain(
          '<w:rPr><w:rStyle w:val="af0"/></w:rPr><w:footnoteReference',
        )
        expect(saved.internal.documentXml).toContain(
          '<w:rPr><w:rStyle w:val="af3"/></w:rPr><w:endnoteReference',
        )
        expect(editor.commands.redo()).toBe(true)
        saved = await save(saved)
        expect(saved[destination].map((note) => note.customMark)).toEqual(['♠', '♥'])
        expect(
          saved.internal.documentXml.match(
            new RegExp(
              `<w:rPr><w:rStyle w:val="${targetStyle}"\\/></w:rPr><w:${targetKind}Reference`,
              'g',
            ),
          ),
        ).toHaveLength(2)
      } finally {
        editor.destroy()
      }
    })
  }

  it('leaves a protected destination note and all references unchanged', async () => {
    const source = await load()
    source.endnotes = source.endnotes.map((note) => ({ ...note, protectedContent: true }))
    const { editor, ctx, notes, status } = setup(source)
    try {
      const before = editor.state.doc.toJSON()
      convertNote(ctx, 'footnote', '1')
      expect(status()).not.toBe('')
      expect(editor.state.doc.toJSON()).toEqual(before)
      expect(notes().footnotes).toHaveLength(1)
      expect(notes().endnotes).toHaveLength(1)
    } finally {
      editor.destroy()
    }
  })
})
