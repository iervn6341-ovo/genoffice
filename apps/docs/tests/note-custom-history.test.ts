import { Editor } from '@tiptap/core'
import type { Transaction } from '@tiptap/pm/state'
import { readFile } from 'node:fs/promises'
import { resolve } from 'node:path'
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
import { submitNote, type ReviewContext } from '../src/renderer/review-actions'

describe('custom note mark history', () => {
  it('keeps a Word-native symbol reference and note body through save, undo, redo and reopen', async () => {
    const source = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(source.blocks) as never,
    })
    try {
      let notes: NoteInfo[] = []
      const baseline = { notes: { footnotes: [], endnotes: [] } } as unknown as LayoutSnapshot
      const ctx = {
        editor,
        notePrompt: { kind: 'footnote' },
        get footnotes() {
          return notes
        },
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
          const before = (editor.state.doc.attrs[LAYOUT_ATTR] ?? baseline) as LayoutSnapshot
          recordLayoutChange(
            editor,
            before,
            { ...before, notes: { footnotes: value, endnotes: [] } },
            markerEdit,
          )
          notes = value
        },
      } as unknown as ReviewContext
      const entries = () => editor.state.doc.attrs.layout?.notes?.footnotes ?? []
      const save = (parsed: typeof source) =>
        saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks, {
          footnotes: entries(),
        })

      submitNote(ctx, 'Symbol note', '♠')
      expect(notes).toEqual([
        {
          id: '1',
          text: 'Symbol note',
          customMark: '♠',
          customMarkSymbol: { font: 'Symbol', char: 'F0AA' },
        },
      ])
      const first = await parseDocx(await save(source))
      expect(first.footnotes[0].customMark).toBe('♠')
      expect(first.blocks[0].runs?.find((run) => run.noteRef)?.noteRef?.customMark).toBe('♠')
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
      const undone = await parseDocx(await save(first))
      expect(undone.footnotes).toEqual([])
      expect(undone.blocks[0].runs?.some((run) => run.noteRef)).toBe(false)
      expect(editor.commands.redo()).toBe(true)
      const redone = await parseDocx(await save(undone))
      expect(redone.footnotes[0].customMark).toBe('♠')
      expect(redone.footnotes[0].text).toBe('Symbol note')
      expect(redone.blocks[0].runs?.find((run) => run.noteRef)?.noteRef?.customMarkSymbol).toEqual({
        font: 'Symbol',
        char: 'F0AA',
      })
    } finally {
      editor.destroy()
    }
  })

  it('changes and restores an imported Word Mac mark across saves without changing endnotes', async () => {
    const source = await parseDocx(
      new Uint8Array(
        await readFile(
          resolve(
            __dirname,
            '../../../docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/word-mac-custom-marks-source.docx',
          ),
        ),
      ),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(source.blocks) as never,
    })
    try {
      let notes = source.footnotes
      const baseline = {
        notes: { footnotes: source.footnotes, endnotes: source.endnotes },
      } as unknown as LayoutSnapshot
      const ctx = {
        editor,
        notePrompt: { kind: 'footnote', id: '1' },
        get footnotes() {
          return notes
        },
        endnotes: source.endnotes,
        setFootnotes: (value: NoteInfo[]) => {
          notes = value
        },
        setStatus: () => undefined,
        setNotesDirty: () => undefined,
        recordNoteChange: (
          _kind: 'footnote' | 'endnote',
          value: NoteInfo[],
          markerEdit?: (tr: Transaction) => Transaction,
        ) => {
          const before = (editor.state.doc.attrs[LAYOUT_ATTR] ?? baseline) as LayoutSnapshot
          recordLayoutChange(
            editor,
            before,
            {
              ...before,
              notes: { footnotes: value, endnotes: source.endnotes },
            },
            markerEdit,
          )
          notes = value
        },
      } as unknown as ReviewContext
      const save = (parsed: typeof source) =>
        saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks, {
          footnotes: editor.state.doc.attrs.layout?.notes?.footnotes ?? source.footnotes,
          endnotes: source.endnotes,
        })

      submitNote(ctx, 'Custom symbol footnote body.', '♥')
      let saved = await parseDocx(await save(source))
      expect(saved.footnotes[0].customMark).toBe('♥')
      expect(saved.endnotes[0].customMark).toBe('♥')
      expect(editor.commands.undo()).toBe(true)
      saved = await parseDocx(await save(saved))
      expect(saved.footnotes[0].customMark).toBe('♠')
      expect(editor.commands.redo()).toBe(true)
      saved = await parseDocx(await save(saved))
      expect(saved.footnotes[0].customMark).toBe('♥')
      expect(saved.footnotes[0].text).toBe('Custom symbol footnote body.')
    } finally {
      editor.destroy()
    }
  })
})
