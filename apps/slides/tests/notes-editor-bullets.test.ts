import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import {
  createBlankPptx,
  openPptx,
  savePptx,
  setSlideNotes,
  getSlideNotesParagraphs,
  notesPathForSlide,
} from '@genoffice/pptx-engine'
import { runTxn } from '@genoffice/pptx-ops'
import { NotesEditor, readNotesEditor } from '../src/renderer/components/NotesEditor'
import { notesParagraphViews } from '../src/main/notes-view'
import type { NotesParagraphView } from '../src/shared/ipc'
import { onParagraphFormat } from '../src/renderer/style-actions'
import type { ActionCtx } from '../src/renderer/action-context'

async function notesFixture() {
  const opened = await openPptx(await createBlankPptx())
  setSlideNotes(opened, 0, 'Original plain line')
  const master = [...opened.archive.entries.keys()].find((p) => /notesMaster\d+\.xml$/.test(p))!
  opened.archive.entries.set(
    master,
    Buffer.from(
      opened.archive
        .readText(master)!
        .replace(
          /(<p:ph type="body"[\s\S]*?)<a:lstStyle\/>/,
          '$1<a:lstStyle><a:lvl1pPr><a:buChar char="●"/></a:lvl1pPr></a:lstStyle>',
        ),
    ),
  )
  const path = notesPathForSlide(opened.archive, opened.deck.slides[0]!.path)!
  opened.archive.entries.set(
    path,
    Buffer.from(opened.archive.readText(path)!.replace('<a:p>', '<a:p><a:pPr><a:buNone/></a:pPr>')),
  )
  return { opened, path }
}

describe('notes authored list intent', () => {
  it('new plain DOM paragraphs persist as no bullet even with a bulleted notes master', async () => {
    const { opened } = await notesFixture()
    const root = document.createElement('div')
    root.innerHTML =
      '<div data-src-para="0"><span data-src-run="0">Original plain line</span></div><div>New plain line</div>'
    const draft = readNotesEditor(root)
    expect(draft.paragraphs[0]!.bullet).toBeUndefined()
    expect(draft.paragraphs[1]!.bullet).toBe('none')
    const old = getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)
    expect(
      runTxn(opened, {
        ops: [{ op: 'setNotes', target: { slide: 0 }, paragraphs: draft.paragraphs }],
      }).applied,
    ).toBe(true)
    const back = await openPptx(await savePptx(opened))
    const paras = getSlideNotesParagraphs(back.archive, back.deck.slides[0]!.path)
    expect(paras.map((p) => p.runs.map((r) => r.text).join(''))).toEqual([
      'Original plain line',
      'New plain line',
    ])
    expect(notesParagraphViews(paras).map((p) => p.bullet)).toEqual([undefined, undefined])
    expect(notesParagraphViews(paras)[0]).toEqual(notesParagraphViews(old)[0])
  })
  it('preserves imported bullet metadata and explicit ribbon choices', () => {
    Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const editorRef = { current: null as HTMLDivElement | null }
    const paragraphs: NotesParagraphView[] = [
      {
        level: 0,
        bullet: '3.',
        bulletKind: 'number',
        numType: 'arabicPeriod',
        runs: [
          {
            text: 'Numbered',
            fontSizePt: 14,
            bold: false,
            italic: false,
            underline: false,
            strike: false,
            baseline: 0,
            fontExplicit: false,
          },
        ],
      },
    ]
    try {
      act(() =>
        root.render(
          React.createElement(NotesEditor, {
            editorRef,
            paragraphs,
            placeholder: 'Notes',
            defaultFont: null,
            onEdit: () => {},
            onFocus: () => {},
            onBlur: () => {},
          }),
        ),
      )
      expect(
        editorRef.current!.querySelector('[data-bullet-text]')!.getAttribute('data-bullet-text'),
      ).toBe('3.')
      expect(readNotesEditor(editorRef.current!).paragraphs[0]!.bullet).toBeUndefined()
      const p = document.createElement('div')
      p.dataset.bullet = 'number'
      p.dataset.numType = 'romanUcPeriod'
      p.textContent = 'Explicit new list'
      editorRef.current!.appendChild(p)
      expect(readNotesEditor(editorRef.current!).paragraphs[1]).toMatchObject({
        bullet: 'number',
        numType: 'romanUcPeriod',
      })
    } finally {
      act(() => root.unmount())
      host.remove()
    }
  })
  it('the notes ribbon removes an inherited bullet, signals the edit and persists no bullet on reopen', async () => {
    Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
    const { opened, path } = await notesFixture()
    opened.archive.entries.set(
      path,
      Buffer.from(opened.archive.readText(path)!.replace('<a:buNone/>', '')),
    )
    const old = getSlideNotesParagraphs(opened.archive, opened.deck.slides[0]!.path)
    expect(notesParagraphViews(old)[0]!.bullet).toBe('●')
    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)
    const editorRef = { current: null as HTMLDivElement | null }
    const onEdit = vi.fn()
    try {
      act(() =>
        root.render(
          React.createElement(NotesEditor, {
            editorRef,
            paragraphs: notesParagraphViews(old),
            placeholder: 'Notes',
            defaultFont: null,
            onEdit,
            onFocus: () => {},
            onBlur: () => {},
          }),
        ),
      )
      const block = editorRef.current!.firstElementChild as HTMLElement
      const range = document.createRange()
      range.selectNodeContents(block)
      window.getSelection()!.removeAllRanges()
      window.getSelection()!.addRange(range)
      await act(async () => {
        onParagraphFormat(
          {
            editingNotes: true,
            editing: null,
            editingCell: null,
            selectedIds: [],
          } as unknown as ActionCtx,
          { bullet: 'none' },
        )
        await Promise.resolve()
      })
      expect(block.hasAttribute('data-bullet-text')).toBe(false)
      expect(onEdit).toHaveBeenCalled()
      const draft = readNotesEditor(editorRef.current!)
      expect(draft.paragraphs[0]!.bullet).toBe('none')
      expect(
        runTxn(opened, {
          ops: [{ op: 'setNotes', target: { slide: 0 }, paragraphs: draft.paragraphs }],
        }).applied,
      ).toBe(true)
      const back = await openPptx(await savePptx(opened))
      expect(
        notesParagraphViews(getSlideNotesParagraphs(back.archive, back.deck.slides[0]!.path))[0]!
          .bullet,
      ).toBeUndefined()
    } finally {
      act(() => root.unmount())
      host.remove()
      window.getSelection()!.removeAllRanges()
    }
  })
})
