/**
 * C7: "type a word, change the margins, ⌘Z" undid the word. Layout changes
 * now commit a snapshot as a doc-attribute step on the same history stack.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { undo, redo } from '@tiptap/pm/history'
import type { SectionSettings } from '@genoffice/docx-engine'
import {
  recordLayoutChange,
  restoredLayout,
  withRestoreDirty,
  type LayoutSnapshot,
} from '../src/renderer/editor/layout-history'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

const margins = (top: number) => ({ marginTop: top }) as unknown as SectionSettings

const BASE: LayoutSnapshot = {
  section: margins(1440),
  sectionDirty: false,
  sections: [{ settings: margins(1440) }],
  sectionsDirty: [],
  titlePg: false,
  titlePgDirty: false,
  evenOddHf: false,
  evenOddHfDirty: false,
  pgNumEdit: null,
  pgNumDirtySections: [],
}

async function setup() {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: { type: 'doc', content: [{ type: 'docParagraph' }] } as never,
  })
  editors.push(editor)
  // the App: live layout state, restored from undo/redo like App.tsx's listener
  let live = BASE
  const restores: LayoutSnapshot[] = []
  editor.on('transaction', ({ transaction }) => {
    const restored = restoredLayout(transaction, transaction.before.attrs.layout)
    if (restored) {
      live = withRestoreDirty(restored, live)
      restores.push(live)
    }
  })
  const commit = (change: Partial<LayoutSnapshot>) => {
    const before = live
    live = { ...before, ...change }
    recordLayoutChange(editor, before, live)
  }
  return {
    editor,
    commit,
    restores,
    get live() {
      return live
    },
  }
}

describe('layout on the undo stack (C7)', () => {
  it('⌘Z after typing then changing margins undoes the margins first', async () => {
    const t = await setup()
    const { editor } = t
    editor.commands.insertContentAt(1, 'word')
    t.commit({ section: margins(720), sectionDirty: true, sections: [{ settings: margins(720) }] })
    expect(t.live.section).toEqual(margins(720))

    undo(editor.state, editor.view.dispatch)
    expect(t.live.section).toEqual(margins(1440)) // margins reverted
    expect(editor.state.doc.textContent).toBe('word') // word kept

    redo(editor.state, editor.view.dispatch)
    expect(t.live.section).toEqual(margins(720))

    undo(editor.state, editor.view.dispatch)
    undo(editor.state, editor.view.dispatch)
    expect(editor.state.doc.textContent).toBe('')
    expect(t.live.section).toEqual(margins(1440))
  })

  it('successive layout changes are separate steps (orientation, first page, page numbers)', async () => {
    const t = await setup()
    const { editor } = t
    t.commit({ section: margins(720), sectionDirty: true })
    t.commit({ titlePg: true, titlePgDirty: true })
    t.commit({ pgNumEdit: { fmt: 'lowerRoman', start: 3 } })
    undo(editor.state, editor.view.dispatch)
    expect(t.live.pgNumEdit).toBeNull()
    expect(t.live.titlePg).toBe(true)
    undo(editor.state, editor.view.dispatch)
    expect(t.live.titlePg).toBe(false)
    expect(t.live.section).toEqual(margins(720))
    undo(editor.state, editor.view.dispatch)
    expect(t.live.section).toEqual(margins(1440))
    expect(undo(editor.state)).toBe(false) // baseline pin is not an undo step
  })

  it('undo past a save keeps the reverted value marked dirty', () => {
    const saved: LayoutSnapshot = { ...BASE, section: margins(720), sectionDirty: false }
    const restored = withRestoreDirty(BASE, saved)
    expect(restored.section).toEqual(margins(1440))
    expect(restored.sectionDirty).toBe(true)
    const untouched = withRestoreDirty(BASE, BASE)
    expect(untouched.sectionDirty).toBe(false)
    const nonFinal = withRestoreDirty(
      { ...BASE, sections: [{ settings: margins(1) }, { settings: margins(2) }] },
      { ...BASE, sections: [{ settings: margins(9) }, { settings: margins(2) }] },
    )
    expect(nonFinal.sectionsDirty).toEqual([0])
  })

  it('a text edit does not trigger a layout restore', async () => {
    const t = await setup()
    t.commit({ titlePg: true, titlePgDirty: true })
    t.editor.commands.insertContentAt(1, 'x')
    expect(t.restores).toHaveLength(0)
  })
})
