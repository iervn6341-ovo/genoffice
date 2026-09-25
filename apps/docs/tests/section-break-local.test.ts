/**
 * C6: a section break used to force a silent save + reparse (ignoring
 * AutoSave off) and so wiped the undo stack. The break is now an ordinary
 * undoable paragraph; the start type chosen for the section AFTER it lives
 * on the paragraph (breakStartType) and is resolved into the next sectPr at
 * save time, so undo/redo and the saved file always agree.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { closeHistory, undo, redo } from '@tiptap/pm/history'
import {
  applySectionStartType,
  buildBlankDocx,
  parseDocx,
  readSections,
  saveDocx,
  type ParsedDocFull,
  type SectionInfo,
} from '@genoffice/docx-engine'
import {
  blocksToPmDoc,
  pmDocOptions,
  pmDocToSavePlan,
  type PmNode,
} from '../src/renderer/editor/convert'
import { insertedBreakTypes } from '../src/renderer/ai/pending-sections'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

async function open(bytes?: Uint8Array) {
  const { editorExtensions } = await import('../src/renderer/editor/extensions')
  const parsed = (await parseDocx(bytes ?? (await buildBlankDocx()))) as ParsedDocFull
  const sections = readSections(parsed)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editors.push(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks, sections, pmDocOptions(parsed)) as never)
  return { editor, parsed, sections }
}

function typeParagraphs(editor: Editor, texts: string[]) {
  const nodes = texts.map((text) =>
    editor.schema.nodeFromJSON({ type: 'docParagraph', content: [{ type: 'text', text }] }),
  )
  editor.view.dispatch(editor.state.tr.replaceWith(0, editor.state.doc.content.size, nodes))
}

/** same node the Layout ▸ Breaks menu inserts (App.insertSectionBreak) */
function insertBreak(
  editor: Editor,
  afterTop: number,
  sectPr: string,
  type: SectionInfo['startType'],
) {
  let pos = 0
  for (let i = 0; i <= afterTop; i++) pos += editor.state.doc.child(i).nodeSize
  // a separate undo step, as when the user picks Breaks from the ribbon
  editor.view.dispatch(
    closeHistory(editor.state.tr).insert(
      pos,
      editor.schema.nodes.docProtected!.create({
        blockType: 'passthrough',
        label: 'Section break paragraph',
        genXml: `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`,
        breakStartType: type,
      }),
    ),
  )
}

async function save(editor: Editor, parsed: ParsedDocFull, sections: SectionInfo[]) {
  const types = insertedBreakTypes(editor.state.doc, sections, applySectionStartType)
  const json = editor.getJSON() as PmNode
  for (const [i, genXml] of types.genXml)
    json.content![i]!.attrs = { ...json.content![i]!.attrs, genXml }
  const bytes = await saveDocx(parsed, pmDocToSavePlan(json, parsed.blocks).saveBlocks, {
    sectionStartType: types.trailing ?? undefined,
  })
  return { bytes, types, reparsed: readSections(await parseDocx(bytes)) }
}

describe('section break without save/reparse (C6)', () => {
  it('insert → undo → redo keeps the derived start type in step with the document', async () => {
    const { editor, sections } = await open()
    typeParagraphs(editor, ['one', 'two'])
    const sectPr = sections[sections.length - 1]!.sectPrXml
    insertBreak(editor, 0, sectPr, 'continuous')
    expect(insertedBreakTypes(editor.state.doc, sections, applySectionStartType).trailing).toBe(
      'continuous',
    )
    // typing before the break stays undoable after it: history was not reset
    undo(editor.state, editor.view.dispatch)
    expect(
      insertedBreakTypes(editor.state.doc, sections, applySectionStartType).trailing,
    ).toBeNull()
    expect(editor.state.doc.textContent).toBe('onetwo')
    redo(editor.state, editor.view.dispatch)
    expect(insertedBreakTypes(editor.state.doc, sections, applySectionStartType).trailing).toBe(
      'continuous',
    )
  })

  it('two breaks in one section: the first break type goes into the second break sectPr', async () => {
    const { editor, sections } = await open()
    typeParagraphs(editor, ['a', 'b', 'c'])
    const sectPr = sections[sections.length - 1]!.sectPrXml
    insertBreak(editor, 0, sectPr, 'continuous') // after "a"
    insertBreak(editor, 2, sectPr, 'oddPage') // after "b"
    const types = insertedBreakTypes(editor.state.doc, sections, applySectionStartType)
    expect(types.trailing).toBe('oddPage')
    expect([...types.genXml.keys()]).toEqual([3])
    expect(types.genXml.get(3)).toMatch(/<w:type w:val="continuous"\/>/)
  })

  it('DOCX round-trip: sections and start types are written by the next save', async () => {
    const { editor, parsed, sections } = await open()
    typeParagraphs(editor, ['a', 'b', 'c'])
    const sectPr = sections[sections.length - 1]!.sectPrXml
    insertBreak(editor, 0, sectPr, 'continuous')
    insertBreak(editor, 2, sectPr, 'evenPage')
    const { reparsed, bytes } = await save(editor, parsed, sections)
    expect(reparsed).toHaveLength(3)
    expect(reparsed.map((s) => s.startType)).toEqual(['nextPage', 'continuous', 'evenPage'])

    // the saved file reopens with all three sections
    const second = await open(bytes)
    expect(second.sections).toHaveLength(3)
  })

  it('an undone break writes nothing', async () => {
    const { editor, parsed, sections } = await open()
    typeParagraphs(editor, ['a', 'b'])
    insertBreak(editor, 0, sections[sections.length - 1]!.sectPrXml, 'continuous')
    undo(editor.state, editor.view.dispatch)
    const { reparsed } = await save(editor, parsed, sections)
    expect(reparsed).toHaveLength(1)
    expect(reparsed[0]!.startType).toBe(sections[0]!.startType)
  })
})
