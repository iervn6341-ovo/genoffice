import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import type { SectionInfo } from '@genoffice/docx-engine'
import {
  numberingSectionBreak,
  withPageNumber,
  withoutPageNumber,
} from '../src/renderer/editor/page-numbering'
import { editorExtensions } from '../src/renderer/editor/extensions'
const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((e) => e.destroy()))
it('splits a measured paragraph with marks and keeps the original anchor on the editable tail', () => {
  const e = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'Before after', marks: [{ type: 'bold' }] }],
        },
      ],
    } as never,
  })
  editors.push(e)
  const sec = {
    firstBlockIndex: 0,
    lastBlockIndex: 1,
    sectPrXml: '<w:sectPr/>',
    settings: {},
    startType: 'nextPage',
    titlePg: false,
    headerRefs: {},
    footerRefs: {},
  } as SectionInfo
  const change = numberingSectionBreak(e, 8, [sec], 0, sec.sectPrXml)
  e.view.dispatch(change.mutate(e.state.tr))
  expect(e.state.doc.childCount).toBe(3)
  expect(e.state.doc.child(0).textContent).toBe('Before ')
  expect(e.state.doc.child(2).textContent).toBe('after')
  expect(e.state.doc.child(2).attrs.docxIndex).toBe(0)
  expect(e.state.doc.child(2).firstChild?.marks[0]?.type.name).toBe('bold')
  expect(change.sections[0]!.lastBlockIndex).toBeLessThan(0)
})
it('removes PAGE without deleting adjacent footer text or other native fields', () => {
  const field = {
    xml: '<w:fldSimple w:instr=" PAGE "/>',
    instruction: ' PAGE ',
    token: '',
    displayText: '1',
    resultRuns: [],
  }
  const before = {
    text: 'Private \uE001',
    pageNumber: true,
    paras: [
      {
        runs: [
          { text: 'Private ', bold: true },
          { text: '\uE001', hfField: field },
        ],
      },
    ],
  }
  const hidden = withoutPageNumber(before)
  expect(hidden.text).toBe('Private ')
  expect(hidden.paras?.[0]?.runs[0]).toEqual(before.paras[0].runs[0])
  expect(hidden.paras?.[0]?.runs[1]?.hfField).toBeUndefined()
  expect(withPageNumber(hidden).paras?.at(-1)?.runs[0]?.text).toBe('\uE001')
})
