import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateTocFieldXml, parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { collectHeadings, collectOutlineItems } from '../src/renderer/editor/headings'

const editors = new Set<Editor>()

afterEach(() => {
  for (const editor of editors) editor.destroy()
  editors.clear()
})

const p = (pPr: string, text: string) =>
  `<w:p><w:pPr>${pPr}</w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

/** H1/H2/H3 via built-in styles, direct outlineLvl and a basedOn custom style */
async function openMixedHeadingsDoc(): Promise<Editor> {
  const bytes = await buildDocx({
    extraStylesXml:
      '<w:style w:type="paragraph" w:styleId="MySub"><w:name w:val="My Sub"/><w:basedOn w:val="Heading2"/></w:style>',
    bodyXml: [
      p('<w:pStyle w:val="Heading1"/>', 'Chapter 1'),
      p('', 'body text'),
      p('<w:pStyle w:val="Heading2"/>', 'Section 1.1'),
      p('<w:outlineLvl w:val="2"/>', 'Topic 1.1.1'),
      p('<w:pStyle w:val="MySub"/>', 'Section 1.2'),
    ].join(''),
  })
  const parsed = await parseDocx(bytes)
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.add(editor)
  return editor
}

describe('TOC heading collection', () => {
  it('collects H1-H3 headings with levels in document order', async () => {
    const editor = await openMixedHeadingsDoc()
    const headings = collectHeadings(editor.state.doc)
    expect(headings.map((h) => [h.text, h.level])).toEqual([
      ['Chapter 1', 1],
      ['Section 1.1', 2],
      ['Topic 1.1.1', 3],
      ['Section 1.2', 2],
    ])
  })

  it('generates a level-indented TOC field from the collected headings', async () => {
    const editor = await openMixedHeadingsDoc()
    const entries = collectHeadings(editor.state.doc).map(({ level, text }) => ({ level, text }))
    const frags = generateTocFieldXml(entries)
    expect(frags[0]).toContain(' TOC \\o "1-3" ')
    expect(frags.map((f) => /w:pStyle w:val="(TOC\d)"/.exec(f)?.[1])).toEqual([
      'TOC1',
      'TOC2',
      'TOC3',
      'TOC2',
    ])
  })

  it('collectOutlineItems falls back to headings when no TOC exists', async () => {
    const editor = await openMixedHeadingsDoc()
    const { items } = collectOutlineItems(editor.state.doc)
    expect(items.map((item) => [item.text, item.level, item.isToc])).toEqual([
      ['Chapter 1', 1, false],
      ['Section 1.1', 2, false],
      ['Topic 1.1.1', 3, false],
      ['Section 1.2', 2, false],
    ])
    // Verifies body text paragraphs are excluded
    expect(items.some((item) => item.text === 'body text')).toBe(false)
  })

  it('collectOutlineItems prioritizes document TOC entries and resolves bookmarks', async () => {
    const bytes = await buildDocx({
      bodyXml: [
        p('', '目次'),
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc101"><w:r><w:t>圖次</w:t></w:r><w:r><w:tab/><w:t>iii</w:t></w:r></w:hyperlink></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:hyperlink w:anchor="_Toc102"><w:r><w:t>第1章 緒論</w:t></w:r><w:r><w:tab/><w:t>1</w:t></w:r></w:hyperlink></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr><w:hyperlink w:anchor="_Toc103"><w:r><w:t>1.1 研究背景</w:t></w:r><w:r><w:tab/><w:t>2</w:t></w:r></w:hyperlink></w:p>',
        p('', 'preliminary content'),
        '<w:p><w:bookmarkStart w:id="1" w:name="_Toc101"/><w:r><w:t>圖次標題內容</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
        '<w:p><w:pStyle w:val="Heading1"/><w:bookmarkStart w:id="2" w:name="_Toc102"/><w:r><w:t>第1章 緒論</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>',
        '<w:p><w:pStyle w:val="Heading2"/><w:bookmarkStart w:id="3" w:name="_Toc103"/><w:r><w:t>1.1 研究背景</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>',
      ].join(''),
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.add(editor)

    const { items } = collectOutlineItems(editor.state.doc)
    expect(items.length).toBe(3)
    expect(items[0]).toMatchObject({
      text: '圖次',
      page: 'iii',
      level: 1,
      isToc: true,
      anchor: '_Toc101',
    })
    expect(items[1]).toMatchObject({
      text: '第1章 緒論',
      page: '1',
      level: 1,
      isToc: true,
      anchor: '_Toc102',
    })
    expect(items[2]).toMatchObject({
      text: '1.1 研究背景',
      page: '2',
      level: 2,
      isToc: true,
      anchor: '_Toc103',
    })
    // Verified bookmark target resolution: targetPos points to destination in body
    expect(items[1].targetPos).toBeDefined()
    expect(items[1].targetPos).toBeGreaterThan(items[1].pos)
  })
})
