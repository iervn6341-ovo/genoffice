import { afterEach, describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import {
  DEFAULT_TOC_OPTIONS,
  generateTocFieldXml,
  parseDocx,
  saveDocx,
  tocInstructionFromXml,
  tocOptionsFromField,
} from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { rebaseTocAnchors, sameBodyAfterTocRebase } from '../src/renderer/editor/save-rebase'
import {
  anchorTocReplacement,
  collectTocSources,
  findTocField,
  preserveTocHistorySource,
  tocNodeXml,
  updateTocPageCaches,
} from '../src/renderer/editor/toc-actions'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

const heading = (style: string, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`

const date =
  '<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>13 October 2026</w:t></w:r></w:fldSimple></w:p>'

async function documentXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

async function save(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  return saveDocx(parsed, pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks).saveBlocks)
}

describe('custom TOC update', () => {
  it('collects four heading levels plus a style mapped to level two', async () => {
    const bytes = await buildDocx({
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="Custom">' +
        '<w:name w:val="Custom Title"/></w:style>',
      bodyXml:
        [1, 2, 3, 4].map((level) => heading(`Heading${level}`, `Heading ${level}`)).join('') +
        heading('Custom', 'Custom chapter'),
    })
    const parsed = await parseDocx(bytes)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.push(editor)
    const entries = collectTocSources(editor, parsed.blocks, parsed.styles, {
      ...DEFAULT_TOC_OPTIONS,
      levels: 4,
      styleLevels: { 'Custom Title': 2 },
    })
    expect(entries.map(({ level, text }) => [level, text])).toEqual([
      [1, 'Heading 1'],
      [2, 'Heading 2'],
      [3, 'Heading 3'],
      [4, 'Heading 4'],
      [2, 'Custom chapter'],
    ])
  })

  it('updates cached pages without changing TOC or DATE field instructions through undo, redo and reopen', async () => {
    const options = { ...DEFAULT_TOC_OPTIONS, levels: 2 }
    const toc = generateTocFieldXml(
      [
        { level: 1, text: 'Chapter One', pageNo: 1 },
        { level: 2, text: 'Section Two', pageNo: 2 },
      ],
      options,
    )
    const source = await buildDocx({
      bodyXml:
        toc.join('') +
        date +
        heading('Heading1', 'Chapter One') +
        heading('Heading2', 'Section Two'),
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.push(editor)
    const region = findTocField(editor, parsed.blocks)
    expect(region?.instruction).toBe(' TOC \\o "1-2" \\h \\z \\u ')
    expect(updateTocPageCaches(editor, region!, ['4', '9'], parsed.blocks)).toBe(true)
    const changed = await documentXml(await save(editor, parsed))
    expect(changed).toContain('<w:t>4</w:t>')
    expect(changed).toContain('<w:t>9</w:t>')
    expect(changed).toContain('<w:fldSimple w:instr=" DATE ">')
    expect(changed).toContain(' TOC \\o "1-2" \\h \\z \\u ')

    expect(editor.commands.undo()).toBe(true)
    const undone = await documentXml(await save(editor, parsed))
    expect(undone).toContain('<w:t>1</w:t>')
    expect(undone).toContain('<w:t>2</w:t>')
    expect(undone).not.toContain('<w:t>9</w:t>')
    expect(editor.commands.redo()).toBe(true)
    const redone = await save(editor, parsed)
    expect(await documentXml(redone)).toContain('<w:t>9</w:t>')
    const reopened = await parseDocx(redone)
    expect(
      sameBodyAfterTocRebase(
        editor.state.doc,
        editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks)),
        reopened.blocks,
      ),
    ).toBe(true)
    expect(
      reopened.blocks
        .filter((block) => block.fieldDisplay?.kind === 'tocLine')
        .map((block) => block.fieldDisplay?.right),
    ).toEqual(['4', '9'])
    rebaseTocAnchors(editor, editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks)))
    expect(editor.commands.undo()).toBe(true)
    const savedUndo = await save(editor, reopened)
    const undoXml = await documentXml(savedUndo)
    expect(undoXml).toContain('<w:t>1</w:t>')
    expect(undoXml).not.toContain('<w:t>9</w:t>')
    expect(undoXml).toContain('<w:fldSimple w:instr=" DATE ">')
    const undoneParsed = await parseDocx(savedUndo)
    rebaseTocAnchors(editor, editor.schema.nodeFromJSON(blocksToPmDoc(undoneParsed.blocks)))
    expect(editor.commands.redo()).toBe(true)
    const redoXml = await documentXml(await save(editor, undoneParsed))
    expect(redoXml).toContain('<w:t>9</w:t>')
  })

  it('keeps an imported TOC content-control wrapper and unrelated body field byte-for-byte', async () => {
    const toc = generateTocFieldXml(
      [{ level: 1, text: 'Wrapped', pageNo: 3 }],
      DEFAULT_TOC_OPTIONS,
    )[0]!
    const source = await buildDocx({
      bodyXml:
        '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' +
        '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' +
        toc +
        '</w:sdtContent></w:sdt>' +
        date +
        heading('Heading1', 'Wrapped'),
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.push(editor)
    const region = findTocField(editor, parsed.blocks)
    expect(region).not.toBeNull()
    expect(updateTocPageCaches(editor, region!, ['7'], parsed.blocks)).toBe(true)
    const xml = await documentXml(await save(editor, parsed))
    expect((xml.match(/<w:sdt>/g) ?? []).length).toBe(1)
    expect((xml.match(/<\/w:sdt>/g) ?? []).length).toBe(1)
    expect(xml).toContain('<w:t>7</w:t>')
    expect(xml).toContain('<w:fldSimple w:instr=" DATE ">')
  })

  it('rebuilds entries inside an imported multi-paragraph content control without losing its title or shell', async () => {
    const original = generateTocFieldXml(
      [
        { level: 1, text: 'Old first', pageNo: 1 },
        { level: 2, text: 'Old second', pageNo: 2 },
      ],
      { ...DEFAULT_TOC_OPTIONS, levels: 2 },
    )
    const source = await buildDocx({
      bodyXml:
        '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' +
        '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' +
        '<w:p><w:r><w:t>My manual title</w:t></w:r></w:p>' +
        original.join('') +
        '</w:sdtContent></w:sdt>' +
        date,
    })
    const parsed = await parseDocx(source)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.push(editor)
    const region = findTocField(editor, parsed.blocks)!
    const entries = [
      { level: 1, text: 'New first', pageNo: 3 },
      { level: 2, text: 'New second', pageNo: 4 },
      { level: 2, text: 'New third', pageNo: 5 },
    ]
    const nodes = generateTocFieldXml(entries, { ...DEFAULT_TOC_OPTIONS, levels: 2 }).map(
      (xml, index) => ({
        type: 'docProtected',
        attrs: {
          docxIndex: null,
          blockType: 'passthrough',
          label: 'TOC line',
          genXml: xml,
          fieldDisplay: {
            kind: 'tocLine',
            left: entries[index]!.text,
            right: String(entries[index]!.pageNo),
            level: entries[index]!.level,
          },
        },
      }),
    )
    expect(anchorTocReplacement(nodes, region, parsed.blocks)).toBe(true)
    preserveTocHistorySource(editor, region, parsed.blocks)
    editor
      .chain()
      .deleteRange({ from: region.from, to: region.to })
      .insertContentAt(region.from, nodes as never)
      .run()
    const xml = await documentXml(await save(editor, parsed))
    expect((xml.match(/<w:sdt>/g) ?? []).length).toBe(1)
    expect((xml.match(/<\/w:sdt>/g) ?? []).length).toBe(1)
    expect(xml.indexOf('My manual title')).toBeLessThan(xml.indexOf('New first'))
    expect(xml.indexOf('New third')).toBeLessThan(xml.indexOf('</w:sdtContent>'))
    expect(xml).toContain('<w:fldSimple w:instr=" DATE ">')
    expect(editor.commands.undo()).toBe(true)
    const undone = await documentXml(await save(editor, parsed))
    expect(undone).toContain('Old first')
    expect(undone).not.toContain('New third')
    expect(undone).toContain('<w:fldSimple w:instr=" DATE ">')
    expect(editor.commands.redo()).toBe(true)
    const redone = await save(editor, parsed)
    expect(await documentXml(redone)).toContain('New third')
    const reopened = await parseDocx(redone)
    expect(
      sameBodyAfterTocRebase(
        editor.state.doc,
        editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks)),
        reopened.blocks,
      ),
    ).toBe(true)
    expect(
      reopened.blocks
        .filter((block) => block.fieldDisplay?.kind === 'tocLine')
        .map((block) => block.fieldDisplay?.left),
    ).toEqual(['New first', 'New second', 'New third'])
    rebaseTocAnchors(editor, editor.schema.nodeFromJSON(blocksToPmDoc(reopened.blocks)))
    expect(editor.commands.undo()).toBe(true)
    const crossSaveUndo = await save(editor, reopened)
    const crossSaveXml = await documentXml(crossSaveUndo)
    expect(crossSaveXml).toContain('Old first')
    expect(crossSaveXml).not.toContain('New third')
    expect((crossSaveXml.match(/<w:sdt>/g) ?? []).length).toBe(1)
    expect((crossSaveXml.match(/<\/w:sdt>/g) ?? []).length).toBe(1)
    expect(crossSaveXml).toContain('<w:fldSimple w:instr=" DATE ">')
    const undoneParsed = await parseDocx(crossSaveUndo)
    expect(
      sameBodyAfterTocRebase(
        editor.state.doc,
        editor.schema.nodeFromJSON(blocksToPmDoc(undoneParsed.blocks)),
        undoneParsed.blocks,
      ),
    ).toBe(true)
    rebaseTocAnchors(editor, editor.schema.nodeFromJSON(blocksToPmDoc(undoneParsed.blocks)))
    expect(editor.commands.redo()).toBe(true)
    const crossSaveRedo = await documentXml(await save(editor, undoneParsed))
    expect(crossSaveRedo).toContain('New third')
    expect(crossSaveRedo).toContain('<w:fldSimple w:instr=" DATE ">')
  })

  it('updates TOC in 碩士論文計畫書 範例.docx', async () => {
    const fs = await import('fs')
    const path = await import('path')
    const filePath = path.resolve(__dirname, '../../../碩士論文計畫書 範例.docx')
    if (!fs.existsSync(filePath)) return
    const buf = fs.readFileSync(filePath)
    const parsed = await parseDocx(buf)
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    editors.push(editor)
    const { computeListMarkerInfos } = await import('@genoffice/docx-engine')
    const refs: any[] = []
    const liNodes: any[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docListItem') {
        refs.push({ numId: node.attrs.numId, ilvl: Number(node.attrs.ilvl) || 0 })
        liNodes.push({ pos, node })
      }
    })
    const markers = computeListMarkerInfos(refs, parsed.numbering as any)
    liNodes.forEach((item, i) => {
      const styleId = item.node.attrs.styleId
      const info = parsed.styles.get(styleId)
      if (info?.headingLevel) {
        console.log(`Heading docListItem pos=${item.pos} marker="${markers[i]?.text}" style=${styleId} text="${item.node.textContent}"`)
      }
    })

    const region = findTocField(editor, parsed.blocks)
    expect(Boolean(region)).toBe(true)
    const options = tocOptionsFromField(region!.instruction, region!.nodes[0]?.xml ?? '')
    expect(Boolean(options)).toBe(true)
    ;(editor.storage as any).listNumbering = { defs: parsed.numbering }
    const actualEntries = collectTocSources(editor, parsed.blocks, parsed.styles, options!)
    expect(actualEntries.length).toBe(42)
    const ch1 = actualEntries.find((e) => e.text === '緒論')
    expect(ch1?.num).toBe('第1章')
    const s11 = actualEntries.find((e) => e.text === '研究動機')
    expect(s11?.num).toBe('1.1')
    const ch2 = actualEntries.find((e) => e.text === '相關研究探討')
    expect(ch2?.num).toBe('第2章')
    const ch3 = actualEntries.find((e) => e.text === '研究方法、步驟及執行速度')
    expect(ch3?.num).toBe('第3章')

    const baseEntryFieldDisplay = region!.nodes.find(
      ({ node }) => node.attrs.fieldDisplay?.kind === 'tocLine',
    )?.node.attrs.fieldDisplay
    expect(baseEntryFieldDisplay?.fontFamily).toBe('標楷體')

    const nodes = generateTocFieldXml(
      actualEntries,
      options!,
      region!.instruction,
      baseEntryFieldDisplay,
    ).map((xml, i) => ({
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: 'TOC Field',
        genXml: xml,
        fieldDisplay: {
          kind: 'tocLine',
          num: actualEntries[i]?.num,
          left: actualEntries[i]?.text ?? '',
          right: '',
          level: actualEntries[i]?.level ?? 1,
          leader: options?.leader ?? 'dot',
          fontFamily: baseEntryFieldDisplay?.fontFamily,
          fontAscii: baseEntryFieldDisplay?.fontAscii,
          eastAsiaFont: baseEntryFieldDisplay?.eastAsiaFont,
        },
      },
    }))
    expect(nodes.length).toBe(42)
    const anchored = anchorTocReplacement(nodes, region!, parsed.blocks)
    expect(anchored).toBe(true)

    preserveTocHistorySource(editor, region!, parsed.blocks)
    editor
      .chain()
      .focus()
      .deleteRange({ from: region!.from, to: region!.to })
      .insertContentAt(region!.from, nodes as never)
      .run()

    let tocLineCount = 0
    editor.state.doc.descendants((node) => {
      if (node.attrs.fieldDisplay?.kind === 'tocLine') {
        tocLineCount++
      }
    })
    expect(tocLineCount).toBe(62)

    const updatedRegion = findTocField(editor, parsed.blocks, region!.from)
    const updatedLines = updatedRegion!.nodes.filter(
      ({ node }) => node.attrs.fieldDisplay?.kind === 'tocLine',
    )
    expect(updatedLines.length).toBe(42)

    const savedBytes = await save(editor, parsed)
    const reopened = await parseDocx(savedBytes)
    const reopenedTocLines = reopened.blocks.filter((b) => b.fieldDisplay?.kind === 'tocLine')
    expect(reopenedTocLines.length).toBe(62)
    const reopenedCh1 = reopenedTocLines.find((b) => b.fieldDisplay?.left === '緒論')
    expect(reopenedCh1?.fieldDisplay?.num).toBe('第1章')
    expect(reopenedCh1?.fieldDisplay?.fontFamily).toBe('標楷體')
  })
})
