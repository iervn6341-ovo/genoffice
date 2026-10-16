import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { PAGE_MARK, TOTAL_PAGES_MARK, parseDocx, readSections, saveDocx } from '../src/index'
import { buildDocx, TINY_PNG_BASE64 } from './helpers/build-docx'

const headerType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const body = '<w:p><w:r><w:t>Body</w:t></w:r></w:p>'
const midSection = (refs: string) =>
  `<w:p><w:pPr><w:sectPr>${refs}<w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>`
const header = (text: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>` +
  '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>KEEP_TABLE</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:hdr>'
const rel = (id: string, target: string) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${target}"/>`
const ref = (variant: string, id: string) => `<w:headerReference w:type="${variant}" r:id="${id}"/>`
const visibleBlocks = (parsed: Awaited<ReturnType<typeof parseDocx>>) =>
  parsed.blocks
    .filter((b) => !b.hidden && b.docxIndex !== null)
    .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))

async function text(bytes: Uint8Array, path: string) {
  return (await JSZip.loadAsync(bytes)).file(path)?.async('string')
}

describe('section header/footer editing', () => {
  it('clones a locally referenced part when another section explicitly shares its rId', async () => {
    const source = await buildDocx({
      bodyXml: body + midSection(ref('default', 'rId60')) + body,
      sectPrExtra: ref('default', 'rId60') + ref('even', 'rId61'),
      extraRels: rel('rId60', 'header1.xml') + rel('rId61', 'header2.xml'),
      extraParts: [
        { path: 'word/header1.xml', xml: header('Shared original'), contentType: headerType },
        { path: 'word/header2.xml', xml: header('Even original'), contentType: headerType },
      ],
    })
    const parsed = await parseDocx(source)
    const last = readSections(parsed).at(-1)!
    const saved = await saveDocx(parsed, visibleBlocks(parsed), {
      sectionHf: [
        {
          lastBlockIndex: last.lastBlockIndex,
          kind: 'header',
          variant: 'default',
          sourceRId: 'rId60',
          hf: { text: 'Only second section' },
        },
      ],
    })
    const sections = readSections(await parseDocx(saved))
    expect(sections[0].headerRefs.default).toBe('rId60')
    expect(sections[1].headerRefs.default).not.toBe('rId60')
    expect(sections[1].headerRefs.even).toBe('rId61')
    expect(await text(saved, 'word/header1.xml')).toBe(await text(source, 'word/header1.xml'))
    expect(await text(saved, 'word/header2.xml')).toBe(await text(source, 'word/header2.xml'))
    expect(await text(saved, 'word/header3.xml')).toContain('Only second section')
  })

  it('expands a self-closing section properties element before inserting a new reference', async () => {
    const source = await buildDocx({
      bodyXml: body + '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>' + body,
    })
    const parsed = await parseDocx(source)
    const first = readSections(parsed)[0]
    const saved = await saveDocx(parsed, visibleBlocks(parsed), {
      sectionHf: [
        {
          lastBlockIndex: first.lastBlockIndex,
          kind: 'header',
          hf: { text: 'First section only' },
        },
      ],
    })
    const document = (await text(saved, 'word/document.xml'))!
    expect(document).toMatch(/<w:sectPr><w:headerReference[^>]*\/><\/w:sectPr>/)
    expect(readSections(await parseDocx(saved))[0].headerRefs.default).toBeTruthy()
  })

  it('creates an independent inherited variant while retaining table, picture, and relationship', async () => {
    const source = await buildDocx({
      bodyXml: body + midSection(ref('first', 'rId60')) + body,
      extraRels: rel('rId60', 'header1.xml'),
      extraParts: [
        { path: 'word/header1.xml', xml: header('First original'), contentType: headerType },
        {
          path: 'word/_rels/header1.xml.rels',
          xml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/></Relationships>',
          contentType: 'application/vnd.openxmlformats-package.relationships+xml',
        },
      ],
      binaryParts: [
        {
          path: 'word/media/logo.png',
          base64: TINY_PNG_BASE64,
          extension: 'png',
          contentType: 'image/png',
        },
      ],
    })
    const parsed = await parseDocx(source)
    const sections = readSections(parsed)
    expect(sections).toHaveLength(2)
    const saved = await saveDocx(parsed, visibleBlocks(parsed), {
      sectionHf: [
        {
          lastBlockIndex: sections[1].lastBlockIndex,
          kind: 'header',
          variant: 'first',
          sourceRId: 'rId60',
          hf: { text: 'First edited', paras: [{ runs: [{ text: 'First edited' }] }] },
        },
      ],
    })
    const document = (await text(saved, 'word/document.xml'))!
    const refs = document.match(/<w:headerReference[^>]*w:type="first"[^>]*\/>/g) ?? []
    expect(refs).toHaveLength(2)
    expect(refs[0]).toContain('rId60')
    expect(refs[1]).not.toContain('rId60')
    expect(await text(saved, 'word/header1.xml')).toContain('First original')
    const second = (await text(saved, 'word/header2.xml'))!
    expect(second).toContain('First edited')
    expect(second).toContain('KEEP_TABLE')
    expect(second).toContain('r:embed="rId1"')
    expect(await text(saved, 'word/_rels/header2.xml.rels')).toContain('media/logo.png')
    expect((await readSections(await parseDocx(saved))).at(-1)?.headerRefs.first).toBeTruthy()
  })

  it('links only the selected variant to the previous section', async () => {
    const source = await buildDocx({
      bodyXml: body + midSection(ref('default', 'rId60') + ref('even', 'rId61')) + body,
      sectPrExtra: ref('default', 'rId62') + ref('even', 'rId63'),
      extraRels:
        rel('rId60', 'header1.xml') +
        rel('rId61', 'header2.xml') +
        rel('rId62', 'header3.xml') +
        rel('rId63', 'header4.xml'),
      extraParts: [1, 2, 3, 4].map((n) => ({
        path: `word/header${n}.xml`,
        xml: header(`H${n}`),
        contentType: headerType,
      })),
    })
    const parsed = await parseDocx(source)
    const sections = readSections(parsed)
    const saved = await saveDocx(parsed, visibleBlocks(parsed), {
      sectionHfLinks: [
        { lastBlockIndex: sections[1].lastBlockIndex, kind: 'header', variant: 'default' },
      ],
    })
    const result = readSections(await parseDocx(saved))
    expect(result[0].headerRefs.default).toBe('rId60')
    expect(result[1].headerRefs.default).toBeUndefined()
    expect(result[1].headerRefs.even).toBe('rId63')
    expect(await text(saved, 'word/document.xml')).toContain(ref('even', 'rId63'))
  })

  it('can undo and redo a saved local reference without rewriting its rich part', async () => {
    const source = await buildDocx({
      bodyXml: body + midSection(ref('default', 'rId60')) + body,
      extraRels: rel('rId60', 'header1.xml'),
      extraParts: [{ path: 'word/header1.xml', xml: header('Previous'), contentType: headerType }],
    })
    const original = await parseDocx(source)
    const finalIndex = readSections(original).at(-1)!.lastBlockIndex
    const edited = await saveDocx(original, visibleBlocks(original), {
      sectionHf: [
        {
          lastBlockIndex: finalIndex,
          kind: 'header',
          sourceRId: 'rId60',
          hf: { text: 'Independent' },
        },
      ],
    })
    const editedParsed = await parseDocx(edited)
    const localRef = readSections(editedParsed).at(-1)!.headerRefs.default!
    const localPart = (await text(edited, 'word/header2.xml'))!
    expect(localRef).toBeTruthy()
    const undone = await saveDocx(editedParsed, visibleBlocks(editedParsed), {
      sectionHfLinks: [{ lastBlockIndex: finalIndex, kind: 'header', variant: 'default' }],
    })
    const undoneParsed = await parseDocx(undone)
    expect(readSections(undoneParsed).at(-1)!.headerRefs.default).toBeUndefined()
    const redone = await saveDocx(undoneParsed, visibleBlocks(undoneParsed), {
      sectionHfRefs: [
        { lastBlockIndex: finalIndex, kind: 'header', variant: 'default', rId: localRef },
      ],
    })
    expect(readSections(await parseDocx(redone)).at(-1)!.headerRefs.default).toBe(localRef)
    expect(await text(redone, 'word/header2.xml')).toBe(localPart)
    expect(await text(redone, 'word/header1.xml')).toBe(await text(source, 'word/header1.xml'))
  })

  it('keeps a final-section first-page edit when an older link command is also pending', async () => {
    const source = await buildDocx({
      bodyXml: body + midSection(ref('first', 'rId60')) + body,
      sectPrExtra: ref('first', 'rId61'),
      extraRels: rel('rId60', 'header1.xml') + rel('rId61', 'header2.xml'),
      extraParts: [
        { path: 'word/header1.xml', xml: header('Previous first'), contentType: headerType },
        { path: 'word/header2.xml', xml: header('Final first'), contentType: headerType },
      ],
    })
    const parsed = await parseDocx(source)
    const last = readSections(parsed).at(-1)!
    const saved = await saveDocx(parsed, visibleBlocks(parsed), {
      sectionHfLinks: [{ lastBlockIndex: last.lastBlockIndex, kind: 'header', variant: 'first' }],
      sectionHf: [
        {
          lastBlockIndex: last.lastBlockIndex,
          kind: 'header',
          variant: 'first',
          hf: { text: 'Edited final first' },
        },
      ],
    })
    const result = readSections(await parseDocx(saved))
    expect(result[0].headerRefs.first).toBe('rId60')
    expect(result[1].headerRefs.first).not.toBe('rId61')
    expect(await text(saved, 'word/header2.xml')).toContain('Final first')
    expect(await text(saved, 'word/header3.xml')).toContain('Edited final first')
    expect(await text(saved, 'word/header1.xml')).toContain('Previous first')
  })

  it('writes PAGE and NUMPAGES fields from rich header runs', async () => {
    const source = await buildDocx({ bodyXml: body })
    const parsed = await parseDocx(source)
    const saved = await saveDocx(parsed, visibleBlocks(parsed), {
      header: {
        text: `Page ${PAGE_MARK} / ${TOTAL_PAGES_MARK}`,
        pageNumber: true,
        paras: [{ align: 'right', runs: [{ text: `Page ${PAGE_MARK} / ${TOTAL_PAGES_MARK}` }] }],
      },
    })
    const part = (await text(saved, 'word/header1.xml'))!
    expect(part).toContain(' PAGE </w:instrText>')
    expect(part).toContain(' NUMPAGES </w:instrText>')
    const reopened = await parseDocx(saved)
    expect(reopened.headerText).toBe(`Page ${PAGE_MARK} / ${TOTAL_PAGES_MARK}`)
    expect(reopened.headerParas?.[0].align).toBe('right')
  })
})
