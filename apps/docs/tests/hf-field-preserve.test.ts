import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx, type HeaderFooter, type SaveBlock } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import {
  applyHfText,
  hfEditText,
  HfProtectedFieldEditError,
  insertHfNativeField,
} from '../src/renderer/editor/hf-text'
import { hfDisplayedRuns } from '../src/renderer/editor/hf-dom'

const DATE =
  '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true" w:fldLock="1"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> DA</w:instrText></w:r>' +
  '<w:r><w:instrText xml:space="preserve">TE \\@ "MMMM d" </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t>October</w:t></w:r>' +
  '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> 13</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const FILE =
  '<w:fldSimple w:instr=" FILENAME \\p " w:dirty="1"><w:r><w:t>report.docx</w:t></w:r></w:fldSimple>'
const EMPTY =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText> AUTHOR </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const PAGE =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText> PAGE </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
const TOTAL = '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>4</w:t></w:r></w:fldSimple>'

async function fixture(xml: string) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId62" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        xml,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
    ],
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId62"/>',
  })
  const parsed = await parseDocx(bytes)
  const blocks: SaveBlock[] = parsed.blocks
    .filter((block) => !block.hidden && block.docxIndex !== null)
    .map((block) => ({ kind: 'original', docxIndex: block.docxIndex! }))
  const value: HeaderFooter = {
    text: parsed.headerText ?? '',
    pageNumber: parsed.headerHasPageNumber,
    paras: parsed.headerParas ?? [],
  }
  return { parsed, blocks, value }
}

describe('header/footer protected fields', () => {
  it('authors a new DATE field as native XML and reopens it as an editable atom', async () => {
    const xml =
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Start End</w:t></w:r></w:p></w:hdr>'
    const { parsed, blocks, value } = await fixture(xml)
    const inserted = insertHfNativeField(
      value,
      hfEditText(value),
      6,
      6,
      'DATE',
      '',
      new Date(2026, 9, 13),
    )
    const saved = await saveDocx(parsed, blocks, { header: inserted.value })
    const savedXml = await (await JSZip.loadAsync(saved)).file('word/header1.xml')!.async('string')
    expect(savedXml).toContain(
      '<w:fldSimple w:instr=" DATE \\@ &quot;yyyy-MM-dd&quot; " w:dirty="true">',
    )
    expect(savedXml).toContain('<w:t xml:space="preserve">2026-10-13</w:t>')
    const reopened = await parseDocx(saved)
    const reopenedValue: HeaderFooter = {
      text: reopened.headerText ?? '',
      paras: reopened.headerParas ?? [],
    }
    expect(hfEditText(reopenedValue)).toContain('⟦DATE·')
    const reopenedDate = reopenedValue.paras
      ?.flatMap((para) => para.runs)
      .find((run) => run.hfField?.instruction.includes('DATE'))
    expect(hfDisplayedRuns(reopenedDate!, new Date(2026, 9, 14))[0].text).toBe('2026-10-14')
    const edited = applyHfText(reopenedValue, hfEditText(reopenedValue).replace('Start', 'Edited'))
    const savedAgain = await saveDocx(reopened, blocks, { header: edited })
    expect(
      await (await JSZip.loadAsync(savedAgain)).file('word/header1.xml')!.async('string'),
    ).toContain('<w:fldSimple w:instr=" DATE \\@ &quot;yyyy-MM-dd&quot; " w:dirty="true">')
  })
  it('reads core creator and keeps native AUTHOR/FILENAME cached XML while filename display follows Save As', async () => {
    const coreXml =
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:meta="http://purl.org/dc/elements/1.1/"><meta:creator>Ada &amp; Lin</meta:creator></cp:coreProperties>'
    const source = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'docProps/core.xml',
          xml: coreXml,
          contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
        },
      ],
    })
    const parsed = await parseDocx(source)
    expect(parsed.coreCreator).toBe('Ada & Lin')
    const authored = insertHfNativeField(
      null,
      '',
      0,
      0,
      'AUTHOR',
      'original.docx',
      new Date(2026, 9, 13),
      'yyyy-MM-dd',
      parsed.coreCreator,
    )
    const filename = insertHfNativeField(
      authored.value,
      authored.text,
      authored.text.length,
      authored.text.length,
      'FILENAME',
      'original.docx',
    )
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: filename.value,
    })
    const zip = await JSZip.loadAsync(saved)
    expect(await zip.file('docProps/core.xml')!.async('string')).toBe(coreXml)
    const headerXml = await zip.file('word/header1.xml')!.async('string')
    expect(headerXml).toContain('<w:fldSimple w:instr=" AUTHOR " w:dirty="true">')
    expect(headerXml).toContain('<w:t xml:space="preserve">Ada &amp; Lin</w:t>')
    expect(headerXml).toContain('<w:fldSimple w:instr=" FILENAME " w:dirty="true">')
    expect(headerXml).toContain('<w:t xml:space="preserve">original.docx</w:t>')
    const reopened = await parseDocx(saved)
    expect(reopened.coreCreator).toBe('Ada & Lin')
    const fields = reopened.headerParas?.flatMap((para) => para.runs.filter((run) => run.hfField))
    expect(fields).toHaveLength(2)
    expect(hfDisplayedRuns(fields![0], new Date(), { author: reopened.coreCreator })[0].text).toBe(
      'Ada & Lin',
    )
    expect(hfDisplayedRuns(fields![1], new Date(), { fileName: 'renamed.docx' })[0].text).toBe(
      'renamed.docx',
    )
    const savedAgain = await saveDocx(reopened, [{ kind: 'original', docxIndex: 0 }], {
      header: { text: reopened.headerText ?? '', paras: reopened.headerParas ?? [] },
    })
    expect(
      await (await JSZip.loadAsync(savedAgain)).file('word/header1.xml')!.async('string'),
    ).toContain('<w:t xml:space="preserve">original.docx</w:t>')
  })
  it('resolves a default Dublin Core namespace and decodes escaped creator text once', async () => {
    const coreXml =
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><creator xmlns="http://purl.org/dc/elements/1.1/">Ada &amp;amp; Lin</creator></cp:coreProperties>'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'docProps/core.xml',
          xml: coreXml,
          contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
        },
      ],
    })
    expect((await parseDocx(bytes)).coreCreator).toBe('Ada &amp; Lin')
  })
  it('preserves adjacent simple, complex, empty, PAGE and NUMPAGES fields through edits and reopen', async () => {
    const table =
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>table sibling</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    const picture = '<w:p><w:r><w:pict><v:shape id="logo"/></w:pict></w:r></w:p>'
    const xml = `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${DATE}${FILE}${EMPTY}<w:r><w:t xml:space="preserve"> after</w:t></w:r>${PAGE}${TOTAL}</w:p>${table}${picture}</w:hdr>`
    const { parsed, blocks, value } = await fixture(xml)
    const fields = value.paras?.flatMap((para) => para.runs.filter((run) => run.hfField)) ?? []
    expect(fields).toHaveLength(5)
    expect(fields.map((run) => run.hfField?.instruction.trim())).toEqual([
      'DATE \\@ "MMMM d"',
      'FILENAME \\p',
      'AUTHOR',
      'PAGE',
      'NUMPAGES',
    ])
    expect(hfDisplayedRuns(fields[1], new Date(), { fileName: 'renamed.docx' })[0].text).toBe(
      'report.docx',
    )
    expect(fields[0].hfField?.resultRuns.map((run) => [run.text, run.bold, run.italic])).toEqual([
      ['October', true, undefined],
      [' 13', undefined, true],
    ])
    const edit = hfEditText(value)
    expect(edit).toContain('Before ')
    expect(new Set(fields.map((run) => run.hfField?.token)).size).toBe(5)
    const next = applyHfText(
      value,
      edit.replace('Before ', 'Before edited\nnew line ').replace(' after', ' after edited'),
    )
    expect(
      next.paras
        ?.slice(0, 3)
        .map((para) => (para.cells ? 'table' : para.runs.map((run) => run.text).join(''))),
    ).toEqual(['Before edited', expect.stringContaining('new line '), 'table'])
    const saved = await saveDocx(parsed, blocks, { header: next })
    const savedXml = await (await JSZip.loadAsync(saved)).file('word/header1.xml')!.async('string')
    for (const fragment of [DATE, FILE, EMPTY, PAGE, TOTAL, table, picture]) {
      expect(savedXml, fragment).toContain(fragment)
      expect(savedXml.split(fragment)).toHaveLength(2)
    }
    const reopened = await parseDocx(saved)
    expect(
      reopened.headerParas?.flatMap((para) => para.runs.filter((run) => run.hfField)).length,
    ).toBe(5)
    expect(reopened.headerText).toContain('Before edited')
    expect(reopened.headerText).toContain('after edited')
    const joined = applyHfText(
      next,
      hfEditText(next).replace('Before edited\nnew line ', 'Before joined '),
    )
    expect(
      joined.paras
        ?.slice(0, 2)
        .map((para) => (para.cells ? 'table' : para.runs.map((run) => run.text).join(''))),
    ).toEqual([expect.stringContaining('Before joined '), 'table'])
    const joinedBytes = await saveDocx(parsed, blocks, { header: joined })
    const joinedXml = await (
      await JSZip.loadAsync(joinedBytes)
    )
      .file('word/header1.xml')!
      .async('string')
    for (const fragment of [DATE, FILE, EMPTY, PAGE, TOTAL])
      expect(joinedXml.split(fragment)).toHaveLength(2)
  })

  it('blocks changes to field tokens and malformed shared boundary runs without flattening', async () => {
    const xml = `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>before</w:t></w:r>${DATE}<w:r><w:t>after</w:t></w:r></w:p></w:hdr>`
    const { value } = await fixture(xml)
    expect(() => applyHfText(value, hfEditText(value).replace(/⟦DATE·\d+⟧/, 'changed'))).toThrow(
      HfProtectedFieldEditError,
    )

    const shared =
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>' +
      '<w:r><w:t>prefix</w:t><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText> DATE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>'
    const unsafe = await fixture(shared)
    expect(unsafe.value.paras?.[0].protectedFields).toBe(true)
    expect(() => applyHfText(unsafe.value, 'changed')).toThrow(HfProtectedFieldEditError)
    const bytes = await saveDocx(unsafe.parsed, unsafe.blocks)
    expect(await (await JSZip.loadAsync(bytes)).file('word/header1.xml')!.async('string')).toBe(
      shared,
    )
    const copied = await saveDocx(unsafe.parsed, unsafe.blocks, { header: unsafe.value })
    expect(await (await JSZip.loadAsync(copied)).file('word/header1.xml')!.async('string')).toBe(
      shared,
    )

    const separateAndCache =
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText> DATE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/><w:t>cached</w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="end"/><w:t>suffix</w:t></w:r></w:p></w:hdr>'
    const mixed = await fixture(separateAndCache)
    expect(mixed.value.paras?.[0].protectedFields).toBe(true)
    expect(() => applyHfText(mixed.value, 'changed')).toThrow(HfProtectedFieldEditError)
    const mixedSaved = await saveDocx(mixed.parsed, mixed.blocks)
    expect(
      await (await JSZip.loadAsync(mixedSaved)).file('word/header1.xml')!.async('string'),
    ).toBe(separateAndCache)
  })

  it('does not confuse literal private-use marker or visible token text with a field', async () => {
    const literal = '\uE100hf1\uE101 ⟦DATE·2⟧'
    const xml = `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>${literal}</w:t></w:r>${DATE}</w:p></w:hdr>`
    const { value } = await fixture(xml)
    const field = value.paras?.[0].runs.find((run) => run.hfField)?.hfField
    expect(field?.token).toBe('⟦DATE·3⟧')
    expect(hfEditText(value)).toContain(literal)
    const next = applyHfText(value, `${hfEditText(value)} more`)
    expect(next.paras?.[0].runs.filter((run) => run.hfField)).toHaveLength(1)
  })

  it('keeps fields enclosed by hyperlink and revision wrappers read-only and byte-exact', async () => {
    const wrappers = [
      `<w:hyperlink w:anchor="bookmark"><w:r><w:t>linked </w:t></w:r>${FILE}</w:hyperlink>`,
      `<w:ins w:id="3" w:author="Writer"><w:r><w:t>inserted </w:t></w:r>${DATE}</w:ins>`,
    ]
    for (const wrapped of wrappers) {
      const xml = `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>before </w:t></w:r>${wrapped}<w:r><w:t> after</w:t></w:r></w:p></w:hdr>`
      const { parsed, blocks, value } = await fixture(xml)
      expect(value.paras?.[0].protectedFields).toBe(true)
      expect(() => applyHfText(value, 'changed')).toThrow(HfProtectedFieldEditError)
      const bytes = await saveDocx(parsed, blocks, { header: value })
      expect(await (await JSZip.loadAsync(bytes)).file('word/header1.xml')!.async('string')).toBe(
        xml,
      )
    }
  })

  it('fails closed for a simple field containing paragraph boundaries', async () => {
    const xml =
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:fldSimple w:instr=" DATE "><w:p><w:r><w:t>first</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>second</w:t></w:r></w:p></w:fldSimple></w:hdr>'
    const { parsed, blocks, value } = await fixture(xml)
    expect(value.paras?.some((para) => para.protectedFields)).toBe(true)
    expect(() => applyHfText(value, 'edited')).toThrow(HfProtectedFieldEditError)
    const saved = await saveDocx(parsed, blocks, { header: value })
    expect(await (await JSZip.loadAsync(saved)).file('word/header1.xml')!.async('string')).toBe(xml)
  })

  it('keeps wrapped field parts byte-exact when only the body is edited', async () => {
    const wrapped = [
      `<w:p><w:hyperlink w:anchor="bookmark">${FILE}</w:hyperlink></w:p>`,
      `<w:sdt><w:sdtContent><w:p>${PAGE}</w:p></w:sdtContent></w:sdt>`,
    ]
    for (const content of wrapped) {
      const xml = `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${content}</w:hdr>`
      const { parsed, blocks, value } = await fixture(xml)
      expect(value.paras?.some((para) => para.protectedFields)).toBe(true)
      const bodyEdited: SaveBlock[] = blocks.map((block, index) =>
        index === 0
          ? {
              kind: 'xml',
              docxIndex: block.kind === 'original' ? block.docxIndex : undefined,
              xml: '<w:p><w:r><w:t>body edited</w:t></w:r></w:p>',
            }
          : block,
      )
      const saved = await saveDocx(parsed, bodyEdited)
      const zip = await JSZip.loadAsync(saved)
      expect(await zip.file('word/header1.xml')!.async('string')).toBe(xml)
      expect(await zip.file('word/document.xml')!.async('string')).toContain('body edited')
    }
  })

  it('keeps a nested result field and a DATE format string mentioning PAGE opaque', async () => {
    const nested =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText> DATE \\@ "PAGE MMMM" </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      '<w:r><w:t>Oct </w:t></w:r>' +
      PAGE +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    const xml = `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Start</w:t></w:r>${nested}<w:r><w:t>End</w:t></w:r></w:p></w:hdr>`
    const { parsed, blocks, value } = await fixture(xml)
    const fields = value.paras?.[0].runs.filter((run) => run.hfField) ?? []
    expect(fields).toHaveLength(1)
    expect(fields[0].hfField?.instruction).toContain('DATE')
    expect(fields[0].text).not.toContain('⟦')
    expect(value.pageNumber).toBe(false)
    const next = applyHfText(value, hfEditText(value).replace('Start', 'Edited'))
    const saved = await saveDocx(parsed, blocks, { header: next })
    const savedXml = await (await JSZip.loadAsync(saved)).file('word/header1.xml')!.async('string')
    expect(savedXml).toContain(nested)
  })
})
