/**
 * Docs → Docs clipboard through the editor's own copy handler and ProseMirror's
 * paste pipeline, with a second document as destination (different style
 * sheet, numbering, media, notes). Every case saves the destination and checks
 * the docx XML, not just the editor model.
 */
import { describe, expect, it } from 'vitest'
import {
  buildAnchoredTextboxParagraphXml,
  generateTocFieldXml,
} from '../../../packages/docx-engine/src/generate'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
  NESTED_TABLE_XML,
} from '../../../packages/docx-engine/tests/helpers/build-docx'
import { TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import {
  blocks,
  caretAtEnd,
  copy,
  openDocx,
  paste,
  selectAll,
  selectNode,
  selectText,
  type OpenedDoc,
} from './helpers/clip-harness'

const P = (inner: string, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${inner}</w:p>`
const R = (t: string, rpr = '') =>
  `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`

const STYLES_EXTRA =
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Callout"><w:name w:val="Callout"/><w:basedOn w:val="Normal"/><w:pPr><w:pBdr><w:left w:val="single" w:sz="24" w:color="FF0000"/></w:pBdr><w:shd w:val="clear" w:fill="FFF2CC"/></w:pPr><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:color w:val="C00000"/><w:sz w:val="26"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:customStyle="1" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:b/><w:color w:val="00B050"/></w:rPr></w:style>'
const NUMBERING =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%1)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>'

const TBL =
  '<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/>' +
  '<w:tblBorders><w:top w:val="single" w:sz="12" w:color="FF0000"/><w:left w:val="single" w:sz="12" w:color="FF0000"/><w:bottom w:val="single" w:sz="12" w:color="FF0000"/><w:right w:val="single" w:sz="12" w:color="FF0000"/><w:insideH w:val="dashed" w:sz="4" w:color="0000FF"/><w:insideV w:val="dashed" w:sz="4" w:color="0000FF"/></w:tblBorders></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="3600"/><w:gridCol w:w="1800"/></w:tblGrid>' +
  '<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr>' +
  '<w:tc><w:tcPr><w:tcW w:w="5400" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Merged AB</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/><w:vMerge w:val="restart"/><w:shd w:val="clear" w:color="auto" w:fill="00FF00"/></w:tcPr><w:p><w:r><w:t>VM</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/><w:tcBorders><w:bottom w:val="double" w:sz="6" w:color="00AA00"/></w:tcBorders><w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>A2</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr><w:p><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="28"/><w:color w:val="800080"/></w:rPr><w:t>B2 courier</w:t></w:r></w:p></w:tc>' +
  '<w:tc><w:tcPr><w:tcW w:w="1800" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>'

const DRAW =
  '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1" descr="alt text here"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId10"/><a:srcRect l="10000" t="5000"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
const FN =
  '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>'
const FOOTNOTES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>footnote body text</w:t></w:r></w:p></w:footnote></w:footnotes>'
const SECT =
  '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:pPr></w:p>'

const RICH_BODY = [
  P(R('Heading text'), '<w:pStyle w:val="Heading1"/>'),
  P(R('plain ') + R('bold', '<w:b/>') + R(' emph', '<w:rStyle w:val="Emph"/>')),
  P(
    R('centered indent spacing'),
    '<w:pStyle w:val="Callout"/><w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:left="720" w:hanging="360"/><w:jc w:val="center"/>',
  ),
  P(R('num item 1'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'),
  P(R('num item 2 lvl1'), '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>'),
  P(R('bullet item'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'),
  P('<w:bookmarkStart w:id="7" w:name="mark1"/>' + R('bookmarked') + '<w:bookmarkEnd w:id="7"/>'),
  P('<w:r><w:br w:type="page"/></w:r>'),
  P(R('after page break')),
].join('')

const FOOTNOTE_PARTS = {
  extraRels:
    '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
  extraParts: [
    {
      path: 'word/footnotes.xml',
      xml: FOOTNOTES_XML,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
    },
  ],
}

const opened: OpenedDoc[] = []
async function open(bodyXml: string, extra: Record<string, unknown> = {}): Promise<OpenedDoc> {
  const doc = await openDocx(await buildDocx({ bodyXml, ...extra } as never))
  opened.push(doc)
  return doc
}
const blank = () => open(P(R('B doc')))
/** copy everything of `src`, paste at the end of `dst` */
function transfer(src: OpenedDoc, dst: OpenedDoc) {
  selectAll(src.editor)
  const clip = copy(src.editor)
  caretAtEnd(dst.editor)
  paste(dst.editor, clip)
  return clip
}
const pasted = (dst: OpenedDoc) => blocks(dst.editor).slice(1)

describe('tables (copy → another document)', () => {
  it('keeps borders, fills, spans, widths, alignment and row height', async () => {
    const src = await open(TBL + '<w:p><w:r><w:t>after</w:t></w:r></w:p>')
    const dst = await blank()
    transfer(src, dst)
    const table = pasted(dst).find((b) => b.type === 'docTable')!
    expect(table.attrs?.borders).toMatchObject({
      top: { style: 'single', color: 'FF0000' },
      insideH: { style: 'dashed', color: '0000FF' },
    })
    const rows = table.content!
    expect(rows[0].attrs?.heightTwips).toBe(600)
    const [merged, vm] = rows[0].content!
    expect(merged.attrs).toMatchObject({ colspan: 2, fill: 'FFFF00' })
    expect(vm.attrs).toMatchObject({ rowspan: 2, fill: '00FF00' })
    const a2 = rows[1].content![0]
    expect(a2.attrs).toMatchObject({ vAlign: 'center', align: 'right' })
    expect(a2.attrs?.borders).toMatchObject({ bottom: { style: 'double', color: '00AA00' } })
    expect(merged.attrs?.colwidth).toEqual([120, 240])
  })

  it('saves the pasted table with borders, shading, spans and merges in the docx', async () => {
    const src = await open(TBL + '<w:p><w:r><w:t>after</w:t></w:r></w:p>')
    const dst = await blank()
    transfer(src, dst)
    const { documentXml } = await dst.save()
    expect(documentXml).toContain('w:gridSpan w:val="2"')
    expect(documentXml).toContain('w:fill="FFFF00"')
    expect(documentXml).toContain('w:fill="00FF00"')
    expect(documentXml).toContain('<w:vMerge w:val="restart"/>')
    expect(documentXml).toMatch(/w:color="FF0000"/)
    expect(documentXml).toMatch(/w:val="dashed"/)
    expect(documentXml).toMatch(/w:val="double"/)
    expect(documentXml).toContain('<w:trHeight')
  })

  it('keeps a nested table', async () => {
    const src = await open(NESTED_TABLE_XML + '<w:p/>')
    const dst = await blank()
    transfer(src, dst)
    const { documentXml } = await dst.save()
    expect(documentXml.match(/<w:tbl>/g)).toHaveLength(2)
    expect(documentXml).toContain('InnerA')
  })

  it('a one-cell table stays a table (not unwrapped to text)', async () => {
    const src = await open(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p>' +
        R('only cell') +
        '</w:p></w:tc></w:tr></w:tbl><w:p/>',
    )
    const dst = await blank()
    transfer(src, dst)
    expect(pasted(dst).some((b) => b.type === 'docTable')).toBe(true)
  })

  it('plain text of a table is tab-separated rows (Excel)', async () => {
    const src = await open(TBL + '<w:p/>')
    selectNode(src.editor, 0)
    const clip = copy(src.editor)
    const lines = clip.text.split('\n')
    expect(lines[0]).toBe('Merged AB\tVM')
    expect(lines[1]).toBe('A2\tB2 courier')
  })
})

describe('paragraphs, styles and lists (copy → another document)', () => {
  const srcExtra = { extraStylesXml: STYLES_EXTRA, numberingXml: NUMBERING }

  it('imports a style the destination lacks and saves its definition', async () => {
    const src = await open(RICH_BODY, srcExtra)
    const dst = await blank()
    transfer(src, dst)
    expect(dst.imported.styles.map((s) => s.id).sort()).toEqual(['Callout', 'Emph'])
    const { stylesXml, documentXml } = await dst.save()
    expect(stylesXml).toContain('w:styleId="Callout"')
    expect(stylesXml).toContain('w:fill="FFF2CC"')
    expect(stylesXml).toContain('w:styleId="Emph"')
    expect(documentXml).toContain('<w:pStyle w:val="Callout"/>')
    expect(documentXml).toContain('<w:rStyle w:val="Emph"/>')
  })

  it('keeps direct paragraph formatting (indent, spacing, alignment)', async () => {
    const src = await open(RICH_BODY, srcExtra)
    const dst = await blank()
    transfer(src, dst)
    const { documentXml } = await dst.save()
    expect(documentXml).toContain(
      '<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>',
    )
    expect(documentXml).toContain('<w:ind w:left="720" w:hanging="360"/>')
    expect(documentXml).toContain('<w:jc w:val="center"/>')
  })

  it('uses the destination definition when the style name already exists (Word rule)', async () => {
    const src = await open(RICH_BODY, srcExtra)
    const dst = await open(P(R('B')), {
      extraStylesXml:
        '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Callout"><w:name w:val="Callout"/><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>',
    })
    transfer(src, dst)
    expect(dst.imported.styles.map((s) => s.id)).toEqual(['Emph'])
    const { stylesXml } = await dst.save()
    expect(stylesXml).toContain('w:val="0000FF"')
    expect(stylesXml).not.toContain('FFF2CC')
  })

  it('imports the list definitions and renumbers the items', async () => {
    const src = await open(RICH_BODY, srcExtra)
    const dst = await blank()
    transfer(src, dst)
    const { documentXml, numberingXml } = await dst.save()
    expect(numberingXml).toContain('w:val="lowerLetter"')
    expect(numberingXml).toContain('(%1)')
    expect(numberingXml).toContain('w:ascii="Symbol"')
    const used = [...documentXml.matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1])
    expect(used.length).toBe(3)
    for (const id of used) expect(numberingXml).toContain(`<w:num w:numId="${id}">`)
  })

  it('keeps an in-paragraph page break', async () => {
    const src = await open(RICH_BODY, srcExtra)
    const dst = await blank()
    transfer(src, dst)
    const { documentXml } = await dst.save()
    expect(documentXml).toContain('<w:br w:type="page"/>')
  })

  it('a pasted paragraph is new content: no anchors, bookmarks or docx identity', async () => {
    const src = await open(RICH_BODY, srcExtra)
    const dst = await blank()
    transfer(src, dst)
    for (const b of pasted(dst)) {
      expect(b.attrs?.docxIndex ?? null).toBeNull()
      expect(b.attrs?.bookmarks ?? null).toBeNull()
    }
  })

  it('pasting into the same document imports nothing and keeps numbering', async () => {
    const src = await open(RICH_BODY, srcExtra)
    selectAll(src.editor)
    const clip = copy(src.editor)
    caretAtEnd(src.editor)
    paste(src.editor, clip)
    expect(src.imported.styles).toEqual([])
    expect(src.imported.numbering).toEqual([])
    const items = blocks(src.editor).filter((b) => b.type === 'docListItem')
    expect(new Set(items.map((i) => i.attrs?.numId))).toEqual(new Set(['1', '2']))
  })

  it('undo restores the destination in one step', async () => {
    const src = await open(RICH_BODY, srcExtra)
    const dst = await blank()
    const before = JSON.stringify(dst.editor.getJSON())
    transfer(src, dst)
    expect(JSON.stringify(dst.editor.getJSON())).not.toBe(before)
    dst.editor.commands.undo()
    expect(JSON.stringify(dst.editor.getJSON())).toBe(before)
  })

  it('cut puts the content on the clipboard and removes it', async () => {
    const src = await open(P(R('cut me')) + P(R('keep')))
    selectText(src.editor, 1, 7)
    const clip = copy(src.editor, true)
    expect(clip.text).toBe('cut me')
    expect(src.editor.getText()).not.toContain('cut me')
  })
})

describe('pictures, notes, boxes and breaks (copy → another document)', () => {
  it('a picture in running text lands in the destination package with a fresh rId', async () => {
    const src = await open(P(R('text before ') + DRAW + R(' text after')), { withImage: true })
    const dst = await blank()
    transfer(src, dst)
    const { documentXml, zip } = await dst.save()
    expect(documentXml).not.toContain('r:embed="rId10"')
    expect(documentXml).not.toContain('aidocs-pending-img')
    const rid = /r:embed="(rId\d+)"/.exec(documentXml)![1]
    const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
    expect(rels).toContain(`Id="${rid}"`)
    expect(Object.keys(zip.files).some((f) => /^word\/media\/aidocs\d+\.png$/.test(f))).toBe(true)
    expect(documentXml).toContain('descr="alt text here"')
    expect(documentXml).toContain('<a:srcRect l="10000" t="5000"/>')
  })

  it('a picture in a table cell travels with the table', async () => {
    const src = await open(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p>' +
        R('cell') +
        DRAW +
        '</w:p></w:tc></w:tr></w:tbl><w:p/>',
      { withImage: true },
    )
    const dst = await blank()
    transfer(src, dst)
    const { documentXml, zip } = await dst.save()
    expect(documentXml).toContain('<w:tbl>')
    expect(documentXml).not.toContain('r:embed="rId10"')
    expect(Object.keys(zip.files).some((f) => /^word\/media\/aidocs\d+\.png$/.test(f))).toBe(true)
  })

  it('the same picture pasted twice shares one media part', async () => {
    const src = await open(P(DRAW), { withImage: true })
    const dst = await blank()
    transfer(src, dst)
    selectAll(src.editor)
    const clip = copy(src.editor)
    caretAtEnd(dst.editor)
    paste(dst.editor, clip)
    const { zip, documentXml } = await dst.save()
    const media = Object.keys(zip.files).filter((f) => /^word\/media\/aidocs\d+\.png$/.test(f))
    expect(media).toHaveLength(1)
    const ids = [...documentXml.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('a block picture keeps its size', async () => {
    const src = await open(IMAGE_PARAGRAPH_XML, { withImage: true })
    const dst = await blank()
    transfer(src, dst)
    const { documentXml, zip } = await dst.save()
    expect(documentXml).toContain('<wp:extent cx="914400" cy="914400"/>')
    expect(Object.keys(zip.files).some((f) => /^word\/media\/aidocs\d+\.png$/.test(f))).toBe(true)
  })

  it('a footnote reference gets its own note in the destination', async () => {
    const src = await open(P(R('body') + FN + R(' more')), FOOTNOTE_PARTS)
    const dst = await blank()
    transfer(src, dst)
    expect(dst.imported.notes).toEqual(['footnote:50:footnote body text'])
    const { documentXml } = await dst.save()
    expect(documentXml).toContain('<w:footnoteReference w:id="50"/>')
  })

  it('a text box is rebuilt in the destination', async () => {
    const xml = buildAnchoredTextboxParagraphXml({
      anchor: 'paragraph',
      xEmu: 91440,
      yEmu: 12700,
      widthEmu: 2486400,
      heightEmu: 771600,
      fillHex: 'FFEEAA',
      insetsEmu: { l: 91440, t: 91440, r: 91440, b: 91440 },
      zOrder: 7,
      id: 42,
      paragraphs: [{ runs: [{ text: 'Textbox content', bold: true }], format: {} }],
      holderLineTwips: 20,
    })
    const src = await open(xml)
    const dst = await blank()
    transfer(src, dst)
    const { documentXml } = await dst.save()
    expect(documentXml).toContain('Textbox content')
    expect(documentXml).toContain('FFEEAA')
  })

  it('a section break keeps its page setup', async () => {
    const src = await open(P(R('before')) + SECT + P(R('after')))
    const dst = await blank()
    transfer(src, dst)
    const { documentXml } = await dst.save()
    expect(documentXml).toContain('<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>')
  })

  it('a table of contents field keeps its begin / separate / end structure', async () => {
    const toc = generateTocFieldXml([
      { level: 1, text: 'Intro', pageNo: 1 },
      { level: 2, text: 'Sub', pageNo: 2 },
    ]).join('')
    const src = await open(toc + '<w:p/>')
    const dst = await blank()
    transfer(src, dst)
    const { documentXml } = await dst.save()
    expect(documentXml).toContain('w:fldCharType="begin"')
    expect(documentXml).toContain('w:fldCharType="separate"')
    expect(documentXml).toContain('w:fldCharType="end"')
    expect(documentXml).toContain('TOC \\o')
    expect(documentXml).toContain('Intro')
    expect(documentXml).toContain('Sub')
  })

  it("a copied cell range pastes into another table's cells with its fills", async () => {
    const cell = (t: string, extra = '') =>
      `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${extra}</w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`
    const table = (rows: string[][], fills: string[] = []) =>
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows
        .map(
          (r, i) =>
            `<w:tr>${r
              .map((c, j) =>
                cell(
                  c,
                  fills[i * 2 + j] ? `<w:shd w:val="clear" w:fill="${fills[i * 2 + j]}"/>` : '',
                ),
              )
              .join('')}</w:tr>`,
        )
        .join('')}</w:tbl><w:p/>`
    const src = await open(
      table(
        [
          ['a1', 'a2'],
          ['b1', 'b2'],
        ],
        ['FF0000', '', '', '00FF00'],
      ),
    )
    const dst = await open(
      table([
        ['x1', 'x2'],
        ['y1', 'y2'],
      ]),
    )
    const posOf = (doc: OpenedDoc, text: string) => {
      let found = -1
      doc.editor.state.doc.descendants((n, pos) => {
        if (
          found < 0 &&
          n.type.name.startsWith('docTable') &&
          n.textContent === text &&
          n.type.name !== 'docTable' &&
          n.type.name !== 'docTableRow'
        )
          found = pos
        return true
      })
      return found
    }
    const s = src.editor
    s.view.dispatch(
      s.state.tr.setSelection(
        CellSelection.create(s.state.doc, posOf(src, 'a1'), posOf(src, 'b2')),
      ),
    )
    const clip = copy(s)
    expect(clip.text).toBe('a1\ta2\nb1\tb2')
    const d = dst.editor
    d.view.dispatch(
      d.state.tr.setSelection(TextSelection.create(d.state.doc, posOf(dst, 'x1') + 2)),
    )
    paste(d, clip)
    const cells = blocks(d)[0]!.content!.map((r) =>
      r.content!.map((c) => [c.attrs?.fill ?? null, c.content![0].content![0].text]),
    )
    expect(cells).toEqual([
      [
        ['FF0000', 'a1'],
        [null, 'a2'],
      ],
      [
        [null, 'b1'],
        ['00FF00', 'b2'],
      ],
    ])
  })
})
