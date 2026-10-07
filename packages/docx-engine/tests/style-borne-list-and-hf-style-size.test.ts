/**
 * Word probes 2026-10-01 (contract sample, Word for Mac):
 *  - a list whose numbering comes from the paragraph style chain lays out at
 *    the style's own w:ind (ECM L2 left=1390 over level left=1106), while the
 *    same paragraph with a direct w:numPr sits at the level's indent — the
 *    parser has to tell the two apart (list.fromStyle);
 *  - a header/footer paragraph styled Footer (10pt under a 12pt Normal) lays
 *    out its unsized runs at 10pt (styleSizeHalfPoints).
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

const NUMBERING =
  `${XML_DECL}<w:numbering ${W}>` +
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/>' +
  '<w:pPr><w:ind w:left="1106" w:hanging="539"/></w:pPr></w:lvl>' +
  '</w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>'

const STYLES =
  `${XML_DECL}<w:styles ${W}>` +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="L1"><w:name w:val="L1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="L2"><w:name w:val="L2"/><w:basedOn w:val="L1"/>' +
  '<w:pPr><w:numPr><w:ilvl w:val="1"/></w:numPr><w:ind w:left="1390"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/>' +
  '<w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
  '</w:styles>'

describe('style-borne list numbering', () => {
  it('flags numbering that comes from the pStyle chain, not a direct numPr', async () => {
    const bytes = await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:pStyle w:val="L2"/></w:pPr><w:r><w:t>from style</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:pStyle w:val="L2"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr></w:pPr>' +
        '<w:r><w:t>direct</w:t></w:r></w:p>',
      stylesXml: STYLES,
      numberingXml: NUMBERING,
    })
    const items = (await parseDocx(bytes)).blocks.filter((b) => b.type === 'listItem')
    expect(items.map((b) => b.list)).toEqual([
      { kind: 'ordered', numId: '2', ilvl: 1, fromStyle: true },
      { kind: 'ordered', numId: '2', ilvl: 1 },
    ])
  })
})

describe('header/footer paragraph style size', () => {
  it('carries the paragraph style run size for unsized text runs', async () => {
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
      stylesXml: STYLES,
      extraRels:
        '<Relationship Id="rId60" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml:
            `${XML_DECL}<w:hdr ${W}>` +
            '<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:t>2</w:t></w:r></w:p>' +
            '<w:p><w:r><w:t>plain</w:t></w:r></w:p></w:hdr>',
          contentType: HEADER_TYPE,
        },
      ],
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId60"/>',
    })
    const [styled, plain] = (await parseDocx(bytes)).headerParas!
    expect(styled.styleSizeHalfPoints).toBe(20)
    expect(plain.styleSizeHalfPoints).toBeUndefined()
  })
})
