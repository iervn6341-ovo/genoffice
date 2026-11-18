/** word/fonts/*.odttf: fontTable embed slots, ECMA-376 17.8.1 de-obfuscation, face listing. */
import { describe, expect, it } from 'vitest'
import {
  deobfuscateOdttf,
  isSfnt,
  isSfntNonBold,
  isSfntNonItalic,
  parseDocx,
  parseFontTable,
} from '../src/index'
import { buildDocx } from './helpers/build-docx'

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const KEY = '{AB6E25B2-8B6E-48E9-B1AF-D7B7CFCD6E9F}'
const KEY_BYTES = Uint8Array.from(
  KEY.replace(/[{}-]/g, '')
    .match(/../g)!
    .map((h) => parseInt(h, 16)),
)

/** TrueType magic + 12 fake table records, distinct byte per position */
function fakeSfnt(seed: number): Uint8Array {
  const b = new Uint8Array(64)
  for (let i = 0; i < b.length; i++) b[i] = (i * 7 + seed) & 0xff
  b.set([0x00, 0x01, 0x00, 0x00], 0)
  return b
}

function obfuscate(sfnt: Uint8Array): Uint8Array {
  const out = sfnt.slice()
  for (let i = 0; i < 32; i++) out[i] ^= KEY_BYTES[15 - (i % 16)]
  return out
}

const b64 = (b: Uint8Array) => Buffer.from(b).toString('base64')

describe('deobfuscateOdttf', () => {
  it('XORs the first 32 bytes with the reversed fontKey GUID', () => {
    const sfnt = fakeSfnt(3)
    const odttf = obfuscate(sfnt)
    expect(isSfnt(odttf)).toBe(false)
    expect(deobfuscateOdttf(odttf, KEY)).toEqual(sfnt)
    expect(odttf).not.toEqual(sfnt)
  })

  it('leaves unobfuscated parts and zero/missing keys alone', () => {
    const sfnt = fakeSfnt(5)
    expect(deobfuscateOdttf(sfnt, KEY)).toBe(sfnt)
    const scrambled = obfuscate(sfnt)
    expect(deobfuscateOdttf(scrambled, undefined)).toBe(scrambled)
    expect(deobfuscateOdttf(scrambled, '{00000000-0000-0000-0000-000000000000}')).toBe(scrambled)
  })

  it('accepts OTTO and true magics', () => {
    expect(isSfnt(Uint8Array.from([0x4f, 0x54, 0x54, 0x4f, ...new Array(12).fill(0)]))).toBe(true)
    expect(isSfnt(Uint8Array.from([0x74, 0x72, 0x75, 0x65, ...new Array(12).fill(0)]))).toBe(true)
    expect(isSfnt(new Uint8Array(4))).toBe(false)
  })
})

describe('fontTable embed slots', () => {
  const fontTableXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts ${W_NS} ${R_NS}>` +
    '<w:font w:name="Arial"><w:family w:val="swiss"/></w:font>' +
    `<w:font w:name="Fixture Sans"><w:family w:val="auto"/>` +
    `<w:embedRegular r:id="rId1" w:fontKey="${KEY}"/>` +
    `<w:embedBold r:id="rId2" w:fontKey="${KEY}"/>` +
    '<w:embedItalic r:id="rId3"/>' +
    `<w:embedBoldItalic r:id="rId4" w:fontKey="${KEY}"/>` +
    '</w:font>' +
    `<w:font w:name="Dangling"><w:embedRegular r:id="rId9" w:fontKey="${KEY}"/></w:font>` +
    '</w:fonts>'

  it('parses each slot with its rId and fontKey', () => {
    const entries = parseFontTable(fontTableXml)
    expect(entries[0].embedded).toBeUndefined()
    expect(entries[1].embedded).toEqual({
      regular: { rId: 'rId1', fontKey: KEY },
      bold: { rId: 'rId2', fontKey: KEY },
      italic: { rId: 'rId3' },
      boldItalic: { rId: 'rId4', fontKey: KEY },
    })
  })

  it('lists de-obfuscated faces from the package, skipping garbage and missing parts', async () => {
    const regular = fakeSfnt(1)
    const bold = fakeSfnt(2)
    const italic = fakeSfnt(3)
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font2.odttf"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="/word/fonts/font3.odttf"/>' +
      '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font4.odttf"/>' +
      '</Relationships>'
    const odttf = 'application/vnd.openxmlformats-officedocument.obfuscatedFont'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'word/fontTable.xml',
          xml: fontTableXml,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
        },
        {
          path: 'word/_rels/fontTable.xml.rels',
          xml: relsXml,
          contentType: 'application/vnd.openxmlformats-package.relationships+xml',
        },
      ],
      binaryParts: [
        {
          path: 'word/fonts/font1.odttf',
          base64: b64(obfuscate(regular)),
          extension: 'odttf',
          contentType: odttf,
        },
        {
          path: 'word/fonts/font2.odttf',
          base64: b64(obfuscate(bold)),
          extension: 'odttf',
          contentType: odttf,
        },
        {
          path: 'word/fonts/font3.odttf',
          base64: b64(italic),
          extension: 'odttf',
          contentType: odttf,
        },
        {
          path: 'word/fonts/font4.odttf',
          base64: b64(new Uint8Array(64).fill(0x42)),
          extension: 'odttf',
          contentType: odttf,
        },
      ],
    })
    const doc = await parseDocx(bytes)
    expect(doc.embeddedFonts?.map((f) => [f.family, f.bold, f.italic])).toEqual([
      ['Fixture Sans', false, false],
      ['Fixture Sans', true, false],
      ['Fixture Sans', false, true],
    ])
    expect(doc.embeddedFonts![0].data).toEqual(regular)
    expect(doc.embeddedFonts![1].data).toEqual(bold)
    expect(doc.embeddedFonts![2].data).toEqual(italic)
  })

  it('detects SFNT non-bold and non-italic weights from OS/2 table', () => {
    function makeSfntWithOs2(weight: number, isBold: boolean, isItalic: boolean): Uint8Array {
      const total = 12 + 16 + 80
      const b = new Uint8Array(total)
      const view = new DataView(b.buffer)
      view.setUint32(0, 0x00010000)
      view.setUint16(4, 1)
      view.setUint32(12, 0x4f532f32) // 'OS/2'
      view.setUint32(16, 0)
      view.setUint32(20, 28)
      view.setUint32(24, 80)
      view.setUint16(28 + 4, weight)
      view.setUint16(28 + 62, (isBold ? 0x0020 : 0) | (isItalic ? 0x0001 : 0x0040))
      return b
    }

    const reg = makeSfntWithOs2(400, false, false)
    const bold = makeSfntWithOs2(700, true, false)
    const italic = makeSfntWithOs2(400, false, true)

    expect(isSfntNonBold(reg)).toBe(true)
    expect(isSfntNonBold(bold)).toBe(false)
    expect(isSfntNonItalic(reg)).toBe(true)
    expect(isSfntNonItalic(italic)).toBe(false)
  })

  it('demotes a regular font file placed in an embedBold slot so faux bold is preserved', async () => {
    // Single-weight font (like 標楷體) where Word duplicates the 400 font into w:embedBold
    const total = 12 + 16 + 80
    const regularSfnt = new Uint8Array(total)
    const view = new DataView(regularSfnt.buffer)
    view.setUint32(0, 0x00010000)
    view.setUint16(4, 1)
    view.setUint32(12, 0x4f532f32)
    view.setUint32(16, 0)
    view.setUint32(20, 28)
    view.setUint32(24, 80)
    view.setUint16(28 + 4, 400) // regular weight
    view.setUint16(28 + 62, 0x0040) // regular selection

    const xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts ${W_NS} ${R_NS}>` +
      '<w:font w:name="標楷體"><w:altName w:val="DFKai-SB"/>' +
      `<w:embedRegular r:id="rId1"/>` +
      `<w:embedBold r:id="rId2"/>` +
      '</w:font></w:fonts>'
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font2.odttf"/>' +
      '</Relationships>'
    const odttf = 'application/vnd.openxmlformats-officedocument.obfuscatedFont'
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>',
      extraParts: [
        {
          path: 'word/fontTable.xml',
          xml,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
        },
        {
          path: 'word/_rels/fontTable.xml.rels',
          xml: relsXml,
          contentType: 'application/vnd.openxmlformats-package.relationships+xml',
        },
      ],
      binaryParts: [
        { path: 'word/fonts/font1.odttf', base64: b64(regularSfnt), extension: 'odttf', contentType: odttf },
        { path: 'word/fonts/font2.odttf', base64: b64(regularSfnt), extension: 'odttf', contentType: odttf },
      ],
    })
    const doc = await parseDocx(bytes)
    // Both entries should have bold: false so FontFace descriptor is 400, allowing synthetic bolding
    const bolds = doc.embeddedFonts?.map((f) => [f.family, f.bold])
    expect(bolds).toEqual([
      ['標楷體', false],
      ['DFKai-SB', false],
      ['標楷體', false],
      ['DFKai-SB', false],
    ])
  })

  it('omits embeddedFonts when the fontTable has no embed slots', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: '<w:p><w:r><w:t>x</w:t></w:r></w:p>' }))
    expect(doc.embeddedFonts).toBeUndefined()
  })
})
