import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  PENDING_IMG_PREFIX,
  parseDocx,
  pendingMediaOfDataUrl,
  registerPendingMedia,
  retargetDrawingToPending,
  saveDocx,
  type SaveBlock,
} from '../src/index'
import { mergeStyleXml } from '../src/style-upsert'
import { abstractNumXml } from '../src/blank'
import { TINY_PNG_BASE64, buildDocx } from './helpers/build-docx'

describe('cross-document style copy', () => {
  it('parse keeps the exact <w:style> element of every style', async () => {
    const parsed = await parseDocx(
      await buildDocx({
        bodyXml: '<w:p/>',
        extraStylesXml:
          '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Callout"><w:name w:val="Callout"/><w:rPr><w:color w:val="C00000"/></w:rPr></w:style>',
      }),
    )
    expect(parsed.extras.styleXml.Callout).toBe(
      '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Callout"><w:name w:val="Callout"/><w:rPr><w:color w:val="C00000"/></w:rPr></w:style>',
    )
    expect(Object.keys(parsed.extras.styleXml)).toContain('Heading1')
  })

  it('a rawXml upsert adds the element verbatim and never overwrites an existing style', async () => {
    const raw =
      '<w:style w:type="paragraph" w:styleId="Callout"><w:name w:val="Callout"/><w:pPr><w:shd w:val="clear" w:fill="FFF2CC"/></w:pPr></w:style>'
    expect(mergeStyleXml(null, { styleId: 'Callout', rawXml: raw })).toBe(raw)
    const existing =
      '<w:style w:type="paragraph" w:styleId="Callout"><w:name w:val="Callout"/><w:rPr><w:b/></w:rPr></w:style>'
    expect(mergeStyleXml(existing, { styleId: 'Callout', rawXml: raw })).toBe(existing)

    const parsed = await parseDocx(await buildDocx({ bodyXml: '<w:p/>' }))
    const blocks: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const saved = await saveDocx(parsed, blocks, {
      styleUpserts: [{ styleId: 'Callout', rawXml: raw }],
    })
    const styles = await (await JSZip.loadAsync(saved)).file('word/styles.xml')!.async('string')
    expect(styles).toContain(raw)
  })
})

describe('numbering level marker font', () => {
  it('writes the marker font, suffix and justification of a custom level', () => {
    const xml = abstractNumXml('9', 'bullet', [
      {
        numFmt: 'bullet',
        lvlText: '',
        indentLeft: 720,
        font: 'Symbol',
        suff: 'space',
        lvlJc: 'right',
      },
    ])
    expect(xml).toContain('<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>')
    expect(xml).toContain('<w:suff w:val="space"/>')
    expect(xml).toContain('<w:lvlJc w:val="right"/>')
    // schema order inside w:lvl: suff before lvlText, rPr after pPr
    expect(xml.indexOf('<w:suff')).toBeLessThan(xml.indexOf('<w:lvlText'))
    expect(xml.indexOf('<w:pPr>')).toBeLessThan(xml.indexOf('<w:rPr>'))
  })
})

describe('pictures copied from another package', () => {
  const DRAWING =
    '<w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="Picture 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'

  it('recognizes only formats the package stores as is', () => {
    expect(pendingMediaOfDataUrl(`data:image/png;base64,${TINY_PNG_BASE64}`)).toEqual({
      base64: TINY_PNG_BASE64,
      mime: 'image/png',
    })
    expect(pendingMediaOfDataUrl('data:image/webp;base64,AAAA')).toBeNull()
    expect(pendingMediaOfDataUrl('genoffice-docx-media://x/word/media/a.png')).toBeNull()
  })

  it('equal bytes share one token', () => {
    const media = { base64: TINY_PNG_BASE64, mime: 'image/png' as const }
    expect(registerPendingMedia(media)).toBe(registerPendingMedia({ ...media }))
    expect(registerPendingMedia(media).startsWith(PENDING_IMG_PREFIX)).toBe(true)
  })

  it('save lands the bytes, swaps the token for a fresh rId and renumbers docPr ids', async () => {
    const parsed = await parseDocx(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>B</w:t></w:r></w:p>' }),
    )
    const original: SaveBlock[] = parsed.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original', docxIndex: b.docxIndex! }))
    const media = { base64: TINY_PNG_BASE64, mime: 'image/png' as const }
    const xml = `<w:p><w:r>${retargetDrawingToPending(DRAWING, media)}</w:r></w:p>`
    const saved = await saveDocx(parsed, [...original, { kind: 'xml', xml }, { kind: 'xml', xml }])
    const zip = await JSZip.loadAsync(saved)
    const doc = await zip.file('word/document.xml')!.async('string')
    expect(doc).not.toContain(PENDING_IMG_PREFIX)
    expect(doc).not.toContain('rId7')
    const rid = /r:embed="(rId\d+)"/.exec(doc)![1]
    expect(doc.match(/r:embed=/g)).toHaveLength(2)
    expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).toContain(`Id="${rid}"`)
    expect(
      Object.keys(zip.files).filter((f) => /^word\/media\/aidocs\d+\.png$/.test(f)),
    ).toHaveLength(1)
    const ids = [...doc.matchAll(/<wp:docPr id="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(2)
  })
})
