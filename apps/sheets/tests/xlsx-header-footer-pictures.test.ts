import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import {
  assembleWithJsZip,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetPageSetupState } from '@genoffice/xlsx-gateway/gateway/xlsx-page-setup'

const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex')

async function save(source: Buffer, state: SheetPageSetupState): Promise<Buffer> {
  const plan = await planCellEditsToXlsx(
    await createBufferEntrySource(source),
    [],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [state],
  )
  return (await assembleWithJsZip(source, plan)).buffer
}

async function part(zip: JSZip, path: string): Promise<string> {
  return zip.file(path)!.async('string')
}

describe('native header/footer pictures', () => {
  it('preserves Excel Mac shared media and an orphan first-page shape while editing active slots', async () => {
    const source = readFileSync(
      new URL(
        '../../../docs/qa/office-mac-field-evidence-2026-10-13/native-wave-2/GenOfficeHeaderImageMacProbe.xlsx',
        import.meta.url,
      ),
    )
    const before = await JSZip.loadAsync(source)
    const originalVml = await part(before, 'xl/drawings/vmlDrawing1.vml')
    const originalRels = await part(before, 'xl/drawings/_rels/vmlDrawing1.vml.rels')
    const originalFirst = /<v:shape\b[^>]*\bid="LHFIRST"[\s\S]*?<\/v:shape>/.exec(originalVml)?.[0]
    expect(originalFirst).toContain('o:relid="rId1"')
    const image1 = await before.file('xl/media/image1.png')!.async('uint8array')
    const image2 = await before.file('xl/media/image2.jpeg')!.async('uint8array')

    const saved = await save(source, {
      sheetName: 'Sheet1',
      headerFooterPictureEdits: [
        {
          position: 'LH',
          widthPt: 120,
          heightPt: 25,
          image: { mediaType: 'image/png', base64: PNG.toString('base64') },
        },
        { position: 'LF', widthPt: 510, heightPt: 500 },
      ],
    })
    const after = await JSZip.loadAsync(saved)
    const vml = await part(after, 'xl/drawings/vmlDrawing1.vml')
    const rels = await part(after, 'xl/drawings/_rels/vmlDrawing1.vml.rels')
    expect(vml).toContain(originalFirst)
    expect(vml).toMatch(/id="LH"[\s\S]*?width:120pt;height:25pt[\s\S]*?o:relid="rId3"/)
    expect(vml).toMatch(/id="LF"[\s\S]*?width:510pt;height:500pt[\s\S]*?o:relid="rId2"/)
    expect(rels).toContain('Id="rId1"')
    expect(rels).toContain('Target="../media/image1.png"')
    expect(rels).toContain('Id="rId2"')
    expect(rels).toContain('Target="../media/image2.jpeg"')
    expect(rels).toContain('Id="rId3"')
    expect(rels).toContain('Target="../media/hfPicture1.png"')
    expect(originalRels).toContain('Id="rId1"')
    expect(await after.file('xl/media/image1.png')!.async('uint8array')).toEqual(image1)
    expect(await after.file('xl/media/image2.jpeg')!.async('uint8array')).toEqual(image2)
    expect(Buffer.from(await after.file('xl/media/hfPicture1.png')!.async('uint8array'))).toEqual(
      PNG,
    )
    expect(await part(after, 'xl/worksheets/sheet1.xml')).toContain(
      '<legacyDrawingHF r:id="rId1"/>',
    )
  })

  it('writes &G slots as worksheet → VML → image relationships and media, then resizes in place', async () => {
    const source = await blankXlsxBuffer('Sheet1')
    const added = await save(source, {
      sheetName: 'Sheet1',
      header: { left: '&G', center: 'Report' },
      firstPageStored: { header: null, footer: { right: '&G' } },
      firstPage: null,
      headerFooterPictureEdits: [
        {
          position: 'LH',
          widthPt: 120,
          heightPt: 24,
          image: { mediaType: 'image/png', base64: PNG.toString('base64') },
        },
        {
          position: 'RFFIRST',
          widthPt: 90,
          heightPt: 18,
          image: { mediaType: 'image/png', base64: PNG.toString('base64') },
        },
      ],
    })
    const zip = await JSZip.loadAsync(added)
    const sheet = await part(zip, 'xl/worksheets/sheet1.xml')
    expect(sheet).toContain('<legacyDrawingHF r:id="rId1"/>')
    expect(sheet).toContain('<oddHeader>&amp;L&amp;G&amp;CReport</oddHeader>')
    expect(sheet).toContain('<firstFooter>&amp;R&amp;G</firstFooter>')
    expect(sheet).not.toContain('differentFirst="1"')
    const sheetRels = await part(zip, 'xl/worksheets/_rels/sheet1.xml.rels')
    expect(sheetRels).toContain('relationships/vmlDrawing')
    const vml = await part(zip, 'xl/drawings/vmlDrawing1.vml')
    expect(vml).toContain('id="LH"')
    expect(vml).toContain('id="RFFIRST"')
    expect(vml).toContain('width:120pt;height:24pt')
    const vmlRels = await part(zip, 'xl/drawings/_rels/vmlDrawing1.vml.rels')
    expect(vmlRels).toContain('Target="../media/hfPicture1.png"')
    expect(vmlRels).toContain('Target="../media/hfPicture2.png"')
    expect(Buffer.from(await zip.file('xl/media/hfPicture1.png')!.async('uint8array'))).toEqual(PNG)
    expect(await part(zip, '[Content_Types].xml')).toContain('Extension="vml"')

    const resized = await save(added, {
      sheetName: 'Sheet1',
      headerFooterPictureEdits: [{ position: 'LH', widthPt: 140, heightPt: 28 }],
    })
    const next = await JSZip.loadAsync(resized)
    const nextVml = await part(next, 'xl/drawings/vmlDrawing1.vml')
    expect(nextVml).toContain('width:140pt;height:28pt')
    expect(nextVml).toContain('id="RFFIRST"')
    expect(await part(next, 'xl/drawings/_rels/vmlDrawing1.vml.rels')).toBe(vmlRels)
    expect(Buffer.from(await next.file('xl/media/hfPicture1.png')!.async('uint8array'))).toEqual(
      PNG,
    )
  })

  it('replaces one imported slot without touching foreign VML shapes or relationships', async () => {
    const source = await blankXlsxBuffer('Sheet1')
    const zip = await JSZip.loadAsync(source)
    let sheet = await part(zip, 'xl/worksheets/sheet1.xml')
    sheet = sheet.replace('</worksheet>', '<legacyDrawingHF r:id="rId7"/></worksheet>')
    sheet = sheet.replace(
      '<worksheet ',
      '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
    )
    zip.file('xl/worksheets/sheet1.xml', sheet)
    zip.file(
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing4.vml"/></Relationships>',
    )
    const foreign =
      '<v:shape id="Button1" custom="a>b" style="width:5pt;height:5pt"><x:ClientData ObjectType="Button"/></v:shape>'
    const original =
      '<v:shape id="LH" custom=\'a>b\' style=\'position:absolute;width:10pt;height:5pt;\n z-index:1\'><v:imagedata o:relid="rId1" o:title="old"/></v:shape>'
    zip.file(
      'xl/drawings/vmlDrawing4.vml',
      `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">${foreign}${original}</xml>`,
    )
    const originalRels =
      '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/original.png"/><Relationship Id="rId9" Type="custom" Target="../media/foreign.bin"/></Relationships>'
    zip.file('xl/drawings/_rels/vmlDrawing4.vml.rels', originalRels)
    zip.file('xl/media/original.png', PNG)
    zip.file('xl/media/foreign.bin', 'FOREIGN')
    const imported = await zip.generateAsync({ type: 'nodebuffer' })
    const saved = await save(imported, {
      sheetName: 'Sheet1',
      headerFooterPictureEdits: [
        {
          position: 'LH',
          widthPt: 22,
          heightPt: 11,
          image: { mediaType: 'image/png', base64: PNG.toString('base64') },
        },
      ],
    })
    const output = await JSZip.loadAsync(saved)
    const vml = await part(output, 'xl/drawings/vmlDrawing4.vml')
    expect(vml).toContain(foreign)
    expect(vml).toContain("custom='a>b'")
    expect(vml).toContain('o:title="old"')
    expect(vml).toContain('width:22pt;height:11pt')
    expect(vml.match(/\bstyle=/g) ?? []).toHaveLength(2)
    expect(await part(output, 'xl/drawings/_rels/vmlDrawing4.vml.rels')).toContain(
      '<Relationship Id="rId9" Type="custom" Target="../media/foreign.bin"/>',
    )
    expect(await part(output, 'xl/drawings/_rels/vmlDrawing4.vml.rels')).not.toContain(
      'Target="../media/original.png"',
    )
    expect(await part(output, 'xl/media/foreign.bin')).toBe('FOREIGN')
    expect(Buffer.from(await output.file('xl/media/original.png')!.async('uint8array'))).toEqual(
      PNG,
    )

    const removed = await save(saved, {
      sheetName: 'Sheet1',
      header: { left: '' },
      headerFooterPictureEdits: [{ position: 'LH', image: null }],
    })
    const cleared = await JSZip.loadAsync(removed)
    expect(await part(cleared, 'xl/drawings/vmlDrawing4.vml')).not.toContain('id="LH"')
    expect(await part(cleared, 'xl/drawings/vmlDrawing4.vml')).toContain(foreign)
    expect(await part(cleared, 'xl/drawings/_rels/vmlDrawing4.vml.rels')).toContain('Type="custom"')
  })
})
