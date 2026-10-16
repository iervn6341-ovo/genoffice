import { createHash } from 'node:crypto'
import { describe, expect, it } from 'vitest'
import { latexToOmml } from '@genoffice/office-math'

import {
  applyVisualEdits,
  VisualEditError,
} from '@genoffice/xlsx-gateway/gateway/xlsx-drawing-edit'
import { equationXml, type MutablePackage } from '@genoffice/xlsx-gateway/gateway/xlsx-drawing-add'

const ANCHOR = {
  fromRow: 2,
  fromColumn: 1,
  fromRowOffset: 0,
  fromColumnOffset: 9525,
  toRow: 12,
  toColumn: 7,
  toRowOffset: -9525,
  toColumnOffset: 0,
}

const marker = (prefix: 'from' | 'to', col: number, row: number): string =>
  `<xdr:${prefix}><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff>` +
  `<xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:${prefix}>`

// Document order: [0] chart graphicFrame, [1] picture, [2] one-cell shape.
const DRAWING =
  '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
  `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}` +
  '<xdr:graphicFrame macro=""><a:graphic><a:graphicData>' +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"/>' +
  '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>' +
  `<xdr:twoCellAnchor editAs="oneCell">${marker('from', 5, 1)}${marker('to', 9, 9)}` +
  '<xdr:pic><xdr:blipFill><a:blip r:embed="rId8"/></xdr:blipFill></xdr:pic>' +
  '<xdr:clientData/></xdr:twoCellAnchor>' +
  `<xdr:oneCellAnchor>${marker('from', 2, 20)}<xdr:ext cx="914400" cy="914400"/>` +
  '<xdr:sp><xdr:txBody/></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>' +
  '</xdr:wsDr>'

const DRAWING_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId7" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" ' +
  'Target="../charts/chart3.xml"/>' +
  '<Relationship Id="rId8" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
  'Target="../media/image1.png"/>' +
  '</Relationships>'

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
  '<Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
  '</Types>'

function fakePackage(entries: Map<string, string>): MutablePackage {
  return {
    paths: () => Promise.resolve([...entries.keys()]),
    has: (path) => Promise.resolve(entries.has(path)),
    readText: (path) => {
      const content = entries.get(path)
      if (content === undefined) return Promise.reject(new Error(`missing ${path}`))
      return Promise.resolve(content)
    },
    write: (path, content) => void entries.set(path, content),
    add: (path, content) => void entries.set(path, content),
    addBinary: () => undefined,
    remove: (path) => void entries.delete(path),
  }
}

const PATH = 'xl/drawings/drawing1.xml'

function packageWithDrawing(): Map<string, string> {
  return new Map([
    [PATH, DRAWING],
    ['xl/drawings/_rels/drawing1.xml.rels', DRAWING_RELS],
    ['xl/charts/chart3.xml', '<c:chartSpace/>'],
    ['xl/charts/_rels/chart3.xml.rels', '<Relationships/>'],
    ['xl/media/image1.png', 'png'],
    ['[Content_Types].xml', CONTENT_TYPES],
  ])
}

describe('applyVisualEdits', () => {
  it('replaces only the selected native equation choice and preserves unrelated anchors byte-for-byte', async () => {
    const before = equationXml(19, { latex: 'x+1', omml: latexToOmml('x+1') })
    const firstAnchor = `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}${before}<xdr:clientData/></xdr:twoCellAnchor>`
    const unrelated = `<xdr:twoCellAnchor>${marker('from', 5, 1)}${marker('to', 9, 9)}<xdr:sp name="keep"/><xdr:clientData/></xdr:twoCellAnchor>`
    const entries = new Map([[PATH, `<xdr:wsDr>${firstAnchor}${unrelated}</xdr:wsDr>`]])
    await applyVisualEdits(
      fakePackage(entries),
      [
        {
          drawingPath: PATH,
          drawingIndex: 0,
          equation: { latex: 'x+2', omml: latexToOmml('x+2') },
        },
      ],
      new Set(),
    )
    const after = entries.get(PATH)!
    expect(after).toContain(latexToOmml('x+2'))
    expect(after).not.toContain(latexToOmml('x+1'))
    expect(after).toContain(unrelated)
    expect(after).toContain('id="19"')
    expect(after).toContain(`GenOfficeEquationV2:${Buffer.from('x+2').toString('base64')}:`)
    await applyVisualEdits(
      fakePackage(entries),
      [
        {
          drawingPath: PATH,
          drawingIndex: 0,
          equation: { latex: 'x+3', omml: latexToOmml('x+3') },
        },
      ],
      new Set(),
    )
    expect(entries.get(PATH)).toContain(latexToOmml('x+3'))
    expect(entries.get(PATH)).not.toContain(latexToOmml('x+2'))
    expect(entries.get(PATH)).toContain(unrelated)
  })

  it('refuses to rewrite a non-equation anchor', async () => {
    await expect(
      applyVisualEdits(
        fakePackage(packageWithDrawing()),
        [
          {
            drawingPath: PATH,
            drawingIndex: 2,
            equation: { latex: 'x', omml: latexToOmml('x') },
          },
        ],
        new Set(),
      ),
    ).rejects.toThrow(/supported native equation/)
  })

  it('edits direct a14:m content without rebuilding its styled shape or ordinary text', async () => {
    const style =
      '<xdr:spPr><a:xfrm rot="5400000"><a:off x="42" y="51"/><a:ext cx="920000" cy="510000"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></xdr:spPr>'
    const direct = `<xdr:wsDr><xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="8" name="Finance equation" title="Quarterly" descr="Analyst note"/></xdr:nvSpPr>${style}<xdr:txBody><a:bodyPr vert="vert270"/><a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1"/><a:t>before </a:t></a:r><a14:m><m:oMath>${latexToOmml('x')}</m:oMath></a14:m><a:r><a:t> after</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`
    const entries = new Map([[PATH, direct]])
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, equation: { latex: 'y', omml: latexToOmml('y') } }],
      new Set(),
    )
    const after = entries.get(PATH) ?? ''
    expect(after).not.toContain('<mc:Choice')
    expect(after).toContain(latexToOmml('y'))
    expect(after).not.toContain(latexToOmml('x'))
    expect(after).toContain(style)
    expect(after).toContain('id="8" name="Finance equation" title="Quarterly" descr="Analyst note"')
    expect(after).toContain('<a:r><a:rPr b="1"/><a:t>before </a:t></a:r>')
    expect(after).toContain('<a:r><a:t> after</a:t></a:r>')
    expect(after).toContain('<a:bodyPr vert="vert270"/><a:p><a:pPr algn="ctr"/>')
  })

  it('keeps imported Choice/Fallback styling and adjacent text while changing matched fallback math text', async () => {
    const original = equationXml(31, { latex: 'q', omml: latexToOmml('q') })
    const custom = original
      .replaceAll(
        /<xdr:cNvPr id="31"[^>]*\/>/g,
        '<xdr:cNvPr id="31" name="Custom equation" title="Board report" descr="User description"/>',
      )
      .replaceAll(
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>',
        '<a:xfrm rot="600000"><a:off x="600" y="700"/><a:ext cx="800000" cy="300000"/></a:xfrm>',
      )
      .replaceAll(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>',
        '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln w="9144"><a:solidFill><a:srgbClr val="010203"/></a:solidFill></a:ln>',
      )
      .replaceAll(
        '<a:bodyPr/><a:lstStyle/><a:p>',
        '<a:bodyPr vert="vert270"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1300"/><a:t>before </a:t></a:r>',
      )
      .replaceAll('</a:p></xdr:txBody>', '<a:r><a:t> after</a:t></a:r></a:p></xdr:txBody>')
    const entries = new Map([
      [
        PATH,
        `<xdr:wsDr><xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}${custom}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      ],
    ])
    const before = entries.get(PATH) ?? ''
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, equation: { latex: 'y', omml: latexToOmml('y') } }],
      new Set(),
    )
    const after = entries.get(PATH) ?? ''
    expect(after).toContain(latexToOmml('y'))
    expect(after).not.toContain(latexToOmml('q'))
    expect(after).toContain('<a:t>before </a:t>')
    expect(after).toContain('<a:t>y</a:t>')
    expect(after).toContain('<a:t> after</a:t>')
    expect(
      after.match(/name="Custom equation" title="Board report" descr="User description"/g),
    ).toHaveLength(2)
    expect(after.match(/<a:xfrm rot="600000"><a:off x="600" y="700"\/>/g)).toHaveLength(2)
    expect(after.match(/<a:prstGeom prst="roundRect">/g)).toHaveLength(2)
    expect(after.match(/<a:solidFill><a:srgbClr val="ABCDEF"\/><\/a:solidFill>/g)).toHaveLength(2)
    expect(after.match(/<a:ln w="9144">/g)).toHaveLength(2)
    expect(after.match(/<a:bodyPr vert="vert270"\/>/g)).toHaveLength(2)
    expect(after.match(/<a:pPr algn="ctr"\/>/g)).toHaveLength(2)
    expect(
      after.replace(latexToOmml('y'), latexToOmml('q')).replace('<a:t>y</a:t>', '<a:t>q</a:t>'),
    ).toBe(before)
  })

  it('refuses an equation with a picture fallback so the old preview cannot survive an edit', async () => {
    const native = equationXml(33, { latex: 'x', omml: latexToOmml('x') }).replace(
      /<mc:Fallback>[\s\S]*?<\/mc:Fallback>/,
      '<mc:Fallback><xdr:pic><a:blip r:embed="rId9"/></xdr:pic></mc:Fallback>',
    )
    const entries = new Map([
      [
        PATH,
        `<xdr:wsDr><xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}${native}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      ],
    ])
    await expect(
      applyVisualEdits(
        fakePackage(entries),
        [{ drawingPath: PATH, drawingIndex: 0, equation: { latex: 'y', omml: latexToOmml('y') } }],
        new Set(),
      ),
    ).rejects.toThrow(/fallback cannot be updated/)
    expect(entries.get(PATH)).toContain(latexToOmml('x'))
  })

  it('uses native OMML when a matching-digest metadata source was tampered', async () => {
    const native = equationXml(34, { latex: 'x', omml: latexToOmml('x') }).replaceAll(
      `GenOfficeEquationV2:${Buffer.from('x').toString('base64')}:`,
      `GenOfficeEquationV2:${Buffer.from('y').toString('base64')}:`,
    )
    const entries = new Map([
      [
        PATH,
        `<xdr:wsDr><xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}${native}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      ],
    ])
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, equation: { latex: 'z', omml: latexToOmml('z') } }],
      new Set(),
    )
    expect(entries.get(PATH)).toContain(latexToOmml('z'))
    expect(entries.get(PATH)).toContain('<a:t>z</a:t>')
    expect(entries.get(PATH)).toContain(
      `GenOfficeEquationV2:${Buffer.from('z').toString('base64')}:`,
    )
  })

  it('rejects unsupported native OMML even when stored source and digest look current', async () => {
    const originalOmml = latexToOmml('x')
    const unsupportedOmml = '<m:unsupportedNativeMath/>'
    const native = equationXml(35, { latex: 'x', omml: originalOmml })
      .replace(originalOmml, unsupportedOmml)
      .replaceAll(
        createHash('sha256').update(originalOmml).digest('hex'),
        createHash('sha256').update(unsupportedOmml).digest('hex'),
      )
    const entries = new Map([
      [
        PATH,
        `<xdr:wsDr><xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}${native}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      ],
    ])
    await expect(
      applyVisualEdits(
        fakePackage(entries),
        [{ drawingPath: PATH, drawingIndex: 0, equation: { latex: 'z', omml: latexToOmml('z') } }],
        new Set(),
      ),
    ).rejects.toThrow(/cannot be reversed safely/)
    expect(entries.get(PATH)).toContain(unsupportedOmml)
  })

  it('rejects an anchor with multiple native math zones instead of editing only the first', async () => {
    const omml = latexToOmml('x')
    const math = `<a14:m><m:oMath>${omml}</m:oMath></a14:m>`
    const native = equationXml(36, { latex: 'x', omml }).replace(math, math + math)
    const entries = new Map([
      [
        PATH,
        `<xdr:wsDr><xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}${native}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      ],
    ])
    await expect(
      applyVisualEdits(
        fakePackage(entries),
        [{ drawingPath: PATH, drawingIndex: 0, equation: { latex: 'z', omml: latexToOmml('z') } }],
        new Set(),
      ),
    ).rejects.toThrow(/one supported native equation/)
    expect(entries.get(PATH)?.match(/<a14:m>/g)).toHaveLength(2)
  })

  it('preserves literal dollar sequences in replacement equation XML', async () => {
    const original = equationXml(25, { latex: 'x', omml: latexToOmml('x') })
    const entries = new Map([
      [
        PATH,
        `<xdr:wsDr><xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}${original}<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`,
      ],
    ])
    const latex = 'x$1'
    const omml = latexToOmml(latex)
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, equation: { latex, omml } }],
      new Set(),
    )
    expect(entries.get(PATH)).toContain(omml)
    expect(entries.get(PATH)).toContain('id="25"')
  })

  it('removes a picture anchor and leaves the others verbatim', async () => {
    const entries = packageWithDrawing()
    const touched = new Set<string>()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, remove: true }],
      touched,
    )
    const xml = entries.get(PATH)!
    expect(xml).not.toContain('<xdr:pic>')
    expect(xml).toContain('<xdr:graphicFrame')
    expect(xml).toContain('<xdr:oneCellAnchor>')
    expect(touched.has(PATH)).toBe(true)
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId8')
    expect(entries.has('xl/media/image1.png')).toBe(false)
  })

  it('keeps media shared by another anchor while removing only the unused relationship', async () => {
    const entries = packageWithDrawing()
    entries.set(
      PATH,
      DRAWING.replace(
        '</xdr:wsDr>',
        `<xdr:twoCellAnchor>${marker('from', 10, 1)}${marker('to', 14, 9)}` +
          '<xdr:pic><xdr:blipFill><a:blip r:embed="rId9"/></xdr:blipFill></xdr:pic>' +
          '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
      ),
    )
    entries.set(
      'xl/drawings/_rels/drawing1.xml.rels',
      DRAWING_RELS.replace(
        '</Relationships>',
        '<Relationship Id="rId9" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/></Relationships>',
      ),
    )

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, remove: true }],
      new Set(),
    )

    const rels = entries.get('xl/drawings/_rels/drawing1.xml.rels')!
    expect(rels).not.toContain('rId8')
    expect(rels).toContain('rId9')
    expect(entries.has('xl/media/image1.png')).toBe(true)
  })

  it('keeps media shared by a picture on another sheet drawing', async () => {
    const entries = packageWithDrawing()
    entries.set(
      'xl/drawings/drawing2.xml',
      '<xdr:wsDr><xdr:twoCellAnchor><xdr:pic><a:blip r:embed="rId1"/></xdr:pic>' +
        '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
    )
    entries.set(
      'xl/drawings/_rels/drawing2.xml.rels',
      '<Relationships><Relationship Id="rId1" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="../media/image1.png"/></Relationships>',
    )

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, remove: true }],
      new Set(),
    )

    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId8')
    expect(entries.has('xl/media/image1.png')).toBe(true)
  })

  it('moves a two-cell anchor by rewriting both markers', async () => {
    const entries = packageWithDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, anchor: ANCHOR }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain(
      '<xdr:from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff>' +
        '<xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>',
    )
    expect(xml).toContain(
      '<xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff>' +
        '<xdr:row>12</xdr:row><xdr:rowOff>-9525</xdr:rowOff></xdr:to>',
    )
    // The editAs attribute and the pic content survive verbatim.
    expect(xml).toContain('editAs="oneCell"')
    expect(xml).toContain('<xdr:pic>')
  })

  it('applies an edge resize that leaves the to marker unchanged', async () => {
    const entries = packageWithDrawing()
    // Picture anchor is from(5,1)→to(9,9); an NW resize moves only `from`.
    await applyVisualEdits(
      fakePackage(entries),
      [
        {
          drawingPath: PATH,
          drawingIndex: 1,
          anchor: {
            fromRow: 3,
            fromColumn: 6,
            fromRowOffset: 0,
            fromColumnOffset: 0,
            toRow: 9,
            toColumn: 9,
            toRowOffset: 0,
            toColumnOffset: 0,
          },
        },
      ],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain(marker('from', 6, 3))
    expect(xml).toContain(marker('to', 9, 9))
  })

  it('rewrites the true frame ext when a resize sends frameSize', async () => {
    const rotated =
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
      `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 4, 8)}` +
      '<xdr:sp><xdr:spPr><a:xfrm rot="2700000"><a:off x="10" y="20"/>' +
      '<a:ext cx="100" cy="50"/></a:xfrm></xdr:spPr></xdr:sp>' +
      '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'
    const entries = new Map([[PATH, rotated]])
    await applyVisualEdits(
      fakePackage(entries),
      [
        {
          drawingPath: PATH,
          drawingIndex: 0,
          anchor: ANCHOR,
          frameSize: { width: 555, height: 333 },
        },
      ],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain('<a:ext cx="555" cy="333"/>')
    // Rotation and offset survive verbatim; only the ext is rewritten.
    expect(xml).toContain('<a:xfrm rot="2700000"><a:off x="10" y="20"/>')
  })

  it('fails closed when frameSize targets an anchor without a frame ext', async () => {
    await expect(
      applyVisualEdits(
        fakePackage(packageWithDrawing()),
        [
          {
            drawingPath: PATH,
            drawingIndex: 1,
            anchor: ANCHOR,
            frameSize: { width: 1, height: 1 },
          },
        ],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
  })

  it('moves a one-cell anchor by rewriting only its from marker', async () => {
    const entries = packageWithDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 2, anchor: ANCHOR }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain('<xdr:ext cx="914400" cy="914400"/>')
    expect(xml).not.toContain(`${marker('from', 2, 20)}<xdr:ext`)
  })

  it('processes several edits on one part high-index first', async () => {
    const entries = packageWithDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [
        { drawingPath: PATH, drawingIndex: 1, remove: true },
        { drawingPath: PATH, drawingIndex: 2, remove: true },
      ],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).not.toContain('<xdr:pic>')
    expect(xml).not.toContain('<xdr:oneCellAnchor>')
    expect(xml).toContain('<xdr:graphicFrame')
  })

  it('deleting a chart cascades its rel, part, own rels, and override', async () => {
    const entries = packageWithDrawing()
    const touched = new Set<string>()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      touched,
    )
    expect(entries.get(PATH)).not.toContain('<xdr:graphicFrame')
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId7')
    // The unrelated image relationship survives verbatim.
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).toContain('rId8')
    expect(entries.has('xl/charts/chart3.xml')).toBe(false)
    expect(entries.has('xl/charts/_rels/chart3.xml.rels')).toBe(false)
    expect(entries.get('[Content_Types].xml')).not.toContain('chart3.xml')
    expect(entries.get('[Content_Types].xml')).toContain('drawing1.xml')
    expect(touched.has('[Content_Types].xml')).toBe(true)
  })

  it('recursively collects chart-owned style, color, embedding, theme, and media parts', async () => {
    const entries = packageWithDrawing()
    entries.set(
      'xl/charts/_rels/chart3.xml.rels',
      '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style3.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors3.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/source3.xlsx"/>' +
        '<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2011/relationships/themeOverride" Target="../theme/themeOverride3.xml"/>' +
        '</Relationships>',
    )
    entries.set('xl/charts/style3.xml', '<cs:chartStyle/>')
    entries.set(
      'xl/charts/_rels/style3.xml.rels',
      '<Relationships><Relationship Id="rId1" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="../media/chart-style.png"/></Relationships>',
    )
    entries.set('xl/charts/colors3.xml', '<cs:colorStyle/>')
    entries.set('xl/embeddings/source3.xlsx', 'embedded')
    entries.set('xl/theme/themeOverride3.xml', '<a:themeOverride/>')
    entries.set('xl/media/chart-style.png', 'png')
    entries.set(
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '</Types>',
        '<Override PartName="/xl/charts/style3.xml" ContentType="chart-style"/>' +
          '<Override PartName="/xl/charts/colors3.xml" ContentType="chart-colors"/>' +
          '<Override PartName="/xl/theme/themeOverride3.xml" ContentType="theme-override"/>' +
          '</Types>',
      ),
    )

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    for (const path of [
      'xl/charts/chart3.xml',
      'xl/charts/_rels/chart3.xml.rels',
      'xl/charts/style3.xml',
      'xl/charts/_rels/style3.xml.rels',
      'xl/charts/colors3.xml',
      'xl/embeddings/source3.xlsx',
      'xl/theme/themeOverride3.xml',
      'xl/media/chart-style.png',
    ]) {
      expect(entries.has(path), path).toBe(false)
    }
    const contentTypes = entries.get('[Content_Types].xml')!
    expect(contentTypes).not.toContain('style3.xml')
    expect(contentTypes).not.toContain('colors3.xml')
    expect(contentTypes).not.toContain('themeOverride3.xml')
  })

  it('retains chart-owned dependencies referenced by another package part', async () => {
    const entries = packageWithDrawing()
    entries.set(
      'xl/charts/_rels/chart3.xml.rels',
      '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style3.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors3.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/source3.xlsx"/>' +
        '</Relationships>',
    )
    entries.set('xl/charts/style3.xml', '<cs:chartStyle/>')
    entries.set(
      'xl/charts/_rels/style3.xml.rels',
      '<Relationships><Relationship Id="rId1" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
        'Target="../media/chart-style.png"/></Relationships>',
    )
    entries.set('xl/charts/colors3.xml', '<cs:colorStyle/>')
    entries.set('xl/embeddings/source3.xlsx', 'embedded')
    entries.set('xl/media/chart-style.png', 'png')
    entries.set('xl/charts/chart4.xml', '<c:chartSpace/>')
    entries.set(
      'xl/charts/_rels/chart4.xml.rels',
      '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style3.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors3.xml"/>' +
        '</Relationships>',
    )

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    expect(entries.has('xl/charts/chart3.xml')).toBe(false)
    expect(entries.has('xl/embeddings/source3.xlsx')).toBe(false)
    expect(entries.has('xl/charts/style3.xml')).toBe(true)
    expect(entries.has('xl/charts/_rels/style3.xml.rels')).toBe(true)
    expect(entries.has('xl/charts/colors3.xml')).toBe(true)
    expect(entries.has('xl/media/chart-style.png')).toBe(true)
  })

  it('fails closed and preserves unsupported internal chart relationships', async () => {
    const entries = packageWithDrawing()
    entries.set(
      'xl/charts/_rels/chart3.xml.rels',
      '<Relationships><Relationship Id="rId1" Type="urn:vendor:chart-extension" ' +
        'Target="../custom/chart-extension.xml"/></Relationships>',
    )
    entries.set('xl/custom/chart-extension.xml', '<vendor:extension/>')
    const before = new Map(entries)

    await expect(
      applyVisualEdits(
        fakePackage(entries),
        [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)

    expect(entries).toEqual(before)
  })

  it('keeps a chart target while another anchor still references its relationship', async () => {
    const entries = packageWithDrawing()
    const duplicated = DRAWING.replace(
      '</xdr:wsDr>',
      `<xdr:twoCellAnchor>${marker('from', 0, 30)}${marker('to', 4, 38)}` +
        '<xdr:graphicFrame><a:graphic><a:graphicData>' +
        '<c:chart xmlns:c="c" xmlns:r="r" r:id="rId7"/>' +
        '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>' +
        '</xdr:wsDr>',
    )
    entries.set(PATH, duplicated)
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).toContain('rId7')
    expect(entries.has('xl/charts/chart3.xml')).toBe(true)
  })

  it('cleans the final picture drawing hookup only when it is empty and unambiguous', async () => {
    const entries = new Map([
      [
        PATH,
        '<xdr:wsDr>' +
          `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 2, 2)}` +
          '<xdr:pic><a:blip r:embed="rId8"/></xdr:pic>' +
          '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
      ],
      [
        'xl/drawings/_rels/drawing1.xml.rels',
        '<Relationships><Relationship Id="rId8" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/></Relationships>',
      ],
      ['xl/media/image1.png', 'png'],
      ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/><drawing r:id="rId5"/></worksheet>'],
      [
        'xl/worksheets/_rels/sheet1.xml.rels',
        '<Relationships><Relationship Id="rId5" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ' +
          'Target="../drawings/drawing1.xml"/></Relationships>',
      ],
      ['[Content_Types].xml', CONTENT_TYPES],
    ])

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    expect(entries.has(PATH)).toBe(false)
    expect(entries.has('xl/drawings/_rels/drawing1.xml.rels')).toBe(false)
    expect(entries.has('xl/media/image1.png')).toBe(false)
    expect(entries.get('xl/worksheets/sheet1.xml')).not.toContain('<drawing')
    expect(entries.has('xl/worksheets/_rels/sheet1.xml.rels')).toBe(false)
    expect(entries.get('[Content_Types].xml')).not.toContain('drawing1.xml')
  })

  it('preserves an empty drawing hookup that still has an unsupported relationship', async () => {
    const entries = new Map([
      [
        PATH,
        '<xdr:wsDr>' +
          `<xdr:twoCellAnchor>${marker('from', 0, 0)}${marker('to', 2, 2)}` +
          '<xdr:pic><a:blip r:embed="rId8"/></xdr:pic>' +
          '<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>',
      ],
      [
        'xl/drawings/_rels/drawing1.xml.rels',
        '<Relationships><Relationship Id="rId8" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/><Relationship Id="rId99" ' +
          'Type="urn:vendor:unsupported" Target="../custom/vendor.xml"/></Relationships>',
      ],
      ['xl/media/image1.png', 'png'],
      ['xl/custom/vendor.xml', 'custom'],
      ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/><drawing r:id="rId5"/></worksheet>'],
      [
        'xl/worksheets/_rels/sheet1.xml.rels',
        '<Relationships><Relationship Id="rId5" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ' +
          'Target="../drawings/drawing1.xml"/></Relationships>',
      ],
      ['[Content_Types].xml', CONTENT_TYPES],
    ])

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    expect(entries.has(PATH)).toBe(true)
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).toContain('rId99')
    expect(entries.has('xl/custom/vendor.xml')).toBe(true)
    expect(entries.get('xl/worksheets/sheet1.xml')).toContain('<drawing')
  })

  it('fails closed on non-chart graphic frames, absolute anchors, and bad indexes', async () => {
    const frame = new Map(packageWithDrawing())
    frame.set(PATH, DRAWING.replace(/<c:chart\b[^>]*\/>/, '<a:tbl/>'))
    await expect(
      applyVisualEdits(
        fakePackage(frame),
        [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
    const absolute = new Map([
      [
        PATH,
        '<xdr:wsDr><xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:sp/><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>',
      ],
    ])
    await expect(
      applyVisualEdits(
        fakePackage(absolute),
        [{ drawingPath: PATH, drawingIndex: 0, anchor: ANCHOR }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
    await expect(
      applyVisualEdits(
        fakePackage(packageWithDrawing()),
        [{ drawingPath: PATH, drawingIndex: 9, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
    await expect(
      applyVisualEdits(
        fakePackage(new Map()),
        [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
    await expect(
      applyVisualEdits(
        fakePackage(packageWithDrawing()),
        [
          { drawingPath: PATH, drawingIndex: 1, remove: true },
          { drawingPath: PATH, drawingIndex: 1, anchor: ANCHOR },
        ],
        new Set(),
      ),
    ).rejects.toThrow(VisualEditError)
  })
})

// openpyxl writes the spreadsheetDrawing namespace as the DEFAULT namespace:
// anchors, markers, pic and graphicFrame all carry no prefix. The sidecar
// counts anchors by local name, so edits must find and rewrite these too.
const DEFAULT_NS_DRAWING =
  '<wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
  '<oneCellAnchor><from><col>3</col><colOff>0</colOff><row>1</row><rowOff>0</rowOff></from>' +
  '<ext cx="5400000" cy="2700000"/><graphicFrame><a:graphic><a:graphicData>' +
  '<c:chart r:id="rId7"/></a:graphicData></a:graphic></graphicFrame>' +
  '<clientData/></oneCellAnchor>' +
  '<twoCellAnchor><from><col>1</col><colOff>0</colOff><row>2</row><rowOff>0</rowOff></from>' +
  '<to><col>4</col><colOff>0</colOff><row>8</row><rowOff>0</rowOff></to>' +
  '<pic><blipFill><a:blip r:embed="rId8"/></blipFill></pic>' +
  '<clientData/></twoCellAnchor>' +
  '</wsDr>'

function packageWithDefaultNsDrawing(): Map<string, string> {
  return new Map([
    [PATH, DEFAULT_NS_DRAWING],
    ['xl/drawings/_rels/drawing1.xml.rels', DRAWING_RELS],
    ['xl/charts/chart3.xml', '<c:chartSpace/>'],
    ['xl/charts/_rels/chart3.xml.rels', '<Relationships/>'],
    ['xl/media/image1.png', 'png'],
    ['[Content_Types].xml', CONTENT_TYPES],
  ])
}

describe('applyVisualEdits on default-namespace (openpyxl) drawings', () => {
  it('moves an unprefixed two-cell anchor, keeping its markers unprefixed', async () => {
    const entries = packageWithDefaultNsDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, anchor: ANCHOR }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain(
      '<from><col>1</col><colOff>9525</colOff><row>2</row><rowOff>0</rowOff></from>',
    )
    expect(xml).toContain(
      '<to><col>7</col><colOff>0</colOff><row>12</row><rowOff>-9525</rowOff></to>',
    )
    expect(xml).not.toContain('xdr:')
  })

  it('moves an unprefixed one-cell anchor by its from marker only', async () => {
    const entries = packageWithDefaultNsDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, anchor: ANCHOR }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain(
      '<from><col>1</col><colOff>9525</colOff><row>2</row><rowOff>0</rowOff></from>',
    )
    expect(xml).toContain('<ext cx="5400000" cy="2700000"/>')
  })

  it('removes an unprefixed picture anchor and cascades its image relationship', async () => {
    const entries = packageWithDefaultNsDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 1, remove: true }],
      new Set(),
    )
    expect(entries.get(PATH)!).not.toContain('<pic>')
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId8')
    expect(entries.has('xl/media/image1.png')).toBe(false)
  })

  it('removes an unprefixed chart graphic frame and cascades the chart part', async () => {
    const entries = packageWithDefaultNsDrawing()
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )
    expect(entries.get(PATH)!).not.toContain('<graphicFrame>')
    expect(entries.get('xl/drawings/_rels/drawing1.xml.rels')).not.toContain('rId7')
    expect(entries.has('xl/charts/chart3.xml')).toBe(false)
  })

  it('keeps anchor indexes aligned across mixed prefixed and unprefixed anchors', async () => {
    const entries = packageWithDefaultNsDrawing()
    entries.set(
      PATH,
      DEFAULT_NS_DRAWING.replace(
        '<oneCellAnchor>',
        '<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
          `${marker('from', 0, 0)}${marker('to', 2, 2)}<xdr:sp/><xdr:clientData/>` +
          '</xdr:twoCellAnchor><oneCellAnchor>',
      ),
    )
    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 2, remove: true }],
      new Set(),
    )
    const xml = entries.get(PATH)!
    expect(xml).toContain('<xdr:sp/>')
    expect(xml).toContain('<graphicFrame>')
    expect(xml).not.toContain('<pic>')
  })

  it('cleans the final default-namespace drawing hookup once it is empty', async () => {
    const entries = new Map([
      [
        PATH,
        '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          '<twoCellAnchor><from><col>0</col><colOff>0</colOff><row>0</row><rowOff>0</rowOff></from>' +
          '<to><col>2</col><colOff>0</colOff><row>2</row><rowOff>0</rowOff></to>' +
          '<pic><a:blip r:embed="rId8"/></pic><clientData/></twoCellAnchor></wsDr>',
      ],
      [
        'xl/drawings/_rels/drawing1.xml.rels',
        '<Relationships><Relationship Id="rId8" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
          'Target="../media/image1.png"/></Relationships>',
      ],
      ['xl/media/image1.png', 'png'],
      ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/><drawing r:id="rId5"/></worksheet>'],
      [
        'xl/worksheets/_rels/sheet1.xml.rels',
        '<Relationships><Relationship Id="rId5" ' +
          'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" ' +
          'Target="../drawings/drawing1.xml"/></Relationships>',
      ],
      ['[Content_Types].xml', CONTENT_TYPES],
    ])

    await applyVisualEdits(
      fakePackage(entries),
      [{ drawingPath: PATH, drawingIndex: 0, remove: true }],
      new Set(),
    )

    expect(entries.has(PATH)).toBe(false)
    expect(entries.get('xl/worksheets/sheet1.xml')).not.toContain('<drawing')
  })
})
