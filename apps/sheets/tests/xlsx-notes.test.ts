import { describe, expect, it } from 'vitest'
import { applySheetNotes } from '@genoffice/xlsx-gateway/gateway/xlsx-notes'

/** In-memory package with the file set an Excel-saved sheet with two notes carries */
function excelPackage() {
  const files = new Map<string, string>([
    [
      '[Content_Types].xml',
      '<Types><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>',
    ],
    ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/><legacyDrawing r:id="rId2"/></worksheet>'],
    [
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
        '</Relationships>',
    ],
    [
      'xl/comments1.xml',
      '<comments><authors><author>Ann</author><author>Bo &amp; Co</author></authors><commentList>' +
        // A1: Excel's usual rich note — bold author line, then red text
        '<comment ref="A1" authorId="0"><text><r><rPr><b/></rPr><t>Ann:</t></r>' +
        '<r><t xml:space="preserve">\n</t></r><r><rPr><color rgb="FFFF0000"/></rPr><t>Check &lt;this&gt;</t></r></text></comment>' +
        '<comment ref="B1" authorId="1"><text><t>old text</t></text></comment>' +
        '</commentList></comments>',
    ],
    [
      'xl/drawings/vmlDrawing1.vml',
      '<xml><v:shapetype id="_x0000_t202"/>' +
        // A1's box was made visible and resized in Excel
        '<v:shape id="_x0000_s1025" type="#_x0000_t202" style="width:300pt;height:120pt;visibility:visible">' +
        '<x:ClientData ObjectType="Note"><x:Anchor>1,0,0,0,6,0,8,0</x:Anchor><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>' +
        '<v:shape id="_x0000_s1026" type="#_x0000_t202" style="width:50pt">' +
        '<x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>' +
        '<v:shape id="_x0000_s1027" type="#_x0000_t201"><x:ClientData ObjectType="Checkbox"/></v:shape>' +
        '</xml>',
    ],
  ])
  return {
    files,
    pkg: {
      paths: async () => [...files.keys()],
      has: async (path: string) => files.has(path),
      readText: async (path: string) => files.get(path) ?? '',
      write: (path: string, content: string) => void files.set(path, content),
      add: (path: string, content: string) => void files.set(path, content),
      remove: (path: string) => void files.delete(path),
    },
  }
}

describe('applySheetNotes keeps the notes an edit did not touch', () => {
  it('edits B1, adds C3, and leaves A1 rich and its box untouched', async () => {
    const { files, pkg } = excelPackage()
    await applySheetNotes(
      pkg,
      'xl/worksheets/sheet1.xml',
      [
        // A1 exactly as the importer read it (all <t> runs, author by id)
        { row: 0, column: 0, author: 'Ann', text: 'Ann:\nCheck <this>' },
        { row: 0, column: 1, author: 'Bo & Co', text: 'new text' },
        { row: 2, column: 2, author: 'Cy', text: 'fresh' },
      ],
      new Set(),
    )
    const comments = files.get('xl/comments1.xml')!
    // A1 keeps its runs; B1 and C3 are plain
    expect(comments).toContain(
      '<comment ref="A1" authorId="0"><text><r><rPr><b/></rPr><t>Ann:</t></r>',
    )
    expect(comments).toContain('<r><rPr><color rgb="FFFF0000"/></rPr><t>Check &lt;this&gt;</t></r>')
    expect(comments).toContain(
      '<comment ref="B1" authorId="1"><text><t xml:space="preserve">new text</t></text></comment>',
    )
    expect(comments).toContain('<author>Bo &amp; Co</author>')
    expect(comments).toContain('<comment ref="C3" authorId="2">')

    const vml = files.get('xl/drawings/vmlDrawing1.vml')!
    // A1's visible, resized box and B1's (edited note, same cell) box survive as they were
    expect(vml).toContain('style="width:300pt;height:120pt;visibility:visible"')
    expect(vml).toContain('<v:shape id="_x0000_s1026" type="#_x0000_t202" style="width:50pt">')
    // the checkbox is untouched, and C3's new box takes an id no other shape uses
    expect(vml).toContain('ObjectType="Checkbox"')
    const ids = [...vml.matchAll(/<v:shape id="_x0000_s(\d+)"/g)].map((m) => Number(m[1]))
    expect(new Set(ids).size).toBe(ids.length)
    expect(ids).toContain(1028)
    expect(vml).toMatch(/<x:Row>2<\/x:Row><x:Column>2<\/x:Column>/)
  })

  it('an author change rebuilds that note as plain text', async () => {
    const { files, pkg } = excelPackage()
    await applySheetNotes(
      pkg,
      'xl/worksheets/sheet1.xml',
      [{ row: 0, column: 0, author: 'Zed', text: 'Ann:\nCheck <this>' }],
      new Set(),
    )
    const comments = files.get('xl/comments1.xml')!
    expect(comments).not.toContain('<b/>')
    expect(comments).toContain('Ann:\nCheck &lt;this&gt;')
    // B1's note was deleted: its box goes, A1 keeps its own
    const vml = files.get('xl/drawings/vmlDrawing1.vml')!
    expect(vml).not.toContain('_x0000_s1026')
    expect(vml).toContain('visibility:visible')
  })
})
