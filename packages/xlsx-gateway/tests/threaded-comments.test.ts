import { describe, expect, it } from 'vitest'

import {
  applySheetThreads,
  buildThreadedCommentsXml,
  isThreadPlaceholderAuthor,
  newGuid,
  parsePersons,
  parseThreadedComments,
  readPersons,
  readSheetThreads,
  type SheetThread,
} from '../src/gateway/xlsx-threaded-comments'

const ROOT = '{AAAAAAAA-0000-4000-8000-000000000001}'
const REPLY = '{AAAAAAAA-0000-4000-8000-000000000002}'
const PERSON = '{BBBBBBBB-0000-4000-8000-000000000001}'

function excelPackage() {
  const files = new Map<string, string>([
    [
      '[Content_Types].xml',
      '<Types><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' +
        '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>' +
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>',
    ],
    ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/><legacyDrawing r:id="rId2"/></worksheet>'],
    [
      'xl/worksheets/_rels/sheet1.xml.rels',
      '<Relationships>' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/>' +
        '</Relationships>',
    ],
    [
      'xl/_rels/workbook.xml.rels',
      '<Relationships><Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>',
    ],
    [
      'xl/comments1.xml',
      '<comments><authors><author>tc=' +
        ROOT +
        '</author><author>Ann</author></authors><commentList>' +
        `<comment ref="B2" authorId="0"><text><t>[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    first &amp; foremost</t></text></comment>` +
        '<comment ref="D4" authorId="1"><text><t>plain note</t></text></comment>' +
        '</commentList></comments>',
    ],
    [
      'xl/drawings/vmlDrawing1.vml',
      '<xml><v:shapetype id="_x0000_t202"/>' +
        '<v:shape id="_x0000_s1025" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>' +
        '<v:shape id="_x0000_s1026" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>3</x:Row><x:Column>3</x:Column></x:ClientData></v:shape>' +
        '</xml>',
    ],
    [
      'xl/threadedComments/threadedComment1.xml',
      '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<threadedComment ref="B2" dT="2026-10-01T09:00:00.00" personId="${PERSON}" id="${ROOT}" done="0"><text>first &amp; foremost</text><mentions><mention mentionpersonId="${PERSON}" mentionId="{M}" startIndex="0" length="5"/></mentions></threadedComment>` +
        `<threadedComment ref="B2" dT="2026-10-01T09:05:00.00" personId="${PERSON}" id="${REPLY}" parentId="${ROOT}"><text>reply text</text><extLst><ext uri="{X}"/></extLst></threadedComment>` +
        '</ThreadedComments>',
    ],
    [
      'xl/persons/person.xml',
      `<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><person displayName="Ann" id="${PERSON}" userId="ann@example.com" providerId="AD"><extLst/></person></personList>`,
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

const SHEET = 'xl/worksheets/sheet1.xml'

describe('threaded comments reader', () => {
  it('groups replies under their root and keeps raw elements, persons and the root tag', async () => {
    const { pkg } = excelPackage()
    const threads = await readSheetThreads(pkg, SHEET)
    expect(threads).toHaveLength(1)
    const [thread] = threads
    expect(thread).toMatchObject({ id: ROOT, row: 1, column: 1, done: false })
    expect(thread!.comments.map((comment) => comment.text)).toEqual([
      'first & foremost',
      'reply text',
    ])
    expect(thread!.comments[1]!.parentId).toBe(ROOT)
    expect(thread!.comments[0]!.raw).toContain('<mentions>')
    const persons = await readPersons(pkg)
    expect(persons).toEqual([
      expect.objectContaining({ id: PERSON, displayName: 'Ann', providerId: 'AD' }),
    ])
    expect(isThreadPlaceholderAuthor(`tc=${ROOT}`)).toBe(true)
  })

  it('keeps comments with a non-cell ref and orphan replies instead of dropping them', () => {
    const parsed = parseThreadedComments(
      '<ThreadedComments><threadedComment ref="not-a-cell" id="{1}" personId="{p}" dT="t"><text>x</text></threadedComment>' +
        '<threadedComment ref="C3" id="{2}" parentId="{missing}" personId="{p}" dT="t"><text>lost reply</text></threadedComment></ThreadedComments>',
    )
    expect(parsed.unparsed).toHaveLength(1)
    expect(parsed.threads).toHaveLength(1)
    expect(parsed.threads[0]!.comments[0]!.parentId).toBe('{missing}')
    expect(
      buildThreadedCommentsXml(parsed.threads, parsed.rootStartTag, parsed.unparsed),
    ).toContain('parentId="{missing}"')
  })
})

describe('applySheetThreads', () => {
  it('rewrites an untouched set byte-for-byte for the threaded part and keeps notes', async () => {
    const { files, pkg } = excelPackage()
    const before = files.get('xl/threadedComments/threadedComment1.xml')!
    const threads = await readSheetThreads(pkg, SHEET)
    await applySheetThreads(
      pkg,
      SHEET,
      [{ row: 3, column: 3, author: 'Ann', text: 'plain note' }],
      threads,
      await readPersons(pkg),
      new Set(),
    )
    const after = files.get('xl/threadedComments/threadedComment1.xml')!
    expect(after.replace(/^<\?xml[^>]*>\n/, '')).toBe(before)
    expect(files.get('xl/comments1.xml')).toContain('plain note')
    expect(files.get('xl/comments1.xml')).toContain(`tc=${ROOT}`)
    expect(files.get('xl/persons/person.xml')).toContain('<extLst/>')
  })

  it('adds a reply, resolves the thread and a new person without touching other comments', async () => {
    const { files, pkg } = excelPackage()
    const [thread] = await readSheetThreads(pkg, SHEET)
    const persons = await readPersons(pkg)
    const newPerson = {
      id: '{CCCCCCCC-0000-4000-8000-000000000001}',
      displayName: 'Bo & Co',
      userId: 'bo',
      providerId: 'None',
    }
    const edited: SheetThread = {
      ...thread!,
      done: true,
      comments: [
        { ...thread!.comments[0]!, raw: undefined },
        thread!.comments[1]!,
        {
          id: '{DDDDDDDD-0000-4000-8000-000000000001}',
          personId: newPerson.id,
          time: '2026-10-15T08:00:00.00',
          text: 'new <reply>',
          parentId: ROOT,
        },
      ],
    }
    await applySheetThreads(
      pkg,
      SHEET,
      [{ row: 3, column: 3, author: 'Ann', text: 'plain note' }],
      [edited],
      [...persons, newPerson],
      new Set(),
    )
    const part = files.get('xl/threadedComments/threadedComment1.xml')!
    expect(part).toContain(`id="${ROOT}" done="1"><text>first &amp; foremost</text>`)
    expect(part).not.toContain('<mentions>') // text/done edit regenerated the root element
    expect(part).toContain('<extLst><ext uri="{X}"/></extLst>') // untouched reply keeps its raw xml
    expect(part).toContain('<text>new &lt;reply&gt;</text>')
    expect(files.get('xl/persons/person.xml')).toContain('displayName="Bo &amp; Co"')
    expect(files.get('xl/persons/person.xml')).toContain('providerId="AD"><extLst/></person>')
    const reread = await readSheetThreads(pkg, SHEET)
    expect(reread[0]).toMatchObject({ done: true })
    expect(reread[0]!.comments).toHaveLength(3)
  })

  it('creates the parts, relationships and content types for a first thread', async () => {
    const files = new Map<string, string>([
      ['[Content_Types].xml', '<Types></Types>'],
      ['xl/worksheets/sheet1.xml', '<worksheet><sheetData/></worksheet>'],
    ])
    const pkg = {
      paths: async () => [...files.keys()],
      has: async (path: string) => files.has(path),
      readText: async (path: string) => files.get(path) ?? '',
      write: (path: string, content: string) => void files.set(path, content),
      add: (path: string, content: string) => void files.set(path, content),
      remove: (path: string) => void files.delete(path),
    }
    const id = newGuid(() => 0.5)
    const personId = newGuid(() => 0.25)
    await applySheetThreads(
      pkg,
      SHEET,
      [],
      [
        {
          id,
          row: 0,
          column: 0,
          done: false,
          comments: [{ id, personId, time: '2026-10-15T08:00:00.00', text: 'hello\nworld' }],
        },
      ],
      [{ id: personId, displayName: 'Me', userId: 'me', providerId: 'None' }],
      new Set(),
    )
    expect(files.get('xl/threadedComments/threadedComment1.xml')).toContain(
      `<threadedComment ref="A1" dT="2026-10-15T08:00:00.00" personId="${personId}" id="${id}" done="0"><text>hello\nworld</text>`,
    )
    expect(files.get('xl/worksheets/_rels/sheet1.xml.rels')).toContain(
      'Target="../threadedComments/threadedComment1.xml"',
    )
    expect(files.get('xl/_rels/workbook.xml.rels')).toContain('Target="persons/person.xml"')
    const types = files.get('[Content_Types].xml')!
    expect(types).toContain('application/vnd.ms-excel.threadedcomments+xml')
    expect(types).toContain('application/vnd.ms-excel.person+xml')
    expect(files.get('xl/comments1.xml')).toContain(`tc=${id}`)
    expect(files.get('xl/comments1.xml')).toContain('Comment:\n    hello\n    world')
    expect(parsePersons(files.get('xl/persons/person.xml')!)).toHaveLength(1)
    expect(await readSheetThreads(pkg, SHEET)).toHaveLength(1)
  })

  it('removes the threaded part, relationship and placeholder when the last thread is deleted', async () => {
    const { files, pkg } = excelPackage()
    await applySheetThreads(
      pkg,
      SHEET,
      [{ row: 3, column: 3, author: 'Ann', text: 'plain note' }],
      [],
      await readPersons(pkg),
      new Set(),
    )
    expect(files.has('xl/threadedComments/threadedComment1.xml')).toBe(false)
    expect(files.get('xl/worksheets/_rels/sheet1.xml.rels')).not.toContain('threadedComment')
    expect(files.get('[Content_Types].xml')).not.toContain('threadedComment1')
    expect(files.get('xl/comments1.xml')).not.toContain('tc=')
    expect(files.get('xl/comments1.xml')).toContain('plain note')
    // persons stay: other sheets may still use them
    expect(files.has('xl/persons/person.xml')).toBe(true)
  })

  it('refuses a note and a thread on the same cell', async () => {
    const { pkg } = excelPackage()
    await expect(
      applySheetThreads(
        pkg,
        SHEET,
        [{ row: 1, column: 1, author: 'x', text: 'clash' }],
        await readSheetThreads(pkg, SHEET),
        [],
        new Set(),
      ),
    ).rejects.toThrow(/both a note and a threaded comment/)
  })
})
