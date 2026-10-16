import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  applyCellEditsToXlsx,
  readWorkbookThreads,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

import { buildEditFixture } from './fixture-builder'

const THREAD = '{AAAAAAAA-0000-4000-8000-000000000001}'
const PERSON = '{BBBBBBBB-0000-4000-8000-000000000001}'

async function withThread(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  zip.file(
    '[Content_Types].xml',
    (await zip.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
        '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/>' +
        '<Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>',
    ),
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    (await zip.file('xl/_rels/workbook.xml.rels')!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>',
    ),
  )
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId7" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>',
  )
  zip.file(
    'xl/threadedComments/threadedComment1.xml',
    '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
      `<threadedComment ref="A1" dT="2026-10-01T09:00:00.00" personId="${PERSON}" id="${THREAD}" done="0"><text>hello</text></threadedComment></ThreadedComments>`,
  )
  zip.file(
    'xl/persons/person.xml',
    `<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="Ann" id="${PERSON}" userId="ann" providerId="None"/></personList>`,
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function reopen(buffer: Buffer): Promise<JSZip> {
  return JSZip.loadAsync(buffer)
}

function readerOf(zip: JSZip) {
  return {
    has: async (path: string) => zip.file(path) !== null,
    readText: async (path: string) => zip.file(path)!.async('string'),
  }
}

describe('threaded comments through the real save path', () => {
  it('reads the thread the editor will show', async () => {
    const zip = await reopen(await withThread())
    const data = await readWorkbookThreads(readerOf(zip), ['Data'])
    expect(data.sheets.get('Data')?.[0]).toMatchObject({ id: THREAD, row: 0, column: 0 })
    expect(data.persons[0]?.displayName).toBe('Ann')
  })

  it('a notes-only snapshot keeps the file thread and its placeholder', async () => {
    const mutation = await applyCellEditsToXlsx(
      await withThread(),
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
      [{ sheetName: 'Data', notes: [{ row: 2, column: 2, author: 'Ann', text: 'a note' }] }],
    )
    const zip = await reopen(mutation.buffer)
    expect(await zip.file('xl/threadedComments/threadedComment1.xml')!.async('string')).toContain(
      '<text>hello</text>',
    )
    const comments = await zip.file('xl/comments1.xml')!.async('string')
    expect(comments).toContain(`tc=${THREAD}`)
    expect(comments).toContain('a note')
  })

  it('writes a reply and a resolved flag from the editor snapshot', async () => {
    const mutation = await applyCellEditsToXlsx(
      await withThread(),
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
      [
        {
          sheetName: 'Data',
          notes: [],
          threads: [
            {
              id: THREAD,
              row: 0,
              column: 0,
              done: true,
              comments: [
                { id: THREAD, personId: PERSON, time: '2026-10-01T09:00:00.00', text: 'hello' },
                {
                  id: '{CCCCCCCC-0000-4000-8000-000000000001}',
                  personId: PERSON,
                  time: '2026-10-02T09:00:00.00',
                  text: 'done',
                  parentId: THREAD,
                },
              ],
            },
          ],
          persons: [{ id: PERSON, displayName: 'Ann', userId: 'ann', providerId: 'None' }],
        },
      ],
    )
    const zip = await reopen(mutation.buffer)
    const part = await zip.file('xl/threadedComments/threadedComment1.xml')!.async('string')
    expect(part).toContain('done="1"')
    expect(part).toContain(`parentId="${THREAD}"`)
    const data = await readWorkbookThreads(readerOf(zip), ['Data'])
    expect(data.sheets.get('Data')?.[0]).toMatchObject({ done: true })
    expect(data.sheets.get('Data')?.[0]?.comments).toHaveLength(2)
  })
})
