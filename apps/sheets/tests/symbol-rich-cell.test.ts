import { describe, expect, it } from 'vitest'
import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { buildEditFixture } from './fixture-builder'
import { createEditJournal, recordSetRangeValues, toSaveEdits } from '../src/renderer/edit-journal'
import { symbolCellPatch } from '../src/renderer/symbol-rich-cell'

describe('font-specific symbol insertion', () => {
  it('keeps existing text and gives only the new symbol its chosen font through the XLSX writer', async () => {
    const patch = symbolCellPatch('Total', { v: 'Total' }, '∑', 'Cambria Math')
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: patch } })
    const entry = journal.cells.get('sheet-1')?.get('0:0')
    expect(entry?.value).toBe('Total∑')
    expect(entry?.rich).toEqual([
      expect.objectContaining({ text: 'Total' }),
      expect.objectContaining({ text: '∑', family: 'Cambria Math' }),
    ])
    const edit = toSaveEdits(journal)[0]!
    const source = await createBufferEntrySource(await buildEditFixture())
    const plan = await planCellEditsToXlsx(source, [
      {
        sheetName: 'Data',
        row: edit.row,
        column: edit.column,
        writeValue: true,
        cell: { value: edit.value },
        rich: edit.rich,
      },
    ])
    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml')
    expect(worksheet).toContain('<t xml:space="preserve">Total</t>')
    expect(worksheet).toContain('<rFont val="Cambria Math"/>')
    expect(worksheet).toContain('<t xml:space="preserve">∑</t>')
  })

  it('keeps existing rich runs and appends a supplementary-plane symbol at UTF-16 offsets', () => {
    const prior = {
      p: { body: { dataStream: 'A\r\n', textRuns: [{ st: 0, ed: 1, ts: { ff: 'Arial' } }] } },
    }
    const patch = symbolCellPatch('A', prior as never, '𝄞', 'Segoe UI Symbol')
    const body = (
      patch.p as {
        body: { dataStream: string; textRuns: { st: number; ed: number; ts: { ff: string } }[] }
      }
    ).body
    expect(body.dataStream).toBe('A𝄞\r\n')
    expect(body.textRuns).toEqual([
      { st: 0, ed: 1, ts: { ff: 'Arial' } },
      { st: 1, ed: 3, ts: { ff: 'Segoe UI Symbol' } },
    ])
  })

  it('carries a plain styled cell font into both XLSX rich runs', async () => {
    const patch = symbolCellPatch(
      'Styled',
      { v: 'Styled', s: 'style-id' } as never,
      'Ω',
      'Cambria Math',
      { ff: 'Arial', fs: 16, bl: 1, cl: { rgb: '#AA0044' } } as never,
    )
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: patch } })
    const entry = journal.cells.get('sheet-1')?.get('0:0')
    expect(entry?.rich).toEqual([
      expect.objectContaining({
        text: 'Styled',
        family: 'Arial',
        size: 16,
        bold: true,
        color: '#AA0044',
      }),
      expect.objectContaining({
        text: 'Ω',
        family: 'Cambria Math',
        size: 16,
        bold: true,
        color: '#AA0044',
      }),
    ])
    const edit = toSaveEdits(journal)[0]!
    const source = await createBufferEntrySource(await buildEditFixture())
    const plan = await planCellEditsToXlsx(source, [
      {
        sheetName: 'Data',
        row: edit.row,
        column: edit.column,
        writeValue: true,
        cell: { value: edit.value },
        rich: edit.rich,
      },
    ])
    const worksheet = plan.replaced.get('xl/worksheets/sheet1.xml') ?? ''
    const xmlRuns = Array.from(worksheet.matchAll(/<r>[\s\S]*?<\/r>/g), (match) => match[0])
    const styledRun = xmlRuns.find((run) => run.includes('>Styled</t>')) ?? ''
    const glyphRun = xmlRuns.find((run) => run.includes('>Ω</t>')) ?? ''
    for (const run of [styledRun, glyphRun]) {
      expect(run).toContain('<b/>')
      expect(run).toContain('<sz val="16"/>')
      expect(run).toContain('<color rgb="FFAA0044"/>')
    }
    expect(styledRun).toContain('<rFont val="Arial"/>')
    expect(glyphRun).toContain('<rFont val="Cambria Math"/>')
  })

  it('inherits the last rich run format while changing only the symbol font family', () => {
    const patch = symbolCellPatch(
      null,
      {
        p: {
          body: {
            dataStream: 'AΩ\r\n',
            textRuns: [
              { st: 0, ed: 1, ts: { ff: 'Arial', fs: 11 } },
              { st: 1, ed: 2, ts: { ff: 'Cambria Math', fs: 18, bl: 1, cl: { rgb: '#008800' } } },
            ],
          },
        },
      } as never,
      '∑',
      'Segoe UI Symbol',
    )
    expect(patch.p?.body?.textRuns?.[2]?.ts).toEqual({
      ff: 'Segoe UI Symbol',
      fs: 18,
      bl: 1,
      cl: { rgb: '#008800' },
    })
  })

  it('uses reopened rich document text when Univer exposes no cell value', () => {
    const reopened = {
      v: null,
      p: {
        body: {
          dataStream: 'StyledΩ\r\n',
          textRuns: [
            { st: 0, ed: 6, ts: { ff: 'Arial' } },
            { st: 6, ed: 7, ts: { ff: 'Cambria Math' } },
          ],
        },
      },
    }
    const patch = symbolCellPatch(null, reopened as never, '∑', 'Segoe UI Symbol')
    expect(patch.v).toBe('StyledΩ∑')
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 0: patch } })
    const entry = journal.cells.get('sheet-1')?.get('0:0')
    expect(entry?.value).toBe('StyledΩ∑')
    expect(entry?.rich).toEqual([
      expect.objectContaining({ text: 'Styled', family: 'Arial' }),
      expect.objectContaining({ text: 'Ω', family: 'Cambria Math' }),
      expect.objectContaining({ text: '∑', family: 'Segoe UI Symbol' }),
    ])
  })

  it('preserves numeric cell text when appending a symbol', () => {
    const patch = symbolCellPatch(123, { v: 123 }, 'Ω', 'Cambria Math')
    expect(patch.v).toBe('123Ω')
    expect((patch.p as { body: { dataStream: string } }).body.dataStream).toBe('123Ω\r\n')

    const reopened = symbolCellPatch(
      null,
      { p: { body: { dataStream: '123\r456\r\n' } } } as never,
      'Ω',
      'Cambria Math',
    )
    expect(reopened.v).toBe('123\n456Ω')
    expect((reopened.p as { body: { dataStream: string } }).body.dataStream).toBe('123\r456Ω\r\n')
  })
})
