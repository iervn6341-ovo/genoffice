import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { applyCellEditsToXlsx, type CellEdit } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { blankWorkbook, formulaReadRanges, writeWorkbook } from '../src/formats/xlsx'
import { xlsxSidecarPath } from '../src/resources'
import { run, tempDir } from './helpers'

const sidecar = Boolean(xlsxSidecarPath())
afterEach(() => vi.unstubAllEnvs())

async function fixture(): Promise<Buffer> {
  const cells: Omit<CellEdit, 'writeValue'>[] = [
    { sheetName: 'Data', row: 0, column: 0, cell: { value: 1 } },
    { sheetName: 'Data', row: 0, column: 1, cell: { value: 2 } },
    { sheetName: 'Data', row: 0, column: 2, cell: { value: 3, formula: '=SUM(A1:B1)' } },
    { sheetName: 'Data', row: 0, column: 3, cell: { value: 6, formula: '=C1*2' } },
    { sheetName: 'Summary', row: 0, column: 0, cell: { value: 6, formula: '=Data!D1' } },
  ]
  const buffer = (
    await applyCellEditsToXlsx(
      await blankWorkbook('Data'),
      cells.map((c) => ({ ...c, writeValue: true })),
      [],
      [],
      {
        renames: [],
        additions: [{ name: 'Summary' }],
        removals: [],
        order: ['Data', 'Summary'],
      },
    )
  ).buffer
  const zip = await JSZip.loadAsync(buffer)
  const data = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    data
      .replace('<f>SUM(A1:B1)</f>', '<f>SUM(A1:B1)</f><v>3</v>')
      .replace('<f>C1*2</f>', '<f>C1*2</f><v>6</v>'),
  )
  const summary = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
  zip.file('xl/worksheets/sheet2.xml', summary.replace('<f>Data!D1</f>', '<f>Data!D1</f><v>6</v>'))
  return zip.generateAsync({ type: 'nodebuffer' })
}

const edit: CellEdit = {
  sheetName: 'Data',
  row: 0,
  column: 1,
  writeValue: true,
  cell: { value: 20 },
}
async function sheets(file: string): Promise<string[]> {
  const zip = await JSZip.loadAsync(readFileSync(file))
  return Promise.all(
    ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'].map((p) =>
      zip.file(p)!.async('string'),
    ),
  )
}

describe('dependent formula caches', () => {
  it('keeps distant formula reads bounded instead of scanning the used rectangle', () => {
    const ranges = formulaReadRanges([
      { row: 0, column: 0 },
      { row: 1048575, column: 16383 },
    ])
    expect(ranges).toHaveLength(2)
    expect(
      ranges.reduce(
        (n, r) => n + (r.endRow - r.startRow + 1) * (r.endColumn - r.startColumn + 1),
        0,
      ),
    ).toBe(2)
  })

  it.skipIf(!sidecar)(
    'refreshes shared-formula followers without expanding or losing their formula tags',
    async () => {
      const zip = await JSZip.loadAsync(await fixture())
      const part = 'xl/worksheets/sheet1.xml'
      const xml = await zip.file(part)!.async('string')
      zip.file(
        part,
        xml
          .replace('<f>SUM(A1:B1)</f>', '<f t="shared" si="0" ref="C1:D1">SUM(A1:B1)</f>')
          .replace('<f>C1*2</f>', '<f t="shared" si="0"/>'),
      )
      const out = join(tempDir(), 'shared.xlsx')
      const result = await writeWorkbook(
        await zip.generateAsync({ type: 'nodebuffer' }),
        [edit],
        out,
      )
      expect(result.cachedValues).toBe(true)
      const [data, summary] = await sheets(out)
      expect(data).toContain('<f t="shared" si="0"/><v>41</v>')
      expect(summary).toContain('<f>Data!D1</f><v>41</v>')
    },
  )

  it.skipIf(!sidecar)('keeps a replacement literal and recalculates its dependents', async () => {
    const out = join(tempDir(), 'literal.xlsx')
    expect(
      (await writeWorkbook(await fixture(), [{ ...edit, column: 2, cell: { value: 100 } }], out))
        .cachedValues,
    ).toBe(true)
    const [data, summary] = await sheets(out)
    expect(data).toContain('<c r="C1"><v>100</v></c>')
    expect(data).toContain('<f>C1*2</f><v>200</v>')
    expect(summary).toContain('<f>Data!D1</f><v>200</v>')
  })

  it.skipIf(!sidecar)(
    'clears an unsupported result rather than retaining its old numeric cache',
    async () => {
      const zip = await JSZip.loadAsync(await fixture())
      const part = 'xl/worksheets/sheet2.xml'
      zip.file(
        part,
        (await zip.file(part)!.async('string')).replace('Data!D1', 'UNKNOWN_FUNCTION(Data!D1)'),
      )
      const out = join(tempDir(), 'unsupported.xlsx')
      const result = await writeWorkbook(
        await zip.generateAsync({ type: 'nodebuffer' }),
        [edit],
        out,
      )
      expect(result.cachedValues).toBe(false)
      expect(result.uncached).toContain('Summary!A1')
      const [data, summary] = await sheets(out)
      expect(data).toContain('<f>C1*2</f><v>42</v>')
      expect(summary).toContain('<f>UNKNOWN_FUNCTION(Data!D1)</f></c>')
    },
  )
  it.skipIf(!sidecar)(
    'refreshes existing chains and other sheets after a constant edit, including CSV export',
    async () => {
      const dir = tempDir()
      const out = join(dir, 'updated.xlsx')
      const result = await writeWorkbook(await fixture(), [edit], out)
      expect(result).toMatchObject({ cells: 1, formulas: 0, cachedValues: true })
      const [data, summary] = await sheets(out)
      expect(data).toMatch(/<c r="C1"[^>]*><f>SUM\(A1:B1\)<\/f><v>21<\/v>/)
      expect(data).toMatch(/<c r="D1"[^>]*><f>C1\*2<\/f><v>42<\/v>/)
      expect(summary).toMatch(/<f>Data!D1<\/f><v>42<\/v>/)
      const csv = join(dir, 'updated.csv')
      expect((await run(['convert', out, '--to', 'csv', '--out', csv])).code).toBe(0)
      expect(
        readFileSync(csv, 'utf8')
          .replace(/^\uFEFF/, '')
          .trim(),
      ).toBe('1,20,21,42')
    },
  )

  it.skipIf(!sidecar)('refreshes caches using final sheet names after a rename', async () => {
    const out = join(tempDir(), 'renamed.xlsx')
    const result = await writeWorkbook(await fixture(), [edit], out, {
      plan: {
        renames: [{ sheetName: 'Data', newName: 'Renamed' }],
        additions: [],
        removals: [],
        order: ['Renamed', 'Summary'],
      },
      renames: { Data: 'Renamed' },
    })
    expect(result.cachedValues).toBe(true)
    const [, summary] = await sheets(out)
    expect(summary).toMatch(/<f>Renamed!D1<\/f><v>42<\/v>/)
  })

  it('drops all dependent caches when the calculation engine cannot run', async () => {
    vi.stubEnv('XLSX_SIDECAR_PATH', join(tempDir(), 'missing-engine'))
    const out = join(tempDir(), 'uncached.xlsx')
    const result = await writeWorkbook(await fixture(), [edit], out)
    expect(result.cachedValues).toBe(false)
    expect(result.warning).toContain('recalculates on open')
    const [data, summary] = await sheets(out)
    expect(data).toContain('<f>SUM(A1:B1)</f></c>')
    expect(data).toContain('<f>C1*2</f></c>')
    expect(summary).toContain('<f>Data!D1</f></c>')
    expect(data).toContain('<v>20</v>')
  })

  it('preserves existing caches for formatting-only edits', async () => {
    vi.stubEnv('XLSX_SIDECAR_PATH', join(tempDir(), 'missing-engine'))
    const out = join(tempDir(), 'styled.xlsx')
    expect(
      (
        await writeWorkbook(
          await fixture(),
          [{ ...edit, writeValue: false, style: { bold: true } }],
          out,
        )
      ).cachedValues,
    ).toBe(true)
    const [data, summary] = await sheets(out)
    expect(data).toContain('<f>SUM(A1:B1)</f><v>3</v>')
    expect(summary).toContain('<f>Data!D1</f><v>6</v>')
  })
})
