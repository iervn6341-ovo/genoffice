import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

const EVIDENCE =
  process.env.GENOFFICE_FORMULA_CACHE_VERIFICATION_DIR ??
  join(process.cwd(), 'docs/qa/sheets-formula-cache-evidence-2026-10-13')
const OFFSCREEN_ADDRESS = process.env.GENOFFICE_FORMULA_CACHE_SMALL === '1' ? 'M2000' : 'M5000'
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const formulas = [
  {
    sheet: 'Sheet1',
    address: 'K1',
    formula:
      process.env.GENOFFICE_FORMULA_CACHE_PROPER_CASE === '1'
        ? "=1/('Inputs'!A1-40)"
        : '=1/(inputs!A1-40)',
    initial: -1 / 30,
    first: '#DIV/0!',
    second: 0.1,
    third: '#DIV/0!',
  },
  {
    sheet: 'Sheet1',
    address: 'A1',
    formula: '="Ω "&SUM(B1:B2)',
    initial: 'Ω 5',
    first: 'Ω 8',
    second: 'Ω 10',
    third: 'Ω 10',
  },
  {
    sheet: 'Sheet1',
    address: 'D1',
    formula: '=SUM(B1:B2)',
    initial: 5,
    first: 8,
    second: 10,
    third: 10,
  },
  {
    sheet: 'Sheet1',
    address: 'E1',
    formula: '=D1*2',
    initial: 10,
    first: 16,
    second: 20,
    third: 20,
  },
  {
    sheet: 'Sheet1',
    address: 'F1',
    formula: '=IF(D1>=8,TRUE,FALSE)',
    initial: false,
    first: true,
    second: true,
    third: true,
  },
  {
    sheet: 'Sheet1',
    address: 'G1',
    formula: '=1/(B1-5)',
    initial: -1 / 3,
    first: '#DIV/0!',
    second: 0.5,
    third: 0.5,
  },
  {
    sheet: 'Sheet1',
    address: 'H1',
    formula: "=SUM('Inputs'!A1:A2)+D1",
    initial: 35,
    first: 68,
    second: 80,
    third: 70,
  },
  {
    sheet: 'Sheet1',
    address: 'I1',
    formula: '=SUM($B$1:B2)',
    initial: 5,
    first: 8,
    second: 10,
    third: 10,
  },
  {
    sheet: 'Sheet1',
    address: 'J1',
    formula: '=D1+H1',
    initial: 40,
    first: 76,
    second: 90,
    third: 80,
  },
  {
    sheet: 'Sheet1',
    address: OFFSCREEN_ADDRESS,
    formula: '=SUM(B1:B2)*3',
    initial: 15,
    first: 24,
    second: 30,
    third: 30,
  },
  {
    sheet: 'Inputs',
    address: 'B1',
    formula: `='Sheet1'!J1+'Sheet1'!${OFFSCREEN_ADDRESS}`,
    initial: 55,
    first: 100,
    second: 120,
    third: 110,
  },
]
test.skip(!process.env.GENOFFICE_OFFICE_AUDIT, 'Opt-in imported formula cache verification')

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}
function xmlCell(address: string, value: string | number | boolean, formula?: string): string {
  const type = typeof value === 'boolean' ? 'b' : typeof value === 'string' ? 'str' : undefined
  return `<c r="${address}"${type ? ` t="${type}"` : ''}>${formula ? `<f>${escapeXml(formula.slice(1))}</f>` : ''}<v>${escapeXml(String(typeof value === 'boolean' ? Number(value) : value))}</v></c>`
}
async function createFixture(path: string): Promise<void> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${[1, 2].map((n) => `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<workbook xmlns="${NS}" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Inputs" sheetId="2" r:id="rId2"/></sheets><calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${[1, 2].map((n) => `<Relationship Id="rId${n}" Type="${REL}/worksheet" Target="worksheets/sheet${n}.xml"/>`).join('')}</Relationships>`,
  )
  for (const [index, sheet] of ['Sheet1', 'Inputs'].entries()) {
    const cells = formulas
      .filter((f) => f.sheet === sheet)
      .map((f) => ({ address: f.address, xml: xmlCell(f.address, f.initial, f.formula) }))
    const constants =
      sheet === 'Sheet1'
        ? ([
            ['B1', 2],
            ['B2', 3],
            ['C1', 99],
          ] as const)
        : ([
            ['A1', 10],
            ['A2', 20],
          ] as const)
    for (const [address, value] of constants) cells.push({ address, xml: xmlCell(address, value) })
    const rows = new Map<number, typeof cells>()
    for (const cell of cells) {
      const row = Number(cell.address.match(/\d+/)![0])
      const list = rows.get(row) ?? []
      list.push(cell)
      rows.set(row, list)
    }
    const data = [...rows]
      .sort((a, b) => a[0] - b[0])
      .map(
        ([row, list]) =>
          `<row r="${row}">${list
            .sort((a, b) => a.address.localeCompare(b.address))
            .map((c) => c.xml)
            .join('')}</row>`,
      )
      .join('')
    zip.file(
      `xl/worksheets/sheet${index + 1}.xml`,
      `<worksheet xmlns="${NS}"><dimension ref="${sheet === 'Sheet1' ? `A1:${OFFSCREEN_ADDRESS}` : 'A1:B2'}"/><sheetData>${data}</sheetData></worksheet>`,
    )
  }
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}
async function openSheet(launched: Awaited<ReturnType<typeof launchShell>>): Promise<Page> {
  const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await page.waitForFunction(() =>
    Boolean((window as any).__genofficeDebug?.univerAPI?.getActiveWorkbook()),
  )
  await setEditorLayoutWidth(launched.app, '://sheets/', 1280)
  return page
}
async function readLive(page: Page) {
  return page.evaluate((wanted) => {
    const book = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
    return wanted.map((f) => {
      const range = book.getSheetByName(f.sheet).getRange(f.address)
      return {
        sheet: f.sheet,
        address: f.address,
        formula: range.getFormula(),
        value:
          range.getCellDatas()[0]?.[0]?.t === 3
            ? Boolean(range.getRawValue())
            : range.getRawValue(),
        display: range.getDisplayValue(),
      }
    })
  }, formulas)
}
async function assertLive(page: Page, stage: 'initial' | 'first' | 'second' | 'third') {
  await expect
    .poll(
      async () => {
        const live = await readLive(page)
        return live.map((cell, index) => ({
          formula: cell.formula,
          value: cell.value,
          expected: formulas[index]![stage],
        }))
      },
      { timeout: 30_000 },
    )
    .toEqual(formulas.map((f) => ({ formula: f.formula, value: f[stage], expected: f[stage] })))
  return readLive(page)
}
async function readNative(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path), { checkCRC32: true })
  return Promise.all(
    formulas.map(async (f) => {
      const xml = await zip
        .file(`xl/worksheets/sheet${f.sheet === 'Sheet1' ? 1 : 2}.xml`)!
        .async('string')
      const cell =
        xml.match(new RegExp(`<c\\b[^>]*\\br="${f.address}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0] ?? ''
      const decode = (s: string) =>
        s
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"')
          .replace(/&apos;/g, "'")
          .replace(/&amp;/g, '&')
      return {
        sheet: f.sheet,
        address: f.address,
        formula: decode(cell.match(/<f\b[^>]*>([\s\S]*?)<\/f>/)?.[1] ?? ''),
        cache: decode(cell.match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? ''),
        type: cell.match(/\bt="([^"]+)"/)?.[1] ?? 'n',
      }
    }),
  )
}
async function assertNative(path: string, stage: 'first' | 'second' | 'third') {
  const wanted = formulas.map((f) => ({
    sheet: f.sheet,
    address: f.address,
    formula: f.formula.slice(1),
    cache: String(typeof f[stage] === 'boolean' ? Number(f[stage]) : f[stage]),
    type:
      typeof f[stage] === 'boolean'
        ? 'b'
        : typeof f[stage] === 'string'
          ? String(f[stage]).startsWith('#')
            ? 'e'
            : 'str'
          : 'n',
  }))
  await expect.poll(() => readNative(path), { timeout: 20_000 }).toEqual(wanted)
  const zip = await JSZip.loadAsync(await readFile(path), { checkCRC32: true })
  for (const [part, address, expected] of [
    ['sheet1.xml', 'B1', stage === 'first' ? 5 : 7],
    ['sheet1.xml', 'B2', 3],
    ['sheet1.xml', 'C1', 99],
    ['sheet2.xml', 'A1', stage === 'second' ? 50 : 40],
    ['sheet2.xml', 'A2', 20],
  ] as const) {
    const xml = await zip.file(`xl/worksheets/${part}`)!.async('string')
    const cell =
      xml.match(new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0] ?? ''
    expect(cell.match(/<v>([\s\S]*?)<\/v>/)?.[1]).toBe(String(expected))
    expect(cell).not.toContain('<f>')
  }
  return readNative(path)
}
async function save(
  launched: Awaited<ReturnType<typeof launchShell>>,
  path: string,
  stage: 'first' | 'second' | 'third',
) {
  await launched.app.evaluate(({ webContents }) =>
    webContents
      .getAllWebContents()
      .find((wc) => wc.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save'),
  )
  return assertNative(path, stage)
}

test('unedited imported formula chains save current caches across precedent edits and process reopens', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-formula-cache-'))
  const path = join(scratch, 'formula-cache.xlsx')
  await createFixture(path)
  await mkdir(EVIDENCE, { recursive: true })
  await copyFile(path, join(EVIDENCE, 'baseline.xlsx'))
  const summary: Record<string, unknown> = { formulas }
  for (const stage of ['first', 'second', 'third', 'final'] as const) {
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: path,
      videoDir: `formula-cache-${stage}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      const page = await openSheet(launched)
      summary[`${stage}Open`] = await assertLive(
        page,
        stage === 'first'
          ? 'initial'
          : stage === 'second'
            ? 'first'
            : stage === 'third'
              ? 'second'
              : 'third',
      )
      if (stage !== 'final') {
        await page.evaluate(
          ({ b1, inputA1 }) => {
            const book = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
            if (b1 !== null) book.getSheetByName('Sheet1').getRange('B1').setValue(b1)
            book.getSheetByName('Inputs').getRange('A1').setValue(inputA1)
          },
          stage === 'first'
            ? { b1: 5, inputA1: 40 }
            : stage === 'second'
              ? { b1: 7, inputA1: 50 }
              : { b1: null, inputA1: 40 },
        )
        if (stage === 'first') summary.firstBeforeSave = await assertLive(page, stage)
        // The second save immediately follows the edits; the product must wait
        // for fresh calculation results rather than reuse a completed old pass.
        summary[`${stage}Native`] = await save(launched, path, stage)
        await copyFile(path, join(EVIDENCE, `${stage}-save.xlsx`))
        summary[`${stage}AfterSave`] = await assertLive(page, stage)
      }
      await writeFile(join(EVIDENCE, `${stage}.png`), await page.screenshot())
      await writeFile(join(EVIDENCE, 'summary.json'), JSON.stringify(summary, null, 2))
    } finally {
      await closeAndSaveVideo(launched, `sheets-formula-cache-${stage}`)
    }
  }
})
