import { expect, test, type Page } from '@playwright/test'
import { execSync } from 'node:child_process'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  type LaunchedApp,
} from './helpers'

/**
 * Copy / cut / paste between workbook windows and with other programs. Two Sheets tabs run in
 * separate renderers, so everything that crosses has to travel through the system clipboard: the
 * GenOffice payload in the copied HTML (formulas, formats, merges, notes, validation, conditional
 * formats), tab-separated text, and a cut that the pasting window reports to the cutting one.
 * Expectations for Excel's side were read from Excel 16 for Mac (docs/qa/sheets-clipboard-*).
 */
const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

async function openSheets(): Promise<{ launched: LaunchedApp; a: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-clipboard-cross-file',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.page.locator('.quick-card').nth(1).click()
  const a = await sheetsPage(launched, 0)
  await a.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  await a.waitForTimeout(1500)
  await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
  return { launched, a }
}

async function sheetsPage(l: LaunchedApp, index: number): Promise<Page> {
  const deadline = Date.now() + 30_000
  for (;;) {
    const pages = l.app.windows().filter((w) => w.url().includes('://sheets/'))
    if (pages[index]) return pages[index]
    if (Date.now() > deadline) throw new Error(`no sheets page ${index}`)
    await l.app.waitForEvent('window', { timeout: 1000 }).catch(() => {})
  }
}

/** a second workbook (a scratch copy of the fixture) in its own tab */
async function openSecond(l: LaunchedApp): Promise<Page> {
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-clip-'))
  const file = join(dir, 'second.xlsx')
  await copyFile(FIXTURE, file)
  await l.app.evaluate(({ app }, path) => {
    app.emit('open-file', { preventDefault() {} }, path)
  }, file)
  const b = await sheetsPage(l, 1)
  await b.waitForSelector('[data-ribbon-body]', { timeout: 30_000 })
  await b.waitForTimeout(2000)
  return b
}

const univer = (expr: string) => `(window).__genofficeDebug.univerAPI.${expr}`
const sheet = `${univer('getActiveWorkbook()')}.getActiveSheet()`
const activate = (p: Page, a1: string) => p.evaluate(`${sheet}.getRange('${a1}').activate()`)
const selectTab = async (l: LaunchedApp, n: number) => {
  await l.page.locator('.tab-item:not(.tab-home)').nth(n).click()
  await l.page.waitForTimeout(700)
}
const cellData = (p: Page, a1: string) =>
  p.evaluate(`${sheet}.getRange('${a1}').getCellData()`) as Promise<Record<string, any> | null>
const formulas = (p: Page, a1: string) =>
  p.evaluate(`${sheet}.getRange('${a1}').getFormulas()`) as Promise<string[][]>
const styleOf = async (p: Page, a1: string) => {
  const data = await cellData(p, a1)
  const s = data?.s
  if (!s) return {} as Record<string, any>
  if (typeof s === 'object') return s as Record<string, any>
  return (await p.evaluate(`${univer('getActiveWorkbook()')}.save().styles['${s}']`)) as Record<
    string,
    any
  >
}
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(700)
}

const BD = (s: number, rgb = '#000000') => ({ s, cl: { rgb } })
const SEED = [
  [
    {
      v: 'Item',
      s: { bl: 1, bg: { rgb: '#FFC000' }, cl: { rgb: '#FFFFFF' }, ht: 2, ff: 'Arial', fs: 12 },
    },
    { v: 'Qty', s: { bl: 1, bg: { rgb: '#FFC000' }, ht: 2 } },
    { v: 'Date' },
  ],
  [
    { v: 'Widget', s: { it: 1, cl: { rgb: '#FF0000' }, ff: 'Times New Roman', fs: 14 } },
    { v: 1234.5, t: 2, s: { n: { pattern: '#,##0.00' } } },
    { v: 46000, t: 2, s: { n: { pattern: 'yyyy-mm-dd' } } },
  ],
  [
    { v: 'line1\nline2', s: { tb: 3, bd: { b: BD(1), t: BD(1), l: BD(1), r: BD(1) } } },
    { f: '=B2*2', v: 2469, t: 2, s: { bd: { b: BD(5, '#0000FF') } } },
    { f: '=SUM(B2:B3)', v: 3703.5, t: 2 },
  ],
  [{ v: 'merged', s: { bg: { rgb: '#00B050' }, ht: 2 } }, { s: { bg: { rgb: '#00B050' } } }, {}],
]

async function seed(p: Page): Promise<void> {
  await p.evaluate(
    ({ cells }) => {
      const api = (window as any).__genofficeDebug.univerAPI
      const s = api.getActiveWorkbook().getActiveSheet()
      s.getRange('A1:C4').setValues(cells)
      s.getRange('A4:B4').merge()
      s.getRange('A2').createOrUpdateNote({
        id: 'n-1-0',
        row: 1,
        col: 0,
        width: 200,
        height: 80,
        note: 'a note',
      })
      s.getRange('C3').setDataValidation(
        api.newDataValidation().requireValueInList(['x', 'y']).build(),
      )
      s.addConditionalFormattingRule(
        s
          .newConditionalFormattingRule()
          .whenNumberGreaterThan(1000)
          .setBackground('#FF9999')
          .setRanges([s.getRange('B2:B3').getRange()])
          .build(),
      )
    },
    { cells: SEED },
  )
}

test('a block pasted into another workbook window keeps formulas, formats, merges and sheet features, in one undo step', async () => {
  const { launched, a } = await openSheets()
  try {
    await seed(a)
    await activate(a, 'A1:C4')
    await a.keyboard.press('Meta+C')
    await a.waitForTimeout(900)
    const b = await openSecond(launched)
    await selectTab(launched, 1)
    await activate(b, 'H10')
    await b.keyboard.press('Meta+V')
    await b.waitForTimeout(1800)

    // formulas move with the paste like a copy inside one window
    expect(await formulas(b, 'I12:J12')).toEqual([['=I11*2', '=SUM(I11:I12)']])
    // formats: fill, font, number format, border, wrap, merge
    expect(await styleOf(b, 'H10')).toMatchObject({
      bl: 1,
      bg: { rgb: '#FFC000' },
      ff: 'Arial',
      fs: 12,
      ht: 2,
    })
    expect(await styleOf(b, 'H11')).toMatchObject({
      it: 1,
      ff: 'Times New Roman',
      cl: { rgb: '#FF0000' },
    })
    expect((await styleOf(b, 'I11')).n).toMatchObject({ pattern: '#,##0.00' })
    expect((await styleOf(b, 'J11')).n).toMatchObject({ pattern: 'yyyy-mm-dd' })
    expect(await styleOf(b, 'H12')).toMatchObject({ tb: 3, bd: { t: { s: 1 }, l: { s: 1 } } })
    expect(await styleOf(b, 'I12')).toMatchObject({ bd: { b: { s: 5 } } })
    expect(await b.evaluate(`${sheet}.getMergedRanges().map((m) => m.getA1Notation())`)).toContain(
      'H13:I13',
    )
    // the cell value, not the conditional format's color, is what the paste carries
    expect((await styleOf(b, 'I11')).bg?.rgb ?? '').not.toBe('#FF9999')
    // notes, validation, conditional formats
    expect(await b.evaluate(`${sheet}.getNotes().map((n) => [n.row, n.col, n.note])`)).toEqual([
      [10, 7, 'a note'],
    ])
    expect(await b.evaluate(`${sheet}.getDataValidations().length`)).toBe(1)
    expect(await b.evaluate(`${sheet}.getConditionalFormattingRules().length`)).toBe(1)

    await undo(b)
    expect(await cellData(b, 'H10')).toBeNull()
    expect(await b.evaluate(`${sheet}.getMergedRanges().length`)).toBe(0)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-clipboard-cross-file')
  }
})

test('cut in one window and paste in another moves the cells', async () => {
  const { launched, a } = await openSheets()
  try {
    await a.evaluate(`${sheet}.getRange('A20:B21').setValues([['x', 'y'], [1, 2]])`)
    await activate(a, 'A20:B21')
    await a.keyboard.press('Meta+X')
    await a.waitForTimeout(800)
    const b = await openSecond(launched)
    await selectTab(launched, 1)
    await activate(b, 'H5')
    await b.keyboard.press('Meta+V')
    await b.waitForTimeout(1800)
    expect(await b.evaluate(`${sheet}.getRange('H5:I6').getRawValues()`)).toEqual([
      ['x', 'y'],
      [1, 2],
    ])
    await selectTab(launched, 0)
    await expect
      .poll(() => a.evaluate(`${sheet}.getRange('A20:B21').getRawValues()`))
      .toEqual([
        [null, null],
        [null, null],
      ])
  } finally {
    await closeAndSaveVideo(launched, 'sheets-clipboard-cut')
  }
})

test('plain text pastes split at tabs and line breaks, with quoted fields kept whole', async () => {
  const { launched, a } = await openSheets()
  try {
    const cases: [string, unknown[][]][] = [
      [
        'a\tb\tc\n1\t2\t3\n',
        [
          ['a', 'b', 'c'],
          [1, 2, 3],
        ],
      ],
      [
        'a,b,c\n1,2,3',
        [
          ['a,b,c', null, null],
          ['1,2,3', null, null],
        ],
      ],
      [
        '"x\ty"\t"say ""hi"""\n5\t=1+2',
        [
          ['x\ty', 'say "hi"', null],
          [5, 3, null],
        ],
      ],
      ['\t\tz', [[null, null, 'z']]],
    ]
    let row = 1
    for (const [text, expected] of cases) {
      await launched.app.evaluate(({ clipboard }, value) => clipboard.writeText(value), text)
      await activate(a, `A${row}`)
      await a.keyboard.press('Meta+V')
      await a.waitForTimeout(900)
      const got = (await a.evaluate(
        `${sheet}.getRange('A${row}:C${row + expected.length - 1}').getRawValues()`,
      )) as unknown[][]
      expect(got).toEqual(expected)
      row += 5
    }
  } finally {
    await closeAndSaveVideo(launched, 'sheets-clipboard-text')
  }
})

test('copy writes the whole display text, never the #### of a narrow column', async () => {
  const { launched, a } = await openSheets()
  try {
    await a.evaluate(
      `${sheet}.getRange('A1').setValues([[{ v: 46000, t: 2, s: { n: { pattern: 'yyyy-mm-dd' } } }]])`,
    )
    await a.evaluate(`${sheet}.setColumnWidth(0, 20)`)
    await activate(a, 'A1')
    await a.keyboard.press('Meta+C')
    await a.waitForTimeout(800)
    const text = await launched.app.evaluate(({ clipboard }) => clipboard.readText())
    expect(text).toBe('2025-12-09')
    const html = await launched.app.evaluate(({ clipboard }) => clipboard.readHTML())
    expect(html).toContain('>2025-12-09<')
    expect(html).toContain('x:num="46000"')
    expect(html).toMatch(/mso-number-format:(&quot;|")yyyy\\-mm\\-dd(&quot;|")/)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-clipboard-narrow')
  }
})

test('several selected areas copy together when they line up and are refused when they do not', async () => {
  const { launched, a } = await openSheets()
  try {
    await a.evaluate(`${sheet}.getRange('A10:B11').setValues([[1, 2], [3, 4]])`)
    await a.evaluate(`${sheet}.getRange('D10:E11').setValues([[5, 6], [7, 8]])`)
    const select = (ranges: number[][]) =>
      a.evaluate(
        async ({ ranges }) => {
          const api = (window as any).__genofficeDebug.univerAPI
          const wb = api.getActiveWorkbook()
          const s = wb.getActiveSheet()
          const sel = (r: number[]) => ({
            range: {
              startRow: r[0],
              endRow: r[1],
              startColumn: r[2],
              endColumn: r[3],
              rangeType: 0,
            },
            primary: {
              startRow: r[0],
              endRow: r[0],
              startColumn: r[2],
              endColumn: r[2],
              actualRow: r[0],
              actualColumn: r[2],
              isMerged: false,
              isMergedMainCell: false,
            },
          })
          await api.executeCommand('sheet.operation.set-selections', {
            unitId: wb.getId(),
            subUnitId: s.getSheetId(),
            selections: ranges.map(sel),
          })
        },
        { ranges },
      )
    await select([
      [9, 10, 0, 1],
      [9, 10, 3, 4],
    ])
    await a.keyboard.press('Meta+C')
    await a.waitForTimeout(800)
    expect(await launched.app.evaluate(({ clipboard }) => clipboard.readText())).toBe(
      '1\t2\t5\t6\n3\t4\t7\t8',
    )
    await launched.app.evaluate(({ clipboard }) => clipboard.writeText('unchanged'))
    await select([
      [9, 10, 0, 1],
      [20, 22, 3, 4],
    ])
    await a.keyboard.press('Meta+C')
    await a.waitForTimeout(800)
    expect(await launched.app.evaluate(({ clipboard }) => clipboard.readText())).toBe('unchanged')
    await expect(a.locator('.workbook-status')).toContainText('multiple selections')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-clipboard-multi')
  }
})

test('a pasted block is written to the file with its own styles, formulas, merges and features', async () => {
  const { launched, a } = await openSheets()
  try {
    await seed(a)
    await activate(a, 'A1:C4')
    await a.keyboard.press('Meta+C')
    await a.waitForTimeout(900)
    const dir = await mkdtemp(join(tmpdir(), 'genoffice-clip-save-'))
    const file = join(dir, 'second.xlsx')
    await copyFile(FIXTURE, file)
    await launched.app.evaluate(({ app }, path) => {
      app.emit('open-file', { preventDefault() {} }, path)
    }, file)
    const b = await sheetsPage(launched, 1)
    await b.waitForSelector('[data-ribbon-body]', { timeout: 30_000 })
    await b.waitForTimeout(2000)
    await selectTab(launched, 1)
    await activate(b, 'H10')
    await b.keyboard.press('Meta+V')
    await b.waitForTimeout(1800)
    await launched.app.evaluate(({ webContents }) => {
      const wcs = webContents.getAllWebContents().filter((w) => w.getURL().includes('://sheets/'))
      wcs.sort((x, y) => x.id - y.id)
      wcs[wcs.length - 1]?.send('menu:action', 'save')
    })
    await expect.poll(() => unzip(file, 'xl/styles.xml'), { timeout: 20_000 }).toContain('dashDot')
    const styles = unzip(file, 'xl/styles.xml')
    expect(styles).toContain('<numFmt numFmtId="164"')
    expect(styles).toContain('FFFFC000')
    expect(styles).toContain('wrapText="1"')
    const sheetXml = unzip(file, 'xl/worksheets/sheet1.xml')
    expect(sheetXml).toContain('<f>I11*2</f>')
    expect(sheetXml).toContain('<mergeCell ref="H13:I13"/>')
    expect(sheetXml).toContain('<conditionalFormatting')
    expect(sheetXml).toContain('<dataValidation ')
    expect(unzip(file, 'xl/comments1.xml')).toContain('a note')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-clipboard-save')
  }
})

function unzip(file: string, entry: string): string {
  try {
    return execSync(`unzip -p "${file}" "${entry}"`).toString()
  } catch {
    return ''
  }
}
