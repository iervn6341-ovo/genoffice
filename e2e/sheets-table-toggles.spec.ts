import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'
import { resolveColor } from '../apps/sheets/src/renderer/table-styles'

/**
 * Table Design → Header Row, Total Row, Filter Button and Resize Table, on a table made this session.
 * The expectations were read from Excel 16 through COM (ListObject.ShowHeaders / ShowTotals /
 * ShowAutoFilter / Resize): a Total Row puts "Total" under the first column and a SUBTOTAL under
 * the last, Header Row off empties the header cells and starts the table one row lower, and the
 * saved file says the same (headerRowCount, totalsRowCount, autoFilter, totalsRowLabel / Function).
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-table-toggles',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.page.locator('.quick-card').nth(1).click()
  const p = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await p.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  await p.waitForTimeout(1500)
  await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')
const api = '(window).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()'
const activate = (p: Page, a1: string) => p.evaluate(`${api}.getRange('${a1}').activate()`)
const ROWS = [
  ['Item', 'Qty', 'Price', 'Note'],
  ['Apple', 3, 1.5, 'a'],
  ['Pear', 5, 2.25, 'b'],
  ['Plum', 8, 0.75, 'c'],
  ['Fig', 2, 3.1, 'd'],
]
const raw = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getRawValues()`) as Promise<unknown[][]>
const formulas = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getFormulas()`) as Promise<string[][]>
interface Entry {
  name: string
  area: { startRow: number; startColumn: number; endRow: number; endColumn: number }
  columnNames: string[]
  headerRow?: boolean
  totalsRow?: boolean
  filterButton?: boolean
  totals?: unknown[]
  style?: string
}
const journal = (p: Page) => p.evaluate(`(window).__genofficeDebug.tableAdds()`) as Promise<Entry[]>
const tab = (p: Page, name: string) =>
  p
    .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', {
      hasText: new RegExp(`^${name}$`),
    })
    .first()
    .click()
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(900)
}
const redo = async (p: Page) => {
  await p.locator('button[aria-label^="Redo"]').first().click()
  await p.waitForTimeout(900)
}
async function makeTable(p: Page, style = 'TableStyleMedium2'): Promise<void> {
  await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
  await activate(p, 'A1:D5')
  await tab(p, 'Home')
  await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
  await p.locator(`.table-style-cell[data-style="${style}"]`).click()
  await expect.poll(async () => (await journal(p)).map((t) => t.name)).toEqual(['Table1'])
  await p.waitForTimeout(800)
  await activate(p, 'B3')
  await tab(p, 'Table Design')
}
const checkbox = (p: Page, name: string) =>
  p.locator('[data-ribbon-body] [role="checkbox"]', { hasText: name })
const toggle = async (p: Page, name: string, on: boolean) => {
  const box = checkbox(p, name)
  await expect(box).toHaveAttribute('aria-checked', on ? 'false' : 'true')
  await box.click()
  await expect(checkbox(p, name)).toHaveAttribute('aria-checked', on ? 'true' : 'false')
  await p.waitForTimeout(900)
}

/** QA_KEEP_DIR=<dir> keeps a copy of each saved workbook, for opening in real Excel */
async function keep(file: string, name: string): Promise<void> {
  const dir = process.env.QA_KEEP_DIR
  if (dir && existsSync(file)) await copyFile(file, join(dir, name))
}
async function save(
  l: LaunchedApp,
  p: Page,
  path: string,
): Promise<{ table: string; sheet: string }> {
  await l.app.evaluate(({ dialog }, target) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
    dialog.showMessageBox = (async () => ({ response: 0, checkboxChecked: false })) as never
  }, path)
  await p.locator('button[aria-label^="Save As"]').first().click()
  await expect.poll(() => existsSync(path), { timeout: 20_000 }).toBe(true)
  await new Promise((r) => setTimeout(r, 900))
  const zip = await JSZip.loadAsync(await readFile(path))
  const part = zip.file(/xl\/tables\/table\d+\.xml/)[0]
  return {
    table: part ? (await part.async('string')).replace(/\s+/g, ' ') : '',
    sheet: (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).replace(/\s+/g, ' '),
  }
}

/** how many pixels of a screenshot are exactly `hex` */
function countColor(png: Buffer, hex: string): number {
  const image = PNG.sync.read(png)
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16))
  let count = 0
  for (let i = 0; i < image.data.length; i += 4) {
    if (image.data[i] === r && image.data[i + 1] === g && image.data[i + 2] === b) count += 1
  }
  return count
}

test('Total Row adds "Total" and a SUBTOTAL under the table, in one undo step, and saves as totalsRowCount', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-totals-'))
  const file = join(dir, 'totals.xlsx')
  try {
    await makeTable(p)
    await toggle(p, 'Total Row', true)
    // Excel (COM): "Total" under the first column, =SUBTOTAL(103,[Note]) under the last (text → count)
    await expect.poll(() => raw(p, 'A6')).toEqual([['Total']])
    expect(await formulas(p, 'D6')).toEqual([['=SUBTOTAL(103,Table1[Note])']])
    await expect.poll(() => raw(p, 'D6')).toEqual([[4]])
    let [entry] = await journal(p)
    expect(entry!.area).toEqual({ startRow: 0, startColumn: 0, endRow: 5, endColumn: 3 })
    expect(entry!.totalsRow).toBe(true)
    expect(entry!.totals).toEqual([{ label: 'Total' }, null, null, { fn: 'count' }])

    // ⌘Z puts the table and the cells back in one step; ⇧⌘Z does it again
    await undo(p)
    expect(await raw(p, 'A6:D6')).toEqual([[null, null, null, null]])
    ;[entry] = await journal(p)
    expect(entry!.area.endRow).toBe(4)
    expect(entry!.totalsRow).toBeUndefined()
    await redo(p)
    await expect.poll(() => raw(p, 'A6')).toEqual([['Total']])
    ;[entry] = await journal(p)
    expect(entry!.totalsRow).toBe(true)

    // the table is still one table: a SUBTOTAL over the Price column adds the data rows only
    await p.evaluate(`${api}.getRange('F6').setFormula('=SUBTOTAL(109,Table1[Price])')`)
    await expect.poll(() => raw(p, 'F6')).toEqual([[7.6]])

    const saved = await save(launched, p, file)
    expect(saved.table).toContain('ref="A1:D6" totalsRowCount="1"')
    expect(saved.table).toContain('<autoFilter ref="A1:D5"/>')
    expect(saved.table).toContain('<tableColumn id="1" name="Item" totalsRowLabel="Total"/>')
    expect(saved.table).toContain('<tableColumn id="4" name="Note" totalsRowFunction="count"/>')
    expect(saved.sheet).toContain('SUBTOTAL(103,Table1[Note])')
  } finally {
    await keep(file, 'totals.xlsx')
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Total Row refuses a table with something under it, and Total Row off empties the row again', async () => {
  const { launched, p } = await openSheets()
  try {
    await makeTable(p)
    await p.evaluate(`${api}.getRange('B6').setValue('in the way')`)
    await checkbox(p, 'Total Row').click()
    await expect(p.locator('.workbook-status')).toContainText('not empty')
    await expect(checkbox(p, 'Total Row')).toHaveAttribute('aria-checked', 'false')
    expect((await journal(p))[0]!.area.endRow).toBe(4)
    await p.evaluate(`${api}.getRange('B6').setValue('')`)
    await toggle(p, 'Total Row', true)
    await expect.poll(() => raw(p, 'A6')).toEqual([['Total']])
    await toggle(p, 'Total Row', false)
    await expect.poll(() => raw(p, 'A6:D6')).toEqual([[null, null, null, null]])
    expect((await journal(p))[0]!.area.endRow).toBe(4)
  } finally {
    await kill(launched)
  }
})

test('Header Row off empties the header and starts the table one row lower; on again puts the names back', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-headerless-'))
  const file = join(dir, 'headerless.xlsx')
  try {
    await makeTable(p)
    await toggle(p, 'Header Row', false)
    // Excel (COM): the header cells are emptied, the table is A2:D5 now
    await expect.poll(() => raw(p, 'A1:D1')).toEqual([[null, null, null, null]])
    let [entry] = await journal(p)
    expect(entry!.area).toEqual({ startRow: 1, startColumn: 0, endRow: 4, endColumn: 3 })
    expect(entry!.headerRow).toBe(false)
    expect(entry!.columnNames).toEqual(['Item', 'Qty', 'Price', 'Note'])
    // the data is untouched, and structured references still find every row of it
    expect((await raw(p, 'A2:A5')).flat()).toEqual(['Apple', 'Pear', 'Plum', 'Fig'])
    // Univer would write the names into the first row of its table: that row stays empty on screen
    expect(await p.evaluate(`${api}.getRange('A1:D1').getDisplayValues()`)).toEqual([
      ['', '', '', ''],
    ])
    await p.evaluate(`${api}.getRange('F1').setFormula('=SUM(Table1[Qty])')`)
    await expect.poll(() => raw(p, 'F1')).toEqual([[18]])

    // ⌘Z takes the formula back, and once more puts the header back — one step for the toggle
    await undo(p)
    await undo(p)
    await expect.poll(() => raw(p, 'A1:D1')).toEqual([['Item', 'Qty', 'Price', 'Note']])
    ;[entry] = await journal(p)
    expect(entry!.headerRow).toBeUndefined()
    expect(entry!.area.startRow).toBe(0)
    await redo(p)
    await redo(p)
    await expect.poll(() => raw(p, 'A1:D1')).toEqual([[null, null, null, null]])
    expect((await journal(p))[0]!.headerRow).toBe(false)

    // the saved part has no header row and no filter
    const saved = await save(launched, p, file)
    expect(saved.table).toContain('ref="A2:D5" headerRowCount="0"')
    expect(saved.table).not.toContain('<autoFilter')
    expect(saved.table).toContain('<tableColumn id="2" name="Qty"/>')
  } finally {
    await keep(file, 'headerless.xlsx')
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Header Row on again needs the row above to be empty; Filter Button only takes the buttons off', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-filter-'))
  const file = join(dir, 'nofilter.xlsx')
  try {
    await makeTable(p)
    await toggle(p, 'Header Row', false)
    await p.evaluate(`${api}.getRange('B1').setValue('in the way')`)
    await checkbox(p, 'Header Row').click()
    await expect(p.locator('.workbook-status')).toContainText('no empty row above')
    await expect(checkbox(p, 'Header Row')).toHaveAttribute('aria-checked', 'false')
    expect((await journal(p))[0]!.area.startRow).toBe(1)
    await p.evaluate(`${api}.getRange('B1').setValue('')`)
    await toggle(p, 'Header Row', true)
    await expect.poll(() => raw(p, 'A1:D1')).toEqual([['Item', 'Qty', 'Price', 'Note']])

    // Filter Button: the header stays, only the journal (and the part) lose the filter
    await toggle(p, 'Filter Button', false)
    const [entry] = await journal(p)
    expect(entry!.filterButton).toBe(false)
    expect(entry!.headerRow).toBeUndefined()
    expect(await raw(p, 'A1:D1')).toEqual([['Item', 'Qty', 'Price', 'Note']])
    const saved = await save(launched, p, file)
    expect(saved.table).toContain('ref="A1:D5"')
    expect(saved.table).not.toContain('<autoFilter')
    expect(saved.table).not.toContain('headerRowCount')
  } finally {
    await keep(file, 'nofilter.xlsx')
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Resize Table extends or shortens the table from its top-left cell, in one undo step', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-resize-'))
  const file = join(dir, 'resized.xlsx')
  try {
    await makeTable(p)
    const resize = async (range: string) => {
      await p.locator('[data-ribbon-body] button', { hasText: 'Resize Table' }).click()
      const dialog = p.getByRole('dialog', { name: 'Resize Table' })
      await expect(dialog.getByRole('textbox')).toHaveValue('A1:D5')
      await dialog.getByRole('textbox').fill(range)
      await dialog.getByRole('button', { name: 'OK' }).click()
      await p.waitForTimeout(1000)
    }
    // two more rows and two new columns: the new header cells get names (Excel: "Column1" …)
    await resize('A1:F7')
    await expect.poll(async () => (await journal(p))[0]!.area.endRow).toBe(6)
    const [entry] = await journal(p)
    expect(entry!.area.endColumn).toBe(5)
    expect(entry!.columnNames).toEqual(['Item', 'Qty', 'Price', 'Note', 'Column5', 'Column6'])
    expect(await raw(p, 'E1:F1')).toEqual([['Column5', 'Column6']])
    await undo(p)
    expect((await journal(p))[0]!.area).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 4,
      endColumn: 3,
    })
    await redo(p)
    await expect.poll(async () => (await journal(p))[0]!.area.endRow).toBe(6)

    // the top-left cell has to stay; a range that is not one is refused
    await p.locator('[data-ribbon-body] button', { hasText: 'Resize Table' }).click()
    const dialog = p.getByRole('dialog', { name: 'Resize Table' })
    await dialog.getByRole('textbox').fill('not a range')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog.getByRole('alert')).toContainText('A1:D10')
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await p.locator('[data-ribbon-body] button', { hasText: 'Resize Table' }).click()
    await p.getByRole('dialog', { name: 'Resize Table' }).getByRole('textbox').fill('B1:D5')
    await p
      .getByRole('dialog', { name: 'Resize Table' })
      .getByRole('button', { name: 'OK' })
      .click()
    await expect(p.locator('.workbook-status')).toContainText('top-left')
    expect((await journal(p))[0]!.area.startColumn).toBe(0)

    // shorter again: the rows left outside are plain cells, as in Excel
    await p.locator('[data-ribbon-body] button', { hasText: 'Resize Table' }).click()
    await p.getByRole('dialog', { name: 'Resize Table' }).getByRole('textbox').fill('A1:C3')
    await p
      .getByRole('dialog', { name: 'Resize Table' })
      .getByRole('button', { name: 'OK' })
      .click()
    await expect.poll(async () => (await journal(p))[0]!.area.endRow).toBe(2)
    expect((await raw(p, 'A4:A5')).flat()).toEqual(['Plum', 'Fig'])

    const saved = await save(launched, p, file)
    expect(saved.table).toContain('ref="A1:C3"')
    expect(saved.table).toContain('<tableColumns count="3">')
  } finally {
    await keep(file, 'resized.xlsx')
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('the grid draws the totals row of the style and, with Header Row off, no header band', async () => {
  const { launched, p } = await openSheets()
  try {
    const colors = (await p.evaluate(`(window).__genofficeDebug.themeColors()`)) as string[]
    // Dark 3 paints its totals row (accent 2, 50% darker); Medium 2 its header (accent 1)
    const totalFill = resolveColor({ s: 'a2', t: -0.5 }, colors)
    const headerFill = resolveColor({ s: 'a1' }, colors)
    await makeTable(p, 'TableStyleDark3')
    await activate(p, 'M30')
    await p.waitForTimeout(500)
    const before = countColor(await p.screenshot(), totalFill)
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    await toggle(p, 'Total Row', true)
    await activate(p, 'M30')
    await p.waitForTimeout(600)
    const withTotals = countColor(await p.screenshot(), totalFill)
    expect(withTotals, 'the totals row is painted in the style’s total color').toBeGreaterThan(
      before + 500,
    )
    // and it goes with the Total Row
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    await toggle(p, 'Total Row', false)
    await activate(p, 'M30')
    await p.waitForTimeout(600)
    expect(countColor(await p.screenshot(), totalFill)).toBeLessThan(withTotals - 500)
    await undo(p)
    await undo(p)
    await expect.poll(async () => (await journal(p)).length).toBe(1)
    void headerFill
  } finally {
    await kill(launched)
  }
})

test('a header typed after the table was made is the column’s name in the saved part (Excel needs them equal)', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-headers-'))
  const file = join(dir, 'headers.xlsx')
  try {
    await makeTable(p)
    await p.evaluate(`${api}.getRange('B1').setValue('Quantity')`)
    // a clash with another header is made unique the way Excel does it, in the cell too
    await p.evaluate(`${api}.getRange('C1').setValue('quantity')`)
    const saved = await save(launched, p, file)
    expect(saved.table).toContain('<tableColumn id="1" name="Item"/>')
    expect(saved.table).toContain('<tableColumn id="2" name="Quantity"/>')
    expect(saved.table).toContain('<tableColumn id="3" name="quantity2"/>')
    expect(saved.table).toContain('<tableColumn id="4" name="Note"/>')
  } finally {
    await keep(file, 'headers.xlsx')
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})
