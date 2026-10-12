import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * The contextual Table Design tab for a table made this session: Table Name, Table Style Options,
 * the style gallery and Convert to Range. Each change is one ⌘Z step, the table the grid draws
 * and the table the saved file names stay one thing, and Excel reads the result.
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-table-design',
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
const tableNames = (p: Page) =>
  p.evaluate(`${api}.getSubTableInfos().map((t) => t.name)`) as Promise<string[]>
const tab = (p: Page, name: string) =>
  p
    .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', {
      hasText: new RegExp(`^${name}$`),
    })
    .first()
    .click()
const tabButton = (p: Page, name: string) =>
  p.locator('nav.ribbon-tabs > button', { hasText: new RegExp(`^${name}$`) })
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(800)
}
const redo = async (p: Page) => {
  await p.locator('button[aria-label^="Redo"]').first().click()
  await p.waitForTimeout(800)
}
async function makeTable(p: Page, style = 'TableStyleMedium2'): Promise<void> {
  await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
  await activate(p, 'A1:D5')
  await tab(p, 'Home')
  await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
  await p.locator(`.table-style-cell[data-style="${style}"]`).click()
  await expect.poll(() => tableNames(p)).toEqual(['Table1'])
  await p.waitForTimeout(800)
}
const checkbox = (p: Page, name: string) =>
  p.locator('[data-ribbon-body] [role="checkbox"]', { hasText: name })

async function save(l: LaunchedApp, p: Page, path: string, first = false): Promise<string> {
  const before = existsSync(path) ? (await readFile(path)).length : -1
  if (first) {
    await l.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
    }, path)
    await p.locator('button[aria-label^="Save As"]').first().click()
  } else {
    await p.locator('button[aria-label^="Save ("]').first().click()
  }
  await expect
    .poll(async () => (existsSync(path) ? (await readFile(path)).length : -1), { timeout: 20_000 })
    .not.toBe(before)
  await new Promise((r) => setTimeout(r, 700))
  const zip = await JSZip.loadAsync(await readFile(path))
  const part = zip.file(/xl\/tables\/table\d+\.xml/)[0]
  return part ? (await part.async('string')).replace(/\s+/g, ' ') : ''
}
const styleInfo = (xml: string): string => /<tableStyleInfo [^>]*\/>/.exec(xml)?.[0] ?? ''

test('Table Design is contextual: it comes with a table of this session and goes without one', async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
    await activate(p, 'A1:D5')
    await expect(tabButton(p, 'Table Design')).toHaveCount(0)
    await makeTable(p)
    await expect(tabButton(p, 'Table Design')).toHaveCount(1)
    await expect(tabButton(p, 'Table Design')).toHaveClass(/contextual/)
    // out of the table the tab goes (it lingers a moment, in case a change is re-creating the table)
    await activate(p, 'M20')
    await expect(tabButton(p, 'Table Design')).toHaveCount(0)
    await activate(p, 'B3')
    await expect(tabButton(p, 'Table Design')).toHaveCount(1)
    await tab(p, 'Table Design')
    await expect(p.locator('[data-ribbon-body] input[aria-label="Table Name:"]')).toHaveValue(
      'Table1',
    )
    await expect(p.locator('[data-ribbon-body] [role="checkbox"]')).toHaveCount(7)
    await expect(checkbox(p, 'Banded Rows')).toHaveAttribute('aria-checked', 'true')
    await expect(checkbox(p, 'Banded Columns')).toHaveAttribute('aria-checked', 'false')
  } finally {
    await kill(launched)
  }
})

test('Table Name renames in one step and refuses what Excel refuses', async () => {
  const { launched, p } = await openSheets()
  try {
    await makeTable(p)
    await tab(p, 'Table Design')
    const name = p.locator('[data-ribbon-body] input[aria-label="Table Name:"]')
    await name.fill('Sales_2024')
    await name.press('Enter')
    await expect.poll(() => tableNames(p)).toEqual(['Sales_2024'])
    await expect(p.locator('.workbook-status')).toContainText('renamed')
    await expect(name).toHaveValue('Sales_2024')
    await undo(p)
    expect(await tableNames(p)).toEqual(['Table1'])
    await redo(p)
    expect(await tableNames(p)).toEqual(['Sales_2024'])

    // a name with a space, one that looks like a cell reference: refused, the table keeps its name
    for (const bad of ['my table', 'A1', '1abc']) {
      await name.fill(bad)
      await name.press('Enter')
      await expect(p.locator('.workbook-status')).toContainText('not a valid table name')
      expect(await tableNames(p)).toEqual(['Sales_2024'])
      await expect(name).toHaveValue('Sales_2024')
    }

    // a second table's name is taken (any case)
    await p.evaluate(`${api}.getRange('G1:H3').setValues([['x','y'],[1,2],[3,4]])`)
    await activate(p, 'G1:H3')
    await tab(p, 'Home')
    await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
    await p.locator('.table-style-cell[data-style="TableStyleLight9"]').click()
    await expect.poll(async () => (await tableNames(p)).length).toBe(2)
    await tab(p, 'Table Design')
    const second = await tableNames(p)
    const other = second.find((n) => n !== 'Sales_2024')!
    await name.fill('SALES_2024')
    await name.press('Enter')
    await expect(p.locator('.workbook-status')).toContainText('already used')
    expect(await tableNames(p)).toEqual(second)
    expect(other).toMatch(/^Table/)
  } finally {
    await kill(launched)
  }
})

// A saved table is the file's own, which Table Design does not change (as Format as Table):
// each scenario below changes the table in the session, then saves once and reads the file.
test('Table Style Options and the style gallery land in the saved tableStyleInfo, and the style keeps the options', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-design-'))
  const file = join(dir, 'design.xlsx')
  try {
    await makeTable(p)
    await tab(p, 'Table Design')
    await checkbox(p, 'First Column').click()
    await expect(checkbox(p, 'First Column')).toHaveAttribute('aria-checked', 'true')
    await checkbox(p, 'Banded Columns').click()
    await expect(checkbox(p, 'Banded Columns')).toHaveAttribute('aria-checked', 'true')
    await checkbox(p, 'Banded Rows').click()
    await expect(checkbox(p, 'Banded Rows')).toHaveAttribute('aria-checked', 'false')
    expect(await tableNames(p)).toEqual(['Table1'])
    // a style picked in the Table Design gallery keeps the options
    await p.locator('[data-ribbon-body] .cover-select[aria-label="Table Styles"]').click()
    await p.locator('.table-style-cell[data-style="TableStyleDark9"]').click()
    await expect.poll(() => tableNames(p)).toEqual(['Table1'])
    await p.waitForTimeout(600)
    await expect(checkbox(p, 'First Column')).toHaveAttribute('aria-checked', 'true')
    await expect(checkbox(p, 'Banded Rows')).toHaveAttribute('aria-checked', 'false')
    const xml = await save(launched, p, file, true)
    expect(styleInfo(xml)).toBe(
      '<tableStyleInfo name="TableStyleDark9" showFirstColumn="1" showLastColumn="0" showRowStripes="0" showColumnStripes="1"/>',
    )
    if (process.env.QA_KEEP_DIR) await copyFile(file, join(process.env.QA_KEEP_DIR, 'design.xlsx'))
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('every Table Design change is one ⌘Z step, and ⇧⌘Z puts it back', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-undo-'))
  const file = join(dir, 'undo.xlsx')
  try {
    await makeTable(p)
    await tab(p, 'Table Design')
    await checkbox(p, 'Last Column').click()
    await expect(checkbox(p, 'Last Column')).toHaveAttribute('aria-checked', 'true')
    await checkbox(p, 'Banded Columns').click()
    await expect(checkbox(p, 'Banded Columns')).toHaveAttribute('aria-checked', 'true')
    await p.locator('[data-ribbon-body] .cover-select[aria-label="Table Styles"]').click()
    await p.locator('.table-style-cell[data-style="TableStyleLight15"]').click()
    await p.waitForTimeout(900)

    await undo(p) // the style
    await expect(checkbox(p, 'Last Column')).toHaveAttribute('aria-checked', 'true')
    await expect(checkbox(p, 'Banded Columns')).toHaveAttribute('aria-checked', 'true')
    await undo(p) // Banded Columns
    await expect(checkbox(p, 'Banded Columns')).toHaveAttribute('aria-checked', 'false')
    await expect(checkbox(p, 'Last Column')).toHaveAttribute('aria-checked', 'true')
    await redo(p)
    await expect(checkbox(p, 'Banded Columns')).toHaveAttribute('aria-checked', 'true')
    await undo(p)
    await undo(p) // Last Column
    await expect(checkbox(p, 'Last Column')).toHaveAttribute('aria-checked', 'false')
    expect(await tableNames(p)).toEqual(['Table1'])
    const xml = await save(launched, p, file, true)
    expect(styleInfo(xml)).toBe(
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>',
    )
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Convert to Range removes the table and keeps its look in the cells; ⌘Z brings the table back', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-range-'))
  const file = join(dir, 'range.xlsx')
  try {
    await makeTable(p, 'TableStyleMedium2')
    await tab(p, 'Table Design')
    await p.locator('[data-ribbon-body] button', { hasText: 'Convert to Range' }).click()
    await expect.poll(() => tableNames(p)).toEqual([])
    await expect(p.locator('.workbook-status')).toContainText('converted')
    await expect(tabButton(p, 'Table Design')).toHaveCount(0)
    // the header keeps its fill, bold and white text; the cells keep their values
    const header = (await p.evaluate(`${api}.getRange('A1').getCellStyleData()`)) as {
      bg?: { rgb?: string }
      bl?: number
      cl?: { rgb?: string }
    }
    expect(header.bg?.rgb).toMatch(/^#[0-9A-F]{6}$/i)
    expect(header.bl).toBe(1)
    expect(header.cl?.rgb?.toUpperCase()).toBe('#FFFFFF')
    const stripe = (await p.evaluate(`${api}.getRange('A2').getCellStyleData()`)) as {
      bg?: { rgb?: string }
    }
    expect(stripe.bg?.rgb).toMatch(/^#[0-9A-F]{6}$/i)
    expect(await p.evaluate(`${api}.getRange('A1:D2').getRawValues()`)).toEqual([
      ['Item', 'Qty', 'Price', 'Note'],
      ['Apple', 3, 1.5, 'a'],
    ])
    const xml = await save(launched, p, file, true)
    expect(xml, 'the saved file has no table part').toBe('')
    if (process.env.QA_KEEP_DIR) await copyFile(file, join(process.env.QA_KEEP_DIR, 'range.xlsx'))

    await undo(p)
    await expect.poll(() => tableNames(p)).toEqual(['Table1'])
    const restored = (await p.evaluate(`${api}.getRange('A1').getCellStyleData()`)) as {
      bg?: { rgb?: string }
    } | null
    expect(restored?.bg?.rgb).toBeUndefined()
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Esc cancels a typed name, a sheet name is refused, and the gallery restyles next to other data', async () => {
  const { launched, p } = await openSheets()
  try {
    await makeTable(p)
    await tab(p, 'Table Design')
    const name = p.locator('[data-ribbon-body] input[aria-label="Table Name:"]')

    // Esc puts the old name back instead of renaming
    await name.fill('Typed')
    await name.press('Escape')
    await expect(name).toHaveValue('Table1')
    expect(await tableNames(p)).toEqual(['Table1'])

    // a table cannot take the name of a sheet (Univer would refuse it and leave the journal behind)
    await p.evaluate(`${api}.setName('Budget')`)
    await name.fill('Budget')
    await name.press('Enter')
    await expect(p.locator('.workbook-status')).toContainText('name of a sheet')
    expect(await tableNames(p)).toEqual(['Table1'])

    // data right beside the table must not turn a restyle into an "overlaps" refusal
    await p.evaluate(`${api}.getRange('E1:E5').setValues([['n'],['1'],['2'],['3'],['4']])`)
    await activate(p, 'B3')
    await p.locator('[data-ribbon-body] .cover-select[aria-label="Table Styles"]').click()
    await p.locator('.table-style-cell[data-style="TableStyleDark5"]').click()
    await expect.poll(() => tableNames(p)).toEqual(['Table1'])
    await p.waitForTimeout(600)
    await expect(p.locator('.workbook-status')).not.toContainText('overlap')
  } finally {
    await kill(launched)
  }
})

test("Convert to Range keeps a cell's own fill above the style", async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
    await p.evaluate(`${api}.getRange('A2').setBackground('#FFFF00')`)
    await activate(p, 'A1:D5')
    await tab(p, 'Home')
    await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
    await p.locator('.table-style-cell[data-style="TableStyleMedium2"]').click()
    await expect.poll(() => tableNames(p)).toEqual(['Table1'])
    await p.waitForTimeout(600)
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    await p.locator('[data-ribbon-body] button', { hasText: 'Convert to Range' }).click()
    await expect.poll(() => tableNames(p)).toEqual([])
    const fill = async (a1: string) =>
      (
        (await p.evaluate(`${api}.getRange('${a1}').getCellStyleData()`)) as {
          bg?: { rgb?: string }
        } | null
      )?.bg?.rgb?.toUpperCase()
    expect(await fill('A2')).toBe('#FFFF00') // its own fill, not the stripe
    expect(await fill('B2')).toMatch(/^#[0-9A-F]{6}$/) // the neighbor takes the style's stripe
    expect(await fill('B2')).not.toBe('#FFFF00')
    expect(await fill('A1')).toMatch(/^#[0-9A-F]{6}$/) // the header
  } finally {
    await kill(launched)
  }
})
