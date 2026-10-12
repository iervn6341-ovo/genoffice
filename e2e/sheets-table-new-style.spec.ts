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

/**
 * Table Styles → New Table Style…: a style of the user's own (name, then fill / font / borders per
 * table element), kept in the workbook, drawn on the table, listed under Custom in the galleries and
 * saved as a <tableStyle> in styles.xml that real Excel reads back.
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-table-new-style',
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
const journal = (p: Page) =>
  p.evaluate(`(window).__genofficeDebug.tableAdds()`) as Promise<{ name: string; style?: string }[]>
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
async function makeTable(p: Page): Promise<void> {
  await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
  await activate(p, 'A1:D5')
  await tab(p, 'Home')
  await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
  await p.locator('.table-style-cell[data-style="TableStyleMedium2"]').click()
  await expect.poll(async () => (await journal(p)).map((t) => t.name)).toEqual(['Table1'])
  await p.waitForTimeout(800)
  await activate(p, 'B3')
}
async function save(l: LaunchedApp, p: Page, path: string): Promise<JSZip> {
  await l.app.evaluate(({ dialog }, target) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
    dialog.showMessageBox = (async () => ({ response: 0, checkboxChecked: false })) as never
  }, path)
  await p.locator('button[aria-label^="Save As"]').first().click()
  await expect.poll(() => existsSync(path), { timeout: 20_000 }).toBe(true)
  await new Promise((r) => setTimeout(r, 900))
  return JSZip.loadAsync(await readFile(path))
}
function countColor(png: Buffer, hex: string): number {
  const image = PNG.sync.read(png)
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16))
  let count = 0
  for (let i = 0; i < image.data.length; i += 4) {
    if (image.data[i] === r && image.data[i + 1] === g && image.data[i + 2] === b) count += 1
  }
  return count
}
const openNewStyle = async (p: Page) => {
  await p.locator('[data-ribbon-body] [aria-label="Table Styles"]').first().click()
  await p.locator('.table-style-gallery .table-style-new').click()
  const dialog = p.getByRole('dialog', { name: 'New Table Style' })
  await expect(dialog).toBeVisible()
  return dialog
}

test('New Table Style makes a style the table wears, the galleries list, and the saved file names', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-new-style-'))
  const file = join(dir, 'custom-style.xlsx')
  try {
    await makeTable(p)
    await tab(p, 'Table Design')
    let dialog = await openNewStyle(p)
    // a default name, the elements of the table, and a preview that draws the draft
    await expect(dialog.getByRole('textbox').first()).toHaveValue('Table Style 1')
    const elements = dialog.getByRole('listbox', { name: 'Table element' })
    await expect(elements.getByRole('option')).toHaveCount(9)
    await expect(dialog.locator('.new-style-preview i')).toHaveCount(30)

    await dialog.getByRole('textbox').first().fill('Ledger')
    // Header Row is selected: a dark fill, white bold text
    await dialog.getByLabel('Choose color: Fill').fill('#1f3864')
    // the first row stripe: a light green
    await elements.getByRole('option', { name: 'First Row Stripe' }).click()
    await dialog.getByLabel('Choose color: Fill').fill('#e2efda')
    // the whole table: thin gray lines between the rows
    await elements.getByRole('option', { name: 'Whole Table' }).click()
    await dialog.getByRole('combobox', { name: 'Border' }).selectOption('thin')
    await dialog.getByRole('group', { name: 'Border' }).getByLabel('Inside horizontal').check()
    await dialog.getByRole('button', { name: 'OK' }).click()

    // the table under the cell wears it at once, as one undo step
    await expect.poll(async () => (await journal(p))[0]!.style).toBe('Ledger')
    await expect(p.locator('.workbook-status')).toContainText('Table style changed')
    await activate(p, 'M30')
    await p.waitForTimeout(600)
    expect(countColor(await p.screenshot(), '#1F3864')).toBeGreaterThan(300)
    expect(countColor(await p.screenshot(), '#E2EFDA')).toBeGreaterThan(300)
    await activate(p, 'B3')
    await undo(p)
    await expect.poll(async () => (await journal(p))[0]!.style).toBe('TableStyleMedium2')

    // the galleries list it under Custom
    await tab(p, 'Table Design')
    await p.locator('[data-ribbon-body] [aria-label="Table Styles"]').first().click()
    await expect(p.locator('.table-style-gallery .rb-drop-title')).toHaveText([
      'Light',
      'Medium',
      'Dark',
      'Custom',
    ])
    await expect(p.locator('.table-style-cell[data-style="Ledger"]')).toHaveCount(1)
    await p.locator('.table-style-cell[data-style="Ledger"]').click()
    await expect.poll(async () => (await journal(p))[0]!.style).toBe('Ledger')

    // a name Excel has, or one already made, is refused in the dialog
    dialog = await openNewStyle(p)
    await dialog.getByRole('textbox').first().fill('TableStyleMedium2')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog.getByRole('alert')).toContainText("Excel's own")
    await dialog.getByRole('textbox').first().fill('ledger')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog.getByRole('alert')).toContainText('already exists')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    // the saved file holds the style, and the table names it
    const zip = await save(launched, p, file)
    const styles = await zip.file('xl/styles.xml')!.async('string')
    expect(styles).toMatch(/<tableStyle name="Ledger" pivot="0" count="3">/)
    expect(styles).toContain('<bgColor rgb="FF1F3864"/>')
    expect(styles).toContain('<bgColor rgb="FFE2EFDA"/>')
    const table = (await zip.file(/xl\/tables\/table\d+\.xml/)[0]!.async('string')).replace(
      /\s+/g,
      ' ',
    )
    expect(table).toContain('<tableStyleInfo name="Ledger"')
  } finally {
    if (process.env.QA_KEEP_DIR && existsSync(file)) {
      await copyFile(file, join(process.env.QA_KEEP_DIR, 'custom-style.xlsx'))
    }
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('a custom style is offered by Format as Table, and survives a save in the same workbook', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-new-style2-'))
  const file = join(dir, 'custom-style2.xlsx')
  try {
    await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
    await activate(p, 'A1:D5')
    await tab(p, 'Home')
    await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
    await p.locator('.table-style-gallery .table-style-new').click()
    const dialog = p.getByRole('dialog', { name: 'New Table Style' })
    await dialog.getByRole('textbox').first().fill('Mine')
    await dialog.getByRole('button', { name: 'OK' }).click()
    // with no table under the cell, the selection becomes one, in the new style
    await expect.poll(async () => (await journal(p))[0]?.style).toBe('Mine')

    await save(launched, p, file)
    // the reopened workbook is the same file: the style is still there to pick
    await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
    await expect(p.locator('.table-style-cell[data-style="Mine"]')).toHaveCount(1)
  } finally {
    if (process.env.QA_KEEP_DIR && existsSync(file)) {
      await copyFile(file, join(process.env.QA_KEEP_DIR, 'custom-style2.xlsx'))
    }
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})
