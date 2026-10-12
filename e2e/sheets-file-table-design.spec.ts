import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { resolveColor } from '../apps/sheets/src/renderer/table-styles'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Table Design on a table the file already has (the table of a saved workbook is the file's own):
 * the style gallery and the banding / column options change its look and the table part's
 * <tableStyleInfo>, one undo step each, also across a Save; the structural controls stay session-only.
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-file-table-design',
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
interface FileEdit {
  name: string
  style?: string
  bandedRows?: boolean
  bandedColumns?: boolean
  firstColumn?: boolean
  lastColumn?: boolean
}
const edits = (p: Page) =>
  p.evaluate(`(window).__genofficeDebug.fileTableEdits()`) as Promise<FileEdit[]>
const sessionTables = (p: Page) =>
  p.evaluate(`(window).__genofficeDebug.tableAdds()`) as Promise<{ name: string }[]>
const tab = (p: Page, name: string) =>
  p
    .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', {
      hasText: new RegExp(`^${name}$`),
    })
    .first()
    .click()
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(1200)
}
const redo = async (p: Page) => {
  await p.locator('button[aria-label^="Redo"]').first().click()
  await p.waitForTimeout(1200)
}
async function save(l: LaunchedApp, p: Page, path: string): Promise<JSZip> {
  await l.app.evaluate(({ dialog }, target) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
    dialog.showMessageBox = (async () => ({ response: 0, checkboxChecked: false })) as never
  }, path)
  await p.locator('button[aria-label^="Save As"]').first().click()
  await expect.poll(() => existsSync(path), { timeout: 20_000 }).toBe(true)
  await new Promise((r) => setTimeout(r, 1500))
  return JSZip.loadAsync(await readFile(path))
}
/** the colorful colors (not white / gray / black) that fill a few hundred pixels or more */
function palette(png: Buffer, min = 250): Map<string, number> {
  const image = PNG.sync.read(png)
  const counts = new Map<number, number>()
  for (let i = 0; i < image.data.length; i += 4) {
    const r = image.data[i]!
    const g = image.data[i + 1]!
    const b = image.data[i + 2]!
    if (Math.max(r, g, b) - Math.min(r, g, b) < 24) continue
    const key = (r << 16) | (g << 8) | b
    counts.set(key, (counts.get(key) ?? 0) + 1)
  }
  const out = new Map<string, number>()
  for (const [key, count] of counts) {
    if (count >= min) out.set(`#${key.toString(16).padStart(6, '0').toUpperCase()}`, count)
  }
  return out
}
/** the table drawn with the selection out of the way (a selected cell tints the cells under it) */
async function shot(p: Page): Promise<Map<string, number>> {
  await activate(p, 'M30')
  await p.waitForTimeout(900)
  const colors = palette(await p.screenshot())
  await activate(p, 'B3')
  await p.waitForTimeout(300)
  await tab(p, 'Table Design')
  return colors
}
const appeared = (before: Map<string, number>, after: Map<string, number>): string[] =>
  [...after.keys()].filter((key) => !before.has(key))
const gone = (before: Map<string, number>, after: Map<string, number>): string[] =>
  [...before.keys()].filter((key) => !after.has(key))
const tableXml = async (zip: JSZip): Promise<string> =>
  (await zip.file(/xl\/tables\/table\d+\.xml/)[0]!.async('string')).replace(/\s+/g, ' ')

/** a table of the file: made, saved, and reopened by the save */
async function fileTable(l: LaunchedApp, p: Page, path: string): Promise<void> {
  await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
  await activate(p, 'A1:D5')
  await tab(p, 'Home')
  await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
  await p.locator('.table-style-cell[data-style="TableStyleMedium2"]').click()
  await expect.poll(async () => (await sessionTables(p)).map((t) => t.name)).toEqual(['Table1'])
  await p.waitForTimeout(800)
  await save(l, p, path)
  // the reopened workbook's table is the file's: nothing of it is a session table any more
  await expect.poll(async () => (await sessionTables(p)).length).toBe(0)
  await p.waitForTimeout(1200)
  await activate(p, 'B3')
  await p.waitForTimeout(600)
  await tab(p, 'Table Design')
}
const pickStyle = async (p: Page, style: string) => {
  await p.locator('[data-ribbon-body] [aria-label="Table Styles"]').first().click()
  await p.locator(`.table-style-cell[data-style="${style}"]`).click()
}
const toggle = (p: Page, name: string) =>
  p.locator('[data-ribbon-body] button.check-item', { hasText: name }).click()

test('Table Design restyles a table of the file, undoably, and the save patches the table part', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-file-table-'))
  const first = join(dir, 'file-table-1.xlsx')
  const second = join(dir, 'file-table-2.xlsx')
  const third = join(dir, 'file-table-3.xlsx')
  try {
    await fileTable(launched, p, first)

    // the structural controls belong to session tables; the style options are open
    const ribbon = p.locator('[data-ribbon-body]')
    for (const name of ['Header Row', 'Total Row', 'Filter Button']) {
      await expect(ribbon.locator('button.check-item', { hasText: name })).toBeDisabled()
    }
    await expect(ribbon.locator('button.check-item', { hasText: 'Banded Rows' })).toBeEnabled()
    const nameField = ribbon.locator('input[aria-label="Table Name:"]')
    await expect(nameField).toBeDisabled()
    await expect(nameField).toHaveValue('Table1')
    await expect(ribbon.locator('button', { hasText: 'Resize Table' })).toBeDisabled()
    await expect(ribbon.locator('button', { hasText: 'Convert to Range' })).toBeDisabled()
    await expect(ribbon.locator('button.check-item', { hasText: 'Banded Rows' })).toHaveAttribute(
      'aria-checked',
      'true',
    )

    const original = await shot(p)
    expect(original.size).toBeGreaterThan(1)

    // a new style: the table is drawn in it at once, and the journal will patch the table part
    await pickStyle(p, 'TableStyleMedium3')
    await expect
      .poll(async () => (await edits(p)).map((e) => `${e.name}:${e.style}`))
      .toEqual(['Table1:TableStyleMedium3'])
    await expect(p.locator('.workbook-status')).toContainText('Table style changed')
    const restyled = await shot(p)
    // the header and the stripes of the old style are gone, those of the new one are there
    expect(gone(original, restyled).length).toBeGreaterThanOrEqual(2)
    expect(appeared(original, restyled).length).toBeGreaterThanOrEqual(2)

    // one undo brings the file's own look back, and takes the edit out of the journal
    await undo(p)
    await expect.poll(async () => (await edits(p)).length).toBe(0)
    const undone = await shot(p)
    expect(gone(original, undone)).toEqual([])
    expect(appeared(original, undone)).toEqual([])
    await redo(p)
    await expect.poll(async () => (await edits(p)).length).toBe(1)
    const redone = await shot(p)
    expect(gone(restyled, redone)).toEqual([])
    expect(appeared(restyled, redone)).toEqual([])

    // an option: no row stripes — the stripe color goes
    await toggle(p, 'Banded Rows')
    await expect.poll(async () => (await edits(p))[0]?.bandedRows).toBe(false)
    const plain = await shot(p)
    expect(gone(restyled, plain).length).toBeGreaterThanOrEqual(1)
    await toggle(p, 'Last Column')
    await expect.poll(async () => (await edits(p))[0]?.lastColumn).toBe(true)

    // the save patches only the style info of the table part
    const zip = await save(launched, p, second)
    const xml = await tableXml(zip)
    expect(xml).toMatch(
      /<tableStyleInfo name="TableStyleMedium3" showFirstColumn="0" showLastColumn="1" showRowStripes="0" showColumnStripes="0"\/>/,
    )
    expect(xml).toContain('<autoFilter')
    expect(xml).toContain('ref="A1:D5"')
    // the reopened workbook draws what the file says, which is what the session drew
    await expect.poll(async () => (await edits(p)).length).toBe(0)
    await p.waitForTimeout(1200)
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    await expect(
      p.locator('[data-ribbon-body] button.check-item', { hasText: 'Last Column' }),
    ).toHaveAttribute('aria-checked', 'true')
    await expect(
      p.locator('[data-ribbon-body] button.check-item', { hasText: 'Banded Rows' }),
    ).toHaveAttribute('aria-checked', 'false')
    const reopened = await shot(p)
    // the header color of the style, as the file's own look (resolved from the file) and as the
    // session drew it, is the same pixel color
    const header = [...restyled.keys()].filter((key) => !original.has(key))[0]!
    expect(reopened.has(header)).toBe(true)

    // undo reaches back across the save: the last two changes are taken back, and the next save
    // writes the table part as it was before them
    await undo(p)
    await undo(p)
    // (the saved file already says Medium 3: only the two options that differ from it are edits)
    await expect
      .poll(async () => (await edits(p)).map((e) => `${e.style}:${e.bandedRows}:${e.lastColumn}`))
      .toEqual(['undefined:true:false'])
    const zipAfter = await save(launched, p, third)
    expect(await tableXml(zipAfter)).toMatch(
      /<tableStyleInfo name="TableStyleMedium3" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"\/>/,
    )
  } finally {
    if (process.env.QA_KEEP_DIR) {
      for (const [file, name] of [
        [second, 'file-table-restyled.xlsx'],
        [third, 'file-table-undone.xlsx'],
      ] as const) {
        if (existsSync(file)) await copyFile(file, join(process.env.QA_KEEP_DIR, name))
      }
    }
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('a custom style can be given to a table of the file, and the save writes the style and names it', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-file-table2-'))
  const first = join(dir, 'file-table-a.xlsx')
  const second = join(dir, 'file-table-b.xlsx')
  try {
    await fileTable(launched, p, first)
    await p.locator('[data-ribbon-body] [aria-label="Table Styles"]').first().click()
    await p.locator('.table-style-gallery .table-style-new').click()
    const dialog = p.getByRole('dialog', { name: 'New Table Style' })
    await dialog.getByRole('textbox').first().fill('Ledger')
    await dialog.getByLabel('Choose color: Fill').fill('#1f3864')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(async () => (await edits(p))[0]?.style).toBe('Ledger')
    const colors = await shot(p)
    expect(colors.has('#1F3864')).toBe(true)

    const zip = await save(launched, p, second)
    const styles = await zip.file('xl/styles.xml')!.async('string')
    expect(styles).toMatch(/<tableStyle name="Ledger" pivot="0"/)
    expect(styles).toContain('<bgColor rgb="FF1F3864"/>')
    expect(await tableXml(zip)).toContain('<tableStyleInfo name="Ledger"')
  } finally {
    if (process.env.QA_KEEP_DIR && existsSync(second)) {
      await copyFile(second, join(process.env.QA_KEEP_DIR, 'file-table-custom.xlsx'))
    }
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

const checkbox = (p: Page, name: string) =>
  p.locator('[data-ribbon-body] [role="checkbox"]', { hasText: name })
const countColor = (png: Buffer, hex: string): number => {
  const image = PNG.sync.read(png)
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16))
  let count = 0
  for (let i = 0; i < image.data.length; i += 4) {
    if (image.data[i] === r && image.data[i + 1] === g && image.data[i + 2] === b) count += 1
  }
  return count
}

test('the totals row of a file table is drawn in the new style', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-file-table3-'))
  const totals = join(dir, 'file-table-totals.xlsx')
  const restyled = join(dir, 'file-table-totals-restyled.xlsx')
  try {
    await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
    await activate(p, 'A1:D5')
    await tab(p, 'Home')
    await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
    await p.locator('.table-style-cell[data-style="TableStyleMedium2"]').click()
    await expect.poll(async () => (await sessionTables(p)).length).toBe(1)
    await p.waitForTimeout(800)
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    await checkbox(p, 'Total Row').click()
    await expect(checkbox(p, 'Total Row')).toHaveAttribute('aria-checked', 'true')
    await p.waitForTimeout(900)
    const zip = await save(launched, p, totals)
    expect(await tableXml(zip)).toContain('totalsRowCount="1"')
    await expect.poll(async () => (await sessionTables(p)).length).toBe(0)
    await p.waitForTimeout(1200)
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    // the table of the file has its totals row: that option shows on, and is the file's to keep
    await expect(checkbox(p, 'Total Row')).toHaveAttribute('aria-checked', 'true')
    await expect(checkbox(p, 'Total Row')).toBeDisabled()

    // Dark 3 paints its totals row in accent 2, 50% darker
    const colors = (await p.evaluate(`(window).__genofficeDebug.themeColors()`)) as string[]
    const totalFill = resolveColor({ s: 'a2', t: -0.5 }, colors)
    await activate(p, 'M30')
    await p.waitForTimeout(700)
    const before = countColor(await p.screenshot(), totalFill)
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    await pickStyle(p, 'TableStyleDark3')
    await expect.poll(async () => (await edits(p))[0]?.style).toBe('TableStyleDark3')
    await activate(p, 'M30')
    await p.waitForTimeout(900)
    expect(countColor(await p.screenshot(), totalFill)).toBeGreaterThan(before + 500)

    const after = await save(launched, p, restyled)
    const xml = await tableXml(after)
    expect(xml).toContain('totalsRowCount="1"')
    expect(xml).toContain('<tableStyleInfo name="TableStyleDark3"')
  } finally {
    if (process.env.QA_KEEP_DIR && existsSync(restyled)) {
      await copyFile(restyled, join(process.env.QA_KEEP_DIR, 'file-table-totals.xlsx'))
    }
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Format as Table over a cell of a file table restyles that table, and undo takes it back', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-file-table4-'))
  try {
    await fileTable(launched, p, join(dir, 'file-table-home.xlsx'))
    await tab(p, 'Home')
    await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
    await p.locator('.table-style-cell[data-style="TableStyleLight9"]').click()
    await expect.poll(async () => (await edits(p))[0]?.style).toBe('TableStyleLight9')
    // no second table over the first: the table of the file is the only one
    await expect.poll(async () => (await sessionTables(p)).length).toBe(0)
    await undo(p)
    await expect.poll(async () => (await edits(p)).length).toBe(0)
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('a file table without a header row is restyled too, and keeps having none', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-file-table5-'))
  const headerless = join(dir, 'file-table-headerless.xlsx')
  const restyled = join(dir, 'file-table-headerless-restyled.xlsx')
  try {
    await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
    await activate(p, 'A1:D5')
    await tab(p, 'Home')
    await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
    await p.locator('.table-style-cell[data-style="TableStyleMedium2"]').click()
    await expect.poll(async () => (await sessionTables(p)).length).toBe(1)
    await p.waitForTimeout(800)
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    await checkbox(p, 'Header Row').click()
    await expect(checkbox(p, 'Header Row')).toHaveAttribute('aria-checked', 'false')
    await p.waitForTimeout(900)
    const zip = await save(launched, p, headerless)
    expect(await tableXml(zip)).toContain('headerRowCount="0"')
    await expect.poll(async () => (await sessionTables(p)).length).toBe(0)
    await p.waitForTimeout(1200)
    await activate(p, 'B3')
    await tab(p, 'Table Design')
    await expect(checkbox(p, 'Header Row')).toHaveAttribute('aria-checked', 'false')
    await expect(checkbox(p, 'Header Row')).toBeDisabled()

    const original = await shot(p)
    await pickStyle(p, 'TableStyleMedium3')
    await expect.poll(async () => (await edits(p))[0]?.style).toBe('TableStyleMedium3')
    const after = await shot(p)
    expect(gone(original, after).length).toBeGreaterThanOrEqual(1)
    expect(appeared(original, after).length).toBeGreaterThanOrEqual(1)
    await undo(p)
    await expect.poll(async () => (await edits(p)).length).toBe(0)
    await redo(p)
    await expect.poll(async () => (await edits(p)).length).toBe(1)

    const saved = await save(launched, p, restyled)
    const xml = await tableXml(saved)
    expect(xml).toContain('headerRowCount="0"')
    expect(xml).toContain('<tableStyleInfo name="TableStyleMedium3"')
  } finally {
    if (process.env.QA_KEEP_DIR && existsSync(restyled)) {
      await copyFile(restyled, join(process.env.QA_KEEP_DIR, 'file-table-headerless.xlsx'))
    }
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})
