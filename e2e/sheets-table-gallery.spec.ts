import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { tableThemeFor, type ThemeColors } from '../apps/sheets/src/renderer/table-styles'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Format as Table is Excel's gallery of 60 built-in styles, drawn in the workbook's own theme
 * colors; the picked style is what the grid shows and what the saved table names. The colors the
 * grid has to show are the ones Excel resolved for the same styles (Medium 2 under Office
 * 2013–2022: header 4472C4, stripe D9E1F2).
 */
async function openSheets(width = 2400): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-table-gallery',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.page.locator('.quick-card').nth(1).click()
  const p = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await p.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  await p.waitForTimeout(1500)
  await setEditorLayoutWidth(launched.app, '://sheets/', width)
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
async function seed(p: Page): Promise<void> {
  await p.evaluate(`${api}.getRange('A1:D5').setValues(${JSON.stringify(ROWS)})`)
  await activate(p, 'A1:D5')
}
const openGallery = async (p: Page) => {
  await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
  const gallery = p.locator('.table-style-gallery')
  await expect(gallery).toBeVisible()
  return gallery
}
const themeColors = (p: Page) =>
  p.evaluate(() =>
    (
      window as unknown as { __genofficeDebug: { themeColors(): string[] | null } }
    ).__genofficeDebug.themeColors(),
  )

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

test("Format as Table is a gallery of Excel's 60 styles that fits the window", async () => {
  const { launched, p } = await openSheets(1280)
  try {
    await seed(p)
    const gallery = await openGallery(p)
    await expect(gallery.locator('.table-style-cell')).toHaveCount(60)
    if (process.env.QA_KEEP_DIR)
      await gallery.screenshot({ path: join(process.env.QA_KEEP_DIR, 'gallery.png') })
    await expect(gallery.locator('.rb-drop-title')).toHaveText(['Light', 'Medium', 'Dark'])
    for (const [family, count] of [
      ['Light', 21],
      ['Medium', 28],
      ['Dark', 11],
    ] as const) {
      await expect(
        gallery.locator(`.table-style-cell[data-style^="TableStyle${family}"]`),
      ).toHaveCount(count)
    }
    // tooltips and labels follow Excel's: "Light 1" … "Dark 11"
    await expect(gallery.locator('.table-style-cell').first()).toHaveAttribute(
      'aria-label',
      'Light 1',
    )
    await expect(gallery.locator('.table-style-cell').nth(21)).toHaveAttribute(
      'aria-label',
      'Medium 1',
    )
    await expect(gallery.locator('.table-style-cell').last()).toHaveAttribute(
      'aria-label',
      'Dark 11',
    )
    // no sideways scroll, and inside the window (scrolls vertically when it is taller)
    const fit = await gallery.evaluate((el) => {
      const rect = el.getBoundingClientRect()
      return {
        sideways: el.scrollWidth - el.clientWidth,
        inside:
          rect.left >= 0 && rect.right <= window.innerWidth && rect.bottom <= window.innerHeight,
      }
    })
    expect(fit.sideways).toBeLessThanOrEqual(1)
    expect(fit.inside).toBe(true)
    // the thumbnails are drawn tables, not blank boxes
    const drawn = await gallery
      .locator('.table-style-preview i')
      .evaluateAll(
        (cells) => cells.filter((c) => (c as HTMLElement).style.background !== '').length,
      )
    expect(drawn).toBeGreaterThan(60 * 4)
    await p.keyboard.press('Escape')
    await expect(gallery).toHaveCount(0)
  } finally {
    await kill(launched)
  }
})

test("the grid shows the colors Excel resolved for the picked style, in the workbook's theme", async () => {
  const { launched, p } = await openSheets()
  try {
    await seed(p)
    const colors = (await themeColors(p)) as ThemeColors | null
    expect(colors, 'a new workbook carries a theme').not.toBeNull()
    const clip = { x: 0, y: 60, width: 1280, height: 420 }

    await openGallery(p)
    await p.locator('.table-style-cell[data-style="TableStyleMedium2"]').click()
    await p.waitForTimeout(1800)
    // the selection's blue wash would tint the cells: look at the table unselected
    await activate(p, 'M20')
    await p.waitForTimeout(400)
    const medium2 = tableThemeFor('TableStyleMedium2', 'x', colors!)!
    const shot = await p.screenshot({ clip })
    if (process.env.QA_KEEP_DIR)
      await writeFile(join(process.env.QA_KEEP_DIR, 'table-medium2.png'), shot)
    // header and the stripe rows are drawn flat in Excel's resolved colors
    expect(countColor(shot, medium2.headerRowStyle!.bg!.rgb)).toBeGreaterThan(100)
    expect(countColor(shot, medium2.firstRowStyle!.bg!.rgb)).toBeGreaterThan(100)
    if (colors![4]!.toUpperCase() === '#4472C4') {
      // Office 2013–2022 (read from Excel): Accent 1 lighter 80% = D9E1F2
      expect(medium2.headerRowStyle!.bg!.rgb).toBe('#4472C4')
      expect(medium2.firstRowStyle!.bg!.rgb).toBe('#D9E1F2')
    }

    // restyled in place: a Dark style swaps the look, a style outside the old six works too
    await activate(p, 'A1:D5')
    await openGallery(p)
    await p.locator('.table-style-cell[data-style="TableStyleDark11"]').click()
    await p.waitForTimeout(1800)
    await activate(p, 'M20')
    await p.waitForTimeout(400)
    const dark11 = tableThemeFor('TableStyleDark11', 'x', colors!)!
    const shot2 = await p.screenshot({ clip })
    if (process.env.QA_KEEP_DIR)
      await writeFile(join(process.env.QA_KEEP_DIR, 'table-dark11.png'), shot2)
    expect(countColor(shot2, dark11.headerRowStyle!.bg!.rgb)).toBeGreaterThan(100)
    expect(countColor(shot2, medium2.headerRowStyle!.bg!.rgb)).toBeLessThan(
      countColor(shot, medium2.headerRowStyle!.bg!.rgb) / 4,
    )
    expect(await p.evaluate(`${api}.getSubTableInfos().length`)).toBe(1)
  } finally {
    await kill(launched)
  }
})

test('the saved file names the picked style, and Excel reads it', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-table-style-'))
  const file = join(dir, 'table.xlsx')
  try {
    await seed(p)
    await openGallery(p)
    await p.locator('.table-style-cell[data-style="TableStyleMedium14"]').click()
    await p.waitForTimeout(1800)
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
    }, file)
    await p.locator('button[aria-label^="Save As"]').first().click()
    await expect.poll(() => existsSync(file), { timeout: 20_000 }).toBe(true)
    await new Promise((r) => setTimeout(r, 800))
    const zip = await JSZip.loadAsync(await readFile(file))
    const tableXml = await zip.file(/xl\/tables\/table\d+\.xml/)[0]!.async('string')
    expect(tableXml).toContain('name="TableStyleMedium14"')
    if (process.env.QA_KEEP_DIR) await copyFile(file, join(process.env.QA_KEEP_DIR, 'table.xlsx'))
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})
