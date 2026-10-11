import { test, expect, type Page } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Table Design / Table Layout tabs of a selected table, laid out like PowerPoint's: the six style
 * options, the built-in style gallery, borders, and row / column / size commands.
 */
async function newDeckWithTable(app: Awaited<ReturnType<typeof launchShell>>): Promise<Page> {
  await app.page.locator('.quick-card').nth(2).click()
  const s = await waitForPageWithUrl(app.app, '://slides/', 30_000)
  await s.waitForSelector('.ribbon', { timeout: 20_000 })
  await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
  await s.locator('.ribbon-body button[data-tip*="Insert a table"]').first().click()
  // the 2nd × 2nd cell of the picker
  const cells = s.locator('.rb-drop button, .rb-drop [role="gridcell"], .rb-table-picker-cell')
  await cells
    .nth(11)
    .click({ force: true })
    .catch(() => undefined)
  return s
}

const heightCm = async (s: Page): Promise<number> =>
  Number(
    await s
      .locator('.rb-tbl-field', { hasText: /^Height/ })
      .locator('input')
      .inputValue(),
  )

test('table tabs: options, gallery, layout commands', async () => {
  const app = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'slides-table-tabs' })
  try {
    const s = await newDeckWithTable(app)
    // inserting selects the table and reveals both contextual tabs
    await expect(s.locator('.ribbon-tab-context')).toHaveCount(2)
    await expect(s.locator('.ribbon-tab-context').first()).toHaveText('Table Design')

    // the six style options, PowerPoint's wording
    for (const name of [
      'Header Row',
      'Total Row',
      'Banded Rows',
      'First Column',
      'Last Column',
      'Banded Columns',
    ]) {
      await expect(s.locator('.rb-tbl-check', { hasText: name })).toBeVisible()
    }
    const firstCol = s.locator('.rb-tbl-check', { hasText: 'First Column' }).locator('input')
    await expect(firstCol).not.toBeChecked()
    await firstCol.click()
    await expect(firstCol).toBeChecked()
    // the gallery shows built-in styles and a pick sticks
    const thumbs = s.locator('.rb-tbl-strip .rb-tbl-thumb')
    await expect(thumbs).toHaveCount(7)
    await thumbs.nth(2).click()
    await expect(s.locator('.rb-tbl-strip .rb-tbl-thumb.selected')).toHaveCount(1)
    // the options survive a style pick
    await expect(
      s.locator('.rb-tbl-check', { hasText: 'First Column' }).locator('input'),
    ).toBeChecked()

    // Table Layout: inserting a row grows the table, distributing keeps its height
    await s.locator('.ribbon-tab-context').nth(1).click()
    const before = await heightCm(s)
    await s.locator('.rb-tbl-small[data-tip="Insert Row Below"]').click()
    await expect.poll(() => heightCm(s)).toBeGreaterThan(before)
    const grown = await heightCm(s)
    await s.locator('.rb-tbl-small', { hasText: 'Distribute Rows' }).click()
    await expect.poll(async () => Math.abs((await heightCm(s)) - grown)).toBeLessThan(0.1)
    // typing a size commits on Enter and moves the table's height
    const rowField = s.locator('.rb-tbl-field', { hasText: 'Row Height' }).locator('input')
    await rowField.fill('2')
    await rowField.press('Enter')
    await expect.poll(() => heightCm(s)).toBeGreaterThan(grown)
  } finally {
    await closeAndSaveVideo(app, 'cleanup')
  }
})
