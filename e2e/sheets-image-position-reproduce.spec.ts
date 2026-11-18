import { test, expect } from '@playwright/test'
import { copyFile, mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type { Page } from '@playwright/test'
import {
  launchShell,
  closeAndSaveVideo,
  waitForPageWithUrl,
  screenshotPath,
} from './helpers'

const FIXTURE = resolve(__dirname, '../apps/sheets/fixtures/generated/compatibility-basic.xlsx')

async function waitForWorkbook(page: Page): Promise<void> {
  await page.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await page.waitForTimeout(1_500)
}

test.describe('sheets: Insert image, selection and column insert drift', () => {
  test('reproduces image selection and column insert behavior', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-image-test-'))
    const workbook = join(scratch, 'test-image.xlsx')
    await copyFile(FIXTURE, workbook)

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-image-test',
      openFile: workbook,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      await waitForWorkbook(sheets)

      sheets.on('console', (msg) => {
        console.log(`[SHEETS CONSOLE] ${msg.type()}: ${msg.text()}`)
      })

      // Stub screen capture
      await launched.app.evaluate(({ desktopCapturer, nativeImage, systemPreferences }) => {
        const pixel =
          'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
        const frame = nativeImage.createFromDataURL(pixel).resize({ width: 200, height: 150 })
        desktopCapturer.getSources = async () => [
          { id: 'screen:0:0', name: 'Mock Screen', thumbnail: frame } as never,
        ]
        systemPreferences.getMediaAccessStatus = () => 'granted'
      })

      // Select cell C3
      await sheets.evaluate(() => {
        const api = (window as any).__genofficeDebug?.univerAPI
        api?.getActiveWorkbook().getActiveSheet().getRange(2, 2).activate()
      })
      await sheets.waitForTimeout(500)

      // Insert screenshot
      await sheets.getByRole('button', { name: 'Insert', exact: true }).click()
      await sheets.getByRole('button', { name: 'Screenshot' }).click()
      const dialog = sheets.getByRole('dialog', { name: 'Screenshot' })
      await expect(dialog).toBeVisible()
      const grid = dialog.locator('.screenshot-grid button')
      await grid.first().click()
      await expect(sheets.locator('.xlsx-image')).toBeVisible({ timeout: 15_000 })

      // Check initial bounding box of .shape-editable and .xlsx-image
      const initialBox = await sheets.locator('.xlsx-image').boundingBox()
      console.log('Initial image bounding box:', initialBox)

      const shapeEditable = sheets.locator('.shape-editable')
      const isSelectedInitially = await shapeEditable.evaluate((el) => el.classList.contains('selected'))
      console.log('Is selected initially?', isSelectedInitially)

      // Try clicking on the image
      console.log('Clicking on the image...')
      await sheets.locator('.xlsx-image').click()
      await sheets.waitForTimeout(500)
      const isSelectedAfterClick = await shapeEditable.evaluate((el) => el.classList.contains('selected'))
      console.log('Is selected after click?', isSelectedAfterClick)

      // Capture screenshot
      await sheets.screenshot({ path: screenshotPath('image-after-click') })

      // Now select column B and insert a column before it
      console.log('Inserting a column before B (to the left of image)...')
      await sheets.evaluate(async () => {
        const api = (window as any).__genofficeDebug?.univerAPI
        // Set selection to column B (index 1)
        api.getActiveWorkbook().getActiveSheet().getRange(0, 1, 10, 1).activate()
        // Execute insert col before
        await api.executeCommand('sheet.command.insert-col-before', { value: 1 })
      })
      await sheets.waitForTimeout(1000)

      const boxAfterLeftCol = await sheets.locator('.xlsx-image').boundingBox()
      console.log('Image bounding box after inserting column to left:', boxAfterLeftCol)
      await sheets.screenshot({ path: screenshotPath('image-after-left-col') })

      // Now select column I (index 8, well to the right of image) and insert a column before it
      console.log('Inserting a column to the right of image (column index 8)...')
      await sheets.evaluate(async () => {
        const api = (window as any).__genofficeDebug?.univerAPI
        api.getActiveWorkbook().getActiveSheet().getRange(0, 8, 10, 1).activate()
        await api.executeCommand('sheet.command.insert-col-before', { value: 1 })
      })
      await sheets.waitForTimeout(1000)

      const boxAfterRightCol = await sheets.locator('.xlsx-image').boundingBox()
      console.log('Image bounding box after inserting column to right:', boxAfterRightCol)
      await sheets.screenshot({ path: screenshotPath('image-after-right-col') })

      // The image must NOT expand when a column is inserted to its right!
      expect(Math.round(boxAfterRightCol?.width ?? 0)).toBe(Math.round(boxAfterLeftCol?.width ?? 0))
      expect(Math.round(boxAfterRightCol?.x ?? 0)).toBe(Math.round(boxAfterLeftCol?.x ?? 0))

      // Now insert a column inside the image (e.g. index 4, which is between fromColumn 3 and toColumn 5)
      console.log('Inserting a column inside the image (column index 4)...')
      await sheets.evaluate(async () => {
        const api = (window as any).__genofficeDebug?.univerAPI
        api.getActiveWorkbook().getActiveSheet().getRange(0, 4, 10, 1).activate()
        await api.executeCommand('sheet.command.insert-col-before', { value: 1 })
      })
      await sheets.waitForTimeout(1000)

      const boxAfterInsideCol = await sheets.locator('.xlsx-image').boundingBox()
      console.log('Image bounding box after inserting column inside:', boxAfterInsideCol)
      await sheets.screenshot({ path: screenshotPath('image-after-inside-col') })

      // The image should expand because a column was inserted inside its bounds
      expect((boxAfterInsideCol?.width ?? 0)).toBeGreaterThan(boxAfterRightCol?.width ?? 0)

      // Now click on the image to re-select it
      console.log('Testing re-selection: Clicking on .xlsx-image...')
      await sheets.locator('.xlsx-image').click()
      await sheets.waitForTimeout(500)
      const isReSelected = await shapeEditable.evaluate((el) => el.classList.contains('selected'))
      console.log('Is re-selected after click?', isReSelected)
      expect(isReSelected).toBe(true)

    } finally {
      await closeAndSaveVideo(launched, 'sheets-image-test')
    }
  })
})
