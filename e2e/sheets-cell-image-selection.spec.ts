import { test, expect } from '@playwright/test'
import { copyFile, mkdtemp, writeFile } from 'node:fs/promises'
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

test.describe('sheets: In-cell picture selection, toolbar and floating conversion', () => {
  test.setTimeout(90_000)

  test('inserts in-cell picture, selects it, and converts to floating picture', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-cell-image-test-'))
    const workbook = join(scratch, 'test-cell-image.xlsx')
    await copyFile(FIXTURE, workbook)

    // Create a 64x64 valid red PNG image file
    const samplePngBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAQklEQVR42u3PMQEAAAgEID/517UGDhy4sXSQQJ+eAIFAQCAQCAQCgUAgEAgEAgGBQCAQCAQCgUAgEAgEAoFAILgKqG1tAQE70t8AAAAASUVORK5CYII='
    const testImagePath = join(scratch, 'sample.png')
    await writeFile(testImagePath, Buffer.from(samplePngBase64, 'base64'))

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-cell-image-test',
      openFile: workbook,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })

    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      await waitForWorkbook(sheets)

      sheets.on('console', (msg) => {
        console.log(`[SHEETS CONSOLE] ${msg.type()}: ${msg.text()}`)
      })

      // Select cell B3
      await sheets.evaluate(() => {
        const api = (window as any).__genofficeDebug?.univerAPI
        api?.getActiveWorkbook().getActiveSheet().getRange(2, 1).activate()
      })
      await sheets.waitForTimeout(500)

      // Trigger insert picture in cell
      const fileChooserPromise = sheets.waitForEvent('filechooser')
      await sheets.evaluate(() => {
        ;(window as any).__genofficeDebug?.ribbonCommand('insert-picture-in-cell')
      })
      const fileChooser = await fileChooserPromise
      await fileChooser.setFiles(testImagePath)

      // Verify the in-cell image appears in the worksheet
      const cellImage = sheets.locator('.xlsx-cell-image')
      await expect(cellImage).toBeVisible({ timeout: 15_000 })

      // Initially, it should not be selected
      const isSelectedInitial = await cellImage.evaluate((el) => el.classList.contains('selected'))
      console.log('Is cell image selected initially?', isSelectedInitial)

      // Click on the in-cell image to select it
      console.log('Clicking on the in-cell image...')
      await cellImage.click()
      await sheets.waitForTimeout(500)

      // Verify selection state and toolbar visibility
      await expect(cellImage).toHaveClass(/selected/)
      const toolbar = sheets.locator('.xlsx-cell-image-toolbar')
      await expect(toolbar).toBeVisible()

      // Capture screenshot showing selected in-cell image with action toolbar
      await sheets.screenshot({ path: screenshotPath('cell-image-selected-with-toolbar') })

      // Click the "放置於儲存格上" button to convert to floating visual
      console.log('Converting in-cell image to floating visual...')
      const convertBtn = toolbar.locator('button', { hasText: '放置於儲存格上' })
      await expect(convertBtn).toBeVisible()
      await convertBtn.click()

      // In-cell image should be removed
      await expect(cellImage).toHaveCount(0)

      // Floating editable image should now be visible and selected
      const floatingVisual = sheets.locator('.shape-editable')
      await expect(floatingVisual).toBeVisible({ timeout: 10_000 })
      await expect(floatingVisual).toHaveClass(/selected/)

      // 8 resize handles should be present
      const handles = floatingVisual.locator('.shape-handle')
      await expect(handles).toHaveCount(8)

      // Capture screenshot of converted floating image with 8 resize handles
      await sheets.screenshot({ path: screenshotPath('cell-image-converted-to-floating') })

      // Wait for focus settling
      await sheets.waitForTimeout(1_200)

      // Click on blank grid cell area to deselect
      await sheets.mouse.click(600, 300)
      await sheets.waitForTimeout(500)
      await expect(floatingVisual).not.toHaveClass(/selected/)

      // Click the floating image again to select it
      await floatingVisual.click()
      await sheets.waitForTimeout(500)
      await expect(floatingVisual).toHaveClass(/selected/)
    } finally {
      await closeAndSaveVideo(launched, 'sheets-cell-image-test')
    }
  })

  test('inserts in-cell picture and deletes it via action toolbar', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-cell-image-del-'))
    const workbook = join(scratch, 'test-cell-image-del.xlsx')
    await copyFile(FIXTURE, workbook)

    const samplePngBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAQklEQVR42u3PMQEAAAgEID/517UGDhy4sXSQQJ+eAIFAQCAQCAQCgUAgEAgEAgGBQCAQCAQCgUAgEAgEAoFAILgKqG1tAQE70t8AAAAASUVORK5CYII='
    const testImagePath = join(scratch, 'sample.png')
    await writeFile(testImagePath, Buffer.from(samplePngBase64, 'base64'))

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-cell-image-del-test',
      openFile: workbook,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })

    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      await waitForWorkbook(sheets)

      await sheets.evaluate(() => {
        const api = (window as any).__genofficeDebug?.univerAPI
        api?.getActiveWorkbook().getActiveSheet().getRange(2, 1).activate()
      })
      await sheets.waitForTimeout(500)

      const fileChooserPromise = sheets.waitForEvent('filechooser')
      await sheets.evaluate(() => {
        ;(window as any).__genofficeDebug?.ribbonCommand('insert-picture-in-cell')
      })
      const fileChooser = await fileChooserPromise
      await fileChooser.setFiles(testImagePath)

      const cellImage = sheets.locator('.xlsx-cell-image')
      await expect(cellImage).toBeVisible({ timeout: 15_000 })

      await cellImage.click()
      await sheets.waitForTimeout(500)

      const deleteBtn = sheets.locator('.xlsx-cell-image-toolbar button.delete')
      await expect(deleteBtn).toBeVisible()
      await deleteBtn.click()

      await expect(cellImage).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, 'sheets-cell-image-del-test')
    }
  })
})
