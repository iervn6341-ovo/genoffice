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

test.describe('sheets: visual scrolling stability and cut-paste size preservation', () => {
  test.setTimeout(90_000)

  test('visual remains stable during scroll and preserves exact size when cut and pasted', async () => {
    const scratch = await mkdtemp(join(tmpdir(), 'genoffice-visual-scroll-clip-'))
    const workbook = join(scratch, 'test-visual.xlsx')
    await copyFile(FIXTURE, workbook)

    // Create a 120x80 valid red PNG image file
    const samplePngBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAQklEQVR42u3PMQEAAAgEID/517UGDhy4sXSQQJ+eAIFAQCAQCAQCgUAgEAgEAgGBQCAQCAQCgUAgEAgEAoFAILgKqG1tAQE70t8AAAAASUVORK5CYII='
    const testImagePath = join(scratch, 'sample.png')
    await writeFile(testImagePath, Buffer.from(samplePngBase64, 'base64'))

    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'sheets-visual-scroll-clipboard',
      openFile: workbook,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })

    try {
      const sheets = await waitForPageWithUrl(launched.app, '://sheets/')
      await waitForWorkbook(sheets)

      sheets.on('console', (msg) => {
        console.log(`[SHEETS CONSOLE] ${msg.type()}: ${msg.text()}`)
      })

      // Select cell C4
      await sheets.evaluate(() => {
        const api = (window as any).__genofficeDebug?.univerAPI
        api?.getActiveWorkbook().getActiveSheet().getRange(3, 2).activate()
      })
      await sheets.waitForTimeout(500)

      // Trigger insert floating picture
      const fileChooserPromise = sheets.waitForEvent('filechooser')
      await sheets.evaluate(() => {
        ;(window as any).__genofficeDebug?.ribbonCommand('insert-picture-floating')
      })
      const fileChooser = await fileChooserPromise
      await fileChooser.setFiles(testImagePath)

      // Verify floating visual is created and visible
      const floatingVisual = sheets.locator('.shape-editable')
      await expect(floatingVisual).toBeVisible({ timeout: 15_000 })
      await sheets.waitForTimeout(600)

      const initialBox = await floatingVisual.boundingBox()
      expect(initialBox).not.toBeNull()
      const initialWidth = initialBox!.width
      const initialHeight = initialBox!.height
      console.log('Initial floating visual size:', initialWidth, 'x', initialHeight)

      // 1. Test Scrolling: Scroll sheet down and verify visual is not disposed or remounted
      await sheets.mouse.move(600, 400)
      await sheets.mouse.wheel(0, 300)
      await sheets.waitForTimeout(500)
      await expect(floatingVisual).toBeVisible()

      // Scroll horizontally
      await sheets.mouse.wheel(200, 0)
      await sheets.waitForTimeout(500)
      await expect(floatingVisual).toBeVisible()

      // Scroll back
      await sheets.mouse.wheel(-200, -300)
      await sheets.waitForTimeout(600)
      await expect(floatingVisual).toBeVisible()

      await sheets.screenshot({ path: screenshotPath('visual-after-scroll') })

      // 2. Test Cut and Paste:
      // Select the visual by clicking it
      await floatingVisual.click()
      await sheets.waitForTimeout(300)
      await expect(floatingVisual).toHaveClass(/selected/)

      const preCutBox = await floatingVisual.boundingBox()
      expect(preCutBox).not.toBeNull()

      // Cut visual with shortcut
      console.log('Cutting visual with ControlOrMeta+x...')
      await sheets.keyboard.press('ControlOrMeta+x')
      await sheets.waitForTimeout(600)

      // Verify visual is removed from DOM
      await expect(sheets.locator('.shape-editable')).toHaveCount(0)

      // Select a different cell (e.g. F8)
      await sheets.evaluate(() => {
        const api = (window as any).__genofficeDebug?.univerAPI
        api?.getActiveWorkbook().getActiveSheet().getRange(7, 5).activate()
      })
      await sheets.waitForTimeout(300)

      // Paste with shortcut
      console.log('Pasting visual with ControlOrMeta+v...')
      await sheets.keyboard.press('ControlOrMeta+v')
      await sheets.waitForTimeout(1000)

      // Verify pasted visual appears
      const pastedVisual = sheets.locator('.shape-editable')
      await expect(pastedVisual).toBeVisible({ timeout: 10_000 })

      const pastedBox = await pastedVisual.boundingBox()
      expect(pastedBox).not.toBeNull()
      console.log('Pasted visual size:', pastedBox!.width, 'x', pastedBox!.height)

      // Verify pasted visual preserves original width and height within 2px tolerance
      expect(Math.abs(pastedBox!.width - preCutBox!.width)).toBeLessThanOrEqual(2)
      expect(Math.abs(pastedBox!.height - preCutBox!.height)).toBeLessThanOrEqual(2)

      await sheets.screenshot({ path: screenshotPath('visual-after-paste') })
    } finally {
      await closeAndSaveVideo(launched, 'sheets-visual-scroll-clipboard')
    }
  })
})
