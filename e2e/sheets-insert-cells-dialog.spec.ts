import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { mkdtemp } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl } from './helpers'

process.env.GENOFFICE_DEBUG_HOOKS = '1'

/**
 * C4 — Excel's Insert Cells (⌃⇧= / ⌘⇧+) and Delete Cells (⌘-) dialogs on a new
 * workbook: Shift cells down moves B2 to B3, a single ⌘Z restores it, ⌘⇧Z
 * redoes, Esc closes without changes; Delete → Shift cells up reverses it.
 */

type Debug = {
  univerAPI: {
    getActiveWorkbook(): {
      getActiveSheet(): {
        getRange(
          row: number,
          column: number,
          rows?: number,
          columns?: number,
        ): { activate(): unknown; getValue(): unknown; setValue(v: unknown): unknown }
      }
    }
  }
}

const valueAt = (p: Page, row: number, column: number) =>
  p.evaluate(
    ([r, c]) =>
      ((window as unknown as Record<string, unknown>).__genofficeDebug as Debug).univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(r!, c!)
        .getValue(),
    [row, column],
  )

test('Insert Cells dialog: shift down, undo, redo, Esc; Delete Cells: shift up', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-insert-cells-'))
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'sheets-insert-cells' })
  try {
    const { app, page } = launched
    await app.evaluate(
      ({ app: electronApp }, dir) => electronApp.setPath('documents', dir),
      scratch,
    )
    await page.locator('.quick-card').nth(1).click()
    const sheets = await waitForPageWithUrl(app, '://sheets/')
    await sheets.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
      timeout: 30_000,
    })
    await sheets.waitForTimeout(1_500)
    const grid = await sheets.evaluate(() => {
      for (const canvas of document.querySelectorAll('canvas')) {
        const rect = canvas.getBoundingClientRect()
        if (rect.width > 500 && rect.height > 300) return { x: rect.x, y: rect.y }
      }
      return null
    })
    if (!grid) throw new Error('worksheet canvas not found')
    await sheets.mouse.click(grid.x + 46 + 43, grid.y + 24 + 12)
    await sheets.evaluate(() => {
      const sheet = (
        (window as unknown as Record<string, unknown>).__genofficeDebug as Debug
      ).univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
      sheet.getRange(1, 1).setValue('moved')
      sheet.getRange(1, 1, 1, 3).activate() // B2:D2 — wider than tall → Shift cells down
    })

    await sheets.keyboard.press('ControlOrMeta+Shift+Equal')
    const dialog = sheets.getByRole('dialog', { name: 'Insert' })
    await expect(dialog).toBeVisible()
    await expect(dialog.getByLabel('Shift cells down')).toBeChecked()
    await sheets.keyboard.press('Escape')
    await expect(dialog).toBeHidden()
    expect(await valueAt(sheets, 1, 1)).toBe('moved')

    await sheets.keyboard.press('ControlOrMeta+Shift+Equal')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => valueAt(sheets, 2, 1)).toBe('moved')
    expect(await valueAt(sheets, 1, 1)).toBeFalsy()

    await sheets.keyboard.press('ControlOrMeta+z')
    await expect.poll(() => valueAt(sheets, 1, 1)).toBe('moved')
    // redo: on macOS ⇧⌘Z is a native menu accelerator Playwright can't fire,
    // so click the QAT Redo (same path); Ctrl+Y elsewhere (Excel's binding)
    if (process.platform === 'darwin') {
      await sheets.getByRole('button', { name: 'Redo', exact: true }).click()
    } else {
      await sheets.keyboard.press('Control+y')
    }
    await expect.poll(() => valueAt(sheets, 2, 1)).toBe('moved')

    await sheets.keyboard.press('ControlOrMeta+Minus')
    const del = sheets.getByRole('dialog', { name: 'Delete' })
    await expect(del.getByLabel('Shift cells up')).toBeChecked()
    await sheets.keyboard.press('Enter')
    await expect.poll(() => valueAt(sheets, 1, 1)).toBe('moved')
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
