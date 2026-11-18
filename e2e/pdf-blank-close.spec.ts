import { expect, test } from '@playwright/test'
import { mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('untouched PDF closes silently; edits prompt and Undo restores a clean blank', async () => {
  test.setTimeout(90_000)
  const root = mkdtempSync(join(tmpdir(), 'genoffice-pdf-close-'))
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { defaultSaveDir: root },
    videoDir: 'pdf-blank-close',
  })
  try {
    await launched.app.evaluate(({ dialog }) => {
      ;(globalThis as any).__pdfClosePrompts = 0
      ;(globalThis as any).__pdfSavePrompts = 0
      dialog.showMessageBox = (async () => {
        ;(globalThis as any).__pdfClosePrompts++
        return { response: 2, checkboxChecked: false }
      }) as never
      dialog.showSaveDialog = (async () => {
        ;(globalThis as any).__pdfSavePrompts++
        return { canceled: true }
      }) as never
    })
    const openBlank = async () => {
      await launched.page.locator('.quick-card', { hasText: 'AI PDF' }).click()
      const pdf = await waitForPageWithUrl(launched.app, '://pdf/', 20_000)
      await expect(pdf.locator('.pdf-page canvas').first()).toBeVisible({ timeout: 20_000 })
      return pdf
    }
    const closeTab = () => launched.page.locator('.tab-item.active .tab-close').click()
    await openBlank()
    await closeTab()
    await expect(launched.page.locator('.tab-close')).toHaveCount(0)
    expect(await launched.app.evaluate(() => (globalThis as any).__pdfClosePrompts)).toBe(0)

    const pdf = await openBlank()
    await pdf.getByRole('button', { name: 'Pages', exact: true }).click()
    await pdf.getByRole('button', { name: 'Rotate right', exact: true }).click()
    await expect(pdf.locator('.tb-save-pending')).toBeVisible()
    await closeTab()
    await expect
      .poll(() => launched.app.evaluate(() => (globalThis as any).__pdfClosePrompts))
      .toBe(1)
    await expect(launched.page.locator('.tab-close')).toHaveCount(1)
    await pdf.keyboard.press('Meta+z')
    await expect(pdf.locator('.tb-save-pending')).toHaveCount(0)
    await pdf.keyboard.press('Meta+Shift+z')
    await expect(pdf.locator('.tb-save-pending')).toBeVisible()
    await pdf.keyboard.press('Meta+z')
    await expect(pdf.locator('.tb-save-pending')).toHaveCount(0)
    await closeTab()
    await expect(launched.page.locator('.tab-close')).toHaveCount(0)
    expect(await launched.app.evaluate(() => (globalThis as any).__pdfClosePrompts)).toBe(1)
    expect(await launched.app.evaluate(() => (globalThis as any).__pdfSavePrompts)).toBe(0)

    const rewritten = await openBlank()
    await rewritten.getByRole('button', { name: 'Pages', exact: true }).click()
    await rewritten.getByRole('button', { name: 'Insert blank page', exact: true }).click()
    await expect(rewritten.locator('.pdf-page')).toHaveCount(2)
    await expect(rewritten.locator('.tb-save-pending')).toBeVisible()
    await closeTab()
    await expect
      .poll(() => launched.app.evaluate(() => (globalThis as any).__pdfClosePrompts))
      .toBe(2)
    await expect(launched.page.locator('.tab-close')).toHaveCount(1)
    expect(readdirSync(root)).toEqual([])
  } finally {
    await closeAndSaveVideo(launched, 'pdf-blank-close')
    rmSync(root, { recursive: true, force: true })
  }
})

test('closing the window with an untouched blank PDF does not prompt', async () => {
  test.setTimeout(60_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'pdf-blank-window-close' })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI PDF' }).click()
    const pdf = await waitForPageWithUrl(launched.app, '://pdf/', 20_000)
    await expect(pdf.locator('.pdf-page canvas').first()).toBeVisible({ timeout: 20_000 })
    await launched.app.evaluate(({ dialog, BrowserWindow }) => {
      ;(globalThis as any).__pdfClosePrompts = 0
      dialog.showMessageBox = (async () => {
        ;(globalThis as any).__pdfClosePrompts++
        return { response: 2, checkboxChecked: false }
      }) as never
      BrowserWindow.getAllWindows()[0]!.close()
    })
    await expect
      .poll(() =>
        launched.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length),
      )
      .toBe(0)
    expect(await launched.app.evaluate(() => (globalThis as any).__pdfClosePrompts)).toBe(0)
  } finally {
    await closeAndSaveVideo(launched, 'pdf-blank-window-close')
  }
})
