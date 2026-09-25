import { test, expect, type Page } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl } from './helpers'

/**
 * C1 — Word parity for style-inherited bold: in Heading 1 the Bold button is lit
 * (effective bold), clicking it un-bolds (w:b w:val="0"), text typed afterwards
 * stays regular, undo/redo step the change, and save/reopen keeps it regular.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

const boldBtn = (p: Page) => p.locator('.ribbon-body button[data-tip^="Bold"]').first()

const headingWeight = (p: Page) =>
  p.evaluate(() => {
    const h = document.querySelector('.doc-page h1')
    const span = h?.querySelector('span, strong') ?? h
    return span ? Number(getComputedStyle(span).fontWeight) : 0
  })

test('Heading 1: turn off Bold, type, undo/redo, save/reopen', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'docs-style-bold-'))
  const out = join(dir, 'heading-bold.docx')
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-style-bold' })
  try {
    await launched.app.evaluate(({ dialog }, file) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
    }, out)
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(() => Boolean((window as unknown as DocsWindow).__aidocs?.editor))
    await p.locator('.doc-page').first().click()
    await p.keyboard.type('Title')
    // setup only: make the paragraph Heading 1 through the style gallery
    await p.locator('.ribbon-tab', { hasText: /^Home$/ }).click()
    await p.locator('.ribbon-body', { hasText: 'Heading 1' }).getByText('Heading 1').first().click()
    await p.keyboard.press('Meta+a')
    await expect(boldBtn(p)).toHaveClass(/active/)

    await boldBtn(p).click()
    await expect(boldBtn(p)).not.toHaveClass(/active/)
    expect(await headingWeight(p)).toBeLessThan(600)
    await p.keyboard.press('End')
    await p.keyboard.type(' more')
    await expect(boldBtn(p)).not.toHaveClass(/active/)

    await p.keyboard.press('Meta+z') // typing
    await p.keyboard.press('Meta+z') // un-bold
    await p.keyboard.press('Meta+a')
    await expect(boldBtn(p)).toHaveClass(/active/)
    await p.keyboard.press('Meta+Shift+z')
    await p.keyboard.press('Meta+a')
    await expect(boldBtn(p)).not.toHaveClass(/active/)

    await p.keyboard.press('Meta+Shift+s')
    await expect.poll(() => existsSync(out), { timeout: 20_000 }).toBe(true)
    await expect
      .poll(() => execFileSync('unzip', ['-p', out, 'word/document.xml']).toString(), {
        timeout: 15_000,
      })
      .toMatch(/<w:b w:val="0"\/>/)
  } finally {
    launched.app.process().kill('SIGKILL')
  }

  const re = await launchShell({ onboardingSeen: true, videoDir: 'docs-style-bold', openFile: out })
  try {
    const p2 = await waitForPageWithUrl(re.app, '://docs/')
    await expect(p2.locator('.doc-page h1').first()).toBeVisible({ timeout: 30_000 })
    expect(await headingWeight(p2)).toBeLessThan(600)
    await p2.locator('.doc-page h1').first().click()
    await expect(boldBtn(p2)).not.toHaveClass(/active/)
    // continue editing: bold it back on
    await p2.keyboard.press('Meta+a')
    await p2.keyboard.press('Meta+b')
    await expect(boldBtn(p2)).toHaveClass(/active/)
  } finally {
    re.app.process().kill('SIGKILL')
  }
})
