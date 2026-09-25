import { test, expect } from '@playwright/test'
import { statSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl } from './helpers'

/**
 * C6 — inserting a section break must not write the file (AutoSave off) nor
 * reset undo: type, insert Breaks ▸ Continuous, the file is untouched, ⌘Z
 * removes the break, ⌘Z again removes the typed text, ⌘⇧Z redoes; an
 * explicit save writes two sections.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

test('section break: no silent save, undo/redo intact', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'docs-section-break-'))
  const file = join(dir, 'sections.docx')
  // start from a saved document so it has a file path
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-section-break' })
  try {
    await launched.app.evaluate(({ dialog }, f) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: f })) as never
    }, file)
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(() => Boolean((window as unknown as DocsWindow).__aidocs?.editor))
    await p.locator('.doc-page').first().click()
    await p.keyboard.type('First section')
    await p.keyboard.press('Meta+s')
    await expect.poll(() => statSync(file, { throwIfNoEntry: false })?.size ?? 0).toBeGreaterThan(0)
    const savedAt = statSync(file).mtimeMs

    await p.keyboard.press('Enter')
    await p.keyboard.type('more')
    await p.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await p.locator('.ribbon-body button', { hasText: 'Breaks' }).first().click()
    await p.getByText('Continuous', { exact: false }).first().click()
    await expect(
      p.locator('.doc-sectbreak-label, [data-label="Section break paragraph"]').first(),
    ).toBeAttached()
    await p.waitForTimeout(2_000)
    expect(statSync(file).mtimeMs).toBe(savedAt) // no silent write

    const breaks = () =>
      p.evaluate(() => {
        let n = 0
        ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((node: any) => {
          if (node.attrs?.breakStartType) n++
        })
        return n
      })
    expect(await breaks()).toBe(1)
    await p.keyboard.press('Meta+z')
    expect(await breaks()).toBe(0)
    await p.keyboard.press('Meta+z')
    await expect(p.locator('.doc-page')).not.toContainText('more')
    await p.keyboard.press('Meta+Shift+z')
    await p.keyboard.press('Meta+Shift+z')
    expect(await breaks()).toBe(1)

    await p.keyboard.press('Meta+s')
    await expect.poll(() => statSync(file).mtimeMs).toBeGreaterThan(savedAt)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
