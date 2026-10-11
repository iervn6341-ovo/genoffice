import { test, expect } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Entering a table keeps the current ribbon tab (Table Design / Layout only become available),
 * and the two-button Table Layout groups stack vertically to save width.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

test('docs: entering a table keeps the tab; merge/split and delete row/column stack', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'table-ribbon' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    const active = () => p.locator('.ribbon-tab.active, .ribbon-tab[aria-selected="true"]').first()
    await expect(active()).toContainText('Home')

    await p.evaluate(() => {
      const ed = (window as unknown as DocsWindow).__aidocs!.editor
      ed.commands.setContent(
        '<table><tbody><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></tbody></table><p></p>',
      )
      ed.commands.focus()
      ed.commands.setTextSelection(4)
    })
    // the contextual tabs appear, the active tab does not move
    await expect(p.locator('.ribbon-tab', { hasText: 'Table Layout' })).toHaveCount(1)
    await expect(active()).toContainText('Home')

    await p.locator('.ribbon-tab', { hasText: 'Table Layout' }).click()
    const merge = await p
      .locator('.table-tool-stack button', { hasText: 'Merge' })
      .first()
      .boundingBox()
    const split = await p
      .locator('.table-tool-stack button', { hasText: 'Split' })
      .first()
      .boundingBox()
    expect(split!.y).toBeGreaterThan(merge!.y + 10)
    expect(Math.abs(split!.x - merge!.x)).toBeLessThan(3)
    const delRow = await p
      .locator('.table-tool-stack button', { hasText: 'Delete Row' })
      .first()
      .boundingBox()
    const delCol = await p
      .locator('.table-tool-stack button', { hasText: 'Delete Column' })
      .first()
      .boundingBox()
    expect(delCol!.y).toBeGreaterThan(delRow!.y + 10)
    expect(Math.abs(delCol!.x - delRow!.x)).toBeLessThan(3)

    // the two 3-button alignment groups use both ribbon rows too, never one long row
    const rowOf = async (name: RegExp) =>
      (await p.locator('.table-tool-grid button', { hasText: name }).first().boundingBox())!.y
    expect(await rowOf(/Align Bottom/)).toBeGreaterThan((await rowOf(/Align Top/)) + 10)
    expect(await rowOf(/^\s*Right/)).toBeGreaterThan((await rowOf(/^\s*Left/)) + 10)
  } finally {
    await closeAndSaveVideo(launched, 'table-ribbon')
  }
})
