import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/** Selecting text shows a Notion-style floating format toolbar; its buttons act on the selection. */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

const doc = (p: Page) =>
  p.evaluate(() => {
    const ed = (window as unknown as DocsWindow).__aidocs!.editor
    const out: Array<{ type: string; level: number | null; text: string; marks: string[] }> = []
    ed.state.doc.forEach((n: any) => {
      const marks = new Set<string>()
      n.descendants((c: any) => c.marks.forEach((m: any) => marks.add(m.type.name)))
      out.push({
        type: n.type.name,
        level: n.attrs.level ?? null,
        text: n.textContent,
        marks: [...marks],
      })
    })
    return out
  })

test('docs: selection toolbar formats the selected text', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'sel-toolbar' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    await p.keyboard.type('Hello toolbar world')
    await expect(p.locator('.sel-tb')).toHaveCount(0)

    await p.keyboard.press('Shift+Home')
    await expect(p.locator('.sel-tb')).toHaveCount(1)

    await p
      .locator('.sel-tb [aria-label="Bold (⌘B)"], .sel-tb button[aria-pressed]')
      .first()
      .click()
    await expect.poll(() => doc(p).then((d) => d[0]!.marks)).toContain('bold')
    // the selection survives the click, so the toolbar is still there
    await expect(p.locator('.sel-tb')).toHaveCount(1)

    // Turn into → Heading 2
    await p.locator('.sel-tb-block').click()
    await p.locator('.sel-tb-item', { hasText: 'Heading 2' }).click()
    await expect
      .poll(() => doc(p).then((d) => d[0]))
      .toMatchObject({ type: 'docHeading', level: 2 })

    // text color
    await p.locator('.sel-tb [aria-label="Font Color"]').click()
    await p.locator('.sel-tb-swatch[aria-label="red"]').first().click()
    await expect.poll(() => doc(p).then((d) => d[0]!.marks)).toContain('docTextStyle')

    // collapsing the selection hides it
    await p.keyboard.press('ArrowRight')
    await expect(p.locator('.sel-tb')).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'sel-toolbar')
  }
})
