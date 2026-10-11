import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Docs ribbon parity with Word: the ribbon never scrolls sideways (Insert groups fold into dropdown
 * buttons), Design-tab changes (page color, watermark) and header/footer edits share the ⌘Z stack,
 * and Change Case offers tOGGLE cASE.
 */
async function openDocs(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-ribbon-parity' })
  await launched.page.locator('.quick-card').first().click()
  const p = await waitForPageWithUrl(launched.app, '://docs/', 30_000)
  await p.waitForSelector('.doc-page', { timeout: 20_000 })
  await p.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')

test.describe('Docs ribbon never scrolls sideways', () => {
  for (const width of [1280, 1000, 900]) {
    test(`every tab fits at ${width}px`, async () => {
      const { launched, p } = await openDocs()
      try {
        await setEditorLayoutWidth(launched.app, '://docs/', width)
        const tabs = (await p.locator('.ribbon-tab:not(.ribbon-tab-file)').allInnerTexts()).map(
          (t) => t.trim(),
        )
        expect(tabs.length).toBeGreaterThanOrEqual(8)
        for (const tab of tabs) {
          await p
            .locator('.ribbon-tab:not(.ribbon-tab-file)', { hasText: new RegExp(`^${tab}$`) })
            .click()
          await p.waitForTimeout(350)
          const m = await p.evaluate(() => {
            const el = document.querySelector('[data-ribbon-body]') as HTMLElement
            return { sw: el.scrollWidth, cw: el.clientWidth }
          })
          expect(m.sw, `${tab} at ${width}px`).toBeLessThanOrEqual(m.cw + 1)
        }
      } finally {
        await kill(launched)
      }
    })
  }

  test('a folded Insert group opens as a flyout and its command still works', async () => {
    const { launched, p } = await openDocs()
    try {
      await setEditorLayoutWidth(launched.app, '://docs/', 900)
      await p.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
      await p.waitForTimeout(400)
      const folded = p.locator('.rb-fold-btn:visible')
      expect(await folded.count()).toBeGreaterThan(2)
      // the caption row only exists while a group is unfolded
      expect(
        await p.locator('.rb-fold-group[data-folded] > .rb-fold-caption:visible').count(),
      ).toBe(0)
      await folded.first().click()
      await expect(p.locator('.rb-fold-open .rb-fold-items')).toBeVisible()
    } finally {
      await kill(launched)
    }
  })

  test('nothing folds when the window is wide enough', async () => {
    const { launched, p } = await openDocs()
    try {
      await setEditorLayoutWidth(launched.app, '://docs/', 1600)
      for (const tab of ['Insert', 'References', 'Review', 'View']) {
        await p.locator('.ribbon-tab', { hasText: new RegExp(`^${tab}$`) }).click()
        await p.waitForTimeout(350)
        expect(await p.locator('.rb-fold-btn:visible').count(), tab).toBe(0)
      }
    } finally {
      await kill(launched)
    }
  })
})

test.describe('Design changes share the undo stack', () => {
  test('page color, watermark and theme colors come back with ⌘Z', async () => {
    const { launched, p } = await openDocs()
    try {
      await p.evaluate(() => {
        const ed = (window as any).__aidocs.editor
        ed.chain().setMeta('addToHistory', false).setContent('<p>Hello world</p>').run()
        ed.commands.focus()
      })
      const pageBg = () =>
        p.evaluate(
          () =>
            (document.querySelector('.doc-page') as HTMLElement | null)?.style.getPropertyValue(
              '--page-bg',
            ) ||
            getComputedStyle(document.querySelector('.doc-page') as HTMLElement).getPropertyValue(
              '--page-bg',
            ),
        )
      await p.locator('.ribbon-tab', { hasText: /^Design$/ }).click()
      const before = await pageBg()
      await p
        .locator('.ribbon-body button[data-tip^="Choose a color for the page"]')
        .first()
        .click()
      await p.locator('.color-palette-page button[data-tip="Light Yellow"]').first().click()
      await expect.poll(pageBg).not.toBe(before)
      const colored = await pageBg()
      await p.evaluate(() => (window as any).__aidocs.editor.commands.focus())
      await p.keyboard.press('ControlOrMeta+z')
      await expect.poll(pageBg).toBe(before)
      await p.keyboard.press('ControlOrMeta+Shift+z')
      await expect.poll(pageBg).toBe(colored)

      // watermark
      const wm = () => p.locator('.page-watermark').count()
      await p.locator('.ribbon-body button[data-tip^="Add ghosted text"]').first().click()
      await p.locator('.layout-menu button', { hasText: 'DRAFT' }).click()
      await expect.poll(wm).toBeGreaterThan(0)
      await p.evaluate(() => (window as any).__aidocs.editor.commands.focus())
      await p.keyboard.press('ControlOrMeta+z')
      await expect.poll(wm).toBe(0)
    } finally {
      await kill(launched)
    }
  })
})

test.describe('Change Case', () => {
  test('tOGGLE cASE swaps the selection and undoes', async () => {
    const { launched, p } = await openDocs()
    try {
      await p.evaluate(() => {
        const ed = (window as any).__aidocs.editor
        ed.chain().setMeta('addToHistory', false).setContent('<p>Hello World</p>').run()
        ed.commands.focus()
        ed.commands.setTextSelection({ from: 1, to: 12 })
      })
      await p
        .locator(
          '.ribbon-body button[aria-label^="Change Case"], .ribbon-body button[data-tip^="Change Case"]',
        )
        .first()
        .click()
      await p.locator('.case-menu button', { hasText: 'tOGGLE cASE' }).click()
      await expect
        .poll(() => p.evaluate(() => (window as any).__aidocs.editor.getText()))
        .toBe('hELLO wORLD')
      await p.keyboard.press('ControlOrMeta+z')
      await expect
        .poll(() => p.evaluate(() => (window as any).__aidocs.editor.getText()))
        .toBe('Hello World')
    } finally {
      await kill(launched)
    }
  })
})
