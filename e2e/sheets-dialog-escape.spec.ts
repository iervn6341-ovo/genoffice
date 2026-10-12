import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Esc closes a Sheets dialog wherever the focus is — Excel's own rule. The ribbon sweep found 18 dialogs
 * that stayed open: most never handled Esc, and the ones that did only heard it with the focus inside.
 * The dialogs are opened from the ribbon (the focus stays on the grid or on the button) and closed with
 * a plain Esc.
 */
const TAB = 'nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)'
const BODY = '[data-ribbon-body]'

async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-dialog-escape',
  })
  await launched.page.locator('.quick-card').nth(1).click()
  const p = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await p.waitForSelector(BODY, { timeout: 20_000 })
  await p.waitForTimeout(1500)
  // wide, so no group is folded into a flyout
  await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')
const tab = (p: Page, name: string) =>
  p
    .locator(TAB, { hasText: new RegExp(`^${name}$`) })
    .first()
    .click()

/** [tab, control (its tooltip text), option in its dropdown or null for a plain button, the dialog's title] */
const DIALOGS: [string, string, string | null, RegExp][] = [
  ['Home', 'Fill', 'Series…', /Series/],
  ['Home', 'Sort & Filter', 'Custom Sort…', /Sort/],
  ['Insert', 'Symbol', null, /Symbol/],
  ['Formulas', 'Name Manager', null, /Name Manager/],
  ['Data', 'Consolidate', null, /Consolidate/],
  ['Data', 'Remove Duplicates', null, /Remove Duplicates/],
]

for (const [tabName, control, option, title] of DIALOGS) {
  test(`Esc closes ${option ?? control} (${tabName}), with the focus outside the dialog`, async () => {
    const { launched, p } = await openSheets()
    try {
      await tab(p, tabName)
      await p.waitForTimeout(400)
      await p.locator(`${BODY} [data-tip="${control}"]`).first().click()
      if (option !== null) {
        await p.waitForTimeout(250)
        await p
          .locator('.menu-select-drop button, .menu-select-drop [role="option"]')
          .filter({ hasText: new RegExp(`^\\s*${option}\\s*$`) })
          .first()
          .click()
      }
      const dialog = p.locator('[role="dialog"]').filter({ hasText: title }).first()
      await expect(dialog).toBeVisible()
      await p.keyboard.press('Escape')
      await expect(dialog).toBeHidden()
    } finally {
      await kill(launched)
    }
  })
}
