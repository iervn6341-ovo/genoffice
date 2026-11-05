import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo } from './helpers'

test.describe('home settings font substitution', () => {
  test('opens settings modal, navigates to font substitutions, adds and filters rules', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'zh-TW',
      videoDir: 'home-font-substitution',
    })
    const { page } = launched
    try {
      // 1. Wait for home screen and open settings modal via account button
      const accountBtn = page.locator('.account-btn')
      await expect(accountBtn).toBeVisible()
      await accountBtn.click()

      // 2. Settings dialog should open
      const dialog = page.locator('.set-dialog')
      await expect(dialog).toBeVisible()

      // 3. Find the font substitution nav item
      const fontNavItem = page.locator('.set-nav-item', { hasText: '字型替換' })
      await expect(fontNavItem).toBeVisible()
      await fontNavItem.click()

      // 4. Verify font substitution pane is displayed
      const fontPane = page.locator('.font-sub-pane')
      await expect(fontPane).toBeVisible()
      await expect(fontPane.locator('.set-pane-title')).toHaveText('字型替換')

      // Verify default substitution rules are listed
      const table = fontPane.locator('.font-sub-table')
      await expect(table).toBeVisible()
      await expect(fontPane.locator('.font-sub-orig-badge', { hasText: '新細明體' })).toBeVisible()
      await expect(fontPane.locator('.font-sub-orig-badge', { hasText: '微軟正黑體' })).toBeVisible()
      await expect(fontPane.locator('.font-sub-orig-badge', { hasText: '標楷體' })).toBeVisible()

      // Capture screenshot of the font substitution settings
      await page.screenshot({
        path: '/Users/iervn6341/.gemini/antigravity/brain/1910f902-d695-4025-9910-eddc7ff264f1/home-settings-font-substitution-verified.png',
      })

      // 5. Test search filter
      const searchInput = fontPane.locator('input[type="search"]')
      await searchInput.fill('新細明體')
      await expect(fontPane.locator('.font-sub-orig-badge', { hasText: '新細明體' })).toBeVisible()
      await expect(fontPane.locator('.font-sub-orig-badge', { hasText: '微軟正黑體' })).toHaveCount(0)

      // Clear search
      await searchInput.fill('')

      // 6. Test adding a new custom font rule
      const addBtn = fontPane.locator('button', { hasText: '新增替換規則' })
      await addBtn.click()

      const addCard = fontPane.locator('.font-sub-add-card')
      await expect(addCard).toBeVisible()

      // Fill in original and substitute font
      const origInput = addCard.locator('input[placeholder*="原始字型名稱"]')
      await origInput.fill('CustomThesisFont')

      const saveBtn = addCard.locator('button', { hasText: '儲存規則' })
      await saveBtn.click()

      // Verify the new custom rule is present and has the custom badge
      const newRuleBadge = fontPane.locator('.font-sub-orig-badge', { hasText: 'CustomThesisFont' })
      await expect(newRuleBadge).toBeVisible()
      const row = fontPane.locator('tr.font-sub-tr', { hasText: 'CustomThesisFont' })
      await expect(row.locator('.badge-custom')).toHaveText('自訂')

      // Capture screenshot showing custom rule added
      await page.screenshot({
        path: '/Users/iervn6341/.gemini/antigravity/brain/1910f902-d695-4025-9910-eddc7ff264f1/home-settings-font-substitution-add-verified.png',
      })
    } finally {
      await closeAndSaveVideo(launched, 'home-font-substitution')
    }
  })
})
