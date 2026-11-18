import { test, expect } from '@playwright/test'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('Docs native TOC retains page labels and follows reading position in the shared outline', async () => {
  test.setTimeout(120_000)
  const root = await mkdtemp(join(tmpdir(), 'genoffice-native-toc-'))
  const path = join(root, 'thesis.docx')
  await copyFile(resolve('我的碩士論文計畫書_修訂版.docx'), path)
  const original = await readFile(path)
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'zh-TW',
    videoDir: 'docs-floating-outline',
  })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI Docs' }).click()
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
    await page.evaluate(async (path) => {
      localStorage.setItem(`aidocs.suppressFontModal.${path}`, '1')
      await (window as any).__aidocs.openPath(path)
    }, path)
    await expect(page.locator('.editor-scroll')).toHaveAttribute('aria-busy', 'false')
    await launched.app.evaluate(({ BrowserWindow }) => {
      BrowserWindow.getAllWindows()[0]!.setSize(1500, 900)
    })
    const trigger = page.locator('.floating-document-outline-trigger')
    await trigger.click()
    const card = page.locator('.floating-document-outline-card')
    await expect(card.locator('header')).toContainText('文件目錄')
    expect(await card.locator('.floating-document-outline-item').count()).toBeGreaterThan(10)
    await expect(card.locator('.floating-document-outline-page').first()).toHaveText('iii')
    const beforeCaret = await page.evaluate(
      () => (window as any).__aidocs.editor.state.selection.from,
    )
    await page.evaluate(() => {
      const editor = (window as any).__aidocs.editor
      let target: HTMLElement | null = null
      let anchor: string | undefined
      editor.state.doc.forEach((node: any) => {
        const field = node.attrs.fieldDisplay
        if (field?.kind === 'tocLine' && String(field.left).includes('緒論')) anchor = field.anchor
      })
      editor.state.doc.forEach((node: any, pos: number) => {
        if (
          anchor &&
          [...(node.attrs.bookmarks ?? []), ...(node.attrs.hiddenBookmarks ?? [])].includes(anchor)
        )
          target = editor.view.nodeDOM(pos)
      })
      const scroller = document.querySelector('.editor-scroll')!
      if (!target) throw new Error('missing chapter heading')
      scroller.scrollTop +=
        (target as HTMLElement).getBoundingClientRect().top -
        scroller.getBoundingClientRect().top -
        24
    })
    await expect(card.locator('[aria-current="location"]')).toContainText('緒論')
    expect(await page.evaluate(() => (window as any).__aidocs.editor.state.selection.from)).toBe(
      beforeCaret,
    )
    await launched.app.evaluate(({ BrowserWindow }) => {
      BrowserWindow.getAllWindows()[0]!.setSize(800, 800)
    })
    await expect
      .poll(async () => {
        const box = await card.boundingBox()
        return box !== null && box.x >= 0 && box.y >= 0 && box.x + box.width <= 800
      })
      .toBe(true)
    await page.screenshot({
      path: '/private/tmp/genoffice-unified-native-toc.png',
      animations: 'disabled',
    })
    await page.keyboard.press('Escape')
    await expect(card).toHaveCount(0)
    await trigger.press('Enter')
    await expect(card).toBeVisible()
    await expect(card.locator('[aria-current="location"]')).toContainText('緒論')
    expect(await readFile(path)).toEqual(original)
  } finally {
    await closeAndSaveVideo(launched, 'docs-floating-outline')
    await rm(root, { recursive: true, force: true })
  }
})
