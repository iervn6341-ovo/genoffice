import { expect, test } from '@playwright/test'
import { copyFile, mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test('Save during progressive TOC loading waits for the complete document and preserves fields', async () => {
  test.setTimeout(120_000)
  const root = await mkdtemp(join(tmpdir(), 'genoffice-progressive-toc-'))
  const file = join(root, 'toc-document.docx')
  await copyFile(resolve(__dirname, '../我的碩士論文計畫書_修訂版.docx'), file)
  const originalXml = await (
    await JSZip.loadAsync(await readFile(file))
  )
    .file('word/document.xml')!
    .async('string')
  let launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-progressive-toc' })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI Docs' }).click()
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
    await page.evaluate(() => {
      const scroll = document.querySelector('.editor-scroll')!
      const observer = new MutationObserver(() => {
        if (scroll.getAttribute('aria-busy') !== 'true') return
        observer.disconnect()
        setTimeout(() => {
          const api = (window as any).__aidocs
          ;(window as any).__saveDuringLoad = { before: api.editor.state.doc.childCount }
          void api.save().then((ok: boolean) => {
            ;(window as any).__saveDuringLoad = {
              ...(window as any).__saveDuringLoad,
              ok,
              done: true,
              after: api.editor.state.doc.childCount,
            }
          })
        }, 0)
      })
      observer.observe(scroll, { attributes: true, attributeFilter: ['aria-busy'] })
    })
    await page.evaluate((path) => (window as any).__aidocs.openPath(path), file)
    await page.waitForFunction(() => (window as any).__saveDuringLoad?.done, null, {
      timeout: 60_000,
    })
    const result = await page.evaluate(() => (window as any).__saveDuringLoad)
    expect(result.ok).toBe(true)
    expect(result.before).toBeLessThan(result.after)
    expect(result.after).toBe(79)
    const savedXml = await (
      await JSZip.loadAsync(await readFile(file))
    )
      .file('word/document.xml')!
      .async('string')
    expect(savedXml.match(/PAGEREF/g)?.length).toBe(originalXml.match(/PAGEREF/g)?.length)
    expect(savedXml.match(/\bTOC\b/g)?.length).toBe(originalXml.match(/\bTOC\b/g)?.length)
    const textRuns = (xml: string) =>
      [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]).join('')
    expect(textRuns(savedXml)).toBe(textRuns(originalXml))
    await closeAndSaveVideo(launched, 'docs-progressive-toc-first-save')
    launched = await launchShell({
      onboardingSeen: true,
      openFile: file,
      videoDir: 'docs-progressive-toc-reopen',
    })
    const reopened = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(reopened.locator('.doc-toc-line')).toHaveCount(44, { timeout: 30_000 })
    await expect(reopened.locator('.editor-scroll')).toHaveAttribute('aria-busy', 'false')
  } finally {
    await closeAndSaveVideo(launched, 'docs-progressive-toc')
  }
})
