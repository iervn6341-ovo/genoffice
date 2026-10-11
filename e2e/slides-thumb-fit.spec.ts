import { test, expect } from '@playwright/test'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * The thumbnail list scrolls; a classic scrollbar takes width from its content box. Every thumbnail
 * canvas must still fit inside its card — otherwise the right edge of each page is cut off.
 */
test('slides: thumbnails fit their cards when the list shows a scrollbar', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'thumb-fit',
    openFile: join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
  })
  try {
    const s = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await s.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
    // a scrollbar that takes width, whatever the OS scrollbar setting
    await s.addStyleTag({ content: '.slide-list{overflow-y:scroll !important}' })
    await s.waitForTimeout(400)
    const worst = () =>
      s.evaluate(() => {
        const cards = [...document.querySelectorAll<HTMLElement>('.slide-list .thumb')]
        return cards.reduce((m, card) => {
          const canvas = card.querySelector('canvas')
          return Math.max(m, canvas ? canvas.getBoundingClientRect().width - card.clientWidth : 0)
        }, -999)
      })
    await expect.poll(worst).toBeLessThanOrEqual(0.5)
  } finally {
    await closeAndSaveVideo(launched, 'thumb-fit')
  }
})
