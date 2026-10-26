import { test, expect } from '@playwright/test'
import { resolve } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from '../helpers'

const PPTX_FILE = resolve(__dirname, '../../10:19 精誠開會簡報.pptx')

test('probe video display in normal edit mode vs reading view', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'probe-video-reading',
    openFile: PPTX_FILE,
  })
  try {
    const s = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await s.waitForSelector('.slide-list .thumb', { timeout: 20_000 })

    // Navigate to slide 3
    const thumbs = s.locator('.slide-list .thumb')
    await expect(thumbs).toHaveCount(3)
    await thumbs.nth(2).click()
    await s.waitForTimeout(1000)

    // Capture Normal view
    const stageWrap = s.locator('.stage-wrap')
    await stageWrap.screenshot({ path: '/tmp/video-normal.png' })
    console.log('Saved /tmp/video-normal.png')

    // Switch to reading view
    console.log('Clicking reading view button...')
    const views = s.locator('.status-views .status-play-btn')
    await views.nth(2).click()

    console.log('Waiting for .slideshow.ss-windowed...')
    const show = s.locator('.slideshow.ss-windowed')
    await expect(show).toBeVisible({ timeout: 10_000 })
    console.log('Reading view visible, waiting 1s...')
    await s.waitForTimeout(1000)

    await s.screenshot({ path: '/tmp/video-reading-full.png' })
    console.log('Saved /tmp/video-reading-full.png')

    const videoCount = await s.locator('video').count()
    console.log('Video count in page:', videoCount)
    for (let i = 0; i < videoCount; i++) {
      const v = s.locator('video').nth(i)
      console.log(`Video ${i} src:`, await v.getAttribute('src'), 'class:', await v.getAttribute('class'))
    }

    const showStage = s.locator('.ss-frame').first()
    await showStage.screenshot({ path: '/tmp/video-reading.png' })
    console.log('Saved /tmp/video-reading.png')

    const videoEl = s.locator('video.ss-video-el').first()
    if (await videoEl.count() > 0) {
      const videoBox = await videoEl.boundingBox()
      const containerBox = await s.locator('.ss-video').boundingBox()
      const videoStyles = await videoEl.evaluate((el) => {
        const cs = window.getComputedStyle(el)
        return {
          objectFit: cs.objectFit,
          width: cs.width,
          height: cs.height,
          position: cs.position,
        }
      })
      console.log('Video styles:', videoStyles)
      console.log('Video element bounding box:', videoBox)
      console.log('Video container bounding box:', containerBox)
    }
  } finally {
    await closeAndSaveVideo(launched, 'probe-video-reading')
  }
})
