import { expect, test } from '@playwright/test'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { openPptx, savePptx, setCustomShows } from '@genoffice/pptx-engine'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_SLIDES_SHOW_SETUP_VERIFICATION,
  'Opt-in native show setup verification',
)

const source = join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')

async function nativeXml(path: string, part: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  return (await zip.file(part)?.async('string')) ?? ''
}

test('Show Setup and custom slide size survive Undo, Save, fresh reopen, and re-edit', async () => {
  test.setTimeout(210_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ?? join(tmpdir(), 'genoffice-slides-show-setup'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'show-setup-'))
  const seeded = join(dir, 'source-with-custom-show.pptx')
  const saved = join(dir, 'configured-show.pptx')
  const opened = await openPptx(await readFile(source))
  expect(setCustomShows(opened, [{ id: '17', name: 'QA subset', slideIndices: [1, 3] }])).toBe(true)
  await writeFile(seeded, await savePptx(opened))
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: seeded,
      videoDir: 'slides-show-setup-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Slide Show$/ })
      .click()
    await page.getByTestId('ribbon-show-setup').click()
    await expect(page.getByRole('dialog', { name: 'Set Up Slide Show' })).toBeVisible()
    await page.getByTestId('show-type-browse').check()
    await page.getByTestId('show-loop').check()
    await page.getByTestId('show-no-narration').check()
    await page.getByTestId('show-no-animation').check()
    await page.getByTestId('show-slides-range').check()
    await page.getByTestId('show-range-start').fill('2')
    await page.getByTestId('show-range-end').fill('4')
    await page.getByTestId('show-manual').check()
    await page.getByTestId('show-setup-apply').click()
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getShowSetup()))
      .toMatchObject({
        mode: 'browse',
        loop: true,
        showNarration: false,
        showAnimation: false,
        useTimings: false,
        slides: { kind: 'range', start: 2, end: 4 },
      })

    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => (await page.evaluate(() => window.slidesApi.getShowSetup()))?.mode)
      .toBe('speaker')
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(async () => (await page.evaluate(() => window.slidesApi.getShowSetup()))?.mode)
      .toBe('browse')
    await page.getByText('From Beginning', { exact: true }).first().click()
    await expect(page.locator('.slideshow.ss-windowed')).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(page.locator('.slideshow')).not.toBeVisible()

    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Design$/ })
      .click()
    await page.getByText('Slide Size', { exact: true }).first().click()
    await page.getByTestId('slide-size-custom').click()
    await expect(page.getByRole('dialog', { name: 'Custom Slide Size' })).toBeVisible()
    await page.getByTestId('slide-width-cm').fill('30')
    await page.getByTestId('slide-height-cm').fill('18')
    await page.getByTestId('slide-size-apply').click()
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getSlideSize()))
      .toEqual({ cx: 10800000, cy: 6480000 })

    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('configured-show.pptx'))).ok).toBe(
      true,
    )
    const presProps = await nativeXml(saved, 'ppt/presProps.xml')
    expect(presProps).toContain('<p:browse/>')
    expect(presProps).toContain('<p:sldRg st="2" end="4"/>')
    expect(presProps).toMatch(
      /<p:showPr\b[^>]*loop="1"[^>]*showNarration="0"[^>]*showAnimation="0"[^>]*useTimings="0"/,
    )
    expect(await nativeXml(saved, 'ppt/presentation.xml')).toMatch(
      /<p:sldSz\b[^>]*cx="10800000"[^>]*cy="6480000"[^>]*type="custom"/,
    )

    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getSlideSize()))
      .not.toEqual({ cx: 10800000, cy: 6480000 })
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getSlideSize()))
      .toEqual({ cx: 10800000, cy: 6480000 })

    await closeAndSaveVideo(launched, 'slides-show-setup-saved')
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-show-setup-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Slide Show$/ })
      .click()
    await page.getByTestId('ribbon-show-setup').click()
    await expect(page.getByTestId('show-type-browse')).toBeChecked()
    await expect(page.getByTestId('show-range-start')).toHaveValue('2')
    await expect(page.getByTestId('show-range-end')).toHaveValue('4')
    await page.getByTestId('show-type-kiosk').check()
    await page.getByTestId('show-slides-custom').check()
    await expect(page.getByTestId('show-custom-id')).toHaveValue('17')
    await page.getByTestId('show-setup-apply').click()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    expect(await nativeXml(saved, 'ppt/presProps.xml')).toContain('<p:kiosk/>')
    expect(await nativeXml(saved, 'ppt/presProps.xml')).toContain('<p:custShow id="17"/>')
    expect(await nativeXml(saved, 'ppt/presProps.xml')).not.toContain('<p:browse/>')
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-show-setup-final')
  }
})
