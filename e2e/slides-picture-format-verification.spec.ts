import { expect, test, type Page } from '@playwright/test'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { addPicture, createBlankPptx, openPptx, savePptx } from '@genoffice/pptx-engine'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_SLIDES_PICTURE_FORMAT_VERIFICATION,
  'Opt-in picture format verification',
)

const PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  ),
)

async function seedPicture(path: string): Promise<void> {
  const opened = await openPptx(await createBlankPptx())
  addPicture(opened, opened.deck.slides[0]!, {
    bytes: PNG,
    ext: 'png',
    offset: { x: 1828800, y: 1371600, cx: 3657600, cy: 2743200 },
  })
  const zip = await JSZip.loadAsync(await savePptx(opened))
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  zip.file(
    'ppt/slides/slide1.xml',
    xml.replace(
      /(<a:blip\b[^>]*)(?:\/>|><\/a:blip>)/,
      '$1><a:alphaModFix amt="80000"/><a:blur rad="18000"/><a:extLst data-preserve="picture-format"/></a:blip>',
    ),
  )
  await writeFile(path, await zip.generateAsync({ type: 'uint8array' }))
}

async function selectPicture(page: Page): Promise<void> {
  const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
  const slide = slides?.[0]
  const pic = slide?.nodes.find((node) => node.type === 'picture')
  expect(slide).toBeTruthy()
  expect(pic).toBeTruthy()
  const stage = page.locator('.stage-rel')
  const rect = await stage.boundingBox()
  expect(rect).toBeTruthy()
  await page.mouse.click(
    rect!.x + (pic!.box.centerX / slide!.widthPx) * rect!.width,
    rect!.y + (pic!.box.centerY / slide!.heightPx) * rect!.height,
  )
  const formatTab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: /^Picture Format$/ })
  await expect(formatTab).toBeVisible()
  if (!(await formatTab.evaluate((el) => el.classList.contains('active')))) {
    await formatTab.click()
  }
  await expect(page.getByTestId('picture-corrections')).toBeEnabled()
}

async function slideXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  return (await zip.file('ppt/slides/slide1.xml')?.async('string')) ?? ''
}

test('picture corrections and recolor preserve imported picture XML through Save and re-edit', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-picture-format'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'picture-format-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'adjusted.pptx')
  await seedPicture(source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-picture-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectPicture(page)
    await page.getByTestId('picture-corrections').click()
    await page.getByTestId('picture-correction-0.2-0').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.lum?.bright
      })
      .toBeCloseTo(0.2)
    await page.getByTestId('picture-color').click()
    await page.getByTestId('picture-color-sepia').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toEqual(['#704214', '#F5DEB3'])

    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toBeUndefined()
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toEqual(['#704214', '#F5DEB3'])

    await selectPicture(page)
    await page.getByTestId('picture-effects').click()
    await page.getByTestId('effect-shadow-br').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.shadow?.dirDeg
      })
      .toBe(45)

    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('adjusted.pptx'))).ok).toBe(true)
    let xml = await slideXml(saved)
    expect(xml).toContain('<a:lum bright="20000" contrast="0"/>')
    expect(xml).toContain(
      '<a:duotone><a:srgbClr val="704214"/><a:srgbClr val="F5DEB3"/></a:duotone>',
    )
    expect(xml).toContain('<a:alphaModFix amt="80000"/>')
    expect(xml).toContain('<a:blur rad="18000"/>')
    expect(xml).toContain('data-preserve="picture-format"')
    expect(xml).toContain('<a:outerShdw')

    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.shadow
      })
      .toBeUndefined()
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.shadow?.dirDeg
      })
      .toBe(45)
    await closeAndSaveVideo(launched, 'slides-picture-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-picture-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectPicture(page)
    await page.getByTestId('picture-color').click()
    await page.getByTestId('picture-color-grayscale').click()
    await selectPicture(page)
    await page.getByTestId('picture-effects').click()
    await page.getByTestId('effect-shadow-tr').click()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await slideXml(saved)
    expect(xml).toContain('<a:grayscl/>')
    expect(xml).not.toContain('<a:duotone>')
    expect(xml).toContain('<a:lum bright="20000" contrast="0"/>')
    expect(xml).toContain('data-preserve="picture-format"')
    expect(xml).toMatch(/<a:outerShdw[^>]*dir="18900000"/)
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-picture-final')
  }
})

test('Reset Picture restores last saved native image/format and rebases after Save As', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-picture-reset'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'picture-reset-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'saved.pptx')
  await seedPicture(source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-picture-reset-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectPicture(page)
    await page.getByTestId('picture-corrections').click()
    await page.getByTestId('picture-correction-0.2-0').click()
    await selectPicture(page)
    await page.getByTestId('picture-color').click()
    await page.getByTestId('picture-color-sepia').click()
    await selectPicture(page)
    await page.getByTestId('picture-reset').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.lum
      })
      .toBeUndefined()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toBeUndefined()
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toEqual(['#704214', '#F5DEB3'])
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toBeUndefined()
    await selectPicture(page)
    await page.getByTestId('picture-color').click()
    await page.getByTestId('picture-color-sepia').click()
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('saved.pptx'))).ok).toBe(true)
    expect(await slideXml(saved)).toContain('<a:duotone>')
    await selectPicture(page)
    await page.getByTestId('picture-color').click()
    await page.getByTestId('picture-color-grayscale').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toEqual(['#000000', '#FFFFFF'])
    await selectPicture(page)
    await page.getByTestId('picture-reset').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toEqual(['#704214', '#F5DEB3'])
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toEqual(['#000000', '#FFFFFF'])
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'picture')?.duotone
      })
      .toEqual(['#704214', '#F5DEB3'])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    let xml = await slideXml(saved)
    expect(xml).toContain('<a:duotone>')
    expect(xml).not.toContain('<a:grayscl/>')
    expect(xml).toContain('data-preserve="picture-format"')
    await closeAndSaveVideo(launched, 'slides-picture-reset-saved')
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-picture-reset-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectPicture(page)
    await page.getByTestId('picture-color').click()
    await page.getByTestId('picture-color-grayscale').click()
    await selectPicture(page)
    await page.getByTestId('picture-reset').click()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await slideXml(saved)
    expect(xml).toContain('<a:duotone>')
    expect(xml).not.toContain('<a:grayscl/>')
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-picture-reset-final')
  }
})
