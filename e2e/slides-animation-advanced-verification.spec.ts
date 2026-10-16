import { expect, test } from '@playwright/test'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import {
  addElement,
  createBlankPptx,
  elementSpid,
  getSlideAnimations,
  openPptx,
  savePptx,
  setSlideAnimations,
} from '@genoffice/pptx-engine'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_SLIDES_ANIMATION_VERIFICATION,
  'Opt-in native animation verification',
)

const fixture = join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')

async function seedNativeAnimations(
  path: string,
): Promise<{ importedId: number; importedXml: string }> {
  const opened = await openPptx(await readFile(fixture))
  const slide = opened.deck.slides[0]!
  const targets = slide.elements.flatMap((el) => {
    const spid = elementSpid(el)
    return spid == null ? [] : [spid]
  })
  expect(targets.length).toBeGreaterThanOrEqual(2)
  const imported =
    '<p:par><p:cTn id="501" presetID="77" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">' +
    '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
    `<p:animEffect transition="in" filter="nativeFilter"><p:cBhvr><p:cTn id="502" dur="600"/><p:tgtEl><p:spTgt spid="${targets[1]}"/></p:tgtEl></p:cBhvr></p:animEffect>` +
    '<p:extLst><p:ext uri="keep-imported-effect"/></p:extLst></p:childTnLst></p:cTn></p:par>'
  setSlideAnimations(slide, [
    { spid: targets[0]!, effect: 'flyIn', trigger: 'onClick', durationMs: 500, delayMs: 0 },
    {
      spid: targets[1]!,
      effect: 'fade',
      trigger: 'onClick',
      durationMs: 600,
      delayMs: 0,
      presetXml: imported,
    },
  ])
  const native = getSlideAnimations(slide, true)[1]!
  await writeFile(path, await savePptx(opened))
  return { importedId: native.nativeId!, importedXml: native.nativeXml! }
}

async function slideTiming(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  return /<p:timing\b[\s\S]*?<\/p:timing>/.exec(xml)?.[0] ?? ''
}

test('advanced animation timing and options preserve imported effects across Undo, Save, and reopen', async () => {
  test.setTimeout(210_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ?? join(tmpdir(), 'genoffice-slides-animation'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'advanced-animation-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'advanced-animation.pptx')
  const { importedId, importedXml } = await seedNativeAnimations(source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-animation-advanced',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Animations$/ })
      .click()
    await page.getByText('Animation Pane', { exact: true }).first().click()
    await expect(page.getByTestId('animation-row-1')).toContainText('Imported effect')
    await page.getByTestId('animation-row-0').click()

    const duration = page.getByTestId('anim-duration').locator('input')
    await duration.fill('1.25')
    await duration.press('Enter')
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getAnimations(0)))
      .toMatchObject([
        { effect: 'flyIn', durationMs: 1250 },
        { nativeId: importedId, customPreset: true },
      ])
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(
        async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.durationMs,
      )
      .toBe(500)
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    const undone = await slideTiming(source)
    expect(undone).toContain(importedXml)
    expect(undone).toContain('dur="500"')
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(
        async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.durationMs,
      )
      .toBe(1250)

    await page.getByTestId('anim-start').locator('.gs-dd-btn').click()
    await page.getByTestId('anim-start').locator('[role="option"][data-value="withPrev"]').click()
    await expect
      .poll(async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.trigger)
      .toBe('withPrev')

    await page.getByTestId('anim-effect-direction').locator('.gs-dd-btn').click()
    await page
      .getByTestId('anim-effect-direction')
      .locator('[role="option"][data-value="left"]')
      .click()
    await expect
      .poll(
        async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.direction,
      )
      .toBe('left')
    await page.getByTestId('anim-after-effect').locator('.gs-dd-btn').click()
    await page
      .getByTestId('anim-after-effect')
      .locator('[role="option"][data-value="hideAfter"]')
      .click()
    await expect
      .poll(async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.after)
      .toBe('hideAfter')

    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('advanced-animation.pptx'))).ok).toBe(
      true,
    )
    const written = await slideTiming(saved)
    expect(written).toContain(importedXml)
    expect(written).toContain(`id="${importedId}"`)
    expect(written).toContain('presetSubtype="8"')
    expect(written).toContain('afterEffect="1"')
    expect(written).toContain('nodeType="withEffect"')

    await closeAndSaveVideo(launched, 'slides-animation-advanced-saved')
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-animation-advanced-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Animations$/ })
      .click()
    await page.getByText('Animation Pane', { exact: true }).first().click()
    await page.getByTestId('animation-row-0').click()
    await page.getByTestId('anim-after-effect').locator('.gs-dd-btn').click()
    await page
      .getByTestId('anim-after-effect')
      .locator('[role="option"][data-value="hideOnNextClick"]')
      .click()
    await expect
      .poll(async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.after)
      .toBe('hideOnNextClick')
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    const final = await slideTiming(saved)
    expect(final).toContain(importedXml)
    expect(final).toContain('masterRel="nextClick"')
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-animation-advanced-final')
  }
})

test('More Paths gallery authors a native motion path and supports Undo, Save, preview, and fresh reopen', async () => {
  test.setTimeout(210_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ?? join(tmpdir(), 'genoffice-slides-animation'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'more-paths-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'motion-path.pptx')
  const opened = await openPptx(await createBlankPptx())
  addElement(opened.deck.slides[0]!, {
    kind: 'rect',
    offset: { x: 914400, y: 914400, cx: 1828800, cy: 914400 },
    fillColor: '#3366CC',
  })
  await writeFile(source, await savePptx(opened))

  const arc = 'M 0 0 C 0.12 -0.16 0.28 -0.16 0.4 0'
  const spiral =
    'M 0 0 C 0.12 -0.14 0.34 -0.04 0.29 0.15 C 0.25 0.31 0.05 0.26 0.08 0.12 C 0.1 0.04 0.21 0.06 0.19 0.14'
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-animation-more-paths',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    const target = await page.evaluate(async () => {
      const slide = (await window.slidesApi.getRenderSlides())[0]!
      const node = slide.nodes.find((item) => item.type === 'shape' && !item.decoration)!
      return {
        x: node.box.centerX,
        y: node.box.centerY,
        width: slide.widthPx,
        height: slide.heightPx,
      }
    })
    const selectShape = async () => {
      const stage = await page.locator('.stage-rel').boundingBox()
      expect(stage).not.toBeNull()
      await page.mouse.click(
        stage!.x + (target.x / target.width) * stage!.width,
        stage!.y + (target.y / target.height) * stage!.height,
      )
    }
    const chooseMorePath = async (label: string) => {
      const group = page.locator('[data-rbgroup="motionPaths"]')
      const more = group.getByRole('button', { name: 'More Paths' })
      if (!(await more.isVisible()))
        await group.locator('.ribbon-group-items > .rb-drop-wrap > button').click()
      await more.click()
      await page.getByTestId(`motion-path-${label}`).click()
    }
    const openAnimationPane = async () => {
      const pane = page.locator('.anim-pane')
      if (!(await pane.isVisible()))
        await page.getByText('Animation Pane', { exact: true }).first().click()
      await expect(pane).toBeVisible()
    }

    await selectShape()
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Animations$/ })
      .click()
    await chooseMorePath('ribbonPathArc')
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getAnimations(0)))
      .toMatchObject([{ effect: 'motionPath', motionPath: arc }])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(() => page.evaluate(() => window.slidesApi.getAnimations(0))).toEqual([])
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(
        async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.motionPath,
      )
      .toBe(arc)

    await launched.app.evaluate(({ dialog }, targetPath) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: targetPath })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('motion-path.pptx'))).ok).toBe(true)
    expect(await slideTiming(saved)).toContain(`path="${arc} E"`)

    await openAnimationPane()
    await page.locator('.anim-pane [aria-label="Play animation preview for this slide"]').click()
    await expect(page.locator('.anim-preview-overlay')).toBeVisible()
    await expect(page.locator('.anim-preview-overlay')).not.toBeVisible({ timeout: 10_000 })

    await closeAndSaveVideo(launched, 'slides-animation-more-paths-saved')
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-animation-more-paths-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await expect
      .poll(
        async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.motionPath,
      )
      .toBe(arc)
    await selectShape()
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Animations$/ })
      .click()
    await openAnimationPane()
    await expect(page.getByTestId('animation-row-0')).toBeVisible()
    await page.getByTestId('animation-row-0').click()
    await expect(page.getByTestId('animation-row-0')).toHaveClass(/selected/)
    await page
      .getByTestId('animation-row-0')
      .getByRole('button', { name: 'Remove this animation' })
      .click()
    await expect.poll(() => page.evaluate(() => window.slidesApi.getAnimations(0))).toEqual([])
    await selectShape()
    await chooseMorePath('ribbonPathSpiral')
    await expect
      .poll(
        async () => (await page.evaluate(() => window.slidesApi.getAnimations(0)))[0]?.motionPath,
      )
      .toBe(spiral)
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    const final = await slideTiming(saved)
    expect(final).toContain(`path="${spiral} E"`)
    expect(final).not.toContain(`path="${arc} E"`)
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-animation-more-paths-final')
  }
})
