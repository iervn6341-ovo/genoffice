import { expect, test, type Page } from '@playwright/test'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { addChart, createBlankPptx, openPptx, savePptx } from '@genoffice/pptx-engine'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_SLIDES_CHART_NATIVE_VERIFICATION,
  'Opt-in imported chart preservation verification',
)

async function seedImportedChart(path: string): Promise<void> {
  const opened = await openPptx(await createBlankPptx())
  addChart(opened, 0, {
    kind: 'bar',
    categories: ['A', 'B'],
    series: [{ name: 'Revenue', values: [4, 7] }],
    offset: { x: 1828800, y: 1371600, cx: 5486400, cy: 3657600 },
    legendPos: 'b',
  })
  const zip = await JSZip.loadAsync(await savePptx(opened))
  zip.file(
    'ppt/slides/slide1.xml',
    (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
      ' descr="aislides-chart"',
      '',
    ),
  )
  zip.file(
    'ppt/charts/chart1.xml',
    (await zip.file('ppt/charts/chart1.xml')!.async('string')).replace(
      '</c:chartSpace>',
      '<c:extLst data-preserve="imported-chart"/></c:chartSpace>',
    ),
  )
  await writeFile(path, await zip.generateAsync({ type: 'uint8array' }))
}

async function selectChart(page: Page): Promise<void> {
  const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
  const slide = slides?.[0]
  const chart = slide?.nodes.find((node) => node.type === 'chart')
  const rect = await page.locator('.stage-rel').boundingBox()
  expect(slide).toBeTruthy()
  expect(chart).toBeTruthy()
  expect(rect).toBeTruthy()
  await page.mouse.click(
    rect!.x + (chart!.box.centerX / slide!.widthPx) * rect!.width,
    rect!.y + (chart!.box.centerY / slide!.heightPx) * rect!.height,
  )
  const designTab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: /^Chart Design$/ })
  await expect(designTab).toBeVisible()
  if (!(await designTab.evaluate((el) => el.classList.contains('active')))) {
    await designTab.click()
  }
}

async function chartXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  return await zip.file('ppt/charts/chart1.xml')!.async('string')
}

const macQuickLayoutFixture = resolve(
  'docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/powerpoint-mac-quick-layout-source.pptx',
)

async function expectOfficeChartPreserved(
  source: string,
  saved: string,
  options?: {
    axisTicksOut?: boolean
    automaticAxisTitles?: boolean
    lastSeriesLabels?: boolean
    allSeriesLabels?: boolean
    valueMinorGridlines?: boolean
    valueMajorGridlinesRemoved?: boolean
    valueAxisHidden?: boolean
  },
): Promise<void> {
  const original = await JSZip.loadAsync(await readFile(source))
  const result = await JSZip.loadAsync(await readFile(saved))
  for (const path of [
    'ppt/charts/_rels/chart1.xml.rels',
    'ppt/charts/style1.xml',
    'ppt/charts/colors1.xml',
    'ppt/embeddings/Microsoft_Excel____.xlsx',
  ]) {
    expect(await result.file(path)!.async('uint8array')).toEqual(
      await original.file(path)!.async('uint8array'),
    )
  }
  const before = await original.file('ppt/charts/chart1.xml')!.async('string')
  const after = await result.file('ppt/charts/chart1.xml')!.async('string')
  const comparableAfter = options?.allSeriesLabels
    ? after.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (series) =>
        series.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, ''),
      )
    : options?.lastSeriesLabels
      ? after.replace(
          /(<c:ser>[\s\S]*?<\/c:ser>[\s\S]*?<c:ser>[\s\S]*?<\/c:ser>[\s\S]*?<c:ser>[\s\S]*?)<c:dLbls>[\s\S]*?<\/c:dLbls>/,
          '$1',
        )
      : after
  for (const name of ['c:ser', 'c:extLst', 'c:catAx', 'c:valAx', 'mc:AlternateContent']) {
    const fragments = (xml: string) =>
      [...xml.matchAll(new RegExp(`<${name}\\b[\\s\\S]*?<\\/${name}>`, 'g'))].map(
        (match) => match[0],
      )
    const actual = fragments(comparableAfter).map((item) => {
      if (!/^(?:c:catAx|c:valAx)$/.test(name)) return item
      let fragment = item
      if (options?.automaticAxisTitles)
        fragment = fragment.replace(/<c:title>[\s\S]*?<\/c:title>/, '')
      if (options?.axisTicksOut)
        fragment = fragment.replace('<c:majorTickMark val="out"/>', '<c:majorTickMark val="none"/>')
      if (name === 'c:valAx' && options?.valueMinorGridlines)
        fragment = fragment.replace(/<c:minorGridlines>[\s\S]*?<\/c:minorGridlines>/, '')
      if (name === 'c:valAx' && options?.valueAxisHidden)
        fragment = fragment.replace('<c:delete val="1"/>', '<c:delete val="0"/>')
      return fragment
    })
    const expected = fragments(before).map((item) =>
      name === 'c:valAx' && options?.valueMajorGridlinesRemoved
        ? item.replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/, '')
        : item,
    )
    expect(actual, `${name} from the PowerPoint Mac chart`).toEqual(expected)
  }
  expect(after).toContain('<c:externalData r:id="rId3"')
}

test('imported chart legend moves without simplifying chart data or extensions', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-native-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'legend-right.pptx')
  await seedImportedChart(source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-native-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.locator('.rb-chart-elem-drop').getByRole('button', { name: 'Right' }).click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'chart')?.styleInfo?.legendPos
      })
      .toBe('r')
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'chart')?.styleInfo?.legendPos
      })
      .toBe('b')
    await page.getByTestId('ribbon-redo').click()
    await selectChart(page)
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Chart Format$/ })
      .click()
    await page.getByTestId('chart-format-element').selectOption('chart')
    await page.getByTestId('chart-area-fill').click()
    await page.getByTestId('chart-fill-D9EAF7').click()
    await page.getByTestId('chart-area-outline').click()
    await page.getByTestId('chart-outline-3').click()
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'chart')?.border
      })
      .toBeUndefined()
    await page.getByTestId('ribbon-redo').click()
    await selectChart(page)
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Chart Format$/ })
      .click()
    await page.getByTestId('chart-format-element').selectOption('plot')
    await page.getByTestId('chart-area-fill').click()
    await page.getByTestId('chart-fill-D9EAD3').click()
    await page.getByTestId('chart-area-outline').click()
    await page.getByTestId('chart-outline-1.5').click()
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return chart?.type === 'chart' ? chart.plotRect?.borderColor : undefined
      })
      .toBeUndefined()
    await page.getByTestId('ribbon-redo').click()

    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('legend-right.pptx'))).ok).toBe(true)
    const xml = await chartXml(saved)
    expect(xml).toContain('<c:legendPos val="r"/>')
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(xml).toContain('data-preserve="imported-chart"')
    expect(xml).toContain('<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>')
    expect(xml).toContain(
      '<a:ln w="38100"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:ln>',
    )
    const plotXml = /<c:plotArea>[\s\S]*?<\/c:plotArea>/.exec(xml)?.[0]
    expect(plotXml).toContain('<a:srgbClr val="D9EAD3"/>')
    expect(plotXml).toContain('<a:ln w="19050">')
    await closeAndSaveVideo(launched, 'chart-native-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-native-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.locator('.rb-chart-elem-drop').getByRole('button', { name: 'Top' }).click()
    await selectChart(page)
    await page
      .locator('.ribbon-tabs .ribbon-tab')
      .filter({ hasText: /^Chart Format$/ })
      .click()
    await page.getByTestId('chart-format-element').selectOption('plot')
    await page.getByTestId('chart-area-fill').click()
    await page.getByTestId('chart-fill-FCE4D6').click()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    expect(await chartXml(saved)).toContain('<c:legendPos val="t"/>')
    expect(await chartXml(saved)).toContain('<a:solidFill><a:srgbClr val="FCE4D6"/></a:solidFill>')
    expect(await chartXml(saved)).toContain('<a:solidFill><a:srgbClr val="D9EAF7"/></a:solidFill>')
    expect(await chartXml(saved)).toContain('data-preserve="imported-chart"')
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-native-final')
  }
})

test('Legend None removes only the native legend and remains editable after Save and reopen', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-legend-none-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'legend-none.pptx')
  await seedImportedChart(source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-legend-none-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.getByTestId('chart-legend-none').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'chart')?.styleInfo?.legendPos
      })
      .toBe('none')
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'chart')?.styleInfo?.legendPos
      })
      .toBe('b')
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'chart')?.styleInfo?.legendPos
      })
      .toBe('none')

    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('legend-none.pptx'))).ok).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).not.toContain('<c:legend>')
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(xml).toContain('data-preserve="imported-chart"')
    const originalZip = await JSZip.loadAsync(await readFile(source))
    const savedZip = await JSZip.loadAsync(await readFile(saved))
    for (const path of Object.keys(originalZip.files).filter(
      (name) => name.startsWith('ppt/charts/_rels/') || name.startsWith('ppt/embeddings/'),
    )) {
      expect(await savedZip.file(path)?.async('uint8array')).toEqual(
        await originalZip.file(path)?.async('uint8array'),
      )
    }
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        return slides?.[0]?.nodes.find((node) => node.type === 'chart')?.styleInfo?.legendPos
      })
      .toBe('b')
    await page.getByTestId('ribbon-redo').click()
    await closeAndSaveVideo(launched, 'chart-legend-none-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-legend-none-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.getByTestId('chart-legend-r').click()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(xml).toContain('<c:legendPos val="r"/>')
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(xml).toContain('data-preserve="imported-chart"')
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-legend-none-final')
  }
})

test('native Quick Layout changes imported chart labels and legend without rebuilding data', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-quick-layout-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'quick-layout.pptx')
  await seedImportedChart(source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-quick-layout-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-right-values').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [chart?.styleInfo?.legendPos, chart?.styleInfo?.dataLabels]
      })
      .toEqual(['r', true])
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [chart?.styleInfo?.legendPos, chart?.styleInfo?.dataLabels]
      })
      .toEqual(['b', false])
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [chart?.styleInfo?.legendPos, chart?.styleInfo?.dataLabels]
      })
      .toEqual(['r', true])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('quick-layout.pptx'))).ok).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).toContain('<c:legendPos val="r"/>')
    expect(xml).toContain('<c:dLbls><c:showVal val="1"/></c:dLbls>')
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(xml).toContain('data-preserve="imported-chart"')
    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [chart?.styleInfo?.legendPos, chart?.styleInfo?.dataLabels]
      })
      .toEqual(['b', false])
    await page.getByTestId('ribbon-redo').click()
    await expect
      .poll(async () => {
        const slides = await page.evaluate(() => window.slidesApi.getRenderSlides())
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [chart?.styleInfo?.legendPos, chart?.styleInfo?.dataLabels]
      })
      .toEqual(['r', true])
    await closeAndSaveVideo(launched, 'chart-quick-layout-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-quick-layout-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-values-only').click()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(xml).not.toContain('<c:legend>')
    expect(xml).toContain('<c:showVal val="1"/>')
    expect(xml).toContain('<c:f>Sheet1!$B$2:$B$3</c:f>')
    expect(xml).toContain('data-preserve="imported-chart"')
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-quick-layout-final')
  }
})

test('PowerPoint Mac Quick Layout keeps the embedded workbook and chart extensions', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-quick-layout-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-edited.pptx')
  await copyFile(macQuickLayoutFixture, source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-quick-layout-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-right-values').click()
    const chartState = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const node = slides?.[0]?.nodes.find((item) => item.type === 'chart')
        return [node?.styleInfo?.legendPos ?? null, Boolean(node?.styleInfo?.dataLabels)]
      })
    await expect.poll(chartState).toEqual(['r', true])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(chartState).toEqual(['b', false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(chartState).toEqual(['r', true])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-edited.pptx'))).ok,
    ).toBe(true)
    let officeXml = await chartXml(saved)
    expect(officeXml).toContain('<c:legendPos val="r"/>')
    expect(officeXml).toMatch(/<c:dLbls>[\s\S]*?<c:showVal val="1"\/>[\s\S]*?<\/c:dLbls>/)
    expect(officeXml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved)
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(chartState).toEqual(['b', false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(chartState).toEqual(['r', true])
    await closeAndSaveVideo(launched, 'chart-office-quick-layout-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-quick-layout-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-values-only').click()
    await expect.poll(chartState).toEqual(['none', true])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    officeXml = await chartXml(saved)
    expect(officeXml).not.toMatch(/<c:legend\b/)
    expect(officeXml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved)
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-quick-layout-final')
  }
})

test('PowerPoint Mac column Layouts 1 and 3 round-trip without changing chart data', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-layouts-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-layouts.pptx')
  await copyFile(macQuickLayoutFixture, source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-layouts-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    const state = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [
          chart?.styleInfo?.legendPos,
          Boolean(chart?.styleInfo?.dataLabels),
          chart?.styleInfo?.gapWidthPct,
        ]
      })
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-1').click()
    await expect.poll(state).toEqual(['r', false, 150])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', false, 150])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-layouts.pptx'))).ok,
    ).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).toContain('<c:legendPos val="r"/>')
    expect(xml).toContain('<c:showVal val="0"/>')
    expect(xml).toContain('<c:gapWidth val="150"/>')
    expect(xml).not.toMatch(/<c:overlap\b/)
    expect(xml).toContain('<c:majorGridlines>')
    await expectOfficeChartPreserved(source, saved)
    await closeAndSaveVideo(launched, 'chart-office-layouts-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-layouts-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-3').click()
    await expect.poll(state).toEqual(['b', false, 75])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(xml).toContain('<c:legendPos val="b"/>')
    expect(xml).toContain('<c:showVal val="0"/>')
    expect(xml).toContain('<c:gapWidth val="75"/>')
    expect(xml).toContain('<c:overlap val="-25"/>')
    await expectOfficeChartPreserved(source, saved)
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-layouts-final')
  }
})

test('PowerPoint Mac column Layout 11 removes the native title and remains editable after reopen', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-layout-11-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-layout-11.pptx')
  await copyFile(macQuickLayoutFixture, source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-layout-11-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    const state = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [
          chart?.styleInfo?.legendPos,
          Boolean(chart?.styleInfo?.dataLabels),
          chart?.styleInfo?.gapWidthPct,
          Boolean(chart?.styleInfo?.title),
        ]
      })
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-11').click()
    await expect.poll(state).toEqual(['r', false, 150, false])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', false, 150, false])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-layout-11.pptx'))).ok,
    ).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).not.toMatch(/<c:title\b/)
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>')
    expect((xml.match(/<c:majorTickMark val="out"\/>/g) ?? []).length).toBe(2)
    expect(xml).toContain('<c:legendPos val="r"/>')
    expect(xml).toContain('<c:gapWidth val="150"/>')
    expect(xml).not.toMatch(/<c:overlap\b/)
    await expectOfficeChartPreserved(source, saved, { axisTicksOut: true })
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', false, 150, false])
    await closeAndSaveVideo(launched, 'chart-office-layout-11-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-layout-11-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await expect.poll(state).toEqual(['r', false, 150, false])
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-right-values').click()
    await expect.poll(state).toEqual(['r', true, 150, false])
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-11').click()
    await expect.poll(state).toEqual(['r', false, 150, false])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(xml).not.toMatch(/<c:title\b/)
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>')
    expect((xml.match(/<c:majorTickMark val="out"\/>/g) ?? []).length).toBe(2)
    await expectOfficeChartPreserved(source, saved, { axisTicksOut: true })
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-layout-11-final')
  }
})

test('PowerPoint Mac column Layout 9 adds native automatic axis titles and survives fresh reopen', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-layout-9-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-layout-9.pptx')
  await copyFile(macQuickLayoutFixture, source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-layout-9-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    const state = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [
          chart?.styleInfo?.legendPos,
          Boolean(chart?.styleInfo?.dataLabels),
          chart?.styleInfo?.gapWidthPct,
          Boolean(chart?.styleInfo?.catAxisTitle),
          Boolean(chart?.styleInfo?.valAxisTitle),
        ]
      })
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-9').click()
    await expect.poll(state).toEqual(['r', false, 150, true, true])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', false, 150, true, true])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-layout-9.pptx'))).ok,
    ).toBe(true)
    let xml = await chartXml(saved)
    expect((xml.match(/<c:title>/g) ?? []).length).toBe(3)
    expect((xml.match(/<c:majorTickMark val="out"\/>/g) ?? []).length).toBe(1)
    expect(xml).toContain('<c:gapWidth val="150"/>')
    expect(xml).not.toMatch(/<c:overlap\b/)
    await expectOfficeChartPreserved(source, saved, {
      automaticAxisTitles: true,
      axisTicksOut: true,
    })
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', false, 150, true, true])
    await closeAndSaveVideo(launched, 'chart-office-layout-9-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-layout-9-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await expect.poll(state).toEqual(['r', false, 150, true, true])
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-right-values').click()
    await expect.poll(state).toEqual(['r', true, 150, true, true])
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-9').click()
    await expect.poll(state).toEqual(['r', false, 150, true, true])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect((xml.match(/<c:title>/g) ?? []).length).toBe(3)
    expect((xml.match(/<c:majorTickMark val="out"\/>/g) ?? []).length).toBe(1)
    await expectOfficeChartPreserved(source, saved, {
      automaticAxisTitles: true,
      axisTicksOut: true,
    })
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-layout-9-final')
  }
})

test('PowerPoint Mac column Layout 10 keeps the Office workbook and labels only its third series', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-layout-10-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-layout-10.pptx')
  await copyFile(macQuickLayoutFixture, source)
  const native10 = await readFile(
    resolve(
      'docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-10-chart.xml',
    ),
    'utf8',
  )
  const series = (xml: string) =>
    [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((match) => match[0])
  const labels = (xml: string) => /<c:dLbls>[\s\S]*?<\/c:dLbls>/.exec(xml)?.[0]
  const nativeLastLabels = labels(series(native10)[2]!)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-layout-10-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    const state = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [
          chart?.styleInfo?.legendPos,
          Boolean(chart?.styleInfo?.dataLabels),
          chart?.styleInfo?.gapWidthPct,
        ]
      })
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-10').click()
    await expect.poll(state).toEqual(['r', true, 75])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', true, 75])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-layout-10.pptx'))).ok,
    ).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    expect(xml).toContain('<c:gapWidth val="75"/>')
    expect(xml).toContain('<c:overlap val="40"/>')
    expect(series(xml).map((fragment) => labels(fragment) ?? null)).toEqual([
      null,
      null,
      nativeLastLabels,
    ])
    await expectOfficeChartPreserved(source, saved, { lastSeriesLabels: true })
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', true, 75])
    await closeAndSaveVideo(launched, 'chart-office-layout-10-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-layout-10-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await expect.poll(state).toEqual(['r', true, 75])
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.locator('.rb-chart-elem-drop').getByRole('button', { name: 'Left' }).click()
    await expect.poll(state).toEqual(['l', true, 75])
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-10').click()
    await expect.poll(state).toEqual(['r', true, 75])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    expect(series(xml).map((fragment) => labels(fragment) ?? null)).toEqual([
      null,
      null,
      nativeLastLabels,
    ])
    await expectOfficeChartPreserved(source, saved, { lastSeriesLabels: true })
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-layout-10-final')
  }
})

test('PowerPoint Mac column Layout 7 adds native minor gridlines and survives fresh reopen', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-layout-7-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-layout-7.pptx')
  await copyFile(macQuickLayoutFixture, source)
  const native7 = await readFile(
    resolve(
      'docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-07-chart.xml',
    ),
    'utf8',
  )
  const minor = (xml: string) => /<c:minorGridlines>[\s\S]*?<\/c:minorGridlines>/.exec(xml)?.[0]
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-layout-7-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    const state = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [
          chart?.styleInfo?.legendPos,
          Boolean(chart?.styleInfo?.dataLabels),
          chart?.styleInfo?.gapWidthPct,
          Boolean(chart?.styleInfo?.title),
          Boolean(chart?.styleInfo?.catAxisTitle),
          Boolean(chart?.styleInfo?.valAxisTitle),
        ]
      })
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-7').click()
    await expect.poll(state).toEqual(['r', false, 300, false, true, true])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', false, 300, false, true, true])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-layout-7.pptx'))).ok,
    ).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).not.toMatch(/<c:chart>\s*<c:title>/)
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>')
    expect(xml).toContain('<c:gapWidth val="300"/>')
    expect(xml).not.toMatch(/<c:overlap\b/)
    expect(minor(xml)).toBe(minor(native7))
    expect(xml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved, {
      automaticAxisTitles: true,
      axisTicksOut: true,
      valueMinorGridlines: true,
    })
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['r', false, 300, false, true, true])
    await closeAndSaveVideo(launched, 'chart-office-layout-7-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-layout-7-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await expect.poll(state).toEqual(['r', false, 300, false, true, true])
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.locator('.rb-chart-elem-drop').getByRole('button', { name: 'Left' }).click()
    await expect.poll(state).toEqual(['l', false, 300, false, true, true])
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-7').click()
    await expect.poll(state).toEqual(['r', false, 300, false, true, true])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(minor(xml)).toBe(minor(native7))
    expect(xml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved, {
      automaticAxisTitles: true,
      axisTicksOut: true,
      valueMinorGridlines: true,
    })
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-layout-7-final')
  }
})

test('PowerPoint Mac column Layout 8 removes only known native gridlines and legend', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-layout-8-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-layout-8.pptx')
  await copyFile(macQuickLayoutFixture, source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-layout-8-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    const state = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [
          chart?.styleInfo?.legendPos,
          Boolean(chart?.styleInfo?.dataLabels),
          chart?.styleInfo?.gapWidthPct,
          Boolean(chart?.styleInfo?.gridlines),
          Boolean(chart?.styleInfo?.title),
          Boolean(chart?.styleInfo?.catAxisTitle),
          Boolean(chart?.styleInfo?.valAxisTitle),
        ]
      })
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-8').click()
    await expect.poll(state).toEqual(['none', false, 0, false, true, true, true])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true, true, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['none', false, 0, false, true, true, true])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-layout-8.pptx'))).ok,
    ).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).toContain('<c:gapWidth val="0"/>')
    expect(xml).not.toMatch(/<c:majorGridlines\b|<c:legend\b|<c:overlap\b/)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved, {
      automaticAxisTitles: true,
      axisTicksOut: true,
      valueMajorGridlinesRemoved: true,
    })
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true, true, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['none', false, 0, false, true, true, true])
    await closeAndSaveVideo(launched, 'chart-office-layout-8-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-layout-8-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await expect.poll(state).toEqual(['none', false, 0, false, true, true, true])
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.locator('.rb-chart-elem-drop').getByRole('button', { name: 'Left' }).click()
    await expect.poll(state).toEqual(['l', false, 0, false, true, true, true])
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-8').click()
    await expect.poll(state).toEqual(['none', false, 0, false, true, true, true])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(xml).not.toMatch(/<c:majorGridlines\b|<c:legend\b/)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved, {
      automaticAxisTitles: true,
      axisTicksOut: true,
      valueMajorGridlinesRemoved: true,
    })
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-layout-8-final')
  }
})

test('PowerPoint Mac column Layout 4 labels all series while retaining its Office workbook', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-layout-4-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-layout-4.pptx')
  await copyFile(macQuickLayoutFixture, source)
  const native4 = await readFile(
    resolve(
      'docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-04-chart.xml',
    ),
    'utf8',
  )
  const series = (xml: string) =>
    [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((match) => match[0])
  const labels = (xml: string) => /<c:dLbls>[\s\S]*?<\/c:dLbls>/.exec(xml)?.[0]
  const nativeLabels = series(native4).map((fragment) => labels(fragment))
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-layout-4-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    const state = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [
          chart?.styleInfo?.legendPos,
          Boolean(chart?.styleInfo?.dataLabels),
          chart?.styleInfo?.gapWidthPct,
          Boolean(chart?.styleInfo?.gridlines),
          Boolean(chart?.styleInfo?.title),
          Boolean(chart?.styleInfo?.catAxisTitle),
          Boolean(chart?.styleInfo?.valAxisTitle),
        ]
      })
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-4').click()
    await expect.poll(state).toEqual(['b', true, 75, false, false, false, false])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true, true, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['b', true, 75, false, false, false, false])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-layout-4.pptx'))).ok,
    ).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).toContain('<c:autoTitleDeleted val="1"/>')
    expect(xml).toContain('<c:gapWidth val="75"/>')
    expect(xml).not.toMatch(/<c:majorGridlines\b|<c:overlap\b/)
    expect(series(xml).map((fragment) => labels(fragment))).toEqual(nativeLabels)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved, {
      allSeriesLabels: true,
      valueMajorGridlinesRemoved: true,
    })
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true, true, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['b', true, 75, false, false, false, false])
    await closeAndSaveVideo(launched, 'chart-office-layout-4-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-layout-4-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await expect.poll(state).toEqual(['b', true, 75, false, false, false, false])
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.locator('.rb-chart-elem-drop').getByRole('button', { name: 'Left' }).click()
    await expect.poll(state).toEqual(['l', true, 75, false, false, false, false])
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-4').click()
    await expect.poll(state).toEqual(['b', true, 75, false, false, false, false])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(series(xml).map((fragment) => labels(fragment))).toEqual(nativeLabels)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved, {
      allSeriesLabels: true,
      valueMajorGridlinesRemoved: true,
    })
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-layout-4-final')
  }
})

test('PowerPoint Mac column Layout 2 hides the value axis and retains its Office workbook', async () => {
  test.setTimeout(180_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
      join(tmpdir(), 'genoffice-slides-chart-native'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'chart-office-layout-2-'))
  const source = join(dir, 'powerpoint-mac-source.pptx')
  const saved = join(dir, 'powerpoint-mac-layout-2.pptx')
  await copyFile(macQuickLayoutFixture, source)
  const native2 = await readFile(
    resolve(
      'docs/qa/office-mac-field-evidence-2026-10-13/slides-chart-native/quick-layout-mac/layout-02-chart.xml',
    ),
    'utf8',
  )
  const series = (xml: string) =>
    [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((match) => match[0])
  const labels = (xml: string) => /<c:dLbls>[\s\S]*?<\/c:dLbls>/.exec(xml)?.[0]
  const nativeLabels = series(native2).map((fragment) => labels(fragment))
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'chart-office-layout-2-first',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    const state = async () =>
      page.evaluate(async () => {
        const slides = await window.slidesApi.getRenderSlides()
        const chart = slides?.[0]?.nodes.find((node) => node.type === 'chart')
        return [
          chart?.styleInfo?.legendPos,
          Boolean(chart?.styleInfo?.dataLabels),
          chart?.styleInfo?.gapWidthPct,
          Boolean(chart?.styleInfo?.gridlines),
          Boolean(chart?.styleInfo?.title),
          Boolean(chart?.styleInfo?.catAxisTitle),
          Boolean(chart?.styleInfo?.valAxisTitle),
          Boolean(chart?.styleInfo?.valAxisHidden),
        ]
      })
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-2').click()
    await expect.poll(state).toEqual(['t', true, 150, false, true, false, false, true])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true, true, false, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['t', true, 150, false, true, false, false, true])
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect(
      (await page.evaluate(() => window.slidesApi.saveAs('powerpoint-mac-layout-2.pptx'))).ok,
    ).toBe(true)
    let xml = await chartXml(saved)
    expect(xml).toContain('<c:gapWidth val="150"/>')
    expect(xml).toContain('<c:overlap val="-25"/>')
    expect(xml).toContain('<c:legendPos val="t"/>')
    expect(xml).toMatch(/<c:valAx>[\s\S]*?<c:delete val="1"\/>[\s\S]*?<\/c:valAx>/)
    expect(xml).not.toMatch(/<c:majorGridlines\b/)
    expect(series(xml).map((fragment) => labels(fragment))).toEqual(nativeLabels)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved, {
      allSeriesLabels: true,
      valueMajorGridlinesRemoved: true,
      valueAxisHidden: true,
    })
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(state).toEqual(['b', false, 219, true, true, false, false, false])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(state).toEqual(['t', true, 150, false, true, false, false, true])
    await closeAndSaveVideo(launched, 'chart-office-layout-2-saved')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'chart-office-layout-2-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await selectChart(page)
    await expect.poll(state).toEqual(['t', true, 150, false, true, false, false, true])
    await page.getByText('Add Chart Element', { exact: true }).click()
    await page.locator('.rb-chart-elem-drop').getByRole('button', { name: 'Left' }).click()
    await expect.poll(state).toEqual(['l', true, 150, false, true, false, false, true])
    await selectChart(page)
    await page.getByTestId('chart-quick-layout').click()
    await page.getByTestId('chart-quick-mac-2').click()
    await expect.poll(state).toEqual(['t', true, 150, false, true, false, false, true])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    xml = await chartXml(saved)
    expect(series(xml).map((fragment) => labels(fragment))).toEqual(nativeLabels)
    expect(xml).toContain('<c:externalData r:id="rId3"')
    await expectOfficeChartPreserved(source, saved, {
      allSeriesLabels: true,
      valueMajorGridlinesRemoved: true,
      valueAxisHidden: true,
    })
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'chart-office-layout-2-final')
  }
})
