import { expect, test, type Page } from '@playwright/test'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

const BASE = join(
  __dirname,
  '../docs/qa/office-mac-math-evidence-2026-10-13/slides/2026-10-13T04-00-55-025Z-HxkaNB/math-symbol-roundtrip.pptx',
)
test.skip(!process.env.GENOFFICE_MATH_RENDER_VERIFICATION, 'Opt-in rotated/mixed math rendering')

async function makeFixture(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(BASE))
  let xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const original = (xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) ?? []).find((shape) =>
    shape.includes('<a14:m'),
  )!
  const math = /<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/
    .exec(original)![0]
    .replace(/<a:rPr lang="en-US" dirty="0"\/>/, '<a:rPr lang="en-US" sz="2600"/>')
  const shape = (
    id: number,
    name: string,
    x: number,
    y: number,
    w: number,
    h: number,
    rotation: number,
    mixed: boolean,
    mathRun = math,
  ) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="${rotation * 60000}"><a:off x="${x * 914400}" y="${y * 914400}"/><a:ext cx="${w * 914400}" cy="${h * 914400}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="${mixed ? 'l' : 'ctr'}"/>${mixed ? '<a:r><a:rPr sz="2600"><a:solidFill><a:srgbClr val="1565C0"/></a:solidFill></a:rPr><a:t>Before </a:t></a:r>' : ''}${mathRun}${mixed ? '<a:r><a:rPr sz="2600" b="1"><a:solidFill><a:srgbClr val="2E7D32"/></a:solidFill></a:rPr><a:t> after</a:t></a:r>' : ''}<a:endParaRPr sz="2600"/></a:p></p:txBody></p:sp>`
  const parallelMath = math
    .replace(
      /(<m:oMath\b[^>]*>)[\s\S]*?(<\/m:oMath>)/,
      (_all, open: string, close: string) =>
        open +
        '<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f><m:r><m:t>+</m:t></m:r><m:f><m:num><m:r><m:t>c</m:t></m:r></m:num><m:den><m:r><m:t>d</m:t></m:r></m:den></m:f>' +
        close,
    )
    .replace('(x^2+√(z))/(n+1)α', '(a)/(b)+(c)/(d)')
  const foreground =
    '<p:sp><p:nvSpPr><p:cNvPr id="43" name="Foreground overlap"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="2179320" y="3931920"/><a:ext cx="274320" cy="822960"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="D32F2F"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp>'
  const start = xml.indexOf('<p:spTree>')
  const end = xml.indexOf('</p:spTree>', start)
  const head = xml
    .slice(start, end)
    .match(/^<p:spTree>[\s\S]*?<p:grpSpPr(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/p:grpSpPr>)/)![0]
  xml =
    xml.slice(0, start) +
    head +
    shape(40, 'Rotated equation', 0.5, 0.7, 4, 1.3, 25, false) +
    shape(41, 'Mixed text equation', 0.5, 2.4, 9, 1.3, 0, true) +
    shape(42, 'Covered equation', 0.5, 4.1, 4, 1.3, 0, false) +
    shape(44, 'Parallel fractions mixed', 5.6, 4.1, 7.2, 1.3, 0, true, parallelMath) +
    foreground +
    xml.slice(end)
  zip.file('ppt/slides/slide1.xml', xml)
  let presentation = await zip.file('ppt/presentation.xml')!.async('string')
  presentation = presentation.replace(
    /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/,
    (_all, inner: string) => `<p:sldIdLst>${inner.match(/<p:sldId\b[^>]*\/>/)![0]}</p:sldIdLst>`,
  )
  zip.file('ppt/presentation.xml', presentation)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function openPage(app: LaunchedApp['app']): Promise<Page> {
  const page = await waitForPageWithUrl(app, '://slides/', 30_000)
  await page.waitForSelector('.stage-rel canvas', { timeout: 30_000 })
  return page
}
async function render(page: Page): Promise<any> {
  return page.evaluate(async () => (await window.slidesApi.getRenderSlides())![0])
}
async function point(page: Page, node: any): Promise<{ x: number; y: number }> {
  const slide = await render(page)
  const bounds = await page.locator('.stage-rel').boundingBox()
  if (!bounds) throw new Error('Missing slide stage')
  const run = node.text?.lines.flatMap((line: any) => line.runs).find((r: any) => r.mathML)
  const localX = run ? (node.text.insets.l ?? 0) + run.x + run.widthPx / 2 : node.box.w / 2
  const localY = run
    ? (node.text.insets.t ?? 0) +
      run.baselineY -
      (run.mathAscentPx ?? run.fontSizePx * 0.8) +
      ((run.mathAscentPx ?? run.fontSizePx * 0.8) + (run.mathDescentPx ?? run.fontSizePx * 0.2)) / 2
    : node.box.h / 2
  const angle = (node.box.rotationDeg * Math.PI) / 180
  const dx = localX - node.box.w / 2
  const dy = localY - node.box.h / 2
  return {
    x:
      bounds.x +
      ((node.box.x + node.box.w / 2 + dx * Math.cos(angle) - dy * Math.sin(angle)) * bounds.width) /
        slide.widthPx,
    y:
      bounds.y +
      ((node.box.y + node.box.h / 2 + dx * Math.sin(angle) + dy * Math.cos(angle)) *
        bounds.height) /
        slide.heightPx,
  }
}
async function screenshotSlide(page: Page, path: string): Promise<void> {
  await page.locator('.stage-rel').screenshot({ path })
}
async function exportPdf(launched: LaunchedApp, target: string): Promise<void> {
  await launched.app.evaluate(({ dialog, webContents }, path) => {
    dialog.showSaveDialog = (async () => ({
      canceled: false,
      filePath: path,
    })) as typeof dialog.showSaveDialog
    webContents
      .getAllWebContents()
      .find((wc) => wc.getURL().includes('://slides/'))
      ?.send('slides:menu', 'export-pdf')
  }, target)
  await expect
    .poll(
      async () => {
        try {
          return (await readFile(target)).subarray(0, 5).toString()
        } catch {
          return ''
        }
      },
      { timeout: 30_000 },
    )
    .toBe('%PDF-')
}

async function exportImages(launched: LaunchedApp, directory: string): Promise<string> {
  await launched.app.evaluate(({ dialog, webContents }, path) => {
    dialog.showOpenDialog = (async () => ({
      canceled: false,
      filePaths: [path],
    })) as typeof dialog.showOpenDialog
    webContents
      .getAllWebContents()
      .find((wc) => wc.getURL().includes('://slides/'))
      ?.send('slides:menu', 'export-images')
  }, directory)
  const path = join(directory, 'rotated-mixed-equations-01.png')
  await expect
    .poll(
      async () => {
        try {
          return PNG.sync.read(await readFile(path)).width
        } catch {
          return 0
        }
      },
      { timeout: 30_000 },
    )
    .toBeGreaterThan(1000)
  return path
}

test('rotated and mixed native equations render on canvas and PDF through save/reopen/edit/reopen', async () => {
  test.setTimeout(240_000)
  const root = resolve(
    process.env.GENOFFICE_MATH_RENDER_VERIFICATION_DIR ??
      'docs/qa/office-mac-math-followup-evidence-2026-10-13/slides',
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, `${new Date().toISOString().replace(/[.:]/g, '-')}-`))
  const working = join(dir, 'rotated-mixed-equations.pptx')
  await makeFixture(working)
  const initialXml = await (
    await JSZip.loadAsync(await readFile(working))
  )
    .file('ppt/slides/slide1.xml')!
    .async('string')
  const initialMath = initialXml.match(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g)!
  let launched: LaunchedApp | undefined
  const observations: Record<string, any> = { dir, initialMath, status: 'running' }
  try {
    launched = await launchShell({ onboardingSeen: true, lang: 'en', openFile: working })
    let page = await openPage(launched.app)
    let slide = await render(page)
    const mathNodes = slide.nodes.filter((n: any) => n.hasEquation)
    expect(mathNodes).toHaveLength(4)
    expect(mathNodes.some((n: any) => n.box.rotationDeg === 25)).toBe(true)
    const mixed = mathNodes.find((n: any) => !n.equationStandalone)!
    const mixedRuns = mixed.text.lines.flatMap((line: any) => line.runs)
    expect(mixedRuns.filter((r: any) => r.mathML)).toHaveLength(1)
    expect(mixedRuns.map((r: any) => r.text).join('')).toContain('Before ')
    expect(mixedRuns.map((r: any) => r.text).join('')).toContain(' after')
    const m = mixedRuns.find((r: any) => r.mathML)
    expect(m.mathML).toContain('<mfrac')
    expect(m.mathML).toContain('<msqrt')
    expect(m.mathML).toContain('<msup')
    observations.initialRender = slide
    await screenshotSlide(page, join(dir, 'editor-initial.png'))
    expect(await page.evaluate(() => window.slidesApi.save())).toMatchObject({ ok: true })
    await copyFile(working, join(dir, 'initial-saved.pptx'))
    const savedXml = await (
      await JSZip.loadAsync(await readFile(working))
    )
      .file('ppt/slides/slide1.xml')!
      .async('string')
    expect(savedXml.match(/<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/g)).toEqual(
      initialMath,
    )
    await closeAndSaveVideo(launched, 'slides-render-first')
    launched = await launchShell({ onboardingSeen: true, lang: 'en', openFile: working })
    page = await openPage(launched.app)
    slide = await render(page)
    const reopenedMixed = slide.nodes.find((n: any) => n.hasEquation && !n.equationStandalone)!
    const p = await point(page, reopenedMixed)
    await page.mouse.click(p.x, p.y)
    await page.mouse.dblclick(p.x, p.y)
    const dialog = page.getByTestId('equation-dialog')
    await expect(dialog).toBeVisible()
    const source = dialog.locator('.go-math-input')
    await expect(source).toHaveValue(reopenedMixed.equationLatex)
    await source.fill('\\frac{x^{2}+\\sqrt{q}}{n+1}α')
    await dialog.getByTestId('equation-commit').click()
    await expect(dialog).toHaveCount(0)
    expect(await page.evaluate(() => window.slidesApi.save())).toMatchObject({ ok: true })
    await copyFile(working, join(dir, 'edited-saved.pptx'))
    const editedXml = await (
      await JSZip.loadAsync(await readFile(working))
    )
      .file('ppt/slides/slide1.xml')!
      .async('string')
    expect(editedXml).toContain('<a:t>Before </a:t>')
    expect(editedXml).toContain('<a:t> after</a:t>')
    expect(editedXml).toContain('id="41"')
    expect(editedXml).toContain('rot="1500000"')
    observations.editedXml = editedXml
    await closeAndSaveVideo(launched, 'slides-render-second')
    launched = await launchShell({ onboardingSeen: true, lang: 'en', openFile: working })
    page = await openPage(launched.app)
    slide = await render(page)
    const finalMixed = slide.nodes.find((n: any) => n.hasEquation && !n.equationStandalone)!
    expect(finalMixed.equationLatex).toContain('q')
    expect(slide.nodes.filter((n: any) => n.hasEquation)).toHaveLength(4)
    observations.finalRender = slide
    await screenshotSlide(page, join(dir, 'editor-final-reopen.png'))
    const finalPoint = await point(page, finalMixed)
    await page.mouse.click(finalPoint.x, finalPoint.y)
    await page.mouse.dblclick(finalPoint.x, finalPoint.y)
    await expect(page.getByTestId('equation-dialog')).toBeVisible()
    await expect(page.getByTestId('equation-dialog').locator('.go-math-input')).toHaveValue(
      finalMixed.equationLatex,
    )
    await page.getByTestId('equation-dialog').getByRole('button', { name: 'Cancel' }).click()
    const exportedImage = await exportImages(launched, dir)
    const png = PNG.sync.read(await readFile(exportedImage))
    const foreground = slide.nodes.find(
      (node: any) => !node.hasEquation && node.fill?.color?.includes('D32F2F'),
    )
    expect(foreground).toBeTruthy()
    const px = Math.round(((foreground.box.x + foreground.box.w / 2) * png.width) / slide.widthPx)
    const py = Math.round(((foreground.box.y + foreground.box.h / 2) * png.height) / slide.heightPx)
    const offset = (py * png.width + px) * 4
    expect([...png.data.subarray(offset, offset + 3)]).toEqual([211, 47, 47])
    observations.exportedImage = exportedImage
    observations.foregroundPixel = [...png.data.subarray(offset, offset + 3)]
    await exportPdf(launched, join(dir, 'structured-equations.pdf'))
    observations.pdf = join(dir, 'structured-equations.pdf')
    observations.status = 'passed'
  } catch (error) {
    observations.status = 'failed'
    observations.error = String(error)
    throw error
  } finally {
    await writeFile(join(dir, 'evidence.json'), JSON.stringify(observations, null, 2))
    if (launched) await closeAndSaveVideo(launched, 'slides-render-final')
  }
})
