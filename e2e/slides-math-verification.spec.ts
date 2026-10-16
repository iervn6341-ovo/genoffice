import { expect, test, type Page } from '@playwright/test'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

const FIXTURE = join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const SYMBOL = '∑'
const SYMBOL_FONT = 'Segoe UI Symbol'
const INITIAL_LATEX = '\\frac{x^{2}+\\sqrt{y}}{n+1}α'
const EDITED_LATEX = '\\frac{x^{2}+\\sqrt{z}}{n+1}α'

test.beforeEach(() => {
  test.skip(
    !process.env.GENOFFICE_SLIDES_MATH_VERIFICATION,
    'Opt-in Slides math verification; set GENOFFICE_SLIDES_MATH_VERIFICATION=1 to run',
  )
})

interface SlidesWindow {
  slidesApi?: {
    save(): Promise<SaveResult>
    getRenderSlides(): Promise<any[] | null>
  }
}

interface SaveResult {
  ok: boolean
  path?: string
  error?: string
  slides?: unknown
}

interface EquationPackageShape {
  slidePath: string
  shapeId: string
  shapeXml: string
  mathXml: string
}

async function evidenceDir(): Promise<string> {
  const root =
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
    join(tmpdir(), 'genoffice-slides-math-verification-2026-10-13')
  const absoluteRoot = resolve(root)
  await mkdir(absoluteRoot, { recursive: true })
  return mkdtemp(join(absoluteRoot, `${new Date().toISOString().replace(/[.:]/g, '-')}-`))
}

async function openRibbonTab(page: Page, name: string): Promise<void> {
  const tab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: new RegExp(`^${name}$`) })
  await expect(tab).toBeVisible()
  await tab.click()
  await expect(tab).toHaveClass(/active/)
}

/** Open the real ribbon fold containing a command when the current layout hides its children. */
async function clickRibbonCommand(page: Page, testId: string): Promise<void> {
  const command = page.getByTestId(testId)
  if (!(await command.isVisible())) {
    const group = command.locator(
      'xpath=ancestor::*[contains(concat(" ", normalize-space(@class), " "), " rb-fold-group ")][1]',
    )
    const foldButton = group.locator(':scope > .rb-fold-btn')
    await expect(foldButton).toBeVisible()
    if ((await foldButton.getAttribute('aria-expanded')) !== 'true') await foldButton.click()
    await expect(command).toBeVisible()
  }
  await command.click()
}

async function openEditor(launched: LaunchedApp): Promise<Page> {
  const page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
  await page.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
  return page
}

async function closeAndOpenFile(
  launched: LaunchedApp,
  path: string,
  videoDir: string,
): Promise<LaunchedApp> {
  await closeAndSaveVideo(launched, `${videoDir}-closed`)
  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: `${videoDir}-reopened`,
  })
  const page = await openEditor(reopened)
  return { ...reopened, page }
}

async function packageSlides(path: string): Promise<Array<{ path: string; xml: string }>> {
  const zip = await JSZip.loadAsync(await readFile(path))
  return Promise.all(
    Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name) && !zip.files[name]?.dir)
      .sort(
        (a, b) =>
          Number(a.match(/slide(\d+)\.xml$/)?.[1]) - Number(b.match(/slide(\d+)\.xml$/)?.[1]),
      )
      .map(async (name) => ({ path: name, xml: await zip.file(name)!.async('string') })),
  )
}

async function equationShape(path: string): Promise<EquationPackageShape> {
  const slides = await packageSlides(path)
  for (const slide of slides) {
    const shapes = slide.xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) ?? []
    const shapeXml = shapes.find((shape) => shape.includes('<a14:m'))
    if (!shapeXml) continue
    const shapeId = /<p:cNvPr\b[^>]*\bid="([^"]+)"/.exec(shapeXml)?.[1]
    const mathXml = /<mc:AlternateContent\b[\s\S]*?<\/mc:AlternateContent>/.exec(shapeXml)?.[0]
    if (!shapeId || !mathXml) throw new Error(`Could not identify native equation in ${slide.path}`)
    return { slidePath: slide.path, shapeId, shapeXml, mathXml }
  }
  throw new Error(`No native equation found in ${path}`)
}

async function symbolShape(path: string): Promise<{ slidePath: string; shapeXml: string }> {
  const slides = await packageSlides(path)
  for (const slide of slides) {
    const shapes = slide.xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) ?? []
    const shapeXml = shapes.find((shape) => shape.includes(SYMBOL))
    if (shapeXml) return { slidePath: slide.path, shapeXml }
  }
  throw new Error(`No ${SYMBOL} text shape found in ${path}`)
}

async function equationBounds(
  page: Page,
): Promise<{ x: number; y: number; width: number; height: number }> {
  const slide = await page.evaluate(async () => (await window.slidesApi.getRenderSlides())![0])
  const node = slide.nodes.find((candidate: any) => candidate.hasEquation) as any
  const bounds = await page.locator('.stage-rel').boundingBox()
  if (!node || !bounds) throw new Error('No native equation canvas bounds')
  return {
    x: bounds.x + (node.box.x * bounds.width) / slide.widthPx,
    y: bounds.y + (node.box.y * bounds.height) / slide.heightPx,
    width: (node.box.w * bounds.width) / slide.widthPx,
    height: (node.box.h * bounds.height) / slide.heightPx,
  }
}

async function currentCanvasMathML(page: Page): Promise<string> {
  return page.evaluate(async () => {
    const slide = (await window.slidesApi.getRenderSlides())![0]
    return (
      (slide.nodes.find((candidate: any) => candidate.hasEquation) as any)?.equationMathML ?? ''
    )
  })
}

async function clickEquation(page: Page): Promise<void> {
  const bounds = await equationBounds(page)
  await page.mouse.click(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2)
}

async function renderObservations(page: Page): Promise<{
  mathNodes: number
  mathLatex: string | null
  symbols: Array<{ slideIndex: number; text: string; fontFamilies: string[] }>
}> {
  return page.evaluate(async () => {
    const slides = await (window as unknown as SlidesWindow).slidesApi!.getRenderSlides()
    const mathNodes = (slides ?? []).flatMap((slide) =>
      (slide.nodes ?? []).filter((node: any) => node.hasEquation),
    )
    const symbols = (slides ?? []).flatMap((slide, slideIndex) =>
      (slide.nodes ?? []).flatMap((node: any) => {
        const runs = (node.text?.lines ?? []).flatMap((line: any) => line.runs ?? [])
        const text = runs.map((run: any) => run.text ?? '').join('')
        if (!text.includes('∑')) return []
        return [
          {
            slideIndex,
            text,
            fontFamilies: runs.map((run: any) => run.srcFontFamily ?? run.fontFamily),
          },
        ]
      }),
    )
    return {
      mathNodes: mathNodes.length,
      mathLatex: mathNodes[0]?.equationLatex ?? null,
      symbols,
    }
  })
}

async function saveCurrentDeck(page: Page): Promise<SaveResult> {
  return page.evaluate(() => (window as unknown as SlidesWindow).slidesApi!.save())
}

/** Inject a legal-but-unhandled OMML phantom so the imported-equation edit path is read-only. */
async function makeUnsupportedEquationCopy(source: string, target: string): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(source))
  const slide = (await packageSlides(source)).find((candidate) => candidate.xml.includes('<a14:m'))
  if (!slide) throw new Error('Cannot make unsupported import fixture without an equation')
  const file = zip.file(slide.path)
  if (!file) throw new Error(`Missing ${slide.path}`)
  const xml = await file.async('string')
  const shape = (xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) ?? []).find((item) =>
    item.includes('<a14:m'),
  )
  if (!shape) throw new Error('Equation shape extraction failed')
  const mathRun = /<m:r\b[\s\S]*?<\/m:r>/.exec(shape)?.[0]
  if (!mathRun) throw new Error('Equation has no OMML run to wrap')
  const phantomRun = `<m:phant><m:e>${mathRun}</m:e></m:phant>`
  const changedShape = shape.replace(mathRun, phantomRun)
  zip.file(slide.path, xml.replace(shape, changedShape))
  await writeFile(target, await zip.generateAsync({ type: 'nodebuffer' }))
}

test('Slides math and symbols: native insert, editable roundtrip, stable shape and imported preservation', async () => {
  test.setTimeout(300_000)
  const dir = await evidenceDir()
  const working = join(dir, 'math-symbol-roundtrip.pptx')
  const unsupported = join(dir, 'unsupported-import.pptx')
  const evidencePath = join(dir, 'slides-math-verification.json')
  const screenshot = join(dir, 'equation-canvas.png')
  await copyFile(FIXTURE, working)

  const evidence: Record<string, unknown> = {
    scenario:
      'Ribbon equation authoring with an internal math symbol; native PPTX save/relaunch/edit/relaunch; chosen-font symbol persistence; unsupported imported OMML preservation',
    startedAt: new Date().toISOString(),
    workingPptx: working,
    unsupportedPptx: unsupported,
    evidencePath,
    screenshot,
    observations: {},
  }
  let launched: LaunchedApp | undefined

  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: working,
      videoDir: 'slides-math-verification-2026-10-13-insert',
    })
    let page = await openEditor(launched)
    await expect(page.locator('.slide-list .thumb').first()).toBeVisible()
    await openRibbonTab(page, 'Insert')
    await clickRibbonCommand(page, 'ribbon-equation')
    const dialog = page.getByTestId('equation-dialog')
    await expect(dialog).toBeVisible()
    const input = dialog.locator('.go-math-input')
    await input.fill('\\frac{x^{2}+\\sqrt{y}}{n+1}')
    await input.evaluate((element: HTMLTextAreaElement) => {
      element.focus()
      element.setSelectionRange(element.value.length, element.value.length)
    })
    await dialog.getByRole('tab', { name: 'Advanced Symbol' }).click()
    await dialog
      .locator('.go-math-symbol-subsets')
      .getByRole('tab', { name: 'Greek', exact: true })
      .click()
    await dialog
      .locator('.go-math-symbol-grid button[aria-label="α GREEK SMALL LETTER ALPHA U+03B1"]')
      .click()
    await expect(input).toHaveValue(INITIAL_LATEX)
    const dialogMathML = await dialog.locator('.go-math-preview math').innerHTML()
    expect(dialogMathML).toContain('<mfrac')
    expect(dialogMathML).toContain('<msqrt')
    expect(dialogMathML).toContain('<msup')
    expect(dialogMathML).toContain('α')
    await expect(dialog.getByTestId('equation-commit')).toBeEnabled()
    await page
      .locator('.go-math-preview')
      .screenshot({ path: join(dir, 'equation-dialog-preview.png') })
    await dialog.getByTestId('equation-commit').click()
    await expect(dialog).toHaveCount(0)

    const canvasMathML = await page.evaluate(async () => {
      const slide = (await window.slidesApi.getRenderSlides())![0]
      return (
        (slide.nodes.find((candidate: any) => candidate.hasEquation) as any)?.equationMathML ?? ''
      )
    })
    expect(canvasMathML).toContain('<mfrac')
    expect(canvasMathML).toContain('<msqrt')
    expect(canvasMathML).toContain('<msup')
    expect(canvasMathML).toContain('α')

    // Put the ordinary symbol on a second slide so the canvas evidence keeps the
    // standalone equation unobscured. It must be a text shape with an explicit font.
    await page.locator('.slide-list .thumb').nth(1).click()
    await clickRibbonCommand(page, 'ribbon-symbol')
    const symbolDialog = page.getByTestId('symbol-dialog')
    await expect(symbolDialog).toBeVisible()
    await symbolDialog.locator('.go-symbol-picker select').nth(0).selectOption(SYMBOL_FONT)
    await symbolDialog.locator('.go-symbol-picker select').nth(1).selectOption('math')
    await symbolDialog.locator(`[role="option"][title="U+2211"]`).click()
    await symbolDialog.locator('.go-symbol-insert').click()
    await expect(symbolDialog).toHaveCount(0)
    await page.locator('.slide-list .thumb').first().click()

    const initialSave = await saveCurrentDeck(page)
    expect(initialSave).toMatchObject({ ok: true })
    const initialEquation = await equationShape(working)
    const initialSymbol = await symbolShape(working)
    expect(initialEquation.mathXml).toContain('<a14:m>')
    expect(initialEquation.mathXml).toContain('<m:oMathPara')
    expect(initialEquation.mathXml).toContain('<m:f>')
    expect(initialEquation.mathXml).toContain('<m:rad>')
    expect(initialEquation.mathXml).toContain('<mc:Fallback>')
    expect(initialEquation.mathXml).toContain('α')
    expect(initialSymbol.slidePath).not.toBe(initialEquation.slidePath)
    expect(initialSymbol.shapeXml).toContain(`typeface="${SYMBOL_FONT}"`)
    evidence.observations = {
      ...(evidence.observations as Record<string, unknown>),
      dialogMathML,
      canvasMathML,
      initialSave: { ok: initialSave.ok, path: initialSave.path },
      initialEquation: {
        slidePath: initialEquation.slidePath,
        shapeId: initialEquation.shapeId,
        hasPowerPointMath: initialEquation.mathXml.includes('<a14:m>'),
        hasOmml: initialEquation.mathXml.includes('<m:oMathPara'),
        hasLinearFallback: initialEquation.mathXml.includes('<mc:Fallback>'),
      },
      symbol: {
        slidePath: initialSymbol.slidePath,
        chosenFont: SYMBOL_FONT,
        persistedFont: initialSymbol.shapeXml.includes(`typeface="${SYMBOL_FONT}"`),
      },
    }

    // First full process restart: open the actual saved file and inspect its
    // render tree/font before exercising both selected and double-click editing.
    launched = await closeAndOpenFile(
      launched,
      working,
      'slides-math-verification-2026-10-13-roundtrip1',
    )
    page = launched.page
    await openRibbonTab(page, 'Insert')
    const reopenedRender = await renderObservations(page)
    expect(reopenedRender.mathNodes).toBe(1)
    expect(reopenedRender.mathLatex).toContain('\\frac')
    expect(reopenedRender.mathLatex).toContain('\\sqrt')
    expect(reopenedRender.mathLatex).toContain('\\alpha')
    expect(reopenedRender.symbols).toHaveLength(1)
    expect(reopenedRender.symbols[0]?.fontFamilies).toContain(SYMBOL_FONT)
    await page.locator('.slide-list .thumb').first().click()
    await clickEquation(page)
    await clickRibbonCommand(page, 'ribbon-equation-edit')
    const selectedEdit = page.getByTestId('equation-dialog')
    await expect(selectedEdit).toBeVisible()
    const selectedInput = selectedEdit.locator('.go-math-input')
    await expect(selectedInput).toHaveValue(reopenedRender.mathLatex!)
    await selectedInput.fill(EDITED_LATEX)
    const editedMathML = await selectedEdit.locator('.go-math-preview math').innerHTML()
    expect(editedMathML).toContain('<mfrac')
    expect(editedMathML).toContain('<msqrt')
    expect(editedMathML).toContain('z')
    await selectedEdit.getByTestId('equation-commit').click()
    await expect(selectedEdit).toHaveCount(0)
    expect(await currentCanvasMathML(page)).toContain('z')

    // A canvas double-click must reopen that same live equation and its latest source.
    const bounds = await equationBounds(page)
    evidence.doubleClickTarget = await page.evaluate(
      ({ x, y }) => ({
        x,
        y,
        element: document.elementFromPoint(x, y)?.outerHTML.slice(0, 600),
        canvasBounds: [...document.querySelectorAll('canvas')].map((e) => ({
          rect: e.getBoundingClientRect().toJSON(),
          className: e.className,
        })),
      }),
      { x: bounds.x + bounds.width / 2, y: bounds.y + bounds.height / 2 },
    )
    await page.screenshot({ path: join(dir, 'before-double-click.png') })
    await page.mouse.dblclick(bounds.x + bounds.width / 2, bounds.y + bounds.height / 2)
    await page.screenshot({ path: join(dir, 'after-double-click.png') })
    const doubleClickEdit = page.getByTestId('equation-dialog')
    await expect(doubleClickEdit).toBeVisible()
    const doubleClickSource = await doubleClickEdit.locator('.go-math-input').inputValue()
    expect(doubleClickSource).toContain('\\frac')
    expect(doubleClickSource).toContain('\\sqrt')
    expect(doubleClickSource).toContain('z')
    expect(doubleClickSource).toContain('\\alpha')
    await doubleClickEdit.getByRole('button', { name: 'Cancel' }).click()
    await expect(doubleClickEdit).toHaveCount(0)

    const secondSave = await saveCurrentDeck(page)
    expect(secondSave).toMatchObject({ ok: true })
    const afterEditEquation = await equationShape(working)
    expect(afterEditEquation.shapeId).toBe(initialEquation.shapeId)
    expect(afterEditEquation.mathXml).toContain('<m:f>')
    expect(afterEditEquation.mathXml).toContain('z')
    expect(afterEditEquation.mathXml).toContain('α')

    // Second full process restart proves both edits committed as native math and
    // still build the actual editable MathML canvas preview.
    launched = await closeAndOpenFile(
      launched,
      working,
      'slides-math-verification-2026-10-13-roundtrip2',
    )
    page = launched.page
    await page.locator('.slide-list .thumb').first().click()
    expect(await currentCanvasMathML(page)).toContain('z')
    const finalRender = await renderObservations(page)
    expect(finalRender.mathNodes).toBe(1)
    expect(finalRender.mathLatex).toContain('z')
    expect(finalRender.symbols).toHaveLength(1)
    expect(finalRender.symbols[0]?.fontFamilies).toContain(SYMBOL_FONT)
    await page.locator('.stage-rel').screenshot({ path: screenshot })
    evidence.observations = {
      ...(evidence.observations as Record<string, unknown>),
      firstReopen: reopenedRender,
      selectedEdit: { sourceBefore: reopenedRender.mathLatex, sourceAfter: EDITED_LATEX },
      doubleClickEdit: { source: doubleClickSource, reopened: true },
      secondSave: { ok: secondSave.ok, path: secondSave.path },
      afterEditEquation: {
        slidePath: afterEditEquation.slidePath,
        shapeId: afterEditEquation.shapeId,
        shapeIdPreserved: afterEditEquation.shapeId === initialEquation.shapeId,
        hasUpdatedMath: afterEditEquation.mathXml.includes('z'),
      },
      secondReopen: finalRender,
    }

    // Unsupported imported OMML is shown as read-only. A real edit elsewhere
    // forces a deck save, after which the original unsupported math block must
    // remain byte-identical in the slide XML.
    await closeAndSaveVideo(launched, 'slides-math-verification-2026-10-13-final')
    launched = undefined
    await makeUnsupportedEquationCopy(working, unsupported)
    const beforeUnsupported = await equationShape(unsupported)
    expect(beforeUnsupported.mathXml).toContain('<m:phant>')
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: unsupported,
      videoDir: 'slides-math-verification-2026-10-13-unsupported',
    })
    page = await openEditor(launched)
    await openRibbonTab(page, 'Insert')
    await page.locator('.slide-list .thumb').first().click()
    await clickEquation(page)
    await clickRibbonCommand(page, 'ribbon-equation-edit')
    const unsupportedDialog = page.getByTestId('equation-dialog')
    await expect(unsupportedDialog).toBeVisible()
    await expect(unsupportedDialog.getByRole('status')).toBeVisible()
    await expect(unsupportedDialog.getByTestId('equation-commit')).toBeDisabled()
    await unsupportedDialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(unsupportedDialog).toHaveCount(0)

    // Add a separate ordinary symbol to mark the imported document dirty and
    // force serialization of its untouched equation bytes.
    await page.locator('.slide-list .thumb').nth(1).click()
    await clickRibbonCommand(page, 'ribbon-symbol')
    const secondSymbolDialog = page.getByTestId('symbol-dialog')
    await expect(secondSymbolDialog).toBeVisible()
    await secondSymbolDialog.locator('.go-symbol-picker select').nth(0).selectOption(SYMBOL_FONT)
    await secondSymbolDialog.locator('.go-symbol-picker select').nth(1).selectOption('math')
    await secondSymbolDialog.locator(`[role="option"][title="U+2211"]`).click()
    await secondSymbolDialog.locator('.go-symbol-insert').click()
    await expect(secondSymbolDialog).toHaveCount(0)
    const unsupportedSave = await saveCurrentDeck(page)
    expect(unsupportedSave).toMatchObject({ ok: true })
    const afterUnsupported = await equationShape(unsupported)
    expect(afterUnsupported.mathXml).toBe(beforeUnsupported.mathXml)
    evidence.observations = {
      ...(evidence.observations as Record<string, unknown>),
      unsupportedImport: {
        readOnly: true,
        disabledCommit: true,
        savedAfterUnrelatedEdit: unsupportedSave.ok,
        originalMathBlockPreservedByteForByte:
          afterUnsupported.mathXml === beforeUnsupported.mathXml,
        shapeId: afterUnsupported.shapeId,
      },
    }
  } finally {
    evidence.finishedAt = new Date().toISOString()
    await writeFile(evidencePath, `${JSON.stringify(evidence, null, 2)}\n`).catch(() => {})
    if (launched) await closeAndSaveVideo(launched, 'slides-math-verification-2026-10-13-cleanup')
  }

  evidence.evidencePath = evidencePath
  await writeFile(evidencePath, `${JSON.stringify(evidence, null, 2)}\n`)
})
