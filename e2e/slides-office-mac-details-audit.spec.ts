import { expect, test, type Page } from '@playwright/test'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Focused Slides audit probes for the Mac Office parity review. These tests use
 * scratch PPTX files, write a JSON observation record, and deliberately assert
 * the expected product semantics after recording what the current build did.
 * Known gaps therefore show up as useful, reproducible failures rather than
 * passing merely because a missing control stayed missing.
 */

const FIXED_DATE = 'January 2, 2001'
const FOOTER_TEXT = 'GENOFFICE_SLIDES_MAC_AUDIT_FOOTER_2026-10-12'
const EQUATION_TEXT = 'x² + y² = z²'

test.beforeEach(() => {
  test.skip(
    !process.env.GENOFFICE_OFFICE_AUDIT,
    'Opt-in Office audit; set GENOFFICE_OFFICE_AUDIT=1 to run',
  )
})

interface SlidesWindow {
  slidesApi?: {
    save(): Promise<{ ok: boolean; path?: string; error?: string }>
    saveAs(name: string): Promise<{ ok: boolean; path?: string; error?: string }>
    getRenderSlides(): Promise<any[] | null>
  }
}

async function openRibbonTab(page: Page, name: string): Promise<void> {
  const tab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: new RegExp(`^${name}$`) })
  await expect(tab).toBeVisible()
  await tab.click()
  await expect(tab).toHaveClass(/active/)
}

async function blankSlides(launched: LaunchedApp): Promise<Page> {
  await launched.page.locator('.quick-card').nth(2).click()
  const page = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
  await page.waitForSelector('.ribbon', { timeout: 15_000 })
  return page
}

async function auditDirectory(): Promise<string> {
  const root =
    process.env.GENOFFICE_SLIDES_AUDIT_DIR ??
    join(tmpdir(), 'genoffice-slides-office-mac-details-audit')
  await mkdir(root, { recursive: true })
  return mkdtemp(join(root, `${new Date().toISOString().replace(/[.:]/g, '-')}-`))
}

async function saveAsScratch(launched: LaunchedApp, page: Page, path: string): Promise<unknown> {
  await launched.app.evaluate(({ dialog }, file) => {
    dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
  }, path)
  return page.evaluate(() =>
    (window as unknown as SlidesWindow).slidesApi!.saveAs('slides-audit.pptx'),
  )
}

async function readSlidePackage(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const slidePaths = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name) && !zip.files[name]?.dir)
    .sort(
      (a, b) => Number(a.match(/slide(\d+)\.xml$/)?.[1]) - Number(b.match(/slide(\d+)\.xml$/)?.[1]),
    )
  const slides = await Promise.all(
    slidePaths.map(async (name) => ({
      path: name,
      xml: await zip.file(name)!.async('string'),
    })),
  )
  return { slides, slidePaths }
}

async function saveSlideXmlEvidence(
  dir: string,
  phase: string,
  slides: Array<{ xml: string }>,
): Promise<string[]> {
  return Promise.all(
    slides.map(async ({ xml }, index) => {
      const file = join(dir, `${phase}-slide-${index + 1}.xml`)
      await writeFile(file, xml)
      return file
    }),
  )
}

test('Slides Mac audit: fixed date dialog echo, Apply to All, then PPTX fields', async () => {
  test.setTimeout(240_000)
  const dir = await auditDirectory()
  const out = join(dir, 'header-footer-audit.pptx')
  const evidencePath = join(dir, 'header-footer-evidence.json')
  const evidence: Record<string, unknown> = {
    scenario: 'fixed header/footer date, save/reopen/dialog echo, Apply to All, inspect OOXML',
    startedAt: new Date().toISOString(),
    scratchPptx: out,
    observations: {},
  }
  let launched: LaunchedApp | undefined
  let reopened: LaunchedApp | undefined
  let page: Page | undefined
  const screenshot = join(dir, 'final-state.png')
  evidence.screenshot = screenshot
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      videoDir: 'slides-office-mac-details-audit-hf',
    })
    page = await blankSlides(launched)

    await openRibbonTab(page, 'Home')
    await page.locator('.ribbon-body button[data-tip^="New blank slide"]').click()
    await expect
      .poll(() =>
        page.evaluate(
          async () =>
            (await (window as unknown as SlidesWindow).slidesApi!.getRenderSlides())?.length,
        ),
      )
      .toBe(2)

    await openRibbonTab(page, 'Insert')
    const headerFooterButton = page.locator('.ribbon-body button[data-tip^="Header & footer"]')
    await expect(headerFooterButton).toHaveCount(1)
    await headerFooterButton.click()
    const dialog = page.locator('.modal', {
      has: page.getByRole('heading', { name: 'Header and Footer' }),
    })
    await expect(dialog).toBeVisible()
    const checks = dialog.locator('input[type="checkbox"]')
    await checks.nth(0).check() // Date and time
    await checks.nth(1).uncheck() // fixed date, not automatically updated
    await dialog.locator('input[type="text"]').first().fill(FIXED_DATE)
    await checks.nth(3).check() // Footer
    await dialog.locator('input[type="text"]').nth(1).fill(FOOTER_TEXT)
    const selectedBeforeSave = {
      dateEnabled: await checks.nth(0).isChecked(),
      autoUpdateEnabled: await checks.nth(1).isChecked(),
      footerEnabled: await checks.nth(3).isChecked(),
      dateText: await dialog.locator('input[type="text"]').first().inputValue(),
      footerText: await dialog.locator('input[type="text"]').nth(1).inputValue(),
      applyButton: await dialog.getByRole('button', { name: 'Apply to All' }).count(),
    }
    await dialog.getByRole('button', { name: 'Apply to All' }).click()
    await expect(dialog).toHaveCount(0)
    const initialSave = await saveAsScratch(launched, page, out)
    expect(initialSave).toMatchObject({ ok: true })
    const firstPackage = await readSlidePackage(out)
    const initialXmlEvidence = await saveSlideXmlEvidence(
      dir,
      'initial-fixed-date',
      firstPackage.slides,
    )
    const firstSaveXml = {
      slideCount: firstPackage.slides.length,
      slidesWithFixedDate: firstPackage.slides.filter(
        ({ xml }) => xml.includes(FIXED_DATE) && !/<a:fld\b[^>]*type="datetime1"/.test(xml),
      ).length,
      slidesWithDynamicDateField: firstPackage.slides.filter(({ xml }) =>
        /<a:fld\b[^>]*type="datetime1"/.test(xml),
      ).length,
      slidesWithFooter: firstPackage.slides.filter(({ xml }) => xml.includes(FOOTER_TEXT)).length,
    }
    evidence.observations = {
      selectedBeforeSave,
      initialSave,
      firstSaveXml,
      initialXmlEvidence,
    }
    expect(selectedBeforeSave).toMatchObject({
      dateEnabled: true,
      autoUpdateEnabled: false,
      footerEnabled: true,
      dateText: FIXED_DATE,
      footerText: FOOTER_TEXT,
      applyButton: 1,
    })
    expect(firstSaveXml).toMatchObject({
      slideCount: 2,
      slidesWithFixedDate: 2,
      slidesWithDynamicDateField: 0,
      slidesWithFooter: 2,
    })

    // Electron tests share a macOS GPU/cache/Dock slot; do not overlap launches.
    await closeAndSaveVideo(launched, 'slides-office-mac-details-audit-hf-initial')
    launched = undefined

    reopened = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: out,
      videoDir: 'slides-office-mac-details-audit-hf-reopen',
    })
    page = await waitForPageWithUrl(reopened.app, '://slides/', 30_000)
    await page.waitForSelector('.ribbon', { timeout: 15_000 })
    await openRibbonTab(page, 'Insert')
    const reopenedHeaderFooterButton = page.locator(
      '.ribbon-body button[data-tip^="Header & footer"]',
    )
    await expect(reopenedHeaderFooterButton).toHaveCount(1)
    await reopenedHeaderFooterButton.click()
    const reopenDialog = page.locator('.modal', {
      has: page.getByRole('heading', { name: 'Header and Footer' }),
    })
    await expect(reopenDialog).toBeVisible()
    const reopenedChecks = reopenDialog.locator('input[type="checkbox"]')
    const reopenMode = {
      dateEnabled: await reopenedChecks.nth(0).isChecked(),
      autoUpdateEnabled: await reopenedChecks.nth(1).isChecked(),
      dateText: await reopenDialog.locator('input[type="text"]').first().inputValue(),
      footerEnabled: await reopenedChecks.nth(3).isChecked(),
      footerText: await reopenDialog.locator('input[type="text"]').nth(1).inputValue(),
    }
    await reopenDialog.getByRole('button', { name: 'Apply to All' }).click()
    await expect(reopenDialog).toHaveCount(0)
    const applyAllSave = await page.evaluate(() =>
      (window as unknown as SlidesWindow).slidesApi!.save(),
    )
    expect(applyAllSave).toMatchObject({ ok: true })
    const afterApplyPackage = await readSlidePackage(out)
    const afterApplyXmlEvidence = await saveSlideXmlEvidence(
      dir,
      'after-apply-all',
      afterApplyPackage.slides,
    )
    const afterApplyXml = {
      slideCount: afterApplyPackage.slides.length,
      slidesWithFixedDate: afterApplyPackage.slides.filter(
        ({ xml }) => xml.includes(FIXED_DATE) && !/<a:fld\b[^>]*type="datetime1"/.test(xml),
      ).length,
      slidesWithDynamicDateField: afterApplyPackage.slides.filter(({ xml }) =>
        /<a:fld\b[^>]*type="datetime1"/.test(xml),
      ).length,
      slidesWithFooter: afterApplyPackage.slides.filter(({ xml }) => xml.includes(FOOTER_TEXT))
        .length,
    }
    evidence.observations = {
      selectedBeforeSave,
      initialSave,
      firstSaveXml,
      reopenMode,
      applyAllSave,
      afterApplyXml,
      initialXmlEvidence,
      afterApplyXmlEvidence,
    }
    evidence.completedAt = new Date().toISOString()
    console.log(`[SLIDES_MAC_AUDIT_HF] ${JSON.stringify(evidence.observations)}`)

    // Apply to All is a real scope assertion; the UI echo assertion catches the
    // fixed→automatic mode regression observed in the current source.
    expect(afterApplyXml.slideCount).toBe(2)
    expect(afterApplyXml.slidesWithFooter).toBe(2)
    expect(afterApplyXml.slidesWithDynamicDateField + afterApplyXml.slidesWithFixedDate).toBe(2)
    expect
      .soft(reopenMode, 'Fixed-date mode should remain fixed after save and reopen')
      .toMatchObject({
        dateEnabled: true,
        autoUpdateEnabled: false,
        dateText: FIXED_DATE,
        footerEnabled: true,
        footerText: FOOTER_TEXT,
      })
    const assertionErrors = test.info().errors.map((error) => error.message)
    evidence.status = assertionErrors.length ? 'failed' : 'passed'
    evidence.parityAssertionErrors = assertionErrors
  } catch (error) {
    evidence.status = 'failed'
    evidence.error =
      error instanceof Error
        ? `${error.name}: ${error.message}\n${error.stack ?? ''}`
        : String(error)
    throw error
  } finally {
    if (page) {
      await page.screenshot({ path: screenshot, fullPage: true }).catch((error) => {
        evidence.screenshotError = String(error)
      })
    }
    await writeFile(evidencePath, JSON.stringify(evidence, null, 2))
    console.log(`[SLIDES_MAC_AUDIT_HF_EVIDENCE] ${evidencePath}`)
    if (reopened) await closeAndSaveVideo(reopened, 'slides-office-mac-details-audit-hf-reopen')
    if (launched) await closeAndSaveVideo(launched, 'slides-office-mac-details-audit-hf')
  }
})

test('Slides Mac audit: Equation entry object and Insert Symbol control', async () => {
  test.setTimeout(180_000)
  const dir = await auditDirectory()
  const out = join(dir, 'equation-audit.pptx')
  const evidencePath = join(dir, 'equation-evidence.json')
  const evidence: Record<string, unknown> = {
    scenario: 'inspect Insert ribbon Equation/Symbol controls and equation object after PPTX save',
    startedAt: new Date().toISOString(),
    scratchPptx: out,
  }
  let launched: LaunchedApp | undefined
  let page: Page | undefined
  const screenshot = join(dir, 'final-state.png')
  evidence.screenshot = screenshot
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      videoDir: 'slides-office-mac-details-audit-equation',
    })
    page = await blankSlides(launched)
    await openRibbonTab(page, 'Insert')
    const insertControls = await page.locator('.ribbon-body button').evaluateAll((buttons) =>
      buttons.map((button) => ({
        text: (button.textContent ?? '').replace(/\s+/g, ' ').trim(),
        tip: button.getAttribute('data-tip'),
        disabled: (button as HTMLButtonElement).disabled,
      })),
    )
    const symbolControls = insertControls.filter((control) =>
      /\bsymbol\b/i.test(`${control.text} ${control.tip ?? ''}`),
    )
    const equationButton = page.locator('.ribbon-body button[data-tip^="Insert an equation"]')
    await expect(equationButton).toHaveCount(1)
    await equationButton.click()
    const dialog = page.locator('.modal', {
      has: page.getByRole('heading', { name: 'Insert Equation' }),
    })
    await expect(dialog).toBeVisible()
    const equationInput = dialog.locator('input[type="text"]')
    await equationInput.fill(EQUATION_TEXT)
    await dialog.getByRole('button', { name: 'Insert' }).click()
    await expect(dialog).toHaveCount(0)

    const renderObservation = await page.evaluate(async (text) => {
      const slides = await (window as unknown as SlidesWindow).slidesApi!.getRenderSlides()
      const slide = slides?.[0]
      const normalize = (value: string) => value.replace(/\s+/g, '')
      const observedNodes = (slide?.nodes ?? []).map((node: any) => {
        const lines = node.text?.lines ?? []
        const fragments = lines.flatMap((line: any) =>
          (line.runs ?? []).map((run: any) => ({
            text: run.text ?? '',
            fontFamily: run.fontFamily,
            italic: run.italic,
          })),
        )
        const textContent = fragments.map((fragment: any) => fragment.text).join('')
        return {
          type: node.type,
          txBox: node.txBox ?? null,
          textContent,
          matchesEquation: normalize(textContent).includes(normalize(text)),
          fragments,
        }
      })
      return {
        slideCount: slides?.length ?? 0,
        matchingNodes: observedNodes.filter((node: any) => node.matchesEquation),
        nodeObservations: observedNodes.map(
          ({ type, txBox, textContent, matchesEquation }: any) => ({
            type,
            txBox,
            textContent,
            matchesEquation,
          }),
        ),
      }
    }, EQUATION_TEXT)
    const saveResult = await saveAsScratch(launched, page, out)
    expect(saveResult).toMatchObject({ ok: true })
    const pkg = await readSlidePackage(out)
    const equationSlide = pkg.slides.find(({ xml }) => xml.includes(EQUATION_TEXT))
    const eqStart = equationSlide?.xml.indexOf(`<a:t>${EQUATION_TEXT}</a:t>`) ?? -1
    const shapeStart = eqStart >= 0 ? equationSlide!.xml.lastIndexOf('<p:sp>', eqStart) : -1
    const shapeEnd =
      eqStart >= 0 ? equationSlide!.xml.indexOf('</p:sp>', eqStart) + '</p:sp>'.length : -1
    const equationShapeXml =
      shapeStart >= 0 && shapeEnd >= 0 ? equationSlide!.xml.slice(shapeStart, shapeEnd) : ''
    const xmlEvidencePath = join(dir, 'equation-slide.xml')
    await writeFile(xmlEvidencePath, equationSlide?.xml ?? '')
    const xmlObservation = {
      slideCount: pkg.slides.length,
      slideWithEquation: equationSlide?.path ?? null,
      equationTextFound: equationSlide?.xml.includes(EQUATION_TEXT) ?? false,
      equationStoredInsideShape:
        equationShapeXml.includes('<p:sp>') &&
        equationShapeXml.includes(`<a:t>${EQUATION_TEXT}</a:t>`),
      equationShapeUsesCambriaMath: equationShapeXml.includes('Cambria Math'),
      hasOmmlOrMathObject: pkg.slides.some(({ xml }) => /<m:oMath\b|<a14:m\b/.test(xml)),
      equationShapePreview: equationShapeXml.slice(0, 1200),
      rawSlideXml: xmlEvidencePath,
      symbolControlCount: symbolControls.length,
      symbolControls,
    }
    evidence.observations = {
      insertControls,
      symbolControls,
      renderObservation,
      saveResult,
      xmlObservation,
    }
    evidence.completedAt = new Date().toISOString()
    console.log(`[SLIDES_MAC_AUDIT_EQUATION] ${JSON.stringify(evidence.observations)}`)

    // Baseline object/save checks keep the probe useful even before parity lands.
    expect(xmlObservation.slideCount).toBeGreaterThan(0)
    expect(xmlObservation.equationStoredInsideShape).toBe(true)
    expect(renderObservation.matchingNodes.length).toBeGreaterThan(0)
    // These are product acceptance checks: current known gaps should be logged as
    // real failures, not inverted into passing assertions that bless the gaps.
    expect
      .soft(xmlObservation.symbolControlCount, 'Insert ribbon should expose a Symbol picker')
      .toBeGreaterThan(0)
    expect
      .soft(
        xmlObservation.hasOmmlOrMathObject,
        'Equation insertion should save a structured math object',
      )
      .toBe(true)
    const assertionErrors = test.info().errors.map((error) => error.message)
    evidence.status = assertionErrors.length ? 'failed' : 'passed'
    evidence.parityAssertionErrors = assertionErrors
  } catch (error) {
    evidence.status = 'failed'
    evidence.error =
      error instanceof Error
        ? `${error.name}: ${error.message}\n${error.stack ?? ''}`
        : String(error)
    throw error
  } finally {
    if (page) {
      await page.screenshot({ path: screenshot, fullPage: true }).catch((error) => {
        evidence.screenshotError = String(error)
      })
    }
    await writeFile(evidencePath, JSON.stringify(evidence, null, 2))
    console.log(`[SLIDES_MAC_AUDIT_EQUATION_EVIDENCE] ${evidencePath}`)
    if (launched) await closeAndSaveVideo(launched, 'slides-office-mac-details-audit-equation')
  }
})
