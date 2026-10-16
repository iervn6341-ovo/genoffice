import { expect, test, type Page } from '@playwright/test'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, posix, resolve } from 'node:path'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

const FIXTURE = join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const FIXED_DATE = 'January 2, 2001'
const CURRENT_FOOTER = 'HF_CURRENT_ONLY_2026-10-13'
const ALL_FOOTER = 'HF_APPLY_ALL_2026-10-13'
const MASTER_FOOTER = 'HF_MASTER_TEMPLATE_2026-10-13'
const MASTER_DATE_CACHE = '1/27/13'
const MASTER_LINK = 'https://example.com/inherited-header-footer'

test.beforeEach(() => {
  test.skip(
    !process.env.GENOFFICE_SLIDES_HF_VERIFICATION,
    'Opt-in Slides Header/Footer verification; set GENOFFICE_SLIDES_HF_VERIFICATION=1 to run',
  )
})

interface SlidesWindow {
  slidesApi?: {
    save(): Promise<SaveResult>
    saveAs(name: string): Promise<SaveResult>
    getRenderSlides(): Promise<any[] | null>
  }
}

interface SaveResult {
  ok: boolean
  path?: string
  error?: string
  slides?: unknown
}

interface PackageSlide {
  path: string
  xml: string
}

async function evidenceDir(): Promise<string> {
  const root =
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ??
    join(tmpdir(), 'genoffice-slides-office-mac-verification-2026-10-13')
  const absoluteRoot = resolve(root)
  await mkdir(absoluteRoot, { recursive: true })
  return mkdtemp(join(absoluteRoot, `${new Date().toISOString().replace(/[.:]/g, '-')}-`))
}

/** Make a private imported PPTX whose active date/footer/number come only from its master. */
async function makeInheritedMasterFixture(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(FIXTURE))
  const masterPath = 'ppt/slideMasters/slideMaster1.xml'
  const masterRelsPath = 'ppt/slideMasters/_rels/slideMaster1.xml.rels'
  let master = await zip.file(masterPath)!.async('string')
  const relsFile = zip.file(masterRelsPath)
  let rels = relsFile
    ? await relsFile.async('string')
    : '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  const ids = [...rels.matchAll(/\bId=["']rId(\d+)["']/g)].map((match) => Number(match[1]))
  const hyperlinkId = `rId${Math.max(0, ...ids) + 1}`
  rels = rels.replace(
    '</Relationships>',
    `<Relationship Id="${hyperlinkId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${MASTER_LINK}" TargetMode="External"/></Relationships>`,
  )

  const placeholderTypes = new Set(['dt', 'ftr', 'sldNum'])
  const present = new Set<string>()
  master = master.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) => {
    const placeholder = /<p:ph\b[^>]*\btype="(dt|ftr|sldNum)"/.exec(shape)?.[1]
    if (!placeholder) return shape
    present.add(placeholder)
    if (placeholder === 'dt') {
      return shape.replace(/<a:t\b[^>]*>[\s\S]*?<\/a:t>/, `<a:t>${MASTER_DATE_CACHE}</a:t>`)
    }
    if (placeholder === 'ftr') {
      const run =
        `<a:r><a:rPr lang="en-US" sz="2400"><a:solidFill><a:srgbClr val="123ABC"/></a:solidFill>` +
        `<a:hlinkClick r:id="${hyperlinkId}"/></a:rPr><a:t>${MASTER_FOOTER}</a:t></a:r>`
      return shape.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/, (paragraph) => {
        const open = /^<a:p\b[^>]*>/.exec(paragraph)?.[0] ?? '<a:p>'
        let inner = paragraph.slice(open.length, -'</a:p>'.length)
        inner = inner.replace(/<a:(r|fld)\b[^>]*>[\s\S]*?<\/a:\1>/g, '')
        const endPara = /<a:endParaRPr\b/.exec(inner)
        inner = endPara ? inner.replace(/<a:endParaRPr\b/, `${run}<a:endParaRPr`) : `${inner}${run}`
        return `${open}${inner}</a:p>`
      })
    }
    return shape
  })
  expect([...placeholderTypes].every((type) => present.has(type))).toBe(true)

  if (/<p:hf\b[^>]*\/>/.test(master)) {
    master = master.replace(/<p:hf\b[^>]*\/>/, '<p:hf dt="1" ftr="1" sldNum="1"/>')
  } else {
    expect(master).toContain('<p:txStyles')
    master = master.replace('<p:txStyles', '<p:hf dt="1" ftr="1" sldNum="1"/><p:txStyles')
  }
  zip.file(masterPath, master)
  zip.file(masterRelsPath, rels)

  // Remove same-type layout placeholders so the imported display and edits resolve
  // through the master shape and its style/relationship chain.
  const layoutPaths = Object.keys(zip.files).filter((name) =>
    /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name),
  )
  for (const layoutPath of layoutPaths) {
    let layout = await zip.file(layoutPath)!.async('string')
    layout = layout.replace(/<p:sp\b[\s\S]*?<\/p:sp>/g, (shape) =>
      /<p:ph\b[^>]*\btype="(?:dt|ftr|sldNum)"/.test(shape) ? '' : shape,
    )
    zip.file(layoutPath, layout)
  }
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function openRibbonTab(page: Page, name: string): Promise<void> {
  const tab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: new RegExp(`^${name}$`) })
  await expect(tab).toBeVisible()
  await tab.click()
  await expect(tab).toHaveClass(/active/)
}

async function openHeaderFooter(page: Page) {
  await openRibbonTab(page, 'Insert')
  const button = page.locator('.ribbon-body button[data-tip^="Header & footer"]')
  await expect(button).toHaveCount(1)
  await button.click()
  const dialog = page.locator('.modal', {
    has: page.getByRole('heading', { name: 'Header and Footer' }),
  })
  await expect(dialog).toBeVisible()
  return dialog
}

async function saveAsScratch(launched: LaunchedApp, page: Page, path: string): Promise<SaveResult> {
  await launched.app.evaluate(({ dialog }, file) => {
    dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
  }, path)
  return page.evaluate(() =>
    (window as unknown as SlidesWindow).slidesApi!.saveAs('slides-hf-verification.pptx'),
  )
}

async function packageSlides(path: string): Promise<PackageSlide[]> {
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

async function titleSlideIndexes(zip: JSZip, slides: PackageSlide[]): Promise<number[]> {
  const indexes: number[] = []
  for (let index = 0; index < slides.length; index++) {
    const slide = slides[index]!
    const number = /slide(\d+)\.xml$/.exec(slide.path)?.[1]
    if (!number) continue
    const rels = await zip.file(`ppt/slides/_rels/slide${number}.xml.rels`)?.async('string')
    const layoutRel = rels
      ?.match(/<Relationship\b[^>]*\/?\s*>/g)
      ?.find((row) => /\bType="[^"]*\/slideLayout"/.test(row))
    const target = layoutRel && /\bTarget="([^"]+)"/.exec(layoutRel)?.[1]
    if (!target) continue
    const layoutPath = posix.normalize(posix.join('ppt/slides', target))
    const layout = await zip.file(layoutPath)?.async('string')
    if (layout && /<p:sldLayout\b[^>]*\btype=["']title["']/.test(layout)) indexes.push(index)
  }
  return indexes
}

function placeholderXml(slide: PackageSlide, type: 'dt' | 'ftr' | 'sldNum'): string {
  return (
    slide.xml
      .match(/<p:sp\b[\s\S]*?<\/p:sp>/g)
      ?.find((shape) => new RegExp(`<p:ph\\b[^>]*\\btype="${type}"`).test(shape)) ?? ''
  )
}

async function decorations(page: Page, index: number): Promise<string[]> {
  return page.evaluate(async (slideIndex) => {
    const slides = await (window as unknown as SlidesWindow).slidesApi!.getRenderSlides()
    const nodes = slides?.[slideIndex]?.nodes ?? []
    return nodes
      .filter((node: any) => node.decoration)
      .map((node: any) =>
        (node.text?.lines ?? [])
          .flatMap((line: any) => (line.runs ?? []).map((run: any) => run.text ?? ''))
          .join(''),
      )
  }, index)
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
  const page = await waitForPageWithUrl(reopened.app, '://slides/', 30_000)
  await page.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
  return { ...reopened, page }
}

test('Slides Header/Footer: inherited master echo, current/all scope, date modes and title exclusion', async () => {
  test.setTimeout(300_000)
  const dir = await evidenceDir()
  const imported = join(dir, 'inherited-master-header-footer.pptx')
  const saved = join(dir, 'header-footer-roundtrip.pptx')
  const evidencePath = join(dir, 'header-footer-verification.json')
  const screenshot = join(dir, 'header-footer-final-state.png')
  const evidence: Record<string, unknown> = {
    scenario:
      'Imported master header/footer, current Apply, fixed-date reopen, dynamic format Apply to All, title exclusion reopen',
    startedAt: new Date().toISOString(),
    importedScratchPptx: imported,
    savedScratchPptx: saved,
    screenshot,
    observations: {},
  }
  let launched: LaunchedApp | undefined
  let page: Page | undefined
  try {
    await makeInheritedMasterFixture(imported)
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: imported,
      videoDir: 'slides-office-mac-hf-verification-2026-10-13',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
    await expect(page.locator('.slide-list .thumb')).toHaveCount(5)

    // The source deck has active master-only values. The dialog should echo the
    // effective values, including the imported uncommon datetime field/cache.
    const titleDecorationBefore = await decorations(page, 0)
    const titleDialog = await openHeaderFooter(page)
    const dateModeRadioBounds = await Promise.all(
      ['hf-date-auto', 'hf-date-fixed'].map((testId) =>
        titleDialog.getByTestId(testId).evaluate((input) => {
          const { width, height } = input.getBoundingClientRect()
          return { width, height }
        }),
      ),
    )
    expect(
      dateModeRadioBounds.every(
        ({ width, height }) => width > 0 && height > 0 && width <= 18 && height <= 18,
      ),
    ).toBe(true)
    const inheritedTitle = {
      dateModeRadioBounds,
      dateEnabled: await titleDialog.getByTestId('hf-date-enabled').isChecked(),
      dateAuto: await titleDialog.getByTestId('hf-date-auto').isChecked(),
      dateFormat: await titleDialog.getByTestId('hf-date-format').inputValue(),
      cachedDate: await titleDialog.getByTestId('hf-date-text').inputValue(),
      slideNumber: await titleDialog.getByTestId('hf-slide-number').isChecked(),
      footerEnabled: await titleDialog.getByTestId('hf-footer-enabled').isChecked(),
      footerText: await titleDialog.getByTestId('hf-footer-text').inputValue(),
      titleExcluded: await titleDialog.getByTestId('hf-exclude-title').isChecked(),
    }
    expect(titleDecorationBefore).toContain(MASTER_DATE_CACHE)
    expect(titleDecorationBefore).toContain(MASTER_FOOTER)
    expect(inheritedTitle).toMatchObject({
      dateEnabled: true,
      dateAuto: true,
      dateFormat: 'datetimeFigureOut',
      cachedDate: MASTER_DATE_CACHE,
      slideNumber: true,
      footerEnabled: true,
      footerText: MASTER_FOOTER,
      titleExcluded: false,
    })

    // Current-only title exclusion should create local blank masks, preserving
    // the other slides' inherited master content without writing slide-level p:hf.
    await titleDialog.getByTestId('hf-exclude-title').check()
    await titleDialog.getByTestId('hf-apply-current').click()
    await expect(titleDialog).toHaveCount(0)
    const titleDecorationAfterCurrent = await decorations(page, 0)
    evidence.observations = {
      inheritedTitle,
      titleDecorationBefore,
      titleDecorationAfterCurrent,
    }
    expect(titleDecorationAfterCurrent).not.toContain(MASTER_DATE_CACHE)
    expect(titleDecorationAfterCurrent).not.toContain(MASTER_FOOTER)

    // Apply to the second slide alone: inherited geometry/style and the footer
    // hyperlink should be materialized on that slide, while peers keep the master.
    await page.locator('.slide-list .thumb').nth(1).click()
    const contentDecorationBefore = await decorations(page, 1)
    const currentDialog = await openHeaderFooter(page)
    await expect(currentDialog.getByTestId('hf-date-enabled')).toBeChecked()
    await expect(currentDialog.getByTestId('hf-date-auto')).toBeChecked()
    await expect(currentDialog.getByTestId('hf-date-format')).toHaveValue('datetimeFigureOut')
    await expect(currentDialog.getByTestId('hf-footer-text')).toHaveValue(MASTER_FOOTER)
    await currentDialog.getByTestId('hf-date-fixed').check()
    await currentDialog.getByTestId('hf-date-text').fill(FIXED_DATE)
    await currentDialog.getByTestId('hf-footer-text').fill(CURRENT_FOOTER)
    await currentDialog.getByTestId('hf-apply-current').click()
    await expect(currentDialog).toHaveCount(0)
    const contentDecorationAfterCurrent = await decorations(page, 1)
    const currentSaveResult = await saveAsScratch(launched, page, saved)
    expect(currentSaveResult).toMatchObject({ ok: true })
    const currentSave = { ok: currentSaveResult.ok, path: currentSaveResult.path }
    let currentSlides = await packageSlides(saved)
    const currentTitle = currentSlides[0]!
    const currentContent = currentSlides[1]!
    const currentXml = {
      titleHasLocalMasks: ['dt', 'ftr', 'sldNum'].every((type) =>
        placeholderXml(currentTitle, type as 'dt' | 'ftr' | 'sldNum'),
      ),
      contentFixedDate: placeholderXml(currentContent, 'dt').includes(FIXED_DATE),
      contentDateIsDynamic: /<a:fld\b[^>]*type="datetime/.test(
        placeholderXml(currentContent, 'dt'),
      ),
      contentFooter: placeholderXml(currentContent, 'ftr').includes(CURRENT_FOOTER),
      untouchedSlideHasNoCurrentFooter: !currentSlides[2]!.xml.includes(CURRENT_FOOTER),
      slideHasNoSlideLevelHfToggle: !currentTitle.xml.includes('<p:hf'),
      contentGeometryMatchesMaster:
        /<a:off x="3124200" y="6356350"\/><a:ext cx="2895600" cy="365125"\/>/.test(
          placeholderXml(currentContent, 'ftr'),
        ),
      contentKeepsMasterTextStyle: /sz="2400"[\s\S]*?val="123ABC"/.test(
        placeholderXml(currentContent, 'ftr'),
      ),
    }
    const savedZip = await JSZip.loadAsync(await readFile(saved))
    const contentRels = await savedZip.file('ppt/slides/_rels/slide2.xml.rels')!.async('string')
    const footerHref = /<a:hlinkClick r:id="([^"]+)"/.exec(
      placeholderXml(currentContent, 'ftr'),
    )?.[1]
    const footerRelationship = footerHref
      ? new RegExp(`<Relationship\\b(?=[^>]*\\bId="${footerHref}")[^>]*/>`).exec(contentRels)?.[0]
      : undefined
    const relationshipRemapped =
      !!footerRelationship &&
      footerRelationship.includes(
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"',
      ) &&
      footerRelationship.includes(`Target="${MASTER_LINK}"`) &&
      footerRelationship.includes('TargetMode="External"')
    evidence.observations = {
      ...(evidence.observations as Record<string, unknown>),
      contentDecorationBefore,
      contentDecorationAfterCurrent,
      currentSave,
      currentXml,
      relationshipRemapped,
      footerHref: footerHref ?? null,
    }
    expect(contentDecorationBefore).toContain(MASTER_FOOTER)
    expect(contentDecorationAfterCurrent).not.toContain(MASTER_FOOTER)
    expect(currentXml).toMatchObject({
      titleHasLocalMasks: true,
      contentFixedDate: true,
      contentDateIsDynamic: false,
      contentFooter: true,
      untouchedSlideHasNoCurrentFooter: true,
      slideHasNoSlideLevelHfToggle: true,
      contentGeometryMatchesMaster: true,
      contentKeepsMasterTextStyle: true,
    })
    expect(relationshipRemapped).toBe(true)

    // Save, close, and reopen: fixed mode and literal date must survive in UI.
    launched = await closeAndOpenFile(launched, saved, 'slides-office-mac-hf-fixed-2026-10-13')
    page = launched.page
    await page.locator('.slide-list .thumb').first().click()
    const titleMaskReopenDialog = await openHeaderFooter(page)
    const currentTitleMaskReopen = {
      dateEnabled: await titleMaskReopenDialog.getByTestId('hf-date-enabled').isChecked(),
      slideNumber: await titleMaskReopenDialog.getByTestId('hf-slide-number').isChecked(),
      footerEnabled: await titleMaskReopenDialog.getByTestId('hf-footer-enabled').isChecked(),
      deckTitleExclusion: await titleMaskReopenDialog.getByTestId('hf-exclude-title').isChecked(),
      inheritedStillHidden: !(await decorations(page, 0)).includes(MASTER_FOOTER),
    }
    await titleMaskReopenDialog.getByRole('button', { name: 'Cancel' }).click()
    expect(currentTitleMaskReopen).toMatchObject({
      dateEnabled: false,
      slideNumber: false,
      footerEnabled: false,
      deckTitleExclusion: false,
      inheritedStillHidden: true,
    })
    await page.locator('.slide-list .thumb').nth(1).click()
    const fixedReopenDialog = await openHeaderFooter(page)
    const fixedReopen = {
      dateEnabled: await fixedReopenDialog.getByTestId('hf-date-enabled').isChecked(),
      fixed: await fixedReopenDialog.getByTestId('hf-date-fixed').isChecked(),
      dateText: await fixedReopenDialog.getByTestId('hf-date-text').inputValue(),
      formatSelectorCount: await fixedReopenDialog.getByTestId('hf-date-format').count(),
      footer: await fixedReopenDialog.getByTestId('hf-footer-text').inputValue(),
      titleDecorationStillHidden: !(await decorations(page, 0)).includes(MASTER_FOOTER),
    }
    expect(fixedReopen).toMatchObject({
      dateEnabled: true,
      fixed: true,
      dateText: FIXED_DATE,
      formatSelectorCount: 0,
      footer: CURRENT_FOOTER,
      titleDecorationStillHidden: true,
    })

    // Select an automatic format, confirm that its cached preview updates, then
    // Apply to All with title exclusion and verify the stored fields after reopen.
    await fixedReopenDialog.getByTestId('hf-date-auto').check()
    await fixedReopenDialog.getByTestId('hf-date-format').selectOption('datetime10')
    await fixedReopenDialog.getByTestId('hf-footer-text').fill(ALL_FOOTER)
    await fixedReopenDialog.getByTestId('hf-exclude-title').check()
    const cachedAutoDate = await fixedReopenDialog.getByTestId('hf-date-text').inputValue()
    const previewAutoDate = await fixedReopenDialog.getByTestId('hf-preview-date').textContent()
    expect(cachedAutoDate.trim().length).toBeGreaterThan(0)
    expect(previewAutoDate?.trim()).toBe(cachedAutoDate)
    await fixedReopenDialog.getByTestId('hf-apply-all').click()
    await expect(fixedReopenDialog).toHaveCount(0)
    const applyAllSave = await page.evaluate(() =>
      (window as unknown as SlidesWindow).slidesApi!.save(),
    )
    expect(applyAllSave).toMatchObject({ ok: true })
    const allSlides = await packageSlides(saved)
    const allZip = await JSZip.loadAsync(await readFile(saved))
    const titleIndexes = await titleSlideIndexes(allZip, allSlides)
    const contentIndexes = allSlides
      .map((_, index) => index)
      .filter((index) => !titleIndexes.includes(index))
    const titleLayoutPaths = new Set<string>()
    for (const titleIndex of titleIndexes) {
      const slideNumber = /slide(\d+)\.xml$/.exec(allSlides[titleIndex]!.path)?.[1]
      if (!slideNumber) continue
      const rels = await allZip
        .file(`ppt/slides/_rels/slide${slideNumber}.xml.rels`)
        ?.async('string')
      const layoutRel = rels
        ?.match(/<Relationship\b[^>]*\/?\s*>/g)
        ?.find((row) => /\bType="[^"]*\/slideLayout"/.test(row))
      const target = layoutRel && /\bTarget="([^"]+)"/.exec(layoutRel)?.[1]
      if (target) titleLayoutPaths.add(posix.normalize(posix.join('ppt/slides', target)))
    }
    const titleLayoutFlags = await Promise.all(
      [...titleLayoutPaths].map(async (path) => allZip.file(path)?.async('string')),
    )
    const applyAllXml = {
      slides: allSlides.length,
      titleIndexes,
      contentIndexes,
      contentSlidesDynamicDate: contentIndexes.every((index) =>
        /<a:fld\b[^>]*type="datetime10"/.test(placeholderXml(allSlides[index]!, 'dt')),
      ),
      contentSlidesFooter: contentIndexes.every((index) =>
        placeholderXml(allSlides[index]!, 'ftr').includes(ALL_FOOTER),
      ),
      titleHasBlankLocalMasks: titleIndexes.every((index) =>
        ['dt', 'ftr', 'sldNum'].every((type) =>
          /<a:t\b[^>]*><\/a:t>/.test(
            placeholderXml(allSlides[index]!, type as 'dt' | 'ftr' | 'sldNum'),
          ),
        ),
      ),
      titleLayoutsSuppressFooterFamily: titleLayoutFlags.every(
        (layout) => !!layout && /<p:hf\b[^>]*dt="0"[^>]*ftr="0"[^>]*sldNum="0"/.test(layout),
      ),
      cachedDateMatchesUi: contentIndexes.every((index) =>
        allSlides[index]!.xml.includes(cachedAutoDate),
      ),
    }
    evidence.observations = {
      ...(evidence.observations as Record<string, unknown>),
      currentTitleMaskReopen,
      fixedReopen,
      cachedAutoDate,
      previewAutoDate,
      applyAllSave,
      applyAllXml,
    }
    expect(applyAllXml).toMatchObject({
      slides: 5,
      titleIndexes: [0, 4],
      contentIndexes: [1, 2, 3],
      contentSlidesDynamicDate: true,
      contentSlidesFooter: true,
      titleHasBlankLocalMasks: true,
      titleLayoutsSuppressFooterFamily: true,
      cachedDateMatchesUi: true,
    })

    launched = await closeAndOpenFile(launched, saved, 'slides-office-mac-hf-auto-2026-10-13')
    page = launched.page
    await page.locator('.slide-list .thumb').first().click()
    const finalTitleDialog = await openHeaderFooter(page)
    const finalTitle = {
      excluded: await finalTitleDialog.getByTestId('hf-exclude-title').isChecked(),
      dateEnabled: await finalTitleDialog.getByTestId('hf-date-enabled').isChecked(),
      auto: await finalTitleDialog.getByTestId('hf-date-auto').isChecked(),
      format: await finalTitleDialog.getByTestId('hf-date-format').inputValue(),
      cachedDate: await finalTitleDialog.getByTestId('hf-date-text').inputValue(),
      footer: await finalTitleDialog.getByTestId('hf-footer-text').inputValue(),
      previewDateCount: await finalTitleDialog.getByTestId('hf-preview-date').count(),
      inheritedMasterHidden: !(await decorations(page, 0)).includes(MASTER_FOOTER),
    }
    expect(finalTitle).toMatchObject({
      excluded: true,
      dateEnabled: true,
      auto: true,
      format: 'datetime10',
      cachedDate: cachedAutoDate,
      footer: ALL_FOOTER,
      previewDateCount: 0,
      inheritedMasterHidden: true,
    })
    evidence.observations = {
      ...(evidence.observations as Record<string, unknown>),
      finalTitle,
    }
    evidence.completedAt = new Date().toISOString()
    evidence.status = 'passed'
    console.log(`[SLIDES_HF_VERIFICATION] ${JSON.stringify(evidence.observations)}`)
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
    evidence.completedAt ??= new Date().toISOString()
    await writeFile(evidencePath, JSON.stringify(evidence, null, 2))
    console.log(`[SLIDES_HF_VERIFICATION_EVIDENCE] ${evidencePath}`)
    if (launched) await closeAndSaveVideo(launched, 'slides-office-mac-hf-verification-2026-10-13')
  }
})
