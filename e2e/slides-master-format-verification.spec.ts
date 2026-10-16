import { expect, test, type Page } from '@playwright/test'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { createRequire } from 'node:module'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_SLIDES_MASTER_VERIFICATION,
  'Opt-in native master format verification',
)

const fixture = join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const req = createRequire(join(process.cwd(), 'package.json'))

async function openMaster(page: Page, kind: 'notes' | 'handout'): Promise<void> {
  await page
    .locator('.ribbon-tabs .ribbon-tab')
    .filter({ hasText: /^View$/ })
    .click()
  await page.getByTestId(kind === 'notes' ? 'view-notes-master' : 'view-handout-master').click()
  await expect(page.locator('.master-view')).toBeVisible()
}

async function zipXml(path: string, part: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  return (await zip.file(part)?.async('string')) ?? ''
}

async function seedNativeMasterArtwork(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const gradient =
    '<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="E6EFFA"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFF2D9"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>'
  const picture =
    '<p:pic><p:nvPicPr><p:cNvPr id="999" name="Native Master Logo"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId999"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="998" name="Native Logo Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="914400" y="365760"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>${picture}</p:grpSp>`
  for (const kind of ['notes', 'handout'] as const) {
    const stem = kind === 'notes' ? 'notesMaster' : 'handoutMaster'
    const dir = kind === 'notes' ? 'notesMasters' : 'handoutMasters'
    const part = `ppt/${dir}/${stem}1.xml`
    const xml = (await zip.file(part)!.async('string'))
      .replace('<p:cSld>', `<p:cSld>${gradient}`)
      .replace('</p:spTree>', `${group}</p:spTree>`)
    zip.file(part, xml)
    const relsPath = `ppt/${dir}/_rels/${stem}1.xml.rels`
    const rels =
      (await zip.file(relsPath)?.async('string')) ??
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    zip.file(
      relsPath,
      rels.replace(
        '</Relationships>',
        '<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>',
      ),
    )
  }
  expect(zip.file('ppt/media/image1.png')).toBeTruthy()
  const notesPart = 'ppt/notesSlides/notesSlide1.xml'
  const tabRuns = (text: string) =>
    `<a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r>`
  const tabs =
    '<a:tabLst><a:tab pos="914400" algn="l"/><a:tab pos="1828800" algn="r"/><a:tab pos="2743200" algn="ctr"/><a:tab pos="3657600" algn="dec"/></a:tabLst>'
  const paragraphs =
    `<a:p><a:pPr algn="l" marL="0" marR="0" indent="0">${tabs}</a:pPr>${tabRuns('TAB_ORIGIN\tTAB_LEFT\tTAB_RIGHT\tTAB_CENTER\t12.34')}</a:p>` +
    `<a:p><a:pPr algn="r" rtl="1" marL="0" marR="0" indent="0">${tabs}</a:pPr>${tabRuns('RTL_ORIGIN\tRTL_LEFT\tRTL_RIGHT')}</a:p>` +
    `<a:p><a:pPr marL="228600" indent="-171450"><a:buBlip><a:blip r:embed="rIdBulletProbe"/></a:buBlip></a:pPr>${tabRuns('PICTURE_BULLET_NATIVE')}</a:p>`
  zip.file(
    notesPart,
    (await zip.file(notesPart)!.async('string')).replace('</p:txBody>', `${paragraphs}</p:txBody>`),
  )
  const notesRelsPath = 'ppt/notesSlides/_rels/notesSlide1.xml.rels'
  zip.file(
    notesRelsPath,
    (await zip.file(notesRelsPath)!.async('string')).replace(
      '</Relationships>',
      '<Relationship Id="rIdBulletProbe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>',
    ),
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function saveAs(launched: LaunchedApp, page: Page, path: string): Promise<void> {
  await launched.app.evaluate(({ dialog }, target) => {
    ;(globalThis as any).__pdfSaveDialogProbe = { calls: 0, target }
    dialog.showSaveDialog = (async () => {
      ;(globalThis as any).__pdfSaveDialogProbe.calls++
      return { canceled: false, filePath: target }
    }) as never
  }, path)
  expect((await page.evaluate(() => window.slidesApi.saveAs('master-formatted.pptx'))).ok).toBe(
    true,
  )
}

async function printPdf(
  launched: LaunchedApp,
  page: Page,
  layout: 'notes' | 'handout2',
  path: string,
): Promise<string> {
  await expect(page.locator('.master-view')).not.toBeVisible()
  await expect(page.locator('.print-dialog')).not.toBeVisible()
  page.on('pageerror', (error) => console.error('PDF renderer error:', String(error)))
  await launched.app.evaluate(({ Menu, BrowserWindow }) => {
    const walk = (menu: Electron.Menu | null | undefined): Electron.MenuItem | undefined => {
      for (const item of menu?.items ?? []) {
        if (/^(CommandOrControl|CmdOrCtrl|Command|Cmd)\+P$/i.test(String(item.accelerator ?? '')))
          return item
        const nested = walk(item.submenu)
        if (nested) return nested
      }
    }
    walk(Menu.getApplicationMenu())?.click(
      undefined,
      BrowserWindow.getAllWindows()[0],
      undefined as never,
    )
  })
  await expect(page.locator('.print-dialog')).toBeVisible()
  await page.getByTestId('print-layout').selectOption(layout)
  await expect
    .poll(async () =>
      page
        .locator('iframe[title="print-preview"]')
        .evaluate((frame: HTMLIFrameElement) => frame.srcdoc.includes('MASTER_HEADER_NATIVE')),
    )
    .toBe(true)
  const preview = await page
    .locator('iframe[title="print-preview"]')
    .evaluate((frame: HTMLIFrameElement) => frame.srcdoc)
  await launched.app.evaluate(({ dialog }, target) => {
    ;(globalThis as any).__pdfSaveDialogProbe = { calls: 0, target }
    dialog.showSaveDialog = (async () => {
      ;(globalThis as any).__pdfSaveDialogProbe.calls++
      return { canceled: false, filePath: target }
    }) as never
  }, path)
  await page.evaluate(() => {
    ;(window as any).__printProbe = { clicks: 0, statuses: [] }
    document.addEventListener(
      'click',
      (event) => {
        if ((event.target as HTMLElement).closest('[data-testid="print-save-pdf"]')) {
          const probe = (window as any).__printProbe
          probe.clicks = (probe.clicks ?? 0) + 1
        }
      },
      true,
    )
    new MutationObserver(() => {
      const value = document.querySelector('.status-msg')?.textContent?.trim()
      const statuses = (window as any).__printProbe.statuses as string[]
      if (value && statuses[statuses.length - 1] !== value) statuses.push(value)
    }).observe(document.body, { childList: true, subtree: true, characterData: true })
  })
  await page.getByTestId('print-save-pdf').click()
  try {
    await expect
      .poll(
        async () => {
          try {
            return (await readFile(path)).subarray(0, 5).toString()
          } catch {
            return ''
          }
        },
        { timeout: 30_000 },
      )
      .toBe('%PDF-')
  } catch (error) {
    console.error(
      'PDF request diagnostic:',
      await page.evaluate(() => ({
        clicks: (window as any).__printProbe.clicks,
        statuses: ((window as any).__printProbe.statuses as string[]).map((s) => s.slice(0, 500)),
      })),
    )
    console.error(
      'PDF save-dialog diagnostic:',
      await launched.app.evaluate(() => (globalThis as any).__pdfSaveDialogProbe),
    )
    const windows = await launched.app.evaluate(async ({ BrowserWindow }) => {
      return Promise.all(
        BrowserWindow.getAllWindows().map(async (win) => ({
          id: win.id,
          visible: win.isVisible(),
          loading: win.webContents.isLoading(),
          urlPrefix: win.webContents.getURL().slice(0, 90),
          content: await win.webContents
            .executeJavaScript(
              '({ ready: document.readyState, fonts: document.fonts.status, images: Array.from(document.images).map(i => ({ complete: i.complete, width: i.naturalWidth, source: i.src.slice(0,60) })) })',
            )
            .catch(String),
        })),
      )
    })
    console.error('PDF windows diagnostic:', JSON.stringify(windows))
    console.error(
      'PDF UI diagnostic:',
      (await page.locator('body').innerText())
        .split('\n')
        .filter((line) => /export|pdf|error|failed/i.test(line))
        .join('\n')
        .slice(0, 500),
    )
    console.error(
      'PDF dialog buttons:',
      await page.locator('.print-dialog button').evaluateAll((buttons) =>
        buttons.map((b) => ({
          text: b.textContent,
          disabled: (b as HTMLButtonElement).disabled,
        })),
      ),
    )
    throw error
  }
  expect(await pdfText(path)).toContain('MASTER_HEADER_NATIVE')
  await expect(page.locator('.print-dialog')).not.toBeVisible()
  return preview
}

interface PdfTextItem {
  str: string
  x: number
  y: number
  width: number
}

async function pdfTextItems(path: string): Promise<PdfTextItem[]> {
  const pdfjs = await import(req.resolve('pdfjs-dist/legacy/build/pdf.mjs'))
  const document = await pdfjs.getDocument({
    data: new Uint8Array(await readFile(path)),
    disableFontFace: true,
    verbosity: 0,
  }).promise
  try {
    const content = await (await document.getPage(1)).getTextContent()
    return content.items.map((item: { str?: string; transform?: number[]; width?: number }) => ({
      str: item.str ?? '',
      x: item.transform?.[4] ?? 0,
      y: item.transform?.[5] ?? 0,
      width: item.width ?? 0,
    }))
  } finally {
    await document.loadingTask.destroy()
  }
}

async function pdfText(path: string): Promise<string> {
  return (await pdfTextItems(path)).map((item) => item.str).join('')
}

function expectTabColumnsApart(items: PdfTextItem[], labels: string[]): void {
  const columns = labels.map((label) => {
    const matches = items.filter((item) => item.str.includes(label))
    expect(matches, `${label} must remain a separate PDF text item`).toHaveLength(1)
    return matches[0]!
  })
  for (let i = 1; i < columns.length; i++) {
    expect(Math.abs(columns[i]!.y - columns[0]!.y)).toBeLessThan(2)
    expect(columns[i - 1]!.x + columns[i - 1]!.width).toBeLessThanOrEqual(columns[i]!.x + 2)
  }
}

async function selectMasterNode(
  page: Page,
  x: number,
  y: number,
  w: number,
  h: number,
): Promise<void> {
  const stage = page.locator('.master-stage')
  await stage.evaluate(
    (element, target) => {
      element.scrollLeft = Math.max(0, target.x + target.w / 2 - element.clientWidth / 2)
      element.scrollTop = Math.max(0, target.y + target.h / 2 - element.clientHeight / 2)
    },
    { x, y, w, h },
  )
  const canvas = page.locator('.master-stage .stage-rel')
  const bounds = await canvas.boundingBox()
  const viewport = await stage.boundingBox()
  expect(bounds).not.toBeNull()
  expect(viewport).not.toBeNull()
  const clickX = bounds!.x + x + w / 2
  const clickY = bounds!.y + y + h / 2
  expect(clickX).toBeGreaterThan(viewport!.x)
  expect(clickX).toBeLessThan(viewport!.x + viewport!.width)
  expect(clickY).toBeGreaterThan(viewport!.y)
  expect(clickY).toBeLessThan(viewport!.y + viewport!.height)
  await page.mouse.click(clickX, clickY)
  await expect(page.getByTestId('master-font-size')).toBeEnabled()
}

async function notesMasterShape(page: Page, placeholder: string) {
  return page.evaluate(async (kind) => {
    const result = await window.slidesApi.masterEnter(1280, 'notesMaster')
    const item = result?.items.find((candidate) => candidate.kind === 'notesMaster')
    const node = item?.slide.nodes.find(
      (candidate) =>
        (candidate.type === 'shape' || candidate.type === 'text') && candidate.placeholder === kind,
    )
    return node
      ? { sourceId: node.sourceId, box: node.box, style: item?.elementStyles?.[node.sourceId] }
      : null
  }, placeholder)
}

test('Notes and Handout masters preserve native fields, geometry, style and history through save/reopen', async () => {
  test.setTimeout(420_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ?? join(tmpdir(), 'genoffice-slides-master'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'master-format-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'master-formatted.pptx')
  const notesPdf = join(dir, 'master-notes.pdf')
  const handoutPdf = join(dir, 'master-handout.pdf')
  await copyFile(fixture, source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-master-format',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await page.evaluate(() =>
      window.slidesApi.applyNotesHandoutsHeaderFooter({
        settings: {
          headerEnabled: true,
          header: 'MASTER_HEADER_NATIVE',
          footerEnabled: true,
          footer: 'MASTER_FOOTER_NATIVE',
          dateEnabled: true,
          date: '10/13/2026',
          dateAuto: true,
          dateFormat: 'datetime10',
          pageNum: true,
        },
      }),
    )
    expect(
      await page.evaluate(() =>
        window.slidesApi.setNotes({ slideIndex: 0, text: 'MASTER BODY STYLE' }),
      ),
    ).toBe(true)
    await saveAs(launched, page, saved)
    const originalDateField = /<a:fld id="([^"]+)" type="datetime10"/.exec(
      await zipXml(saved, 'ppt/notesMasters/notesMaster1.xml'),
    )?.[1]
    expect(originalDateField).toBeTruthy()

    await openMaster(page, 'notes')
    await expect(page.getByTestId('master-flag-dt')).toBeChecked()
    await expect(page.getByTestId('master-placeholder-sldImg')).toBeChecked()
    await expect(page.getByTestId('master-placeholder-body')).toBeChecked()
    const date = (await notesMasterShape(page, 'dt'))!
    await selectMasterNode(page, date.box.x, date.box.y, date.box.w, date.box.h)
    await expect(page.getByTestId('master-font-size')).toBeEnabled()
    await page.getByTestId('master-font-size').fill('19')
    await page.getByTestId('master-font-size').press('Tab')
    await expect.poll(async () => (await notesMasterShape(page, 'dt'))?.style?.fontSizePt).toBe(19)
    await page.getByTestId('master-paragraph-align').selectOption('right')
    await expect.poll(async () => (await notesMasterShape(page, 'dt'))?.style?.align).toBe('right')
    const transformed = await page.evaluate(
      ({ id, box }) =>
        window.slidesApi.masterEditTransform({
          sourceId: id,
          xPx: box.x + 45,
          yPx: box.y - 45,
          wPx: box.w,
          hPx: box.h,
          rotationDeg: 30,
          fitWidthPx: 1280,
        }),
      { id: date.sourceId, box: date.box },
    )
    expect(transformed).not.toBeNull()
    await expect.poll(async () => (await notesMasterShape(page, 'dt'))?.box.rotationDeg).toBe(30)
    await page.getByTestId('master-placeholder-body').uncheck()
    await expect(page.getByTestId('master-placeholder-body')).not.toBeChecked()
    await page.getByTestId('master-undo').click()
    await expect(page.getByTestId('master-placeholder-body')).toBeChecked()
    await page.getByTestId('master-redo').click()
    await expect(page.getByTestId('master-placeholder-body')).not.toBeChecked()
    await page.getByTestId('master-placeholder-body').check()
    const bodyNode = (await notesMasterShape(page, 'body'))!
    await selectMasterNode(page, bodyNode.box.x, bodyNode.box.y, bodyNode.box.w, bodyNode.box.h)
    await page.getByTestId('master-font-size').fill('18')
    await page.getByTestId('master-font-size').press('Tab')
    await expect
      .poll(async () => (await notesMasterShape(page, 'body'))?.style?.fontSizePt)
      .toBe(18)
    await page.getByTestId('master-paragraph-align').selectOption('right')
    await expect
      .poll(async () => (await notesMasterShape(page, 'body'))?.style?.align)
      .toBe('right')
    await page.getByTestId('master-page-orientation').selectOption('landscape')
    await expect(page.getByTestId('master-page-orientation')).toHaveValue('landscape')
    await saveAs(launched, page, saved)
    const notesXml = await zipXml(saved, 'ppt/notesMasters/notesMaster1.xml')
    const presXml = await zipXml(saved, 'ppt/presentation.xml')
    const dateField = /<a:fld id="([^"]+)" type="datetime10"/.exec(notesXml)?.[1]
    expect(dateField).toBe(originalDateField)
    expect(notesXml).toContain('sz="1900"')
    expect(notesXml).toContain('algn="r"')
    expect(notesXml).toContain('rot="1800000"')
    expect(notesXml).toContain('<p:notesStyle>')
    expect(notesXml).toContain('<p:ph type="sldImg" idx="2"')
    expect(notesXml).toContain('<p:ph type="body" sz="quarter" idx="3"')
    expect(presXml).toMatch(/<p:notesSz\b[^>]*cx="9144000"[^>]*cy="6858000"/)
    await page.getByTestId('master-undo').click()
    await expect(page.getByTestId('master-page-orientation')).toHaveValue('portrait')
    await page.getByTestId('master-redo').click()
    await expect(page.getByTestId('master-page-orientation')).toHaveValue('landscape')
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    expect(await zipXml(saved, 'ppt/notesMasters/notesMaster1.xml')).toContain(
      `<a:fld id="${dateField}" type="datetime10"`,
    )
    await page.getByRole('button', { name: 'Close Master View' }).click()

    await openMaster(page, 'handout')
    await expect(page.getByTestId('handout-master-preview-layout')).toHaveValue('6')
    await page.getByTestId('handout-master-preview-layout').selectOption('3')
    await expect(page.locator('.master-handout-preview .master-slide-slot')).toHaveCount(3)
    await page.getByTestId('master-page-orientation').selectOption('portrait')
    await expect(page.getByTestId('master-page-orientation')).toHaveValue('portrait')
    await page.getByTestId('master-undo').click()
    await expect(page.getByTestId('master-page-orientation')).toHaveValue('landscape')
    await page.getByTestId('master-redo').click()
    await expect(page.getByTestId('master-page-orientation')).toHaveValue('portrait')
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    const handoutXml = await zipXml(saved, 'ppt/handoutMasters/handoutMaster1.xml')
    expect(handoutXml).toContain('type="datetime10"')
    await page.getByRole('button', { name: 'Close Master View' }).click()
    const notesPreview = await printPdf(launched, page, 'notes', notesPdf)
    expect(notesPreview).toContain('hf-field hf-date')
    expect(notesPreview).toContain('font-size:19pt')
    const handoutPreview = await printPdf(launched, page, 'handout2', handoutPdf)
    expect(handoutPreview).toContain('hf-field hf-pageNum')
    await closeAndSaveVideo(launched, 'slides-master-format-before-reopen')
    await seedNativeMasterArtwork(saved)

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-master-format-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await openMaster(page, 'notes')
    const fresh = await page.evaluate(() => window.slidesApi.masterEnter(1280, 'notesMaster'))
    const freshNotes = fresh!.items.find((item) => item.kind === 'notesMaster')!
    const freshDate = freshNotes.slide.nodes.find(
      (node) => (node.type === 'shape' || node.type === 'text') && node.placeholder === 'dt',
    )!
    expect(freshNotes.elementStyles?.[freshDate.sourceId]).toMatchObject({
      fontSizePt: 19,
      align: 'right',
    })
    await expect(page.getByTestId('master-placeholder-body')).toBeChecked()
    await selectMasterNode(page, freshDate.box.x, freshDate.box.y, freshDate.box.w, freshDate.box.h)
    await page.getByTestId('master-font-size').fill('21')
    await page.getByTestId('master-font-size').press('Tab')
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    await page.getByRole('button', { name: 'Close Master View' }).click()
    expect(await zipXml(saved, 'ppt/notesMasters/notesMaster1.xml')).toContain(
      `<a:fld id="${dateField}" type="datetime10"`,
    )
    const notesArtworkPreview = await printPdf(
      launched,
      page,
      'notes',
      join(dir, 'master-notes-artwork.pdf'),
    )
    expect(notesArtworkPreview).toContain('class="master-artwork"')
    expect(notesArtworkPreview).toContain('rotate(30deg)')
    expect(notesArtworkPreview).toContain('MASTER BODY STYLE')
    expect(notesArtworkPreview).toContain('TAB_LEFT')
    expect(notesArtworkPreview).toContain('RTL_RIGHT')
    expect(notesArtworkPreview).toContain('PICTURE_BULLET_NATIVE')
    const artworkPdf = join(dir, 'master-notes-artwork.pdf')
    expect(await pdfText(artworkPdf)).toContain('TAB_LEFT')
    const positionedText = await pdfTextItems(artworkPdf)
    expectTabColumnsApart(positionedText, [
      'TAB_ORIGIN',
      'TAB_LEFT',
      'TAB_RIGHT',
      'TAB_CENTER',
      '12.34',
    ])
    expectTabColumnsApart(positionedText, ['RTL_RIGHT', 'RTL_LEFT', 'RTL_ORIGIN'])
    const handoutArtworkPreview = await printPdf(
      launched,
      page,
      'handout2',
      join(dir, 'master-handout-artwork.pdf'),
    )
    expect(handoutArtworkPreview).toContain('class="master-artwork"')
    await openMaster(page, 'notes')
    await page.getByTestId('master-placeholder-sldImg').uncheck()
    await page.getByTestId('master-placeholder-body').uncheck()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    expect(await zipXml(saved, 'ppt/notesMasters/notesMaster1.xml')).not.toContain(
      '<p:ph type="sldImg"',
    )
    expect(await zipXml(saved, 'ppt/notesMasters/notesMaster1.xml')).not.toContain(
      '<p:ph type="body"',
    )
    expect(await zipXml(saved, 'ppt/notesSlides/notesSlide1.xml')).toContain('MASTER BODY STYLE')
    await page.getByRole('button', { name: 'Close Master View' }).click()
    const localSurvives = await printPdf(
      launched,
      page,
      'notes',
      join(dir, 'master-off-local-survives.pdf'),
    )
    expect(localSurvives).toContain('class="notes-slide-image"')
    expect(localSurvives).toContain('MASTER BODY STYLE')
    expect(await pdfText(join(dir, 'master-off-local-survives.pdf'))).toContain('MASTER BODY STYLE')
    await closeAndSaveVideo(launched, 'slides-master-format-off-before-reopen')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-master-format-off-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await openMaster(page, 'notes')
    await expect(page.getByTestId('master-placeholder-sldImg')).not.toBeChecked()
    await expect(page.getByTestId('master-placeholder-body')).not.toBeChecked()
    await page.getByRole('button', { name: 'Close Master View' }).click()
    const reopenedLocal = await printPdf(
      launched,
      page,
      'notes',
      join(dir, 'master-off-reopened.pdf'),
    )
    expect(reopenedLocal).toContain('class="notes-slide-image"')
    expect(reopenedLocal).toContain('MASTER BODY STYLE')
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-master-format-final')
  }
})
