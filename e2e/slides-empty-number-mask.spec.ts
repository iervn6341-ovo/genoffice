import { expect, test, type Page } from '@playwright/test'
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(!process.env.GENOFFICE_EMPTY_NUMBER_REPRO_PATH, 'Opt-in imported placeholder audit')
const root = resolve('docs/qa/powerpoint-empty-number-2026-11-04')
const file = join(root, 'engine-fixed-empty-number.pptx')

async function values(page: Page) {
  return page.evaluate(async () =>
    (await (window as any).slidesApi.getRenderSlides()).map((s: any) =>
      s.nodes
        .filter((n: any) => n.placeholder === 'sldNum')
        .map(
          (n: any) =>
            n.text?.lines?.map((p: any) => p.runs.map((r: any) => r.text).join('')).join('') ?? '',
        )
        .join(''),
    ),
  )
}
async function numberDialog(page: Page) {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^Insert$/ }).click()
  await page.locator('.ribbon-body button[data-tip^="Header & footer"]').click()
  return page.locator('.modal', { has: page.getByRole('heading', { name: 'Header and Footer' }) })
}
async function placeholder(path: string, slide: number) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const xml = await zip.file(`ppt/slides/slide${slide}.xml`)!.async('string')
  return [...xml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)]
    .map((m) => m[0])
    .filter((s) => /<p:ph\b[^>]*type="sldNum"/.test(s))
}

test('migrates imported empty number masks on Save; restores, undoes and reopens styled fields', async () => {
  test.setTimeout(180_000)
  await mkdir(root, { recursive: true })
  await copyFile(resolve(process.env.GENOFFICE_EMPTY_NUMBER_REPRO_PATH!), file)
  const original = (await placeholder(file, 2))[0]!
  let l = await launchShell({
    onboardingSeen: true,
    openFile: file,
    videoDir: 'empty-number-mask',
    settings: { autoSave: false },
  })
  try {
    let page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(25)
    expect((await values(page)).slice(0, 4)).toEqual(['', '', '1', '2'])
    expect((await page.evaluate(() => (window as any).slidesApi.save())).ok).toBe(true)
    const hidden = await placeholder(file, 2)
    expect(hidden).toHaveLength(1)
    expect(hidden[0]).toMatch(/<p:cNvPr\b[^>]*hidden="1"/)
    // Preserve the engine-generated file for independent PowerPoint opening.
    await copyFile(file, join(root, 'powerpoint-check-engine-fixed.pptx'))
    await closeAndSaveVideo(l, 'empty-number-mask-saved')
    l = await launchShell({
      onboardingSeen: true,
      openFile: file,
      videoDir: 'empty-number-mask-reopen',
      settings: { autoSave: false },
    })
    page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(25)
    expect((await values(page)).slice(0, 4)).toEqual(['', '', '1', '2'])
    await page.locator('.slide-list .thumb').nth(1).click()
    await page.locator('.stage-rel').screenshot({ path: join(root, 'genoffice-hidden-prefix.png') })
    const d = await numberDialog(page)
    await d.getByTestId('hf-slide-number').check()
    await d.getByTestId('hf-number-from-slide').fill('1')
    await d.getByTestId('hf-apply-all').click()
    expect((await values(page)).slice(0, 4)).toEqual(['1', '2', '3', '4'])
    await page.evaluate(() => (window as any).slidesApi.undo())
    expect((await values(page)).slice(0, 4)).toEqual(['', '', '1', '2'])
    await page.evaluate(() => (window as any).slidesApi.redo())
    expect((await values(page)).slice(0, 4)).toEqual(['1', '2', '3', '4'])
    expect((await page.evaluate(() => (window as any).slidesApi.save())).ok).toBe(true)
    const restored = await placeholder(file, 2)
    expect(restored).toHaveLength(1)
    expect(restored[0]).not.toContain('hidden="1"')
    expect(restored[0]).toContain('type="slidenum"')
    for (const pattern of [
      /<p:cNvPr\b[^>]*\bid="(\d+)"/,
      /<a:xfrm\b[\s\S]*?<\/a:xfrm>/,
      /<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/,
    ])
      expect(restored[0]!.match(pattern)?.[0]).toBe(original.match(pattern)?.[0])
    await closeAndSaveVideo(l, 'empty-number-mask-restored')
    l = await launchShell({
      onboardingSeen: true,
      openFile: file,
      videoDir: 'empty-number-mask-restored-reopen',
      settings: { autoSave: false },
    })
    page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(25)
    expect((await values(page)).slice(0, 4)).toEqual(['1', '2', '3', '4'])
    await writeFile(
      join(root, 'engine-result.json'),
      JSON.stringify(
        {
          physicalSlides: 25,
          offset: { startSlide: 3, startNumber: 1 },
          prefix: ['', ''],
          numbered: ['1', '2'],
          saveMigration: true,
          restored: ['1', '2', '3', '4'],
          undoRedo: true,
          saveReopen: true,
          preservedShapeIdGeometryAndRunStyle: true,
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(l, 'empty-number-mask-end')
  }
})

test('reopens the PowerPoint-saved file and can edit its offset numbering', async () => {
  test.skip(!process.env.GENOFFICE_EMPTY_NUMBER_OFFICE_PATH, 'Requires native PowerPoint Save As')
  test.setTimeout(120_000)
  const officeCopy = join(root, 'office-roundtrip-reedit.pptx')
  await copyFile(resolve(process.env.GENOFFICE_EMPTY_NUMBER_OFFICE_PATH!), officeCopy)
  let l = await launchShell({
    onboardingSeen: true,
    openFile: officeCopy,
    videoDir: 'empty-number-office-roundtrip',
    settings: { autoSave: false },
  })
  try {
    let page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(25)
    expect((await values(page)).slice(0, 4)).toEqual(['', '', '1', '2'])
    await page.locator('.slide-list .thumb').nth(2).click()
    const d = await numberDialog(page)
    await d.getByTestId('hf-slide-start-at').fill('7')
    await d.getByTestId('hf-apply-all').click()
    expect((await values(page)).slice(0, 4)).toEqual(['', '', '7', '8'])
    await page.evaluate(() => (window as any).slidesApi.undo())
    expect((await values(page)).slice(0, 4)).toEqual(['', '', '1', '2'])
    await page.evaluate(() => (window as any).slidesApi.redo())
    expect((await page.evaluate(() => (window as any).slidesApi.save())).ok).toBe(true)
    await closeAndSaveVideo(l, 'empty-number-office-reedit-saved')
    l = await launchShell({
      onboardingSeen: true,
      openFile: officeCopy,
      videoDir: 'empty-number-office-reedit-reopen',
      settings: { autoSave: false },
    })
    page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(25)
    expect((await values(page)).slice(0, 4)).toEqual(['', '', '7', '8'])
    const active = (await placeholder(officeCopy, 3))[0]!
    expect(active).toContain('type="genofficeSlideNumber"')
    expect((await placeholder(officeCopy, 2))[0]).toContain('hidden="1"')
    await writeFile(
      join(root, 'office-roundtrip-result.json'),
      JSON.stringify(
        {
          officeSaveReopen: true,
          fieldsEditable: true,
          undoRedo: true,
          savedReopenedNumbers: ['', '', '7', '8'],
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(l, 'empty-number-office-end')
  }
})
