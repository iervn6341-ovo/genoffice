import { expect, test, type Page } from '@playwright/test'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_SLIDES_CUSTOM_SHOW_VERIFICATION,
  'Opt-in native custom-show verification',
)

const fixture = join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')

async function seedImportedShow(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const xml = await zip.file('ppt/presentation.xml')!.async('string')
  const rids = [...xml.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)].map((match) => match[1]!)
  expect(rids).toHaveLength(5)
  const native = `<p:custShowLst><p:custShow name="Imported Board" id="7"><p:sldLst><p:sld r:id="${rids[3]}"/><p:sld r:id="${rids[1]}"/></p:sldLst></p:custShow></p:custShowLst>`
  const insertBefore =
    /<p:(?:photoAlbum|custDataLst|kinsoku|defaultTextStyle|modifyVerifier|showPr|extLst)\b|<\/p:presentation>/.exec(
      xml,
    )
  expect(insertBefore).toBeTruthy()
  zip.file(
    'ppt/presentation.xml',
    xml.slice(0, insertBefore!.index) + native + xml.slice(insertBefore!.index),
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function nativeShows(path: string): Promise<{ id: string; name: string; rids: string[] }[]> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const xml = await zip.file('ppt/presentation.xml')!.async('string')
  return [...xml.matchAll(/<p:custShow\b([^>]+)>([\s\S]*?)<\/p:custShow>/g)].map((match) => ({
    id: /\bid="([^"]+)"/.exec(match[1]!)?.[1] ?? '',
    name: /\bname="([^"]+)"/.exec(match[1]!)?.[1] ?? '',
    rids: [...match[2]!.matchAll(/<p:sld\b[^>]*r:id="([^"]+)"/g)].map((ref) => ref[1]!),
  }))
}

async function openDialog(page: Page): Promise<void> {
  await page
    .locator('.ribbon-tabs .ribbon-tab')
    .filter({ hasText: /^Slide Show$/ })
    .click()
  await page.getByTestId('ribbon-custom-show').click()
  await expect(page.locator('.modal.csd')).toBeVisible()
}

test('imported custom shows edit, undo, save and reopen as native ordered references', async () => {
  test.setTimeout(210_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ?? join(tmpdir(), 'genoffice-slides-shows'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'custom-shows-'))
  const source = join(dir, 'source.pptx')
  const saved = join(dir, 'custom-shows.pptx')
  await copyFile(fixture, source)
  await seedImportedShow(source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-custom-shows',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getCustomShows()))
      .toEqual([{ id: '7', name: 'Imported Board', slideIndices: [3, 1] }])

    await openDialog(page)
    await expect(page.getByTestId('custom-show-7')).toContainText('Imported Board')
    await page.getByTestId('custom-show-name').fill('Board Review')
    await page.getByTestId('custom-show-new').click()
    await page.getByTestId('custom-show-name').fill('Short Show')
    await page.getByTestId('custom-show-slide-0').check()
    await page.getByTestId('custom-show-slide-4').check()
    await page.getByTestId('custom-show-up-1').click()
    await page.getByTestId('custom-show-close').click()
    await expect(page.locator('.modal.csd')).not.toBeVisible()
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getCustomShows()))
      .toEqual([
        { id: '7', name: 'Board Review', slideIndices: [3, 1] },
        { id: '0', name: 'Short Show', slideIndices: [4, 0] },
      ])

    await page.getByTestId('ribbon-undo').click()
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getCustomShows()))
      .toEqual([{ id: '7', name: 'Imported Board', slideIndices: [3, 1] }])
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(() => page.evaluate(() => window.slidesApi.getCustomShows())).toHaveLength(2)
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, saved)
    expect((await page.evaluate(() => window.slidesApi.saveAs('custom-shows.pptx'))).ok).toBe(true)
    const raw = await nativeShows(saved)
    expect(raw.map(({ id, name }) => ({ id, name }))).toEqual([
      { id: '7', name: 'Board Review' },
      { id: '0', name: 'Short Show' },
    ])
    expect(raw[0]!.rids).toHaveLength(2)
    expect(raw[1]!.rids).toHaveLength(2)

    await closeAndSaveVideo(launched, 'slides-custom-shows-saved')
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-custom-shows-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await openDialog(page)
    await page.getByTestId('custom-show-7').click()
    await page.getByTestId('custom-show-name').fill('Board Final')
    await page.getByTestId('custom-show-close').click()
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getCustomShows()))
      .toMatchObject([
        { id: '7', name: 'Board Final', slideIndices: [3, 1] },
        { id: '0', name: 'Short Show', slideIndices: [4, 0] },
      ])
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    expect((await nativeShows(saved)).map((show) => show.name)).toEqual([
      'Board Final',
      'Short Show',
    ])
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-custom-shows-final')
  }
})

test('legacy path-keyed shows migrate only into an empty native list and retire after Save', async () => {
  test.setTimeout(210_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ?? join(tmpdir(), 'genoffice-slides-shows'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'legacy-shows-'))
  const source = join(dir, 'legacy-source.pptx')
  await copyFile(fixture, source)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-custom-shows-legacy-seed',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    const legacyKey = `ai-slides-custom-shows:${source}`
    await page.evaluate((key) => {
      localStorage.setItem(
        key,
        JSON.stringify([{ id: 'cs-old-id', name: 'Legacy Board', slideIndices: [4, 1, 4] }]),
      )
    }, legacyKey)
    const userDataDir = launched.userDataDir
    await closeAndSaveVideo(launched, 'slides-custom-shows-legacy-seeded')
    launched = await launchShell({
      userDataDir,
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-custom-shows-legacy-migrate',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await expect
      .poll(() => page.evaluate(() => window.slidesApi.getCustomShows()))
      .toEqual([{ id: '0', name: 'Legacy Board', slideIndices: [4, 1, 4] }])
    expect(await page.evaluate((key) => localStorage.getItem(key), legacyKey)).not.toBeNull()
    expect(await nativeShows(source)).toEqual([])
    await expect(page.getByTestId('ribbon-save')).toBeEnabled()
    await page.getByTestId('ribbon-save').click()
    await expect.poll(() => page.evaluate((key) => localStorage.getItem(key), legacyKey)).toBeNull()
    await expect.poll(async () => (await nativeShows(source)).length).toBe(1)
    expect(await nativeShows(source)).toMatchObject([
      { id: '0', name: 'Legacy Board', rids: expect.any(Array) },
    ])
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-custom-shows-legacy-final')
  }
})

test('deleting every migrated show survives Undo, Redo, Save and a fresh process', async () => {
  test.setTimeout(210_000)
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ?? join(tmpdir(), 'genoffice-slides-shows'),
  )
  await mkdir(root, { recursive: true })
  const dir = await mkdtemp(join(root, 'legacy-clear-'))
  const source = join(dir, 'legacy-clear.pptx')
  await copyFile(fixture, source)
  const key = `ai-slides-custom-shows:${source}`
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-custom-shows-clear-seed',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await page.evaluate((legacyKey) => {
      localStorage.setItem(
        legacyKey,
        JSON.stringify([{ id: 'cs-clear', name: 'Remove Me', slideIndices: [2, 0] }]),
      )
    }, key)
    const userDataDir = launched.userDataDir
    await closeAndSaveVideo(launched, 'slides-custom-shows-clear-seeded')

    launched = await launchShell({
      userDataDir,
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-custom-shows-clear-edit',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await expect.poll(() => page.evaluate(() => window.slidesApi.getCustomShows())).toHaveLength(1)
    await openDialog(page)
    await page.getByTestId('custom-show-delete').click()
    await page.getByTestId('custom-show-close').click()
    await expect.poll(() => page.evaluate(() => window.slidesApi.getCustomShows())).toEqual([])
    await page.getByTestId('ribbon-undo').click()
    await expect.poll(() => page.evaluate(() => window.slidesApi.getCustomShows())).toHaveLength(1)
    await page.getByTestId('ribbon-redo').click()
    await expect.poll(() => page.evaluate(() => window.slidesApi.getCustomShows())).toEqual([])
    await expect(page.getByTestId('ribbon-save')).toBeEnabled()
    await page.getByTestId('ribbon-save').click()
    await expect
      .poll(() => page.evaluate((legacyKey) => localStorage.getItem(legacyKey), key))
      .toBeNull()
    expect(await nativeShows(source)).toEqual([])
    await closeAndSaveVideo(launched, 'slides-custom-shows-clear-saved')

    launched = await launchShell({
      userDataDir,
      onboardingSeen: true,
      lang: 'en',
      openFile: source,
      videoDir: 'slides-custom-shows-clear-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb')
    await expect.poll(() => page.evaluate(() => window.slidesApi.getCustomShows())).toEqual([])
    expect(await page.evaluate((legacyKey) => localStorage.getItem(legacyKey), key)).toBeNull()
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-custom-shows-clear-final')
  }
})
