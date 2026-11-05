import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { existsSync } from 'node:fs'
import { mkdir, mkdtemp, readFile, readdir } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(!process.env.GENOFFICE_SHEETS_UNSAVED_VERIFY, 'Opt-in unsaved workbook lifecycle')

type Launched = Awaited<ReturnType<typeof launchShell>>

async function openedBlank(page: Page): Promise<void> {
  try {
    await page.waitForFunction(
      () => {
        const workbook = (window as any).__genofficeDebug?.univerAPI?.getActiveWorkbook()
        return typeof workbook?.getId() === 'string' && workbook.getId().startsWith('file-')
      },
      null,
      { timeout: 20_000 },
    )
  } catch (error) {
    const state = await page.evaluate(() => ({
      debug: Boolean((window as any).__genofficeDebug),
      id: (window as any).__genofficeDebug?.univerAPI?.getActiveWorkbook()?.getId() ?? null,
      text: document.body.innerText.slice(0, 500),
    }))
    throw new Error(`Blank workbook did not open: ${JSON.stringify(state)}`, { cause: error })
  }
}

async function recentPaths(page: Page): Promise<string[]> {
  return page.evaluate(async () =>
    ((await (window as any).desktopApi.getRecentWorkbooks()) as Array<{ path: string }>).map(
      (entry) => entry.path,
    ),
  )
}

async function saveAction(launched: Launched, action: 'save' | 'save-as'): Promise<void> {
  await launched.app.evaluate(({ webContents }, value) => {
    webContents
      .getAllWebContents()
      .find((item) => item.getURL().includes('://sheets/'))
      ?.send('menu:action', value)
  }, action)
}

async function captureSaveDialog(launched: Launched, target: string): Promise<void> {
  await launched.app.evaluate(({ dialog }, chosen) => {
    ;(globalThis as any).__sheetsBlankDialogDefaults = [] as string[]
    dialog.showSaveDialog = (async (...args: unknown[]) => {
      const options = args.at(-1) as { defaultPath?: string }
      ;(globalThis as any).__sheetsBlankDialogDefaults.push(options.defaultPath ?? '')
      return { canceled: false, filePath: chosen }
    }) as never
  }, target)
}

async function capturedDefaults(launched: Launched): Promise<string[]> {
  return launched.app.evaluate(() => (globalThis as any).__sheetsBlankDialogDefaults ?? [])
}

async function nativeA1(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const strings = (await zip.file('xl/sharedStrings.xml')?.async('string')) ?? ''
  return sheet + strings
}

test('folder-origin blank remains unsaved until Save and only then enters recents', async () => {
  test.setTimeout(120_000)
  const folder = await mkdtemp(join(tmpdir(), 'genoffice-unsaved-sheet-folder-'))
  const personal = join(folder, 'Personal')
  await mkdir(personal)
  const target = join(personal, 'Untitled Spreadsheet.xlsx')
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    settings: { defaultSaveDir: folder },
    videoDir: 'sheets-unsaved-folder',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    await launched.page.locator('.folder-panel .tree .tree-name', { hasText: 'Personal' }).click()
    await launched.page.locator('.quick-card', { hasText: 'AI Sheets' }).click()
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await openedBlank(page)
    expect(await readdir(personal)).toEqual([])
    expect((await readdir(folder)).filter((name) => name.endsWith('.xlsx'))).toEqual([])
    expect(await recentPaths(page)).not.toContain(target)

    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('A1').setValue('UNSAVED_DRAFT')
    })
    await expect
      .poll(() =>
        page.evaluate(() =>
          (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange('A1')
            .getValue(),
        ),
      )
      .toBe('UNSAVED_DRAFT')
    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() =>
          (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange('A1')
            .getValue(),
        ),
      )
      .toBeNull()
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() =>
          (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange('A1')
            .getValue(),
        ),
      )
      .toBe('UNSAVED_DRAFT')
    expect(existsSync(target)).toBe(false)
    expect(await recentPaths(page)).not.toContain(target)

    await captureSaveDialog(launched, target)
    await saveAction(launched, 'save')
    await expect.poll(() => existsSync(target), { timeout: 30_000 }).toBe(true)
    await expect.poll(() => nativeA1(target)).toContain('UNSAVED_DRAFT')
    expect(await capturedDefaults(launched)).toEqual([target])
    await expect.poll(() => recentPaths(page)).toContain(target)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-unsaved-folder')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    openFile: target,
    videoDir: 'sheets-unsaved-folder-reopen',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await expect
      .poll(() =>
        page.evaluate(() =>
          (window as any).__genofficeDebug?.univerAPI
            ?.getActiveWorkbook()
            ?.getActiveSheet()
            ?.getRange('A1')
            .getValue(),
        ),
      )
      .toBe('UNSAVED_DRAFT')
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-unsaved-folder-reopen')
  }
})

test('closing an untouched blank leaves no file; first Save As may create a clean workbook', async () => {
  test.setTimeout(120_000)
  const folder = await mkdtemp(join(tmpdir(), 'genoffice-unsaved-sheet-clean-'))
  const target = join(folder, 'Clean.xlsx')
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    settings: { defaultSaveDir: folder },
    videoDir: 'sheets-unsaved-clean',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI Sheets' }).click()
    let page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await openedBlank(page)
    expect((await readdir(folder)).filter((name) => name.endsWith('.xlsx'))).toEqual([])
    await launched.page.locator('.tab-bar .tab-item.active .tab-close').click()
    await expect(launched.page.locator('.quick-card', { hasText: 'AI Sheets' })).toBeVisible()
    expect((await readdir(folder)).filter((name) => name.endsWith('.xlsx'))).toEqual([])

    await launched.page.locator('.quick-card', { hasText: 'AI Sheets' }).click()
    page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await openedBlank(page)
    await captureSaveDialog(launched, target)
    await saveAction(launched, 'save-as')
    await expect.poll(() => existsSync(target), { timeout: 30_000 }).toBe(true)
    const defaults = await capturedDefaults(launched)
    expect(defaults).toEqual([join(folder, 'Untitled Spreadsheet.xlsx')])
    await expect.poll(() => recentPaths(page)).toContain(target)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-unsaved-clean')
  }
})
