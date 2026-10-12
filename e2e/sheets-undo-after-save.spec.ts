import { existsSync, statSync } from 'node:fs'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * ⌘Z across a save, for the edits that only change what the save writes (page setup, the document
 * theme): Excel keeps its undo history through a save, and so does this app (the stack is carried over
 * the reopen). An undo after the save has to write the older value into the file at the next save,
 * not only repaint — the saved file is no longer "the older value" the journal could fall back to.
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-undo-after-save',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.page.locator('.quick-card').nth(1).click()
  const p = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await p.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  await p.waitForTimeout(1500)
  await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')
const tab = (p: Page, name: string) =>
  p
    .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', {
      hasText: new RegExp(`^${name}$`),
    })
    .first()
    .click()
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(1200)
}
const redo = async (p: Page) => {
  await p.locator('button[aria-label^="Redo"]').first().click()
  await p.waitForTimeout(1200)
}

async function save(
  l: LaunchedApp,
  p: Page,
  path: string,
  first: boolean,
): Promise<{ sheet: string; theme: string }> {
  const before = existsSync(path) ? statSync(path).mtimeMs : -1
  if (first) {
    await l.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
      dialog.showMessageBox = (async () => ({ response: 0, checkboxChecked: false })) as never
    }, path)
    await p.locator('button[aria-label^="Save As"]').first().click()
  } else {
    await p.locator('button[aria-label^="Save ("]').first().click()
  }
  await expect
    .poll(async () => (existsSync(path) ? statSync(path).mtimeMs : -1), { timeout: 20_000 })
    .not.toBe(before)
  await new Promise((resolve) => setTimeout(resolve, 1200))
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    sheet: (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).replace(/\s+/g, ' '),
    theme: (await zip.file('xl/theme/theme1.xml')?.async('string')) ?? '',
  }
}
const orientation = (sheet: string): string =>
  /<pageSetup\b[^>]*\borientation="(\w+)"/.exec(sheet)?.[1] ?? 'portrait'
const schemeName = (theme: string): string => /<a:clrScheme name="([^"]*)"/.exec(theme)?.[1] ?? ''

test('⌘Z after a save puts the older page orientation into the file at the next save', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-undo-save-'))
  const file = join(dir, 'orientation.xlsx')
  try {
    await tab(p, 'Page Layout')
    await p.locator('[data-ribbon-body] [aria-label="Orientation"]').first().click()
    await p.locator('.menu-select-drop button', { hasText: 'Landscape' }).first().click()
    await p.waitForTimeout(600)
    expect(orientation((await save(launched, p, file, true)).sheet)).toBe('landscape')

    // the edit is behind us; Undo still takes it back, and the next save says so
    await undo(p)
    expect(orientation((await save(launched, p, file, false)).sheet)).toBe('portrait')

    // and ⇧⌘Z brings it back across another save
    await redo(p)
    expect(orientation((await save(launched, p, file, false)).sheet)).toBe('landscape')
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('⌘Z after a save puts the older theme colors into the file at the next save', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-undo-save-theme-'))
  const file = join(dir, 'theme.xlsx')
  try {
    const first = await (async () => {
      await tab(p, 'Page Layout')
      return null
    })()
    void first
    await p.locator('[data-ribbon-body] .cover-select[aria-label="Colors"]').first().click()
    await p.locator('.theme-card[data-theme-id="indigo"]').click()
    await p.waitForTimeout(800)
    const saved = await save(launched, p, file, true)
    expect(schemeName(saved.theme)).toBe('Indigo')

    await undo(p)
    expect(schemeName((await save(launched, p, file, false)).theme)).not.toBe('Indigo')

    await redo(p)
    expect(schemeName((await save(launched, p, file, false)).theme)).toBe('Indigo')
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})
