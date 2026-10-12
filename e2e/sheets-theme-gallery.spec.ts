import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
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
import { COLOR_SCHEMES, FONT_SCHEMES, THEME_PRESETS } from '../apps/sheets/src/renderer/themes'

/**
 * Page Layout → Themes / Colors / Fonts are galleries of cards with a live preview: while the
 * pointer is on a card the workbook shows that theme; leaving, Esc or closing puts the workbook's
 * own theme back without touching the journal or the undo stack; a click applies it as one undo step.
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-theme-gallery',
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
type Debug = {
  themeColors(): string[] | null
  themeFonts(): { major: string; minor: string } | null
  theme(): { colors?: { name: string }; fonts?: { name: string } }
}
const colors = (p: Page) =>
  p.evaluate(() =>
    (window as unknown as { __genofficeDebug: Debug }).__genofficeDebug.themeColors(),
  )
const fonts = (p: Page) =>
  p.evaluate(() => (window as unknown as { __genofficeDebug: Debug }).__genofficeDebug.themeFonts())
const journal = (p: Page) =>
  p.evaluate(() => (window as unknown as { __genofficeDebug: Debug }).__genofficeDebug.theme())
const tab = (p: Page, name: string) =>
  p
    .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', {
      hasText: new RegExp(`^${name}$`),
    })
    .first()
    .click()
const undoDisabled = (p: Page) => p.locator('button[aria-label^="Undo"]').first().isDisabled()
const open = async (p: Page, label: string) => {
  await p.locator(`[data-ribbon-body] .cover-select[aria-label="${label}"]`).first().click()
  const gallery = p.locator('.theme-gallery')
  await expect(gallery).toBeVisible()
  return gallery
}

test('Themes is a gallery of cards that previews on hover and restores on leave, Esc and close', async () => {
  const { launched, p } = await openSheets()
  try {
    await tab(p, 'Page Layout')
    const before = await colors(p)
    expect(before, 'a new workbook carries a theme').not.toBeNull()
    const indigo = COLOR_SCHEMES.find((s) => s.id === 'indigo')!.values
    const forest = COLOR_SCHEMES.find((s) => s.id === 'forest')!.values
    expect(before).not.toEqual(indigo)

    const gallery = await open(p, 'Themes')
    await expect(gallery.locator('.theme-card')).toHaveCount(COLOR_SCHEMES.length)
    // a card shows its font sample, its six accents and its name
    const indigoCard = gallery.locator('.theme-card[data-theme-id="indigo"]')
    await expect(indigoCard.locator('.theme-card-swatches i')).toHaveCount(6)
    await expect(indigoCard.locator('.theme-card-name')).toHaveText('Indigo')

    // hover previews: the workbook's theme is Indigo's, yet nothing is journaled or undoable
    await indigoCard.hover()
    await expect.poll(() => colors(p)).toEqual(indigo)
    expect((await journal(p)).colors).toBeUndefined()
    expect(await undoDisabled(p)).toBe(true)
    // moving to another card previews that one
    await gallery.locator('.theme-card[data-theme-id="forest"]').hover()
    await expect.poll(() => colors(p)).toEqual(forest)
    // leaving the gallery puts the workbook's own theme back
    await p.mouse.move(700, 500)
    await expect.poll(() => colors(p)).toEqual(before)
    expect((await journal(p)).colors).toBeUndefined()

    // Esc closes and restores too
    await gallery.locator('.theme-card[data-theme-id="forest"]').hover()
    await expect.poll(() => colors(p)).toEqual(forest)
    await p.keyboard.press('Escape')
    await expect(p.locator('.theme-gallery')).toHaveCount(0)
    await expect.poll(() => colors(p)).toEqual(before)

    // a click applies it for real: journaled, one undo step, and no leftover preview
    await open(p, 'Themes')
    await p.locator('.theme-card[data-theme-id="indigo"]').hover()
    await expect.poll(() => colors(p)).toEqual(indigo)
    await p.locator('.theme-card[data-theme-id="indigo"]').click()
    await expect(p.locator('.theme-gallery')).toHaveCount(0)
    await expect.poll(async () => (await journal(p)).colors?.name).toBe('Indigo')
    expect(await colors(p)).toEqual(indigo)
    await p.mouse.move(700, 500)
    await p.waitForTimeout(500)
    expect(await colors(p), 'leaving after a pick does not restore the old theme').toEqual(indigo)
    // ⌘Z puts the workbook's own theme back (the preview left no stale "before")
    await p.locator('button[aria-label^="Undo"]').first().click()
    await expect.poll(() => colors(p)).toEqual(before)
    expect((await journal(p)).colors).toBeUndefined()
  } finally {
    await kill(launched)
  }
})

test('Colors and Fonts are galleries with the same live preview', async () => {
  const { launched, p } = await openSheets()
  try {
    await tab(p, 'Page Layout')
    const beforeColors = await colors(p)
    const beforeFonts = await fonts(p)
    expect(beforeFonts).not.toBeNull()
    const rose = COLOR_SCHEMES.find((s) => s.id === 'rose')!.values
    const georgia = FONT_SCHEMES.find((s) => s.id === 'georgia')!

    let gallery = await open(p, 'Colors')
    await expect(gallery.locator('.theme-card')).toHaveCount(COLOR_SCHEMES.length)
    await gallery.locator('.theme-card[data-theme-id="rose"]').hover()
    await expect.poll(() => colors(p)).toEqual(rose)
    // colors alone: the fonts stay
    expect(await fonts(p)).toEqual(beforeFonts)
    await p.keyboard.press('Escape')
    await expect.poll(() => colors(p)).toEqual(beforeColors)

    gallery = await open(p, 'Fonts')
    await expect(gallery.locator('.theme-card')).toHaveCount(FONT_SCHEMES.length)
    await gallery.locator('.theme-card[data-theme-id="georgia"]').hover()
    await expect.poll(() => fonts(p)).toEqual({ major: georgia.major, minor: georgia.minor })
    expect(await colors(p)).toEqual(beforeColors)
    await p.keyboard.press('Escape')
    await expect.poll(() => fonts(p)).toEqual(beforeFonts)
    expect(await undoDisabled(p)).toBe(true)

    // clicking a font scheme applies it
    gallery = await open(p, 'Fonts')
    await gallery.locator('.theme-card[data-theme-id="georgia"]').click()
    await expect.poll(async () => (await journal(p)).fonts?.name).toBe('Georgia')
    expect(await fonts(p)).toEqual({ major: georgia.major, minor: georgia.minor })
  } finally {
    await kill(launched)
  }
})

test("Excel's built-in themes are in the gallery, apply with their own colors and fonts, and save as that theme", async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-theme-'))
  const file = join(dir, 'facet.xlsx')
  try {
    await tab(p, 'Page Layout')
    const gallery = await open(p, 'Themes')
    // Excel 16's own themes sit next to the app's presets (24 documents themes, one of them the
    // classic Office palette the app already had)
    for (const id of ['facet', 'ion-boardroom', 'retrospect', 'wisp', 'office-theme', 'indigo']) {
      await expect(gallery.locator(`.theme-card[data-theme-id="${id}"]`)).toHaveCount(1)
    }
    await expect(gallery.locator('.theme-card[data-theme-id="facet"] .theme-card-name')).toHaveText(
      'Facet',
    )
    const facet = THEME_PRESETS.find((preset) => preset.id === 'facet')!
    await gallery.locator('.theme-card[data-theme-id="facet"]').hover()
    await expect.poll(() => colors(p)).toEqual(facet.colors.values)
    await expect.poll(() => fonts(p)).toEqual({ major: 'Trebuchet MS', minor: 'Trebuchet MS' })
    await gallery.locator('.theme-card[data-theme-id="facet"]').click()
    await expect.poll(async () => (await journal(p)).colors?.name).toBe('Facet')

    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
      dialog.showMessageBox = (async () => ({ response: 0, checkboxChecked: false })) as never
    }, file)
    await p.locator('button[aria-label^="Save As"]').first().click()
    await expect.poll(() => existsSync(file), { timeout: 20_000 }).toBe(true)
    await new Promise((resolve) => setTimeout(resolve, 800))
    const zip = await JSZip.loadAsync(await readFile(file))
    const theme = await zip.file('xl/theme/theme1.xml')!.async('string')
    expect(theme).toContain('<a:clrScheme name="Facet">')
    expect(theme).toMatch(/<a:accent1><a:srgbClr val="90C226"\/><\/a:accent1>/)
    expect(theme).toMatch(/<a:majorFont><a:latin typeface="Trebuchet MS"/)
    const keepDir = process.env.QA_KEEP_DIR
    if (keepDir) await copyFile(file, join(keepDir, 'facet.xlsx'))
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})
