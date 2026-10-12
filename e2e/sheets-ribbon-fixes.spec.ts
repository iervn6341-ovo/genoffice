import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Regressions found by the ribbon audit of Sheets: the ribbon Paste button did nothing, multi-column
 * AutoSum and Cell Styles took several ⌘Z to undo, and two popups scrolled / clipped sideways.
 */
const TAB = 'nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)'

async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-ribbon-fixes',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.page.locator('.quick-card').nth(1).click()
  const p = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await p.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  await p.waitForTimeout(1500)
  // wide enough that no ribbon group folds: the buttons are all directly clickable
  await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')
const api = '(window).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()'
const activate = (p: Page, a1: string) => p.evaluate(`${api}.getRange('${a1}').activate()`)
const values = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getValues()`) as Promise<unknown[][]>
const tab = (p: Page, name: string) =>
  p
    .locator(TAB, { hasText: new RegExp(`^${name}$`) })
    .first()
    .click()
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(400)
}
const pick = async (p: Page, trigger: string, option: string) => {
  await p.locator(`[data-ribbon-body] [aria-label="${trigger}"]`).first().click()
  await p.locator('.menu-select-drop button', { hasText: option }).first().click()
  await p.waitForTimeout(500)
}

const SEED = [
  ['Item', 'Qty', 'Price'],
  ['Apple', 3, 1.5],
  ['Pear', 5, 2.25],
]

test('the ribbon Copy / Cut / Paste buttons move cells', async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1:C3').setValues(${JSON.stringify(SEED)})`)
    // Copy → Paste keeps the source
    await activate(p, 'A1:C3')
    await p.locator('[data-tip^="Copy"]').first().click()
    await p.waitForTimeout(300)
    await activate(p, 'F8')
    await p.locator('[data-tip^="Paste"]').first().click()
    await expect.poll(() => values(p, 'F8:H10')).toEqual(SEED)
    expect(await values(p, 'A1:C3')).toEqual(SEED)
    // Cut → Paste moves them
    await activate(p, 'A1:C3')
    await p.locator('[data-tip^="Cut"]').first().click()
    await p.waitForTimeout(300)
    await activate(p, 'K1')
    await p.locator('[data-tip^="Paste"]').first().click()
    await expect.poll(() => values(p, 'K1:M3')).toEqual(SEED)
    expect(await values(p, 'A1:C3')).toEqual([
      [null, null, null],
      [null, null, null],
      [null, null, null],
    ])
  } finally {
    await kill(launched)
  }
})

test('AutoSum over several columns and a Cell Style each undo in one step', async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1:C3').setValues(${JSON.stringify(SEED)})`)
    await activate(p, 'A1:C3')
    await pick(p, 'AutoSum', 'Sum')
    await expect.poll(() => values(p, 'B4:C4')).toEqual([[8, 3.75]])
    await undo(p)
    expect(await values(p, 'A4:C4')).toEqual([[null, null, null]])

    await activate(p, 'A1:C1')
    const before = await p.evaluate(`${api}.getRange('A1:C1').getBackgrounds()`)
    await pick(p, 'Cell Styles', 'Input')
    expect(await p.evaluate(`${api}.getRange('A1:C1').getBackgrounds()`)).not.toEqual(before)
    await undo(p)
    expect(await p.evaluate(`${api}.getRange('A1:C1').getBackgrounds()`)).toEqual(before)
  } finally {
    await kill(launched)
  }
})

test('the Shapes gallery and the equation presets fit their boxes', async () => {
  const { launched, p } = await openSheets()
  try {
    await tab(p, 'Insert')
    await p.locator('[data-ribbon-body] [aria-label="Shapes"]').first().click()
    const gallery = p.locator('.menu-select-drop.rb-shape-gallery')
    await expect(gallery).toBeVisible()
    const m = await gallery.evaluate((el) => ({ sw: el.scrollWidth, cw: el.clientWidth }))
    expect(m.sw).toBeLessThanOrEqual(m.cw + 1)
    await p.keyboard.press('Escape')

    await p
      .locator('[data-ribbon-body] [data-tip^="Equation"], [data-ribbon-body] button', {
        hasText: /Equation/,
      })
      .first()
      .click()
    const presets = p.locator('.equation-presets button')
    await expect(presets).toHaveCount(6)
    const clipped = await presets.evaluateAll(
      (els) => els.filter((b) => b.scrollWidth > b.clientWidth + 1).length,
    )
    expect(clipped).toBe(0)
  } finally {
    await kill(launched)
  }
})

test('Conditional Formatting is a menu, Borders has Excel presets, font sizes and number formats match', async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1:C3').setValues(${JSON.stringify(SEED)})`)
    await activate(p, 'A1:C3')
    // Conditional Formatting ▾: Excel's categories, each opening the rule panel preset to that kind
    await p.locator('[data-ribbon-body] [aria-label="Conditional Formatting"]').first().click()
    const cf = p.locator('.menu-select-drop button')
    await expect(cf).toHaveCount(10)
    for (const name of [
      'Highlight Cells Rules',
      'Top/Bottom Rules',
      'Data Bars',
      'Color Scales',
      'Icon Sets',
      'New Rule',
      'Clear Rules from Selected Cells',
      'Clear Rules from Entire Sheet',
      'Manage Rules',
    ])
      await expect(cf.filter({ hasText: name }).first()).toBeVisible()
    await cf.filter({ hasText: 'Data Bars' }).first().click()
    await expect(
      p.locator('[class*="univer-"]', { hasText: /Data Bar|data bar/i }).first(),
    ).toBeVisible({ timeout: 10_000 })
    await p.keyboard.press('Escape')

    // Borders ▾: compound presets set several edges in one undo step
    const cell = (a1: string) =>
      p.evaluate(
        `(() => { const r = ${api}.getRange('${a1}'); const c = r.getCellData?.(); return JSON.stringify(c?.s ?? null) })()`,
      )
    await activate(p, 'A2:C2')
    await p.locator('[data-ribbon-body] [aria-label="Borders"]').first().click()
    const items = await p.locator('.menu-select-drop button').allInnerTexts()
    for (const name of [
      'Bottom Double Border',
      'Thick Bottom Border',
      'Top and Bottom Border',
      'Top and Thick Bottom Border',
      'Top and Double Bottom Border',
      'Inside Borders',
      'Inside Horizontal Border',
      'Inside Vertical Border',
    ])
      expect(items.map((t) => t.trim())).toContain(name)
    const before = await cell('B2')
    await p
      .locator('.menu-select-drop button', { hasText: 'Top and Double Bottom Border' })
      .first()
      .click()
    await p.waitForTimeout(400)
    expect(await cell('B2')).not.toEqual(before)
    await undo(p)
    expect(await cell('B2')).toEqual(before)

    // Font size: Excel's preset list
    await p
      .locator('[data-ribbon-body] button.menu-select-caret[aria-label="Font size"]')
      .first()
      .click()
    const sizes = (await p.locator('.menu-select-drop button').allInnerTexts()).map((t) => t.trim())
    for (const size of ['8', '20', '24', '28', '36', '48', '72']) expect(sizes).toContain(size)
    await p.keyboard.press('Escape')

    // Number format: More Number Formats… opens Format Cells
    await p.locator('[data-ribbon-body] [aria-label="Number format"]').first().click()
    await p.locator('.menu-select-drop button', { hasText: 'More Number Formats' }).first().click()
    await expect(p.locator('.dialog-backdrop')).toBeVisible()
  } finally {
    await kill(launched)
  }
})

test('Format as Table draws the picked style, restyles in place and ⌘Z walks it back', async () => {
  const { launched, p } = await openSheets()
  try {
    const rows = [
      ['Item', 'Qty', 'Price', 'Date'],
      ['Apple', 3, 1.5, 'x'],
      ['Pear', 5, 2.25, 'y'],
      ['Plum', 8, 0.75, 'z'],
    ]
    await p.evaluate(`${api}.getRange('A1:D4').setValues(${JSON.stringify(rows)})`)
    await activate(p, 'A1:D4')
    const tables = () =>
      p.evaluate(`${api}.getSubTableInfos().map((t) => t.name)`) as Promise<string[]>
    // the contextual Table Design tab lingers 700 ms after its table goes: let the ribbon settle before comparing pixels
    const look = async () => {
      await p.waitForTimeout(900)
      return p.screenshot({ clip: { x: 0, y: 60, width: 900, height: 130 } })
    }
    const formatAsTable = async (style: string) => {
      await p.locator('[data-ribbon-body] [aria-label="Format as Table"]').first().click()
      await p.locator(`.table-style-cell[data-style="${style}"]`).click()
      await p.waitForTimeout(1500)
    }

    await formatAsTable('TableStyleLight1')
    expect(await tables()).toEqual(['Table1'])
    const light = await look()
    // a second style over the same range restyles that table instead of failing on the overlap
    await formatAsTable('TableStyleMedium3')
    expect(await tables()).toEqual(['Table1'])
    const orange = await look()
    expect(orange.equals(light), 'the Medium 3 look differs from Light 1').toBe(false)
    await formatAsTable('TableStyleMedium7')
    expect(await tables()).toEqual(['Table1'])
    expect((await look()).equals(orange), 'the Medium 7 look differs from Medium 3').toBe(false)

    // one ⌘Z per style: back to orange, back to gray, then no table at all
    await undo(p)
    expect(await tables()).toEqual(['Table1'])
    expect((await look()).equals(orange)).toBe(true)
    await undo(p)
    expect((await look()).equals(light)).toBe(true)
    await undo(p)
    expect(await tables()).toEqual([])
    // the undone table is gone from the journal too: the same range takes a new table
    await formatAsTable('TableStyleDark2')
    expect(await tables()).toEqual(['Table1'])
  } finally {
    await kill(launched)
  }
})

test('Page Layout edits are on the undo stack, not the cell edit before them', async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1').setValue('keep me')`)
    await activate(p, 'A1')
    const sheetId = (await p.evaluate(`${api}.getSheetId()`)) as string
    const setup = () =>
      p.evaluate(`window.__genofficeDebug.pageSetup(${JSON.stringify(sheetId)})`) as Promise<{
        orientation?: string
        fitToWidth?: number
      } | null>
    await tab(p, 'Page Layout')
    await pick(p, 'Orientation', 'Landscape')
    await expect.poll(async () => (await setup())?.orientation).toBe('landscape')
    await pick(p, 'Fit to width', '1 page')
    await expect.poll(async () => (await setup())?.fitToWidth).toBe(1)

    // ⌘Z takes back the page setup, one edit at a time, and leaves the cell alone
    await undo(p)
    await expect.poll(async () => (await setup())?.fitToWidth ?? 0).toBe(0)
    // the ribbon shows the restored value, not the one that was undone
    await expect(p.locator('[data-ribbon-body] [aria-label="Fit to width"]').first()).toContainText(
      /Automatic/i,
    )
    expect((await setup())?.orientation).toBe('landscape')
    expect(await values(p, 'A1')).toEqual([['keep me']])
    await undo(p)
    await expect.poll(async () => (await setup())?.orientation).toBeUndefined()
    expect(await values(p, 'A1')).toEqual([['keep me']])
    // redo puts them back
    await p.locator('button[aria-label^="Redo"]').first().click()
    await expect.poll(async () => (await setup())?.orientation).toBe('landscape')
    // only now does Undo reach the cell edit
    await undo(p)
    await undo(p)
    expect(await values(p, 'A1')).toEqual([[null]])
  } finally {
    await kill(launched)
  }
})

test('a document theme is one undo step too', async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1').setValue('keep me')`)
    await activate(p, 'A1')
    const themeName = () =>
      p.evaluate(`window.__genofficeDebug.theme().colors?.name ?? null`) as Promise<string | null>
    await tab(p, 'Page Layout')
    await p
      .locator('[data-ribbon-body] .ribbon-tool:not(.rb-fold-btn)', { hasText: 'Themes' })
      .click()
    await p.getByRole('option', { name: 'Indigo' }).click()
    await expect.poll(themeName).toBe('Indigo')
    await undo(p)
    await expect.poll(themeName).toBeNull()
    expect(await values(p, 'A1')).toEqual([['keep me']])
    await p.locator('button[aria-label^="Redo"]').first().click()
    await expect.poll(themeName).toBe('Indigo')
  } finally {
    await kill(launched)
  }
})
