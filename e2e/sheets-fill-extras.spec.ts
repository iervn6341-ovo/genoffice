import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Home → Fill → Justify and Across Worksheets, through the real menus and dialogs. Justify's rules
 * were read from Excel 16 through COM (Range.Justify): the text of the column is joined with spaces and
 * broken again to the width of the selection; numbers or formulas in the range refuse it.
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-fill-extras',
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
const wb = '(window).__genofficeDebug.univerAPI.getActiveWorkbook()'
const api = `${wb}.getActiveSheet()`
const activate = (p: Page, a1: string) => p.evaluate(`${api}.getRange('${a1}').activate()`)
const raw = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getRawValues()`) as Promise<unknown[][]>
const setValues = (p: Page, a1: string, data: unknown[][]) =>
  p.evaluate(`${api}.getRange('${a1}').setValues(${JSON.stringify(data)})`)
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(600)
}
const redo = async (p: Page) => {
  await p.locator('button[aria-label^="Redo"]').first().click()
  await p.waitForTimeout(600)
}
const pick = async (p: Page, trigger: string, option: string | RegExp) => {
  await p.locator(`[data-ribbon-body] [aria-label="${trigger}"]`).first().click()
  await p.locator('.menu-select-drop button', { hasText: option }).first().click()
  await p.waitForTimeout(700)
}
const column = async (p: Page, rows: number): Promise<string[]> =>
  ((await raw(p, `A1:A${rows}`)) as unknown[][]).map((row) => String(row[0] ?? ''))

const TEXT = 'The quick brown fox jumps over the lazy dog and keeps running far away from here'

test('Justify breaks the text of a column into lines the width of the selection, in one undo step', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:A3', [[TEXT], ['second cell'], ['third']])
    await activate(p, 'A1:A3')
    await pick(p, 'Fill', 'Justify')
    // the words are all there, in order, and no line is longer than the others could fit
    await expect.poll(async () => (await column(p, 2))[0]).not.toBe(TEXT)
    const lines = (await column(p, 20)).filter((line) => line !== '')
    expect(lines.length).toBeGreaterThan(3)
    expect(lines.join(' ')).toBe(`${TEXT} second cell third`)
    for (const line of lines) expect(line.length).toBeLessThan(26)
    await expect(p.locator('.workbook-status')).toContainText('justified')

    // ⌘Z puts the three cells back, and ⇧⌘Z breaks the lines again
    await undo(p)
    expect(await column(p, 4)).toEqual([TEXT, 'second cell', 'third', ''])
    await redo(p)
    expect((await column(p, 20)).filter((line) => line !== '').join(' ')).toBe(
      `${TEXT} second cell third`,
    )
  } finally {
    await kill(launched)
  }
})

test('Justify over a wider selection uses its width, and refuses numbers and formulas (Excel)', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:A2', [[TEXT], ['more']])
    await activate(p, 'A1:A2')
    await pick(p, 'Fill', 'Justify')
    const narrow = (await column(p, 40)).filter((line) => line !== '')
    await undo(p)
    await activate(p, 'A1:F2')
    await pick(p, 'Fill', 'Justify')
    const wide = (await column(p, 40)).filter((line) => line !== '')
    expect(wide.length).toBeLessThan(narrow.length)
    expect(wide.join(' ')).toBe(`${TEXT} more`)
    await undo(p)

    // a number in the range: refused, nothing changes
    await p.evaluate(`${api}.getRange('A2').setValue(42)`)
    await activate(p, 'A1:A3')
    await pick(p, 'Fill', 'Justify')
    await expect(p.locator('.workbook-status')).toContainText('numbers or formulas')
    expect(await raw(p, 'A1:A2')).toEqual([[TEXT], [42]])
    // nothing but empty cells: nothing to justify
    await activate(p, 'H1:H3')
    await pick(p, 'Fill', 'Justify')
    await expect(p.locator('.workbook-status')).toContainText('Select cells with text')
  } finally {
    await kill(launched)
  }
})

test('Justify will not run over the cells under the selection when they are not empty', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1', [[TEXT]])
    await setValues(p, 'A2', [['in the way']])
    await activate(p, 'A1')
    await pick(p, 'Fill', 'Justify')
    await expect(p.locator('.workbook-status')).toContainText('not empty')
    expect(await raw(p, 'A1:A2')).toEqual([[TEXT], ['in the way']])
    // with room below, the lines run under the selection (Excel asks first)
    await p.evaluate(`${api}.getRange('A2').setValue('')`)
    await pick(p, 'Fill', 'Justify')
    await expect
      .poll(async () => (await column(p, 8)).filter((line) => line !== '').length)
      .toBeGreaterThan(2)
  } finally {
    await kill(launched)
  }
})

test('Across Worksheets copies the selection to the sheets that are ticked: All, Contents or Formats', async () => {
  const { launched, p } = await openSheets()
  try {
    // a second sheet
    await pick(p, 'Insert cells', /^Insert Sheet$/)
    await expect.poll(() => p.evaluate(`${wb}.getSheets().length`)).toBe(2)
    await p.evaluate(`${wb}.getSheets()[0].activate()`)
    await setValues(p, 'A1:B2', [
      ['x', 1],
      ['y', 2],
    ])
    await p.evaluate(`${api}.getRange('C1').setFormula('=A1&B1')`)
    await p.evaluate(`${api}.getRange('A1:C1').setFontWeight('bold')`)
    await p.evaluate(`${api}.getRange('A1:C2').activate()`)

    const other = (a1: string) =>
      p.evaluate(`${wb}.getSheets()[1].getRange('${a1}').getRawValues()`) as Promise<unknown[][]>
    const otherBold = (a1: string) =>
      p.evaluate(
        `${wb}.getSheets()[1].getRange('${a1}').getCellStyleData()?.bl ?? 0`,
      ) as Promise<number>
    const otherFormula = (a1: string) =>
      p.evaluate(`${wb}.getSheets()[1].getRange('${a1}').getFormulas()`) as Promise<string[][]>
    const openAcross = async () => {
      await p.locator('[data-ribbon-body] [aria-label="Fill"]').first().click()
      await p.locator('.menu-select-drop button', { hasText: 'Across Worksheets…' }).click()
      return p.getByRole('dialog', { name: 'Fill Across Worksheets' })
    }

    // Contents: the values and formulas, not the bold
    let dialog = await openAcross()
    await expect(dialog.getByRole('button', { name: 'OK' })).toBeDisabled()
    await dialog.getByRole('checkbox', { name: 'Sheet2' }).check()
    await dialog.getByRole('radio', { name: 'Contents' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(() => other('A1:B2'))
      .toEqual([
        ['x', 1],
        ['y', 2],
      ])
    expect(await otherFormula('C1')).toEqual([['=A1&B1']])
    expect(await otherBold('A1')).toBe(0)
    await expect(p.locator('.workbook-status')).toContainText('Filled 1 sheet')
    // one ⌘Z empties the other sheet again
    await undo(p)
    // ⌘Z follows the cells to the sheet they are on; back to the first one
    await p.evaluate(`${wb}.getSheets()[0].activate()`)
    expect(await other('A1:C2')).toEqual([
      [null, null, null],
      [null, null, null],
    ])

    // Formats: the bold, no values
    await p.evaluate(`${api}.getRange('A1:C2').activate()`)
    dialog = await openAcross()
    await dialog.getByRole('checkbox', { name: 'Sheet2' }).check()
    await dialog.getByRole('radio', { name: 'Formats' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => otherBold('A1')).toBe(1)
    expect(await other('A1:B2')).toEqual([
      [null, null],
      [null, null],
    ])
    await undo(p)
    await p.evaluate(`${wb}.getSheets()[0].activate()`)

    // All: both
    await p.evaluate(`${api}.getRange('A1:C2').activate()`)
    dialog = await openAcross()
    await dialog.getByRole('checkbox', { name: 'Sheet2' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => otherBold('B1')).toBe(1)
    expect(await other('A1:B2')).toEqual([
      ['x', 1],
      ['y', 2],
    ])
  } finally {
    await kill(launched)
  }
})
