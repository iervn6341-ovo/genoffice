import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Home tab gaps against Excel, driven through the real ribbon menus and dialogs: Fill Up / Left /
 * Series…, Clear Comments and Notes, and the Cells group's Insert ▾ / Delete ▾ menus. The expected
 * values were read from Excel 16 through COM (see docs/qa/handoff-ribbon-audit-2026-10-11.md).
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-home-gaps',
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
/// raw numbers: getValues() returns the display text (and ####### in a narrow column) for formatted cells
const raw = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getRawValues()`) as Promise<unknown[][]>
const formulas = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getFormulas()`) as Promise<string[][]>
const style = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getCellStyleData()`) as Promise<{ bl?: number } | null>
const setValues = (p: Page, a1: string, data: unknown[][]) =>
  p.evaluate(`${api}.getRange('${a1}').setValues(${JSON.stringify(data)})`)
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(400)
}
const redo = async (p: Page) => {
  await p.locator('button[aria-label^="Redo"]').first().click()
  await p.waitForTimeout(400)
}
const pick = async (p: Page, trigger: string, option: string | RegExp) => {
  await p.locator(`[data-ribbon-body] [aria-label="${trigger}"]`).first().click()
  await p.locator('.menu-select-drop button', { hasText: option }).first().click()
  await p.waitForTimeout(500)
}

test('Fill ▾ has Up, Left and Series like Excel, and Up / Left copy the way Excel does', async () => {
  const { launched, p } = await openSheets()
  try {
    // the menu lists Excel's entries in order
    await p.locator('[data-ribbon-body] [aria-label="Fill"]').first().click()
    await expect(p.locator('.menu-select-drop button')).toHaveText([
      'Fill Down',
      'Fill Right',
      'Fill Up',
      'Fill Left',
      'Series…',
      'Justify',
      'Across Worksheets…',
    ])
    await p.keyboard.press('Escape')

    // several rows: the bottom row goes up (Excel: A1:A4 = 1,2,3,9 → 9,9,9,9)
    await setValues(p, 'A1:A4', [[1], [2], [3], [9]])
    await activate(p, 'A1:A4')
    await pick(p, 'Fill', 'Fill Up')
    await expect.poll(() => values(p, 'A1:A4')).toEqual([[9], [9], [9], [9]])
    await undo(p)
    expect(await values(p, 'A1:A4')).toEqual([[1], [2], [3], [9]])
    await redo(p)
    expect(await values(p, 'A1:A4')).toEqual([[9], [9], [9], [9]])

    // one cell: the cell below is copied into it
    await setValues(p, 'C3', [[5]])
    await activate(p, 'C2')
    await pick(p, 'Fill', 'Fill Up')
    await expect.poll(() => values(p, 'C2:C3')).toEqual([[5], [5]])

    // formulas move their relative references (Excel: B4 =A4*2 filled up → =A1*2 … )
    await setValues(p, 'A10:A13', [[1], [2], [3], [4]])
    await p.evaluate(`${api}.getRange('B13').setFormula('=A13*2')`)
    await activate(p, 'B10:B13')
    await pick(p, 'Fill', 'Fill Up')
    await expect
      .poll(() => formulas(p, 'B10:B13'))
      .toEqual([['=A10*2'], ['=A11*2'], ['=A12*2'], ['=A13*2']])
    expect(await values(p, 'B10:B13')).toEqual([[2], [4], [6], [8]])

    // formats ride along (Excel: bold + red fill on the source reaches every cell)
    await setValues(p, 'E3', [[1]])
    await p.evaluate(`${api}.getRange('E3').setFontWeight('bold').setBackground('#FF0000')`)
    await activate(p, 'E1:E3')
    await pick(p, 'Fill', 'Fill Up')
    await expect.poll(async () => (await style(p, 'E1'))?.bl).toBe(1)
    expect(await p.evaluate(`${api}.getRange('E1').getBackgrounds()[0][0].toLowerCase()`)).toBe(
      '#ff0000',
    )

    // Fill Left: the right column goes left
    await setValues(p, 'G1:J1', [[1, 2, 3, 9]])
    await activate(p, 'G1:J1')
    await pick(p, 'Fill', 'Fill Left')
    await expect.poll(() => values(p, 'G1:J1')).toEqual([[9, 9, 9, 9]])
    await undo(p)
    expect(await values(p, 'G1:J1')).toEqual([[1, 2, 3, 9]])

    // one cell: the cell to its right is copied into it
    await setValues(p, 'M2', [[4]])
    await activate(p, 'L2')
    await pick(p, 'Fill', 'Fill Left')
    await expect.poll(() => values(p, 'L2:M2')).toEqual([[4, 4]])
  } finally {
    await kill(launched)
  }
})

async function openSeries(p: Page) {
  await pick(p, 'Fill', /Series/)
  const dialog = p.locator('.fill-series-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

test('Series…: linear, growth, stop value, trend and one-step undo match Excel', async () => {
  const { launched, p } = await openSheets()
  try {
    // linear, columns, step 2 (Excel: 1 3 5 7 9)
    await setValues(p, 'A1', [[1]])
    await activate(p, 'A1:A5')
    let dialog = await openSeries(p)
    // a column selection opens on Columns / Linear
    await expect(dialog.getByRole('radio', { name: 'Columns' })).toBeChecked()
    await expect(dialog.getByRole('radio', { name: 'Linear' })).toBeChecked()
    await dialog.getByLabel('Step value:').fill('2')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => values(p, 'A1:A5')).toEqual([[1], [3], [5], [7], [9]])
    await undo(p) // one ⌘Z takes back the whole series
    expect(await values(p, 'A1:A5')).toEqual([[1], [null], [null], [null], [null]])
    await redo(p)
    expect(await values(p, 'A1:A5')).toEqual([[1], [3], [5], [7], [9]])

    // growth with a stop value (Excel: start 2 step 2 stop 20 → 2 4 8 16, the rest untouched)
    await setValues(p, 'C1', [[2]])
    await activate(p, 'C1:C10')
    dialog = await openSeries(p)
    await dialog.getByRole('radio', { name: 'Growth' }).check()
    await dialog.getByLabel('Step value:').fill('2')
    await dialog.getByLabel('Stop value:').fill('20')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => values(p, 'C1:C6')).toEqual([[2], [4], [8], [16], [null], [null]])

    // one cell + stop value runs past the selection (Excel: 1 … 5)
    await setValues(p, 'E1', [[1]])
    await activate(p, 'E1')
    dialog = await openSeries(p)
    await dialog.getByLabel('Stop value:').fill('5')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => values(p, 'E1:E7')).toEqual([[1], [2], [3], [4], [5], [null], [null]])

    // a one-row selection opens on Rows; each row is its own series
    await setValues(p, 'G1:G2', [[1], [100]])
    await activate(p, 'G1:I2')
    dialog = await openSeries(p)
    await dialog.getByRole('radio', { name: 'Rows' }).check()
    await dialog.getByLabel('Step value:').fill('5')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(() => values(p, 'G1:I2'))
      .toEqual([
        [1, 6, 11],
        [100, 105, 110],
      ])

    // trend replaces the numbers with the least-squares line (Excel: 1,3 → 1 3 5 7 9)
    await setValues(p, 'K1:K2', [[1], [3]])
    await activate(p, 'K1:K5')
    dialog = await openSeries(p)
    await dialog.getByLabel('Trend').check()
    await expect(dialog.getByLabel('Step value:')).toBeDisabled()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => values(p, 'K1:K5')).toEqual([[1], [3], [5], [7], [9]])

    // the first cell's format goes with the series (Excel: 0.00 + bold)
    await setValues(p, 'M1', [[1]])
    await p.evaluate(`${api}.getRange('M1').setNumberFormat('0.00').setFontWeight('bold')`)
    await activate(p, 'M1:M3')
    dialog = await openSeries(p)
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => raw(p, 'M2')).toEqual([[2]])
    expect(await p.evaluate(`${api}.getRange('M2').getNumberFormat()`)).toBe('0.00')
    expect((await style(p, 'M2'))?.bl).toBe(1)

    // the format further down the line is replaced, not merged: a percent format under a General first cell goes
    await setValues(p, 'O1', [[1]])
    await p.evaluate(`${api}.getRange('O2:O5').setNumberFormat('0%')`)
    await activate(p, 'O1:O5')
    dialog = await openSeries(p)
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => raw(p, 'O5')).toEqual([[5]])
    expect(await p.evaluate(`${api}.getRange('O3').getNumberFormat()`)).not.toBe('0%')
  } finally {
    await kill(launched)
  }
})

test('Series…: a date start opens on Date; month steps clamp to the month end like Excel', async () => {
  const { launched, p } = await openSheets()
  try {
    // 2024-01-31 → Feb 29, Mar 31, Apr 30, May 31 (Excel steps from the start each time)
    await setValues(p, 'A1', [[45322]])
    await p.evaluate(`${api}.getRange('A1:A5').setNumberFormat('yyyy-mm-dd')`)
    await activate(p, 'A1:A5')
    const dialog = await openSeries(p)
    await expect(dialog.getByRole('radio', { name: 'Date', exact: true })).toBeChecked()
    await dialog.getByRole('radio', { name: 'Month' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => raw(p, 'A1:A5')).toEqual([[45322], [45351], [45382], [45412], [45443]])
    // the date format carried over
    expect(await p.evaluate(`${api}.getRange('A3').getNumberFormat()`)).toBe('yyyy-mm-dd')

    // weekdays skip the weekend (Fri 2024-02-02 → Mon, Tue, Wed …)
    await setValues(p, 'C1', [[45324]])
    await p.evaluate(`${api}.getRange('C1:C4').setNumberFormat('yyyy-mm-dd')`)
    await activate(p, 'C1:C4')
    const weekday = await openSeries(p)
    await weekday.getByRole('radio', { name: 'Weekday' }).check()
    await weekday.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => raw(p, 'C1:C4')).toEqual([[45324], [45327], [45328], [45329]])
  } finally {
    await kill(launched)
  }
})

test('Series…: AutoFill type continues a pattern; bad numbers keep the dialog open', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:A2', [[1], [3]])
    await activate(p, 'A1:A5')
    let dialog = await openSeries(p)
    await dialog.getByRole('radio', { name: 'AutoFill' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    // Excel: 1,3 + AutoFill over A1:A5 → 1 3 5 7 9
    await expect.poll(() => values(p, 'A1:A5')).toEqual([[1], [3], [5], [7], [9]])

    await setValues(p, 'C1', [[1]])
    await activate(p, 'C1:C3')
    dialog = await openSeries(p)
    await dialog.getByLabel('Step value:').fill('abc')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog.getByRole('alert')).toBeVisible()
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(dialog).toHaveCount(0)
    expect(await values(p, 'C1:C3')).toEqual([[1], [null], [null]])

    // a text start is left alone, with a message (Excel writes nothing)
    await setValues(p, 'E1', [['abc']])
    await activate(p, 'E1:E3')
    dialog = await openSeries(p)
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(p.locator('.workbook-status')).toContainText('number')
    expect(await values(p, 'E1:E3')).toEqual([['abc'], [null], [null]])
  } finally {
    await kill(launched)
  }
})

test("Clear ▾ lists Excel's entries; Clear Comments and Notes removes notes (one undo)", async () => {
  const { launched, p } = await openSheets()
  try {
    await p.locator('[data-ribbon-body] [aria-label="Clear"]').first().click()
    await expect(p.locator('.menu-select-drop button')).toHaveText([
      'Clear All',
      'Clear Formats',
      'Clear Contents',
      'Clear Comments and Notes',
      'Clear Hyperlinks',
      'Remove Hyperlinks',
    ])
    await p.keyboard.press('Escape')

    await setValues(p, 'A1:A3', [[1], [2], [3]])
    for (const [row, text] of [
      [0, 'first'],
      [2, 'third'],
    ] as const) {
      await p.evaluate(
        `${api}.getRange(${row}, 0, 1, 1).createOrUpdateNote({ id: 'n${row}', row: ${row}, col: 0, width: 200, height: 80, note: '${text}' })`,
      )
    }
    const noteCount = () => p.evaluate(`${api}.getNotes().length`)
    expect(await noteCount()).toBe(2)

    // contents go, notes stay (Excel: ClearContents keeps the note)
    await activate(p, 'A1')
    await pick(p, 'Clear', 'Clear Contents')
    await expect.poll(() => values(p, 'A1')).toEqual([[null]])
    expect(await noteCount()).toBe(2)
    await undo(p)

    // Clear Comments and Notes over A1:A2 takes only the first note, text stays
    await activate(p, 'A1:A2')
    await pick(p, 'Clear', 'Clear Comments and Notes')
    await expect.poll(noteCount).toBe(1)
    expect(await values(p, 'A1:A3')).toEqual([[1], [2], [3]])
    await undo(p)
    await expect.poll(noteCount).toBe(2)
    await redo(p)
    await expect.poll(noteCount).toBe(1)

    // Clear All takes text and notes together
    await activate(p, 'A3')
    await pick(p, 'Clear', 'Clear All')
    await expect.poll(noteCount).toBe(0)
    expect(await values(p, 'A3')).toEqual([[null]])
  } finally {
    await kill(launched)
  }
})

test("Cells group: Insert ▾ / Delete ▾ carry Excel's entries and act on the whole selection", async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:C4', [
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
      [10, 11, 12],
    ])
    await p.locator('[data-ribbon-body] [aria-label="Insert cells"]').first().click()
    await expect(p.locator('.menu-select-drop button')).toHaveText([
      'Insert Cells…',
      'Insert Sheet Rows',
      'Insert Sheet Columns',
      'Insert Sheet',
    ])
    await p.keyboard.press('Escape')
    await p.locator('[data-ribbon-body] [aria-label="Delete cells"]').first().click()
    await expect(p.locator('.menu-select-drop button')).toHaveText([
      'Delete Cells…',
      'Delete Sheet Rows',
      'Delete Sheet Columns',
      'Delete Sheet',
    ])
    await p.keyboard.press('Escape')

    // Insert Sheet Rows inserts as many rows as are selected (Excel: two rows → two rows)
    await activate(p, 'A2:C3')
    await pick(p, 'Insert cells', 'Insert Sheet Rows')
    await expect
      .poll(() => values(p, 'A1:C6'))
      .toEqual([
        [1, 2, 3],
        [null, null, null],
        [null, null, null],
        [4, 5, 6],
        [7, 8, 9],
        [10, 11, 12],
      ])
    await undo(p)
    expect(await values(p, 'A1:C4')).toEqual([
      [1, 2, 3],
      [4, 5, 6],
      [7, 8, 9],
      [10, 11, 12],
    ])

    // Delete Sheet Columns takes the selected columns
    await activate(p, 'B1:B2')
    await pick(p, 'Delete cells', 'Delete Sheet Columns')
    await expect
      .poll(() => values(p, 'A1:C2'))
      .toEqual([
        [1, 3, null],
        [4, 6, null],
      ])
    await undo(p)

    // Insert Cells… opens Excel's dialog, preselected by the selection's shape
    await activate(p, 'A1:B1')
    await pick(p, 'Insert cells', 'Insert Cells…')
    const dialog = p.locator('.insert-cells-dialog')
    await expect(dialog).toBeVisible()
    await expect(dialog.getByRole('radio', { name: 'Shift cells down' })).toBeChecked()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(() => values(p, 'A1:C2'))
      .toEqual([
        [null, null, 3],
        [1, 2, 6],
      ])
    await undo(p)

    // Delete Sheet is refused for the only sheet; with two it asks, then removes the active one
    await pick(p, 'Delete cells', /^Delete Sheet$/)
    await expect(p.locator('.workbook-status')).toContainText('at least one visible sheet')
    await pick(p, 'Insert cells', /^Insert Sheet$/)
    await expect
      .poll(() =>
        p.evaluate(`(window).__genofficeDebug.univerAPI.getActiveWorkbook().getSheets().length`),
      )
      .toBe(2)
    await pick(p, 'Delete cells', /^Delete Sheet$/)
    await p
      .getByRole('button', { name: /confirm|ok|delete/i })
      .last()
      .click()
    await expect
      .poll(() =>
        p.evaluate(`(window).__genofficeDebug.univerAPI.getActiveWorkbook().getSheets().length`),
      )
      .toBe(1)
  } finally {
    await kill(launched)
  }
})
