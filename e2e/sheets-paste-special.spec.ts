import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Paste ▾ / Paste Special… against Excel, through the real ribbon and dialog: Transpose, Values &
 * Number Formatting, Formulas & Number Formatting, Operation, Skip blanks. The expected results
 * were read from Excel 16 through COM (see docs/qa/handoff-ribbon-audit-2026-10-11.md).
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-paste-special',
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
const api = '(window).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()'
const activate = (p: Page, a1: string) => p.evaluate(`${api}.getRange('${a1}').activate()`)
const raw = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getRawValues()`) as Promise<unknown[][]>
const formulas = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getFormulas()`) as Promise<string[][]>
const style = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getCellStyleData()`) as Promise<{
    bl?: number
    n?: { pattern?: string }
  } | null>
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
const bold = async (p: Page) => {
  await p.locator('[data-ribbon-body] [data-tip="Bold"]').first().click()
  await p.waitForTimeout(300)
}
const copy = async (p: Page) => {
  await p.locator('[data-tip^="Copy"]').first().click()
  await p.waitForTimeout(400)
}
const pasteMenu = async (p: Page, option: string | RegExp) => {
  await p.locator('[data-ribbon-body] [aria-label="Paste Special"]').first().click()
  await p.locator('.menu-select-drop button', { hasText: option }).first().click()
  await p.waitForTimeout(500)
}
const openDialog = async (p: Page) => {
  await pasteMenu(p, /^Paste Special…$/)
  const dialog = p.locator('.paste-special-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

test("the Paste Special menu lists Excel's choices, with Transpose and the number-format variants", async () => {
  const { launched, p } = await openSheets()
  try {
    await p.locator('[data-ribbon-body] [aria-label="Paste Special"]').first().click()
    await expect(p.locator('.menu-select-drop button')).toHaveText([
      'Values Only',
      'Formulas Only',
      'Formatting Only',
      'Column Widths Only',
      'All Except Borders',
      'Values & Number Formatting',
      'Formulas & Number Formatting',
      'Transpose',
      'Paste Special…',
    ])
  } finally {
    await kill(launched)
  }
})

test('Transpose turns rows into columns, moves formulas with their cells, and undoes in one step', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:C2', [
      [1, 2, 3],
      [4, 5, 0],
    ])
    await p.evaluate(`${api}.getRange('C2').setFormula('=A2+B2')`)
    await activate(p, 'A1:C2')
    await copy(p)
    await activate(p, 'E1')
    await pasteMenu(p, 'Transpose')
    // Excel: E1:F3 = 1 4 / 2 5 / 3 9, and C2 =A2+B2 becomes F3 =F1+F2
    await expect
      .poll(() => raw(p, 'E1:F3'))
      .toEqual([
        [1, 4],
        [2, 5],
        [3, 9],
      ])
    expect((await formulas(p, 'F3:F3'))[0]![0]).toBe('=F1+F2')
    // the pasted block is selected, the source is untouched
    expect(await raw(p, 'A1:C2')).toEqual([
      [1, 2, 3],
      [4, 5, 9],
    ])

    await undo(p)
    expect(await raw(p, 'E1:F3')).toEqual([
      [null, null],
      [null, null],
      [null, null],
    ])
    await redo(p)
    await expect.poll(() => raw(p, 'E1:F1')).toEqual([[1, 4]])
  } finally {
    await kill(launched)
  }
})

test('Transpose of a block onto its own corner is refused, as in Excel', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:C1', [[1, 2, 3]])
    await activate(p, 'A1:C1')
    await copy(p)
    await activate(p, 'A1')
    await pasteMenu(p, 'Transpose')
    await expect(p.locator('.workbook-status')).toContainText('overlap')
    expect(await raw(p, 'A1:C3')).toEqual([
      [1, 2, 3],
      [null, null, null],
      [null, null, null],
    ])
  } finally {
    await kill(launched)
  }
})

test('Paste Special with nothing copied says so instead of failing silently', async () => {
  const { launched, p } = await openSheets()
  try {
    await activate(p, 'A1')
    await pasteMenu(p, 'Transpose')
    await expect(p.locator('.workbook-status')).toContainText('Copy cells')
  } finally {
    await kill(launched)
  }
})

test('Values & Number Formatting and Formulas & Number Formatting carry only the number format', async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1').setFormula('=1/3').setNumberFormat('0.00')`)
    await activate(p, 'A1')
    await bold(p)
    await copy(p)

    await activate(p, 'B1')
    await pasteMenu(p, 'Values & Number Formatting')
    // Excel: the value 0.33 as a constant, number format kept, bold not
    await expect.poll(async () => (await style(p, 'B1'))?.n?.pattern).toBe('0.00')
    expect((await raw(p, 'B1'))[0]![0]).toBeCloseTo(1 / 3, 10)
    expect((await formulas(p, 'B1'))[0]![0]).toBe('')
    expect((await style(p, 'B1'))?.bl).not.toBe(1)
    await undo(p)
    expect((await raw(p, 'B1'))[0]![0]).toBeNull()

    await activate(p, 'C1')
    await pasteMenu(p, 'Formulas & Number Formatting')
    await expect.poll(async () => (await formulas(p, 'C1'))[0]![0]).toBe('=1/3')
    expect((await style(p, 'C1'))?.n?.pattern).toBe('0.00')
    expect((await style(p, 'C1'))?.bl).not.toBe(1)
  } finally {
    await kill(launched)
  }
})

test('Paste Special… dialog: Add / Subtract / Multiply / Divide act on the cells underneath', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:A2', [[1], [2]])
    await setValues(p, 'C1:C2', [[10], [20]])
    await activate(p, 'A1:A2')
    await copy(p)

    // defaults, like Excel: All + None, nothing ticked
    await activate(p, 'C1:C2')
    let dialog = await openDialog(p)
    await expect(dialog.getByRole('radio', { name: 'All', exact: true })).toBeChecked()
    await expect(dialog.getByRole('radio', { name: 'None' })).toBeChecked()
    await expect(dialog.getByLabel('Skip blanks')).not.toBeChecked()
    await expect(dialog.getByLabel('Transpose')).not.toBeChecked()
    // Column widths takes no operation, skip blanks or transpose
    await dialog.getByRole('radio', { name: 'Column widths' }).check()
    await expect(dialog.getByLabel('Transpose')).toBeDisabled()
    await dialog.getByRole('radio', { name: 'Values', exact: true }).check()
    await expect(dialog.getByLabel('Transpose')).toBeEnabled()
    await dialog.getByRole('radio', { name: 'Add' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    // Excel: 10+1, 20+2
    await expect.poll(() => raw(p, 'C1:C2')).toEqual([[11], [22]])
    await undo(p)
    expect(await raw(p, 'C1:C2')).toEqual([[10], [20]])

    for (const [operation, expected] of [
      ['Subtract', [[9], [18]]],
      ['Multiply', [[10], [40]]],
      ['Divide', [[10], [10]]],
    ] as const) {
      await activate(p, 'C1:C2')
      dialog = await openDialog(p)
      await dialog.getByRole('radio', { name: 'Values', exact: true }).check()
      await dialog.getByRole('radio', { name: operation }).check()
      await dialog.getByRole('button', { name: 'OK' }).click()
      await expect.poll(() => raw(p, 'C1:C2')).toEqual(expected)
      await undo(p)
      expect(await raw(p, 'C1:C2')).toEqual([[10], [20]])
    }

    // a formula destination is wrapped like Excel: =2+3 plus 1 → =(2+3)+1
    await p.evaluate(`${api}.getRange('E1').setFormula('=2+3')`)
    await setValues(p, 'G1', [[1]])
    await activate(p, 'G1')
    await copy(p)
    await activate(p, 'E1')
    dialog = await openDialog(p)
    await dialog.getByRole('radio', { name: 'Values', exact: true }).check()
    await dialog.getByRole('radio', { name: 'Add' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(async () => (await formulas(p, 'E1'))[0]![0]).toBe('=(2+3)+1')
    await expect.poll(() => raw(p, 'E1')).toEqual([[6]])
  } finally {
    await kill(launched)
  }
})

test('Paste Special… dialog: Skip blanks and Transpose with formats', async () => {
  const { launched, p } = await openSheets()
  try {
    // Skip blanks: A1=1, A2 blank copied over C1=9, C2=8 → 1, 8 (Excel); without it → 1, blank
    await setValues(p, 'A1', [[1]])
    await setValues(p, 'C1:C2', [[9], [8]])
    await activate(p, 'A1:A2')
    await copy(p)
    await activate(p, 'C1:C2')
    let dialog = await openDialog(p)
    await dialog.getByLabel('Skip blanks').check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => raw(p, 'C1:C2')).toEqual([[1], [8]])
    await undo(p)
    expect(await raw(p, 'C1:C2')).toEqual([[9], [8]])
    await activate(p, 'C1:C2')
    dialog = await openDialog(p)
    await dialog.getByRole('radio', { name: 'Values', exact: true }).check()
    await dialog.getByLabel('Transpose').check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    // transposed: the 2×1 block becomes 1×2 at C1, never tiled over the selection
    await expect.poll(() => raw(p, 'C1:D1')).toEqual([[1, null]])

    // Transpose keeps each cell's format: bold goes with the first cell
    await setValues(p, 'F1:G1', [[1, 2]])
    await activate(p, 'F1')
    await bold(p)
    await activate(p, 'F1:G1')
    await copy(p)
    await activate(p, 'I1')
    await pasteMenu(p, 'Transpose')
    await expect.poll(() => raw(p, 'I1:I2')).toEqual([[1], [2]])
    expect((await style(p, 'I1'))?.bl).toBe(1)
    expect((await style(p, 'I2'))?.bl).not.toBe(1)
  } finally {
    await kill(launched)
  }
})

test('Paste Special keeps the formulas of cells a fill made (shared-formula followers)', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A2:A6', [[1], [2], [3], [4], [5]])
    await p.evaluate(`${api}.getRange('B2').setFormula('=A2*2')`)
    await activate(p, 'B2:B6')
    await p.locator('[data-ribbon-body] [aria-label="Fill"]').first().click()
    await p.locator('.menu-select-drop button', { hasText: 'Fill Down' }).first().click()
    await expect
      .poll(() => formulas(p, 'B2:B6'))
      .toEqual([['=A2*2'], ['=A3*2'], ['=A4*2'], ['=A5*2'], ['=A6*2']])

    await activate(p, 'B2:B6')
    await copy(p)
    await activate(p, 'D2')
    await pasteMenu(p, 'Formulas & Number Formatting')
    // every pasted cell is still a formula, moved two columns (Excel: =C2*2 …)
    await expect
      .poll(() => formulas(p, 'D2:D6'))
      .toEqual([['=C2*2'], ['=C3*2'], ['=C4*2'], ['=C5*2'], ['=C6*2']])
    await undo(p)

    // an operation onto filled cells wraps their formulas instead of replacing them with numbers
    await setValues(p, 'F1', [[5]])
    await activate(p, 'F1')
    await copy(p)
    await activate(p, 'B2:B3')
    const dialog = await openDialog(p)
    await dialog.getByRole('radio', { name: 'Values', exact: true }).check()
    await dialog.getByRole('radio', { name: 'Add' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => formulas(p, 'B2:B3')).toEqual([['=(A2*2)+5'], ['=(A3*2)+5']])
  } finally {
    await kill(launched)
  }
})
