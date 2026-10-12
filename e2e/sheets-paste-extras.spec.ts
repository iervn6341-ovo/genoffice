import { expect, test, type Page } from '@playwright/test'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Paste Special… beyond the cells: Comments and notes, Validation, merged cells, "All merging
 * conditional formats", and a clipboard another program replaced. The expectations were read from
 * Excel 16 through COM (Range.PasteSpecial) — see docs/qa/ribbon-dropdown-audit-2026-10-11.md.
 */
async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-paste-extras',
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
const univer = '(window).__genofficeDebug.univerAPI'
const api = `${univer}.getActiveWorkbook().getActiveSheet()`
const activate = (p: Page, a1: string) => p.evaluate(`${api}.getRange('${a1}').activate()`)
const raw = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getRawValues()`) as Promise<unknown[][]>
const setValues = (p: Page, a1: string, data: unknown[][]) =>
  p.evaluate(`${api}.getRange('${a1}').setValues(${JSON.stringify(data)})`)
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(500)
}
const redo = async (p: Page) => {
  await p.locator('button[aria-label^="Redo"]').first().click()
  await p.waitForTimeout(500)
}
const copy = async (p: Page) => {
  await p.locator('[data-tip^="Copy"]').first().click()
  await p.waitForTimeout(400)
}
const openDialog = async (p: Page) => {
  await p.locator('[data-ribbon-body] [aria-label="Paste Special"]').first().click()
  await p
    .locator('.menu-select-drop button', { hasText: /^Paste Special…$/ })
    .first()
    .click()
  const dialog = p.locator('.paste-special-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}
const pasteWith = async (
  p: Page,
  what: string | RegExp,
  extra: { transpose?: boolean } = {},
): Promise<void> => {
  const dialog = await openDialog(p)
  await dialog.getByRole('radio', { name: what, exact: true }).check()
  if (extra.transpose) await dialog.getByLabel('Transpose').check()
  await dialog.getByRole('button', { name: 'OK' }).click()
  await p.waitForTimeout(700)
}
/** another program takes the system clipboard: plain text, no sheet html */
async function foreignClipboard(l: LaunchedApp): Promise<void> {
  await l.app.evaluate(({ clipboard }) => {
    clipboard.writeText('from elsewhere')
  })
}
const notes = (p: Page) =>
  p.evaluate(
    `${api}.getNotes().map((n) => [n.row, n.col, n.note]).sort((a, b) => a[0] - b[0] || a[1] - b[1])`,
  ) as Promise<[number, number, string][]>
const addNote = (p: Page, a1: string, row: number, col: number, text: string) =>
  p.evaluate(
    `${api}.getRange('${a1}').createOrUpdateNote({ id: 'n-${row}-${col}', row: ${row}, col: ${col}, width: 220, height: 90, note: ${JSON.stringify(text)} })`,
  )
const ruleAt = (p: Page, a1: string) =>
  p.evaluate(
    `(() => { const v = ${api}.getRange('${a1}').getDataValidation(); return v ? { type: v.rule.type, formula1: v.rule.formula1 } : null })()`,
  ) as Promise<{ type: string; formula1?: string } | null>
const merges = (p: Page) =>
  p.evaluate(
    `${api}.getMergedRanges().map((r) => [r.getRow(), r.getColumn(), r.getHeight(), r.getWidth()])`,
  ) as Promise<number[][]>

test('the Paste Special dialog lists Excel’s twelve choices in Excel’s order', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1', [[1]])
    await activate(p, 'A1')
    await copy(p)
    await activate(p, 'C1')
    const dialog = await openDialog(p)
    await expect(dialog.locator('fieldset').first().locator('label')).toHaveText([
      'All',
      'Formulas',
      'Values',
      'Formats',
      'Comments and notes',
      'Validation',
      'All using source theme',
      'All except borders',
      'Column widths',
      'Formulas and number formats',
      'Values and number formats',
      'All merging conditional formats',
    ])
    // Comments and Validation take no operation or Skip blanks; Transpose still applies
    await dialog.getByRole('radio', { name: 'Comments and notes' }).check()
    await expect(dialog.getByLabel('Skip blanks')).toBeDisabled()
    await expect(dialog.getByRole('radio', { name: 'Add' })).toBeDisabled()
    await expect(dialog.getByLabel('Transpose')).toBeEnabled()
  } finally {
    await kill(launched)
  }
})

test('Comments replaces the destination notes with the source’s and removes them where the source has none', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:A2', [['src'], ['plain']])
    await addNote(p, 'A1', 0, 0, 'note one')
    await setValues(p, 'C1:C2', [['keep'], ['keep2']])
    await addNote(p, 'C2', 1, 2, 'old note')
    await activate(p, 'A1:A2')
    await copy(p)
    await activate(p, 'C1')
    await pasteWith(p, 'Comments and notes')
    // Excel (COM): C1 gets 'note one'; C2's old note goes (A2 has none); the values stay
    await expect
      .poll(() => notes(p))
      .toEqual(
        [
          [0, 0, 'note one'],
          [0, 2, 'note one'],
        ].sort(),
      )
    expect(await raw(p, 'C1:C2')).toEqual([['keep'], ['keep2']])
    // one ⌘Z puts the old note back and takes the pasted one off; ⇧⌘Z does it again
    await undo(p)
    await expect
      .poll(() => notes(p))
      .toEqual([
        [0, 0, 'note one'],
        [1, 2, 'old note'],
      ])
    await redo(p)
    await expect
      .poll(() => notes(p))
      .toEqual([
        [0, 0, 'note one'],
        [0, 2, 'note one'],
      ])
  } finally {
    await kill(launched)
  }
})

test('Validation copies the source rule over the pasted cells and clears it where the source has none', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:A2', [[1], [2]])
    await p.evaluate(
      `${api}.getRange('A1').setDataValidation(${univer}.newDataValidation().requireValueInList(['a', 'b', 'c']).build())`,
    )
    await p.evaluate(
      `${api}.getRange('E2').setDataValidation(${univer}.newDataValidation().requireNumberBetween(1, 9).build())`,
    )
    expect((await ruleAt(p, 'A1'))?.type).toBe('list')
    await activate(p, 'A1')
    await copy(p)
    // one source cell over a 1 × 2 selection: the rule lands on both (Excel tiles it)
    await activate(p, 'E1:E2')
    await pasteWith(p, 'Validation')
    await expect.poll(async () => (await ruleAt(p, 'E1'))?.type).toBe('list')
    expect((await ruleAt(p, 'E2'))?.type).toBe('list')
    await undo(p)
    await expect.poll(async () => await ruleAt(p, 'E1')).toBeNull()
    expect((await ruleAt(p, 'E2'))?.type).toBe('decimal')
    await redo(p)
    await expect.poll(async () => (await ruleAt(p, 'E2'))?.type).toBe('list')

    // a source cell without a rule clears the destination's
    await activate(p, 'A2')
    await copy(p)
    await activate(p, 'E2')
    await pasteWith(p, 'Validation')
    await expect.poll(async () => await ruleAt(p, 'E2')).toBeNull()
    expect((await ruleAt(p, 'E1'))?.type).toBe('list')
  } finally {
    await kill(launched)
  }
})

test('merged cells travel with All and Formats, turn with Transpose, and stay out of Values', async () => {
  const { launched, p } = await openSheets()
  try {
    await p.evaluate(`${api}.getRange('A1:B1').merge()`)
    await setValues(p, 'A1', [['m']])
    await p.evaluate(`${api}.getRange('A1').setBackground('#00ff00')`)
    await activate(p, 'A1:B1')
    await copy(p)

    // Values: no merge (Excel: D3 merge=False)
    await activate(p, 'D3')
    await pasteWith(p, 'Values')
    expect((await merges(p)).filter((m) => m[0] === 2)).toEqual([])
    expect(await raw(p, 'D3')).toEqual([['m']])

    // All + Transpose: the merge turns — A1:B1 becomes D13:D14 (Excel: mergearea $D$13:$D$14)
    await activate(p, 'D13')
    await pasteWith(p, 'All', { transpose: true })
    await expect.poll(() => merges(p)).toContainEqual([12, 3, 2, 1])
    expect(await raw(p, 'D13')).toEqual([['m']])

    // Formats + Transpose carries the merge too, and no value
    await activate(p, 'F3')
    await pasteWith(p, 'Formats', { transpose: true })
    await expect.poll(() => merges(p)).toContainEqual([2, 5, 2, 1])
    expect(await raw(p, 'F3')).toEqual([[null]])

    // one ⌘Z takes the last paste's merge off
    await undo(p)
    await expect.poll(() => merges(p)).not.toContainEqual([2, 5, 2, 1])
    expect(await merges(p)).toContainEqual([12, 3, 2, 1])
  } finally {
    await kill(launched)
  }
})

test('All merging conditional formats keeps the destination’s rules and adds the source’s', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:B1', [[10, 1]])
    const addRule = (a1: string, threshold: number, color: string) =>
      p.evaluate(
        `(() => { const s = ${api}; s.addConditionalFormattingRule(s.newConditionalFormattingRule().whenNumberGreaterThan(${threshold}).setBackground('${color}').setRanges([s.getRange('${a1}').getRange()]).build()) })()`,
      )
    await addRule('A1', 5, '#ff0000')
    await addRule('B1', 100, '#0000ff')
    const rulesAt = (a1: string) =>
      p.evaluate(
        `(() => { const [r, c] = [${api}.getRange('${a1}').getRow(), ${api}.getRange('${a1}').getColumn()]; return ${api}.getConditionalFormattingRules().filter((rule) => rule.ranges.some((x) => r >= x.startRow && r <= x.endRow && c >= x.startColumn && c <= x.endColumn)).length })()`,
      ) as Promise<number>
    expect(await rulesAt('B1')).toBe(1)
    await activate(p, 'A1')
    await copy(p)
    await activate(p, 'B1')
    await pasteWith(p, 'All merging conditional formats')
    // Excel: B1 now has its own rule and the source's (COM: FormatConditions.Count = 2)
    await expect.poll(() => rulesAt('B1')).toBe(2)
    expect(await raw(p, 'B1')).toEqual([[10]])
    await undo(p)
    await expect.poll(() => rulesAt('B1')).toBe(1)
    expect(await raw(p, 'B1')).toEqual([[1]])
  } finally {
    await kill(launched)
  }
})

test('a clipboard another program replaced is not pasted from the older copy', async () => {
  const { launched, p } = await openSheets()
  try {
    await setValues(p, 'A1:A2', [[1], [2]])
    await setValues(p, 'C1:C2', [[10], [20]])
    await activate(p, 'A1:A2')
    await copy(p)
    // another program takes the system clipboard (this is what failed once in a long run, when an
    // Excel automation script copied cells at the same time)
    await foreignClipboard(launched)
    await activate(p, 'C1:C2')
    const dialog = await openDialog(p)
    await dialog.getByRole('radio', { name: 'Values', exact: true }).check()
    await dialog.getByRole('radio', { name: 'Add' }).check()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(p.locator('.workbook-status')).toContainText('Copy cells')
    expect(await raw(p, 'C1:C2')).toEqual([[10], [20]])
  } finally {
    await kill(launched)
  }
})
