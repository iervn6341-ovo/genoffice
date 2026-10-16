import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { copyFile, mkdir, mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

const EVIDENCE_DIRECTORY =
  process.env.GENOFFICE_FORMULA_SYMBOL_VERIFICATION_DIR ??
  join(process.cwd(), 'docs/qa/office-mac-formula-symbol-evidence-2026-10-13/sheets')

test.skip(!process.env.GENOFFICE_OFFICE_AUDIT, 'Opt-in Sheets formula symbol verification')
test.use({ actionTimeout: 10_000 })

async function sheetsPage(app: Awaited<ReturnType<typeof launchShell>>['app']): Promise<Page> {
  const page = await waitForPageWithUrl(app, '://sheets/', 30_000)
  await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
  await setEditorLayoutWidth(app, '://sheets/', 1280)
  await page.waitForFunction(() =>
    [...document.querySelectorAll('canvas')].some((c) => c.getBoundingClientRect().height > 100),
  )
  return page
}

async function prepareSymbolRibbon(page: Page): Promise<void> {
  await page
    .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', {
      hasText: /^Insert$/,
    })
    .first()
    .click()
  const group = page.locator('.rb-fold-group[aria-label="Symbols"]')
  if ((await group.getAttribute('data-folded')) !== null) {
    await group.locator(':scope > .rb-fold-btn').click()
  }
}

async function clickSymbol(page: Page): Promise<void> {
  const group = page.locator('.rb-fold-group[aria-label="Symbols"]')
  const command = group.locator('button[data-tip="Symbol"]')
  if (!(await command.isVisible())) await group.locator(':scope > .rb-fold-btn').click()
  await command.click()
}

async function beginCellEdit(page: Page, reference: string): Promise<void> {
  const point = await page.evaluate((ref) => {
    const book = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
    const range = book.getActiveSheet().getRange(ref)
    range.activate()
    const cell = range.getCellRect()
    const canvas = [...document.querySelectorAll('canvas')]
      .filter((node) => {
        const rect = node.getBoundingClientRect()
        return rect.width > 100 && rect.height > 100
      })
      .at(-1)
    if (!canvas) throw new Error('Worksheet canvas missing')
    const bounds = canvas.getBoundingClientRect()
    return {
      x: bounds.left + (cell.left + cell.right) / 2,
      y: bounds.top + (cell.top + cell.bottom) / 2,
    }
  }, reference)
  // Native input establishes Univer's edit location; a facade startEditing
  // immediately after modal focus loss can see an unset edit-location state.
  await page.mouse.click(point.x, point.y)
  await page.keyboard.press('F2')
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().isCellEditing(),
      ),
    )
    .toBe(true)
  await page.waitForFunction(
    () => (document.activeElement as HTMLElement | null)?.isContentEditable === true,
  )
}

async function selectLiveFormulaLiteral(page: Page): Promise<void> {
  await beginCellEdit(page, 'A1')
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().isCellEditing(),
      ),
    )
    .toBe(true)
  await page.evaluate(() => {
    const debug = (window as any).__genofficeDebug
    const document = debug.univerAPI.getActiveDocument()
    if (!document || document.getId() !== debug.formulaSymbolSelectionDiagnostic().editorId)
      throw new Error('The active document is not the live cell editor')
    // Use Univer's public selection facade: it updates the real rendered
    // caret/range. Arrow keys in a formula can instead select cell references.
    document.setSelection(2, 7)
  })
  await expect
    .poll(
      () =>
        page.evaluate(() => {
          const draft = (window as any).__genofficeDebug.formulaSymbolDraft()
          return { start: draft?.selectionStart, end: draft?.selectionEnd }
        }),
      { timeout: 5000 },
    )
    .toEqual({ start: 2, end: 7 })
}

test('symbol replaces formula literal selection while references recalculate after save and reopen', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-formula-symbol-'))
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.app.evaluate(({ app }, path) => app.setPath('documents', path), scratch)
  let workbook = ''
  try {
    await launched.page.locator('.quick-card').nth(1).click()
    const page = await sheetsPage(launched.app)
    workbook = join(
      scratch,
      'GenOffice',
      (await readdir(join(scratch, 'GenOffice'))).find((file) => file.endsWith('.xlsx'))!,
    )
    await page.evaluate(() => {
      const book = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
      const sheet = book.getActiveSheet()
      sheet.getRange('B1').setValue(2)
      sheet.getRange('B2').setValue(3)
      sheet.getRange('A1').setValue({ f: '="Total "&SUM(B1:B2)' })
      sheet.getRange('D1').setValue({ f: '=SUM(B1:B2)' })
      sheet.getRange('A1').activate()
    })
    await prepareSymbolRibbon(page)
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('D1').activate()
    })
    await clickSymbol(page)
    await expect(
      page.getByRole('dialog', { name: 'Symbol' }).getByRole('textbox', { name: 'Formula' }),
    ).toHaveValue('=SUM(B1:B2)')
    await page
      .getByRole('dialog', { name: 'Symbol' })
      .getByRole('button', { name: 'Cancel' })
      .click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('D1').getFormula()
        }),
      )
      .toBe('=SUM(B1:B2)')
    await beginCellEdit(page, 'A1')
    await page.keyboard.press('End')
    await page.keyboard.type('+1')
    await clickSymbol(page)
    const dialog = page.getByRole('dialog', { name: 'Symbol' })
    const source = dialog.getByRole('textbox', { name: 'Formula' })
    await expect(source).toHaveValue('="Total "&SUM(B1:B2)+1')
    // Cancel discards the unsaved live-editor draft, retaining the saved formula.
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('A1').getFormula()
        }),
      )
      .toBe('="Total "&SUM(B1:B2)')
    await selectLiveFormulaLiteral(page)
    await clickSymbol(page)
    const secondDialog = page.getByRole('dialog', { name: 'Symbol' })
    const secondSource = secondDialog.getByRole('textbox', { name: 'Formula' })
    await expect(secondSource).toHaveValue('="Total "&SUM(B1:B2)')
    await secondDialog.getByLabel('Character code').fill('03A9')
    await secondDialog.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(secondSource).toHaveValue('="Ω "&SUM(B1:B2)')
    await secondDialog.getByRole('button', { name: 'Apply' }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('A1').getFormula()
        }),
      )
      .toBe('="Ω "&SUM(B1:B2)')
    // A currently typed ordinary cell must be committed before the rich
    // symbol patch reads it; the newly typed prefix may not be in the grid yet.
    await beginCellEdit(page, 'C1')
    await page.keyboard.type('Draft')
    await clickSymbol(page)
    const plainDialog = page.getByRole('dialog', { name: 'Symbol' })
    await expect(plainDialog.getByRole('textbox', { name: 'Formula' })).toHaveCount(0)
    await plainDialog.getByLabel('Character code').fill('03A9')
    await plainDialog.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('C1').getValue()
        }),
      )
      .toBe('DraftΩ')
    expect(
      await page.evaluate(
        () =>
          (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange('C2')
            .getValue() ?? null,
      ),
    ).toBeNull()
    await plainDialog.getByRole('button', { name: 'Close' }).click()
    await launched.app.evaluate(({ webContents }) => {
      webContents
        .getAllWebContents()
        .find((wc) => wc.getURL().includes('://sheets/'))
        ?.send('menu:action', 'save')
    })
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(await readFile(workbook))
          return (
            (await zip.file('xl/worksheets/sheet1.xml')?.async('string'))?.includes(
              '<f>"Ω "&amp;SUM(B1:B2)</f>',
            ) ?? false
          )
        },
        { timeout: 20_000 },
      )
      .toBe(true)
    await mkdir(EVIDENCE_DIRECTORY, { recursive: true })
    await writeFile(
      join(EVIDENCE_DIRECTORY, 'formula-symbol-first-apply.png'),
      await page.screenshot(),
    )
    await copyFile(workbook, join(EVIDENCE_DIRECTORY, 'formula-symbol-first-save.xlsx'))
  } finally {
    await closeAndSaveVideo(launched, 'sheets-formula-symbol-first')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    openFile: workbook,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await sheetsPage(reopened.app)
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('A1').getFormula()
        }),
      )
      .toBe('="Ω "&SUM(B1:B2)')
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('B1').setValue(5)
    })
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('A1').getDisplayValue()
        }),
      )
      .toBe('Ω 8')
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('A1').activate()
    })
    await prepareSymbolRibbon(page)
    await clickSymbol(page)
    const dialog = page.getByRole('dialog', { name: 'Symbol' })
    const source = dialog.getByRole('textbox', { name: 'Formula' })
    await expect(source).toHaveValue('="Ω "&SUM(B1:B2)')
    await source.evaluate((element) => {
      const field = element as HTMLTextAreaElement
      field.focus()
      field.setSelectionRange(3, 3)
    })
    await dialog.getByLabel('Character code').fill('03B3')
    await dialog.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(source).toHaveValue('="Ωγ "&SUM(B1:B2)')
    await dialog.getByRole('button', { name: 'Apply' }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('A1').getDisplayValue()
        }),
      )
      .toBe('Ωγ 8')
    await reopened.app.evaluate(({ webContents }) => {
      webContents
        .getAllWebContents()
        .find((wc) => wc.getURL().includes('://sheets/'))
        ?.send('menu:action', 'save')
    })
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(await readFile(workbook))
          const xml = (await zip.file('xl/worksheets/sheet1.xml')?.async('string')) ?? ''
          return (
            xml.includes('<f>"Ωγ "&amp;SUM(B1:B2)</f>') &&
            xml.includes('r="B1"') &&
            xml.includes('<v>5</v>')
          )
        },
        { timeout: 20_000 },
      )
      .toBe(true)
    await writeFile(
      join(EVIDENCE_DIRECTORY, 'formula-symbol-second-apply.png'),
      await page.screenshot(),
    )
    await copyFile(workbook, join(EVIDENCE_DIRECTORY, 'formula-symbol-second-save.xlsx'))
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-formula-symbol-reopen')
  }

  const third = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    openFile: workbook,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await sheetsPage(third.app)
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('A1').getFormula()
        }),
      )
      .toBe('="Ωγ "&SUM(B1:B2)')
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('A1').getDisplayValue()
        }),
      )
      .toBe('Ωγ 8')
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('A1').activate()
    })
    await prepareSymbolRibbon(page)
    await clickSymbol(page)
    await expect(
      page.getByRole('dialog', { name: 'Symbol' }).getByRole('textbox', { name: 'Formula' }),
    ).toHaveValue('="Ωγ "&SUM(B1:B2)')
    await writeFile(
      join(EVIDENCE_DIRECTORY, 'formula-symbol-third-reopen.png'),
      await page.screenshot(),
    )
    await writeFile(
      join(EVIDENCE_DIRECTORY, 'formula-symbol-summary.json'),
      JSON.stringify(
        {
          firstFormula: '="Ω "&SUM(B1:B2)',
          finalFormula: '="Ωγ "&SUM(B1:B2)',
          finalDisplay: 'Ωγ 8',
          firstWorkbook: 'formula-symbol-first-save.xlsx',
          secondWorkbook: 'formula-symbol-second-save.xlsx',
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(third, 'sheets-formula-symbol-third-reopen')
  }
})
