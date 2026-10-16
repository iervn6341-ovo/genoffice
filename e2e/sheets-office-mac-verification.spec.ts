import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { basename, dirname, join } from 'node:path'
import { createRequire } from 'node:module'
import { mkdtemp, readdir, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'

import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

const RIBBON_TAB = 'nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)'
const req = createRequire(join(process.cwd(), 'package.json'))

type Launched = Awaited<ReturnType<typeof launchShell>>
type SavedPackagePredicate = (zip: JSZip) => Promise<boolean>

test.skip(!process.env.GENOFFICE_OFFICE_AUDIT, 'Opt-in Office parity verification')
test.use({ actionTimeout: 10_000 })

async function waitForSheetsPage(launched: Launched): Promise<Page> {
  const sheet = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await sheet.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await sheet.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI), null, {
    timeout: 30_000,
  })
  await sheet.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  return sheet
}

async function openBlank(videoDir: string): Promise<{
  launched: Launched
  sheet: Page
  workbook: string
}> {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-office-verify-'))
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.app.evaluate(({ app }, path) => app.setPath('documents', path), scratch)
  await launched.page.locator('.quick-card').nth(1).click()
  const sheet = await waitForSheetsPage(launched)
  await sheet.waitForTimeout(1_000)
  const files = await readdir(join(scratch, 'GenOffice'))
  const workbooks = files.filter((file) => file.endsWith('.xlsx'))
  expect(workbooks).toHaveLength(1)
  return { launched, sheet, workbook: join(scratch, 'GenOffice', workbooks[0]!) }
}

async function openExisting(
  workbook: string,
  videoDir: string,
): Promise<{
  launched: Launched
  sheet: Page
}> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir,
    openFile: workbook,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  return { launched, sheet: await waitForSheetsPage(launched) }
}

const selectTab = (page: Page, name: string): Promise<void> =>
  page
    .locator(RIBBON_TAB, { hasText: new RegExp(`^${name}$`) })
    .first()
    .click()

async function packageXml(zip: JSZip, path: string): Promise<string> {
  return (await zip.file(path)?.async('string')) ?? ''
}

function decodeXmlEntities(text: string): string {
  const entities: Record<string, string> = {
    amp: '&',
    quot: '"',
    apos: "'",
    lt: '<',
    gt: '>',
  }
  return text.replace(/&(amp|quot|apos|lt|gt);/g, (_match, entity: string) => entities[entity]!)
}

function headerFooterSectionText(xml: string, tag: string): string | null {
  const section = new RegExp('<' + tag + '\\b[^>]*>([\\s\\S]*?)</' + tag + '>').exec(xml)
  return section ? decodeXmlEntities(section[1] ?? '') : null
}

async function saveFromApp(
  launched: Launched,
  workbook: string,
  expected: SavedPackagePredicate,
  onExpectationFailure?: () => Promise<void>,
): Promise<JSZip> {
  await launched.app.evaluate(({ webContents }) => {
    const wc = webContents.getAllWebContents().find((item) => item.getURL().includes('://sheets/'))
    wc?.send('menu:action', 'save')
  })
  try {
    await expect
      .poll(async () => expected(await JSZip.loadAsync(await readFile(workbook))), {
        timeout: 20_000,
      })
      .toBe(true)
  } catch (error) {
    await onExpectationFailure?.()
    throw error
  }
  return JSZip.loadAsync(await readFile(workbook))
}

async function attachScreenshot(page: Page, name: string): Promise<void> {
  await test.info().attach(name, { body: await page.screenshot(), contentType: 'image/png' })
}

async function createValidationRule(page: Page, options: readonly string[]): Promise<unknown[]> {
  await page.evaluate(() => {
    const api = (window as any).__genofficeDebug.univerAPI
    api.getActiveWorkbook().getActiveSheet().getRange('B2:B4').activate()
  })
  await selectTab(page, 'Data')
  await page.locator('[data-ribbon-body] [data-tip="Data Validation"]').first().click()
  await expect(page.getByRole('button', { name: 'Add Rule' })).toBeVisible()
  await page.getByRole('button', { name: 'Add Rule' }).click()

  const detail = page.locator('[data-u-comp="data-validation-detail"]')
  await expect(detail).toBeVisible()
  const typeSelect = detail
    .getByText('Type', { exact: true })
    .locator('..')
    .locator('[data-u-comp="select"]')
  await typeSelect.click()
  await expect(page.getByText('Dropdown', { exact: true }).last()).toBeVisible()
  await page.getByText('Dropdown', { exact: true }).last().click()
  await expect(detail.getByText('Options', { exact: true })).toBeVisible()

  const listValues = () =>
    detail.locator('[data-draggable-list-item-id] [data-u-comp="input"] input')
  for (const [index, option] of options.entries()) {
    await detail.getByText('Add', { exact: true }).click()
    await expect(listValues()).toHaveCount(index + 1)
    await listValues().nth(index).fill(option)
  }
  const liveRules = async () =>
    page.evaluate(() =>
      (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getDataValidations()
        .map(({ rule }: { rule: Record<string, unknown> }) => ({
          type: rule.type,
          formula1: rule.formula1,
          formula2: rule.formula2,
          ranges: rule.ranges,
        })),
    )
  await expect.poll(liveRules).toEqual([
    expect.objectContaining({
      type: 'list',
      formula1: JSON.stringify(options),
      formula2: ',',
    }),
  ])
  const rules = await liveRules()
  await detail.getByRole('button', { name: 'Done' }).click()
  await expect(detail).toBeHidden()
  return rules
}

async function clickAppAction(launched: Launched, action: string): Promise<void> {
  await launched.app.evaluate(({ webContents }, requestedAction) => {
    const wc = webContents.getAllWebContents().find((item) => item.getURL().includes('://sheets/'))
    wc?.send('menu:action', requestedAction)
  }, action)
}

async function openHeaderFooterDialog(launched: Launched, page: Page) {
  // At the default 1280px viewport, Insert's Text group can fold into a flyout.
  // Expand the logical ribbon width and target the actual command by its tooltip.
  await setEditorLayoutWidth(launched.app, '://sheets/', 2200)
  await selectTab(page, 'Insert')
  const button = page.locator('[data-ribbon-body] button[data-tip="Header & Footer"]')
  await expect(button).toBeVisible()
  await button.click()
  const dialog = page.getByRole('dialog', { name: 'Header & Footer' })
  await expect(dialog).toBeVisible()
  return dialog
}

async function exportedPageTexts(file: string): Promise<string[]> {
  const pdfjs = await import(req.resolve('pdfjs-dist/legacy/build/pdf.mjs'))
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(await readFile(file)),
    disableFontFace: true,
    verbosity: 0,
  }).promise
  try {
    const pages: string[] = []
    for (let number = 1; number <= doc.numPages; number += 1) {
      const content = await (await doc.getPage(number)).getTextContent()
      pages.push(content.items.map((item: { str?: string }) => item.str ?? '').join(''))
    }
    return pages
  } finally {
    await doc.loadingTask.destroy()
  }
}

test('Data Validation: UI list preserves label spaces in OOXML and after reopen', async () => {
  const { launched, sheet, workbook } = await openBlank('sheets-office-mac-verify-dv')
  try {
    await sheet.evaluate(() => {
      const api = (window as any).__genofficeDebug.univerAPI
      api.getActiveWorkbook().getActiveSheet().getRange('B2:B4').activate()
    })
    await selectTab(sheet, 'Data')
    await sheet.locator('[data-ribbon-body] [data-tip="Data Validation"]').first().click()
    await expect(sheet.getByRole('button', { name: 'Add Rule' })).toBeVisible()
    await sheet.getByRole('button', { name: 'Add Rule' }).click()

    const detail = sheet.locator('[data-u-comp="data-validation-detail"]')
    await expect(detail).toBeVisible()
    const typeSelect = detail
      .getByText('Type', { exact: true })
      .locator('..')
      .locator('[data-u-comp="select"]')
    await typeSelect.click()
    await sheet.getByText('Dropdown', { exact: true }).last().click()
    const listValues = () =>
      detail.locator('[data-draggable-list-item-id] [data-u-comp="input"] input')
    await detail.getByText('Add', { exact: true }).click()
    await expect(listValues()).toHaveCount(1)
    await listValues().nth(0).fill('  Open  ')
    await detail.getByText('Add', { exact: true }).click()
    await expect(listValues()).toHaveCount(2)
    await listValues().nth(1).fill('Closed ')

    const liveRules = async () =>
      sheet.evaluate(() =>
        (window as any).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveSheet()
          .getDataValidations()
          .map(({ rule }: { rule: Record<string, unknown> }) => ({
            type: rule.type,
            formula1: rule.formula1,
            formula2: rule.formula2,
            ranges: rule.ranges,
          })),
      )
    await expect.poll(liveRules).toEqual([
      expect.objectContaining({
        type: 'list',
        formula1: JSON.stringify(['  Open  ', 'Closed ']),
        formula2: ',',
      }),
    ])
    await attachScreenshot(sheet, 'sheets-dv-whitespace-editor.png')
    await detail.getByRole('button', { name: 'Done' }).click()
    await expect(detail).toBeHidden()
    await expect.poll(liveRules).toEqual([
      expect.objectContaining({
        type: 'list',
        formula1: JSON.stringify(['  Open  ', 'Closed ']),
        formula2: ',',
      }),
    ])
    const appliedRules = await liveRules()
    await test.info().attach('sheets-dv-whitespace-live-rule.json', {
      body: JSON.stringify(appliedRules, null, 2),
      contentType: 'application/json',
    })

    const zip = await saveFromApp(
      launched,
      workbook,
      async (candidate) => {
        const xml = await packageXml(candidate, 'xl/worksheets/sheet1.xml')
        return (
          xml.includes('sqref="B2:B4"') && xml.includes('<formula1>"  Open  ,Closed "</formula1>')
        )
      },
      async () => {
        const candidate = await JSZip.loadAsync(await readFile(workbook))
        const xml = await packageXml(candidate, 'xl/worksheets/sheet1.xml')
        const status = await sheet
          .locator('.status-msg')
          .innerText()
          .catch(() => '')
        const rules = await liveRules().catch((error: unknown) => ({
          readError: error instanceof Error ? error.message : String(error),
        }))
        await test.info().attach('sheets-dv-valid-save-failure.json', {
          body: JSON.stringify(
            {
              status,
              rules,
              workbookBytes: (await readFile(workbook)).byteLength,
              expectedRangeAndFormulaFound:
                xml.includes('sqref="B2:B4"') &&
                xml.includes('<formula1>"  Open  ,Closed "</formula1>'),
            },
            null,
            2,
          ),
          contentType: 'application/json',
        })
        await test.info().attach('sheets-dv-valid-save-failure-sheet1.xml', {
          body: xml,
          contentType: 'application/xml',
        })
        await attachScreenshot(sheet, 'sheets-dv-valid-save-failure.png')
      },
    )
    const xml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('sqref="B2:B4"')
    expect(xml).toContain('<formula1>"  Open  ,Closed "</formula1>')
    expect(xml).not.toContain('<formula2>')
    await test.info().attach('sheets-dv-whitespace-sheet1.xml', {
      body: xml,
      contentType: 'application/xml',
    })
  } finally {
    await closeAndSaveVideo(launched, 'sheets-office-mac-verify-dv')
  }

  const reopened = await openExisting(workbook, 'sheets-office-mac-verify-dv-reopen')
  try {
    const rules = await reopened.sheet.evaluate(() =>
      (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getDataValidations()
        .map(({ rule }: { rule: Record<string, unknown> }) => ({
          type: rule.type,
          formula1: rule.formula1,
          formula2: rule.formula2,
          ranges: rule.ranges,
        })),
    )
    expect(rules).toEqual([
      expect.objectContaining({
        type: 'list',
        formula1: JSON.stringify(['  Open  ', 'Closed ']),
        ranges: [expect.objectContaining({ startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 })],
      }),
    ])
    await test.info().attach('sheets-dv-whitespace-reopened-live-rule.json', {
      body: JSON.stringify(rules, null, 2),
      contentType: 'application/json',
    })
    await reopened.sheet.evaluate(() => {
      const api = (window as any).__genofficeDebug.univerAPI
      api.getActiveWorkbook().getActiveSheet().getRange('B2').activate()
    })
    // The dropdown affordance is a floating active-cell component, so it is
    // not guaranteed to exist as a normal DOM button in every renderer state.
    // Reopen the supported Data Validation panel and inspect the saved rule's
    // actual editable option fields instead of assuming a CSS implementation.
    await selectTab(reopened.sheet, 'Data')
    await reopened.sheet.locator('[data-ribbon-body] [data-tip="Data Validation"]').first().click()
    await expect(reopened.sheet.getByRole('button', { name: 'Add Rule' })).toBeVisible()
    await reopened.sheet.getByText('B2:B4', { exact: true }).click()
    const reopenedDetail = reopened.sheet.locator('[data-u-comp="data-validation-detail"]')
    await expect(reopenedDetail).toBeVisible()
    const reopenedOptions = reopenedDetail.locator(
      '[data-draggable-list-item-id] [data-u-comp="input"] input',
    )
    await expect(reopenedOptions).toHaveCount(2)
    await expect(reopenedOptions.nth(0)).toHaveValue('  Open  ')
    await expect(reopenedOptions.nth(1)).toHaveValue('Closed ')
    await test.info().attach('sheets-dv-whitespace-reopened-editor.png', {
      body: await reopened.sheet.screenshot(),
      contentType: 'image/png',
    })
    await attachScreenshot(reopened.sheet, 'sheets-dv-whitespace-reopened.png')
  } finally {
    await closeAndSaveVideo(reopened.launched, 'sheets-office-mac-verify-dv-reopen')
  }
})

test('Header & Footer: first/even variants survive UI save, reopen, and PDF export', async () => {
  test.setTimeout(300_000)
  const { launched, sheet, workbook } = await openBlank('sheets-office-mac-verify-hf')
  try {
    await sheet.evaluate(() => {
      const api = (window as any).__genofficeDebug.univerAPI
      const worksheet = api.getActiveWorkbook().getActiveSheet()
      worksheet
        .getRange('A1:A180')
        .setValues(Array.from({ length: 180 }, (_, index) => [`SHEETS_HF_PAGE_ROW_${index + 1}`]))
    })
    const dialog = await openHeaderFooterDialog(launched, sheet)

    const fillField = async (label: string, value: string): Promise<void> => {
      await dialog.getByLabel(label).fill(value)
    }
    await fillField('Header · Center', 'ODD &12&"Arial,Regular"&B&P/&N&B')
    await fillField('Footer · Right', '&Z')

    await dialog.getByRole('tab', { name: 'First page' }).click()
    await dialog.getByLabel('Different first page').check()
    await fillField('Header · Center', 'FIRST &F')
    await fillField('Footer · Left', '&Z')

    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await dialog.getByLabel('Different odd and even pages').check()
    await fillField('Header · Center', 'EVEN &A')
    await fillField('Footer · Left', 'EVEN &I italic&I')
    await attachScreenshot(sheet, 'sheets-hf-variants-editor.png')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()

    const zip = await saveFromApp(launched, workbook, async (candidate) => {
      const xml = await packageXml(candidate, 'xl/worksheets/sheet1.xml')
      return (
        xml.includes('differentFirst="1"') &&
        xml.includes('differentOddEven="1"') &&
        headerFooterSectionText(xml, 'oddHeader') === '&CODD &12&"Arial,Regular"&B&P/&N&B' &&
        headerFooterSectionText(xml, 'firstHeader') === '&CFIRST &F' &&
        headerFooterSectionText(xml, 'evenHeader') === '&CEVEN &A'
      )
    })
    const xml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
    expect(xml).toContain('differentFirst="1"')
    expect(xml).toContain('differentOddEven="1"')
    expect(headerFooterSectionText(xml, 'oddFooter')).toBe('&R&Z')
    expect(headerFooterSectionText(xml, 'firstFooter')).toBe('&L&Z')
    expect(headerFooterSectionText(xml, 'evenFooter')).toBe('&LEVEN &I italic&I')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-office-mac-verify-hf')
  }

  const reopened = await openExisting(workbook, 'sheets-office-mac-verify-hf-reopen')
  try {
    const dialog = await openHeaderFooterDialog(reopened.launched, reopened.sheet)
    await expect(dialog.getByLabel('Header · Center')).toHaveValue(
      'ODD &12&"Arial,Regular"&B&P/&N&B',
    )
    await dialog.getByRole('tab', { name: 'First page' }).click()
    await expect(dialog.getByLabel('Different first page')).toBeChecked()
    await expect(dialog.getByLabel('Header · Center')).toHaveValue('FIRST &F')
    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await expect(dialog.getByLabel('Different odd and even pages')).toBeChecked()
    await expect(dialog.getByLabel('Header · Center')).toHaveValue('EVEN &A')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    const scratchPdf = join(dirname(workbook), 'sheets-hf-preview.pdf')
    await reopened.launched.app.evaluate(({ dialog: nativeDialog }, path) => {
      nativeDialog.showSaveDialog = (async () => ({ canceled: false, filePath: path })) as never
    }, scratchPdf)
    await clickAppAction(reopened.launched, 'export-pdf')
    await expect
      .poll(async () => (await readFile(scratchPdf).catch(() => Buffer.alloc(0))).byteLength, {
        timeout: 30_000,
      })
      .toBeGreaterThan(1_000)
    const pages = await exportedPageTexts(scratchPdf)
    expect(pages.length).toBeGreaterThanOrEqual(3)
    expect(pages[0]).toContain('FIRST')
    expect(pages[0]).toContain(basename(workbook))
    expect(pages[0]).toContain(dirname(workbook))
    expect(pages[1]).toContain('EVEN')
    expect(pages[1]).toContain('Sheet1')
    expect(pages[2]).toContain('ODD')
    await test.info().attach('sheets-hf-preview-page-text.json', {
      body: JSON.stringify(pages.slice(0, 4), null, 2),
      contentType: 'application/json',
    })
    await attachScreenshot(reopened.sheet, 'sheets-hf-reopened-before-disable.png')
  } finally {
    await closeAndSaveVideo(reopened.launched, 'sheets-office-mac-verify-hf-reopen')
  }

  const disabled = await openExisting(workbook, 'sheets-office-mac-verify-hf-disabled')
  try {
    const dialog = await openHeaderFooterDialog(disabled.launched, disabled.sheet)
    await dialog.getByRole('tab', { name: 'First page' }).click()
    await expect(dialog.getByLabel('Different first page')).toBeChecked()
    await dialog.getByLabel('Different first page').uncheck()
    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await expect(dialog.getByLabel('Different odd and even pages')).toBeChecked()
    await dialog.getByLabel('Different odd and even pages').uncheck()
    await dialog.getByRole('button', { name: 'OK' }).click()

    const zip = await saveFromApp(disabled.launched, workbook, async (candidate) => {
      const xml = await packageXml(candidate, 'xl/worksheets/sheet1.xml')
      return !xml.includes('differentFirst="1"') && !xml.includes('differentOddEven="1"')
    })
    const xml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
    expect(xml).not.toContain('differentFirst="1"')
    expect(xml).not.toContain('differentOddEven="1"')
    expect(headerFooterSectionText(xml, 'firstHeader')).toBe('&CFIRST &F')
    expect(headerFooterSectionText(xml, 'evenHeader')).toBe('&CEVEN &A')
  } finally {
    await closeAndSaveVideo(disabled.launched, 'sheets-office-mac-verify-hf-disabled')
  }

  const reenabled = await openExisting(workbook, 'sheets-office-mac-verify-hf-reenabled')
  try {
    const dialog = await openHeaderFooterDialog(reenabled.launched, reenabled.sheet)
    await dialog.getByRole('tab', { name: 'First page' }).click()
    await expect(dialog.getByLabel('Different first page')).not.toBeChecked()
    await dialog.getByLabel('Different first page').check()
    await expect(dialog.getByLabel('Header · Center')).toHaveValue('FIRST &F')
    await expect(dialog.getByLabel('Footer · Left')).toHaveValue('&Z')

    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await expect(dialog.getByLabel('Different odd and even pages')).not.toBeChecked()
    await dialog.getByLabel('Different odd and even pages').check()
    await expect(dialog.getByLabel('Header · Center')).toHaveValue('EVEN &A')
    await expect(dialog.getByLabel('Footer · Left')).toHaveValue('EVEN &I italic&I')
    await dialog.getByRole('button', { name: 'OK' }).click()

    const zip = await saveFromApp(reenabled.launched, workbook, async (candidate) => {
      const xml = await packageXml(candidate, 'xl/worksheets/sheet1.xml')
      return (
        xml.includes('differentFirst="1"') &&
        xml.includes('differentOddEven="1"') &&
        headerFooterSectionText(xml, 'firstHeader') === '&CFIRST &F' &&
        headerFooterSectionText(xml, 'evenHeader') === '&CEVEN &A'
      )
    })
    const xml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
    expect(headerFooterSectionText(xml, 'firstHeader')).toBe('&CFIRST &F')
    expect(headerFooterSectionText(xml, 'evenHeader')).toBe('&CEVEN &A')
  } finally {
    await closeAndSaveVideo(reenabled.launched, 'sheets-office-mac-verify-hf-reenabled')
  }
})

test('Sheets folded Home keeps nested MenuSelect and font color palette operable', async () => {
  const { launched, sheet } = await openBlank('sheets-office-mac-verify-folded-home')
  try {
    await setEditorLayoutWidth(launched.app, '://sheets/', 1280)
    await selectTab(sheet, 'Home')
    const alignment = sheet.locator('.rb-fold-group[aria-label="Alignment"]')
    let layoutWidth = 1280
    if ((await alignment.getAttribute('data-folded')) === null) {
      // Keep the normal 1280px check above; some display scaling leaves enough
      // room there, so narrow further only when needed to exercise the flyout.
      layoutWidth = 900
      await setEditorLayoutWidth(launched.app, '://sheets/', layoutWidth)
      await selectTab(sheet, 'Home')
    }
    await expect(alignment).toHaveAttribute('data-folded', '')
    await test.info().attach('sheets-folded-home-layout.json', {
      body: JSON.stringify({ initialLayoutWidth: 1280, interactionLayoutWidth: layoutWidth }),
      contentType: 'application/json',
    })

    const foldButton = alignment.locator(':scope > .rb-fold-btn')
    await foldButton.click()
    await expect(foldButton).toHaveAttribute('aria-expanded', 'true')
    await expect(alignment.locator('.rb-fold-items')).toBeVisible()

    const orientation = alignment.locator('.rb-fold-items [aria-label="Orientation"]')
    await orientation.click()
    const options = sheet.locator('.menu-select-drop[role="listbox"][aria-label="Orientation"]')
    await expect(options).toBeVisible()
    await options.getByRole('option', { name: 'Rotate Text Up' }).click()
    await expect(foldButton).toHaveAttribute('aria-expanded', 'false')
    await expect(options).toBeHidden()

    const fontColorOptions = sheet.locator(
      '[data-ribbon-body] button[aria-label="Font color options"]',
    )
    await expect(fontColorOptions).toBeVisible()
    await fontColorOptions.click()
    const palette = sheet.locator('.gcp-palette')
    await expect(fontColorOptions).toHaveAttribute('aria-expanded', 'true')
    await expect(palette).toBeVisible()
    await attachScreenshot(sheet, 'sheets-folded-home-font-color-palette.png')
    await palette.locator('.gcp-theme-base .gcp-swatch').first().click()
    await expect(palette).toBeHidden()
    await expect(fontColorOptions).toHaveAttribute('aria-expanded', 'false')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-office-mac-verify-folded-home')
  }
})

test('Header & Footer: native codes survive edit, Undo/Redo, save and fresh reopen', async () => {
  test.setTimeout(300_000)
  const { launched, sheet, workbook } = await openBlank('sheets-hf-native-codes')
  const original = 'SHIFT&P+2/&N &D &T &F &A &"Arial,Regular"&K04+050&Edouble&E'
  const edited = `EDITED ${original}`
  const apply = async (value: string) => {
    const dialog = await openHeaderFooterDialog(launched, sheet)
    await dialog.getByLabel('Header · Center').fill(value)
    await dialog.getByLabel('Footer · Right').fill('OFFSET&P-1 END')
    await dialog.getByRole('tab', { name: 'First page' }).click()
    await dialog.getByLabel('Different first page').check()
    await dialog.getByLabel('Header · Center').fill(`FIRST ${value}`)
    await dialog.getByLabel('Footer · Right').fill('OFFSET&P-1 END')
    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await dialog.getByLabel('Different odd and even pages').check()
    await dialog.getByLabel('Header · Center').fill(`EVEN ${value}`)
    await dialog.getByLabel('Footer · Right').fill('OFFSET&P-1 END')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()
  }
  const assertDraft = async (value: string) => {
    const dialog = await openHeaderFooterDialog(launched, sheet)
    await expect(dialog.getByLabel('Header · Center')).toHaveValue(value)
    await expect(dialog.getByLabel('Footer · Right')).toHaveValue('OFFSET&P-1 END')
    await dialog.getByRole('tab', { name: 'First page' }).click()
    await expect(dialog.getByLabel('Different first page')).toBeChecked()
    await expect(dialog.getByLabel('Header · Center')).toHaveValue(`FIRST ${value}`)
    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await expect(dialog.getByLabel('Different odd and even pages')).toBeChecked()
    await expect(dialog.getByLabel('Header · Center')).toHaveValue(`EVEN ${value}`)
    await dialog.getByRole('button', { name: 'Cancel' }).click()
  }
  const saveCodes = async (value: string) =>
    saveFromApp(launched, workbook, async (zip) => {
      const xml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
      return (
        headerFooterSectionText(xml, 'oddHeader') === `&C${value}` &&
        headerFooterSectionText(xml, 'oddFooter') === '&ROFFSET&P-1 END' &&
        headerFooterSectionText(xml, 'firstHeader') === `&CFIRST ${value}` &&
        headerFooterSectionText(xml, 'evenHeader') === `&CEVEN ${value}`
      )
    })
  try {
    await sheet.evaluate(() => {
      const worksheet = (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
      worksheet
        .getRange('A1:A180')
        .setValues(Array.from({ length: 180 }, (_, index) => [`NATIVE_CODE_ROW_${index + 1}`]))
    })
    await apply(original)
    await saveCodes(original)
    await apply(edited)
    await sheet.getByRole('button', { name: 'Undo', exact: true }).click()
    await assertDraft(original)
    await sheet.getByRole('button', { name: 'Redo', exact: true }).click()
    await assertDraft(edited)
    await saveCodes(edited)
    await sheet.getByRole('button', { name: 'Undo', exact: true }).click()
    await assertDraft(original)
    await saveCodes(original)
    await sheet.getByRole('button', { name: 'Redo', exact: true }).click()
    await assertDraft(edited)
    await saveCodes(edited)
    await attachScreenshot(sheet, 'sheets-hf-native-codes-after-redo.png')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-hf-native-codes')
  }
  const reopened = await openExisting(workbook, 'sheets-hf-native-codes-reopened')
  try {
    const dialog = await openHeaderFooterDialog(reopened.launched, reopened.sheet)
    await expect(dialog.getByLabel('Header · Center')).toHaveValue(edited)
    await expect(dialog.getByLabel('Footer · Right')).toHaveValue('OFFSET&P-1 END')
    await dialog.getByRole('tab', { name: 'First page' }).click()
    await expect(dialog.getByLabel('Different first page')).toBeChecked()
    await expect(dialog.getByLabel('Header · Center')).toHaveValue(`FIRST ${edited}`)
    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await expect(dialog.getByLabel('Different odd and even pages')).toBeChecked()
    await expect(dialog.getByLabel('Header · Center')).toHaveValue(`EVEN ${edited}`)
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    const pdf = join(dirname(workbook), 'native-codes.pdf')
    await reopened.launched.app.evaluate(({ dialog: nativeDialog }, path) => {
      nativeDialog.showSaveDialog = (async () => ({ canceled: false, filePath: path })) as never
    }, pdf)
    await clickAppAction(reopened.launched, 'export-pdf')
    await expect
      .poll(async () => (await readFile(pdf).catch(() => Buffer.alloc(0))).length, {
        timeout: 30_000,
      })
      .toBeGreaterThan(1_000)
    const pages = await exportedPageTexts(pdf)
    expect(pages.length).toBeGreaterThanOrEqual(3)
    for (const [index, text] of pages.entries()) {
      expect(text).toContain(`SHIFT${index + 3}/${pages.length}`)
      expect(text).toContain(`OFFSET${index} END`)
      expect(text).toContain(basename(workbook))
      expect(text).toContain('Sheet1')
      expect(text).not.toContain('&K04+050')
      expect(text).not.toContain('&P')
    }
    await test.info().attach('sheets-hf-native-code-page-text.json', {
      body: JSON.stringify(pages, null, 2),
      contentType: 'application/json',
    })
    await attachScreenshot(reopened.sheet, 'sheets-hf-native-codes-reopened.png')
  } finally {
    await closeAndSaveVideo(reopened.launched, 'sheets-hf-native-codes-reopened')
  }
})

test('Data Validation: an unrepresentable UI label gives a retryable save error', async () => {
  const { launched, sheet, workbook } = await openBlank('sheets-office-mac-verify-dv-error')
  try {
    const original = await readFile(workbook)
    const liveRules = await createValidationRule(sheet, ['Red, blue', 'Green'])
    expect(liveRules).toHaveLength(1)
    await clickAppAction(launched, 'save')
    const status = sheet.locator('.status-msg')
    await expect(status).toContainText('cell-range source', { timeout: 20_000 })
    expect(await readFile(workbook)).toEqual(original)

    // A second attempt reaching the same gateway validation proves the failed
    // save did not drain or replace the live, still-dirty rule state.
    await clickAppAction(launched, 'save')
    await expect(status).toContainText('cell-range source', { timeout: 20_000 })
    const remaining = await sheet.evaluate(() =>
      (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getDataValidations()
        .map(({ rule }: { rule: Record<string, unknown> }) => ({
          formula1: rule.formula1,
          formula2: rule.formula2,
        })),
    )
    expect(remaining).toEqual([
      {
        formula1: JSON.stringify(['Red, blue', 'Green']),
        formula2: ',',
      },
    ])
  } finally {
    await closeAndSaveVideo(launched, 'sheets-office-mac-verify-dv-error')
  }
})
