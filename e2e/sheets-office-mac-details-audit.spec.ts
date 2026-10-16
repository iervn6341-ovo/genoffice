import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtemp, readdir, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

const RIBBON_TAB = 'nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)'

type Launched = Awaited<ReturnType<typeof launchShell>>
type SavedPackagePredicate = (zip: JSZip) => Promise<boolean>
interface FailureCapture {
  readonly screenshot?: Buffer
  readonly state: { readonly url: string; readonly title: string; readonly bodyText: string }
}

test.skip(!process.env.GENOFFICE_OFFICE_AUDIT, 'Opt-in Office detail audit')
test.use({ actionTimeout: 10_000 })

let lastFailureCapture: FailureCapture | null = null

test.beforeEach(() => {
  lastFailureCapture = null
})

test.afterEach(async ({ browserName }, testInfo) => {
  void browserName
  const capture = lastFailureCapture
  lastFailureCapture = null
  if (testInfo.status === testInfo.expectedStatus || !capture) return
  if (capture.screenshot) {
    await testInfo.attach('sheets-office-audit-failure.png', {
      body: capture.screenshot,
      contentType: 'image/png',
    })
  }
  await testInfo.attach('sheets-office-audit-failure.json', {
    body: JSON.stringify(
      {
        title: testInfo.title,
        status: testInfo.status,
        expectedStatus: testInfo.expectedStatus,
        error: testInfo.error?.message ?? null,
        ...capture.state,
      },
      null,
      2,
    ),
    contentType: 'application/json',
  })
})

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

async function widenScratchWindow(launched: Launched, sheet: Page): Promise<void> {
  const contentSize = await launched.app.evaluate(({ BrowserWindow }) => {
    const win = BrowserWindow.getAllWindows()[0]
    win?.setContentSize(2200, 1100)
    return win?.getContentSize() ?? null
  })
  const rendererSize = await sheet.evaluate(() => ({ width: innerWidth, height: innerHeight }))
  test.info().annotations.push({
    type: 'sheets-audit-window-size',
    description: `content=${contentSize?.join('x') ?? 'unknown'}; renderer=${rendererSize.width}x${rendererSize.height}`,
  })
  await sheet.waitForTimeout(300)
}

async function openBlank(): Promise<{
  launched: Launched
  sheet: Page
  workbook: string
}> {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-details-audit-'))
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-office-mac-details-audit',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.app.evaluate(({ app }, path) => app.setPath('documents', path), scratch)
  await launched.page.locator('.quick-card').nth(1).click()
  const sheet = await waitForSheetsPage(launched)
  await widenScratchWindow(launched, sheet)
  await sheet.waitForTimeout(1_000)
  const files = await readdir(join(scratch, 'GenOffice'))
  const workbooks = files.filter((file) => file.endsWith('.xlsx'))
  expect(workbooks).toHaveLength(1)
  const workbook = join(scratch, 'GenOffice', workbooks[0]!)
  await test.info().attach('sheets-audit-scratch-workbook.json', {
    body: JSON.stringify({ test: test.info().title, workbook }, null, 2),
    contentType: 'application/json',
  })
  return { launched, sheet, workbook }
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
  const sheet = await waitForSheetsPage(launched)
  await widenScratchWindow(launched, sheet)
  return { launched, sheet }
}

const selectTab = (page: Page, name: string): Promise<void> =>
  page
    .locator(RIBBON_TAB, { hasText: new RegExp(`^${name}$`) })
    .first()
    .click()

const setCell = (page: Page, address: string, value: string): Promise<void> =>
  page.evaluate(
    ({ address, value }) => {
      const api = (window as any).__genofficeDebug.univerAPI
      const sheet = api.getActiveWorkbook().getActiveSheet()
      const range = sheet.getRange(address)
      range.setValue(value)
      range.activate()
    },
    { address, value },
  )

const readValues = (page: Page, address: string): Promise<unknown[][]> =>
  page.evaluate(
    ({ address }) =>
      (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(address)
        .getValues(),
    { address },
  )

async function packageXml(zip: JSZip, path: string): Promise<string> {
  return (await zip.file(path)?.async('string')) ?? ''
}

async function packageXmlContains(zip: JSZip, fragment: string): Promise<boolean> {
  const files = Object.values(zip.files).filter(
    (file) => !file.dir && /\.(?:xml|vml)$/i.test(file.name),
  )
  const texts = await Promise.all(files.map((file) => file.async('string')))
  return texts.some((text) => text.includes(fragment))
}

function xmlElementAttributes(xml: string, name: string): Record<string, string> | null {
  const openingTag = new RegExp(`<${name}\\b([^>]*)>`).exec(xml)?.[1]
  if (openingTag === undefined) return null
  return Object.fromEntries(
    [...openingTag.matchAll(/([\w:.-]+)="([^"]*)"/g)].map((match) => [match[1]!, match[2]!]),
  )
}

function xmlElementHasAttributes(
  xml: string,
  name: string,
  expected: Readonly<Record<string, string>>,
): boolean {
  const attributes = xmlElementAttributes(xml, name)
  return (
    attributes !== null &&
    Object.entries(expected).every(([key, value]) => attributes[key] === value)
  )
}

async function cellValue(zip: JSZip, address: string): Promise<string | number | undefined> {
  const sheetXml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
  const escapedAddress = address.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const cell = new RegExp(`<c\\b(?=[^>]*\\br="${escapedAddress}")([^>]*)>([\\s\\S]*?)<\\/c>`).exec(
    sheetXml,
  )
  if (!cell) return undefined
  const attributes = cell[1] ?? ''
  const content = cell[2] ?? ''
  const value = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(content)?.[1]
  if (/\bt="inlineStr"/.test(attributes)) {
    return decodeXml(
      [...content.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
        .map((match) => match[1] ?? '')
        .join(''),
    )
  }
  if (value === undefined) return undefined
  if (/\bt="s"/.test(attributes)) {
    const stringsXml = await packageXml(zip, 'xl/sharedStrings.xml')
    const items = [...stringsXml.matchAll(/<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g)]
    const item = items[Number(value)]?.[1] ?? ''
    const fragments = [...item.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
      .map((match) => match[1] ?? '')
      .join('')
    return decodeXml(fragments)
  }
  const numeric = Number(value)
  return Number.isFinite(numeric) ? numeric : decodeXml(value)
}

function decodeXml(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

async function saveFromApp(
  launched: Launched,
  workbook: string,
  expected: SavedPackagePredicate,
): Promise<JSZip> {
  await launched.app.evaluate(({ webContents }) => {
    const wc = webContents.getAllWebContents().find((item) => item.getURL().includes('://sheets/'))
    wc?.send('menu:action', 'save')
  })
  // Do not accept a pre-existing workbook as a successful save. Each caller
  // waits for its own changed value/rule/part to appear in the saved package.
  await expect
    .poll(async () => expected(await JSZip.loadAsync(await readFile(workbook))), {
      timeout: 20_000,
    })
    .toBe(true)
  return JSZip.loadAsync(await readFile(workbook))
}

async function close(launched: Launched, videoName: string): Promise<void> {
  const page = launched.app
    .context()
    .pages()
    .find((candidate) => candidate.url().includes('://sheets/'))
  if (page) {
    const screenshot = await page.screenshot({ timeout: 5_000 }).catch(() => undefined)
    const state = {
      url: page.url(),
      title: await page.title().catch(() => ''),
      bodyText: (
        await page
          .locator('body')
          .innerText({ timeout: 2_000 })
          .catch(() => '')
      ).slice(0, 4_000),
    }
    lastFailureCapture = { ...(screenshot ? { screenshot } : {}), state }
  }
  await closeAndSaveVideo(launched, videoName)
}

test('Format Cells: inspect current menu options and persist the selected protection value', async () => {
  const { launched, sheet, workbook } = await openBlank()
  try {
    await setCell(sheet, 'A1', 'SHEETS_AUDIT_FORMAT')
    await selectTab(sheet, 'Home')
    await sheet
      .locator('[data-ribbon-body] [data-tip="Format Cells"]')
      .first()
      .click({ timeout: 10_000 })
    const format = sheet.locator('.format-cells-dialog[role="dialog"]')
    await expect(format).toBeVisible()
    await expect(format.locator('.dialog-tabs button')).toHaveText([
      'Number',
      'Alignment',
      'Font',
      'Border',
      'Fill',
      'Protection',
    ])

    const numberOptions = await format.locator('.numfmt-cats [role="option"]').allInnerTexts()
    expect(numberOptions).not.toContain('Special')

    await format.locator('.dialog-tabs button', { hasText: 'Alignment' }).click()
    await format.getByRole('button', { name: 'Horizontal' }).click()
    await expect(format.locator('.gs-dd-pop [role="option"]')).toHaveText([
      '(unchanged)',
      'Left',
      'Center',
      'Right',
      'Justify',
      'Distributed',
    ])
    await sheet.keyboard.press('Escape')
    await format.getByRole('button', { name: 'Vertical' }).click()
    await expect(format.locator('.gs-dd-pop [role="option"]')).toHaveText([
      '(unchanged)',
      'Top',
      'Center',
      'Bottom',
    ])
    await sheet.keyboard.press('Escape')
    expect(await format.innerText()).not.toMatch(/Merge cells|Shrink to fit/i)

    await format.locator('.dialog-tabs button', { hasText: 'Protection' }).click()
    const locked = format.getByRole('button', { name: 'Locked' })
    const hidden = format.getByRole('button', { name: 'Hidden (formula)' })
    // This records the current echo behavior: both protection fields start in
    // the unchanged state rather than reflecting the selected cell's flags.
    await expect(locked).toHaveAttribute('data-value', '')
    await expect(hidden).toHaveAttribute('data-value', '')
    await locked.click()
    await format.getByRole('option', { name: 'Off' }).click()
    await format.getByRole('button', { name: 'OK' }).click()

    const zip = await saveFromApp(launched, workbook, (candidate) =>
      packageXmlContains(candidate, 'locked="0"'),
    )
    expect(await cellValue(zip, 'A1')).toBe('SHEETS_AUDIT_FORMAT')
    expect(await packageXmlContains(zip, 'locked="0"')).toBe(true)
  } finally {
    await close(launched, 'sheets-office-mac-details-audit-format')
  }

  const reopened = await openExisting(workbook, 'sheets-office-mac-details-audit-format-reopen')
  try {
    await expect.poll(() => readValues(reopened.sheet, 'A1')).toEqual([['SHEETS_AUDIT_FORMAT']])
    const zip = await JSZip.loadAsync(await readFile(workbook))
    expect(await packageXmlContains(zip, 'locked="0"')).toBe(true)
  } finally {
    await close(reopened.launched, 'sheets-office-mac-details-audit-format-reopen')
  }
})

test('Text to Columns: delimited wizard splits cell values and survives workbook reopen', async () => {
  const { launched, sheet, workbook } = await openBlank()
  try {
    await setCell(sheet, 'A1', 'SHEETS_AUDIT_LEFT,SHEETS_AUDIT_RIGHT')
    await selectTab(sheet, 'Data')
    await sheet.locator('[data-ribbon-body] [data-tip="Text to Columns"]').first().click()
    const wizard = sheet.getByRole('dialog', { name: 'Text to Columns' })
    await expect(wizard).toBeVisible()
    await wizard.getByRole('button', { name: 'Next' }).click()
    await expect(wizard.getByRole('checkbox', { name: 'Split by Comma' })).toBeChecked()
    await wizard.getByRole('button', { name: 'Next' }).click()
    await wizard.getByRole('button', { name: 'Finish' }).click()
    await expect
      .poll(() => readValues(sheet, 'A1:B1'))
      .toEqual([['SHEETS_AUDIT_LEFT', 'SHEETS_AUDIT_RIGHT']])

    const zip = await saveFromApp(
      launched,
      workbook,
      async (candidate) =>
        (await cellValue(candidate, 'A1')) === 'SHEETS_AUDIT_LEFT' &&
        (await cellValue(candidate, 'B1')) === 'SHEETS_AUDIT_RIGHT',
    )
    expect(await cellValue(zip, 'A1')).toBe('SHEETS_AUDIT_LEFT')
    expect(await cellValue(zip, 'B1')).toBe('SHEETS_AUDIT_RIGHT')
  } finally {
    await close(launched, 'sheets-office-mac-details-audit-text-to-columns')
  }

  const reopened = await openExisting(
    workbook,
    'sheets-office-mac-details-audit-text-to-columns-reopen',
  )
  try {
    await expect
      .poll(() => readValues(reopened.sheet, 'A1:B1'))
      .toEqual([['SHEETS_AUDIT_LEFT', 'SHEETS_AUDIT_RIGHT']])
  } finally {
    await close(reopened.launched, 'sheets-office-mac-details-audit-text-to-columns-reopen')
  }
})

test('Symbol and Equation: insert a cell character plus a raster picture, then reopen', async () => {
  const { launched, sheet, workbook } = await openBlank()
  try {
    await setCell(sheet, 'A1', 'SHEETS_AUDIT_SYMBOL')
    await selectTab(sheet, 'Insert')
    await sheet.locator('[data-ribbon-body] [data-tip="Symbol"]').first().click({ timeout: 10_000 })
    const symbolDialog = sheet.getByRole('dialog', { name: 'Symbol' })
    await expect(symbolDialog).toBeVisible()
    await expect(symbolDialog.locator('.dialog-tabs button')).toHaveText([
      'Currency & Common',
      'Math',
      'Arrows',
      'Greek',
      'Shapes & Misc',
    ])
    await symbolDialog.getByRole('button', { name: 'Greek', exact: true }).click()
    await test.info().attach('sheets-symbol-greek-dialog.png', {
      body: await sheet.screenshot(),
      contentType: 'image/png',
    })
    await symbolDialog.getByRole('button', { name: 'Ω', exact: true }).click()
    await expect.poll(() => readValues(sheet, 'A1')).toEqual([['SHEETS_AUDIT_SYMBOLΩ']])
    await symbolDialog.getByRole('button', { name: 'Close' }).click()

    await sheet.locator('[data-ribbon-body] [data-tip="Equation"]').first().click()
    const equation = sheet.locator('[role="dialog"]').filter({ hasText: 'Insert Equation' })
    await expect(equation).toBeVisible()
    await expect(equation.locator('.equation-presets button')).toHaveCount(6)
    await equation.locator('.equation-latex').fill('x = \\frac{a}{b}')
    await test.info().attach('sheets-equation-editor.png', {
      body: await sheet.screenshot(),
      contentType: 'image/png',
    })
    await equation.getByRole('button', { name: 'OK' }).click()
    await expect(equation).toBeHidden()

    const zip = await saveFromApp(launched, workbook, async (candidate) => {
      const hasSymbol = (await cellValue(candidate, 'A1')) === 'SHEETS_AUDIT_SYMBOLΩ'
      const hasDrawing = (await packageXml(candidate, 'xl/worksheets/sheet1.xml')).includes(
        '<drawing',
      )
      const hasMedia = Object.keys(candidate.files).some((name) => /^xl\/media\//.test(name))
      return hasSymbol && hasDrawing && hasMedia
    })
    expect(await cellValue(zip, 'A1')).toBe('SHEETS_AUDIT_SYMBOLΩ')
    expect(await packageXml(zip, 'xl/worksheets/sheet1.xml')).toContain('<drawing')
    expect(Object.keys(zip.files).some((name) => /^xl\/media\//.test(name))).toBe(true)
  } finally {
    await close(launched, 'sheets-office-mac-details-audit-symbol-equation')
  }

  const reopened = await openExisting(
    workbook,
    'sheets-office-mac-details-audit-symbol-equation-reopen',
  )
  try {
    await expect.poll(() => readValues(reopened.sheet, 'A1')).toEqual([['SHEETS_AUDIT_SYMBOLΩ']])
    // The exporter names this scratch workbook's sole drawing "Picture 2";
    // imported visual names are not the original media filename Equation.png.
    const equationImage = reopened.sheet.locator('.xlsx-image')
    await expect(equationImage).toHaveCount(1, { timeout: 15_000 })
    await expect(equationImage).toBeVisible({ timeout: 15_000 })
    await expect
      .poll(() => equationImage.evaluate((image) => (image as HTMLImageElement).naturalWidth))
      .toBeGreaterThan(0)
    await test.info().attach('sheets-equation-reopened.png', {
      body: await reopened.sheet.screenshot(),
      contentType: 'image/png',
    })
    const zip = await JSZip.loadAsync(await readFile(workbook))
    expect(Object.keys(zip.files).some((name) => /^xl\/media\//.test(name))).toBe(true)
    expect(await packageXml(zip, 'xl/worksheets/sheet1.xml')).toContain('<drawing')
  } finally {
    await close(reopened.launched, 'sheets-office-mac-details-audit-symbol-equation-reopen')
  }
})

test('New Comment opens and saves a note editor, not a threaded comment composer', async () => {
  const { launched, sheet, workbook } = await openBlank()
  try {
    await selectTab(sheet, 'Review')
    await sheet.locator('[data-ribbon-body] [data-tip="New Comment"]').first().click()
    const editor = sheet.locator('[data-u-comp="note-textarea"]')
    await expect(editor).toBeVisible()
    await editor.fill('SHEETS_AUDIT_NOTE_FROM_UI')
    const liveNotes = () =>
      sheet.evaluate(() =>
        (window as any).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveSheet()
          .getNotes()
          .map((note: { note: string }) => note.note),
      )
    // The textarea debounces commits to the live note model. Wait for that
    // commit before dismissing the popup and asking the app to save.
    await expect.poll(liveNotes).toContain('SHEETS_AUDIT_NOTE_FROM_UI')
    const a2 = await sheet.evaluate(() => {
      const canvas = [...document.querySelectorAll('canvas')].find((candidate) => {
        const rect = candidate.getBoundingClientRect()
        return rect.width > 500 && rect.height > 300
      })
      if (!canvas) return null
      const rect = canvas.getBoundingClientRect()
      return { x: rect.x + 46 + 43, y: rect.y + 24 + 12 + 20 }
    })
    if (!a2) throw new Error('worksheet canvas not found while committing the note popup')
    await sheet.mouse.click(a2.x, a2.y)
    await expect(editor).not.toBeVisible()
    await expect.poll(liveNotes).toContain('SHEETS_AUDIT_NOTE_FROM_UI')

    const zip = await saveFromApp(launched, workbook, async (candidate) =>
      packageXmlContains(candidate, 'SHEETS_AUDIT_NOTE_FROM_UI'),
    )
    const comments = await packageXml(zip, 'xl/comments1.xml')
    expect(comments).toContain('SHEETS_AUDIT_NOTE_FROM_UI')
    expect(comments).toContain('<comment ref="A1"')
    expect(
      Object.keys(zip.files).some((name) => /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(name)),
    ).toBe(true)
    expect(Object.keys(zip.files).some((name) => /^xl\/threadedComments\//.test(name))).toBe(false)
  } finally {
    await close(launched, 'sheets-office-mac-details-audit-comments')
  }

  const reopened = await openExisting(workbook, 'sheets-office-mac-details-audit-comments-reopen')
  try {
    await expect
      .poll(() =>
        reopened.sheet.evaluate(() =>
          (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getNotes()
            .map((note: { note: string }) => note.note),
        ),
      )
      .toContain('SHEETS_AUDIT_NOTE_FROM_UI')
    await selectTab(reopened.sheet, 'Review')
    await reopened.sheet.locator('[data-ribbon-body] [data-tip="New Comment"]').first().click()
    const editor = reopened.sheet.locator('[data-u-comp="note-textarea"]')
    await expect(editor).toBeVisible()
    await expect(editor).toHaveValue('SHEETS_AUDIT_NOTE_FROM_UI')
  } finally {
    await close(reopened.launched, 'sheets-office-mac-details-audit-comments-reopen')
  }
})

test('Data Validation: record the Univer list model and exported OOXML (engine probe)', async () => {
  const { launched, sheet, workbook } = await openBlank()
  try {
    await selectTab(sheet, 'Data')
    await sheet.locator('[data-ribbon-body] [data-tip="Data Validation"]').first().click()
    await expect(sheet.locator('body')).toContainText('Data validation')
    await expect(sheet.getByRole('button', { name: 'Add Rule' })).toBeVisible()

    // The current E2E driver seeds Univer's public rule model after opening
    // the real UI panel. This captures model/export behavior, not authoring
    // through the panel controls.
    await sheet.evaluate(() => {
      const api = (window as any).__genofficeDebug.univerAPI
      const activeSheet = api.getActiveWorkbook().getActiveSheet()
      const rule = api
        .newDataValidation()
        .requireValueInList(['Open', 'Closed'], false, true)
        .setOptions({
          allowBlank: true,
          showInputMessage: true,
          prompt: 'Choose a status',
          promptTitle: 'Status',
          showErrorMessage: true,
          error: 'Choose Open or Closed',
          errorTitle: 'Invalid status',
        })
        .build()
      activeSheet.getRange('B2:B4').setDataValidation(rule)
    })

    const liveRules = await sheet.evaluate(() =>
      (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getDataValidations()
        .map(({ rule }: { rule: Record<string, unknown> }) => ({
          type: rule.type,
          formula1: rule.formula1,
          formula2: rule.formula2,
          ranges: rule.ranges,
          allowBlank: rule.allowBlank,
          showInputMessage: rule.showInputMessage,
          promptTitle: rule.promptTitle,
          prompt: rule.prompt,
          showErrorMessage: rule.showErrorMessage,
          errorTitle: rule.errorTitle,
          error: rule.error,
        })),
    )
    expect(liveRules).toHaveLength(1)
    expect(liveRules[0]).toMatchObject({
      type: 'list',
      formula1: JSON.stringify(['Open', 'Closed']),
      showInputMessage: true,
      promptTitle: 'Status',
      showErrorMessage: true,
      errorTitle: 'Invalid status',
    })
    await test.info().attach('sheets-dv-live-rule.json', {
      body: JSON.stringify(liveRules, null, 2),
      contentType: 'application/json',
    })

    const zip = await saveFromApp(launched, workbook, async (candidate) => {
      const xml = await packageXml(candidate, 'xl/worksheets/sheet1.xml')
      const element = /<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/.exec(xml)
      const attributes = element?.[1] ?? ''
      const serializedListLiteral = `"${JSON.stringify(['Open', 'Closed'])}"`
      return (
        attributes.includes('type="list"') &&
        attributes.includes('sqref="B2:B4"') &&
        attributes.includes('showInputMessage="1"') &&
        attributes.includes('showErrorMessage="1"') &&
        attributes.includes('promptTitle="Status"') &&
        attributes.includes('errorTitle="Invalid status"') &&
        element?.[2]?.includes(`<formula1>${serializedListLiteral}</formula1>`) === true
      )
    })
    const sheetXml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
    await test.info().attach('sheets-dv-export-sheet1.xml', {
      body: sheetXml,
      contentType: 'application/xml',
    })
    // Capture current serializer output as the probe result, then separately
    // evaluate the documented OOXML literal-list shape. This is intentionally
    // soft so reopen and live-rule evidence are still collected.
    expect.soft(sheetXml).toContain('<formula1>"Open,Closed"</formula1>')
    const serializedListLiteral = `"${JSON.stringify(['Open', 'Closed'])}"`
    expect(
      xmlElementHasAttributes(sheetXml, 'dataValidation', {
        type: 'list',
        sqref: 'B2:B4',
        showInputMessage: '1',
        showErrorMessage: '1',
        promptTitle: 'Status',
        errorTitle: 'Invalid status',
      }),
    ).toBe(true)
    expect(sheetXml).toContain(`<formula1>${serializedListLiteral}</formula1>`)
  } finally {
    await close(launched, 'sheets-office-mac-details-audit-validation')
  }

  const reopened = await openExisting(workbook, 'sheets-office-mac-details-audit-validation-reopen')
  try {
    const rules = await reopened.sheet.evaluate(() => {
      const api = (window as any).__genofficeDebug.univerAPI
      return api
        .getActiveWorkbook()
        .getActiveSheet()
        .getDataValidations()
        .map(({ rule }: { rule: Record<string, any> }) => ({
          ranges: rule.ranges,
          type: rule.type,
          formula1: rule.formula1,
          showInputMessage: rule.showInputMessage,
          showErrorMessage: rule.showErrorMessage,
          promptTitle: rule.promptTitle,
          errorTitle: rule.errorTitle,
        }))
    })
    expect(rules).toEqual([
      expect.objectContaining({
        ranges: [expect.objectContaining({ startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 })],
        type: 'list',
        formula1: JSON.stringify(['Open', 'Closed']),
        showInputMessage: true,
        showErrorMessage: true,
        promptTitle: 'Status',
        errorTitle: 'Invalid status',
      }),
    ])
  } finally {
    await close(reopened.launched, 'sheets-office-mac-details-audit-validation-reopen')
  }
})

test('Data Validation: author a literal list through the panel, export, and reopen', async () => {
  const { launched, sheet, workbook } = await openBlank()
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
    await expect(sheet.getByText('Dropdown', { exact: true }).last()).toBeVisible()
    await sheet.getByText('Dropdown', { exact: true }).last().click()
    await expect(detail.getByText('Options', { exact: true })).toBeVisible()

    const listValues = () =>
      detail.locator('[data-draggable-list-item-id] [data-u-comp="input"] input')
    await detail.getByText('Add', { exact: true }).click()
    await expect(listValues()).toHaveCount(1)
    await listValues().nth(0).fill('Open')
    await detail.getByText('Add', { exact: true }).click()
    await expect(listValues()).toHaveCount(2)
    await listValues().nth(1).fill('Closed')
    await test.info().attach('sheets-dv-ui-authored-list-editor.png', {
      body: await sheet.screenshot(),
      contentType: 'image/png',
    })

    const liveRules = async () =>
      sheet.evaluate(() =>
        (window as any).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveSheet()
          .getDataValidations()
          .map(({ rule }: { rule: Record<string, unknown> }) => ({
            type: rule.type,
            formula1: rule.formula1,
            ranges: rule.ranges,
          })),
      )
    await expect.poll(liveRules).toEqual([
      expect.objectContaining({
        type: 'list',
        formula1: JSON.stringify(['Open', 'Closed']),
      }),
    ])
    await test.info().attach('sheets-dv-ui-authored-live-rule.json', {
      body: JSON.stringify(await liveRules(), null, 2),
      contentType: 'application/json',
    })

    await detail.getByRole('button', { name: 'Done' }).click()
    await expect(detail).toBeHidden()
    const zip = await saveFromApp(launched, workbook, async (candidate) => {
      const xml = await packageXml(candidate, 'xl/worksheets/sheet1.xml')
      const element = /<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/.exec(xml)
      const attributes = Object.fromEntries(
        [...(element?.[1] ?? '').matchAll(/([\w:.-]+)="([^"]*)"/g)].map((match) => [
          match[1]!,
          match[2]!,
        ]),
      )
      return (
        attributes.type === 'list' &&
        attributes.sqref === 'B2:B4' &&
        /<formula1>/.test(element?.[2] ?? '')
      )
    })
    const sheetXml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
    const validation = /<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/.exec(sheetXml)
    const actualFormula1 = /<formula1>([\s\S]*?)<\/formula1>/.exec(validation?.[2] ?? '')?.[1]
    await test.info().attach('sheets-dv-ui-authored-export-sheet1.xml', {
      body: sheetXml,
      contentType: 'application/xml',
    })
    await test.info().attach('sheets-dv-ui-authored-export-formula1.json', {
      body: JSON.stringify(
        {
          validationAttributes: xmlElementAttributes(sheetXml, 'dataValidation'),
          formula1: actualFormula1 ?? null,
        },
        null,
        2,
      ),
      contentType: 'application/json',
    })
    expect(
      xmlElementHasAttributes(sheetXml, 'dataValidation', {
        type: 'list',
        sqref: 'B2:B4',
      }),
    ).toBe(true)
    expect(actualFormula1).toBeDefined()
    // OOXML literal lists should carry a CSV string literal. Keep this soft so
    // artifacts and close/reopen evidence survive if the current gateway emits
    // the Univer JSON-array model verbatim instead.
    expect.soft(actualFormula1).toBe('"Open,Closed"')
  } finally {
    await close(launched, 'sheets-office-mac-details-audit-validation-ui')
  }

  const reopened = await openExisting(
    workbook,
    'sheets-office-mac-details-audit-validation-ui-reopen',
  )
  try {
    const reopenedRules = await reopened.sheet.evaluate(() =>
      (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getDataValidations()
        .map(({ rule }: { rule: Record<string, unknown> }) => ({
          type: rule.type,
          formula1: rule.formula1,
          ranges: rule.ranges,
        })),
    )
    await test.info().attach('sheets-dv-ui-authored-reopened-rule.json', {
      body: JSON.stringify(reopenedRules, null, 2),
      contentType: 'application/json',
    })
    expect(reopenedRules).toHaveLength(1)
    expect(reopenedRules[0]).toMatchObject({ type: 'list' })
    expect(reopenedRules[0]?.ranges).toEqual([
      expect.objectContaining({ startRow: 1, endRow: 3, startColumn: 1, endColumn: 1 }),
    ])
  } finally {
    await close(reopened.launched, 'sheets-office-mac-details-audit-validation-ui-reopen')
  }
})

test('Print pane persists selected page setup fields into the reopened workbook', async () => {
  const { launched, sheet, workbook } = await openBlank()
  try {
    await launched.app.evaluate(({ Menu, BrowserWindow }) => {
      const walk = (menu: Electron.Menu | null | undefined): Electron.MenuItem | undefined => {
        for (const item of menu?.items ?? []) {
          if (/^(CommandOrControl|CmdOrCtrl|Command|Cmd)\+P$/i.test(String(item.accelerator ?? '')))
            return item
          const found = walk(item.submenu)
          if (found) return found
        }
      }
      walk(Menu.getApplicationMenu())?.click(
        undefined,
        BrowserWindow.getAllWindows()[0],
        undefined as never,
      )
    })
    await sheet.waitForSelector('.print-sheet-dialog', { timeout: 15_000 })
    await sheet.locator('.print-section-head', { hasText: 'Page Setup' }).click()
    await sheet.locator('[data-testid="print-orientation"]').selectOption('landscape')
    await sheet.locator('[data-testid="print-paper"]').selectOption('1')
    await sheet.locator('[data-testid="print-margins"]').selectOption('narrow')
    await sheet.locator('[data-testid="print-scaling"]').selectOption('fitSheet')
    await sheet.locator('.print-section-head', { hasText: 'Options' }).click()
    await sheet
      .locator('.print-section-body .ps-check', { hasText: 'Gridlines' })
      .locator('input')
      .check()
    await sheet.getByRole('button', { name: 'Cancel' }).click()

    const zip = await saveFromApp(launched, workbook, async (candidate) => {
      const xml = await packageXml(candidate, 'xl/worksheets/sheet1.xml')
      return (
        xmlElementHasAttributes(xml, 'pageSetup', { orientation: 'landscape', paperSize: '1' }) &&
        xmlElementHasAttributes(xml, 'pageSetUpPr', { fitToPage: '1' }) &&
        xmlElementHasAttributes(xml, 'printOptions', { gridLines: '1' })
      )
    })
    const sheetXml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
    expect(
      xmlElementHasAttributes(sheetXml, 'pageSetup', { orientation: 'landscape', paperSize: '1' }),
    ).toBe(true)
    expect(xmlElementHasAttributes(sheetXml, 'pageSetUpPr', { fitToPage: '1' })).toBe(true)
    expect(xmlElementHasAttributes(sheetXml, 'printOptions', { gridLines: '1' })).toBe(true)
  } finally {
    await close(launched, 'sheets-office-mac-details-audit-print')
  }

  const reopened = await openExisting(workbook, 'sheets-office-mac-details-audit-print-reopen')
  try {
    await reopened.launched.app.evaluate(({ Menu, BrowserWindow }) => {
      const walk = (menu: Electron.Menu | null | undefined): Electron.MenuItem | undefined => {
        for (const item of menu?.items ?? []) {
          if (/^(CommandOrControl|CmdOrCtrl|Command|Cmd)\+P$/i.test(String(item.accelerator ?? '')))
            return item
          const found = walk(item.submenu)
          if (found) return found
        }
      }
      walk(Menu.getApplicationMenu())?.click(
        undefined,
        BrowserWindow.getAllWindows()[0],
        undefined as never,
      )
    })
    await reopened.sheet.waitForSelector('.print-sheet-dialog', { timeout: 15_000 })
    await reopened.sheet.locator('.print-section-head', { hasText: 'Page Setup' }).click()
    await expect(reopened.sheet.locator('[data-testid="print-orientation"]')).toHaveValue(
      'landscape',
    )
    await expect(reopened.sheet.locator('[data-testid="print-paper"]')).toHaveValue('1')
    await expect(reopened.sheet.locator('[data-testid="print-scaling"]')).toHaveValue('fitSheet')
    const reopenedZip = await JSZip.loadAsync(await readFile(workbook))
    const reopenedXml = await packageXml(reopenedZip, 'xl/worksheets/sheet1.xml')
    expect(
      xmlElementHasAttributes(reopenedXml, 'pageSetup', {
        orientation: 'landscape',
        paperSize: '1',
      }),
    ).toBe(true)
    expect(xmlElementHasAttributes(reopenedXml, 'pageSetUpPr', { fitToPage: '1' })).toBe(true)
    expect(xmlElementHasAttributes(reopenedXml, 'printOptions', { gridLines: '1' })).toBe(true)
  } finally {
    await close(reopened.launched, 'sheets-office-mac-details-audit-print-reopen')
  }
})
