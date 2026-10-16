import { expect, test, type Page } from '@playwright/test'
import { XMLParser } from 'fast-xml-parser'
import JSZip from 'jszip'
import { copyFile, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(
  !process.env.GENOFFICE_SHEETS_TEXT_TO_COLUMNS_VERIFY,
  'Opt-in Text to Columns verification',
)

const nativeBefore = join(
  process.cwd(),
  'docs/qa/office-mac-field-evidence-2026-10-13/sheets-text-to-columns/excel-mac-text-to-columns-before.xlsx',
)

type Launched = Awaited<ReturnType<typeof launchShell>>

const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
})

function asArray<T>(value: T | T[] | undefined): T[] {
  return value === undefined ? [] : Array.isArray(value) ? value : [value]
}

function xmlText(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'object' && '#text' in value)
    return xmlText((value as { '#text': unknown })['#text'])
  return String(value)
}

async function openWorkbook(
  path: string,
  videoDir: string,
): Promise<{ launched: Launched; page: Page }> {
  const launched = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
  return { launched, page }
}

async function nativeState(
  path: string,
): Promise<Record<string, { value: string | null; type: string | null; formula: string | null }>> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const strings = (await zip.file('xl/sharedStrings.xml')?.async('string')) ?? ''
  const sheetData = xmlParser.parse(sheet).worksheet?.sheetData
  const cells = asArray<any>(sheetData?.row).flatMap((row) => asArray<any>(row.c))
  const shared = asArray<any>(xmlParser.parse(strings).sst?.si).map((item) =>
    item.t !== undefined
      ? xmlText(item.t)
      : asArray<any>(item.r)
          .map((run) => xmlText(run.t))
          .join(''),
  )
  return Object.fromEntries(
    ['A1', 'A2', 'B1', 'D1', 'E1', 'F1', 'D2', 'E2', 'F2'].map((address) => {
      const cell = cells.find((candidate) => candidate['@_r'] === address)
      const type = cell?.['@_t'] ?? null
      const raw = cell?.v !== undefined ? xmlText(cell.v) : null
      const inline = cell?.is?.t !== undefined ? xmlText(cell.is.t) : null
      const value = type === 's' && raw != null ? (shared[Number(raw)] ?? null) : (inline ?? raw)
      return [address, { value, type, formula: cell?.f !== undefined ? xmlText(cell.f) : null }]
    }),
  )
}

async function save(launched: Launched): Promise<void> {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((part) => part.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
}

async function grid(page: Page, address: string): Promise<unknown[][]> {
  return page.evaluate(
    (target) =>
      (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange(target)
        .getValues(),
    address,
  )
}

async function select(page: Page, address: string): Promise<void> {
  await page.evaluate((target) => {
    const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
    workbook.getActiveSheet().getRange(target).activate()
  }, address)
  await expect
    .poll(() =>
      page.evaluate(() => {
        const range = (window as any).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveRange()
        return range ? `${range.getRow()}:${range.getHeight()}:${range.getWidth()}` : null
      }),
    )
    .toBe('0:2:1')
}

async function finishStep(page: Page): Promise<void> {
  await page
    .getByRole('dialog', { name: 'Text to Columns' })
    .getByRole('button', { name: 'Next' })
    .click()
}

test('native Excel quoted CSV splits into empty destination, blocks formula overwrite, and survives history and reopen', async () => {
  test.setTimeout(180_000)
  const directory = await mkdtemp(join(tmpdir(), 'genoffice-text-to-columns-'))
  const path = join(directory, 'ExcelTextToColumns.xlsx')
  await copyFile(nativeBefore, path)
  const before = await nativeState(path)
  expect(before.A1?.value).toBe('"A, B",003,4')
  expect(before.B1?.formula).toBe('10+5')
  await test.info().attach('text-to-columns-workbook.json', {
    body: JSON.stringify({ path, source: nativeBefore }, null, 2),
    contentType: 'application/json',
  })

  const { launched, page } = await openWorkbook(path, 'sheets-text-to-columns')
  try {
    await expect.poll(() => grid(page, 'A1:A2')).toEqual([['"A, B",003,4'], ['"C,D",005,6']])
    await select(page, 'A1:A2')
    await page
      .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', { hasText: /^Data$/ })
      .first()
      .click()
    await page.locator('[data-ribbon-body] [data-tip="Text to Columns"]').first().click()
    const dialog = page.getByRole('dialog', { name: 'Text to Columns' })
    await expect(dialog).toBeVisible()
    await expect(dialog.getByRole('radio', { name: /Fixed width/ })).toBeDisabled()
    await finishStep(page)
    await expect(dialog.locator('.text-columns-preview tbody tr').first().locator('td')).toHaveText(
      ['A, B', '003', '4'],
    )
    await finishStep(page)
    await expect(dialog.locator('option[value="date"]')).toHaveAttribute('disabled', '')
    await expect(dialog.getByRole('button', { name: /Advanced/ })).toBeDisabled()
    await dialog.getByRole('button', { name: 'Column data format 2' }).click()
    await dialog.getByRole('combobox', { name: 'Column data format' }).selectOption('text')
    await dialog.getByRole('button', { name: 'Finish' }).click()
    await expect(dialog.getByRole('alert')).toContainText('destination contains data or formulas')
    expect((await nativeState(path)).B1?.formula).toBe('10+5')
    await dialog.getByRole('textbox', { name: 'Destination cell' }).fill('D1')
    await dialog.getByRole('button', { name: 'Finish' }).click()
    await expect(dialog).toBeHidden()
    await expect
      .poll(() => grid(page, 'D1:F2'))
      .toEqual([
        ['A, B', '003', 4],
        ['C,D', '005', 6],
      ])
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).E1?.value).toBe('003')
    let saved = await nativeState(path)
    expect(saved.B1?.formula).toBe('10+5')
    expect(saved.B1?.value).toBe('15')
    expect(saved.E1?.type).not.toBeNull()
    expect(saved.D2?.value).toBe('C,D')
    expect(saved.F2?.value).toBe('6')

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).E1?.value).toBeNull()
    saved = await nativeState(path)
    expect(saved.A1?.value).toBe(before.A1?.value)
    expect(saved.B1).toEqual(before.B1)
    for (const address of ['D1', 'E1', 'F1', 'D2', 'E2', 'F2']) {
      expect(saved[address]?.value, `${address} must be cleared by one Undo`).toBeNull()
    }
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).E1?.value).toBe('003')
    expect((await nativeState(path)).B1).toEqual(before.B1)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-text-to-columns')
  }

  const reopened = await openWorkbook(path, 'sheets-text-to-columns-reopen')
  try {
    await expect
      .poll(() => grid(reopened.page, 'D1:F2'))
      .toEqual([
        ['A, B', '003', 4],
        ['C,D', '005', 6],
      ])
    expect((await nativeState(path)).B1).toEqual(before.B1)
    await select(reopened.page, 'A1:A2')
    await reopened.page
      .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', { hasText: /^Data$/ })
      .first()
      .click()
    await reopened.page.locator('[data-ribbon-body] [data-tip="Text to Columns"]').first().click()
    await expect(reopened.page.getByRole('dialog', { name: 'Text to Columns' })).toBeVisible()
    await reopened.page
      .getByRole('dialog', { name: 'Text to Columns' })
      .getByRole('button', { name: 'Cancel' })
      .click()
  } finally {
    await closeAndSaveVideo(reopened.launched, 'sheets-text-to-columns-reopen')
  }
})

test('fixed-width ruler boundaries split text, move and remove, then persist through history and reopen', async () => {
  test.setTimeout(180_000)
  const directory = await mkdtemp(join(tmpdir(), 'genoffice-text-to-columns-fixed-'))
  const path = join(directory, 'FixedWidth.xlsx')
  const zip = await JSZip.loadAsync(await readFile(nativeBefore))
  const strings = await zip.file('xl/sharedStrings.xml')!.async('string')
  if (!strings.includes('"A, B",003,4') || !strings.includes('"C,D",005,6'))
    throw new Error('The native Excel Text to Columns source fixture changed.')
  zip.file(
    'xl/sharedStrings.xml',
    strings.replace('"A, B",003,4', 'ABC0036').replace('"C,D",005,6', 'XYZ0057'),
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
  await test.info().attach('fixed-width-workbook.json', {
    body: JSON.stringify({ path, source: nativeBefore }, null, 2),
    contentType: 'application/json',
  })

  const { launched, page } = await openWorkbook(path, 'sheets-text-to-columns-fixed')
  try {
    await expect.poll(() => grid(page, 'A1:A2')).toEqual([['ABC0036'], ['XYZ0057']])
    await select(page, 'A1:A2')
    await page
      .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', { hasText: /^Data$/ })
      .first()
      .click()
    await page.locator('[data-ribbon-body] [data-tip="Text to Columns"]').first().click()
    const dialog = page.getByRole('dialog', { name: 'Text to Columns' })
    await dialog.getByRole('radio', { name: 'Fixed width' }).check()
    await finishStep(page)
    const at = (position: number) =>
      dialog.locator(`.text-columns-fixed-ruler button[data-position="${position}"]`)
    await expect(dialog.getByRole('alert')).toContainText('column break')
    await at(3).click()
    await at(5).click()
    await expect(dialog.locator('.text-columns-preview tbody tr').first().locator('td')).toHaveText([
      'ABC', '00', '36',
    ])
    await at(5).click()
    await at(6).click()
    await expect(dialog.locator('.text-columns-preview tbody tr').first().locator('td')).toHaveText([
      'ABC', '003', '6',
    ])
    await at(6).dragTo(at(5))
    await expect(at(5)).toHaveAttribute('aria-pressed', 'true')
    await expect(dialog.locator('.text-columns-preview tbody tr').first().locator('td')).toHaveText([
      'ABC', '00', '36',
    ])
    await at(5).dragTo(at(6))
    await expect(at(6)).toHaveAttribute('aria-pressed', 'true')
    await expect(dialog.locator('.text-columns-preview tbody tr').first().locator('td')).toHaveText([
      'ABC', '003', '6',
    ])
    await finishStep(page)
    await dialog.getByRole('button', { name: 'Column data format 2' }).click()
    await dialog.getByRole('combobox', { name: 'Column data format' }).selectOption('text')
    await dialog.getByRole('textbox', { name: 'Destination cell' }).fill('D1')
    await dialog.getByRole('button', { name: 'Finish' }).click()
    await expect(dialog).toBeHidden()
    await expect.poll(() => grid(page, 'D1:F2')).toEqual([
      ['ABC', '003', 6],
      ['XYZ', '005', 7],
    ])
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).E1?.value).toBe('003')
    expect((await nativeState(path)).B1?.formula).toBe('10+5')

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).E1?.value).toBeNull()
    const undone = await nativeState(path)
    for (const address of ['D1', 'E1', 'F1', 'D2', 'E2', 'F2'])
      expect(undone[address]?.value, `${address} should be cleared by one Undo`).toBeNull()
    expect(undone.B1?.formula).toBe('10+5')

    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).E1?.value).toBe('003')
    expect((await nativeState(path)).E2?.value).toBe('005')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-text-to-columns-fixed')
  }

  const reopened = await openWorkbook(path, 'sheets-text-to-columns-fixed-reopen')
  try {
    await expect.poll(() => grid(reopened.page, 'D1:F2')).toEqual([
      ['ABC', '003', 6],
      ['XYZ', '005', 7],
    ])
    expect((await nativeState(path)).B1?.formula).toBe('10+5')
  } finally {
    await closeAndSaveVideo(reopened.launched, 'sheets-text-to-columns-fixed-reopen')
  }
})
