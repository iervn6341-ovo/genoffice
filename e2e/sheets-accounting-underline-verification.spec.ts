import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  TEST_WINDOW,
  waitForPageWithUrl,
} from './helpers'

test.skip(
  !process.env.GENOFFICE_SHEETS_ACCOUNTING_UNDERLINE_VERIFY,
  'Opt-in accounting underline visual verification',
)

const source = join(
  process.cwd(),
  'docs/qa/office-mac-field-evidence-2026-10-13/sheets-accounting-underline/excel-mac-rich-rotated-source.xlsx',
)
const evidence =
  process.env.GENOFFICE_SHEETS_ACCOUNTING_UNDERLINE_DIR ??
  join(process.cwd(), 'docs/qa/sheets-accounting-underline-evidence-2026-10-14')
const addresses = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'A2', 'B2', 'C2', 'D2'] as const

async function fontModes(path: string): Promise<Record<string, string | null>> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const xfs = [
    ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? '').matchAll(
      /<xf\b[^>]*?\/>|<xf\b[^>]*?>[\s\S]*?<\/xf>/g,
    ),
  ].map((match) => match[0])
  const fonts = [
    ...(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(styles)?.[1] ?? '').matchAll(
      /<font\b[^>]*>[\s\S]*?<\/font>/g,
    ),
  ].map((match) => match[0])
  return Object.fromEntries(
    addresses.map((address) => {
      const cell = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>`).exec(sheet)?.[0] ?? ''
      const xf = xfs[Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)] ?? ''
      const font = fonts[Number(/\bfontId="(\d+)"/.exec(xf)?.[1] ?? 0)] ?? ''
      const underline = /<u\b[^>]*\/>/.exec(font)?.[0] ?? null
      return [address, underline]
    }),
  )
}

async function visualWidths(
  launched: Awaited<ReturnType<typeof launchShell>>,
  page: Page,
  filename: string,
): Promise<Record<string, number>> {
  await setEditorLayoutWidth(launched.app, '://sheets/', 1800)
  const zoom = TEST_WINDOW.width / 1800
  const candidates = await page.evaluate(
    ({ zoom, addresses }) => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      const local = addresses.map((address: string) => sheet.getRange(address).getCellRect())
      return [
        { name: 'viewport', x: 0, y: 0 },
        ...[...document.querySelectorAll('canvas')].map((canvas, index) => {
          const rect = canvas.getBoundingClientRect()
          return { name: `canvas-${index}`, x: rect.left, y: rect.top }
        }),
      ].map((origin) => ({
        name: origin.name,
        rects: local.map((rect: any) => ({
          left: (origin.x + rect.left) * zoom,
          right: (origin.x + rect.right) * zoom,
          top: (origin.y + rect.top) * zoom,
          bottom: (origin.y + rect.bottom) * zoom,
        })),
      }))
    },
    { zoom, addresses },
  )
  const buffer = await page.screenshot({ path: join(evidence, filename) })
  const image = PNG.sync.read(buffer)
  const scale = image.width / (await page.evaluate(() => window.innerWidth))
  const measured = candidates.map((candidate) => {
    const widths = candidate.rects.map((rect) => {
      const left = Math.round((rect.left + 4) * scale)
      const right = Math.round((rect.right - 4) * scale)
      const top = Math.round((rect.top + 7) * scale)
      const bottom = Math.round((rect.bottom - 1) * scale)
      let longest = 0
      for (let y = top; y < bottom; y++) {
        let run = 0
        for (let x = left; x < right; x++) {
          const offset = (y * image.width + x) * 4
          const dark =
            image.data[offset]! < 100 &&
            image.data[offset + 1]! < 100 &&
            image.data[offset + 2]! < 100
          run = dark ? run + 1 : 0
          longest = Math.max(longest, run)
        }
      }
      return longest / scale
    })
    return { ...candidate, widths }
  })
  // A real accounting line spans most of B1's cell and its text/numeric
  // partner D1. This selects the actual grid origin, not unrelated canvas ink.
  measured.sort(
    (a, b) =>
      Math.min(b.widths[1] ?? 0, b.widths[3] ?? 0) - Math.min(a.widths[1] ?? 0, a.widths[3] ?? 0),
  )
  await writeFile(
    join(evidence, filename.replace(/\.png$/, '.json')),
    JSON.stringify({ scale, measured }, null, 2),
  )
  return Object.fromEntries(
    addresses.map((address, index) => [address, measured[0]?.widths[index] ?? 0]),
  )
}

async function activate(page: Page, address: string): Promise<void> {
  await page.evaluate((target) => {
    const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
    workbook.getActiveSheet().getRange(target).activate()
  }, address)
}

async function chooseUnderline(page: Page, address: string, label: string): Promise<void> {
  await activate(page, address)
  await page.keyboard.press('Meta+1')
  const dialog = page.getByRole('dialog', { name: 'Format Cells' })
  await expect(dialog).toBeVisible()
  await dialog.locator('.dialog-tabs button', { hasText: 'Font' }).click()
  const trigger = dialog.getByRole('button', { name: 'Underline' })
  await trigger.click()
  await expect(trigger).toHaveAttribute('aria-expanded', 'true')
  await dialog.getByRole('option', { name: label, exact: true }).click()
  await dialog.getByRole('button', { name: 'OK' }).click()
}

async function save(launched: Awaited<ReturnType<typeof launchShell>>): Promise<void> {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((part) => part.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
}

test('Excel Mac accounting lines span the cell and survive Format Cells, history, and reopen', async () => {
  test.setTimeout(180_000)
  await mkdir(evidence, { recursive: true })
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-accounting-underline-'))
  const path = join(scratch, 'AccountingUnderline.xlsx')
  await copyFile(source, path)
  const originalModes = await fontModes(path)
  expect(originalModes).toMatchObject({
    A1: '<u/>',
    B1: '<u val="singleAccounting"/>',
    C1: '<u/>',
    D1: '<u val="singleAccounting"/>',
    A2: '<u val="double"/>',
    B2: '<u val="doubleAccounting"/>',
    C2: '<u val="double"/>',
    D2: '<u val="doubleAccounting"/>',
  })
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-accounting-underline',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    const imported = await visualWidths(launched, page, 'imported.png')
    const documentCells = await page.evaluate(() => {
      const worksheet = (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getSheet()
      return Object.fromEntries(
        ['E1', 'F1'].map((address) => {
          const column = address === 'E1' ? 4 : 5
          const cell = worksheet.getCell(0, column)
          return [
            address,
            {
              marker: cell?.custom?.genofficeAccountingUnderline,
              hasDocument: Boolean(cell?.p),
              style: cell?.s,
              value: cell?.v,
            },
          ]
        }),
      )
    })
    await writeFile(join(evidence, 'document-cells.json'), JSON.stringify(documentCells, null, 2))
    expect(documentCells).toMatchObject({
      E1: { marker: 'singleAccounting', hasDocument: true },
      F1: { marker: 'singleAccounting' },
    })
    expect(imported.B1).toBeGreaterThan(imported.A1 * 1.8)
    expect(imported.D1).toBeGreaterThan(imported.C1 * 1.8)
    expect(imported.B2).toBeGreaterThan(imported.A2 * 1.8)
    expect(imported.D2).toBeGreaterThan(imported.C2 * 1.8)

    for (const address of ['E1', 'F1']) {
      await activate(page, address)
      await page.keyboard.press('Meta+1')
      const dialog = page.getByRole('dialog', { name: 'Format Cells' })
      await expect(dialog).toBeVisible()
      await dialog.locator('.dialog-tabs button', { hasText: 'Font' }).click()
      const underline = dialog.getByRole('button', { name: 'Underline' })
      await expect(underline).toHaveAttribute('data-value', 'singleAccounting')
      await expect(dialog.getByRole('status')).toContainText(
        'Accounting underline cannot be applied to rich text or rotated cells',
      )
      await underline.click()
      await expect(dialog.getByRole('option', { name: 'Single Accounting' })).toBeDisabled()
      await expect(dialog.getByRole('option', { name: 'Double Accounting' })).toBeDisabled()
      await underline.press('Escape')
      await expect(underline).toHaveAttribute('aria-expanded', 'false')
      await dialog.getByRole('button', { name: 'Cancel' }).click()
    }

    await activate(page, 'B1')
    await page.keyboard.press('Meta+1')
    const rotationDialog = page.getByRole('dialog', { name: 'Format Cells' })
    await expect(rotationDialog).toBeVisible()
    await rotationDialog.locator('.dialog-tabs button', { hasText: 'Alignment' }).click()
    await rotationDialog.getByRole('spinbutton', { name: 'Degrees' }).fill('45')
    await rotationDialog.locator('.dialog-tabs button', { hasText: 'Font' }).click()
    await expect(rotationDialog.getByRole('status')).toContainText(
      'Accounting underline cannot be applied to rich text or rotated cells',
    )
    await expect(rotationDialog.getByRole('button', { name: 'OK' })).toBeDisabled()
    await rotationDialog.getByRole('button', { name: 'Cancel' }).click()

    await chooseUnderline(page, 'A1', 'Single Accounting')
    await save(launched)
    await expect.poll(async () => (await fontModes(path)).A1).toContain('singleAccounting')
    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await fontModes(path)).A1).toBe('<u/>')
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await fontModes(path)).A1).toContain('singleAccounting')
    const edited = await visualWidths(launched, page, 'edited.png')
    expect(edited.A1).toBeGreaterThan(imported.A1 * 1.8)
    expect((await fontModes(path)).D2).toContain('doubleAccounting')
    expect((await fontModes(path)).E1).toContain('singleAccounting')
    expect((await fontModes(path)).F1).toContain('singleAccounting')
    const savedZip = await JSZip.loadAsync(await readFile(path))
    const savedSheet = await savedZip.file('xl/worksheets/sheet1.xml')!.async('string')
    const savedStyles = await savedZip.file('xl/styles.xml')!.async('string')
    expect(savedSheet).toContain('<c r="E1" s="4" t="inlineStr">')
    expect(savedSheet).toContain('<c r="F1" s="5" t="inlineStr">')
    expect(savedStyles).toContain('<alignment textRotation="45"/>')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-accounting-underline')
  }
  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-accounting-underline-reopen',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    const widths = await visualWidths(reopened, page, 'reopened.png')
    expect(widths.A1).toBeGreaterThan(widths.C1 * 1.8)
    expect(widths.B2).toBeGreaterThan(widths.A2 * 1.8)
    await activate(page, 'A1')
    await page.keyboard.press('Meta+1')
    const dialog = page.getByRole('dialog', { name: 'Format Cells' })
    await expect(dialog).toBeVisible()
    await dialog.locator('.dialog-tabs button', { hasText: 'Font' }).click()
    await expect(dialog.getByRole('button', { name: 'Underline' })).toHaveAttribute(
      'data-value',
      'singleAccounting',
    )
    await dialog.getByRole('button', { name: 'Cancel' }).click()
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-accounting-underline-reopen')
  }
})
