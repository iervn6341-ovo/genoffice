import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { buildEditFixture } from '../apps/sheets/tests/fixture-builder'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(
  !process.env.GENOFFICE_SHEETS_FONT_EFFECTS_VERIFY,
  'Opt-in native font effects verification',
)

async function fixture(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  const source = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const originalRow = '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="C1"><v>5</v></c></row>'
  if (!source.includes(originalRow)) throw new Error('Font-effects fixture base row changed')
  zip.file(
    'xl/worksheets/sheet1.xml',
    source.replace(
      originalRow,
      '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" t="inlineStr"><is><t>B</t></is></c><c r="C1"><v>5</v></c><c r="D1" t="inlineStr"><is><t>D</t></is></c><c r="E1" t="inlineStr"><is><t>E</t></is></c></row>',
    ),
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function nativeFont(path: string, address: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const cell = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*(?:/>|>[\\s\\S]*?<\\/c>)`).exec(
    sheet,
  )?.[0]
  if (!cell) throw new Error(`Missing cell ${address}`)
  const index = Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? ''
  const xf = [...xfs.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)][index]?.[0]
  const fontId = Number(/\bfontId="(\d+)"/.exec(xf ?? '')?.[1] ?? 0)
  const fonts = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(styles)?.[1] ?? ''
  const font = [...fonts.matchAll(/<font\b[^>]*\/>|<font\b[^>]*>[\s\S]*?<\/font>/g)][fontId]?.[0]
  if (!font) throw new Error(`Missing font ${fontId}`)
  return {
    cell,
    index,
    underline: /<u\b[^>]*\/>/.exec(font)?.[0] ?? null,
    position: /<vertAlign\b[^>]*\/>/.exec(font)?.[0] ?? null,
    formula: /<f>SUM\(C1:C2\)<\/f><v>5<\/v>/.test(sheet),
  }
}

async function activate(page: Page, address: string): Promise<void> {
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
        return range ? `${range.getRow()}:${range.getColumn()}` : null
      }),
    )
    .toBe(`${Number(/\d+/.exec(address)?.[0] ?? 1) - 1}:${address.charCodeAt(0) - 65}`)
}

async function fontDialog(page: Page) {
  await page.keyboard.press('Meta+1')
  const dialog = page.getByRole('dialog', { name: 'Format Cells' })
  await expect(dialog).toBeVisible()
  await dialog.locator('.dialog-tabs button', { hasText: 'Font' }).click()
  return dialog
}

async function choose(page: Page, address: string, underline: string, position?: string) {
  await activate(page, address)
  const dialog = await fontDialog(page)
  const trigger = dialog.getByRole('button', { name: 'Underline' })
  await trigger.click()
  await expect(trigger).toHaveAttribute('aria-expanded', 'true')
  const option = dialog.getByRole('option', { name: underline, exact: true })
  await expect(option).toBeVisible()
  await option.click()
  if (position) {
    await dialog.getByRole('button', { name: 'Position' }).click()
    await dialog.getByRole('option', { name: position, exact: true }).click()
  }
  await dialog.getByRole('button', { name: 'OK' }).click()
}

async function save(launched: Awaited<ReturnType<typeof launchShell>>) {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((part) => part.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
}

test('native underline modes and script positions survive Undo, Save, and process reopen', async () => {
  test.setTimeout(180_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-font-effects-'))
  const path = join(scratch, 'FontEffects.xlsx')
  await fixture(path)
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-font-effects',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await choose(page, 'A1', 'Single Accounting', 'Superscript')
    await choose(page, 'B1', 'Double Accounting', 'Subscript')
    await choose(page, 'C1', 'Double')
    await choose(page, 'D1', 'Single')
    await choose(page, 'E1', 'Single')
    await choose(page, 'E1', 'None')
    await save(launched)
    await expect
      .poll(async () => (await nativeFont(path, 'B1')).underline)
      .toContain('doubleAccounting')
    expect((await nativeFont(path, 'A1')).underline).toContain('singleAccounting')
    expect((await nativeFont(path, 'A1')).position).toContain('superscript')
    expect((await nativeFont(path, 'B1')).position).toContain('subscript')
    expect((await nativeFont(path, 'C1')).underline).toContain('double')
    expect((await nativeFont(path, 'D1')).underline).toBe('<u/>')
    expect((await nativeFont(path, 'E1')).underline).toBeNull()
    expect((await nativeFont(path, 'B1')).formula).toBe(true)

    await page.evaluate(() => {
      const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
      workbook.getActiveSheet().getRange('A1:B1').activate()
    })
    const mixed = await fontDialog(page)
    await expect(mixed.getByRole('button', { name: 'Underline' })).toHaveAttribute('data-value', '')
    await expect(mixed.getByRole('button', { name: 'Position' })).toHaveAttribute('data-value', '')
    await mixed.getByRole('button', { name: 'Cancel' }).click()

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeFont(path, 'E1')).underline).toBe('<u/>')
    expect((await nativeFont(path, 'B1')).formula).toBe(true)
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeFont(path, 'E1')).underline).toBeNull()
    expect((await nativeFont(path, 'B1')).formula).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-font-effects')
  }
  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-font-effects-reopen',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await activate(page, 'A1')
    const dialog = await fontDialog(page)
    await expect(dialog.getByRole('button', { name: 'Underline' })).toHaveAttribute(
      'data-value',
      'singleAccounting',
    )
    await expect(dialog.getByRole('button', { name: 'Position' })).toHaveAttribute(
      'data-value',
      'superscript',
    )
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await activate(page, 'B1')
    const second = await fontDialog(page)
    await expect(second.getByRole('button', { name: 'Underline' })).toHaveAttribute(
      'data-value',
      'doubleAccounting',
    )
    await expect(second.getByRole('button', { name: 'Position' })).toHaveAttribute(
      'data-value',
      'subscript',
    )
    await second.getByRole('button', { name: 'Cancel' }).click()
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-font-effects-reopen')
  }
})
