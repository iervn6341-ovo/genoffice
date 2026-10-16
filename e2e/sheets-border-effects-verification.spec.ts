import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { copyFile, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { buildEditFixture } from '../apps/sheets/tests/fixture-builder'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(!process.env.GENOFFICE_SHEETS_BORDER_VERIFY, 'Opt-in native border verification')

async function fixture(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const original = '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="C1"><v>5</v></c></row>'
  if (!sheet.includes(original)) throw new Error('The border fixture base row changed.')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace(
      original,
      '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" t="inlineStr"><is><t>B</t></is></c><c r="C1"><v>5</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>3</v></c></row>',
    ),
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function nativeBorder(path: string, address: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const cell = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*(?:/>|>[\\s\\S]*?<\\/c>)`).exec(
    sheet,
  )?.[0]
  if (!cell) throw new Error(`Missing ${address}`)
  const index = Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)
  const xfs = [
    ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? '').matchAll(
      /<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g,
    ),
  ].map((match) => match[0])
  const borderId = Number(/\bborderId="(\d+)"/.exec(xfs[index] ?? '')?.[1] ?? 0)
  const borders = [
    ...(/<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(styles)?.[1] ?? '').matchAll(
      /<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g,
    ),
  ].map((match) => match[0])
  return {
    cell,
    border: borders[borderId] ?? '',
    formula: /<f>SUM\(C1:C2\)<\/f><v>5<\/v>/.test(sheet),
  }
}

async function selectRange(page: Page, address: string, height: number, width: number) {
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
        return range ? `${range.getHeight()}:${range.getWidth()}` : null
      }),
    )
    .toBe(`${height}:${width}`)
}

async function borderDialog(page: Page) {
  await page.keyboard.press('Meta+1')
  const dialog = page.getByRole('dialog', { name: 'Format Cells' })
  await expect(dialog).toBeVisible()
  await dialog.locator('.dialog-tabs button', { hasText: 'Border' }).click()
  return dialog
}

async function pickOn(dialog: ReturnType<Page['getByRole']>, label: string) {
  const trigger = dialog.getByRole('button', { name: label, exact: true })
  const active = dialog.locator('.gs-dd-pop .gs-dd-item.active')
  await trigger.focus()
  await trigger.press('ArrowDown')
  await expect(trigger).toHaveAttribute('aria-expanded', 'true')
  await trigger.press('Home')
  await expect(active).toHaveAttribute('data-value', '')
  await trigger.press('ArrowDown')
  await expect(active).toHaveAttribute('data-value', 'on')
  await trigger.press('Enter')
  await expect(trigger).toHaveAttribute('data-value', 'on')
}

async function pickOnByPointer(dialog: ReturnType<Page['getByRole']>, label: string) {
  const trigger = dialog.getByRole('button', { name: label, exact: true })
  await trigger.click()
  await expect(trigger).toHaveAttribute('aria-expanded', 'true')
  await dialog.getByRole('option', { name: 'On', exact: true }).click()
  await expect(trigger).toHaveAttribute('data-value', 'on')
}

async function pickDoubleLineByKeyboard(dialog: ReturnType<Page['getByRole']>) {
  const trigger = dialog.getByRole('button', { name: 'Line style', exact: true })
  const active = dialog.locator('.gs-dd-pop .gs-dd-item.active')
  await trigger.focus()
  await trigger.press('ArrowDown')
  await expect(trigger).toHaveAttribute('aria-expanded', 'true')
  await trigger.press('Home')
  await expect(active).toHaveAttribute('data-value', 'thin')
  for (const value of ['medium', 'thick', 'double']) {
    await trigger.press('ArrowDown')
    await expect(active).toHaveAttribute('data-value', value)
  }
  await trigger.press('Enter')
  await expect(trigger).toHaveAttribute('data-value', 'double')
}

async function save(launched: Awaited<ReturnType<typeof launchShell>>) {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((part) => part.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
}

test('per-edge, inside, and diagonal borders survive Undo, Save, and fresh reopen', async () => {
  test.setTimeout(180_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-border-effects-'))
  const path = join(scratch, 'BorderEffects.xlsx')
  await fixture(path)
  const launched = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-border-effects',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await selectRange(page, 'A1:B2', 2, 2)
    const dialog = await borderDialog(page)
    await pickDoubleLineByKeyboard(dialog)
    await dialog.getByRole('button', { name: 'Color' }).click()
    await dialog.locator('.gcp-standard-row .gcp-swatch[title="Blue"]').click()
    for (const edge of ['Top', 'Left', 'Inside horizontal', 'Inside vertical', 'Diagonal up']) {
      await pickOn(dialog, edge)
    }
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return sheet.getRange('A1').getCellStyleData()?.bd?.r?.s ?? null
        }),
      )
      .toBe(7)
    await save(launched)
    await expect
      .poll(async () => (await nativeBorder(path, 'A1')).border)
      .toContain('diagonalUp="1"')
    expect((await nativeBorder(path, 'A1')).border).toContain('<right style="double">')
    expect((await nativeBorder(path, 'A1')).border).toContain('<bottom style="double">')
    expect((await nativeBorder(path, 'B1')).border).toContain('<left style="double">')
    expect((await nativeBorder(path, 'A2')).border).toContain('<top style="double">')
    expect((await nativeBorder(path, 'B2')).border).toContain('<diagonal style="double">')
    expect((await nativeBorder(path, 'A1')).border).toContain('rgb="FF0070C0"')
    expect((await nativeBorder(path, 'B2')).formula).toBe(true)

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeBorder(path, 'A1')).border).toBe('<border/>')
    expect((await nativeBorder(path, 'B1')).cell).toContain('<t>B</t>')
    expect((await nativeBorder(path, 'B2')).formula).toBe(true)
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched)
    await expect
      .poll(async () => (await nativeBorder(path, 'A1')).border)
      .toContain('diagonalUp="1"')
    expect((await nativeBorder(path, 'B2')).formula).toBe(true)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-border-effects')
  }

  const reopened = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-border-effects-reopen',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await selectRange(page, 'A1:B2', 2, 2)
    const dialog = await borderDialog(page)
    await expect(dialog.getByRole('button', { name: 'Inside horizontal' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await expect(dialog.getByRole('button', { name: 'Inside vertical' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await expect(dialog.getByRole('button', { name: 'Diagonal up' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await expect(dialog.getByRole('button', { name: 'Line style' })).toHaveAttribute(
      'data-value',
      'double',
    )
    await pickOn(dialog, 'Diagonal down')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(reopened)
    await expect
      .poll(async () => (await nativeBorder(path, 'A1')).border)
      .toContain('diagonalDown="1"')
    await selectRange(page, 'A1', 1, 1)
    const one = await borderDialog(page)
    await expect(one.getByRole('button', { name: 'Inside horizontal' })).toBeDisabled()
    await expect(one.getByRole('button', { name: 'Inside vertical' })).toBeDisabled()
    await one.getByRole('button', { name: 'Cancel' }).click()
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-border-effects-reopen')
  }
})

test('Excel Mac dotted borders backfill and retain automatic color across direction edits', async () => {
  test.setTimeout(180_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-mac-border-'))
  const path = join(scratch, 'MacBorder.xlsx')
  await copyFile(
    join(
      process.cwd(),
      'docs/qa/office-mac-field-evidence-2026-10-13/sheets-border-effects/excel-mac-border-source.xlsx',
    ),
    path,
  )
  const launched = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-mac-border',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await selectRange(page, 'A1:B2', 2, 2)
    const imported = await borderDialog(page)
    for (const edge of [
      'Top',
      'Bottom',
      'Left',
      'Right',
      'Inside horizontal',
      'Inside vertical',
      'Diagonal up',
    ])
      await expect(imported.getByRole('button', { name: edge, exact: true })).toHaveAttribute(
        'data-value',
        'on',
      )
    await expect(imported.getByRole('button', { name: 'Diagonal down' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await expect(imported.getByRole('button', { name: 'Line style' })).toHaveAttribute(
      'data-value',
      'dotted',
    )
    await imported.getByRole('button', { name: 'Cancel' }).click()

    await selectRange(page, 'A1', 1, 1)
    const single = await borderDialog(page)
    await pickOnByPointer(single, 'Diagonal down')
    await single.getByRole('button', { name: 'OK' }).click()
    await save(launched)
    await expect
      .poll(async () => (await nativeBorder(path, 'A1')).border)
      .toContain('diagonalDown="1"')
    expect((await nativeBorder(path, 'A1')).border).toContain(
      '<diagonal style="dotted"><color auto="1"/></diagonal>',
    )
    expect((await nativeBorder(path, 'B1')).border).not.toContain('diagonalDown="1"')

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched)
    await expect
      .poll(async () => (await nativeBorder(path, 'A1')).cell)
      .toContain('<c r="A1" s="1"><v>1</v></c>')
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched)
    await expect
      .poll(async () => (await nativeBorder(path, 'A1')).border)
      .toContain('diagonalDown="1"')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-mac-border')
  }

  const reopened = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-mac-border-reopen',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await selectRange(page, 'A1', 1, 1)
    const dialog = await borderDialog(page)
    await expect(dialog.getByRole('button', { name: 'Diagonal up' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await expect(dialog.getByRole('button', { name: 'Diagonal down' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await expect(dialog.getByRole('button', { name: 'Line style' })).toHaveAttribute(
      'data-value',
      'dotted',
    )
    await dialog.getByRole('button', { name: 'Cancel' }).click()
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-mac-border-reopen')
  }
})
