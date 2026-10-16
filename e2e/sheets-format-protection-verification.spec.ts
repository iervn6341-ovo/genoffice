import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdir, mkdtemp, readFile, writeFile, copyFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { buildEditFixture } from '../apps/sheets/tests/fixture-builder'
import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

const evidence =
  process.env.GENOFFICE_FORMAT_PROTECTION_VERIFICATION_DIR ??
  join(process.cwd(), 'docs/qa/sheets-format-protection-evidence-2026-10-13')

test.skip(
  !process.env.GENOFFICE_FORMAT_PROTECTION_VERIFY,
  'Opt-in Format Cells protection verification',
)
test.use({ actionTimeout: 10_000 })

async function fixture(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file(
    'xl/styles.xml',
    styles.replace(
      '<cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>',
      '<cellXfs count="3"><xf/><xf fontId="1" applyFont="1" applyProtection="1"><protection locked="0"/></xf><xf applyProtection="1"><protection hidden="1"/></xf></cellXfs>',
    ),
  )
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet
      .replace(
        '<c r="A1" t="s" s="1"><v>0</v></c>',
        '<c r="A1" t="s" s="1"><v>0</v></c><c r="B1" s="2"><v>8</v></c>',
      )
      .replace(
        '<row r="5" ht="20" customHeight="1"/>',
        '<row r="5" ht="20" customHeight="1"><c r="C5" s="2"/></row>',
      ),
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function pageOf(launched: Awaited<ReturnType<typeof launchShell>>): Promise<Page> {
  const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
  await setEditorLayoutWidth(launched.app, '://sheets/', 2200)
  return page
}

async function activate(page: Page, address: string): Promise<void> {
  const expected = await page.evaluate((address) => {
    const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
    const range = workbook.getActiveSheet().getRange(address)
    const bounds = range.getRange()
    range.activate()
    return [bounds.startRow, bounds.endRow, bounds.startColumn, bounds.endColumn]
  }, address)
  await expect
    .poll(async () =>
      page.evaluate(() => {
        const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
        const bounds = workbook.getActiveRange()?.getRange()
        return bounds
          ? [bounds.startRow, bounds.endRow, bounds.startColumn, bounds.endColumn]
          : null
      }),
    )
    .toEqual(expected)
}

async function openProtection(page: Page) {
  await page
    .locator('nav.ribbon-tabs > button', { hasText: /^Home$/ })
    .first()
    .click()
  await page.locator('[data-ribbon-body] [data-tip="Format Cells"]').first().click()
  const dialog = page.getByRole('dialog', { name: 'Format Cells' })
  await expect(dialog).toBeVisible()
  await dialog.locator('.dialog-tabs button', { hasText: 'Protection' }).click()
  return dialog
}

async function savedCell(path: string, address: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const cell =
    new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`).exec(sheet)?.[0] ?? ''
  const index = Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? ''
  const xf = [...xfs.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)][index]?.[0] ?? ''
  return {
    cell,
    locked: !/<protection\b[^>]*\blocked="0"/.test(xf),
    hidden: /<protection\b[^>]*\bhidden="1"/.test(xf),
  }
}

async function save(
  launched: Awaited<ReturnType<typeof launchShell>>,
  path: string,
  check: () => Promise<boolean>,
): Promise<void> {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((part) => part.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
  await expect.poll(check, { timeout: 30_000 }).toBe(true)
  await expect
    .poll(async () => (await readFile(path)).length, { timeout: 10_000 })
    .toBeGreaterThan(0)
}

test('Format Cells backfills mixed flags and preserves the opposite flag through Undo, Save, and reopen', async () => {
  test.setTimeout(180_000)
  await mkdir(evidence, { recursive: true })
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-cell-protection-'))
  const path = join(scratch, 'Protection.xlsx')
  await fixture(path)
  expect(await savedCell(path, 'A1')).toMatchObject({ locked: false, hidden: false })
  expect(await savedCell(path, 'B1')).toMatchObject({ locked: true, hidden: true })
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-cell-protection',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await pageOf(launched)
    await activate(page, 'A1:K1000')
    let dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Hidden' })).toBeDisabled()
    await expect(dialog.getByRole('status')).toContainText('10,000 cells or fewer')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await activate(page, 'A1')
    dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await expect(dialog.getByRole('button', { name: 'Hidden' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await activate(page, 'D5')
    dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute('data-value', 'on')
    await expect(dialog.getByRole('button', { name: 'Hidden' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await activate(page, 'C5')
    dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute('data-value', 'on')
    await expect(dialog.getByRole('button', { name: 'Hidden' })).toHaveAttribute('data-value', 'on')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await activate(page, 'A1:B1')
    dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute('data-value', '')
    await expect(dialog.getByRole('button', { name: 'Hidden' })).toHaveAttribute('data-value', '')
    await expect(dialog.getByRole('button', { name: 'Locked' })).toContainText('(unchanged)')
    await expect(
      dialog.locator('label', { has: page.getByRole('button', { name: 'Locked' }) }),
    ).toContainText('Mixed')
    await dialog.getByRole('button', { name: 'Locked' }).click()
    await dialog.getByRole('option', { name: 'On' }).click()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await activate(page, 'A1')
    dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await activate(page, 'A1')
    dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute('data-value', 'on')
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await save(launched, path, async () => (await savedCell(path, 'A1')).locked)
    expect(await savedCell(path, 'A1')).toMatchObject({ locked: true, hidden: false })
    expect(await savedCell(path, 'B1')).toMatchObject({ locked: true, hidden: true })
    expect((await savedCell(path, 'B3')).cell).toContain('<f>SUM(C1:C2)</f><v>5</v>')
    expect(await savedCell(path, 'C5')).toMatchObject({ locked: true, hidden: true })
    await copyFile(path, join(evidence, 'format-protection-saved.xlsx'))

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched, path, async () => !(await savedCell(path, 'A1')).locked)
    expect(await savedCell(path, 'B1')).toMatchObject({ locked: true, hidden: true })
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched, path, async () => (await savedCell(path, 'A1')).locked)

    await activate(page, 'B3')
    dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute('data-value', 'on')
    await expect(dialog.getByRole('button', { name: 'Hidden' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await dialog.getByRole('button', { name: 'Hidden' }).click()
    await dialog.getByRole('option', { name: 'On' }).click()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(launched, path, async () => (await savedCell(path, 'B3')).hidden)
    expect(await savedCell(path, 'B3')).toMatchObject({
      locked: true,
      hidden: true,
      cell: expect.stringContaining('<f>SUM(C1:C2)</f><v>5</v>'),
    })
    await copyFile(path, join(evidence, 'format-protection-formula-saved.xlsx'))
    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched, path, async () => !(await savedCell(path, 'B3')).hidden)
    expect((await savedCell(path, 'B3')).cell).toContain('<f>SUM(C1:C2)</f><v>5</v>')
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched, path, async () => (await savedCell(path, 'B3')).hidden)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-cell-protection')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-cell-protection-reopen',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await pageOf(reopened)
    await activate(page, 'B1')
    let dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute('data-value', 'on')
    await expect(dialog.getByRole('button', { name: 'Hidden' })).toHaveAttribute('data-value', 'on')
    await dialog.getByRole('button', { name: 'Hidden' }).click()
    await dialog.getByRole('option', { name: 'Off' }).click()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(reopened, path, async () => !(await savedCell(path, 'B1')).hidden)
    expect(await savedCell(path, 'B1')).toMatchObject({ locked: true, hidden: false })
    expect((await savedCell(path, 'B3')).cell).toContain('<f>SUM(C1:C2)</f><v>5</v>')

    await activate(page, 'B3')
    dialog = await openProtection(page)
    await expect(dialog.getByRole('button', { name: 'Locked' })).toHaveAttribute('data-value', 'on')
    await expect(dialog.getByRole('button', { name: 'Hidden' })).toHaveAttribute('data-value', 'on')
    await page.screenshot({ path: join(evidence, 'format-protection-reopened-dialog.png') })
    await dialog.getByRole('button', { name: 'Hidden' }).click()
    await dialog.getByRole('option', { name: 'Off' }).click()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(reopened, path, async () => !(await savedCell(path, 'B3')).hidden)
    expect(await savedCell(path, 'B3')).toMatchObject({
      locked: true,
      hidden: false,
      cell: expect.stringContaining('<f>SUM(C1:C2)</f><v>5</v>'),
    })
    await page.screenshot({ path: join(evidence, 'format-protection-reopened.png') })
    await copyFile(path, join(evidence, 'format-protection-reopened.xlsx'))
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-cell-protection-reopen')
  }
})
