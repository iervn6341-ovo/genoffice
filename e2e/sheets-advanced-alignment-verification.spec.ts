import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { buildEditFixture } from '../apps/sheets/tests/fixture-builder'
import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

const evidence =
  process.env.GENOFFICE_ALIGNMENT_VERIFICATION_DIR ??
  join(process.cwd(), 'docs/qa/sheets-advanced-alignment-evidence-2026-10-13')

test.skip(!process.env.GENOFFICE_ALIGNMENT_VERIFY, 'Opt-in Format Cells alignment verification')
test.use({ actionTimeout: 10_000 })

async function fixture(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file(
    'xl/styles.xml',
    styles.replace(
      '<cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>',
      '<cellXfs count="3"><xf/><xf fontId="1" applyFont="1" applyAlignment="1"><alignment horizontal="fill" vertical="distributed" textRotation="37" shrinkToFit="1" readingOrder="2"/></xf><xf applyAlignment="1"><alignment horizontal="general" vertical="justify" textRotation="135"/></xf></cellXfs>',
    ),
  )
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace(
      '<c r="A1" t="s" s="1"><v>0</v></c>',
      '<c r="A1" t="s" s="1"><v>0</v></c><c r="B1" s="2"><v>8</v></c>',
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

async function openAlignment(page: Page) {
  await page
    .locator('nav.ribbon-tabs > button', { hasText: /^Home$/ })
    .first()
    .click()
  await page.locator('[data-ribbon-body] [data-tip="Format Cells"]').first().click()
  const dialog = page.getByRole('dialog', { name: 'Format Cells' })
  await expect(dialog).toBeVisible()
  await dialog.locator('.dialog-tabs button', { hasText: 'Alignment' }).click()
  return dialog
}

async function pick(
  dialog: Awaited<ReturnType<typeof openAlignment>>,
  name: string,
  option: string,
) {
  await dialog.getByRole('button', { name, exact: true }).click()
  await dialog.getByRole('option', { name: option, exact: true }).click()
}

async function nativeCell(path: string, address: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const cell =
    new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`).exec(sheet)?.[0] ?? ''
  const index = Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? ''
  const xf = [...xfs.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)][index]?.[0] ?? ''
  const alignment = /<alignment\b[^>]*\/>/.exec(xf)?.[0] ?? ''
  const attr = (name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(alignment)?.[1] ?? null
  return {
    cell,
    xf,
    alignment,
    horizontal: attr('horizontal'),
    vertical: attr('vertical'),
    rotation: attr('textRotation'),
    shrink: attr('shrinkToFit'),
    readingOrder: attr('readingOrder'),
    merge: /<mergeCell\b[^>]*ref="E1:F1"/.test(sheet),
  }
}

async function save(
  launched: Awaited<ReturnType<typeof launchShell>>,
  check: () => Promise<boolean>,
) {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((part) => part.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
  await expect.poll(check, { timeout: 30_000 }).toBe(true)
}

test('Alignment backfills native modes, previews edited cells, and round-trips independent XF fields', async () => {
  test.setTimeout(180_000)
  await mkdir(evidence, { recursive: true })
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-advanced-alignment-'))
  const path = join(scratch, 'Alignment.xlsx')
  await fixture(path)
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-advanced-alignment',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await pageOf(launched)
    await activate(page, 'A1')
    await expect
      .poll(async () =>
        page.evaluate(() => {
          const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
          const cell = workbook.getActiveSheet().getRange('A1').getCellData()
          return { value: cell?.v, preview: cell?.p?.body?.dataStream ?? '' }
        }),
      )
      .toMatchObject({ value: 'Hello', preview: expect.stringContaining('HelloHello') })
    let dialog = await openAlignment(page)
    await expect(dialog.getByRole('button', { name: 'Horizontal' })).toHaveAttribute(
      'data-value',
      'fill',
    )
    await expect(dialog.getByRole('button', { name: 'Vertical' })).toHaveAttribute(
      'data-value',
      'distributed',
    )
    await expect(dialog.getByRole('spinbutton', { name: 'Degrees' })).toHaveValue('37')
    await expect(dialog.getByRole('button', { name: 'Shrink to fit' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await activate(page, 'B1')
    dialog = await openAlignment(page)
    await expect(dialog.getByRole('button', { name: 'Horizontal' })).toHaveAttribute(
      'data-value',
      'general',
    )
    await expect(dialog.getByRole('button', { name: 'Vertical' })).toHaveAttribute(
      'data-value',
      'justify',
    )
    await expect(dialog.getByRole('spinbutton', { name: 'Degrees' })).toHaveValue('-45')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await activate(page, 'A1:B1')
    dialog = await openAlignment(page)
    await expect(dialog.getByRole('button', { name: 'Horizontal' })).toHaveAttribute(
      'data-value',
      '',
    )
    await expect(
      dialog.locator('label', { has: page.getByRole('button', { name: 'Horizontal' }) }),
    ).toContainText('Mixed')
    await pick(dialog, 'Horizontal', 'Center Across Selection')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(async () =>
        page.evaluate(() => {
          const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
          return workbook.getActiveSheet().getRange('A1').getCellStyleData()?.ht
        }),
      )
      .toBe(2)
    await save(
      launched,
      async () => (await nativeCell(path, 'A1')).horizontal === 'centerContinuous',
    )
    expect(await nativeCell(path, 'A1')).toMatchObject({
      horizontal: 'centerContinuous',
      vertical: 'distributed',
      rotation: '37',
      shrink: '1',
      readingOrder: '2',
    })
    expect(await nativeCell(path, 'B1')).toMatchObject({
      horizontal: 'centerContinuous',
      vertical: 'justify',
      rotation: '135',
    })
    expect((await nativeCell(path, 'B3')).cell).toContain('<f>SUM(C1:C2)</f><v>5</v>')
    expect((await nativeCell(path, 'A1')).merge).toBe(false)

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched, async () => (await nativeCell(path, 'A1')).horizontal === 'fill')
    expect((await nativeCell(path, 'B1')).horizontal).toBe('general')
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(
      launched,
      async () => (await nativeCell(path, 'A1')).horizontal === 'centerContinuous',
    )

    await activate(page, 'A1')
    dialog = await openAlignment(page)
    await pick(dialog, 'Vertical', 'Justify')
    await dialog.getByRole('spinbutton', { name: 'Degrees' }).fill('-37')
    await pick(dialog, 'Shrink to fit', 'Off')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(launched, async () => (await nativeCell(path, 'A1')).rotation === '127')
    expect(await nativeCell(path, 'A1')).toMatchObject({
      horizontal: 'centerContinuous',
      vertical: 'justify',
      rotation: '127',
      shrink: null,
      readingOrder: '2',
    })
    await copyFile(path, join(evidence, 'advanced-alignment-saved.xlsx'))
    await page.screenshot({ path: join(evidence, 'advanced-alignment-edited.png') })
  } finally {
    await closeAndSaveVideo(launched, 'sheets-advanced-alignment')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-advanced-alignment-reopen',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await pageOf(reopened)
    await activate(page, 'A1')
    let dialog = await openAlignment(page)
    await expect(dialog.getByRole('button', { name: 'Horizontal' })).toHaveAttribute(
      'data-value',
      'centerContinuous',
    )
    await expect(dialog.getByRole('button', { name: 'Vertical' })).toHaveAttribute(
      'data-value',
      'justify',
    )
    await expect(dialog.getByRole('spinbutton', { name: 'Degrees' })).toHaveValue('-37')
    await expect(dialog.getByRole('button', { name: 'Shrink to fit' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await dialog.getByRole('spinbutton', { name: 'Degrees' }).fill('42')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(reopened, async () => (await nativeCell(path, 'A1')).rotation === '42')
    expect((await nativeCell(path, 'B3')).cell).toContain('<f>SUM(C1:C2)</f><v>5</v>')

    await activate(page, 'E1:F1')
    dialog = await openAlignment(page)
    await pick(dialog, 'Merge Cells', 'On')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(reopened, async () => (await nativeCell(path, 'A1')).merge)
    expect((await nativeCell(path, 'A1')).horizontal).toBe('centerContinuous')

    await activate(page, 'B1')
    dialog = await openAlignment(page)
    await pick(dialog, 'Horizontal', 'Left')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(async () =>
        page.evaluate(() => {
          const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
          return workbook.getActiveSheet().getRange('B1').getCellStyleData()?.ht
        }),
      )
      .toBe(1)
    await save(reopened, async () => (await nativeCell(path, 'B1')).horizontal === 'left')

    await activate(page, 'B1')
    dialog = await openAlignment(page)
    await pick(dialog, 'Horizontal', 'General')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(async () =>
        page.evaluate(() => {
          const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
          return workbook.getActiveSheet().getRange('B1').getCellStyleData()?.ht
        }),
      )
      .toBe(0)
    await save(reopened, async () => (await nativeCell(path, 'B1')).horizontal === 'general')

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await expect
      .poll(async () =>
        page.evaluate(() => {
          const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
          return workbook.getActiveSheet().getRange('B1').getCellStyleData()?.ht
        }),
      )
      .toBe(1)
    await save(reopened, async () => (await nativeCell(path, 'B1')).horizontal === 'left')
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await expect
      .poll(async () =>
        page.evaluate(() => {
          const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
          return workbook.getActiveSheet().getRange('B1').getCellStyleData()?.ht
        }),
      )
      .toBe(0)
    await save(reopened, async () => (await nativeCell(path, 'B1')).horizontal === 'general')

    await activate(page, 'B1')
    dialog = await openAlignment(page)
    await pick(dialog, 'Horizontal', 'Center')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect
      .poll(async () =>
        page.evaluate(() => {
          const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
          return workbook.getActiveSheet().getRange('B1').getCellStyleData()?.ht
        }),
      )
      .toBe(2)
    await save(reopened, async () => (await nativeCell(path, 'B1')).horizontal === 'center')
    await page.screenshot({ path: join(evidence, 'advanced-alignment-reopened.png') })
    await copyFile(path, join(evidence, 'advanced-alignment-reopened.xlsx'))
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-advanced-alignment-reopen')
  }

  const finalReopen = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-advanced-alignment-final-reopen',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await pageOf(finalReopen)
    await activate(page, 'B1')
    const dialog = await openAlignment(page)
    await expect(dialog.getByRole('button', { name: 'Horizontal' })).toHaveAttribute(
      'data-value',
      'center',
    )
    expect((await nativeCell(path, 'B3')).cell).toContain('<f>SUM(C1:C2)</f><v>5</v>')
  } finally {
    await closeAndSaveVideo(finalReopen, 'sheets-advanced-alignment-final-reopen')
  }
})
