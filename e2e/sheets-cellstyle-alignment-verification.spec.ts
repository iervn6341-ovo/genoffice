import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { buildEditFixture } from '../apps/sheets/tests/fixture-builder'
import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

const evidence =
  process.env.GENOFFICE_CELLSTYLE_ALIGNMENT_DIR ??
  join(process.cwd(), 'docs/qa/sheets-cellstyle-alignment-evidence-2026-10-13')

test.skip(
  !process.env.GENOFFICE_CELLSTYLE_ALIGNMENT_VERIFY,
  'Opt-in named-style cell alignment verification',
)

async function fixture(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const original =
    '<cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>'
  if (!styles.includes(original)) throw new Error('Alignment fixture base stylesheet changed')
  zip.file(
    'xl/styles.xml',
    styles.replace(
      original,
      '<cellStyleXfs count="2"><xf/><xf applyAlignment="1"><alignment horizontal="center" vertical="justify" wrapText="1" indent="3"/></xf></cellStyleXfs>' +
        '<cellXfs count="4"><xf xfId="0"/><xf xfId="1" applyAlignment="0"/><xf xfId="1" applyAlignment="1"/><xf xfId="1" applyAlignment="1"><alignment horizontal="left"/></xf></cellXfs>' +
        '<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="ProbeStyle" xfId="1"/></cellStyles>',
    ),
  )
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const source = '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="C1"><v>5</v></c></row>'
  if (!sheet.includes(source)) throw new Error('Alignment fixture base worksheet changed')
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace(
      source,
      '<row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" s="2" t="inlineStr"><is><t>true-empty</t></is></c><c r="C1" s="3"><v>5</v></c></row>',
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
    .poll(() =>
      page.evaluate(() => {
        const range = (window as any).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveRange()
          ?.getRange()
        return range ? [range.startRow, range.endRow, range.startColumn, range.endColumn] : null
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

async function nativeAlignment(path: string, address: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const cell = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>[\\s\\S]*?<\\/c>`).exec(sheet)?.[0]
  if (!cell) throw new Error(`Missing ${address}`)
  const index = Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? ''
  const xf = [...xfs.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)][index]?.[0]
  if (!xf) throw new Error(`Missing XF ${index}`)
  const alignment = /<alignment\b[^>]*\/>/.exec(xf)?.[0] ?? ''
  const attr = (name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(alignment)?.[1] ?? null
  return {
    cell,
    xf,
    styles,
    horizontal: attr('horizontal'),
    vertical: attr('vertical'),
    wrap: attr('wrapText'),
    indent: attr('indent'),
  }
}

async function save(
  launched: Awaited<ReturnType<typeof launchShell>>,
  path: string,
  vertical: string | null,
) {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((part) => part.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
  await expect
    .poll(async () => (await nativeAlignment(path, 'A1')).vertical, { timeout: 30_000 })
    .toBe(vertical)
}

test('cell XF controls current alignment while named style survives edit, history, and reopen', async () => {
  test.setTimeout(180_000)
  await mkdir(evidence, { recursive: true })
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-cellstyle-alignment-'))
  const path = join(scratch, 'CellStyleAlignment.xlsx')
  await fixture(path)
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-cellstyle-alignment',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await pageOf(launched)
    for (const [address, horizontal] of [
      ['A1', 'general'],
      ['B1', 'general'],
      ['C1', 'left'],
    ]) {
      await activate(page, address!)
      const dialog = await openAlignment(page)
      await expect(dialog.getByRole('button', { name: 'Horizontal' })).toHaveAttribute(
        'data-value',
        horizontal!,
      )
      await expect(dialog.getByRole('button', { name: 'Vertical' })).toHaveAttribute(
        'data-value',
        'bottom',
      )
      await expect(dialog.getByRole('button', { name: 'Wrap text' })).toHaveAttribute(
        'data-value',
        'off',
      )
      await expect(dialog.getByRole('spinbutton', { name: 'Indent' })).toHaveValue('')
      await dialog.getByRole('button', { name: 'Cancel' }).click()
    }
    await activate(page, 'A1:C1')
    const dialog = await openAlignment(page)
    await dialog.getByRole('button', { name: 'Vertical' }).click()
    await dialog.getByRole('option', { name: 'Distributed', exact: true }).click()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(launched, path, 'distributed')
    for (const address of ['A1', 'B1', 'C1']) {
      const native = await nativeAlignment(path, address)
      expect(native.vertical).toBe('distributed')
      expect(native.wrap).toBeNull()
      expect(native.indent).toBeNull()
      expect(native.xf).toContain('xfId="1"')
      expect(native.horizontal).toBe(address === 'C1' ? 'left' : null)
      expect(native.styles).toContain(
        '<cellStyleXfs count="2"><xf/><xf applyAlignment="1"><alignment horizontal="center" vertical="justify" wrapText="1" indent="3"/></xf></cellStyleXfs>',
      )
    }
    expect((await nativeAlignment(path, 'B3')).cell).toContain('<f>SUM(C1:C2)</f><v>5</v>')

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched, path, null)
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched, path, 'distributed')
    await page.screenshot({ path: join(evidence, 'cellstyle-alignment-edited.png') })
    await writeFile(join(evidence, 'cellstyle-alignment-saved.xlsx'), await readFile(path))
  } finally {
    await closeAndSaveVideo(launched, 'sheets-cellstyle-alignment')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-cellstyle-alignment-reopen',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await pageOf(reopened)
    for (const [address, horizontal] of [
      ['A1', 'general'],
      ['B1', 'general'],
      ['C1', 'left'],
    ]) {
      await activate(page, address!)
      const dialog = await openAlignment(page)
      await expect(dialog.getByRole('button', { name: 'Horizontal' })).toHaveAttribute(
        'data-value',
        horizontal!,
      )
      await expect(dialog.getByRole('button', { name: 'Vertical' })).toHaveAttribute(
        'data-value',
        'distributed',
      )
      await expect(dialog.getByRole('button', { name: 'Wrap text' })).toHaveAttribute(
        'data-value',
        'off',
      )
      await dialog.getByRole('button', { name: 'Cancel' }).click()
    }
    await page.screenshot({ path: join(evidence, 'cellstyle-alignment-reopened.png') })
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-cellstyle-alignment-reopen')
  }
})
