import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { buildEditFixture } from '../apps/sheets/tests/fixture-builder'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  TEST_WINDOW,
  waitForPageWithUrl,
} from './helpers'

const editorLayoutWidth = 1800
const editorPageZoom = TEST_WINDOW.width / editorLayoutWidth

const evidence =
  process.env.GENOFFICE_VERTICAL_DISTRIBUTION_DIR ??
  join(process.cwd(), 'docs/qa/sheets-vertical-distribution-evidence-2026-10-13')

test.skip(
  !process.env.GENOFFICE_VERTICAL_DISTRIBUTION_VERIFY,
  'Opt-in vertical alignment visual verification',
)

async function fixture(path: string): Promise<void> {
  const zip = await JSZip.loadAsync(await buildEditFixture())
  const styles = await zip.file('xl/styles.xml')!.async('string')
  zip.file(
    'xl/styles.xml',
    styles.replace(
      '<cellXfs count="2"><xf/><xf fontId="1" applyFont="1"/></cellXfs>',
      '<cellXfs count="5"><xf/><xf fontId="1" applyFont="1"/><xf applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf applyAlignment="1"><alignment vertical="justify" wrapText="1"/></xf><xf applyAlignment="1"><alignment vertical="distributed" wrapText="1"/></xf></cellXfs>',
    ),
  )
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const text = '<is><t xml:space="preserve">ONE\nTWO\nTHREE</t></is>'
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace(
      '  </sheetData>',
      `    <row r="8" ht="105" customHeight="1"><c r="D8" t="inlineStr" s="2">${text}</c><c r="E8" t="inlineStr" s="3">${text}</c><c r="F8" t="inlineStr" s="4">${text}</c></row>\n  </sheetData>`,
    ),
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function inheritedFixture(path: string): Promise<void> {
  await fixture(path)
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const text = '<is><t xml:space="preserve">ONE\nTWO\nTHREE</t></is>'
  zip.file(
    'xl/worksheets/sheet1.xml',
    sheet
      .replace('<sheetData>', '<cols><col min="4" max="5" style="4"/></cols><sheetData>')
      .replace(
        '<row r="8" ht="105" customHeight="1">',
        '<row r="8" s="3" customFormat="1" ht="105" customHeight="1">',
      )
      .replace('<c r="E8" t="inlineStr" s="3">', '<c r="E8" t="inlineStr">')
      .replace(
        '  </sheetData>',
        `    <row r="9" ht="105" customHeight="1"><c r="D9" t="inlineStr">${text}</c><c r="E9" t="inlineStr">${text}</c></row>\n  </sheetData>`,
      ),
  )
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

interface CellScreenRect {
  left: number
  top: number
  right: number
  bottom: number
}

interface RectCandidate {
  source: string
  rects: CellScreenRect[]
}

async function cellRectCandidates(
  page: Page,
  addresses: readonly string[],
): Promise<RectCandidate[]> {
  return page.evaluate(
    ({ zoom, addresses }) => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      const local = addresses.map((address) => sheet.getRange(address).getCellRect())
      const origins = [
        { source: 'viewport', left: 0, top: 0 },
        ...[...document.querySelectorAll('canvas')].map((canvas, index) => {
          const bounds = canvas.getBoundingClientRect()
          return { source: `canvas-${index}`, left: bounds.left, top: bounds.top }
        }),
      ]
      return origins.map(({ source, left, top }) => ({
        source,
        rects: local.map((rect) => ({
          left: (left + rect.left) * zoom,
          top: (top + rect.top) * zoom,
          right: (left + rect.right) * zoom,
          bottom: (top + rect.bottom) * zoom,
        })),
      }))
    },
    { zoom: editorPageZoom, addresses },
  )
}

function inkBands(image: PNG, rect: CellScreenRect, scale: number): number[] {
  const left = Math.max(0, Math.round((rect.left + 7) * scale))
  const right = Math.min(image.width, Math.round((rect.right - 7) * scale))
  const top = Math.max(0, Math.round((rect.top + 4) * scale))
  const bottom = Math.min(image.height, Math.round((rect.bottom - 4) * scale))
  const bands: number[] = []
  let activeStart = -1
  for (let y = top; y < bottom; y++) {
    let dark = 0
    for (let x = left; x < right; x++) {
      const pixel = (y * image.width + x) * 4
      if (image.data[pixel]! < 105 && image.data[pixel + 1]! < 105 && image.data[pixel + 2]! < 105)
        dark++
    }
    if (dark > 2 && activeStart < 0) activeStart = y
    if ((dark <= 2 || y === bottom - 1) && activeStart >= 0) {
      bands.push((activeStart + y) / 2 / scale - rect.top)
      activeStart = -1
    }
  }
  return bands
}

function darkPixels(image: PNG, rect: CellScreenRect, scale: number): number {
  const left = Math.max(0, Math.round((rect.left + 7) * scale))
  const right = Math.min(image.width, Math.round((rect.right - 7) * scale))
  const top = Math.max(0, Math.round((rect.top + 4) * scale))
  const bottom = Math.min(image.height, Math.round((rect.bottom - 4) * scale))
  let dark = 0
  for (let y = top; y < bottom; y++) {
    for (let x = left; x < right; x++) {
      const pixel = (y * image.width + x) * 4
      if (image.data[pixel]! < 105 && image.data[pixel + 1]! < 105 && image.data[pixel + 2]! < 105)
        dark++
    }
  }
  return dark
}

async function visibleBands(
  page: Page,
  filename: string,
  addresses: readonly string[] = ['D8', 'E8', 'F8'],
): Promise<number[][]> {
  const candidates = await cellRectCandidates(page, addresses)
  const shot = await page.screenshot({ path: join(evidence, filename) })
  const image = PNG.sync.read(shot)
  const scale = image.width / (await page.evaluate(() => window.innerWidth))
  const measured = candidates.map(({ source, rects }) => ({
    source,
    rects,
    scores: rects.map((rect) => darkPixels(image, rect, scale)),
    bands: rects.map((rect) => inkBands(image, rect, scale)),
  }))
  measured.sort((a, b) => Math.min(...b.scores) - Math.min(...a.scores))
  await writeFile(
    join(evidence, filename.replace(/\.png$/, '-geometry.json')),
    JSON.stringify({ scale, measured }, null, 2),
  )
  return measured[0]?.bands ?? []
}

async function nativeCellVerticalMode(path: string, address: string): Promise<string | null> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const cell = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>`).exec(sheet)?.[0] ?? ''
  const index = Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? ''
  const xf = [...xfs.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)][index]?.[0] ?? ''
  expect(sheet).not.toContain('genofficeVerticalDistribution')
  expect(styles).not.toContain('genofficeVerticalDistribution')
  return /\bvertical="([^"]+)"/.exec(xf)?.[1] ?? null
}

test('imported and edited multiline Justify and Distributed visibly spread line baselines', async () => {
  test.setTimeout(180_000)
  await mkdir(evidence, { recursive: true })
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-vertical-distribution-'))
  const path = join(scratch, 'Vertical.xlsx')
  await fixture(path)
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-vertical-distribution',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await setEditorLayoutWidth(launched.app, '://sheets/', editorLayoutWidth)
    await expect
      .poll(() =>
        page.evaluate(() => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return ['D8', 'E8', 'F8'].map(
            (address) => sheet.getRange(address).getCellData()?.p?.body?.dataStream ?? '',
          )
        }),
      )
      .toEqual([
        expect.stringContaining('THREE'),
        expect.stringContaining('THREE'),
        expect.stringContaining('THREE'),
      ])
    // Move the selection away so its fill cannot affect the pixel measurements.
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('M20').activate()
    })
    await page.waitForTimeout(300)
    const [top, justified, distributed] = await visibleBands(
      page,
      'vertical-distribution-imported.png',
    )
    expect(top?.length).toBeGreaterThanOrEqual(3)
    expect(justified?.length).toBeGreaterThanOrEqual(3)
    expect(distributed?.length).toBeGreaterThanOrEqual(3)
    expect(justified!.at(-1)! - top!.at(-1)!).toBeGreaterThan(35)
    expect(distributed![0]! - justified![0]!).toBeGreaterThan(7)
    expect(distributed!.at(-1)!).toBeLessThan(justified!.at(-1)!)
    await writeFile(
      join(evidence, 'vertical-distribution-pixels.json'),
      JSON.stringify({ top, justified, distributed }, null, 2),
    )
    expect(await nativeCellVerticalMode(path, 'D8')).toBe('top')

    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('D8').activate()
    })
    await page
      .locator('nav.ribbon-tabs > button', { hasText: /^Home$/ })
      .first()
      .click()
    await page.locator('[data-ribbon-body] [data-tip="Format Cells"]').first().click()
    const dialog = page.getByRole('dialog', { name: 'Format Cells' })
    await expect(dialog).toBeVisible()
    await dialog.locator('.dialog-tabs button', { hasText: 'Alignment' }).click()
    await dialog.getByRole('button', { name: 'Vertical', exact: true }).click()
    await dialog.getByRole('option', { name: 'Distributed', exact: true }).click()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('M20').activate()
    })
    await expect
      .poll(async () => (await visibleBands(page, 'vertical-distribution-edited.png'))[0]?.at(-1), {
        timeout: 10_000,
      })
      .toBeGreaterThan(top!.at(-1)! + 25)
    const [edited] = await visibleBands(page, 'vertical-distribution-edited.png')
    expect(Math.abs(edited![0]! - distributed![0]!)).toBeLessThan(5)
    expect(Math.abs(edited!.at(-1)! - distributed!.at(-1)!)).toBeLessThan(5)
    await launched.app.evaluate(({ webContents }) => {
      webContents
        .getAllWebContents()
        .find((part) => part.getURL().includes('://sheets/'))
        ?.send('menu:action', 'save')
    })
    await expect
      .poll(() => nativeCellVerticalMode(path, 'D8'), { timeout: 30_000 })
      .toBe('distributed')
    await writeFile(
      join(evidence, 'vertical-distribution-edited-pixels.json'),
      JSON.stringify({ edited, distributed }, null, 2),
    )
  } finally {
    await closeAndSaveVideo(launched, 'sheets-vertical-distribution')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-vertical-distribution-reopen',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await setEditorLayoutWidth(reopened.app, '://sheets/', editorLayoutWidth)
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('M20').activate()
    })
    await expect
      .poll(
        async () => (await visibleBands(page, 'vertical-distribution-reopened.png'))[0]?.at(-1),
        {
          timeout: 10_000,
        },
      )
      .toBeGreaterThan(65)
    const [edited, justified, distributed] = await visibleBands(
      page,
      'vertical-distribution-reopened.png',
    )
    expect(Math.abs(edited![0]! - distributed![0]!)).toBeLessThan(5)
    expect(Math.abs(edited!.at(-1)! - distributed!.at(-1)!)).toBeLessThan(5)
    expect(justified!.at(-1)!).toBeGreaterThan(edited!.at(-1)!)
    expect(await nativeCellVerticalMode(path, 'D8')).toBe('distributed')
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-vertical-distribution-reopen')
  }
})

test('row and column inherited modes render by native XF precedence and survive an edit', async () => {
  test.setTimeout(180_000)
  await mkdir(evidence, { recursive: true })
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-vertical-inherited-'))
  const path = join(scratch, 'Inherited.xlsx')
  await inheritedFixture(path)
  const addresses = ['D8', 'E8', 'D9']
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-vertical-inherited',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await setEditorLayoutWidth(launched.app, '://sheets/', editorLayoutWidth)
    await expect
      .poll(() =>
        page.evaluate((cells) => {
          const sheet = (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
          return cells.map(
            (address) => sheet.getRange(address).getCellData()?.p?.body?.dataStream ?? '',
          )
        }, addresses),
      )
      .toEqual(addresses.map(() => expect.stringContaining('THREE')))
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('M20').activate()
    })
    const [top, rowJustify, columnDistributed] = await visibleBands(
      page,
      'vertical-inherited-imported.png',
      addresses,
    )
    expect(top?.length).toBeGreaterThanOrEqual(3)
    expect(rowJustify?.length).toBeGreaterThanOrEqual(3)
    expect(columnDistributed?.length).toBeGreaterThanOrEqual(3)
    expect(rowJustify!.at(-1)! - top!.at(-1)!).toBeGreaterThan(35)
    expect(columnDistributed![0]! - rowJustify![0]!).toBeGreaterThan(7)
    expect(columnDistributed!.at(-1)!).toBeLessThan(rowJustify!.at(-1)!)

    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('E8').activate()
    })
    await page
      .locator('nav.ribbon-tabs > button', { hasText: /^Home$/ })
      .first()
      .click()
    await page.locator('[data-ribbon-body] [data-tip="Format Cells"]').first().click()
    const dialog = page.getByRole('dialog', { name: 'Format Cells' })
    await dialog.locator('.dialog-tabs button', { hasText: 'Alignment' }).click()
    await expect(dialog.getByRole('button', { name: 'Vertical' })).toHaveAttribute(
      'data-value',
      'justify',
    )
    await dialog.getByRole('button', { name: 'Vertical', exact: true }).click()
    await dialog.getByRole('option', { name: 'Distributed', exact: true }).click()
    await dialog.getByRole('button', { name: 'OK' }).click()
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('M20').activate()
    })
    await expect
      .poll(
        async () =>
          (await visibleBands(page, 'vertical-inherited-edited.png', addresses))[1]?.at(-1),
        {
          timeout: 10_000,
        },
      )
      .toBeLessThan(rowJustify!.at(-1)! - 5)
    const [topStill, edited, columnStill] = await visibleBands(
      page,
      'vertical-inherited-edited.png',
      addresses,
    )
    expect(Math.abs(edited![0]! - columnStill![0]!)).toBeLessThan(5)
    expect(Math.abs(edited!.at(-1)! - columnStill!.at(-1)!)).toBeLessThan(5)
    expect(Math.abs(topStill!.at(-1)! - top!.at(-1)!)).toBeLessThan(5)
    await launched.app.evaluate(({ webContents }) => {
      webContents
        .getAllWebContents()
        .find((part) => part.getURL().includes('://sheets/'))
        ?.send('menu:action', 'save')
    })
    await expect
      .poll(() => nativeCellVerticalMode(path, 'E8'), { timeout: 30_000 })
      .toBe('distributed')
    const zip = await JSZip.loadAsync(await readFile(path))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toMatch(/<col\b[^>]*min="4"[^>]*max="5"[^>]*style="4"/)
    expect(sheet).toMatch(/<row\b[^>]*r="8"[^>]*s="3"/)
  } finally {
    await closeAndSaveVideo(launched, 'sheets-vertical-inherited')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-vertical-inherited-reopen',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await setEditorLayoutWidth(reopened.app, '://sheets/', editorLayoutWidth)
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('M20').activate()
    })
    const [top, edited, column] = await visibleBands(
      page,
      'vertical-inherited-reopened.png',
      addresses,
    )
    expect(top?.length).toBeGreaterThanOrEqual(3)
    expect(edited?.length).toBeGreaterThanOrEqual(3)
    expect(column?.length).toBeGreaterThanOrEqual(3)
    expect(Math.abs(edited![0]! - column![0]!)).toBeLessThan(5)
    expect(Math.abs(edited!.at(-1)! - column!.at(-1)!)).toBeLessThan(5)
    expect(edited!.at(-1)! - top!.at(-1)!).toBeGreaterThan(35)
    expect(await nativeCellVerticalMode(path, 'E8')).toBe('distributed')
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-vertical-inherited-reopen')
  }
})

test('new multiline text in a row-inherited blank keeps Justify through undo, save, and reopen', async () => {
  test.setTimeout(180_000)
  await mkdir(evidence, { recursive: true })
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-vertical-inherited-blank-'))
  const path = join(scratch, 'InheritedBlank.xlsx')
  await inheritedFixture(path)
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-vertical-inherited-blank',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await setEditorLayoutWidth(launched.app, '://sheets/', editorLayoutWidth)
    const cellState = () =>
      page.evaluate(() => {
        const sheet = (window as any).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveSheet()
        const data = sheet.getRange('G8').getCellData()
        return {
          text: data?.p?.body?.dataStream ?? data?.v ?? null,
          mode: data?.custom?.genofficeVerticalDistribution ?? null,
        }
      })
    expect((await cellState()).text).toBeNull()
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('G8').setValue('ONE\nTWO\nTHREE')
      sheet.getRange('M20').activate()
    })
    await expect.poll(cellState, { timeout: 15_000 }).toMatchObject({
      text: expect.stringContaining('THREE'),
      mode: 'justify',
    })
    const [top, justified] = await visibleBands(page, 'vertical-inherited-blank-edited.png', [
      'D8',
      'G8',
    ])
    expect(top?.length).toBeGreaterThanOrEqual(3)
    expect(justified?.length).toBeGreaterThanOrEqual(3)
    expect(justified!.at(-1)! - top!.at(-1)!).toBeGreaterThan(35)

    await page.evaluate(async () => (window as any).__genofficeDebug.univerAPI.undo())
    await expect.poll(async () => (await cellState()).text).toBeNull()
    await page.evaluate(async () => (window as any).__genofficeDebug.univerAPI.redo())
    await expect.poll(cellState, { timeout: 15_000 }).toMatchObject({
      text: expect.stringContaining('THREE'),
      mode: 'justify',
    })
    await launched.app.evaluate(({ webContents }) => {
      webContents
        .getAllWebContents()
        .find((part) => part.getURL().includes('://sheets/'))
        ?.send('menu:action', 'save')
    })
    await expect
      .poll(
        async () => {
          const zip = await JSZip.loadAsync(await readFile(path))
          return (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).includes('r="G8"')
        },
        { timeout: 30_000 },
      )
      .toBe(true)
    const zip = await JSZip.loadAsync(await readFile(path))
    const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheetXml).toMatch(/<row\b[^>]*r="8"[^>]*s="3"/)
    expect(sheetXml).toMatch(/<c\b[^>]*r="G8"(?![^>]*\bs=)[^>]*>/)
    expect(sheetXml).not.toContain('genofficeVerticalDistribution')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-vertical-inherited-blank')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-vertical-inherited-blank-reopen',
    openFile: path,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await setEditorLayoutWidth(reopened.app, '://sheets/', editorLayoutWidth)
    await expect
      .poll(
        () =>
          page.evaluate(() => {
            const sheet = (window as any).__genofficeDebug.univerAPI
              .getActiveWorkbook()
              .getActiveSheet()
            const data = sheet.getRange('G8').getCellData()
            return {
              text: data?.p?.body?.dataStream ?? data?.v ?? null,
              mode: data?.custom?.genofficeVerticalDistribution ?? null,
            }
          }),
        { timeout: 15_000 },
      )
      .toMatchObject({
        text: expect.stringContaining('THREE'),
        mode: 'justify',
      })
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet.getRange('M20').activate()
    })
    const [top, justified] = await visibleBands(page, 'vertical-inherited-blank-reopened.png', [
      'D8',
      'G8',
    ])
    expect(top?.length).toBeGreaterThanOrEqual(3)
    expect(justified?.length).toBeGreaterThanOrEqual(3)
    expect(justified!.at(-1)! - top!.at(-1)!).toBeGreaterThan(35)
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-vertical-inherited-blank-reopen')
  }
})
