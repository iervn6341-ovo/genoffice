import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { copyFile, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(!process.env.GENOFFICE_SHEETS_MERGED_BORDER_VERIFY, 'Opt-in merged-border verification')

const SOURCE = join(
  process.cwd(),
  'docs/qa/office-mac-field-evidence-2026-10-13/sheets-merged-border/excel-mac-merged-border-source.xlsx',
)

async function addFormulaProbe(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const part = zip.file('xl/worksheets/sheet1.xml')
  if (!part) throw new Error('Excel Mac fixture has no worksheet')
  const original = await part.async('string')
  const patched = original.replace(
    /(<row r="2"[^>]*>[\s\S]*?)(<\/row>)/,
    (_match, cells: string, close: string) => `${cells}<c r="C2"><f>1+2</f><v>3</v></c>${close}`,
  )
  if (patched === original) throw new Error('Excel Mac fixture has no row 2')
  zip.file('xl/worksheets/sheet1.xml', patched)
  await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
}

async function nativeState(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const xfs = [
    ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? '').matchAll(
      /<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g,
    ),
  ].map((match) => match[0])
  const borders = [
    ...(/<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(styles)?.[1] ?? '').matchAll(
      /<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g,
    ),
  ].map((match) => match[0])
  const cell = (address: string) =>
    new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*(?:/>|>[\\s\\S]*?<\\/c>)`).exec(sheet)?.[0] ?? ''
  const styleRef = (address: string) => Number(/\bs="(\d+)"/.exec(cell(address))?.[1] ?? 0)
  const border = (address: string) =>
    borders[Number(/\bborderId="(\d+)"/.exec(xfs[styleRef(address)] ?? '')?.[1] ?? 0)] ?? ''
  return {
    merged: sheet.includes('<mergeCell ref="A1:B2"/>'),
    refs: ['A1', 'B1', 'A2', 'B2'].map(styleRef),
    borders: ['A1', 'B1', 'A2', 'B2'].map(border),
    neighbor: cell('C1'),
    neighborBorder: border('C1'),
    belowNeighborBorder: border('C2'),
    text: cell('A1'),
    formulaCell: cell('C2'),
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
        const active = (window as any).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveRange()
        return active ? `${active.getHeight()}:${active.getWidth()}` : null
      }),
    )
    .toBe(`${height}:${width}`)
}

async function selectedCellMergeState(page: Page): Promise<boolean | null> {
  return page.evaluate(() => {
    const workbook = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
    const cell = workbook.getActiveSheet().getSelection()?.getCurrentCell()
    return cell ? cell.isMerged === true : null
  })
}

async function formatCells(page: Page, tab: 'Border' | 'Alignment') {
  await page.keyboard.press('Meta+1')
  const dialog = page.getByRole('dialog', { name: 'Format Cells' })
  await expect(dialog).toBeVisible()
  await dialog.locator('.dialog-tabs button', { hasText: tab }).click()
  return dialog
}

async function pickState(
  dialog: ReturnType<Page['getByRole']>,
  label: string,
  state: 'on' | 'off',
) {
  const trigger = dialog.getByRole('button', { name: label, exact: true })
  const active = dialog.locator('.gs-dd-pop .gs-dd-item.active')
  for (let attempt = 0; attempt < 3; attempt++) {
    await dialog.page().bringToFront()
    await dialog.page().mouse.move(1, 1)
    await trigger.focus()
    if ((await trigger.getAttribute('aria-expanded')) !== 'true') await trigger.press('ArrowDown')
    await expect(trigger).toHaveAttribute('aria-expanded', 'true')
    await trigger.press('Home')
    try {
      await expect(active).toHaveAttribute('data-value', '', { timeout: 1000 })
      await trigger.press(state === 'off' ? 'End' : 'ArrowDown')
      await expect(active).toHaveAttribute('data-value', state, { timeout: 1000 })
      await trigger.press('Enter')
      await expect(trigger).toHaveAttribute('data-value', state)
      return
    } catch {
      // Electron focus or pointer hover can dismiss or move a newly opened menu.
      if (attempt === 2) throw new Error(`${label} dropdown could not select ${state}`)
    }
  }
}

async function save(launched: Awaited<ReturnType<typeof launchShell>>) {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((part) => part.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
}

test('Excel Mac merged outline and diagonals edit, undo, reopen, and unmerge without hidden seams', async () => {
  test.setTimeout(180_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-merged-border-'))
  const path = join(scratch, 'MergedBorder.xlsx')
  await copyFile(SOURCE, path)
  await addFormulaProbe(path)
  const original = await nativeState(path)
  expect(original.merged).toBe(true)
  expect(original.refs).toEqual([1, 2, 3, 4])
  expect(original.formulaCell).toContain('<f>1+2</f><v>3</v>')
  const launched = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-merged-border',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await selectRange(page, 'A1:B2', 2, 2)
    const dialog = await formatCells(page, 'Border')
    for (const label of ['Top', 'Bottom', 'Left', 'Right', 'Diagonal up'])
      await expect(dialog.getByRole('button', { name: label, exact: true })).toHaveAttribute(
        'data-value',
        'on',
      )
    await expect(dialog.getByRole('button', { name: 'Inside horizontal' })).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Inside vertical' })).toBeDisabled()
    await pickState(dialog, 'Top', 'off')
    await pickState(dialog, 'Diagonal down', 'on')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).refs).not.toEqual(original.refs)
    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).refs).toEqual(original.refs)
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched)
    await expect.poll(async () => (await nativeState(path)).refs).not.toEqual(original.refs)
    const changed = await nativeState(path)
    expect(changed.merged).toBe(true)
    expect(changed.neighbor).toBe(original.neighbor)
    expect(changed.formulaCell).toContain('<f>1+2</f><v>3</v>')
    expect(changed.text).toContain('<v>0</v>')
    for (const border of changed.borders) {
      expect(border).toContain('diagonalUp="1" diagonalDown="1"')
      expect(border).toContain('<diagonal style="dotted"><color auto="1"/></diagonal>')
      expect(border).not.toContain('<top style=')
    }
    expect(changed.borders[0]).toContain('<left style="dotted">')
    expect(changed.borders[1]).toContain('<right style="dotted">')
    expect(changed.borders[2]).toContain('<bottom style="dotted">')
    expect(changed.borders[3]).toContain('<bottom style="dotted">')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-merged-border')
  }

  const reopened = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-merged-border-reopen',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await selectRange(page, 'A1:B2', 2, 2)
    const dialog = await formatCells(page, 'Border')
    await expect(dialog.getByRole('button', { name: 'Top' })).toHaveAttribute('data-value', 'off')
    await expect(dialog.getByRole('button', { name: 'Diagonal down' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await pickState(dialog, 'Top', 'on')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(reopened)
    await expect
      .poll(async () => (await nativeState(path)).borders[0])
      .toContain('<top style="dotted">')
    const beforeUnmerge = await nativeState(path)
    await selectRange(page, 'A1:B2', 2, 2)
    const alignment = await formatCells(page, 'Alignment')
    await pickState(alignment, 'Merge Cells', 'off')
    await alignment.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => selectedCellMergeState(page)).toBe(false)
    await save(reopened)
    await expect.poll(async () => (await nativeState(path)).merged).toBe(false)
    const unmerged = await nativeState(path)
    expect(unmerged.refs).toEqual(beforeUnmerge.refs)
    expect(unmerged.borders).toEqual(beforeUnmerge.borders)
    expect(unmerged.neighbor).toBe(original.neighbor)
    expect(unmerged.formulaCell).toContain('<f>1+2</f><v>3</v>')
    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await expect.poll(() => selectedCellMergeState(page)).toBe(true)
    await save(reopened)
    await expect.poll(async () => (await nativeState(path)).merged).toBe(true)
    const remerged = await nativeState(path)
    expect(remerged.refs).toEqual(beforeUnmerge.refs)
    expect(remerged.borders).toEqual(beforeUnmerge.borders)
    expect(remerged.formulaCell).toContain('<f>1+2</f><v>3</v>')
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await expect.poll(() => selectedCellMergeState(page)).toBe(false)
    await save(reopened)
    await expect.poll(async () => (await nativeState(path)).merged).toBe(false)
    const unmergedAgain = await nativeState(path)
    expect(unmergedAgain.refs).toEqual(beforeUnmerge.refs)
    expect(unmergedAgain.borders).toEqual(beforeUnmerge.borders)
    expect(unmergedAgain.formulaCell).toContain('<f>1+2</f><v>3</v>')
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-merged-border-reopen')
  }
})

test('Inside borders on a selection containing a merge touch only visible seams beside it', async () => {
  test.setTimeout(120_000)
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-merged-border-inside-'))
  const path = join(scratch, 'MergedInside.xlsx')
  await copyFile(SOURCE, path)
  const launched = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-merged-border-inside',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await selectRange(page, 'A1:C2', 2, 3)
    const dialog = await formatCells(page, 'Border')
    await expect(dialog.getByRole('button', { name: 'Inside horizontal' })).toBeEnabled()
    await expect(dialog.getByRole('button', { name: 'Inside vertical' })).toBeEnabled()
    await pickState(dialog, 'Inside horizontal', 'on')
    await pickState(dialog, 'Inside vertical', 'off')
    await expect(dialog.getByRole('button', { name: 'Inside horizontal' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await expect(dialog.getByRole('button', { name: 'Inside vertical' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(launched)
    await expect
      .poll(async () => (await nativeState(path)).neighborBorder)
      .toContain('<bottom style=')
    const zip = await JSZip.loadAsync(await readFile(path))
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    const styles = await zip.file('xl/styles.xml')!.async('string')
    const xfs = [
      ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1] ?? '').matchAll(
        /<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g,
      ),
    ].map((match) => match[0])
    const borders = [
      ...(/<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(styles)?.[1] ?? '').matchAll(
        /<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g,
      ),
    ].map((match) => match[0])
    const border = (address: string) => {
      const cell = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>`).exec(sheet)?.[0] ?? ''
      const xf = xfs[Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)] ?? ''
      return borders[Number(/\bborderId="(\d+)"/.exec(xf)?.[1] ?? 0)] ?? ''
    }
    expect(sheet).toContain('<mergeCell ref="A1:B2"/>')
    expect(sheet).toContain('<c r="C1"')
    expect(border('A1')).not.toContain('<right style=')
    expect(border('A1')).not.toContain('<bottom style=')
    expect(border('B1')).not.toContain('<left style=')
    expect(border('B1')).not.toContain('<bottom style=')
    expect(border('B1')).not.toContain('<right style=')
    expect(border('B2')).not.toContain('<right style=')
    expect(border('C1')).not.toContain('<left style=')
    expect(border('C1')).toContain('<bottom style=')
    expect(border('C2')).not.toContain('<left style=')
    expect(border('C2')).toContain('<top style=')
  } finally {
    await closeAndSaveVideo(launched, 'sheets-merged-border-inside')
  }
  const reopened = await launchShell({
    foreground: true,
    onboardingSeen: true,
    lang: 'en',
    openFile: path,
    videoDir: 'sheets-merged-border-inside-reopen',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await waitForPageWithUrl(reopened.app, '://sheets/', 30_000)
    await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
    await selectRange(page, 'A1:C2', 2, 3)
    const dialog = await formatCells(page, 'Border')
    await expect(dialog.getByRole('button', { name: 'Inside horizontal' })).toHaveAttribute(
      'data-value',
      'on',
    )
    await expect(dialog.getByRole('button', { name: 'Inside vertical' })).toHaveAttribute(
      'data-value',
      'off',
    )
    await pickState(dialog, 'Inside vertical', 'on')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await save(reopened)
    await expect
      .poll(async () => (await nativeState(path)).neighborBorder)
      .toContain('<left style=')
    const restored = await JSZip.loadAsync(await readFile(path))
    const restoredSheet = await restored.file('xl/worksheets/sheet1.xml')!.async('string')
    const restoredStyles = await restored.file('xl/styles.xml')!.async('string')
    const restoredXfs = [
      ...(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(restoredStyles)?.[1] ?? '').matchAll(
        /<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g,
      ),
    ].map((match) => match[0])
    const restoredBorders = [
      ...(/<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(restoredStyles)?.[1] ?? '').matchAll(
        /<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g,
      ),
    ].map((match) => match[0])
    const restoredBorder = (address: string) => {
      const cell = new RegExp(`<c\\b[^>]*\\br="${address}"[^>]*>`).exec(restoredSheet)?.[0] ?? ''
      const xf = restoredXfs[Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)] ?? ''
      return restoredBorders[Number(/\bborderId="(\d+)"/.exec(xf)?.[1] ?? 0)] ?? ''
    }
    expect(restoredBorder('A1')).not.toContain('<right style=')
    expect(restoredBorder('B1')).toContain('<right style=')
    expect(restoredBorder('B2')).toContain('<right style=')
    expect(restoredBorder('C1')).toContain('<left style=')
    expect(restoredBorder('C2')).toContain('<left style=')
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-merged-border-inside-reopen')
  }
})
