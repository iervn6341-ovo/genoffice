import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  closeAndSaveVideo,
  launchShell,
  setEditorLayoutWidth,
  waitForPageWithUrl,
  type LaunchedApp,
} from './helpers'

/**
 * Lifecycle for the Home-tab additions: perform → undo → redo → Save → read the xlsx back →
 * reopen. Links, notes, series and Paste Special results must reach the saved file, and the
 * undo of a clear must bring them back into the next save.
 */
const api = '(window).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()'

async function openSheets(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-home-gaps-save',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.page.locator('.quick-card').nth(1).click()
  const p = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await p.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  await p.waitForTimeout(1500)
  await setEditorLayoutWidth(launched.app, '://sheets/', 2400)
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')
/** QA_KEEP_DIR=<dir> keeps a copy of each saved workbook, for opening in real Excel */
async function keep(file: string, name: string): Promise<void> {
  const dir = process.env.QA_KEEP_DIR
  if (dir && existsSync(file)) await copyFile(file, join(dir, name))
}
const activate = (p: Page, a1: string) => p.evaluate(`${api}.getRange('${a1}').activate()`)
const raw = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getRawValues()`) as Promise<unknown[][]>
const style = (p: Page, a1: string) =>
  p.evaluate(`${api}.getRange('${a1}').getCellStyleData()`) as Promise<{
    ul?: { s?: number } | null
    bl?: number
    n?: { pattern?: string }
  } | null>
const setValues = (p: Page, a1: string, data: unknown[][]) =>
  p.evaluate(`${api}.getRange('${a1}').setValues(${JSON.stringify(data)})`)
const undo = async (p: Page) => {
  await p.locator('button[aria-label^="Undo"]').first().click()
  await p.waitForTimeout(500)
}
const redo = async (p: Page) => {
  await p.locator('button[aria-label^="Redo"]').first().click()
  await p.waitForTimeout(500)
}
const pick = async (p: Page, trigger: string, option: string | RegExp) => {
  await p.locator(`[data-ribbon-body] [aria-label="${trigger}"]`).first().click()
  await p.locator('.menu-select-drop button', { hasText: option }).first().click()
  await p.waitForTimeout(700)
}
const tab = (p: Page, name: string) =>
  p
    .locator('nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)', {
      hasText: new RegExp(`^${name}$`),
    })
    .first()
    .click()

async function stubSaveAs(l: LaunchedApp, path: string): Promise<void> {
  await l.app.evaluate(({ dialog }, target) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: target })
    dialog.showMessageBox = (async () => ({ response: 0, checkboxChecked: false })) as never
  }, path)
}

/** Save As the first time, plain Save afterwards; waits until the file is rewritten */
async function save(l: LaunchedApp, p: Page, path: string, first = false): Promise<JSZip> {
  const before = existsSync(path) ? (await readFile(path)).length : -1
  if (first) {
    await stubSaveAs(l, path)
    await p.locator('button[aria-label^="Save As"]').first().click()
  } else {
    await p.locator('button[aria-label^="Save ("]').first().click()
  }
  await expect
    .poll(async () => (existsSync(path) ? JSON.stringify((await readFile(path)).length) : '-1'), {
      timeout: 20_000,
    })
    .not.toBe(JSON.stringify(before))
  // the writer finishes its rename after the size changes
  await new Promise((r) => setTimeout(r, 600))
  return JSZip.loadAsync(await readFile(path))
}
const sheetXml = async (zip: JSZip): Promise<string> =>
  (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).replace(/\s+/g, ' ')
const hyperlinkRefs = (xml: string): string[] =>
  [...xml.matchAll(/<hyperlink [^>]*ref="([A-Z]+\d+)"/g)].map((m) => m[1]!).sort()

async function addLink(p: Page, cell: string, url: string): Promise<void> {
  await activate(p, cell)
  await tab(p, 'Insert')
  await p.locator('[data-ribbon-body] button', { hasText: /^Link/ }).first().click()
  const dialog = p.getByRole('dialog', { name: 'Insert link' })
  await dialog.getByRole('textbox').fill(url)
  await dialog.getByRole('button', { name: 'OK' }).click()
  await p.waitForTimeout(500)
  await tab(p, 'Home')
}

test('Clear Hyperlinks keeps the look, Remove Hyperlinks drops it, Clear Contents takes the link; saves follow, undo restores', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-clear-links-'))
  const file = join(dir, 'links.xlsx')
  try {
    await setValues(p, 'A1:A3', [['one'], ['two'], ['three']])
    await addLink(p, 'A1', 'https://example.com/1')
    await addLink(p, 'A2', 'https://example.com/2')
    await addLink(p, 'A3', 'https://example.com/3')
    expect((await style(p, 'A1'))?.ul?.s).toBe(1)

    let zip = await save(launched, p, file, true)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A1', 'A2', 'A3'])

    // Clear Hyperlinks over A1:A2: the links go, text and the blue underline stay (Excel)
    await activate(p, 'A1:A2')
    await pick(p, 'Clear', 'Clear Hyperlinks')
    expect((await style(p, 'A1'))?.ul?.s).toBe(1)
    expect(await raw(p, 'A1:A2')).toEqual([['one'], ['two']])
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A3'])

    // ⌘Z puts both links back — into the next save too
    await undo(p)
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A1', 'A2', 'A3'])
    await redo(p)
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A3'])
    await undo(p)

    // Remove Hyperlinks takes the link AND the hyperlink formatting
    await activate(p, 'A1:A2')
    await pick(p, 'Clear', 'Remove Hyperlinks')
    expect((await style(p, 'A1'))?.ul?.s ?? 0).not.toBe(1)
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A3'])
    await undo(p)
    expect((await style(p, 'A1'))?.ul?.s).toBe(1)
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A1', 'A2', 'A3'])

    // Clear Contents on a link cell takes the link with the text (Excel), the format stays
    await activate(p, 'A3')
    await pick(p, 'Clear', 'Clear Contents')
    expect(await raw(p, 'A3')).toEqual([[null]])
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A1', 'A2'])

    // one ⌘Z brings back the text AND the link: the clear is a single undo step
    await undo(p)
    expect(await raw(p, 'A3')).toEqual([['three']])
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A1', 'A2', 'A3'])
  } finally {
    await keep(file, 'links.xlsx')
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Clear Comments and Notes removes the saved note; ⌘Z brings it back into the file', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-clear-notes-'))
  const file = join(dir, 'notes.xlsx')
  try {
    await setValues(p, 'A1:A2', [[1], [2]])
    for (const row of [0, 1]) {
      await p.evaluate(
        `${api}.getRange(${row}, 0, 1, 1).createOrUpdateNote({ id: 'n${row}', row: ${row}, col: 0, width: 200, height: 80, note: 'note ${row + 1}' })`,
      )
    }
    const comments = async (zip: JSZip): Promise<string> => {
      const name = Object.keys(zip.files).find((f) => /^xl\/comments\d*\.xml$/.test(f))
      return name ? (await zip.file(name)!.async('string')).replace(/\s+/g, ' ') : ''
    }
    let zip = await save(launched, p, file, true)
    const withNotes = await comments(zip)
    expect(withNotes).toContain('note 1')
    expect(withNotes).toContain('note 2')

    await activate(p, 'A1')
    await pick(p, 'Clear', 'Clear Comments and Notes')
    zip = await save(launched, p, file)
    const afterClear = await comments(zip)
    expect(afterClear).not.toContain('note 1')
    expect(afterClear).toContain('note 2')
    expect(await raw(p, 'A1:A2')).toEqual([[1], [2]])

    await undo(p)
    zip = await save(launched, p, file)
    expect(await comments(zip)).toContain('note 1')

    // Clear All takes the text and the note together
    await activate(p, 'A2')
    await pick(p, 'Clear', 'Clear All')
    zip = await save(launched, p, file)
    const afterAll = await comments(zip)
    expect(afterAll).not.toContain('note 2')
    expect(afterAll).toContain('note 1')
    expect(await raw(p, 'A2')).toEqual([[null]])
  } finally {
    await keep(file, 'notes.xlsx')
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Series, Fill Up and Paste Special results are in the saved file and survive a reopen', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-series-save-'))
  const file = join(dir, 'series.xlsx')
  try {
    // linear series down A1:A5 (1 3 5 7 9)
    await setValues(p, 'A1', [[1]])
    await activate(p, 'A1:A5')
    await pick(p, 'Fill', /Series/)
    const dialog = p.locator('.fill-series-dialog')
    await dialog.getByLabel('Step value:').fill('2')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => raw(p, 'A1:A5')).toEqual([[1], [3], [5], [7], [9]])

    // Fill Up B1:B3 ← bottom cell
    await setValues(p, 'B3', [[42]])
    await activate(p, 'B1:B3')
    await pick(p, 'Fill', 'Fill Up')
    await expect.poll(() => raw(p, 'B1:B3')).toEqual([[42], [42], [42]])

    // Transpose A1:A5 → D1:H1
    await activate(p, 'A1:A5')
    await p.locator('[data-tip^="Copy"]').first().click()
    await p.waitForTimeout(400)
    await activate(p, 'D1')
    await p.locator('[data-ribbon-body] [aria-label="Paste Special"]').first().click()
    await p.locator('.menu-select-drop button', { hasText: 'Transpose' }).first().click()
    await expect.poll(() => raw(p, 'D1:H1')).toEqual([[1, 3, 5, 7, 9]])

    // Paste Special… Values + Add onto D1:H1 from a single cell (10)
    await setValues(p, 'J1', [[10]])
    await activate(p, 'J1')
    await p.locator('[data-tip^="Copy"]').first().click()
    await p.waitForTimeout(400)
    await activate(p, 'D1:H1')
    await p.locator('[data-ribbon-body] [aria-label="Paste Special"]').first().click()
    await p
      .locator('.menu-select-drop button', { hasText: /^Paste Special…$/ })
      .first()
      .click()
    const paste = p.locator('.paste-special-dialog')
    await paste.getByRole('radio', { name: 'Values', exact: true }).check()
    await paste.getByRole('radio', { name: 'Add' }).check()
    await paste.getByRole('button', { name: 'OK' }).click()
    await expect.poll(() => raw(p, 'D1:H1')).toEqual([[11, 13, 15, 17, 19]])

    const zip = await save(launched, p, file, true)
    const xml = await sheetXml(zip)
    for (const [ref, value] of [
      ['A1', 1],
      ['A3', 5],
      ['A5', 9],
      ['B1', 42],
      ['D1', 11],
      ['H1', 19],
    ] as const) {
      expect(xml).toMatch(new RegExp(`<c r="${ref}"[^>]*><v>${value}</v></c>`))
    }
  } finally {
    await keep(file, 'series.xlsx')
    await kill(launched)
  }

  // reopen the saved file in a fresh session: everything is really in the grid
  const second = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-home-gaps-reopen',
    openFile: file,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const q = await waitForPageWithUrl(second.app, '://sheets/', 30_000)
    await q.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
    await expect.poll(() => raw(q, 'A1:A5'), { timeout: 30_000 }).toEqual([[1], [3], [5], [7], [9]])
    expect(await raw(q, 'D1:H1')).toEqual([[11, 13, 15, 17, 19]])
    expect(await raw(q, 'B1:B3')).toEqual([[42], [42], [42]])
  } finally {
    await closeAndSaveVideo(second, 'cleanup')
    await rm(dir, { recursive: true, force: true })
  }
})

test('the Delete key on link cells takes the links with the text, as Clear Contents does; ⌘Z brings both back', async () => {
  const { launched, p } = await openSheets()
  const dir = await mkdtemp(join(tmpdir(), 'genoffice-delete-links-'))
  const file = join(dir, 'delete-links.xlsx')
  try {
    await setValues(p, 'A1:A3', [['one'], ['two'], ['three']])
    await addLink(p, 'A1', 'https://example.com/1')
    await addLink(p, 'A2', 'https://example.com/2')
    await addLink(p, 'A3', 'https://example.com/3')
    let zip = await save(launched, p, file, true)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A1', 'A2', 'A3'])

    // Excel (COM): ClearContents removes the link, typing over a link cell keeps it
    await activate(p, 'A2:A3')
    await p.keyboard.press('Delete')
    await expect.poll(() => raw(p, 'A1:A3')).toEqual([['one'], [null], [null]])
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A1'])

    // one ⌘Z: the text and the links come back together, and the next save writes them
    await undo(p)
    await expect.poll(() => raw(p, 'A1:A3')).toEqual([['one'], ['two'], ['three']])
    zip = await save(launched, p, file)
    expect(hyperlinkRefs(await sheetXml(zip))).toEqual(['A1', 'A2', 'A3'])
  } finally {
    await keep(file, 'delete-links.xlsx')
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})
