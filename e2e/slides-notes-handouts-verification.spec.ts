import { expect, test, type Page } from '@playwright/test'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

const fixture = join(__dirname, '../packages/pptx-engine/tests/fixtures/01_standard_business.pptx')
const req = createRequire(join(process.cwd(), 'package.json'))
const header = 'NOTES_HEADER_NATIVE_2026'
const footer = 'HANDOUT_FOOTER_NATIVE_2026'

test.skip(
  !process.env.GENOFFICE_SLIDES_NOTES_HF_VERIFICATION,
  'Opt-in native Notes/Handouts verification',
)

async function scratch(): Promise<string> {
  const root = resolve(
    process.env.GENOFFICE_SLIDES_VERIFICATION_DIR ?? join(tmpdir(), 'genoffice-slides-notes-hf'),
  )
  await mkdir(root, { recursive: true })
  return mkdtemp(join(root, 'notes-handouts-'))
}

async function openHeaderFooter(page: Page) {
  const tab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: /^Insert$/ })
  await tab.click()
  await page.locator('.ribbon-body button[data-tip^="Header & footer"]').click()
  const dialog = page.locator('.modal', {
    has: page.getByRole('heading', { name: 'Header and Footer' }),
  })
  await expect(dialog).toBeVisible()
  return dialog
}

async function saveAs(launched: LaunchedApp, page: Page, path: string): Promise<void> {
  await launched.app.evaluate(({ dialog }, target) => {
    dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
  }, path)
  const result = await page.evaluate(() => window.slidesApi.saveAs('notes-handouts-native.pptx'))
  expect(result.ok).toBe(true)
}

async function openPrint(page: Page, launched: LaunchedApp): Promise<void> {
  await launched.app.evaluate(({ Menu, BrowserWindow }) => {
    const walk = (menu: Electron.Menu | null | undefined): Electron.MenuItem | undefined => {
      for (const item of menu?.items ?? []) {
        if (/^(CommandOrControl|CmdOrCtrl|Command|Cmd)\+P$/i.test(String(item.accelerator ?? '')))
          return item
        const nested = walk(item.submenu)
        if (nested) return nested
      }
    }
    walk(Menu.getApplicationMenu())?.click(
      undefined,
      BrowserWindow.getAllWindows()[0],
      undefined as never,
    )
  })
  await expect(page.locator('.print-dialog')).toBeVisible()
  await expect(page.locator('iframe[title="print-preview"]')).toBeVisible()
}

async function exportPrintPdf(
  page: Page,
  launched: LaunchedApp,
  layout: 'notes' | 'handout2',
  path: string,
): Promise<void> {
  await openPrint(page, launched)
  await page.getByTestId('print-layout').selectOption(layout)
  await expect
    .poll(() =>
      page
        .locator('iframe[title="print-preview"]')
        .evaluate(
          (frame: HTMLIFrameElement, expected: { header: string; footer: string }) =>
            frame.srcdoc.includes(expected.header) && frame.srcdoc.includes(expected.footer),
          { header, footer: `${footer}_EDITED` },
        ),
    )
    .toBe(true)
  await launched.app.evaluate(({ dialog }, target) => {
    dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
  }, path)
  await page.getByTestId('print-save-pdf').click()
  await expect
    .poll(
      async () => {
        try {
          return (await readFile(path)).subarray(0, 5).toString()
        } catch {
          return ''
        }
      },
      { timeout: 30_000 },
    )
    .toBe('%PDF-')
}

async function pdfText(path: string): Promise<string[]> {
  const pdfjs = await import(req.resolve('pdfjs-dist/legacy/build/pdf.mjs'))
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(await readFile(path)),
    disableFontFace: true,
    verbosity: 0,
  }).promise
  try {
    const pages: string[] = []
    for (let i = 1; i <= doc.numPages; i++) {
      const content = await (await doc.getPage(i)).getTextContent()
      pages.push(content.items.map((item: { str?: string }) => item.str ?? '').join(''))
    }
    return pages
  } finally {
    await doc.loadingTask.destroy()
  }
}

function fieldIds(xml: string): string[] {
  return [...xml.matchAll(/<a:fld\b[^>]*\bid="([^"]+)"[^>]*\btype="(?:datetime10|slidenum)"/g)].map(
    (match) => match[1]!,
  )
}

test('Slides Notes/Handouts native fields, start at zero, history, save/reopen, and layout PDFs', async () => {
  test.setTimeout(300_000)
  const dir = await scratch()
  const imported = join(dir, 'source.pptx')
  const saved = join(dir, 'notes-handouts-native.pptx')
  const notesPdf = join(dir, 'notes-pages.pdf')
  const handoutPdf = join(dir, 'handout-pages.pdf')
  await copyFile(fixture, imported)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: imported,
      videoDir: 'slides-notes-handouts-hf',
    })
    let page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
    await expect(page.locator('.slide-list .thumb')).toHaveCount(5)

    let dialog = await openHeaderFooter(page)
    await dialog.getByTestId('hf-slide-start-at').fill('0')
    await dialog.getByTestId('hf-slide-number').check()
    await expect(dialog.getByTestId('hf-preview-number')).toHaveText('0')
    await dialog.getByTestId('hf-apply-all').click()
    await expect(dialog).toHaveCount(0)

    dialog = await openHeaderFooter(page)
    await expect(dialog.getByTestId('hf-slide-start-at')).toHaveValue('0')
    await dialog.getByTestId('hf-tab-notes-handouts').click()
    await dialog.getByTestId('nh-header-enabled').check()
    await dialog.getByTestId('nh-header-text').fill(header)
    await dialog.getByTestId('nh-footer-enabled').check()
    await dialog.getByTestId('nh-footer-text').fill(footer)
    await dialog.getByTestId('nh-date-enabled').check()
    await dialog.getByTestId('nh-date-auto').check()
    await dialog.getByTestId('nh-date-format').selectOption('datetime10')
    await dialog.getByTestId('nh-page-number').check()
    await expect(dialog.getByTestId('nh-preview-header')).toHaveText(header)
    await expect(dialog.getByTestId('nh-preview-footer')).toHaveText(footer)
    await expect(dialog.getByTestId('nh-preview-number')).toHaveText('0')
    await dialog.getByTestId('nh-apply-all').click()
    await expect(dialog).toHaveCount(0)

    const undo = page.getByRole('button', { name: 'Undo', exact: true })
    const redo = page.getByRole('button', { name: 'Redo', exact: true })
    await undo.click()
    dialog = await openHeaderFooter(page)
    await dialog.getByTestId('hf-tab-notes-handouts').click()
    await expect(dialog.getByTestId('nh-header-enabled')).not.toBeChecked()
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await redo.click()
    dialog = await openHeaderFooter(page)
    await dialog.getByTestId('hf-tab-notes-handouts').click()
    await expect(dialog.getByTestId('nh-header-text')).toHaveValue(header)
    await expect(dialog.getByTestId('nh-date-format')).toHaveValue('datetime10')
    await dialog.getByRole('button', { name: 'Cancel' }).click()

    await saveAs(launched, page, saved)
    const before = await JSZip.loadAsync(await readFile(saved))
    const presentation = await before.file('ppt/presentation.xml')!.async('string')
    const notesMaster = await before.file('ppt/notesMasters/notesMaster1.xml')!.async('string')
    const handoutMaster = await before
      .file('ppt/handoutMasters/handoutMaster1.xml')!
      .async('string')
    expect(presentation).toContain('firstSlideNum="0"')
    for (const xml of [notesMaster, handoutMaster]) {
      expect(xml).toMatch(/<p:hf\b[^>]*hdr="1"[^>]*ftr="1"[^>]*dt="1"[^>]*sldNum="1"/)
      expect(xml).toContain('type="datetime10"')
      expect(xml).toContain('type="slidenum"')
      expect(xml).toContain(header)
      expect(xml).toContain(footer)
      expect(fieldIds(xml)).toHaveLength(2)
    }

    await undo.click()
    dialog = await openHeaderFooter(page)
    await dialog.getByTestId('hf-tab-notes-handouts').click()
    await expect(dialog.getByTestId('nh-header-enabled')).not.toBeChecked()
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    const undone = await JSZip.loadAsync(await readFile(saved))
    expect(
      (await undone.file('ppt/notesMasters/notesMaster1.xml')?.async('string')) ?? '',
    ).not.toContain(header)
    await redo.click()
    const savedAgain = await page.evaluate(() => window.slidesApi.save())
    expect(savedAgain.ok).toBe(true)
    const after = await JSZip.loadAsync(await readFile(saved))
    expect(
      fieldIds(await after.file('ppt/notesMasters/notesMaster1.xml')!.async('string')),
    ).toEqual(fieldIds(notesMaster))
    expect(
      fieldIds(await after.file('ppt/handoutMasters/handoutMaster1.xml')!.async('string')),
    ).toEqual(fieldIds(handoutMaster))

    await closeAndSaveVideo(launched, 'slides-notes-handouts-hf-before-reopen')
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      openFile: saved,
      videoDir: 'slides-notes-handouts-hf-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://slides/', 30_000)
    await page.waitForSelector('.slide-list .thumb', { timeout: 20_000 })
    dialog = await openHeaderFooter(page)
    await expect(dialog.getByTestId('hf-slide-start-at')).toHaveValue('0')
    await dialog.getByTestId('hf-tab-notes-handouts').click()
    await expect(dialog.getByTestId('nh-header-text')).toHaveValue(header)
    await expect(dialog.getByTestId('nh-footer-text')).toHaveValue(footer)
    await expect(dialog.getByTestId('nh-date-format')).toHaveValue('datetime10')
    await expect(dialog.getByTestId('nh-page-number')).toBeChecked()
    await dialog.getByTestId('nh-footer-text').fill(`${footer}_EDITED`)
    await dialog.getByTestId('nh-apply-all').click()
    expect((await page.evaluate(() => window.slidesApi.save())).ok).toBe(true)
    const reopenedEdited = await JSZip.loadAsync(await readFile(saved))
    for (const masterPath of [
      'ppt/notesMasters/notesMaster1.xml',
      'ppt/handoutMasters/handoutMaster1.xml',
    ]) {
      const xml = await reopenedEdited.file(masterPath)!.async('string')
      expect(xml).toContain(`${footer}_EDITED`)
      expect(fieldIds(xml)).toEqual(fieldIds(await before.file(masterPath)!.async('string')))
    }
    const evidenceDir = resolve(__dirname, '../docs/qa/office-mac-field-evidence-2026-10-13/slides')
    await mkdir(evidenceDir, { recursive: true })
    await page.screenshot({ path: join(evidenceDir, 'reopened-edit.png') })

    const effectiveFields = await page.evaluate(() =>
      window.slidesApi.getNotesHandoutsPrintFields(),
    )
    expect(effectiveFields.notesPages[0]?.footer?.text).toBe(`${footer}_EDITED`)
    expect(effectiveFields.handout.footer?.text).toBe(`${footer}_EDITED`)
    await exportPrintPdf(page, launched, 'notes', notesPdf)
    await exportPrintPdf(page, launched, 'handout2', handoutPdf)
    const notesText = await pdfText(notesPdf)
    const handoutText = await pdfText(handoutPdf)
    expect(notesText).toHaveLength(5)
    expect(handoutText).toHaveLength(3)
    expect(notesText.every((text) => text.includes(header) && text.includes(footer))).toBe(true)
    expect(handoutText.every((text) => text.includes(header) && text.includes(footer))).toBe(true)
    for (const [i, text] of notesText.entries()) {
      expect(text).toMatch(new RegExp(`\\d{1,2}:\\d{2} ${footer}_EDITED ${i}$`))
    }
    for (const [i, text] of handoutText.entries()) {
      expect(text).toMatch(new RegExp(`\\d{1,2}:\\d{2} ${footer}_EDITED ${i + 1}$`))
    }
    await copyFile(saved, join(evidenceDir, 'notes-handouts-native.pptx'))
    await copyFile(notesPdf, join(evidenceDir, 'notes-pages.pdf'))
    await copyFile(handoutPdf, join(evidenceDir, 'handout-pages.pdf'))
    await writeFile(
      join(evidenceDir, 'verification.json'),
      JSON.stringify(
        { saved, notesPdf, handoutPdf, notesText, handoutText, fieldIds: fieldIds(notesMaster) },
        null,
        2,
      ),
    )
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'slides-notes-handouts-hf-final')
  }
})
