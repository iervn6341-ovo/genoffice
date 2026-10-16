import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx, TINY_PNG_BASE64 } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

const VERIFY = process.env.GENOFFICE_DOCS_HF_AUTHORING_VERIFICATION
test.skip(!VERIFY, 'Opt-in Docs native header/footer authoring verification')

const headerType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const importedField =
  '<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Protected Author</w:t></w:r></w:fldSimple>'
const importedPicture =
  '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">' +
  '<wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="Imported logo"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill>' +
  '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>' +
  '</pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
const headerXml = (text: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>` +
  '</w:hdr>'

async function currentSecondHeader(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  const documentXml = await zip.file('word/document.xml')!.async('string')
  const docRels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  const sections = [...documentXml.matchAll(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)]
  const ref = sections[1]?.[0].match(/<w:headerReference\b(?=[^>]*w:type="default")[^>]*\/>/)?.[0]
  const rId = ref?.match(/r:id="([^"]+)"/)?.[1]
  const target = rId
    ? docRels
        .match(new RegExp(`<Relationship\\b(?=[^>]*Id="${rId}")[^>]*/>`))?.[0]
        .match(/Target="([^"]+)"/)?.[1]
    : null
  if (!target) throw new Error('Section 2 header relationship is missing')
  const part = target.startsWith('/') ? target.slice(1) : `word/${target}`
  const xml = await zip.file(part)!.async('string')
  const rels = (await zip.file(part.replace(/([^/]+)$/, '_rels/$1.rels'))?.async('string')) ?? ''
  return { zip, xml, rels }
}

async function openHeaderEditor(page: Page) {
  const tab = page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^Insert$/ })
  if (!(await tab.evaluate((element) => element.classList.contains('active')))) await tab.click()
  await page
    .locator('.ribbon-body button', { hasText: /^Header$/ })
    .first()
    .click()
  const modal = page.locator('.hf-menu')
  await expect(modal.locator('textarea')).toBeVisible()
  return modal
}

test('native fields, picture, template and section navigation survive save history and reopen', async () => {
  test.setTimeout(300_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-hf-authoring-'))
  const file = join(scratch, 'authoring.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  const evidenceDir = resolve(
    __dirname,
    '../docs/qa/office-mac-field-evidence-2026-10-13/docs-authoring',
  )
  await mkdir(evidenceDir, { recursive: true })
  const sectionBreak =
    '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId20"/>' +
    '<w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>' +
    '</w:sectPr></w:pPr></w:p>'
  const bytes = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>Section one body</w:t></w:r></w:p>' +
      sectionBreak +
      '<w:p><w:r><w:t>Section two body</w:t></w:r></w:p>',
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId21"/>',
    extraRels:
      '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
      '<Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>',
    extraParts: [
      {
        path: 'docProps/core.xml',
        xml: '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>HF_USER_AUTHOR</dc:creator></cp:coreProperties>',
        contentType: 'application/vnd.openxmlformats-package.core-properties+xml',
      },
      { path: 'word/header1.xml', xml: headerXml('First physical page'), contentType: headerType },
      {
        path: 'word/header2.xml',
        xml: headerXml('Second section').replace(
          '</w:hdr>',
          Array.from(
            { length: 8 },
            (_, index) => `<w:p><w:r><w:t>Tall header line ${index + 1}</w:t></w:r></w:p>`,
          ).join('') + `<w:p>${importedField}</w:p>${importedPicture}</w:hdr>`,
        ),
        contentType: headerType,
      },
      {
        path: 'word/_rels/header2.xml.rels',
        xml: '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/original.png"/></Relationships>',
        contentType: 'application/vnd.openxmlformats-package.relationships+xml',
      },
    ],
    binaryParts: [
      {
        path: 'word/media/original.png',
        base64: TINY_PNG_BASE64,
        extension: 'png',
        contentType: 'image/png',
      },
    ],
  })
  await writeFile(file, bytes)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-hf-authoring',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    await launched.app.evaluate(({ ipcMain }, base64) => {
      ipcMain.removeHandler('docs:pick-image')
      ipcMain.handle('docs:pick-image', () => ({ base64, mime: 'image/png', name: 'logo.png' }))
    }, TINY_PNG_BASE64)
    const topHeader = page.locator('.page-hf-header:not(.page-gap-hf)').first()
    await expect(topHeader).toContainText('First physical page')
    await page
      .locator('.hf-variant-chips')
      .getByRole('button', { name: 'Next', exact: true })
      .click()
    await expect(page.locator('.hf-variant-chips')).toContainText('Section 2')
    await expect(topHeader).toContainText('First physical page')
    await topHeader.dblclick()
    await expect(topHeader.locator('.page-hf-edit-surface')).toHaveCount(0)
    await page
      .locator('.hf-variant-chips')
      .getByRole('button', { name: 'Previous', exact: true })
      .click()
    await page
      .locator('.hf-variant-chips')
      .getByRole('button', { name: 'Next', exact: true })
      .click()
    const sectionBody = page
      .locator('.doc-page.ProseMirror [data-idx]', { hasText: 'Section two body' })
      .first()
    await expect
      .poll(async () => {
        const box = await sectionBody.boundingBox()
        const height = await page.evaluate(() => window.innerHeight)
        return Boolean(box && box.y > 0 && box.y < height)
      })
      .toBe(true)
    const sectionHeader = page
      .locator('[data-testid="hf-page-gap-header"][data-hf-section="2"]')
      .first()
    await expect(sectionHeader).toContainText('Tall header line 8')
    await expect
      .poll(async () => {
        const headerLastLine = await sectionHeader.locator('.page-hf-para').last().boundingBox()
        const bodyBox = await sectionBody.boundingBox()
        return headerLastLine && bodyBox
          ? bodyBox.y - headerLastLine.y - headerLastLine.height
          : -Infinity
      })
      .toBeGreaterThanOrEqual(-1)

    let modal = await openHeaderEditor(page)
    const textarea = modal.locator('textarea')
    await textarea.focus()
    await textarea.press('End')
    for (const format of ['yyyy-MM-dd', 'dd/MM/yyyy', 'yyyy/MM/dd', 'MM/dd/yyyy']) {
      await modal.getByRole('combobox', { name: 'Date' }).selectOption(format)
      await expect(modal.getByRole('combobox', { name: 'Date' })).toHaveValue(format)
    }
    await modal.getByRole('button', { name: 'Fixed date', exact: true }).click()
    await modal.getByRole('button', { name: 'Date', exact: true }).click()
    await modal.getByRole('button', { name: 'File name', exact: true }).click()
    await modal.getByRole('button', { name: 'Page', exact: true }).click()
    await modal.getByRole('button', { name: 'Total pages', exact: true }).click()
    await modal.getByRole('button', { name: 'Author', exact: true }).click()
    await modal.getByRole('button', { name: 'Picture', exact: true }).click()
    await modal.getByRole('button', { name: 'Date and page', exact: true }).click()
    await modal.getByRole('button', { name: 'OK', exact: true }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await currentSecondHeader(file)).rels)
      .toMatch(/media\/aidocs\d+\.png/)
    let saved = await currentSecondHeader(file)
    expect(saved.xml).toContain('Second section')
    expect(saved.xml).toContain(importedField)
    expect(saved.xml).toContain(importedPicture)
    expect(saved.xml).toContain(' DATE \\@ &quot;MM/dd/yyyy&quot; ')
    expect(saved.xml).toContain(' FILENAME ')
    expect(saved.xml).toContain(' PAGE ')
    expect(saved.xml).toContain(' NUMPAGES ')
    expect(saved.xml).toContain('HF_USER_AUTHOR')
    expect(saved.xml).toContain('<w:tabs>')
    expect(saved.xml).toContain('<wp:extent')
    expect(saved.rels).toMatch(/Target="media\/aidocs\d+\.png"/)
    expect(saved.rels).toContain('Target="media/original.png"')
    expect(
      Object.keys(saved.zip.files).some((name) => /^word\/media\/aidocs\d+\.png$/.test(name)),
    ).toBe(true)

    const undo = page.locator('.qa-btn[aria-label="Undo"]')
    const redo = page.locator('.qa-btn[aria-label="Redo"]')
    await undo.click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await currentSecondHeader(file)).rels)
      .not.toMatch(/media\/aidocs\d+\.png/)
    await redo.click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await currentSecondHeader(file)).rels)
      .toMatch(/media\/aidocs\d+\.png/)
    await closeAndSaveVideo(launched, 'docs-hf-authoring-before-reopen')
    launched = undefined

    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-hf-authoring-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await page
      .locator('.hf-variant-chips')
      .getByRole('button', { name: 'Next', exact: true })
      .click()
    modal = await openHeaderEditor(page)
    const reopenedText = await modal.locator('textarea').inputValue()
    expect(reopenedText).toMatch(/⟦DATE·\d+⟧/)
    expect(reopenedText).toMatch(/⟦FILENAME·\d+⟧/)
    await modal.locator('textarea').fill(reopenedText.replace('Second section', 'Reedited section'))
    await modal.getByRole('button', { name: 'OK', exact: true }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await currentSecondHeader(file)).xml)
      .toContain('Reedited section')
    saved = await currentSecondHeader(file)
    expect(saved.xml).toContain('<w:drawing>')
    expect(saved.xml).toContain(importedField)
    expect(saved.xml).toContain(importedPicture)
    expect(saved.xml).toContain(' FILENAME ')
    expect(saved.rels).toContain('/image')
    await expect(
      page.locator('[data-testid="hf-page-gap-header"][data-hf-section="2"]').first(),
    ).toContainText('HF_USER_AUTHOR')
    const renamed = join(scratch, 'authoring-renamed.docx')
    await launched.app.evaluate(({ dialog }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
    }, renamed)
    await page.keyboard.press('ControlOrMeta+Shift+s')
    await expect.poll(async () => (await currentSecondHeader(renamed)).xml).toContain(' FILENAME ')
    await expect(
      page.locator('[data-testid="hf-page-gap-header"][data-hf-section="2"]').first(),
    ).toContainText('authoring-renamed.docx')
    await page.screenshot({ path: join(evidenceDir, 'reopened-edit.png') })
    await copyFile(renamed, join(evidenceDir, 'authoring-native.docx'))
    const pdfPath = join(scratch, 'authoring-native.pdf')
    await launched.app.evaluate(({ dialog, shell }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
      shell.openPath = (async () => '') as never
    }, pdfPath)
    await page.getByRole('button', { name: 'File', exact: true }).click()
    await page.getByRole('button', { name: 'Export', exact: true }).click()
    await page.locator('.bs-exports').getByRole('button', { name: /PDF/ }).click()
    await expect
      .poll(
        async () =>
          (await readFile(pdfPath).catch(() => Buffer.alloc(0))).subarray(0, 5).toString(),
        { timeout: 30_000 },
      )
      .toBe('%PDF-')
    await copyFile(pdfPath, join(evidenceDir, 'authoring-native.pdf'))
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'docs-hf-authoring')
  }
})
