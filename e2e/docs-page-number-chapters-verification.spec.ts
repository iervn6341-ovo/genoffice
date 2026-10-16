import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx, IMAGE_PARAGRAPH_XML } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_DOCS_PAGE_NUMBER_CHAPTERS_VERIFICATION,
  'Opt-in Docs page numbering and chapter verification',
)

const headerXml =
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:p><w:r><w:t>Page </w:t></w:r>' +
  '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
  '<w:r><w:t> of </w:t></w:r>' +
  '<w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>2</w:t></w:r></w:fldSimple>' +
  '<w:r><w:t> • </w:t></w:r>' +
  '<w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple>' +
  '</w:p></w:hdr>'

async function documentParts(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    body: await zip.file('word/document.xml')!.async('string'),
    header: await zip.file('word/header1.xml')!.async('string'),
  }
}

async function openPageFormat(page: Page) {
  const insertTab = page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^Insert$/ })
  await insertTab.click()
  await page.locator('.ribbon-body button[data-tip="Page Number"]').click()
  await page.locator('[data-rb-panel] button', { hasText: /Format/ }).click()
  const dialog = page.locator('.modal', {
    has: page.getByRole('heading', { name: 'Page Number Format' }),
  })
  await expect(dialog).toBeVisible()
  return dialog
}

test('page format and numbered chapter survive Save, Undo, Redo and fresh reopen', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-page-number-chapters-'))
  const file = join(scratch, 'page-number-chapters.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  const sectionBreak =
    '<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rId20"/>' +
    '<w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>' +
    '</w:sectPr></w:pPr></w:p>'
  await writeFile(
    file,
    await buildDocx({
      withNumbering: true,
      withImage: true,
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>' +
        '<w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="0"/>' +
        '<w:numId w:val="2"/></w:numPr><w:outlineLvl w:val="2"/></w:pPr></w:style>',
      bodyXml:
        '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>CHAPTER_TARGET</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>First section body</w:t></w:r></w:p>' +
        sectionBreak +
        '<w:p><w:r><w:t>SECOND_SECTION_TARGET</w:t></w:r></w:p>' +
        IMAGE_PARAGRAPH_XML,
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
      extraRels:
        '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: headerXml,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
      ],
    }),
  )

  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-page-number-chapters',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    await page
      .locator('.doc-page.ProseMirror p')
      .filter({ hasText: 'SECOND_SECTION_TARGET' })
      .click()
    let dialog = await openPageFormat(page)
    await dialog.getByRole('button', { name: 'Number format' }).click()
    await dialog.locator('.gs-dd-item[data-value="decimalEnclosedCircle"]').click()
    await dialog.getByLabel('Include chapter number').check()
    await dialog.getByRole('button', { name: 'Chapter starts with style' }).click()
    await dialog.locator('.gs-dd-item[data-value="3"]').click()
    await dialog.getByRole('button', { name: 'Separator' }).click()
    await dialog.locator('.gs-dd-item[data-value="period"]').click()
    await dialog.getByText('Start at', { exact: true }).click()
    await dialog.locator('input[type="number"]').fill('7')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => document.activeElement?.classList.contains('ProseMirror')))
      .toBe(true)
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await documentParts(file)).body)
      .toContain('w:fmt="decimalEnclosedCircle"')
    let parts = await documentParts(file)
    expect(parts.body).toMatch(
      /<w:pgNumType\b[^>]*w:start="7"[^>]*w:chapStyle="3"[^>]*w:chapSep="period"/,
    )
    expect(parts.header).toBe(headerXml)
    expect(parts.body).toContain('<w:drawing>')

    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await documentParts(file)).body).not.toContain('w:chapStyle="3"')
    expect((await documentParts(file)).header).toBe(headerXml)

    await page.keyboard.press('ControlOrMeta+Shift+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await documentParts(file)).body).toContain('w:chapStyle="3"')
    parts = await documentParts(file)
    expect(parts.header).toBe(headerXml)
    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-page-number-chapters-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await page
      .locator('.doc-page.ProseMirror p')
      .filter({ hasText: 'SECOND_SECTION_TARGET' })
      .click()
    dialog = await openPageFormat(page)
    await expect(dialog.getByLabel('Include chapter number')).toBeChecked()
    await expect(dialog.getByRole('button', { name: 'Chapter starts with style' })).toContainText(
      'Heading 3',
    )
    await expect(dialog.locator('input[type="number"]')).toHaveValue('7')
    const evidenceDir = process.env.GENOFFICE_DOCS_PAGE_NUMBER_CHAPTERS_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'page-number-chapters-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'page-number-chapters-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})
