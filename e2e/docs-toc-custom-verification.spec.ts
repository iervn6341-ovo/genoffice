import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx, IMAGE_PARAGRAPH_XML } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(!process.env.GENOFFICE_DOCS_TOC_VERIFICATION, 'Opt-in Docs custom TOC verification')

async function documentXml(path: string): Promise<string> {
  return (await JSZip.loadAsync(await readFile(path))).file('word/document.xml')!.async('string')
}

async function showReferences(page: Page): Promise<void> {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^References$/ }).click()
}

async function updateToc(page: Page, mode: 'pages' | 'full'): Promise<void> {
  await showReferences(page)
  await page
    .locator(
      '.ribbon-body button[data-tip="Update the table of contents to match the current headings"]',
    )
    .click()
  const dialog = page.getByRole('dialog', { name: 'Update Table' })
  await expect(dialog).toBeVisible()
  await dialog
    .getByLabel(mode === 'pages' ? 'Update page numbers only' : 'Update entire table')
    .check()
  await dialog.getByRole('button', { name: 'Update Table' }).click()
}

test('custom TOC settings and page-only/full updates retain native fields through Save, Undo, Redo and reopen', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-custom-toc-'))
  const file = join(scratch, 'custom-toc.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  const para = (style: string, text: string) =>
    `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`
  await writeFile(
    file,
    await buildDocx({
      withImage: true,
      extraStylesXml:
        '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>' +
        '<w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/>' +
        '<w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="CustomTitle"><w:name w:val="Custom Title"/></w:style>',
      bodyXml:
        '<w:p/>' +
        para('Heading1', 'CHAPTER_ONE') +
        para('Heading2', 'SECTION_TWO') +
        para('Heading3', 'SUBSECTION_THREE') +
        para('Heading4', 'DETAIL_FOUR') +
        para('CustomTitle', 'CUSTOM_CHAPTER') +
        '<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple></w:p>' +
        IMAGE_PARAGRAPH_XML,
    }),
  )

  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-custom-toc',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    await page.locator('.doc-page.ProseMirror p').first().click()
    await showReferences(page)
    await page.locator('.ribbon-body button[data-tip="Custom table of contents"]').click()
    let dialog = page.getByRole('dialog', { name: 'Custom table of contents' })
    await expect(dialog).toBeVisible()
    await dialog.getByLabel('Format').selectOption('modern')
    await dialog.getByLabel('Show levels').fill('4')
    await dialog.getByLabel('Tab leader').selectOption('hyphen')
    await dialog.getByLabel('Search styles').fill('Custom Title')
    await dialog.getByLabel('Custom Title Level').selectOption('2')
    await dialog.getByRole('button', { name: 'Insert' }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(() => documentXml(file)).toContain('\\t "Custom Title,2"')
    let xml = await documentXml(file)
    expect(xml).toContain(' TOC \\o "1-4" \\t "Custom Title,2" \\h \\z \\u ')
    expect(xml).toContain('<w:tab w:val="right" w:leader="hyphen"')
    expect(xml).toContain('<w:color w:val="4472C4"/>')
    expect(xml).toContain('CUSTOM_CHAPTER')
    expect(xml).toContain('<w:fldSimple w:instr=" DATE ">')
    expect(xml).toContain('<w:drawing>')

    await page
      .locator('.doc-page.ProseMirror')
      .getByText('CHAPTER_ONE', { exact: true })
      .last()
      .click()
    await page.keyboard.press('End')
    await page.keyboard.type('_EDITED')
    await updateToc(page, 'pages')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(() => documentXml(file)).toContain('CHAPTER_ONE_EDITED')
    xml = await documentXml(file)
    expect((xml.match(/CHAPTER_ONE_EDITED/g) ?? []).length).toBe(1)
    expect((xml.match(/CHAPTER_ONE/g) ?? []).length).toBe(2)
    await updateToc(page, 'full')
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await documentXml(file)).match(/CHAPTER_ONE_EDITED/g)?.length)
      .toBe(2)
    expect((await documentXml(file)).match(/<w:sdt>/g)?.length ?? 0).toBe(0)

    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await documentXml(file)).match(/CHAPTER_ONE_EDITED/g)?.length)
      .toBe(1)
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await documentXml(file)).match(/CHAPTER_ONE_EDITED/g)?.length)
      .toBe(2)
    xml = await documentXml(file)
    expect(xml).toContain('<w:fldSimple w:instr=" DATE ">')
    expect(xml).toContain('<w:drawing>')

    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-custom-toc-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await showReferences(page)
    await page.locator('.ribbon-body button[data-tip="Custom table of contents"]').click()
    dialog = page.getByRole('dialog', { name: 'Custom table of contents' })
    await expect(dialog.getByLabel('Format')).toHaveValue('modern')
    await expect(dialog.getByLabel('Show levels')).toHaveValue('4')
    await expect(dialog.getByLabel('Tab leader')).toHaveValue('hyphen')
    await dialog.getByLabel('Search styles').fill('Custom Title')
    await expect(dialog.getByLabel('Custom Title Level')).toHaveValue('2')
    const evidenceDir = process.env.GENOFFICE_DOCS_TOC_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'custom-toc-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'custom-toc-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})
