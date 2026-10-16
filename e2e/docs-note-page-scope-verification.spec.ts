import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_DOCS_NOTE_PAGE_SCOPE_VERIFICATION,
  'Opt-in Docs per-page note preview and section scope verification',
)

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const notePart =
  `<w:footnotes ${W_NS}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  [2, 3, 4, 5]
    .map(
      (id) =>
        `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>NOTE ${id}</w:t></w:r></w:p></w:footnote>`,
    )
    .join('') +
  '</w:footnotes>'

async function parts(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    body: await zip.file('word/document.xml')!.async('string'),
    settings: await zip.file('word/settings.xml')!.async('string'),
    notes: await zip.file('word/footnotes.xml')!.async('string'),
  }
}

async function marks(page: Page, expected: string[]) {
  const refs = page.locator('.editor-scroll .ProseMirror .doc-note-ref[data-note-kind="footnote"]')
  await expect(refs).toHaveCount(expected.length)
  for (const [index, value] of expected.entries()) await expect(refs.nth(index)).toHaveText(value)
}

async function dialog(page: Page) {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^References$/ }).click()
  const group = page.locator('.ribbon-body .rb-fold-group[aria-label="Footnotes"]')
  const command = group.getByRole('button', { name: 'Note Settings…', exact: true })
  if (!(await command.isVisible())) await group.locator('button.rb-fold-btn').click()
  await command.click()
  const modal = page.getByRole('dialog', { name: 'Footnote and Endnote Settings' })
  await expect(modal).toBeVisible()
  return modal
}

test('actual-page restart and current-section edits survive history, native save and reopen', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-note-page-scope-'))
  const file = join(scratch, 'page-scope.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  const ref = (id: number) => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`
  await writeFile(
    file,
    await buildDocx({
      bodyXml:
        `<w:p><w:r><w:t>SECTION_ONE_A</w:t></w:r>${ref(2)}</w:p>` +
        `<w:p><w:r><w:t>SECTION_ONE_B</w:t></w:r>${ref(3)}</w:p>` +
        `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>SECTION_ONE_PAGE_TWO</w:t></w:r>${ref(4)}</w:p>` +
        '<w:p><w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/><w:footnote w:id="0"/></w:footnotePr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>SECTION_ONE_END</w:t></w:r></w:p>' +
        `<w:p><w:r><w:t>SECTION_TWO</w:t></w:r>${ref(5)}</w:p>`,
      sectPrExtra:
        '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="8"/><w:numRestart w:val="eachSect"/><w:footnote w:id="1"/></w:footnotePr>',
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      extraParts: [
        {
          path: 'word/settings.xml',
          xml: `<w:settings ${W_NS}><w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr></w:settings>`,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        },
        {
          path: 'word/footnotes.xml',
          xml: notePart,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
      ],
    }),
  )
  const original = await parts(file)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-note-page-scope',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    await marks(page, ['III', 'IV', 'III', 'h'])
    await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^View$/ }).click()
    const views = page.locator('.ribbon-body .rb-fold-group[aria-label="Views"]')
    const preview = views.getByRole('button', { name: 'Page Preview', exact: true })
    if (!(await preview.isVisible())) await views.locator('button.rb-fold-btn').click()
    await preview.click()
    await expect(page.locator('.pv-page')).toHaveCount(3)
    await expect(
      page.locator('.pv-page[data-pv-page="0"] .pv-footnotes .pv-footnote sup'),
    ).toHaveText(['III', 'IV'])
    await expect(
      page.locator('.pv-page[data-pv-page="1"] .pv-footnotes .pv-footnote sup'),
    ).toHaveText(['III'])
    await expect(
      page.locator('.pv-page[data-pv-page="2"] .pv-footnotes .pv-footnote sup'),
    ).toHaveText(['h'])
    await page.locator('.pagination-preview .pv-close', { hasText: 'Close' }).click()

    await page.locator('.editor-scroll .ProseMirror').click()
    await page.keyboard.press('ControlOrMeta+a')
    let modal = await dialog(page)
    await expect(
      modal.getByLabel('Apply changes to').locator('option[value="section"]'),
    ).toHaveAttribute('disabled', '')
    await modal.getByRole('button', { name: 'Cancel' }).click()
    await page.locator('.editor-scroll .ProseMirror .doc-note-ref').first().click()
    modal = await dialog(page)
    await modal.getByLabel('Apply changes to').selectOption('section')
    await expect(modal.getByLabel('Number format')).toHaveValue('upperRoman')
    await expect(modal.getByLabel('Numbering')).toHaveValue('eachPage')
    await modal.getByLabel('Number format').selectOption('decimal')
    await modal.getByLabel('Start at').fill('5')
    await modal.getByRole('button', { name: 'OK' }).click()
    await marks(page, ['5', '6', '5', 'h'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).body).toContain('<w:numStart w:val="5"/>')
    let saved = await parts(file)
    expect(saved.settings).toBe(original.settings)
    expect(saved.notes).toBe(original.notes)
    expect(saved.body).toContain('<w:numFmt w:val="decimal"/><w:numStart w:val="5"/>')
    expect(saved.body).toContain('<w:numFmt w:val="lowerLetter"/><w:numStart w:val="8"/>')
    for (const id of [2, 3, 4, 5])
      expect(saved.body).toContain(`<w:footnoteReference w:id="${id}"/>`)

    await page.keyboard.press('ControlOrMeta+z')
    await marks(page, ['III', 'IV', 'III', 'h'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).body).not.toContain('<w:numStart w:val="5"/>')
    saved = await parts(file)
    expect(saved.body).toContain('<w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/>')
    expect(saved.settings).toBe(original.settings)
    expect(saved.notes).toBe(original.notes)

    await page.keyboard.press('ControlOrMeta+Shift+z')
    await marks(page, ['5', '6', '5', 'h'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).body).toContain('<w:numStart w:val="5"/>')
    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-note-page-scope-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await marks(page, ['5', '6', '5', 'h'])
    await page.locator('.editor-scroll .ProseMirror .doc-note-ref').first().click()
    modal = await dialog(page)
    await modal.getByLabel('Apply changes to').selectOption('section')
    await expect(modal.getByLabel('Number format')).toHaveValue('decimal')
    await expect(modal.getByLabel('Start at')).toHaveValue('5')
    await modal.getByRole('button', { name: 'Cancel' }).click()

    await page.locator('.editor-scroll .ProseMirror .doc-note-ref').last().click()
    modal = await dialog(page)
    await modal.getByLabel('Apply changes to').selectOption('section')
    await expect(modal.getByLabel('Number format')).toHaveValue('lowerLetter')
    await modal.getByLabel('Number format').selectOption('upperLetter')
    await modal.getByLabel('Start at').fill('9')
    await modal.getByRole('button', { name: 'OK' }).click()
    await marks(page, ['5', '6', '5', 'I'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).body).toContain('<w:numStart w:val="9"/>')
    saved = await parts(file)
    expect(saved.body).toContain('<w:numFmt w:val="decimal"/><w:numStart w:val="5"/>')
    expect(saved.body).toContain('<w:numFmt w:val="upperLetter"/><w:numStart w:val="9"/>')
    expect(saved.settings).toBe(original.settings)
    expect(saved.notes).toBe(original.notes)
    const evidenceDir = process.env.GENOFFICE_DOCS_NOTE_PAGE_SCOPE_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'page-scope-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'page-scope-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})
