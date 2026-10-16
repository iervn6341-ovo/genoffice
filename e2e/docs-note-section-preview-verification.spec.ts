import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_DOCS_NOTE_SECTION_PREVIEW_VERIFICATION,
  'Opt-in Docs note section preview verification',
)

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
const notePart =
  `<w:footnotes ${W_NS}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>FIRST NOTE</w:t></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="3"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>SECOND NOTE</w:t></w:r></w:p></w:footnote>' +
  '</w:footnotes>'

async function xmlParts(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    body: await zip.file('word/document.xml')!.async('string'),
    settings: await zip.file('word/settings.xml')!.async('string'),
    footnotes: await zip.file('word/footnotes.xml')!.async('string'),
  }
}

async function settings(page: Page) {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^References$/ }).click()
  const group = page.locator('.ribbon-body .rb-fold-group[aria-label="Footnotes"]')
  const command = group.getByRole('button', { name: 'Note Settings…', exact: true })
  if (!(await command.isVisible())) await group.locator('button.rb-fold-btn').click()
  await command.click()
  const dialog = page.getByRole('dialog', { name: 'Footnote and Endnote Settings' })
  await expect(dialog).toBeVisible()
  return dialog
}

async function marks(page: Page, expected: [string, string]) {
  const references = page.locator(
    '.editor-scroll .ProseMirror .doc-note-ref[data-note-kind="footnote"]',
  )
  await expect(references).toHaveCount(2)
  await expect(references.nth(0)).toHaveText(expected[0])
  await expect(references.nth(1)).toHaveText(expected[1])
}

test('section-owned native footnote references keep their formats across edit, history and reopen', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-note-section-preview-'))
  const file = join(scratch, 'notes-by-section.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  await writeFile(
    file,
    await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="upperRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachSect"/><w:footnote w:id="0"/></w:footnotePr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:r><w:t>FIRST_SECTION</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>' +
        '<w:p><w:r><w:t>SECOND_SECTION</w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>' +
        '<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple></w:p>',
      sectPrExtra:
        '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:numStart w:val="8"/><w:numRestart w:val="eachSect"/><w:footnote w:id="1"/></w:footnotePr>',
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      extraParts: [
        {
          path: 'word/settings.xml',
          xml: `<w:settings ${W_NS}><w:footnotePr><w:numFmt w:val="lowerRoman"/><w:footnote w:id="-1"/></w:footnotePr></w:settings>`,
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
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-note-section-preview',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    await marks(page, ['III', 'h'])
    let dialog = await settings(page)
    await dialog.getByLabel('Number format').selectOption('decimal')
    await dialog.getByLabel('Start at').fill('4')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await marks(page, ['4', '4'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await xmlParts(file)).settings)
      .toContain('<w:numStart w:val="4"/>')
    let saved = await xmlParts(file)
    expect(saved.body).toContain('<w:footnoteReference w:id="2"/>')
    expect(saved.body).toContain('<w:footnoteReference w:id="3"/>')
    expect(saved.body).toContain('<w:fldSimple w:instr=" DATE ">')
    expect(saved.body.match(/<w:numFmt w:val="decimal"\/>/g)).toHaveLength(2)
    expect(saved.footnotes).toContain(notePart)

    await page.keyboard.press('ControlOrMeta+z')
    await marks(page, ['III', 'h'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await xmlParts(file)).settings)
      .not.toContain('<w:numStart w:val="4"/>')
    saved = await xmlParts(file)
    expect(saved.body).toContain('<w:numFmt w:val="upperRoman"/>')
    expect(saved.body).toContain('<w:numFmt w:val="lowerLetter"/>')
    expect(saved.footnotes).toContain(notePart)

    await page.keyboard.press('ControlOrMeta+Shift+z')
    await marks(page, ['4', '4'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await xmlParts(file)).settings)
      .toContain('<w:numStart w:val="4"/>')
    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-note-section-preview-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await marks(page, ['4', '4'])
    dialog = await settings(page)
    await expect(dialog.getByLabel('Number format')).toHaveValue('decimal')
    await expect(dialog.getByLabel('Start at')).toHaveValue('4')
    await dialog.getByLabel('Number format').selectOption('upperRoman')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await marks(page, ['IV', 'IV'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await xmlParts(file)).settings)
      .toContain('<w:numFmt w:val="upperRoman"/>')
    dialog = await settings(page)
    await dialog.getByLabel('Numbering').selectOption('eachPage')
    await expect(dialog.getByLabel('Numbering')).toHaveValue('eachPage')
    await dialog.getByRole('tab', { name: 'Endnote' }).click()
    await dialog.getByLabel('Position').selectOption('sectEnd')
    await expect(dialog.getByLabel('Position')).toHaveValue('sectEnd')
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    const evidenceDir = process.env.GENOFFICE_DOCS_NOTE_SECTION_PREVIEW_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'note-section-preview-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'note-section-preview-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})

test('genuine Word Mac section settings preserve untouched Office XML through edit and reopen', async () => {
  test.setTimeout(240_000)
  const source = resolve(
    __dirname,
    '../docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/word-mac-two-section-notes-source.docx',
  )
  const scratch = await mkdtemp(join(tmpdir(), 'docs-word-section-notes-'))
  const file = join(scratch, 'word-section-notes.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  await copyFile(source, file)
  const original = await xmlParts(file)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-word-section-notes',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    // Both references in this Word-authored specimen happen to be in section 2.
    await marks(page, ['h', 'i'])
    let dialog = await settings(page)
    await expect(dialog.getByLabel('Number format')).toHaveValue('lowerLetter')
    await expect(dialog.getByLabel('Start at')).toHaveValue('8')
    await dialog.getByLabel('Number format').selectOption('decimal')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await marks(page, ['8', '9'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await xmlParts(file)).settings)
      .toContain('<w:numFmt w:val="decimal"/>')
    let saved = await xmlParts(file)
    expect(saved.footnotes).toBe(original.footnotes)
    expect(saved.body).toContain('<w:numStart w:val="3"/>')
    expect(saved.body).toContain('<w:numStart w:val="8"/>')
    expect(saved.body).toContain('<w:footnoteReference w:id="1"/>')
    expect(saved.body).toContain('<w:footnoteReference w:id="2"/>')
    expect(
      saved.settings.replace('<w:numFmt w:val="decimal"/>', '<w:numFmt w:val="upperRoman"/>'),
    ).toBe(original.settings)

    await page.keyboard.press('ControlOrMeta+z')
    await marks(page, ['h', 'i'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await xmlParts(file)).settings).toBe(original.settings)
    saved = await xmlParts(file)
    expect(saved.footnotes).toBe(original.footnotes)

    await page.keyboard.press('ControlOrMeta+Shift+z')
    await marks(page, ['8', '9'])
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await xmlParts(file)).settings)
      .toContain('<w:numFmt w:val="decimal"/>')
    await closeAndSaveVideo(launched, 'word-first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-word-section-notes-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await marks(page, ['8', '9'])
    dialog = await settings(page)
    await expect(dialog.getByLabel('Number format')).toHaveValue('decimal')
    await expect(dialog.getByLabel('Start at')).toHaveValue('8')
    const evidenceDir = process.env.GENOFFICE_DOCS_NOTE_SECTION_PREVIEW_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'word-mac-two-section-notes-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'word-mac-two-section-notes-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'word-final-session')
  }
})
