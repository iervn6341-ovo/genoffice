import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_DOCS_NOTE_SETTINGS_VERIFICATION,
  'Opt-in Docs note settings verification',
)

async function parts(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    body: await zip.file('word/document.xml')!.async('string'),
    settings: await zip.file('word/settings.xml')!.async('string'),
    footnotes: await zip.file('word/footnotes.xml')!.async('string'),
    endnotes: await zip.file('word/endnotes.xml')!.async('string'),
  }
}

async function openSettings(page: Page) {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^References$/ }).click()
  const group = page.locator('.ribbon-body .rb-fold-group[aria-label="Footnotes"]')
  const command = group.getByRole('button', { name: 'Note Settings…', exact: true })
  if (!(await command.isVisible())) await group.locator('button.rb-fold-btn').click()
  await command.click()
  const dialog = page.getByRole('dialog', { name: 'Footnote and Endnote Settings' })
  await expect(dialog).toBeVisible()
  return dialog
}

test('whole-document note settings preserve native parts and distinct section overrides across save history', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-note-settings-'))
  const file = join(scratch, 'note-settings.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  const wNs = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
  const footnoteEntry =
    '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>FOOTNOTE_BODY</w:t></w:r></w:p></w:footnote>'
  const endnoteEntry =
    '<w:endnote w:id="3"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>ENDNOTE_BODY</w:t></w:r></w:p></w:endnote>'
  await writeFile(
    file,
    await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:sectPr><w:footnotePr><w:numFmt w:val="upperRoman"/><w:footnote w:id="0"/></w:footnotePr></w:sectPr></w:pPr><w:r><w:t>FIRST_SECTION</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>SECOND_SECTION</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:endnoteReference w:id="3"/></w:r></w:p>' +
        '<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple></w:p>',
      sectPrExtra:
        '<w:footnotePr><w:numFmt w:val="lowerLetter"/><w:footnote w:id="1"/></w:footnotePr>' +
        '<w:endnotePr><w:pos w:val="sectEnd"/><w:endnote w:id="0"/></w:endnotePr>',
      extraRels:
        '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>' +
        '<Relationship Id="rId41" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>',
      extraParts: [
        {
          path: 'word/settings.xml',
          xml: `<w:settings ${wNs}><w:footnotePr><w:numFmt w:val="lowerRoman"/><w:footnote w:id="-1"/></w:footnotePr><w:endnotePr><w:numFmt w:val="upperRoman"/><w:endnote w:id="-1"/></w:endnotePr></w:settings>`,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        },
        {
          path: 'word/footnotes.xml',
          xml: `<w:footnotes ${wNs}>${footnoteEntry}</w:footnotes>`,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
        },
        {
          path: 'word/endnotes.xml',
          xml: `<w:endnotes ${wNs}>${endnoteEntry}</w:endnotes>`,
          contentType:
            'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
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
      videoDir: 'docs-note-settings',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    let dialog = await openSettings(page)
    await dialog.getByLabel('Position').selectOption('beneathText')
    await dialog.getByLabel('Number format').selectOption('decimal')
    await dialog.getByLabel('Start at').fill('4')
    await dialog.getByLabel('Numbering').selectOption('eachSect')
    await dialog.getByRole('tab', { name: 'Endnote' }).click()
    await dialog.getByLabel('Position').selectOption('docEnd')
    await dialog.getByLabel('Number format').selectOption('lowerLetter')
    await dialog.getByLabel('Start at').fill('3')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(page.locator('.doc-note-ref').first()).toContainText('4')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).settings).toContain('<w:numStart w:val="4"/>')
    let xml = await parts(file)
    expect(xml.body).toContain('<w:fldSimple w:instr=" DATE ">')
    expect(xml.body).toContain('<w:footnoteReference w:id="2"/>')
    expect(xml.body).toContain('<w:endnoteReference w:id="3"/>')
    expect(xml.body).toContain('<w:footnote w:id="0"/>')
    expect(xml.body).toContain('<w:footnote w:id="1"/>')
    expect(xml.body).toContain('<w:endnote w:id="0"/>')
    expect(xml.footnotes).toContain(footnoteEntry)
    expect(xml.endnotes).toContain(endnoteEntry)
    expect(xml.settings).toContain('<w:numRestart w:val="eachSect"/>')
    expect(xml.settings).toContain('<w:numFmt w:val="lowerLetter"/>')

    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await parts(file)).settings)
      .toContain('<w:numFmt w:val="lowerRoman"/>')
    xml = await parts(file)
    expect(xml.body).toContain('<w:numFmt w:val="upperRoman"/>')
    expect(xml.body).toContain('<w:numFmt w:val="lowerLetter"/>')
    expect(xml.body).not.toContain('<w:numStart w:val="4"/>')
    expect(xml.footnotes).toContain(footnoteEntry)
    expect(xml.endnotes).toContain(endnoteEntry)

    await page.keyboard.press('ControlOrMeta+Shift+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).settings).toContain('<w:numStart w:val="4"/>')
    xml = await parts(file)
    expect(xml.body).toContain('<w:numFmt w:val="decimal"/>')
    expect(xml.body).toContain('<w:numStart w:val="4"/>')

    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-note-settings-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    dialog = await openSettings(page)
    await expect(dialog.getByLabel('Position')).toHaveValue('beneathText')
    await expect(dialog.getByLabel('Number format')).toHaveValue('decimal')
    await expect(dialog.getByLabel('Start at')).toHaveValue('4')
    await expect(dialog.getByLabel('Numbering')).toHaveValue('eachSect')
    await dialog.getByRole('tab', { name: 'Endnote' }).click()
    await expect(dialog.getByLabel('Position')).toHaveValue('docEnd')
    await expect(dialog.getByLabel('Start at')).toHaveValue('3')
    const evidenceDir = process.env.GENOFFICE_DOCS_NOTE_SETTINGS_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'note-settings-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'note-settings-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})

test('Word Mac note settings import keeps untouched Office XML through edit, history and reopen', async () => {
  test.setTimeout(240_000)
  const source = resolve(
    __dirname,
    '../docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/word-mac-footnote-settings-source.docx',
  )
  const scratch = await mkdtemp(join(tmpdir(), 'docs-word-note-settings-'))
  const file = join(scratch, 'word-note-settings.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  await copyFile(source, file)
  const original = await parts(file)
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-word-note-settings',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    let dialog = await openSettings(page)
    await expect(dialog.getByLabel('Position')).toHaveValue('beneathText')
    await expect(dialog.getByLabel('Number format')).toHaveValue('lowerRoman')
    await expect(dialog.getByLabel('Numbering')).toHaveValue('eachSect')
    await dialog.getByLabel('Start at').fill('8')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).settings).toContain('<w:numStart w:val="8"/>')
    let edited = await parts(file)
    expect(edited.settings.replace('<w:numStart w:val="8"/>', '')).toBe(original.settings)
    expect(edited.footnotes).toBe(original.footnotes)
    expect(edited.endnotes).toBe(original.endnotes)
    expect(edited.body).toContain('<w:footnoteReference w:id="1"/>')
    expect(edited.body).toContain('<w:pos w:val="beneathText"/>')
    expect(edited.body).toContain('<w:numFmt w:val="lowerRoman"/>')
    expect(edited.body).toContain('<w:numRestart w:val="eachSect"/>')

    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await parts(file)).settings)
      .not.toContain('<w:numStart w:val="8"/>')
    edited = await parts(file)
    expect(edited.settings).toBe(original.settings)
    expect(edited.footnotes).toBe(original.footnotes)
    expect(edited.endnotes).toBe(original.endnotes)

    await page.keyboard.press('ControlOrMeta+Shift+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).settings).toContain('<w:numStart w:val="8"/>')
    await closeAndSaveVideo(launched, 'word-first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-word-note-settings-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    dialog = await openSettings(page)
    await expect(dialog.getByLabel('Start at')).toHaveValue('8')
    await expect(dialog.getByLabel('Position')).toHaveValue('beneathText')
    await expect(dialog.getByLabel('Number format')).toHaveValue('lowerRoman')
    await expect(dialog.getByLabel('Numbering')).toHaveValue('eachSect')
    const evidenceDir = process.env.GENOFFICE_DOCS_NOTE_SETTINGS_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'word-note-settings-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'word-note-settings-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'word-final-session')
  }
})
