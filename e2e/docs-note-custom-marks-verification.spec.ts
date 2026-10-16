import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_DOCS_NOTE_CUSTOM_MARKS_VERIFICATION,
  'Opt-in native note custom-mark verification',
)

async function parts(path: string): Promise<{ body: string; footnotes: string; endnotes: string }> {
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    body: await zip.file('word/document.xml')!.async('string'),
    footnotes: (await zip.file('word/footnotes.xml')?.async('string')) ?? '',
    endnotes: (await zip.file('word/endnotes.xml')?.async('string')) ?? '',
  }
}

async function insertNote(page: Page, kind: 'Footnote' | 'Endnote', text: string, mark: string) {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^References$/ }).click()
  const group = page.locator('.ribbon-body .rb-fold-group[aria-label="Footnotes"]')
  const command = group.getByRole('button', { name: `Insert ${kind}`, exact: true })
  if (!(await command.isVisible())) await group.locator('button.rb-fold-btn').click()
  await command.click()
  const dialog = page.locator('.modal-backdrop .modal')
  await dialog.locator('textarea').fill(text)
  await dialog.getByRole('textbox', { name: 'Custom mark' }).fill(mark)
  await dialog.getByRole('button', { name: 'OK' }).click()
}

async function assertPagePreviewMarks(page: Page, footnoteBody: string, endnoteBody: string) {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^View$/ }).click()
  const group = page.locator('.ribbon-body .rb-fold-group[aria-label="Views"]')
  const command = group.getByRole('button', { name: 'Page Preview', exact: true })
  if (!(await command.isVisible())) await group.locator('button.rb-fold-btn').click()
  await command.click()

  const preview = page.locator('.pagination-preview')
  const footnote = preview.locator('.pv-footnotes .pv-footnote').filter({ hasText: footnoteBody })
  const endnote = preview.locator('.pv-endnotes .pv-footnote').filter({ hasText: endnoteBody })
  await expect(footnote.locator('sup')).toHaveText('♠')
  await expect(endnote.locator('sup')).toHaveText('♥')
  await preview.getByRole('button', { name: 'Close', exact: true }).click()
}

test('custom footnote and endnote marks remain native through save, undo, redo and reopen', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-note-custom-marks-'))
  const file = join(scratch, 'custom-marks.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  await writeFile(
    file,
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>CUSTOM_MARK_TARGET</w:t></w:r></w:p>' }),
  )
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-custom-marks',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    await page.locator('.doc-page.ProseMirror p').filter({ hasText: 'CUSTOM_MARK_TARGET' }).click()
    await page.keyboard.press('End')
    await insertNote(page, 'Footnote', 'CUSTOM_FOOTNOTE_BODY', '♠')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).footnotes).toContain('CUSTOM_FOOTNOTE_BODY')
    let xml = await parts(file)
    expect(xml.body).toMatch(
      /<w:footnoteReference[^>]*w:customMarkFollows="1"[^>]*w:id="1"\/><w:sym w:font="Symbol" w:char="F0AA"\/>/,
    )
    expect(xml.footnotes).toMatch(
      /<w:footnote w:id="1">[\s\S]*?<w:sym w:font="Symbol" w:char="F0AA"\/>/,
    )
    await expect(page.locator('.doc-note-ref').first()).toHaveText('♠')

    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).body).not.toContain('<w:footnoteReference')
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).body).toContain('w:customMarkFollows="1"')

    await page.locator('.doc-page.ProseMirror p').filter({ hasText: 'CUSTOM_MARK_TARGET' }).click()
    await page.keyboard.press('End')
    await insertNote(page, 'Footnote', 'AUTO_FOOTNOTE_BODY', '')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).footnotes).toContain('AUTO_FOOTNOTE_BODY')
    await expect(page.locator('.doc-note-ref').nth(1)).toHaveText('1')

    await page.locator('.doc-page.ProseMirror p').filter({ hasText: 'CUSTOM_MARK_TARGET' }).click()
    await page.keyboard.press('End')
    await insertNote(page, 'Endnote', 'CUSTOM_ENDNOTE_BODY', '♥')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).endnotes).toContain('CUSTOM_ENDNOTE_BODY')
    xml = await parts(file)
    expect(xml.body).toMatch(
      /<w:endnoteReference[^>]*w:customMarkFollows="1"[^>]*w:id="1"\/><w:sym w:font="Symbol" w:char="F0A9"\/>/,
    )
    expect(xml.endnotes).toMatch(
      /<w:endnote w:id="1">[\s\S]*?<w:sym w:font="Symbol" w:char="F0A9"\/>/,
    )
    await expect(
      page.locator('.page-endnotes .page-note', { hasText: 'CUSTOM_ENDNOTE_BODY' }).locator('sup'),
    ).toHaveText('♥')
    await assertPagePreviewMarks(page, 'CUSTOM_FOOTNOTE_BODY', 'CUSTOM_ENDNOTE_BODY')

    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-custom-marks-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-note-ref').first()).toHaveText('♠')
    await expect(page.locator('.doc-note-ref').last()).toHaveText('♥')
    await expect(
      page.locator('.page-endnotes .page-note', { hasText: 'CUSTOM_ENDNOTE_BODY' }).locator('sup'),
    ).toHaveText('♥')
    await page.locator('.page-note', { hasText: 'CUSTOM_FOOTNOTE_BODY' }).hover()
    await page.getByRole('button', { name: 'Edit footnote' }).first().click()
    const dialog = page.locator('.modal-backdrop .modal')
    await expect(dialog.getByRole('textbox', { name: 'Custom mark' })).toHaveValue('♠')
    await dialog.locator('textarea').fill('CUSTOM_FOOTNOTE_EDITED')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).footnotes).toContain('CUSTOM_FOOTNOTE_EDITED')
    expect((await parts(file)).footnotes).toContain('w:char="F0AA"')

    const evidenceDir = process.env.GENOFFICE_DOCS_NOTE_CUSTOM_MARKS_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'custom-marks-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'custom-marks-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})

test('Word Mac custom-mark fixture retains its source symbol runs when note text changes', async () => {
  test.setTimeout(150_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-note-word-mac-custom-'))
  const file = join(scratch, 'word-custom.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  await copyFile(
    join(
      process.cwd(),
      'docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/word-mac-custom-marks-source.docx',
    ),
    file,
  )
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-word-custom',
    })
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.doc-note-ref').first()).toHaveText('♠')
    await expect(page.locator('.doc-note-ref').last()).toHaveText('♥')
    await assertPagePreviewMarks(
      page,
      'Custom symbol footnote body.',
      'Custom symbol endnote body.',
    )
    await page.locator('.page-note', { hasText: 'Custom symbol footnote body.' }).hover()
    await page.getByRole('button', { name: 'Edit footnote' }).first().click()
    const dialog = page.locator('.modal-backdrop .modal')
    await expect(dialog.getByRole('textbox', { name: 'Custom mark' })).toHaveValue('♠')
    await dialog.locator('textarea').fill('Word Mac native mark edited body.')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await parts(file)).footnotes)
      .toContain('Word Mac native mark edited body.')
    const xml = await parts(file)
    expect(xml.body).toContain('w:customMarkFollows="1"')
    expect(xml.footnotes).toContain('<w:sym w:font="Symbol" w:char="F0AA"/>')
    expect(xml.endnotes).toContain('<w:sym w:font="Symbol" w:char="F0A9"/>')
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'word-fixture')
  }
})
