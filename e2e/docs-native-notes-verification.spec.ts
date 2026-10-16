import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(!process.env.GENOFFICE_DOCS_NOTES_VERIFICATION, 'Opt-in Docs native notes verification')

async function parts(path: string): Promise<{ body: string; footnotes: string; endnotes: string }> {
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    body: await zip.file('word/document.xml')!.async('string'),
    footnotes: (await zip.file('word/footnotes.xml')?.async('string')) ?? '',
    endnotes: (await zip.file('word/endnotes.xml')?.async('string')) ?? '',
  }
}

async function insertNote(page: Page, kind: 'Footnote' | 'Endnote', text: string): Promise<void> {
  await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^References$/ }).click()
  const group = page.locator('.ribbon-body .rb-fold-group[aria-label="Footnotes"]')
  const command = group.getByRole('button', { name: `Insert ${kind}`, exact: true })
  if (!(await command.isVisible())) await group.locator('button.rb-fold-btn').click()
  await command.click()
  const dialog = page.locator('.modal-backdrop .modal')
  await expect(dialog).toContainText(`Insert ${kind}`)
  await dialog.locator('textarea').fill(text)
  await dialog.getByRole('button', { name: 'OK' }).click()
}

test('footnote and endnote keep native references and bodies through Save, Undo, Redo and a new process', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-native-notes-'))
  const file = join(scratch, 'native-notes.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  await writeFile(
    file,
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>NOTE_TARGET</w:t></w:r></w:p>' }),
  )
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-native-notes',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    await page.locator('.doc-page.ProseMirror p').filter({ hasText: 'NOTE_TARGET' }).click()
    await page.keyboard.press('End')
    await insertNote(page, 'Footnote', 'NATIVE_FOOTNOTE')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).footnotes).toContain('NATIVE_FOOTNOTE')
    let xml = await parts(file)
    expect(xml.body).toMatch(/<w:footnoteReference\b[^>]*w:id="1"/)
    expect(xml.footnotes).toContain('<w:footnoteRef/>')

    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).body).not.toContain('<w:footnoteReference')
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).footnotes).toContain('NATIVE_FOOTNOTE')

    await page.locator('.doc-page.ProseMirror p').filter({ hasText: 'NOTE_TARGET' }).click()
    await page.keyboard.press('End')
    await insertNote(page, 'Endnote', 'NATIVE_ENDNOTE')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).endnotes).toContain('NATIVE_ENDNOTE')
    xml = await parts(file)
    expect(xml.body).toMatch(/<w:endnoteReference\b[^>]*w:id="1"/)
    expect(xml.endnotes).toContain('<w:endnoteRef/>')

    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-native-notes-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(page.locator('.page-note-text', { hasText: 'NATIVE_ENDNOTE' })).toBeVisible()
    await page.locator('.page-note', { hasText: 'NATIVE_ENDNOTE' }).hover()
    await page.getByRole('button', { name: 'Edit endnote' }).click()
    const dialog = page.locator('.modal-backdrop .modal')
    await expect(dialog.locator('textarea')).toHaveValue('NATIVE_ENDNOTE')
    await dialog.locator('textarea').fill('NATIVE_ENDNOTE_EDITED')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => (await parts(file)).endnotes).toContain('NATIVE_ENDNOTE_EDITED')
    const evidenceDir = process.env.GENOFFICE_DOCS_NOTES_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'native-notes-final.docx'))
      await page.screenshot({ path: join(evidenceDir, 'native-notes-reopened.png') })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})
