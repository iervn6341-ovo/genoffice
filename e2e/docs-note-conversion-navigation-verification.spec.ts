import { copyFile, mkdir, mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_DOCS_NOTE_CONVERSION_VERIFICATION,
  'Opt-in Word Mac note conversion and navigation verification',
)

const source = join(
  process.cwd(),
  'docs/qa/office-mac-field-evidence-2026-10-13/docs-native-notes/conversion/word-mac-conversion-before.docx',
)

async function parts(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    body: await zip.file('word/document.xml')!.async('string'),
    footnotes: await zip.file('word/footnotes.xml')!.async('string'),
    endnotes: await zip.file('word/endnotes.xml')!.async('string'),
  }
}

async function convert(page: Page, from: 'footnote' | 'endnote') {
  if (from === 'endnote') {
    const row = page.locator('.page-endnotes .page-note', {
      hasText: 'Custom symbol endnote body.',
    })
    const button = row.getByRole('button', { name: 'Convert to Footnote' })
    await expect(button).toBeVisible()
    await button.click()
    return
  }
  await page.locator(`.doc-note-ref[data-note-kind="${from}"]`).first().click({ button: 'right' })
  await page
    .locator('.ctx-menu')
    .getByRole('button', {
      name: from === 'footnote' ? 'Convert to Endnote' : 'Convert to Footnote',
    })
    .click()
}

for (const from of ['footnote', 'endnote'] as const) {
  test(`Word Mac ${from} converts on an ID collision with native parts, history and navigation`, async () => {
    test.setTimeout(240_000)
    const scratch = await mkdtemp(join(tmpdir(), `docs-note-conversion-${from}-`))
    const file = join(scratch, `conversion-${from}.docx`)
    const userDataDir = join(scratch, 'user-data')
    await mkdir(userDataDir)
    await copyFile(source, file)
    const target = from === 'footnote' ? 'endnote' : 'footnote'
    let launched: LaunchedApp | undefined
    try {
      launched = await launchShell({
        onboardingSeen: true,
        userDataDir,
        openFile: file,
        videoDir: `docs-conversion-${from}`,
      })
      let page = await waitForPageWithUrl(launched.app, '://docs/')
      page.setDefaultTimeout(10_000)
      await expect(page.locator('.doc-note-ref[data-note-kind="footnote"]')).toHaveText('♠')
      await expect(page.locator('.doc-note-ref[data-note-kind="endnote"]')).toHaveText('♥')
      await convert(page, from)
      await expect(page.locator(`.doc-note-ref[data-note-kind="${target}"]`)).toHaveText(['♠', '♥'])

      await page.locator(`.doc-note-ref[data-note-kind="${target}"]`).first().click()
      await page.locator('.ribbon-tabs .ribbon-tab', { hasText: /^References$/ }).click()
      const group = page.locator('.ribbon-body .rb-fold-group[aria-label="Footnotes"]')
      const next = group.getByRole('button', {
        name: target === 'footnote' ? 'Next Footnote' : 'Next Endnote',
        exact: true,
      })
      if (!(await next.isVisible())) await group.locator('button.rb-fold-btn').click()
      await next.click()
      await expect(page.locator('.doc-note-ref.ProseMirror-selectednode')).toHaveAttribute(
        'data-note-ref',
        '2',
      )

      await page.keyboard.press('ControlOrMeta+s')
      await expect.poll(async () => (await parts(file)).body).toContain('w:id="2"')
      let xml = await parts(file)
      expect(xml.body).not.toContain(`<w:${from}Reference`)
      expect(xml.body).toContain(`<w:${target}Reference w:customMarkFollows="1" w:id="1"/>`)
      expect(xml.body).toContain(`<w:${target}Reference w:customMarkFollows="1" w:id="2"/>`)
      const targetStyle = target === 'footnote' ? 'af0' : 'af3'
      expect(
        [
          ...xml.body.matchAll(
            new RegExp(`<w:rStyle w:val="${targetStyle}"\\/>[\\s\\S]*?<w:${target}Reference`, 'g'),
          ),
        ].length,
      ).toBe(2)
      expect(xml[target === 'footnote' ? 'footnotes' : 'endnotes']).toContain('w:char="F0AA"')
      expect(xml[target === 'footnote' ? 'footnotes' : 'endnotes']).toContain('w:char="F0A9"')
      expect(xml[from === 'footnote' ? 'footnotes' : 'endnotes']).not.toContain(
        `<w:${from} w:id="1">`,
      )

      await page.locator(`.doc-note-ref[data-note-kind="${target}"]`).first().click()
      await page.keyboard.press('ControlOrMeta+z')
      await page.keyboard.press('ControlOrMeta+s')
      await expect.poll(async () => (await parts(file)).body).toContain(`<w:${from}Reference`)
      xml = await parts(file)
      expect(xml.footnotes).toContain('w:char="F0AA"')
      expect(xml.endnotes).toContain('w:char="F0A9"')
      expect(xml.body).toContain(`<w:rStyle w:val="${from === 'footnote' ? 'af0' : 'af3'}"/>`)

      await page.locator(`.doc-note-ref[data-note-kind="${from}"]`).first().click()
      await page.keyboard.press('ControlOrMeta+Shift+z')
      await page.keyboard.press('ControlOrMeta+s')
      await expect.poll(async () => (await parts(file)).body).not.toContain(`<w:${from}Reference`)
      expect((await parts(file)).body).toContain(`<w:rStyle w:val="${targetStyle}"/>`)
      await closeAndSaveVideo(launched, 'first-session')

      launched = await launchShell({
        onboardingSeen: true,
        userDataDir,
        openFile: file,
        videoDir: `docs-conversion-${from}-reopen`,
      })
      page = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(page.locator(`.doc-note-ref[data-note-kind="${target}"]`)).toHaveText(['♠', '♥'])
      const convertedBody =
        from === 'footnote' ? 'Custom symbol footnote body.' : 'Custom symbol endnote body.'
      const noteRow = page.locator('.page-note', { hasText: convertedBody })
      const editButton = noteRow.getByRole('button', { name: `Edit ${target}` })
      await expect(editButton).toBeVisible()
      await editButton.click()
      const dialog = page.locator('.modal-backdrop .modal')
      await dialog.locator('textarea').fill('Converted note edited after reopen.')
      await dialog.getByRole('button', { name: 'OK' }).click()
      await page.keyboard.press('ControlOrMeta+s')
      await expect
        .poll(async () => (await parts(file))[target === 'footnote' ? 'footnotes' : 'endnotes'])
        .toContain('Converted note edited after reopen.')

      const evidenceDir = process.env.GENOFFICE_DOCS_NOTE_CONVERSION_VERIFICATION_DIR
      if (evidenceDir) {
        await mkdir(evidenceDir, { recursive: true })
        await copyFile(file, join(evidenceDir, `conversion-${from}-final.docx`))
        await page.screenshot({ path: join(evidenceDir, `conversion-${from}-reopened.png`) })
      }
    } finally {
      if (launched) await closeAndSaveVideo(launched, 'final-session')
    }
  })
}
