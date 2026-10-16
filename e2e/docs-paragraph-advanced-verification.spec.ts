import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import {
  buildDocx,
  IMAGE_PARAGRAPH_XML,
  MATH_PARAGRAPH_XML,
} from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_DOCS_PARAGRAPH_ADVANCED_VERIFICATION,
  'Opt-in Docs Paragraph Advanced verification',
)

const originalPPr =
  '<w:pPr><w:keepNext w:val="0"/><w:widowControl w:val="0"/>' +
  '<w:tabs><w:tab w:val="right" w:pos="1440" w:leader="heavy"/></w:tabs>' +
  '<w:spacing w:after="200"/><w:contextualSpacing/>' +
  '<w:rPr><w:lang w:val="ja-JP"/></w:rPr></w:pPr>'

async function documentParts(path: string) {
  const zip = await JSZip.loadAsync(await readFile(path))
  return {
    body: await zip.file('word/document.xml')!.async('string'),
    settings: await zip.file('word/settings.xml')!.async('string'),
  }
}

function targetPPr(body: string): string {
  const para =
    body.match(
      /<w:p(?:\s[^>]*)?>[\s\S]*?<w:t(?:\s[^>]*)?>ADVANCED_TARGET<\/w:t>[\s\S]*?<\/w:p>/,
    )?.[0] ?? ''
  return para.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? ''
}

function expectNativeDate(body: string): void {
  expect(body).toContain(
    '<w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple>',
  )
}

async function openParagraph(page: Page) {
  await page
    .locator('.doc-page.ProseMirror p')
    .filter({ hasText: 'ADVANCED_TARGET' })
    .first()
    .click({ button: 'right' })
  await page.locator('.ctx-item', { hasText: 'Paragraph…' }).click()
  const dialog = page.getByTestId('paragraph-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

test('Paragraph dialog edits native pagination, tabs and Asian settings with Undo across save', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-paragraph-advanced-'))
  const file = join(scratch, 'paragraph-advanced.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  const zip = await JSZip.loadAsync(
    await buildDocx({
      withImage: true,
      bodyXml:
        `<w:p>${originalPPr}<w:r><w:t>ADVANCED_TARGET</w:t></w:r>` +
        '<w:fldSimple w:instr=" DATE "><w:r><w:t>October</w:t></w:r></w:fldSimple></w:p>' +
        '<w:p><w:r><w:t>SECOND_PARAGRAPH</w:t></w:r></w:p>' +
        IMAGE_PARAGRAPH_XML +
        MATH_PARAGRAPH_XML,
    }),
  )
  zip.file(
    'word/settings.xml',
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="compressPunctuation"/>' +
      '</w:settings>',
  )
  const relsPath = 'word/_rels/document.xml.rels'
  const rels = await zip.file(relsPath)!.async('string')
  zip.file(
    relsPath,
    rels.replace(
      '</Relationships>',
      '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>',
    ),
  )
  await writeFile(file, await zip.generateAsync({ type: 'uint8array' }))

  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-paragraph-advanced',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    let dialog = await openParagraph(page)
    await dialog.getByTestId('paragraph-pane-breaks').click()
    await expect(dialog.getByLabel('Widow/orphan control')).not.toBeChecked()
    await dialog.getByLabel('Keep with next').check()
    await dialog.getByLabel('Do not hyphenate').check()
    await dialog.getByTestId('paragraph-pane-asian').click()
    await dialog.getByLabel('Use line-breaking rules').uncheck()
    await dialog.getByLabel('Compress punctuation at start of line').check()
    await dialog.getByRole('button', { name: 'Character compression' }).click()
    await dialog.locator('.gs-dd-item[data-value="compressPunctuationAndJapaneseKana"]').click()
    await dialog.getByTestId('paragraph-pane-tabs').click()
    await dialog.getByLabel('Default tab stops').fill('24')
    await dialog.getByLabel('Tab stop position').fill('90')
    await dialog.getByRole('button', { name: 'Alignment' }).click()
    await dialog.locator('.gs-dd-item[data-value="decimal"]').click()
    await dialog.getByRole('button', { name: 'Leader' }).click()
    await dialog.locator('.gs-dd-item[data-value="dot"]').click()
    await dialog.getByRole('button', { name: 'Set', exact: true }).click()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => targetPPr((await documentParts(file)).body))
      .toContain('<w:keepNext/>')
    let { body, settings } = await documentParts(file)
    expect(targetPPr(body)).toContain('<w:suppressAutoHyphens/>')
    expect(targetPPr(body)).toContain('<w:kinsoku w:val="0"/>')
    expect(targetPPr(body)).toContain('<w:topLinePunct/>')
    expect(targetPPr(body)).toContain('<w:tab w:val="decimal" w:pos="1800" w:leader="dot"/>')
    expect(settings).toContain('<w:defaultTabStop w:val="480"/>')
    expect(settings).toContain(
      '<w:characterSpacingControl w:val="compressPunctuationAndJapaneseKana"/>',
    )
    expectNativeDate(body)
    expect(body).toContain('<w:drawing>')
    expect(body).toContain('<m:oMath>')

    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => targetPPr((await documentParts(file)).body))
      .toContain('<w:keepNext w:val="0"/>')
    ;({ body, settings } = await documentParts(file))
    expect(targetPPr(body)).toContain('<w:tab w:val="right" w:pos="1440" w:leader="heavy"/>')
    expect(settings).toContain('<w:defaultTabStop w:val="720"/>')
    expect(settings).toContain('<w:characterSpacingControl w:val="compressPunctuation"/>')
    expectNativeDate(body)
    expect(body).toContain('<w:drawing>')
    expect(body).toContain('<m:oMath>')
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => (await documentParts(file)).settings)
      .toContain('<w:defaultTabStop w:val="480"/>')
    ;({ body } = await documentParts(file))
    expectNativeDate(body)
    expect(body).toContain('<w:drawing>')
    expect(body).toContain('<m:oMath>')

    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-paragraph-advanced-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    dialog = await openParagraph(page)
    await dialog.getByTestId('paragraph-pane-breaks').click()
    await expect(dialog.getByLabel('Keep with next')).toBeChecked()
    await dialog.getByTestId('paragraph-pane-tabs').click()
    await expect(dialog.getByLabel('Default tab stops')).toHaveValue('24')
    await expect(dialog.getByRole('listitem').filter({ hasText: '90 pt' })).toHaveCount(1)
    expectNativeDate((await documentParts(file)).body)
    const evidenceDir = process.env.GENOFFICE_DOCS_PARAGRAPH_ADVANCED_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'paragraph-advanced-final.docx'))
      await page.screenshot({
        path: join(evidenceDir, 'paragraph-advanced-reopened.png'),
        fullPage: true,
      })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})
