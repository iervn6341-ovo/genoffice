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
  !process.env.GENOFFICE_DOCS_FONT_ADVANCED_VERIFICATION,
  'Opt-in Docs Font Advanced verification',
)

const originalRPr =
  '<w:rFonts w:ascii="Cambria" w:eastAsia="SimSun" w:cs="Arial"/>' +
  '<w:u w:val="double" w:color="C0504D"/><w:spacing w:val="30"/>' +
  '<w:w w:val="125"/><w:kern w:val="28"/><w:position w:val="4"/>' +
  '<w:smallCaps/><w:lang w:val="en-US"/><w14:ligatures w14:val="all"/>' +
  '<w14:stylisticSets><w14:styleSet w14:id="4"/></w14:stylisticSets>'

async function documentXml(path: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(path))
  return zip.file('word/document.xml')!.async('string')
}

function alphaRPr(xml: string): string {
  const run = (xml.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<w:t[^>]*>FONT_ALPHA<\/w:t>[\s\S]*?<\/w:r>/) ??
    [])[0]
  return (run?.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) ?? [])[0] ?? ''
}

async function openFont(page: Page) {
  const alpha = page.locator('.doc-page.ProseMirror').getByText('FONT_ALPHA', { exact: true })
  await alpha.dblclick()
  await page.keyboard.press('ControlOrMeta+d')
  const dialog = page.locator('.font-advanced-dialog')
  await expect(dialog).toBeVisible()
  return dialog
}

test('Font and Advanced controls preserve mixed imported formatting, fields, picture and math across saves', async () => {
  test.setTimeout(240_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-font-advanced-'))
  const file = join(scratch, 'font-advanced.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  await writeFile(
    file,
    await buildDocx({
      docRootExtraAttrs: 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
      withImage: true,
      bodyXml:
        `<w:p><w:r><w:rPr>${originalRPr}</w:rPr><w:t>FONT_ALPHA</w:t></w:r>` +
        '<w:r><w:t xml:space="preserve"> FONT_BETA</w:t></w:r></w:p>' +
        '<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>October 13</w:t></w:r></w:fldSimple></w:p>' +
        IMAGE_PARAGRAPH_XML +
        MATH_PARAGRAPH_XML,
    }),
  )
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-font-advanced',
    })
    let page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    let dialog = await openFont(page)
    await dialog.getByRole('tab', { name: 'Advanced' }).click()
    await expect(dialog.getByLabel('Scale')).toHaveValue('125')
    await expect(dialog.getByLabel('By (pt)').first()).toHaveValue('1.5')
    await dialog.getByLabel('Scale').fill('90')
    await dialog.getByRole('button', { name: 'Position' }).click()
    await dialog.locator('.gs-dd-item[data-value="lowered"]').click()
    await dialog.getByLabel('By (pt)').last().fill('2')
    await dialog.getByLabel('Kerning').uncheck()
    await dialog.getByRole('button', { name: 'Stylistic sets' }).click()
    await dialog.locator('.gs-dd-item[data-value="7"]').click()
    await dialog.getByRole('tab', { name: 'Font', exact: true }).click()
    await dialog.getByRole('button', { name: 'Underline style' }).click()
    await dialog.locator('.gs-dd-item[data-value="wave"]').click()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => alphaRPr(await documentXml(file))).toContain('<w:w w:val="90"/>')
    let xml = await documentXml(file)
    expect(alphaRPr(xml)).toContain('<w:u w:val="wave" w:color="C0504D"/>')
    expect(alphaRPr(xml)).toContain('<w:kern w:val="0"/>')
    expect(alphaRPr(xml)).toContain('<w:spacing w:val="30"/>')
    expect(alphaRPr(xml)).toContain(
      '<w:rFonts w:ascii="Cambria" w:eastAsia="SimSun" w:cs="Arial"/>',
    )
    expect(alphaRPr(xml)).toContain('<w14:ligatures w14:val="all"/>')
    expect(alphaRPr(xml)).toContain('<w14:styleSet w14:id="7"/>')
    expect(xml).toContain('<w:fldSimple w:instr=" DATE ">')
    expect(xml).toContain('<w:drawing>')
    expect(xml).toContain('<m:oMath>')

    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => alphaRPr(await documentXml(file))).toContain('<w:w w:val="125"/>')
    expect(alphaRPr(await documentXml(file))).toContain('<w14:styleSet w14:id="4"/>')
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => alphaRPr(await documentXml(file))).toContain('<w:w w:val="90"/>')
    expect(alphaRPr(await documentXml(file))).toContain('<w14:styleSet w14:id="7"/>')

    await closeAndSaveVideo(launched, 'first-session')
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-font-advanced-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    dialog = await openFont(page)
    await dialog.getByRole('tab', { name: 'Advanced' }).click()
    await expect(dialog.getByLabel('Scale')).toHaveValue('90')
    await expect(dialog.getByLabel('Kerning')).not.toBeChecked()
    await expect(dialog.getByRole('button', { name: 'Stylistic sets' })).toContainText('Set 7')
    await dialog.getByRole('tab', { name: 'Font', exact: true }).click()
    await expect(dialog.getByRole('button', { name: 'Underline style' })).toContainText('Wavy')
    xml = await documentXml(file)
    expect(xml).toContain('<w:drawing>')
    expect(xml).toContain('<m:oMath>')
    const evidenceDir = process.env.GENOFFICE_DOCS_FONT_ADVANCED_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'font-advanced-final.docx'))
      await page.screenshot({
        path: join(evidenceDir, 'font-advanced-reopened.png'),
        fullPage: true,
      })
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'final-session')
  }
})

test('Font dialog resolves theme values and Undo after save restores theme references', async () => {
  test.setTimeout(150_000)
  const scratch = await mkdtemp(join(tmpdir(), 'docs-font-theme-'))
  const file = join(scratch, 'theme-font.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir)
  const originalThemeRPr =
    '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"' +
    ' w:eastAsiaTheme="majorEastAsia" w:ascii="Legacy Latin" w:hAnsi="Legacy Latin"' +
    ' w:eastAsia="Legacy East"/>' +
    '<w:color w:val="365F91" w:themeColor="accent1"/>'
  const themeXml =
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">' +
    '<a:themeElements><a:clrScheme name="Test"><a:accent1><a:srgbClr val="4472C4"/>' +
    '</a:accent1></a:clrScheme><a:fontScheme name="Test">' +
    '<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface="SimSun"/>' +
    '<a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/>' +
    '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme></a:themeElements></a:theme>'
  await writeFile(
    file,
    await buildDocx({
      bodyXml:
        `<w:p><w:r><w:rPr>${originalThemeRPr}</w:rPr><w:t>THEME_ALPHA</w:t></w:r></w:p>` +
        '<w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>October 13</w:t></w:r></w:fldSimple></w:p>',
      extraParts: [
        {
          path: 'word/theme/theme1.xml',
          xml: themeXml,
          contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
        },
      ],
    }),
  )
  const themeRPr = (xml: string) => {
    const run = xml.match(/<w:r(?:\s[^>]*)?>[\s\S]*?<w:t[^>]*>THEME_ALPHA<\/w:t>[\s\S]*?<\/w:r>/)?.[0]
    return run?.match(/<w:rPr>[^]*?<\/w:rPr>/)?.[0] ?? ''
  }
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: file,
      videoDir: 'docs-font-theme',
    })
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    await page.locator('.doc-page.ProseMirror').getByText('THEME_ALPHA', { exact: true }).dblclick()
    await page.keyboard.press('ControlOrMeta+d')
    const dialog = page.locator('.font-advanced-dialog')
    await expect(dialog).toBeVisible()
    await expect(dialog.getByRole('button', { name: 'Latin font' })).toContainText('Calibri')
    await expect(dialog.getByRole('button', { name: 'East Asian font' })).toContainText('SimSun')
    await expect(dialog.getByRole('textbox', { name: 'Font color' })).toHaveValue('#4472c4')
    await dialog.getByRole('button', { name: 'Latin font' }).click()
    await dialog.locator('.gs-dd-item[data-value="Arial"]').click()
    await dialog.getByRole('textbox', { name: 'Font color' }).fill('#c0504d')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => themeRPr(await documentXml(file))).toContain('w:ascii="Arial"')
    let changed = themeRPr(await documentXml(file))
    expect(changed).not.toContain('w:asciiTheme=')
    expect(changed).not.toContain('w:hAnsiTheme=')
    expect(changed).toContain('w:eastAsiaTheme="majorEastAsia"')
    expect(changed).toContain('<w:color w:val="C0504D"/>')
    expect(changed).not.toContain('w:themeColor=')
    await page.keyboard.press('ControlOrMeta+z')
    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => themeRPr(await documentXml(file)))
      .toContain('w:asciiTheme="minorHAnsi"')
    changed = themeRPr(await documentXml(file))
    expect(changed).toContain('w:hAnsiTheme="minorHAnsi"')
    expect(changed).toContain('w:eastAsiaTheme="majorEastAsia"')
    expect(changed).toContain('<w:color w:val="365F91" w:themeColor="accent1"/>')
    expect(await documentXml(file)).toContain('<w:fldSimple w:instr=" DATE ">')
    const evidenceDir = process.env.GENOFFICE_DOCS_FONT_ADVANCED_VERIFICATION_DIR
    if (evidenceDir) {
      await mkdir(evidenceDir, { recursive: true })
      await copyFile(file, join(evidenceDir, 'font-theme-undo-restored.docx'))
    }
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'theme-session')
  }
})
