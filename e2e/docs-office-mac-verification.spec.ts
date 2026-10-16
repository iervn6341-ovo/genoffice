import { existsSync } from 'node:fs'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Opt-in desktop verification for the 2026-10-13 Docs header/footer work.
 * It exercises the actual 1280px ribbon and a generated scratch DOCX with
 * three sections. The generated file is isolated under the OS temp directory;
 * captured evidence is written to the dated QA folder (or the env override).
 */

const VERIFY = process.env.GENOFFICE_DOCS_MAC_VERIFICATION
const EVIDENCE_ROOT =
  process.env.GENOFFICE_DOCS_VERIFICATION_DIR ??
  resolve(__dirname, '../docs/qa/office-mac-evidence-2026-10-13/docs-verification')
const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
test.skip(!VERIFY, 'Opt-in Docs Office for Mac verification')

function sectionProperties(xml: string): string[] {
  return [...xml.matchAll(/<w:sectPr\b[^>]*\/>|<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)].map(
    ([part]) => part,
  )
}

function referenceId(sectPr: string, kind: 'header' | 'footer', variant: string): string | null {
  const refs = sectPr.match(new RegExp(`<w:${kind}Reference\\b[^>]*/>`, 'g')) ?? []
  const ref = refs.find((tag) => {
    const type = /w:type="([^"]+)"/.exec(tag)?.[1] ?? 'default'
    return type === variant || (variant === 'default' && type === 'odd')
  })
  return ref ? (/r:id="([^"]+)"/.exec(ref)?.[1] ?? null) : null
}

function relTarget(relsXml: string, rId: string): string | null {
  const rel = relsXml.match(new RegExp(`<Relationship\\b(?=[^>]*\\bId="${rId}")[^>]*/>`))?.[0]
  return rel ? (/Target="([^"]+)"/.exec(rel)?.[1] ?? null) : null
}

async function savedPackage(file: string) {
  if (!existsSync(file)) return { exists: false as const }
  const zip = await JSZip.loadAsync(await readFile(file))
  const documentXml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  const documentRels = (await zip.file('word/_rels/document.xml.rels')?.async('string')) ?? ''
  const sections = sectionProperties(documentXml)
  return {
    exists: true as const,
    zip,
    documentXml,
    documentRels,
    sections,
    ref: (sectionIndex: number, kind: 'header' | 'footer', variant: string) =>
      sections[sectionIndex] ? referenceId(sections[sectionIndex]!, kind, variant) : null,
    part: (rId: string | null) => {
      if (!rId) return null
      const target = relTarget(documentRels, rId)
      if (!target) return null
      return target.startsWith('/') ? target.slice(1) : `word/${target}`
    },
  }
}

function richHeader(text: string, withImage = false): string {
  const imageParagraph = withImage
    ? '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    : ''
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr><w:t>${text}</w:t></w:r></w:p>` +
    '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:t xml:space="preserve"> of </w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
    '<w:tr><w:tc><w:p><w:r><w:t>KEEP_TABLE_LEFT</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>KEEP_TABLE_RIGHT</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
    imageParagraph +
    '</w:hdr>'
  )
}

function simpleHf(root: 'w:hdr' | 'w:ftr', text: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></${root}>`
  )
}

const headerRel = (rId: string, target: string) =>
  `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="${target}"/>`
const footerRel = (rId: string, target: string) =>
  `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="${target}"/>`
const headerRef = (variant: string, rId: string) =>
  `<w:headerReference w:type="${variant}" r:id="${rId}"/>`
const footerRef = (variant: string, rId: string) =>
  `<w:footerReference w:type="${variant}" r:id="${rId}"/>`
const sectionBreak = (refs: string, differentFirst = false) =>
  `<w:p><w:pPr><w:sectPr>${refs}<w:type w:val="nextPage"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>${differentFirst ? '<w:titlePg/>' : ''}</w:sectPr></w:pPr></w:p>`

/** Three-section scratch fixture with variant refs and a rich first-page header. */
async function createHeaderFooterFixture(file: string): Promise<string> {
  const zip = new JSZip()
  const relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
  const contentTypeNs = 'http://schemas.openxmlformats.org/package/2006/content-types'
  const headerType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
  const footerType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'
  const relsType = 'application/vnd.openxmlformats-package.relationships+xml'
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${contentTypeNs}"><Default Extension="rels" ContentType="${relsType}"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>${[1, 2, 3, 4, 5].map((n) => `<Override PartName="/word/header${n}.xml" ContentType="${headerType}"/>`).join('')}<Override PartName="/word/footer1.xml" ContentType="${footerType}"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${relNs}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${relNs}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>${headerRel('rId20', 'header1.xml')}${headerRel('rId21', 'header2.xml')}${headerRel('rId22', 'header3.xml')}${headerRel('rId23', 'header4.xml')}${headerRel('rId24', 'header5.xml')}${footerRel('rId30', 'footer1.xml')}</Relationships>`,
  )
  zip.file(
    'word/settings.xml',
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:evenAndOddHeaders/></w:settings>',
  )
  zip.file(
    'word/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>',
  )

  const firstSectionRefs =
    headerRef('default', 'rId20') +
    headerRef('first', 'rId21') +
    headerRef('even', 'rId23') +
    footerRef('default', 'rId30')
  const secondSectionRefs = headerRef('default', 'rId22') + headerRef('even', 'rId23')
  const thirdSectionRefs = headerRef('default', 'rId24')
  const body =
    '<w:p><w:r><w:t>Section one body</w:t></w:r></w:p>' +
    sectionBreak(firstSectionRefs, true) +
    '<w:p><w:r><w:t>Section two body</w:t></w:r></w:p>' +
    sectionBreak(secondSectionRefs) +
    '<w:p><w:r><w:t>Section three body</w:t></w:r></w:p>'
  const trailingSectPr = `<w:sectPr>${thirdSectionRefs}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/></w:sectPr>`
  zip.file(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>' +
      body +
      trailingSectPr +
      '</w:body></w:document>',
  )
  zip.file('word/header1.xml', simpleHf('w:hdr', 'Section 1 odd header'))
  zip.file('word/header2.xml', richHeader('First page heading', true))
  zip.file('word/header3.xml', simpleHf('w:hdr', 'Section 2 odd header'))
  zip.file('word/header4.xml', simpleHf('w:hdr', 'Section 2 even header'))
  zip.file('word/header5.xml', simpleHf('w:hdr', 'Section 3 odd header'))
  const originalFooterXml =
    '<?xml version="1.0" encoding="UTF-8"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer line one</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> of </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'
  zip.file('word/footer1.xml', originalFooterXml)
  zip.file(
    'word/_rels/header2.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${relNs}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/></Relationships>`,
  )
  zip.file('word/media/logo.png', TINY_PNG_BASE64, { base64: true })
  await writeFile(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
  return originalFooterXml
}

async function openRibbonTab(page: Page, name: string): Promise<void> {
  const tab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: new RegExp(`^${name}$`) })
  await expect(tab).toBeVisible()
  if (!(await tab.evaluate((el) => el.classList.contains('active')))) await tab.click()
  await expect(tab).toHaveClass(/active/)
}

async function openHeaderEditor(page: Page): Promise<ReturnType<Page['locator']>> {
  await openRibbonTab(page, 'Insert')
  await page
    .locator('.ribbon-body button', { hasText: /^Header$/ })
    .first()
    .click()
  const editor = page.locator('.hf-menu')
  await expect(editor).toBeVisible()
  await expect(editor.locator('textarea')).toBeVisible()
  return editor
}

async function clickSectionChip(page: Page, label: string): Promise<void> {
  const chips = page.locator('.hf-variant-chips')
  await chips.getByRole('button', { name: label, exact: true }).click()
}

async function expectSelectedHeader(page: Page, text: string): Promise<void> {
  const editor = await openHeaderEditor(page)
  await expect.poll(() => editor.locator('textarea').inputValue()).toContain(text)
  await editor.locator('textarea').press('Escape')
  await expect(editor).not.toBeVisible()
}

async function evidenceFolder(testName: string): Promise<string> {
  const run = new Date().toISOString().replace(/[.:]/g, '-')
  const dir = join(EVIDENCE_ROOT, `${testName}-${run}`)
  await mkdir(dir, { recursive: true })
  return dir
}

test('Docs 1280px folded Symbols group keeps Symbol and Equation pickers usable', async () => {
  test.setTimeout(180_000)
  const evidenceDir = await evidenceFolder('folded-symbols')
  const evidence: Record<string, unknown> = { status: 'running', evidenceDir, steps: [] }
  let launched: LaunchedApp | undefined
  let page: Page | undefined
  const step = (value: string) => {
    evidence.steps = [...(evidence.steps as string[]), value]
  }
  try {
    launched = await launchShell({
      onboardingSeen: true,
      videoDir: 'docs-office-mac-verification-folded',
    })
    await launched.page.locator('.quick-card').first().click()
    page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    const editor = page.locator('.doc-page.ProseMirror').first()
    await expect(editor).toBeVisible()
    await editor.click()
    await page.keyboard.type('Folded ribbon target')
    await page.keyboard.press('End')
    await openRibbonTab(page, 'Insert')

    const symbolsGroup = page.locator('.rb-fold-group[aria-label="Symbols"]')
    await expect(symbolsGroup).toHaveAttribute('data-folded', '')
    const foldButton = symbolsGroup.locator(':scope > .rb-fold-btn')
    await foldButton.click()
    await expect(foldButton).toHaveAttribute('aria-expanded', 'true')
    await symbolsGroup.locator('button[data-tip="Insert a symbol"]').click()
    const palette = symbolsGroup.locator('.symbol-palette')
    await expect(palette).toBeVisible()
    await expect(foldButton).toHaveAttribute('aria-expanded', 'true')
    await palette.getByRole('button', { name: '≤', exact: true }).click()
    await expect(editor).toContainText('Folded ribbon target≤')
    await expect(foldButton).toHaveAttribute('aria-expanded', 'false')
    step(
      'At 1280px, folded Symbols launcher stayed open while its Symbol palette opened; ≤ inserted and completed the parent fold.',
    )
    await page.screenshot({ path: join(evidenceDir, 'after-symbol.png'), fullPage: true })

    await editor.getByText('Folded ribbon target≤', { exact: true }).click()
    await page.keyboard.press('End')
    await foldButton.click()
    await symbolsGroup.locator('button[data-tip="Insert an equation"]').click()
    const gallery = symbolsGroup.locator('.equation-gallery')
    await expect(gallery).toBeVisible()
    await expect(foldButton).toHaveAttribute('aria-expanded', 'true')
    await gallery.locator('.equation-gallery-custom').click()
    const equation = page.locator('.equation-modal')
    await expect(equation).toBeVisible()
    await equation.locator('input:not([type="checkbox"])').fill('\\frac{x}{y}')
    await equation.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(editor.locator('mfrac')).toHaveCount(1)
    step(
      'At the same 1280px folded state, the Equation gallery and custom equation dialog stayed operable; the fraction inserted.',
    )
    await page.screenshot({ path: join(evidenceDir, 'after-equation.png'), fullPage: true })
    evidence.status = 'passed'
  } catch (error) {
    evidence.status = 'failed'
    evidence.error = error instanceof Error ? `${error.name}: ${error.message}` : String(error)
    if (page)
      await page
        .screenshot({ path: join(evidenceDir, 'failure.png'), fullPage: true })
        .catch(() => undefined)
    throw error
  } finally {
    evidence.finishedAt = new Date().toISOString()
    await writeFile(join(evidenceDir, 'evidence.json'), JSON.stringify(evidence, null, 2))
    if (launched) await closeAndSaveVideo(launched, 'docs-office-mac-verification-folded')
  }
})

test('Docs section headers preserve rich content, variants, links, and save/reopen state', async () => {
  test.setTimeout(300_000)
  const scratchDir = await mkdtemp(join(tmpdir(), 'docs-office-mac-verification-'))
  const fixture = join(scratchDir, 'three-sections.docx')
  const originalFooterXml = await createHeaderFooterFixture(fixture)
  const evidenceDir = await evidenceFolder('sections-rich-header')
  const evidence: Record<string, unknown> = {
    status: 'running',
    evidenceDir,
    scratchDocument: fixture,
    steps: [],
  }
  const userDataDir = join(scratchDir, 'user-data')
  await mkdir(userDataDir, { recursive: true })
  let launched: LaunchedApp | undefined
  let page: Page | undefined
  const step = (value: string) => {
    evidence.steps = [...(evidence.steps as string[]), value]
  }
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      videoDir: 'docs-office-mac-verification-sections',
      openFile: fixture,
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    const header = page.locator('.page-hf-header').first()
    const footer = page.locator('.page-hf-footer').first()
    const chips = page.locator('.hf-variant-chips')
    await expect(page.locator('.doc-page.ProseMirror').first()).toContainText('Section one body')
    await expect(header).toContainText('First page heading')

    await clickSectionChip(page, 'Next')
    await expect(chips).toContainText('Page Setup (Section 2)')
    await clickSectionChip(page, 'Odd Pages')
    await expectSelectedHeader(page, 'Section 2 odd header')
    await expect(header).toContainText('First page heading')
    await expect(footer).toContainText('Footer line one')
    await clickSectionChip(page, 'Even Pages')
    await expectSelectedHeader(page, 'Section 2 even header')
    await clickSectionChip(page, 'Odd Pages')
    step(
      'Next navigated to section 2; Odd Page and Even Page selected and displayed their own saved content.',
    )

    await chips.getByRole('button', { name: 'Different First Page', exact: false }).click()
    await expect(chips.getByRole('button', { name: 'First Page', exact: true })).toBeVisible()
    await clickSectionChip(page, 'First Page')
    await expect(header).toContainText('First page heading')
    await expect(footer.locator('.page-hf-para')).toHaveText('')
    const firstLink = chips.getByRole('button', { name: /^Header:\s*Link to Previous$/ })
    await expect(firstLink).toHaveAttribute('aria-pressed', 'true')
    const hf = await openHeaderEditor(page)
    await hf
      .locator('textarea')
      .fill(
        (await hf.locator('textarea').inputValue()).replace(
          'First page heading',
          'First page heading edited',
        ),
      )
    await hf.getByRole('button', { name: 'OK', exact: true }).click()
    await expectSelectedHeader(page, 'First page heading edited')
    await expect(firstLink).toHaveAttribute('aria-pressed', 'false')
    step(
      'Section 2 enabled a distinct first page and edited its inherited first header through the Insert > Header UI.',
    )

    const sectionBeforeSave = await savedPackage(fixture)
    expect(sectionBeforeSave.sections).toHaveLength(3)
    expect(sectionBeforeSave.ref(0, 'header', 'first')).toBe('rId21')
    const undoButton = page.locator('.qa-btn[aria-label="Undo"]')
    const redoButton = page.locator('.qa-btn[aria-label="Redo"]')
    await expect(undoButton).toBeEnabled()
    await undoButton.click()
    await expectSelectedHeader(page, 'First page heading\n')
    await expect(firstLink).toHaveAttribute('aria-pressed', 'true')
    await expect(redoButton).toBeEnabled()
    await redoButton.click()
    await expectSelectedHeader(page, 'First page heading edited')
    await expect(firstLink).toHaveAttribute('aria-pressed', 'false')
    step('Undo and Redo before save restored both first-header content and its link state.')

    await clickSectionChip(page, 'Next')
    await expect(chips).toContainText('Page Setup (Section 3)')
    await clickSectionChip(page, 'Odd Pages')
    const thirdLink = chips.getByRole('button', { name: /^Header:\s*Link to Previous$/ })
    await expect(thirdLink).toHaveAttribute('aria-pressed', 'false')
    await thirdLink.click()
    await expect(thirdLink).toHaveAttribute('aria-pressed', 'true')
    await expectSelectedHeader(page, 'Section 2 odd header')
    step(
      'Section 3 linked its default header to section 2 through the on-page Link to Previous control.',
    )

    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => {
        const saved = await savedPackage(fixture)
        return {
          sections: saved.sections.length,
          section2FirstRef: saved.ref(1, 'header', 'first'),
          section3DefaultRef: saved.ref(2, 'header', 'default'),
        }
      })
      .toMatchObject({ sections: 3, section3DefaultRef: null })
    let saved = await savedPackage(fixture)
    const clonedFirstId = saved.ref(1, 'header', 'first')
    expect(clonedFirstId).toBeTruthy()
    expect(clonedFirstId).not.toBe('rId21')
    const clonedPart = saved.part(clonedFirstId)
    expect(clonedPart).toBeTruthy()
    const clonedXml = (await saved.zip.file(clonedPart!)?.async('string')) ?? ''
    expect(clonedXml).toContain('First page heading edited')
    expect(clonedXml).toContain('KEEP_TABLE_LEFT')
    expect(clonedXml).toContain('KEEP_TABLE_RIGHT')
    expect(clonedXml).toMatch(/<w:b(?:\s|\/>)/)
    expect(clonedXml).toContain('C00000')
    expect(clonedXml).toContain(' PAGE ')
    expect(clonedXml).toContain(' NUMPAGES ')
    expect(clonedXml).toContain('<w:drawing>')
    const clonedRelPath = clonedPart!.replace(/^word\//, 'word/_rels/') + '.rels'
    expect(await saved.zip.file(clonedRelPath)?.async('string')).toContain('media/logo.png')
    expect(await saved.zip.file('word/media/logo.png')?.async('base64')).toBe(TINY_PNG_BASE64)
    expect(await saved.zip.file('word/footer1.xml')?.async('string')).toBe(originalFooterXml)
    expect((await saved.zip.file('word/header2.xml')?.async('string')) ?? '').toContain(
      'First page heading',
    )
    expect(saved.ref(0, 'header', 'first')).toBe('rId21')
    expect(saved.ref(2, 'header', 'default')).toBeNull()
    expect(saved.ref(0, 'footer', 'default')).toBe('rId30')
    expect(saved.ref(1, 'footer', 'first')).toBeNull()
    evidence.savedRefs = {
      section1First: saved.ref(0, 'header', 'first'),
      section2First: clonedFirstId,
      section3Default: saved.ref(2, 'header', 'default'),
    }
    evidence.clonedHeaderPart = clonedPart
    step(
      'Saved DOCX contains an independent section 2 first-header part with styled runs, both field instructions, table, picture bytes, and copied relationship; section 3 has no local default ref.',
    )
    await page.screenshot({ path: join(evidenceDir, 'saved-sections.png'), fullPage: true })
    await copyFile(fixture, join(evidenceDir, 'three-sections-saved.docx'))

    await closeAndSaveVideo(launched, 'docs-office-mac-verification-sections-first-launch')
    launched = undefined
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      videoDir: 'docs-office-mac-verification-sections-reopen',
      openFile: fixture,
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    const reopenedHeader = page.locator('.page-hf-header').first()
    const reopenedChips = page.locator('.hf-variant-chips')
    await clickSectionChip(page, 'Next')
    await expect(reopenedChips).toContainText('Page Setup (Section 2)')
    await expect(
      reopenedChips.getByRole('button', { name: 'First Page', exact: true }),
    ).toBeVisible()
    await clickSectionChip(page, 'First Page')
    await expectSelectedHeader(page, 'First page heading edited')
    await expect(reopenedHeader).toContainText('First page heading')
    const reopenedFooter = page.locator('.page-hf-footer').first()
    await expect(reopenedFooter.locator('.page-hf-para')).toHaveText('')
    await clickSectionChip(page, 'Odd Pages')
    await expect(reopenedFooter).toContainText('Footer line one')
    await clickSectionChip(page, 'Next')
    await expect(reopenedChips).toContainText('Page Setup (Section 3)')
    await clickSectionChip(page, 'Odd Pages')
    await expect(
      reopenedChips.getByRole('button', { name: /^Header:\s*Link to Previous$/ }),
    ).toHaveAttribute('aria-pressed', 'true')
    await expectSelectedHeader(page, 'Section 2 odd header')
    step(
      'A fresh Electron launch reopened the DOCX with the edited first header, untouched footer, and section 3 link state visible in the UI.',
    )
    await page.screenshot({ path: join(evidenceDir, 'reopened-sections.png'), fullPage: true })

    saved = await savedPackage(fixture)
    expect(saved.sections).toHaveLength(3)
    expect(saved.ref(1, 'header', 'first')).toBe(clonedFirstId)
    expect(saved.ref(2, 'header', 'default')).toBeNull()
    evidence.status = 'passed'
  } catch (error) {
    evidence.status = 'failed'
    evidence.error = error instanceof Error ? `${error.name}: ${error.message}` : String(error)
    if (page)
      await page
        .screenshot({ path: join(evidenceDir, 'failure.png'), fullPage: true })
        .catch(() => undefined)
    throw error
  } finally {
    evidence.finishedAt = new Date().toISOString()
    await writeFile(join(evidenceDir, 'evidence.json'), JSON.stringify(evidence, null, 2))
    if (launched) await closeAndSaveVideo(launched, 'docs-office-mac-verification-sections')
  }
})

test('Docs section header Undo stays correct across a save', async () => {
  test.setTimeout(240_000)
  const scratchDir = await mkdtemp(join(tmpdir(), 'docs-office-mac-undo-'))
  const fixture = join(scratchDir, 'undo-sections.docx')
  await createHeaderFooterFixture(fixture)
  const evidenceDir = await evidenceFolder('section-header-undo')
  const userDataDir = join(scratchDir, 'user-data')
  await mkdir(userDataDir, { recursive: true })
  const evidence: Record<string, unknown> = {
    status: 'running',
    evidenceDir,
    scratchDocument: fixture,
    steps: [],
  }
  let launched: LaunchedApp | undefined
  let page: Page | undefined
  const step = (value: string) => {
    evidence.steps = [...(evidence.steps as string[]), value]
  }
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      videoDir: 'docs-office-mac-verification-undo',
      openFile: fixture,
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    const chips = page.locator('.hf-variant-chips')
    await clickSectionChip(page, 'Next')
    await expect(chips).toContainText('Page Setup (Section 2)')
    await clickSectionChip(page, 'Odd Pages')
    await expectSelectedHeader(page, 'Section 2 odd header')
    const initialPackage = await savedPackage(fixture)
    const initialHeaderId = initialPackage.ref(1, 'header', 'default')
    const initialHeaderPart = initialPackage.part(initialHeaderId)
    expect(initialHeaderId).toBeTruthy()
    expect(initialHeaderPart).toBeTruthy()
    const initialHeaderXml = await initialPackage.zip.file(initialHeaderPart!)?.async('string')
    expect(initialHeaderXml).toContain('Section 2 odd header')
    expect(initialHeaderXml).not.toContain('after edit')
    const editor = await openHeaderEditor(page)
    await editor.locator('textarea').fill('Section 2 odd header after edit')
    await editor.getByRole('button', { name: 'OK', exact: true }).click()
    await expectSelectedHeader(page, 'Section 2 odd header after edit')

    const undoButton = page.locator('.qa-btn[aria-label="Undo"]')
    const redoButton = page.locator('.qa-btn[aria-label="Redo"]')
    await expect(undoButton).toBeEnabled()
    await undoButton.click()
    await expectSelectedHeader(page, 'Section 2 odd header')
    await expect(redoButton).toBeEnabled()
    await redoButton.click()
    await expectSelectedHeader(page, 'Section 2 odd header after edit')
    step('Undo and Redo before save restored the section 2 header content.')

    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => {
        const saved = await savedPackage(fixture)
        const id = saved.ref(1, 'header', 'default')
        const path = saved.part(id)
        return path ? ((await saved.zip.file(path)?.async('string')) ?? '') : ''
      })
      .toContain('Section 2 odd header after edit')
    await expect(undoButton).toBeEnabled()
    await undoButton.click()
    await expectSelectedHeader(page, 'Section 2 odd header')
    step('After save, Undo restored the previous section header in the editor.')

    await page.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(async () => {
        const saved = await savedPackage(fixture)
        const id = saved.ref(1, 'header', 'default')
        const part = saved.part(id)
        const xml = part ? await saved.zip.file(part)?.async('string') : null
        return id === initialHeaderId && xml === initialHeaderXml
      })
      .toBe(true)
    const packageAfterUndo = await savedPackage(fixture)
    const section2Part = packageAfterUndo.part(packageAfterUndo.ref(1, 'header', 'default'))
    expect(section2Part).toBeTruthy()
    expect(await packageAfterUndo.zip.file(section2Part!)?.async('string')).toBe(initialHeaderXml)
    evidence.status = 'passed'
    evidence.section2DefaultHeaderPart = section2Part
    step('A second save persisted the undone header text to the section 2 default header part.')
    await page.screenshot({ path: join(evidenceDir, 'undo-after-save.png'), fullPage: true })
    await copyFile(fixture, join(evidenceDir, 'undo-after-save.docx'))
  } catch (error) {
    evidence.status = 'failed'
    evidence.error = error instanceof Error ? `${error.name}: ${error.message}` : String(error)
    if (page)
      await page
        .screenshot({ path: join(evidenceDir, 'failure.png'), fullPage: true })
        .catch(() => undefined)
    throw error
  } finally {
    evidence.finishedAt = new Date().toISOString()
    await writeFile(join(evidenceDir, 'evidence.json'), JSON.stringify(evidence, null, 2))
    if (launched) await closeAndSaveVideo(launched, 'docs-office-mac-verification-undo')
  }
})

test('Docs DATE field survives modal and canvas edits, Undo across saves, and reopen', async () => {
  test.setTimeout(300_000)
  const scratchDir = await mkdtemp(join(tmpdir(), 'docs-hf-field-'))
  const fixture = join(scratchDir, 'date-field.docx')
  await createHeaderFooterFixture(fixture)
  const fieldXml =
    '<w:r><w:fldChar w:fldCharType="begin" w:dirty="1" w:fldLock="1"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> DA</w:instrText></w:r>' +
    '<w:r><w:instrText xml:space="preserve">TE \\@ "MMMM d" </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:rPr><w:b/></w:rPr><w:t>October</w:t></w:r>' +
    '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> 13</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  const originalZip = await JSZip.loadAsync(await readFile(fixture))
  const originalHeader = (await originalZip.file('word/header2.xml')!.async('string')).replace(
    '<w:t>First page heading</w:t></w:r>',
    `<w:t xml:space="preserve">Before </w:t></w:r>${fieldXml}<w:r><w:t xml:space="preserve"> after</w:t></w:r>`,
  )
  originalZip.file('word/header2.xml', originalHeader)
  await writeFile(
    fixture,
    await originalZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
  )
  const currentHeaderXml = async (): Promise<string> => {
    const saved = await savedPackage(fixture)
    if (!saved.exists || !('ref' in saved)) throw new Error('Saved DOCX is missing')
    const part = saved.part(saved.ref(0, 'header', 'first'))
    if (!part) throw new Error('Current first-page header reference is missing')
    const xml = (await saved.zip.file(part)?.async('string')) ?? ''
    expect(xml).toContain(fieldXml)
    expect(xml).toMatch(/<w:instrText[^>]*> PAGE <\/w:instrText>/)
    expect(xml).toMatch(/<w:instrText[^>]*> NUMPAGES <\/w:instrText>/)
    expect(xml).toContain('KEEP_TABLE_LEFT')
    expect(xml).toContain('KEEP_TABLE_RIGHT')
    expect(xml).toContain('<w:drawing>')
    return xml
  }
  const userDataDir = join(scratchDir, 'user-data')
  await mkdir(userDataDir, { recursive: true })
  let launched: LaunchedApp | undefined
  let page: Page | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: fixture,
      videoDir: 'docs-hf-date-field',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    page.setDefaultTimeout(10_000)
    const header = page.locator('.page-hf-header:not(.page-gap-hf)').first()
    await expect(header).toContainText('Before October 13 after')
    const modal = await openHeaderEditor(page)
    const input = modal.locator('textarea')
    const tokenText = await input.inputValue()
    expect(tokenText).toMatch(/Before ⟦DATE·\d+⟧ after/)
    await input.fill(tokenText.replace('Before ', 'Modal ').replace(' after', ' changed\nNew line'))
    await modal.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(header).toContainText('Modal October 13 changed')
    await expect(header).toContainText('New line')
    const undo = page.locator('.qa-btn[aria-label="Undo"]')
    const redo = page.locator('.qa-btn[aria-label="Redo"]')
    await undo.click()
    await expect(header).toContainText('Before October 13 after')
    await redo.click()
    await expect(header).toContainText('Modal October 13 changed')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => await currentHeaderXml()).toContain('Modal ')
    let savedXml = await currentHeaderXml()
    expect(savedXml).toContain(fieldXml)
    await undo.click()
    await expect(header).toContainText('Before October 13 after')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => await currentHeaderXml()).toContain('Before ')
    savedXml = await currentHeaderXml()
    expect(savedXml).toContain(fieldXml)
    await redo.click()
    await expect(header).toContainText('Modal October 13 changed')
    await header.dblclick()
    const canvas = page.locator('.page-hf-header .page-hf-edit-surface').first()
    await expect(canvas).toBeVisible()
    await canvas.fill((await canvas.innerText()).replace('Modal ', 'Canvas '))
    await page.locator('.doc-page.ProseMirror').first().click()
    await expect(header).toContainText('Canvas October 13 changed')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => await currentHeaderXml()).toContain('Canvas ')
    savedXml = await currentHeaderXml()
    expect(savedXml).toContain(fieldXml)
    await undo.click()
    await expect(header).toContainText('Modal October 13 changed')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => await currentHeaderXml()).toContain('Modal ')
    expect(await currentHeaderXml()).toContain(fieldXml)
    await redo.click()
    await expect(header).toContainText('Canvas October 13 changed')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(async () => await currentHeaderXml()).toContain('Canvas ')
    await closeAndSaveVideo(launched, 'docs-hf-date-field-before-reopen')
    launched = undefined
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      openFile: fixture,
      videoDir: 'docs-hf-date-field-reopen',
    })
    page = await waitForPageWithUrl(launched.app, '://docs/')
    const reopenedHeader = page.locator('.page-hf-header:not(.page-gap-hf)').first()
    await expect(reopenedHeader).toContainText('Canvas October 13 changed')
    const reopenedModal = await openHeaderEditor(page)
    const reopenedInput = reopenedModal.locator('textarea')
    const reopenedText = await reopenedInput.inputValue()
    expect(reopenedText).toMatch(/⟦DATE·\d+⟧/)
    expect(reopenedText).toMatch(/⟦PAGE·\d+⟧/)
    expect(reopenedText).toMatch(/⟦NUMPAGES·\d+⟧/)
    await reopenedInput.fill(
      reopenedText.replace('Canvas ', 'Reopened ').replace(' changed\nNew line', ' joined'),
    )
    await reopenedModal.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(reopenedHeader).toContainText('Reopened October 13 joined')
    await expect(reopenedHeader).not.toContainText('New line')
    await expect
      .poll(async () => {
        const headerContentBottom = await reopenedHeader.evaluate((el) => {
          const css = getComputedStyle(el)
          return (
            el.getBoundingClientRect().bottom -
            parseFloat(css.paddingBottom) -
            parseFloat(css.borderBottomWidth)
          )
        })
        const bodyBox = await page!
          .locator('.doc-page.ProseMirror')
          .first()
          .getByText('Section one body', { exact: true })
          .boundingBox()
        return bodyBox ? bodyBox.y - headerContentBottom : -Infinity
      })
      .toBeGreaterThanOrEqual(-1)
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(currentHeaderXml).toContain('Reopened ')
    const evidenceDir = resolve(__dirname, '../docs/qa/office-mac-field-evidence-2026-10-13/docs')
    await mkdir(evidenceDir, { recursive: true })
    await page.screenshot({ path: join(evidenceDir, 'reopened-edit.png') })
    await copyFile(fixture, join(evidenceDir, 'date-field.docx'))
    const pdfPath = join(scratchDir, 'date-field.pdf')
    await launched.app.evaluate(({ dialog, shell }, target) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: target })) as never
      shell.openPath = (async () => '') as never
    }, pdfPath)
    await page.getByRole('button', { name: 'File', exact: true }).click()
    await page.getByRole('button', { name: 'Export', exact: true }).click()
    await page.locator('.bs-exports').getByRole('button', { name: /PDF/ }).click()
    await expect
      .poll(
        async () => {
          try {
            return (await readFile(pdfPath)).subarray(0, 5).toString()
          } catch {
            return ''
          }
        },
        { timeout: 30_000 },
      )
      .toBe('%PDF-')
    await copyFile(pdfPath, join(evidenceDir, 'date-field.pdf'))
    await writeFile(
      join(evidenceDir, 'verification.json'),
      JSON.stringify(
        {
          fixture,
          nativeFieldXmlPreserved: true,
          livePageFieldsPreserved: true,
          tableAndDrawingPreserved: true,
          modalAndCanvasEdits: true,
          newlineSplitJoin: true,
          undoRedoBeforeAndAfterSave: true,
          freshReopenAndEdit: true,
        },
        null,
        2,
      ),
    )
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'docs-hf-date-field')
  }
})
