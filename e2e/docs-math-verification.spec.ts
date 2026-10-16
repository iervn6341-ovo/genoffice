import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/** Opt-in desktop acceptance for the Docs math workbench and advanced symbol picker. */
const VERIFY = process.env.GENOFFICE_DOCS_MATH_VERIFICATION
const EVIDENCE_ROOT = resolve(
  process.env.GENOFFICE_DOCS_MATH_VERIFICATION_DIR ??
    join(__dirname, '../docs/qa/office-mac-evidence-2026-10-13/docs-math-verification'),
)
const SYMBOL_ANCHOR = 'SYMBOL_ANCHOR'
const EQUATION_ANCHOR = 'EQUATION_ANCHOR'
const SYMBOL = '𝄞'
const SYMBOL_FONT = 'Segoe UI Symbol'
const INITIAL_LATEX_PREFIX = '\\frac{\\sqrt{a_1^2+x}}{\\sum_{i=0}^{n}{\\frac{x_i}{i!}}}+'
const INITIAL_LATEX = `${INITIAL_LATEX_PREFIX}β`
const EDITED_LATEX = '\\frac{\\sqrt{c_3^4+\\gamma}}{\\prod_{j=1}^{k}{\\frac{z_j}{j!}}}'

async function evidenceDir(): Promise<string> {
  await mkdir(EVIDENCE_ROOT, { recursive: true })
  return mkdtemp(join(EVIDENCE_ROOT, `${new Date().toISOString().replace(/[.:]/g, '-')}-`))
}

async function createScratchDocx(path: string): Promise<void> {
  const bytes = await buildDocx({
    bodyXml:
      `<w:p><w:r><w:rPr><w:b/><w:color w:val="C00000"/><w:sz w:val="28"/></w:rPr><w:t>${SYMBOL_ANCHOR}</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>${EQUATION_ANCHOR}</w:t></w:r></w:p>`,
  })
  await writeFile(path, bytes)
}

async function openEditor(launched: LaunchedApp): Promise<Page> {
  const page = await waitForPageWithUrl(launched.app, '://docs/', 30_000)
  page.setDefaultTimeout(12_000)
  await expect(page.locator('.doc-page.ProseMirror').first()).toBeVisible()
  return page
}

async function openRibbonTab(page: Page, name: string): Promise<void> {
  const tab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: new RegExp(`^${name}$`) })
  await expect(tab).toBeVisible()
  if (!(await tab.evaluate((el) => el.classList.contains('active')))) await tab.click()
  await expect(tab).toHaveClass(/active/)
}

async function openInsertAction(page: Page, tooltip: string): Promise<ReturnType<Page['locator']>> {
  await openRibbonTab(page, 'Insert')
  const group = page.locator('.rb-fold-group[aria-label="Symbols"]')
  await expect(group).toBeVisible()
  if ((await group.getAttribute('data-folded')) !== null) {
    const foldButton = group.locator(':scope > .rb-fold-btn')
    if ((await foldButton.getAttribute('aria-expanded')) !== 'true') await foldButton.click()
  }
  const action = group.locator(`button[data-tip="${tooltip}"]`)
  await expect(action).toBeVisible()
  await action.click()
  return group
}

async function savedDocumentXml(path: string): Promise<{ zip: JSZip; xml: string }> {
  const zip = await JSZip.loadAsync(await readFile(path))
  const xml = (await zip.file('word/document.xml')?.async('string')) ?? ''
  return { zip, xml }
}

async function waitForPackage(path: string, check: (xml: string) => boolean): Promise<void> {
  await expect
    .poll(async () => {
      try {
        const { xml } = await savedDocumentXml(path)
        return check(xml)
      } catch {
        // A save may be replacing the package while Playwright polls.
        return false
      }
    })
    .toBe(true)
}

function mathFragment(xml: string): string {
  return (xml.match(/<m:oMath\b[\s\S]*?<\/m:oMath>/) ?? [])[0] ?? ''
}

function symbolRun(xml: string): string {
  return (xml.match(/<w:r\b[\s\S]*?<\/w:r>/g) ?? []).find((run) => run.includes(SYMBOL)) ?? ''
}

function expectNativeMath(xml: string, expectedGreek: string): void {
  const formula = mathFragment(xml)
  expect(formula).not.toBe('')
  expect(formula).toMatch(/<m:f\b/)
  expect(formula).toMatch(/<m:rad\b/)
  expect(formula).toMatch(/<m:nary\b/)
  expect(formula).toMatch(/<m:s(?:Sub|Sup|SubSup)\b/)
  expect(formula).toContain(expectedGreek)
}

function expectStyledSymbol(xml: string): void {
  const run = symbolRun(xml)
  expect(run).not.toBe('')
  expect(run).toMatch(/<w:rFonts\b[^>]*\bw:ascii="Segoe UI Symbol"/)
  expect(run).toMatch(/<w:b(?:\s|\/|>)/)
  expect(run).toMatch(/<w:color\b[^>]*\bw:val="C00000"/)
  expect(run).toMatch(/<w:sz\b[^>]*\bw:val="28"/)
}

test.skip(
  !VERIFY,
  'Opt-in Docs math and symbol verification; set GENOFFICE_DOCS_MATH_VERIFICATION=1 after build:all',
)

test('Docs math and symbols: nested equation and styled font symbol survive save, edit, and reopen', async () => {
  test.setTimeout(360_000)
  const evidence = await evidenceDir()
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-docs-math-'))
  const documentPath = join(scratch, 'math-symbol-roundtrip.docx')
  const userDataDir = join(scratch, 'user-data')
  await mkdir(userDataDir, { recursive: true })
  await createScratchDocx(documentPath)

  const record: Record<string, unknown> = {
    scenario:
      '1280px Docs equation authoring with nested fraction/radical/script/sum and workbench Greek insertion; styled font-symbol insertion; save/relaunch/re-edit/resave/relaunch',
    startedAt: new Date().toISOString(),
    evidenceDir: evidence,
    scratchDocument: documentPath,
    userDataDir,
    steps: [],
  }
  const step = (value: string) => {
    record.steps = [...(record.steps as string[]), value]
  }
  let launched: LaunchedApp | undefined
  let currentVideo = 'docs-math-verification-insert'
  let page: Page | undefined

  try {
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      lang: 'en',
      openFile: documentPath,
      videoDir: currentVideo,
    })
    page = await openEditor(launched)
    const editor = page.locator('.doc-page.ProseMirror').first()
    await expect(editor).toContainText(SYMBOL_ANCHOR)
    await expect(editor).toContainText(EQUATION_ANCHOR)

    await editor.getByText(EQUATION_ANCHOR, { exact: true }).click()
    await page.keyboard.press('End')
    const group = await openInsertAction(page, 'Insert an equation')
    await expect(group.locator('.equation-gallery')).toBeVisible()
    await group.locator('.equation-gallery-custom').click()

    const equation = page.locator('.equation-modal')
    await expect(equation).toBeVisible()
    const input = equation.locator('textarea.go-math-input')
    await input.fill(INITIAL_LATEX_PREFIX)
    await input.evaluate((element: HTMLTextAreaElement) => {
      element.focus()
      element.setSelectionRange(element.value.length, element.value.length)
    })
    await equation
      .locator('.go-math-categories')
      .getByRole('tab', { name: 'Mathematical symbols', exact: true })
      .click()
    await equation
      .locator('.go-math-symbol-subsets')
      .getByRole('tab', { name: 'Greek', exact: true })
      .click()
    await equation
      .locator('.go-math-symbol-grid button[aria-label*="GREEK SMALL LETTER BETA"]')
      .click()
    await expect(input).toHaveValue(INITIAL_LATEX)
    const preview = await equation.locator('.go-math-preview').innerHTML()
    expect(preview).toContain('<mfrac')
    expect(preview).toContain('<msqrt')
    expect(preview).toContain('β')
    await equation.getByRole('checkbox').check()
    await page.screenshot({
      path: join(evidence, 'equation-dialog-narrow-1280.png'),
      fullPage: true,
    })
    await equation.getByRole('button', { name: 'Insert', exact: true }).click()
    await expect(equation).toHaveCount(0)
    const inlineMath = editor.locator('.doc-inline-math')
    await expect(inlineMath).toHaveCount(1)
    await expect(inlineMath.locator('math')).toBeVisible()
    await expect(inlineMath).toContainText('β')
    step(
      'At the 1280px app window, a nested inline equation was authored and β was inserted from the Greek workbench gallery.',
    )

    await editor.getByText(SYMBOL_ANCHOR, { exact: true }).click()
    await page.keyboard.press('End')
    const symbolGroup = await openInsertAction(page, 'Insert a symbol')
    await expect(symbolGroup.locator('.symbol-palette')).toBeVisible()
    await symbolGroup.locator('.symbol-more').click()
    const symbolDialog = page.locator('.docs-symbol-modal')
    await expect(symbolDialog).toBeVisible()
    await symbolDialog.locator('.go-symbol-picker select').first().selectOption(SYMBOL_FONT)
    await symbolDialog
      .locator('.go-symbol-bottom input[aria-label="Character code"]')
      .fill('U+1D11E')
    await symbolDialog.locator('.go-symbol-insert').click()
    await expect(symbolDialog).toHaveCount(0)
    await expect(editor).toContainText(`${SYMBOL_ANCHOR}${SYMBOL}`)
    step('The advanced picker inserted U+1D11E at the styled cursor with Segoe UI Symbol selected.')

    await page.keyboard.press('ControlOrMeta+s')
    await waitForPackage(documentPath, (xml) => xml.includes('<m:oMath') && xml.includes(SYMBOL))
    let packageState = await savedDocumentXml(documentPath)
    expectNativeMath(packageState.xml, 'β')
    expectStyledSymbol(packageState.xml)
    await copyFile(documentPath, join(evidence, 'initial-saved.docx'))
    record.initialPackage = {
      mathXml: mathFragment(packageState.xml),
      symbolRun: symbolRun(packageState.xml),
    }
    step(
      'Initial save contains structured OMML fraction, radical, n-ary operator, scripts, β, and a bold red 14pt Segoe UI Symbol run.',
    )

    await closeAndSaveVideo(launched, currentVideo)
    launched = undefined
    currentVideo = 'docs-math-verification-first-reopen'
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      lang: 'en',
      openFile: documentPath,
      videoDir: currentVideo,
    })
    page = await openEditor(launched)
    const reopenedEditor = page.locator('.doc-page.ProseMirror').first()
    const reopenedMath = reopenedEditor.locator('.doc-inline-math')
    await expect(reopenedMath).toHaveCount(1)
    await expect(reopenedMath).toContainText('β')
    await page.screenshot({ path: join(evidence, 'first-reopen-narrow-1280.png'), fullPage: true })

    await reopenedMath.dblclick()
    const editDialog = page.locator('.equation-modal')
    await expect(editDialog).toBeVisible()
    const editInput = editDialog.locator('textarea.go-math-input')
    const recoveredSource = await editInput.inputValue()
    expect(recoveredSource).toContain('\\frac')
    expect(recoveredSource).toContain('\\sqrt')
    expect(recoveredSource).toContain('\\sum')
    expect(recoveredSource).toContain('\\beta')
    record.recoveredSource = recoveredSource
    await page.screenshot({
      path: join(evidence, 'equation-reopened-for-edit.png'),
      fullPage: true,
    })
    await editInput.fill(EDITED_LATEX)
    await editDialog.getByRole('button', { name: 'Update', exact: true }).click()
    await expect(editDialog).toHaveCount(0)
    await expect(reopenedMath).toContainText('γ')
    await page.keyboard.press('ControlOrMeta+s')
    await waitForPackage(documentPath, (xml) => mathFragment(xml).includes('γ'))
    packageState = await savedDocumentXml(documentPath)
    expectNativeMath(packageState.xml, 'γ')
    expect(mathFragment(packageState.xml)).not.toContain('β')
    expectStyledSymbol(packageState.xml)
    await copyFile(documentPath, join(evidence, 'edited-saved.docx'))
    record.editedPackage = {
      mathXml: mathFragment(packageState.xml),
      symbolRun: symbolRun(packageState.xml),
    }
    step(
      'After relaunch, the equation reopened from its DOCX OMML, was edited to a new nested product/radical/fraction, and saved with the styled symbol intact.',
    )

    await closeAndSaveVideo(launched, currentVideo)
    launched = undefined
    currentVideo = 'docs-math-verification-final-reopen'
    launched = await launchShell({
      onboardingSeen: true,
      userDataDir,
      lang: 'en',
      openFile: documentPath,
      videoDir: currentVideo,
    })
    page = await openEditor(launched)
    const finalEditor = page.locator('.doc-page.ProseMirror').first()
    const finalMath = finalEditor.locator('.doc-inline-math')
    await expect(finalMath).toHaveCount(1)
    await expect(finalMath).toContainText('γ')
    await expect(finalEditor).toContainText(`${SYMBOL_ANCHOR}${SYMBOL}`)
    const finalPackage = await savedDocumentXml(documentPath)
    expectNativeMath(finalPackage.xml, 'γ')
    expectStyledSymbol(finalPackage.xml)
    await page.screenshot({ path: join(evidence, 'final-reopen-narrow-1280.png'), fullPage: true })
    record.finalPackage = {
      mathXml: mathFragment(finalPackage.xml),
      symbolRun: symbolRun(finalPackage.xml),
    }
    record.status = 'passed'
    step(
      'A second relaunch displays the edited equation and font/style-preserving symbol; the final DOCX retains native OMML and run properties.',
    )
  } catch (error) {
    record.status = 'failed'
    record.error = error instanceof Error ? `${error.name}: ${error.message}` : String(error)
    if (page) {
      await page
        .screenshot({ path: join(evidence, 'failure.png'), fullPage: true })
        .catch(() => undefined)
    }
    throw error
  } finally {
    record.finishedAt = new Date().toISOString()
    await writeFile(join(evidence, 'evidence.json'), JSON.stringify(record, null, 2))
    if (launched) await closeAndSaveVideo(launched, currentVideo)
  }
})
