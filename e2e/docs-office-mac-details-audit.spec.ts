import { existsSync } from 'node:fs'
import { appendFile, mkdir, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Narrow UI + saved OOXML probe for the Word-for-Mac Docs audit. One scratch
 * document/launch exercises the actual GenOffice controls so the audit does
 * not rely on code inspection alone.
 */

const UI_SEED = 'GENOFFICE_DOCS_MAC_AUDIT_SEED_2026-10-12'

test.skip(!process.env.GENOFFICE_OFFICE_AUDIT, 'Opt-in Office detail audit')

async function openRibbonTab(page: Page, name: string): Promise<void> {
  const tab = page.locator('.ribbon-tabs .ribbon-tab').filter({ hasText: new RegExp(`^${name}$`) })
  await expect(tab).toBeVisible()
  if (!(await tab.evaluate((el) => el.classList.contains('active')))) await tab.click()
  await expect(tab).toHaveClass(/active/)
}

async function savedMarkers(path: string) {
  if (!existsSync(path)) return { exists: false }
  try {
    const zip = await JSZip.loadAsync(await readFile(path))
    const documentXml = (await zip.file('word/document.xml')!.async('string')).replace(/\s+/g, ' ')
    const headerPath = Object.keys(zip.files).find((name) => /^word\/header\d+\.xml$/.test(name))
    const headerXml = headerPath ? await zip.file(headerPath)!.async('string') : ''
    return {
      exists: true,
      xmlReady: true,
      headerPath: headerPath ?? null,
      markers: {
        uiSeed: documentXml.includes(UI_SEED),
        pageNumberUpperRomanStart5: /<w:pgNumType[^>]*w:fmt="upperRoman"[^>]*w:start="5"/.test(
          documentXml,
        ),
        headerText: headerXml.includes('Audit header'),
        symbol: documentXml.includes('≤'),
        equationOMMLFraction:
          documentXml.includes('<m:f>') && /<m:t(?:\s[^>]*)?>x<\/m:t>/.test(documentXml),
        paragraphLineSpacing360: /w:spacing[^>]*w:line="360"[^>]*w:lineRule="auto"/.test(
          documentXml,
        ),
      },
    }
  } catch (error) {
    return { exists: true, xmlReady: false, parseError: String(error) }
  }
}

test('Docs detail audit: header/page numbers, symbol/equation, font/paragraph', async () => {
  test.setTimeout(240_000)
  const artifactsRoot =
    process.env.GENOFFICE_DOCS_AUDIT_DIR ?? join(tmpdir(), 'genoffice-docs-mac-details-audit')
  const runId = new Date().toISOString().replace(/[.:]/g, '-')
  const dir = join(artifactsRoot, runId)
  const userDataDir = join(dir, 'user-data')
  await mkdir(userDataDir, { recursive: true })
  const out = join(dir, 'docs-audit.docx')
  const jsonPath = join(dir, 'evidence.json')
  const stageLogPath = join(dir, 'stage-log.jsonl')
  const screenshotPath = join(dir, 'final-state.png')
  const evidence: Record<string, unknown> = {
    status: 'running',
    startedAt: new Date().toISOString(),
    uiSeed: UI_SEED,
    artifactDirectory: dir,
    scratchDocument: out,
    screenshot: screenshotPath,
    stageLog: stageLogPath,
    screenshots: {},
    steps: [],
    pageNumberOptions: null,
    savedMarkers: null,
  }
  let launched: LaunchedApp | undefined
  let p: Page | undefined
  const recordStep = async (step: string) => {
    evidence.steps = [...(evidence.steps as string[]), step]
    await appendFile(stageLogPath, `${JSON.stringify({ at: new Date().toISOString(), step })}\n`)
  }
  const captureUi = async (name: string) => {
    if (!p) return
    const path = join(dir, `${name}.png`)
    await p.screenshot({ path, fullPage: true })
    evidence.screenshots = { ...(evidence.screenshots as Record<string, string>), [name]: path }
  }
  const verifySeedPresent = async (stage: string) => {
    const text = await p!.locator('.doc-page.ProseMirror').first().innerText()
    evidence[`editorTextAfter_${stage}`] = text
    expect(text).toContain(UI_SEED)
    await recordStep(`${stage}: verified typed document seed remains in the editor`)
  }
  try {
    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      userDataDir,
      videoDir: 'docs-office-mac-details-audit',
    })
    p = launched.page
    p.setDefaultTimeout(10_000)
    await launched.page.locator('.quick-card').first().click()
    p = await waitForPageWithUrl(launched.app, '://docs/')
    p.setDefaultTimeout(10_000)
    const editorCanvas = p.locator('.doc-page.ProseMirror').first()
    await expect(editorCanvas).toBeVisible()
    await editorCanvas.click()
    await p.keyboard.type(UI_SEED)
    await p.keyboard.press('Enter')
    await p.keyboard.type('Audit body')
    await p.keyboard.press('Enter')
    await p.keyboard.type('Format target')
    await expect(p.getByText('Format target', { exact: true })).toBeVisible()
    evidence.seedText = await editorCanvas.innerText()
    await recordStep('typed scratch seed and target paragraphs through the actual Docs editor')
    await openRibbonTab(p, 'Insert')

    // Header dropdown: text-only editing currently exposes one field and OK/Cancel.
    await p
      .locator('.ribbon-body button', { hasText: /^Header$/ })
      .first()
      .click()
    const hf = p.locator('.hf-menu')
    await expect(hf.locator('input')).toBeVisible()
    await hf.locator('input').fill('Audit header')
    await captureUi('header-dialog')
    await hf.getByRole('button', { name: 'OK' }).click()
    await recordStep('inserted header text through Header UI')
    await verifySeedPresent('header')

    // Editing a header can activate the Header & Footer contextual ribbon. Return
    // to Insert explicitly before looking up Page Number.
    await openRibbonTab(p, 'Insert')
    await recordStep('reselected Insert after header edit')
    await p.locator('.ribbon-body button[data-tip="Page Number"]').click()
    await recordStep('opened Page Number dropdown')
    await p.locator('[data-rb-panel] button', { hasText: /Format/ }).click()
    const pn = p.locator('.modal', { has: p.getByRole('heading', { name: 'Page Number Format' }) })
    await expect(pn).toBeVisible()
    await pn.locator('.pgnum-row .gs-dd-btn').first().click()
    await captureUi('page-number-format-options')
    const numberOptions = pn.getByRole('option')
    const optionValues = await numberOptions.evaluateAll((els) =>
      els.map((el) => (el as HTMLElement).dataset.value),
    )
    evidence.pageNumberOptions = optionValues
    expect(optionValues).toEqual([
      'decimal',
      'numberInDash',
      'lowerLetter',
      'upperLetter',
      'lowerRoman',
      'upperRoman',
      'chineseCounting',
      'chineseCountingThousand',
      'decimalFullWidth',
      'decimalEnclosedCircle',
      'japaneseLegal',
      'ideographLegalTraditional',
      'taiwaneseCounting',
      'ideographTraditional',
      'ideographZodiac',
    ])
    await pn.getByRole('option', { name: 'I, II, III, …', exact: true }).click()
    await pn.getByText('Start at', { exact: true }).click()
    await pn.locator('input[type="number"]').fill('5')
    await pn.getByRole('button', { name: 'OK' }).click()
    await recordStep('set upperRoman page numbering starting at 5')
    await verifySeedPresent('pageNumber')

    // The 1280px shared test window folds the Symbols group into a flyout whose
    // nested menus close with the flyout. Widen this scratch window so the real
    // ribbon controls and their palettes remain independently operable.
    const appWindowSize = await launched.app.evaluate(({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0]
      win?.setContentSize(2200, 1100)
      return win?.getContentSize()
    })
    evidence.appWindowSize = appWindowSize
    await expect
      .poll(() => p!.locator('.rb-fold-group[aria-label="Symbols"]').getAttribute('data-folded'))
      .toBeNull()
    await recordStep('widened scratch app window to keep Symbols group expanded')

    // The Symbol palette inserts a literal text character; the Equation dialog
    // inserts structured OMML. Re-select Insert after every contextual action.
    await p.getByText('Audit body', { exact: true }).click()
    await p.keyboard.press('End')
    await openRibbonTab(p, 'Insert')
    await p.locator('button[data-tip="Insert a symbol"]').click()
    const symbols = p.locator('.symbol-palette button')
    expect(await symbols.count()).toBe(50)
    await captureUi('symbol-palette')
    await symbols.filter({ hasText: '≤' }).click()
    const symbolPage = p.locator('.doc-page').filter({ hasText: '≤' })
    await expect(symbolPage).toHaveCount(1)
    await recordStep('inserted ≤ from Symbol palette')
    await verifySeedPresent('symbol')

    await p.getByText('Format target', { exact: true }).click()
    await p.keyboard.press('End')
    await openRibbonTab(p, 'Insert')
    await p.locator('button[data-tip="Insert an equation"]').click()
    await p.locator('.equation-gallery-custom').click()
    const eq = p.locator('.equation-modal')
    await eq.locator('input:not([type="checkbox"])').fill('\\frac{x}{y}')
    await expect(eq.locator('.equation-preview')).not.toBeEmpty()
    await captureUi('equation-modal')
    await eq.getByRole('button', { name: 'Insert' }).click()
    evidence.editorTextAfterEquation = await p.locator('.doc-page.ProseMirror').first().innerText()
    await recordStep('inserted fraction from Equation LaTeX dialog')
    await verifySeedPresent('equation')

    // Context-menu dialogs: capture the controls exposed in the app, then use
    // Paragraph's line-spacing control once to verify it changes document XML.
    await p.getByText('Format target', { exact: true }).click({ button: 'right' })
    await p.locator('.ctx-item', { hasText: 'Font…' }).click()
    const font = p.locator('.modal', { has: p.getByRole('heading', { name: 'Font' }) })
    await expect(font).toBeVisible()
    expect(await font.locator('.gs-dd-btn').count()).toBe(3)
    expect(await font.locator('input[type="checkbox"]').count()).toBe(4)
    await captureUi('font-dialog')
    await font.getByRole('button', { name: 'Cancel' }).click()
    evidence.fontControls = { dropdowns: 3, checkboxes: 4 }
    await recordStep('opened Font dialog and captured controls')
    await verifySeedPresent('font')

    await p.getByText('Format target', { exact: true }).click({ button: 'right' })
    await p.locator('.ctx-item', { hasText: 'Paragraph…' }).click()
    const para = p.locator('.modal', { has: p.getByRole('heading', { name: 'Paragraph' }) })
    await expect(para).toBeVisible()
    expect(await para.locator('.gs-dd-btn').count()).toBe(3)
    await captureUi('paragraph-dialog')
    await para.getByRole('button', { name: 'Line spacing' }).click()
    await p.getByRole('option', { name: '1.5 lines' }).click()
    await para.getByRole('button', { name: 'OK' }).click()
    evidence.paragraphControls = { dropdowns: 3, selectedLineSpacing: '1.5 lines' }
    await recordStep('changed Paragraph line spacing to 1.5')
    await verifySeedPresent('paragraph')

    await launched.app.evaluate(({ dialog }, file) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
    }, out)
    await p.keyboard.press('ControlOrMeta+Shift+s')
    const expectedMarkers = {
      uiSeed: true,
      pageNumberUpperRomanStart5: true,
      headerText: true,
      symbol: true,
      equationOMMLFraction: true,
      paragraphLineSpacing360: true,
    }
    await expect
      .poll(
        async () => {
          evidence.savedMarkers = await savedMarkers(out)
          return evidence.savedMarkers
        },
        { timeout: 30_000, intervals: [250, 500, 1_000] },
      )
      .toMatchObject({ exists: true, xmlReady: true, markers: expectedMarkers })
    evidence.savedMarkers = await savedMarkers(out)
    await recordStep('saved DOCX and polled all OOXML markers')

    // Verify the actual file can be reopened in a fresh shell process, not just
    // parsed as a ZIP. The second launch receives the saved DOCX as its file arg.
    await captureUi('saved-document-before-close')
    await closeAndSaveVideo(launched, 'docs-office-mac-details-audit-first-launch')
    launched = undefined
    await recordStep('closed the first Electron launch after saving DOCX')

    launched = await launchShell({
      onboardingSeen: true,
      lang: 'en',
      userDataDir,
      videoDir: 'docs-office-mac-details-audit-reopen',
      openFile: out,
    })
    p = await waitForPageWithUrl(launched.app, '://docs/')
    p.setDefaultTimeout(10_000)
    const reopenedDocument = p.locator('.doc-page.ProseMirror').first()
    await expect(reopenedDocument).toBeVisible()
    await expect(reopenedDocument).toContainText(UI_SEED)
    await expect(reopenedDocument).toContainText('Audit body≤')
    const reopenedFraction = reopenedDocument.locator('mfrac')
    await expect(reopenedFraction).toHaveCount(1)
    evidence.reopenedEquationMathML = await reopenedFraction.evaluate(
      (el) => el.parentElement?.outerHTML ?? el.outerHTML,
    )
    await expect(p.getByText('Audit header', { exact: true })).toBeVisible()
    const reopenedText = await reopenedDocument.innerText()
    evidence.reopenedEditorText = reopenedText
    await captureUi('reopened-document')
    await recordStep('reopened DOCX and verified seed, symbol, fraction rendering, and header')

    await openRibbonTab(p, 'Insert')
    await p.locator('button[data-tip="Page Number"]').click()
    await p.locator('[data-rb-panel] button', { hasText: /Format/ }).click()
    const reopenedPageNumber = p.locator('.modal', {
      has: p.getByRole('heading', { name: 'Page Number Format' }),
    })
    await expect(reopenedPageNumber).toBeVisible()
    await captureUi('reopened-page-number-format')
    const reopenedNumberFormat = await reopenedPageNumber
      .locator('.pgnum-row .gs-dd-btn')
      .first()
      .innerText()
    const reopenedStartAt = await reopenedPageNumber.locator('input[type="number"]').inputValue()
    expect(reopenedNumberFormat).toContain('I, II, III')
    expect(reopenedStartAt).toBe('5')
    evidence.reopenedPageNumber = { format: reopenedNumberFormat, startAt: reopenedStartAt }
    await reopenedPageNumber.getByRole('button', { name: 'Cancel' }).click()
    await recordStep('reopened Page Number Format dialog and verified upperRoman/start=5')
    evidence.status = 'passed'
  } catch (error) {
    evidence.status = 'failed'
    evidence.error =
      error instanceof Error
        ? `${error.name}: ${error.message}\n${error.stack ?? ''}`
        : String(error)
    if (p) {
      evidence.ribbonButtonsAtFailure = await p
        .locator('.ribbon-body button')
        .evaluateAll((buttons) =>
          buttons.map((button) => ({
            text: (button as HTMLElement).innerText,
            title: button.getAttribute('title'),
            ariaLabel: button.getAttribute('aria-label'),
            dataTip: button.getAttribute('data-tip'),
          })),
        )
        .catch((snapshotError) => `Could not inspect ribbon buttons: ${String(snapshotError)}`)
    }
    await appendFile(
      stageLogPath,
      `${JSON.stringify({ at: new Date().toISOString(), error: String(error) })}\n`,
    )
    throw error
  } finally {
    evidence.finishedAt = new Date().toISOString()
    if (p) {
      await p.screenshot({ path: screenshotPath, fullPage: true }).catch((error) => {
        evidence.screenshotError = String(error)
      })
    } else {
      evidence.screenshotError = 'No renderer page was available to capture.'
    }
    await writeFile(jsonPath, JSON.stringify(evidence, null, 2))
    console.log(
      `DOCS_AUDIT_ARTIFACTS ${JSON.stringify({ directory: dir, jsonPath, screenshotPath, out })}`,
    )
    if (launched) await closeAndSaveVideo(launched, 'docs-office-mac-details-audit')
  }
})
