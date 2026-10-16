import { createHash } from 'node:crypto'
import { mkdir, mkdtemp, readdir, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, posix } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'

import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

const RIBBON_TAB = 'nav.ribbon-tabs > button:not(.qa-btn):not(.ribbon-tab-file)'
const EVIDENCE_DIRECTORY =
  process.env.GENOFFICE_SHEETS_MATH_VERIFICATION_DIR ??
  join(process.cwd(), 'docs/qa/office-mac-math-evidence-2026-10-13/sheets')

type Launched = Awaited<ReturnType<typeof launchShell>>
type SavedPackagePredicate = (zip: JSZip) => Promise<boolean>
const saveSequences = new Map<string, number>()

test.skip(!process.env.GENOFFICE_OFFICE_AUDIT, 'Opt-in Sheets native math verification')
test.use({ actionTimeout: 10_000 })

async function waitForSheetsPage(launched: Launched): Promise<Page> {
  const sheet = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await sheet.waitForFunction(() => document.body.textContent?.includes('Sheet1'), null, {
    timeout: 30_000,
  })
  await sheet.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI), null, {
    timeout: 30_000,
  })
  await sheet.waitForSelector('[data-ribbon-body]', { timeout: 20_000 })
  return sheet
}

async function openBlank(videoDir: string): Promise<{
  launched: Launched
  sheet: Page
  workbook: string
}> {
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-sheets-math-verify-'))
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  await launched.app.evaluate(({ app }, path) => app.setPath('documents', path), scratch)
  await launched.page.locator('.quick-card').nth(1).click()
  const sheet = await waitForSheetsPage(launched)
  await setEditorLayoutWidth(launched.app, '://sheets/', 1280)
  const files = await readdir(join(scratch, 'GenOffice'))
  const workbooks = files.filter((file) => file.endsWith('.xlsx'))
  expect(workbooks).toHaveLength(1)
  return { launched, sheet, workbook: join(scratch, 'GenOffice', workbooks[0]!) }
}

async function openExisting(
  workbook: string,
  videoDir: string,
): Promise<{ launched: Launched; sheet: Page }> {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir,
    openFile: workbook,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  const sheet = await waitForSheetsPage(launched)
  await setEditorLayoutWidth(launched.app, '://sheets/', 1280)
  return { launched, sheet }
}

const selectTab = (page: Page, name: string): Promise<void> =>
  page
    .locator(RIBBON_TAB, { hasText: new RegExp(`^${name}$`) })
    .first()
    .click()

async function packageXml(zip: JSZip, path: string): Promise<string> {
  return (await zip.file(path)?.async('string')) ?? ''
}

async function saveFromApp(
  launched: Launched,
  workbook: string,
  expected: SavedPackagePredicate,
): Promise<JSZip> {
  await launched.app.evaluate(({ webContents }) => {
    const wc = webContents.getAllWebContents().find((item) => item.getURL().includes('://sheets/'))
    wc?.send('menu:action', 'save')
  })
  await expect
    .poll(
      async () => {
        try {
          return await expected(await JSZip.loadAsync(await readFile(workbook)))
        } catch {
          // The save is asynchronous; the first poll can run before the new
          // drawing part or workbook bytes have reached the scratch file.
          return false
        }
      },
      {
        timeout: 25_000,
      },
    )
    .toBe(true)
  const bytes = await readFile(workbook)
  const title = test
    .info()
    .title.toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
  const sequence = (saveSequences.get(title) ?? 0) + 1
  saveSequences.set(title, sequence)
  await attachArtifact(
    `${title}-saved-${sequence}.xlsx`,
    bytes,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  )
  return JSZip.loadAsync(bytes)
}

async function attachScreenshot(page: Page, name: string): Promise<void> {
  await attachArtifact(name, await page.screenshot(), 'image/png')
}

async function attachArtifact(
  name: string,
  body: Buffer | string,
  contentType: string,
): Promise<void> {
  await mkdir(EVIDENCE_DIRECTORY, { recursive: true })
  await test.info().attach(name, { body, contentType })
  await writeFile(join(EVIDENCE_DIRECTORY, name), body)
}

async function close(launched: Launched, videoDir: string): Promise<void> {
  await closeAndSaveVideo(launched, videoDir)
}

function packageRelationships(xml: string): { id: string; target: string }[] {
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?\s*>/g)].flatMap((match) => {
    const attrs = Object.fromEntries(
      [...(match[1] ?? '').matchAll(/([\w:.-]+)="([^"]*)"/g)].map((attr) => [attr[1]!, attr[2]!]),
    )
    return attrs['Id'] && attrs['Target'] ? [{ id: attrs['Id'], target: attrs['Target'] }] : []
  })
}

async function drawingPart(zip: JSZip): Promise<{ path: string; xml: string }> {
  const sheet = await packageXml(zip, 'xl/worksheets/sheet1.xml')
  const relationshipId = /<drawing\b[^>]*\br:id="([^"]+)"/.exec(sheet)?.[1]
  if (!relationshipId) throw new Error('The worksheet has no drawing relationship.')
  const rels = await packageXml(zip, 'xl/worksheets/_rels/sheet1.xml.rels')
  const relationship = packageRelationships(rels).find((item) => item.id === relationshipId)
  if (!relationship) throw new Error(`Drawing relationship ${relationshipId} is missing.`)
  const path = relationship.target.startsWith('/')
    ? relationship.target.slice(1)
    : posix.normalize(posix.join('xl/worksheets', relationship.target))
  const xml = await packageXml(zip, path)
  if (!xml) throw new Error(`Drawing part ${path} is empty or missing.`)
  return { path, xml }
}

function drawingAnchors(xml: string): string[] {
  return [
    ...xml.matchAll(
      /<((?:[\w.-]+:)?)((?:twoCellAnchor|oneCellAnchor|absoluteAnchor))\b[\s\S]*?<\/\1\2>/g,
    ),
  ].map((match) => match[0])
}

function officeMath(anchor: string): string {
  return /<(?:[\w.-]+:)?oMath\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?oMath>/.exec(anchor)?.[1] ?? ''
}

function anchorMarkers(anchor: string): { from: string; to: string } {
  const marker = (tag: 'from' | 'to') =>
    new RegExp(`<((?:[\\w.-]+:)?)${tag}\\b[^>]*>([\\s\\S]*?)<\\/\\1${tag}>`).exec(anchor)?.[2] ?? ''
  return { from: marker('from'), to: marker('to') }
}

function equationMetadata(anchor: string): { shapeId: string; source: string; digest: string } {
  const shapeId = /<(?:[\w.-]+:)?cNvPr\b[^>]*\bid="(\d+)"/.exec(anchor)?.[1]
  const metadata = /\bdescr="GenOfficeEquationV2:([A-Za-z0-9+/=]+):([a-f0-9]{64})"/.exec(anchor)
  if (!shapeId || !metadata) throw new Error('The equation is missing its stable shape metadata.')
  return {
    shapeId,
    source: Buffer.from(metadata[1]!, 'base64').toString('utf8'),
    digest: metadata[2]!,
  }
}

async function equationDialog(page: Page) {
  await clickSymbolsCommand(page, 'Equation')
  const dialog = page.getByRole('dialog', { name: 'Insert Equation' })
  await expect(dialog).toBeVisible()
  return dialog
}

async function clickSymbolsCommand(page: Page, name: 'Equation' | 'Symbol'): Promise<void> {
  await selectTab(page, 'Insert')
  const group = page.locator('.rb-fold-group[aria-label="Symbols"]')
  await expect(group).toBeVisible()
  if ((await group.getAttribute('data-folded')) !== null) {
    const foldButton = group.locator(':scope > .rb-fold-btn')
    if ((await foldButton.getAttribute('aria-expanded')) !== 'true') await foldButton.click()
  }
  await group.locator(`button[data-tip="${name}"]`).click()
}

async function selectLatex(
  source: ReturnType<typeof pageLocator>,
  fragment: string,
): Promise<void> {
  const value = await source.inputValue()
  const start = value.indexOf(fragment)
  if (start < 0) throw new Error(`Could not find ${JSON.stringify(fragment)} in LaTeX ${value}.`)
  await source.evaluate(
    (element, range) => {
      const textarea = element as HTMLTextAreaElement
      textarea.focus()
      textarea.setSelectionRange(range.start, range.end)
    },
    { start, end: start + fragment.length },
  )
}

function pageLocator(page: Page) {
  return page.getByLabel('LaTeX source')
}

async function liveRichCell(page: Page): Promise<{
  value: unknown
  dataStream?: string
  textRuns: { st?: number; ed?: number; ts?: { ff?: string; bl?: number } }[]
}> {
  return page.evaluate(() => {
    const api = (window as any).__genofficeDebug.univerAPI
    const range = api.getActiveWorkbook().getActiveSheet().getRange('A1')
    const cell = range.getCellDatas()[0]?.[0] as
      | {
          p?: {
            body?: {
              dataStream?: string
              textRuns?: { st?: number; ed?: number; ts?: { ff?: string; bl?: number } }[]
            }
          }
        }
      | undefined
    const dataStream = cell?.p?.body?.dataStream
    const documentText =
      dataStream === undefined
        ? undefined
        : dataStream.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n+$/, '')
    return {
      // Univer can leave getValue() null for imported rich text until its lazy
      // cell reader hydrates. The document stream is the authoritative live
      // text in that case; persisted XLSX rich runs are asserted separately.
      value: range.getValue() ?? documentText,
      ...(dataStream === undefined ? {} : { dataStream }),
      textRuns: cell?.p?.body?.textRuns ?? [],
    }
  })
}

async function insertSymbol(page: Page, font: string, codePoint: string): Promise<void> {
  await clickSymbolsCommand(page, 'Symbol')
  const dialog = page.getByRole('dialog', { name: 'Symbol' })
  await expect(dialog).toBeVisible()
  await dialog.getByLabel('Font', { exact: true }).selectOption(font)
  await dialog.getByLabel('Character code').fill(codePoint)
  await dialog.getByRole('button', { name: 'Insert', exact: true }).click()
}

function richRunsFromWorksheet(xml: string): { font: string; text: string; bold: boolean }[] {
  const cell = /<c\b(?=[^>]*\br="A1")[^>]*>([\s\S]*?)<\/c>/.exec(xml)?.[1] ?? ''
  return [...cell.matchAll(/<r\b[^>]*>([\s\S]*?)<\/r>/g)].map((match) => {
    const run = match[1] ?? ''
    const font = /<(?:[\w.-]+:)?rFont\b[^>]*\bval="([^"]*)"/.exec(run)?.[1] ?? ''
    const text = [...run.matchAll(/<(?:[\w.-]+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
      .map((item) => item[1] ?? '')
      .join('')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&')
    return { font, text, bold: /<(?:[\w.-]+:)?b(?:\s[^>]*)?\/>/.test(run) }
  })
}

test('editable native equation survives save, reopen, F2 edit, and a second reopen', async () => {
  test.setTimeout(300_000)
  const { launched, sheet, workbook } = await openBlank('sheets-math-native-equation')
  let latex: string
  try {
    await sheet.evaluate(() => {
      const api = (window as any).__genofficeDebug.univerAPI
      api.getActiveWorkbook().getActiveSheet().getRange('B3').activate()
    })
    const dialog = await equationDialog(sheet)
    // Build the expression with the actual catalog UI to exercise selection
    // retention between fraction, radical, and script insertions.
    await dialog.getByRole('tab', { name: 'Fraction', exact: true }).click()
    const source = pageLocator(sheet)
    await source.fill('x')
    await selectLatex(source, 'x')
    await dialog.locator('.go-math-template').nth(0).click()
    await selectLatex(source, 'x')
    await dialog.getByRole('tab', { name: 'Radical', exact: true }).click()
    await dialog.locator('.go-math-template').nth(0).click()
    await selectLatex(source, 'x')
    await dialog.getByRole('tab', { name: 'Script', exact: true }).click()
    await dialog.locator('.go-math-template').nth(0).click()
    const structuredLatex = await source.inputValue()
    expect(structuredLatex).toBe('\\frac{\\sqrt{x^{2}}}{y}')
    await source.evaluate((element) => {
      const textarea = element as HTMLTextAreaElement
      textarea.focus()
      textarea.setSelectionRange(textarea.value.length, textarea.value.length)
    })
    await dialog.getByRole('tab', { name: 'Symbol', exact: true }).click()
    await dialog.getByRole('tab', { name: 'Greek', exact: true }).click()
    await dialog.getByRole('button', { name: /^Ω/ }).click()
    latex = await source.inputValue()
    expect(latex).toBe(`${structuredLatex}Ω`)
    await attachScreenshot(sheet, 'sheets-math-narrow-equation-editor.png')
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()
    const visual = sheet
      .locator('.xlsx-equation')
      .locator('xpath=ancestor::*[@data-print-visual][1]')
    await expect(visual).toBeVisible({ timeout: 20_000 })

    const saved = await saveFromApp(launched, workbook, async (zip) => {
      const drawing = await drawingPart(zip)
      const anchors = drawingAnchors(drawing.xml)
      return anchors.length === 1 && officeMath(anchors[0]!).includes('<m:f>')
    })
    const firstDrawing = await drawingPart(saved)
    const firstAnchors = drawingAnchors(firstDrawing.xml)
    expect(firstAnchors).toHaveLength(1)
    const firstAnchor = firstAnchors[0]!
    const firstOmml = officeMath(firstAnchor)
    const firstMetadata = equationMetadata(firstAnchor)
    expect(firstMetadata.source).toBe(latex)
    expect(firstMetadata.digest).toBe(createHash('sha256').update(firstOmml).digest('hex'))
    expect(firstOmml).toMatch(/<(?:[\w.-]+:)?f\b/)
    expect(firstOmml).toMatch(/<(?:[\w.-]+:)?rad\b/)
    expect(firstOmml).toMatch(/<(?:[\w.-]+:)?sSup\b/)
    await attachArtifact('sheets-math-inserted-drawing.xml', firstDrawing.xml, 'application/xml')
  } finally {
    await close(launched, 'sheets-math-native-equation')
  }

  const reopened = await openExisting(workbook, 'sheets-math-native-equation-reopen')
  let stableVisualId: string
  let stableMarkers: { from: string; to: string }
  const editedLatex = `${latex} + 1`
  try {
    const visual = reopened.sheet
      .locator('.xlsx-equation')
      .locator('xpath=ancestor::*[@data-print-visual][1]')
    await expect(visual).toBeVisible({ timeout: 20_000 })
    stableVisualId = (await visual.getAttribute('data-print-visual')) ?? ''
    const drawing = await drawingPart(await JSZip.loadAsync(await readFile(workbook)))
    const anchor = drawingAnchors(drawing.xml)[0]
    expect(anchor).toBeTruthy()
    stableMarkers = anchorMarkers(anchor!)
    const shapeId = equationMetadata(anchor!).shapeId

    const editable = visual.locator('.shape-editable')
    await editable.focus()
    await expect(editable).toHaveClass(/selected/)
    // Double-click and F2 both route into the native equation editor.
    await editable.dblclick()
    let dialog = reopened.sheet.getByRole('dialog', { name: 'Insert Equation' })
    await expect(dialog).toBeVisible()
    await expect(reopened.sheet.getByLabel('LaTeX source')).toHaveValue(latex)
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    await expect(dialog).toBeHidden()

    await editable.focus()
    await editable.press('F2')
    dialog = reopened.sheet.getByRole('dialog', { name: 'Insert Equation' })
    await expect(dialog).toBeVisible()
    await expect(reopened.sheet.getByLabel('LaTeX source')).toHaveValue(latex)
    await reopened.sheet.getByLabel('LaTeX source').fill(editedLatex)
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()
    await expect(reopened.sheet.locator('.xlsx-equation')).toContainText('+1')
    await attachScreenshot(reopened.sheet, 'sheets-math-reopened-edit.png')

    const saved = await saveFromApp(reopened.launched, workbook, async (zip) => {
      const drawing = await drawingPart(zip)
      const equation = drawingAnchors(drawing.xml)[0]
      return Boolean(equation && equationMetadata(equation).source === editedLatex)
    })
    const savedDrawing = await drawingPart(saved)
    const savedAnchor = drawingAnchors(savedDrawing.xml)[0]
    expect(savedAnchor).toBeTruthy()
    expect(equationMetadata(savedAnchor!).shapeId).toBe(shapeId)
    expect(anchorMarkers(savedAnchor!)).toEqual(stableMarkers)
    expect(equationMetadata(savedAnchor!).source).toBe(editedLatex)
    await attachArtifact('sheets-math-edited-drawing.xml', savedDrawing.xml, 'application/xml')
  } finally {
    await close(reopened.launched, 'sheets-math-native-equation-reopen')
  }

  const secondReopen = await openExisting(workbook, 'sheets-math-native-equation-second-reopen')
  try {
    const visual = secondReopen.sheet
      .locator('.xlsx-equation')
      .locator('xpath=ancestor::*[@data-print-visual][1]')
    await expect(visual).toBeVisible({ timeout: 20_000 })
    await expect(visual).toHaveAttribute('data-print-visual', stableVisualId)
    const editable = visual.locator('.shape-editable')
    await editable.focus()
    await editable.press('F2')
    const dialog = secondReopen.sheet.getByRole('dialog', { name: 'Insert Equation' })
    await expect(dialog).toBeVisible()
    await expect(secondReopen.sheet.getByLabel('LaTeX source')).toHaveValue(editedLatex)
    await dialog.getByRole('button', { name: 'Cancel' }).click()
    const zip = await JSZip.loadAsync(await readFile(workbook))
    const drawing = await drawingPart(zip)
    const anchor = drawingAnchors(drawing.xml)[0]
    expect(anchor).toBeTruthy()
    expect(anchorMarkers(anchor!)).toEqual(stableMarkers)
    const omml = officeMath(anchor!)
    expect(omml).toMatch(/<(?:[\w.-]+:)?f\b/)
    expect(omml).toMatch(/<(?:[\w.-]+:)?rad\b/)
    expect(omml).toMatch(/<(?:[\w.-]+:)?sSup\b/)
    await attachScreenshot(secondReopen.sheet, 'sheets-math-second-reopen.png')
    await attachArtifact(
      'sheets-math-native-equation-before-unsupported.xlsx',
      await readFile(workbook),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    await attachArtifact(
      'sheets-math-native-equation-summary.json',
      JSON.stringify(
        {
          workbook,
          visualId: stableVisualId,
          anchor: stableMarkers,
          shapeId: equationMetadata(anchor!).shapeId,
          insertedLatex: latex,
          editedLatex,
          hasFraction: /<(?:[\w.-]+:)?f\b/.test(omml),
          hasRadical: /<(?:[\w.-]+:)?rad\b/.test(omml),
          hasScript: /<(?:[\w.-]+:)?sSup\b/.test(omml),
        },
        null,
        2,
      ),
      'application/json',
    )
  } finally {
    await close(secondReopen.launched, 'sheets-math-native-equation-second-reopen')
  }

  // Mutate the scratch file like an external Office producer added a structure
  // outside the authorable subset. Stale LaTeX metadata must not make it editable.
  const mutated = await JSZip.loadAsync(await readFile(workbook))
  const drawing = await drawingPart(mutated)
  const changedXml = drawing.xml.replace(
    /(<(?:[\w.-]+:)?oMath\b[^>]*>)[\s\S]*?(<\/(?:[\w.-]+:)?oMath>)/,
    '$1<m:box><m:e><m:r><m:t>x</m:t></m:r></m:e></m:box>$2',
  )
  expect(changedXml).not.toBe(drawing.xml)
  mutated.file(drawing.path, changedXml)
  await writeFile(workbook, await mutated.generateAsync({ type: 'nodebuffer' }))
  await attachArtifact(
    'sheets-math-native-equation-unsupported-scratch.xlsx',
    await readFile(workbook),
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  )
  const unsupported = await openExisting(workbook, 'sheets-math-unsupported-readonly')
  try {
    const visual = unsupported.sheet
      .locator('.xlsx-equation')
      .locator('xpath=ancestor::*[@data-print-visual][1]')
    await expect(visual).toBeVisible({ timeout: 20_000 })
    const editable = visual.locator('.shape-editable')
    await editable.focus()
    await editable.press('F2')
    const dialog = unsupported.sheet.getByRole('dialog', { name: 'Insert Equation' })
    await expect(dialog).toBeVisible()
    await expect(
      dialog.getByText(
        'This Office equation uses math structures that cannot be edited safely here.',
      ),
    ).toBeVisible()
    await expect(dialog.getByRole('button', { name: 'OK' })).toBeDisabled()
    await attachScreenshot(unsupported.sheet, 'sheets-math-unsupported-readonly.png')
  } finally {
    await close(unsupported.launched, 'sheets-math-unsupported-readonly')
  }
})

test('symbol insertion preserves existing rich text and appends chosen-font runs across reopen', async () => {
  test.setTimeout(240_000)
  const { launched, sheet, workbook } = await openBlank('sheets-math-rich-symbol')
  try {
    await sheet.evaluate(() => {
      const api = (window as any).__genofficeDebug.univerAPI
      const range = api.getActiveWorkbook().getActiveSheet().getRange('A1')
      range.setValue({
        v: 'Styled',
        t: 1,
        f: null,
        p: {
          id: 'rich-cell',
          body: {
            dataStream: 'Styled\r\n',
            textRuns: [{ st: 0, ed: 6, ts: { ff: 'Arial', bl: 1 } }],
            paragraphs: [{ startIndex: 6 }],
            sectionBreaks: [{ startIndex: 7 }],
          },
          documentStyle: {},
        },
      })
      range.activate()
    })
    await expect.poll(() => liveRichCell(sheet)).toMatchObject({ value: 'Styled' })
    await insertSymbol(sheet, 'Segoe UI Symbol', '03A9')
    await expect.poll(() => liveRichCell(sheet)).toMatchObject({ value: 'StyledΩ' })
    const firstLive = await liveRichCell(sheet)
    expect(firstLive.textRuns).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          st: 0,
          ed: 6,
          ts: expect.objectContaining({ ff: 'Arial', bl: 1 }),
        }),
        expect.objectContaining({
          st: 6,
          ed: 7,
          ts: expect.objectContaining({ ff: 'Segoe UI Symbol', bl: 1 }),
        }),
      ]),
    )
    await attachScreenshot(sheet, 'sheets-math-symbol-chosen-font.png')
    await sheet
      .getByRole('dialog', { name: 'Symbol' })
      .getByRole('button', { name: 'Close' })
      .click()

    const saved = await saveFromApp(launched, workbook, async (zip) => {
      const xml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
      const runs = richRunsFromWorksheet(xml)
      return (
        runs.some((run) => run.text === 'Styled' && run.font === 'Arial' && run.bold) &&
        runs.some((run) => run.text === 'Ω' && run.font === 'Segoe UI Symbol' && run.bold)
      )
    })
    const runs = richRunsFromWorksheet(await packageXml(saved, 'xl/worksheets/sheet1.xml'))
    expect(runs).toEqual([
      { font: 'Arial', text: 'Styled', bold: true },
      { font: 'Segoe UI Symbol', text: 'Ω', bold: true },
    ])
    await attachArtifact(
      'sheets-math-symbol-first-save-sheet1.xml',
      await packageXml(saved, 'xl/worksheets/sheet1.xml'),
      'application/xml',
    )
  } finally {
    await close(launched, 'sheets-math-rich-symbol')
  }

  const reopened = await openExisting(workbook, 'sheets-math-rich-symbol-reopen')
  try {
    const readCell = () => liveRichCell(reopened.sheet)
    await expect.poll(readCell).toMatchObject({ value: 'StyledΩ' })
    await reopened.sheet.evaluate(() => {
      const api = (window as any).__genofficeDebug.univerAPI
      api.getActiveWorkbook().getActiveSheet().getRange('A1').activate()
    })
    await insertSymbol(reopened.sheet, 'Cambria Math', '2211')
    await expect.poll(readCell).toMatchObject({ value: 'StyledΩ∑' })
    const appended = await readCell()
    expect(appended.textRuns).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          st: 0,
          ed: 6,
          ts: expect.objectContaining({ ff: 'Arial', bl: 1 }),
        }),
        expect.objectContaining({
          st: 6,
          ed: 7,
          ts: expect.objectContaining({ ff: 'Segoe UI Symbol', bl: 1 }),
        }),
        expect.objectContaining({
          st: 7,
          ed: 8,
          ts: expect.objectContaining({ ff: 'Cambria Math', bl: 1 }),
        }),
      ]),
    )
    await reopened.sheet
      .getByRole('dialog', { name: 'Symbol' })
      .getByRole('button', { name: 'Close' })
      .click()
    const saved = await saveFromApp(reopened.launched, workbook, async (zip) => {
      const xml = await packageXml(zip, 'xl/worksheets/sheet1.xml')
      const runs = richRunsFromWorksheet(xml)
      return (
        runs.some((run) => run.text === 'Styled' && run.font === 'Arial' && run.bold) &&
        runs.some((run) => run.text === 'Ω' && run.font === 'Segoe UI Symbol' && run.bold) &&
        runs.some((run) => run.text === '∑' && run.font === 'Cambria Math' && run.bold)
      )
    })
    expect(richRunsFromWorksheet(await packageXml(saved, 'xl/worksheets/sheet1.xml'))).toEqual([
      { font: 'Arial', text: 'Styled', bold: true },
      { font: 'Segoe UI Symbol', text: 'Ω', bold: true },
      { font: 'Cambria Math', text: '∑', bold: true },
    ])
    await attachScreenshot(reopened.sheet, 'sheets-math-symbol-reopened-appended.png')
  } finally {
    await close(reopened.launched, 'sheets-math-rich-symbol-reopen')
  }

  const secondReopen = await openExisting(workbook, 'sheets-math-rich-symbol-second-reopen')
  try {
    await expect.poll(() => liveRichCell(secondReopen.sheet)).toMatchObject({ value: 'StyledΩ∑' })
    const zip = await JSZip.loadAsync(await readFile(workbook))
    const savedRuns = richRunsFromWorksheet(await packageXml(zip, 'xl/worksheets/sheet1.xml'))
    expect(savedRuns).toEqual([
      { font: 'Arial', text: 'Styled', bold: true },
      { font: 'Segoe UI Symbol', text: 'Ω', bold: true },
      { font: 'Cambria Math', text: '∑', bold: true },
    ])
    await attachArtifact(
      'sheets-math-symbol-summary.json',
      JSON.stringify({ value: 'StyledΩ∑', runs: savedRuns }, null, 2),
      'application/json',
    )
  } finally {
    await close(secondReopen.launched, 'sheets-math-rich-symbol-second-reopen')
  }
})
