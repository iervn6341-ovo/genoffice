import { expect, test, type Page } from '@playwright/test'
import { existsSync } from 'node:fs'
import { copyFile, mkdir, readFile, unlink, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

test.skip(
  !process.env.GENOFFICE_BIDIRECTIONAL,
  'Opt-in native Office bidirectional compatibility audit',
)
const root = resolve('docs/qa/office-roundtrip-2026-11-04')
const imagePath = resolve('apps/docs/src/renderer/assets/app-icon.png')

test.beforeAll(async () => {
  for (const folder of ['office', 'genoffice', 'roundtrip', 'pdf', 'screenshots', 'checks'])
    await mkdir(join(root, folder), { recursive: true })
})
const html = `<p style="font-family:Arial;font-size:20pt"><b>GenOffice Native Compatibility</b></p><p>原生 GenOffice 雙向格式測試</p><h1>Chapter 1 Text and tables</h1><p style="font-family:Arial;font-size:12pt">中文與 English 混合文字。α β Ω ∑ x² H₂O。<b>Bold text</b>, <i>italic text</i>, <u>underlined text</u>, <span style="color:#C00000">red text</span>.</p><ul><li>First bullet</li><li>Second bullet<ul><li>Nested bullet</li></ul></li></ul><table style="border-collapse:collapse;width:560px" border="1"><tr style="background-color:#D9E2F3"><th colspan="2">Merged group heading</th><th>Amount</th><th>Notes</th></tr><tr><td rowspan="2">Vertical merge</td><td>Alpha</td><td>1,250.50</td><td>Long wrapped text over multiple lines</td></tr><tr><td>Beta</td><td>2,000.00</td><td>中文內容與換行測試</td></tr><tr><td>Group B</td><td>Gamma</td><td>325.25</td><td>Final data row</td></tr><tr style="background-color:#FFF2CC"><td colspan="2">Total</td><td>3,575.75</td><td>Native merged table</td></tr></table><h2>Section 1.1 Research overview</h2><p>This document tests paragraph spacing, mixed language text, lists, merged tables and fields.</p><h1 style="page-break-before:always">Chapter 2 Fields and equations</h1><h2>Section 2.1 Results</h2><p>Editable note: GENOFFICE_EDIT_MARKER.</p><p>EQUATION_ANCHOR</p><p>End of specimen.</p>`

async function dialogTarget(l: LaunchedApp, target: string) {
  await l.app.evaluate(({ dialog }, path) => {
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
  }, target)
}
async function ribbon(page: Page, name: string) {
  await page
    .locator('.ribbon-tabs button')
    .filter({ hasText: new RegExp(`^${name}$`) })
    .first()
    .click()
}
async function command(page: Page, tip: string) {
  const button = page.locator(`[data-tip="${tip}"]`).first()
  if (!(await button.isVisible())) {
    const group = button.locator(
      'xpath=ancestor::*[contains(concat(" ",normalize-space(@class)," ")," rb-fold-group ")][1]',
    )
    await group.locator(':scope > .rb-fold-btn').click()
  }
  await button.click()
}
async function zip(path: string) {
  return JSZip.loadAsync(await readFile(path))
}

test('author Docs objects in a new GenOffice document', async () => {
  test.setTimeout(180_000)
  await mkdir(join(root, 'genoffice'), { recursive: true })
  const path = join(root, 'genoffice/genoffice-docs.docx')
  const l = await launchShell({
    onboardingSeen: true,
    settings: { defaultSaveDir: join(root, 'genoffice') },
    videoDir: 'office-roundtrip-docs-author',
  })
  try {
    await l.page.locator('.quick-card', { hasText: 'AI Docs' }).click()
    const page = await waitForPageWithUrl(l.app, '://docs/')
    await page.waitForFunction(() => Boolean((window as any).__aidocs?.editor?.isEditable))
    await expect(page.locator('.doc-page').first()).toBeVisible({ timeout: 30_000 })
    await page.evaluate((markup) => {
      ;(window as any).__aidocs.editor.commands.setContent(markup)
    }, html)
    await expect(page.locator('table.doc-table')).toHaveCount(1)
    await page.evaluate(() => {
      ;(window as any).__aidocs.editor.commands.setTextSelection(1)
    })
    await ribbon(page, 'References')
    await command(page, 'Custom table of contents')
    const toc = page.getByRole('dialog', { name: 'Custom table of contents' })
    await toc.getByLabel('Show levels').fill('2')
    await toc.getByRole('button', { name: 'Insert', exact: true }).click()
    await page.evaluate(() => {
      const e = (window as any).__aidocs.editor
      e.state.doc.forEach((n: any, pos: number) => {
        if (n.textContent === 'EQUATION_ANCHOR') e.commands.setTextSelection(pos + n.nodeSize - 1)
      })
    })
    await ribbon(page, 'Insert')
    await command(page, 'Insert an equation')
    await page.locator('.equation-gallery-custom').click()
    const math = page.locator('.equation-modal')
    await math.locator('textarea.go-math-input').fill('x=\\frac{-b+\\sqrt{b^{2}-4ac}}{2a}')
    await math.getByRole('button', { name: 'Insert', exact: true }).click()
    await dialogTarget(l, path)
    await page.evaluate(() =>
      window.dispatchEvent(
        new KeyboardEvent('keydown', { key: 's', metaKey: true, shiftKey: true, bubbles: true }),
      ),
    )
    await expect.poll(() => existsSync(path), { timeout: 60_000 }).toBe(true)
    await page.waitForTimeout(1200)
    const pkg = await zip(path)
    const xml = await pkg.file('word/document.xml')!.async('string')
    expect(xml).toContain(' TOC ')
    expect(xml).toContain('<m:oMath')
    expect(xml).toContain('w:gridSpan')
    expect(xml).toContain('w:vMerge')
    await page.screenshot({ path: join(root, 'screenshots/genoffice-docs-authored.png') })
    await exportGenPdf(l, page, 'docs', 'genoffice-docs')
    await writeFile(
      join(root, 'checks/genoffice-docs-author.json'),
      JSON.stringify(
        {
          provenance:
            'Blank document; native editor HTML insertion, ribbon TOC and equation creation, application Save and PDF export.',
          objects: [
            'rich text',
            'headings',
            'multilevel lists',
            'horizontal and vertical merged table',
            'TOC field',
            'OMML equation',
          ],
          xmlSize: xml.length,
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(l, 'office-roundtrip-docs-author')
  }
})

test('author Sheets objects in a new GenOffice workbook', async () => {
  test.setTimeout(180_000)
  const path = join(root, 'genoffice/genoffice-sheets.xlsx')
  const l = await launchShell({
    onboardingSeen: true,
    videoDir: 'office-roundtrip-sheets-author',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    await l.page.locator('.quick-card', { hasText: 'AI Sheets' }).click()
    const page = await waitForPageWithUrl(l.app, '://sheets/')
    await page.waitForFunction(() =>
      Boolean((window as any).__genofficeDebug?.univerAPI?.getActiveWorkbook()),
    )
    await page.evaluate(() => {
      const s = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      s.getRange('A1').setValue('GenOffice Native Workbook')
      s.getRange('A1:F1')
        .merge()
        .setFontWeight('bold')
        .setFontSize(18)
        .setHorizontalAlignment('center')
      s.getRange('A3:D8').setValues([
        ['Item', 'Quantity', 'Price', 'Amount'],
        ['Alpha', 2, 1250.5, '=B4*C4'],
        ['Beta', 3, 2000, '=B5*C5'],
        ['Gamma', 5, 325.25, '=B6*C6'],
        ['Delta', 4, 400, '=B7*C7'],
        ['Total', '=SUM(B4:B7)', null, '=SUM(D4:D7)'],
      ])
      s.getRange('A3:D3').setFontWeight('bold').setBackgroundColor('#D9E2F3')
      s.getRange('C4:D8').setNumberFormat('#,##0.00')
      s.getRange('A8:D8').setBackgroundColor('#FFF2CC').setFontWeight('bold')
      s.getRange('F3').setValue('Mixed / wrapped')
      s.getRange('F4').setValue('中文 English Ω ∑ long wrapped text')
      s.getRange('F4:F5').merge().setWrap(true)
      s.getRange('F7').setValue(0.1234).setNumberFormat('0.00%')
      s.getRange('F8').setValue('GENOFFICE_EDIT_MARKER')
    })
    await dialogTarget(l, path)
    await page.locator('button[aria-label^="Save ("]').first().click()
    await expect.poll(() => existsSync(path), { timeout: 60_000 }).toBe(true)
    await page.waitForTimeout(1500)
    const pkg = await zip(path)
    const xml = await pkg.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(xml).toContain('mergeCell')
    expect(xml).toContain('<f>B4*C4</f>')
    await page.screenshot({ path: join(root, 'screenshots/genoffice-sheets-authored.png') })
    await writeFile(
      join(root, 'checks/genoffice-sheets-author.json'),
      JSON.stringify(
        {
          provenance:
            'Blank workbook; native Univer edit/format/merge commands and application toolbar Save.',
          objects: [
            'merged title',
            'vertical merge',
            'wrapped mixed text',
            'number formats',
            'percent',
            'formulas',
            'cell fills',
            'bold',
          ],
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(l, 'office-roundtrip-sheets-author')
  }
})

test('author Slides objects in a new GenOffice presentation', async () => {
  test.setTimeout(240_000)
  const path = join(root, 'genoffice/genoffice-slides.pptx')
  let l = await launchShell({ onboardingSeen: true, videoDir: 'office-roundtrip-slides-author' })
  try {
    await l.page.locator('.quick-card', { hasText: 'AI Slides' }).click()
    let page = await waitForPageWithUrl(l.app, '://slides/')
    await page.waitForSelector('.slide-list .thumb')
    const base64 = (await readFile(imagePath)).toString('base64')
    await page.evaluate(async (image) => {
      const api = (window as any).slidesApi
      const width = 1280
      const p = (text: string, fontSize = 20) => [
        { runs: [{ text, fontFamily: 'Arial', fontSize }] },
      ]
      await api.addElement({
        slideIndex: 0,
        kind: 'textbox',
        xPx: 50,
        yPx: 25,
        wPx: 1160,
        hPx: 70,
        fitWidthPx: width,
        paragraphs: p('GenOffice native objects', 32),
      })
      await api.addElement({
        slideIndex: 0,
        kind: 'textbox',
        xPx: 50,
        yPx: 105,
        wPx: 1100,
        hPx: 85,
        fitWidthPx: width,
        paragraphs: [
          {
            runs: [
              { text: '中文 English Ω ∑ ', fontFamily: 'Arial', fontSize: 20 },
              { text: 'Bold ', bold: true, fontSize: 20 },
              { text: 'Italic ', italic: true, fontSize: 20 },
              { text: 'Underlined', underline: true, fontSize: 20 },
            ],
          },
        ],
      })
      const table = await api.addTable({
        slideIndex: 0,
        rows: 4,
        cols: 4,
        xPx: 50,
        yPx: 220,
        wPx: 850,
        hPx: 280,
        fitWidthPx: width,
      })
      for (let row = 0; row < 4; row++)
        for (let col = 0; col < 4; col++)
          await api.editTableCell({
            slideIndex: 0,
            sourceId: table.sourceId,
            row,
            col,
            paragraphs: p(
              row === 0
                ? ['Group', 'Quarter', 'Amount', 'Notes'][col]!
                : [`Region ${row}`, `Q${row}`, `${row * 1250}.50`, 'Wrapped table text'][col]!,
              16,
            ),
          })
      await api.tableMerge({
        slideIndex: 0,
        sourceId: table.sourceId,
        kind: 'merge-right',
        row: 0,
        col: 0,
      })
      await api.addImageBytes({
        slideIndex: 0,
        base64: image,
        ext: 'png',
        xPx: 970,
        yPx: 240,
        wPx: 180,
        hPx: 180,
        fitWidthPx: width,
        name: 'Image specimen',
      })
      const shape = await api.addElement({
        slideIndex: 0,
        kind: 'roundRect',
        xPx: 960,
        yPx: 470,
        wPx: 210,
        hPx: 110,
        fitWidthPx: width,
        fillColor: '#D9E2F3',
        stroke: { color: '#4472C4', widthPt: 2 },
        text: 'Rotated shape',
      })
      await api.editTransform({
        slideIndex: 0,
        sourceId: shape.sourceId,
        xPx: 960,
        yPx: 470,
        wPx: 210,
        hPx: 110,
        rotationDeg: 15,
        fitWidthPx: width,
      })
      await api.addBlankSlide({ sourceIndex: 0, fitWidthPx: width })
      await api.addElement({
        slideIndex: 1,
        kind: 'textbox',
        xPx: 50,
        yPx: 25,
        wPx: 1160,
        hPx: 70,
        fitWidthPx: width,
        paragraphs: p('Chart and editable objects', 30),
      })
      await api.addChart({
        slideIndex: 1,
        kind: 'bar',
        title: 'Native chart',
        categories: ['Alpha', 'Beta', 'Gamma'],
        series: [
          { name: '2025', values: [12, 18, 9] },
          { name: '2026', values: [16, 21, 15] },
        ],
        xPx: 50,
        yPx: 140,
        wPx: 800,
        hPx: 440,
        fitWidthPx: width,
      })
      await api.addElement({
        slideIndex: 1,
        kind: 'textbox',
        xPx: 890,
        yPx: 180,
        wPx: 320,
        hPx: 160,
        fitWidthPx: width,
        text: 'GENOFFICE_EDIT_MARKER\nEditable native chart and text',
      })
      await api.applyHeaderFooter({
        footer: 'Compatibility specimen',
        slideNum: true,
        date: '2026/11/04',
        firstSlideNum: 1,
        dontShowOnTitleSlide: false,
        fitWidthPx: width,
      })
    }, base64)
    await dialogTarget(l, path)
    const saved = await page.evaluate(() =>
      (window as any).slidesApi.saveAs('genoffice-slides.pptx'),
    )
    expect(saved.ok).toBe(true)
    await closeAndSaveVideo(l, 'office-roundtrip-slides-authored')
    l = await launchShell({
      onboardingSeen: true,
      openFile: path,
      videoDir: 'office-roundtrip-slides-preview',
    })
    page = await waitForPageWithUrl(l.app, '://slides/')
    await expect(page.locator('.slide-list .thumb')).toHaveCount(2)
    await page.screenshot({ path: join(root, 'screenshots/genoffice-slides-authored-page1.png') })
    await page.locator('.slide-list .thumb').nth(1).click()
    await page.screenshot({ path: join(root, 'screenshots/genoffice-slides-authored-page2.png') })
    await writeFile(
      join(root, 'checks/genoffice-slides-author.json'),
      JSON.stringify(
        {
          provenance:
            'Blank presentation; native slidesApi object insertion/format/merge/chart/footer operations and application Save As; fresh application reopen.',
          objects: [
            'rich text',
            'merged native table',
            'picture',
            'rotated shape',
            'native chart',
            'footer/date/slide number',
          ],
        },
        null,
        2,
      ),
    )
  } finally {
    await closeAndSaveVideo(l, 'office-roundtrip-slides-author')
  }
})

async function exportGenPdf(l: LaunchedApp, page: Page, kind: string, name: string) {
  const path = join(root, `pdf/${name}-genoffice.pdf`)
  await unlink(path).catch(() => undefined)
  if (kind === 'docs') {
    await dialogTarget(l, path)
    await page.evaluate(() => (window as any).__aidocs.exportPdfTo())
  } else {
    await dialogTarget(l, path)
    await l.app.evaluate(({ webContents }, k) => {
      webContents
        .getAllWebContents()
        .find((w) => w.getURL().includes(`://${k}/`))
        ?.send(k === 'slides' ? 'slides:menu' : 'menu:action', 'export-pdf')
    }, kind)
  }
  await expect
    .poll(
      async () => {
        try {
          return (await readFile(path)).subarray(0, 5).toString()
        } catch {
          return ''
        }
      },
      { timeout: 60_000 },
    )
    .toBe('%PDF-')
}

for (const kind of ['docs', 'sheets', 'slides'] as const)
  test(`import Office ${kind}, edit/history/save and reopen`, async () => {
    test.setTimeout(240_000)
    const ext = { docs: 'docx', sheets: 'xlsx', slides: 'pptx' }[kind]
    const path = join(root, `roundtrip/office-${kind}-genoffice.${ext}`)
    await copyFile(join(root, `office/office-${kind}.${ext}`), path)
    let l = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: join(root, 'roundtrip'), autoSave: false },
      openFile: path,
      videoDir: `office-import-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      let page = await waitForPageWithUrl(l.app, `://${kind}/`)
      const ready = async () => {
        if (kind === 'docs') {
          await page.waitForFunction(() =>
            (window as any).__aidocs?.editor?.state.doc.textContent.includes('OFFICE_EDIT_MARKER'),
          )
          await expect(page.locator('.editor-scroll')).toHaveAttribute('aria-busy', 'false')
        } else if (kind === 'sheets')
          await page.waitForFunction(
            () =>
              (window as any).__genofficeDebug?.univerAPI?.getActiveWorkbook()?.getSheets()
                .length === 2,
          )
        else await expect(page.locator('.slide-list .thumb')).toHaveCount(3)
      }
      await ready()
      await page.screenshot({ path: join(root, `screenshots/office-${kind}-in-genoffice.png`) })
      if (kind === 'slides')
        for (let i = 0; i < 3; i++) {
          await page.locator('.slide-list .thumb').nth(i).click()
          await page.locator('.stage-rel').screenshot({
            path: join(root, `screenshots/office-slides-in-genoffice-page${i + 1}.png`),
          })
        }
      const before = await page.evaluate((k) => {
        if (k === 'docs') return (window as any).__aidocs.editor.getJSON()
        if (k === 'slides') return (window as any).slidesApi.getRenderSlides()
        const w = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook()
        return w.getSheets().map((s: any) => ({
          name: s.getSheetName(),
          data: s.getRange('A1:F8').getValues(),
          formulas: s.getRange('A1:F8').getFormulas(),
        }))
      }, kind)
      await writeFile(
        join(root, `checks/office-${kind}-import-model.json`),
        JSON.stringify(before, null, 2),
      )
      await exportGenPdf(l, page, kind, `office-${kind}`)
      if (kind === 'docs') {
        await page.evaluate(() => {
          const e = (window as any).__aidocs.editor
          let from = -1
          e.state.doc.descendants((n: any, p: number) => {
            if (n.isText && n.text.includes('OFFICE_EDIT_MARKER'))
              from = p + n.text.indexOf('OFFICE_EDIT_MARKER')
          })
          if (from < 0) throw Error('Marker missing')
          e.chain()
            .setTextSelection({ from, to: from + 'OFFICE_EDIT_MARKER'.length })
            .insertContent('GENOFFICE_ROUNDTRIP_EDIT')
            .run()
        })
        await page.evaluate(() => (window as any).__aidocs.editor.commands.undo())
        expect(
          await page.evaluate(() => (window as any).__aidocs.editor.state.doc.textContent),
        ).toContain('OFFICE_EDIT_MARKER')
        await page.evaluate(() => (window as any).__aidocs.editor.commands.redo())
        expect(
          await page.evaluate(() => (window as any).__aidocs.editor.state.doc.textContent),
        ).toContain('GENOFFICE_ROUNDTRIP_EDIT')
        expect(await page.evaluate(() => (window as any).__aidocs.save())).toBe(true)
      } else if (kind === 'sheets') {
        await page.evaluate(() => {
          ;(window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getSheets()[0]
            .getRange('B4')
            .setValue(4)
        })
        const value = () =>
          page.evaluate(() =>
            (window as any).__genofficeDebug.univerAPI
              .getActiveWorkbook()
              .getSheets()[0]
              .getRange('B4')
              .getValue(),
          )
        await expect.poll(value).toBe(4)
        await page.locator('button[aria-label^="Undo"]').first().click()
        await expect.poll(value).toBe(2)
        await page.locator('button[aria-label^="Redo"]').first().click()
        await expect.poll(value).toBe(4)
        await page.locator('button[aria-label^="Save ("]').first().click()
        await page.waitForTimeout(2200)
      } else {
        const result = await page.evaluate(async () => {
          const a = (window as any).slidesApi
          const slides = await a.getRenderSlides()
          const n = slides[0].nodes.find((x: any) =>
            x.text?.lines
              ?.map((p: any) => p.runs.map((r: any) => r.text).join(''))
              .join(' ')
              .includes('Office Native'),
          )
          if (!n) throw Error('Native text missing')
          await a.editText({
            slideIndex: 0,
            sourceId: n.sourceId,
            paragraphs: [{ runs: [{ text: 'GENOFFICE_ROUNDTRIP_EDIT', fontSize: 44 }] }],
          })
          await a.undo()
          const undo = (await a.getRenderSlides())[0].nodes.some((x: any) =>
            x.text?.lines
              ?.map((p: any) => p.runs.map((r: any) => r.text).join(''))
              .join(' ')
              .includes('Office Native Compatibility'),
          )
          await a.redo()
          const redo = JSON.stringify((await a.getRenderSlides())[0]).includes(
            'GENOFFICE_ROUNDTRIP_EDIT',
          )
          return { undo, redo, saved: await a.save() }
        })
        expect(result.undo).toBe(true)
        expect(result.redo).toBe(true)
        expect(result.saved.ok).toBe(true)
      }
      await closeAndSaveVideo(l, `office-${kind}-saved`)
      l = await launchShell({
        onboardingSeen: true,
        openFile: path,
        videoDir: `office-${kind}-reopened`,
        env: { GENOFFICE_DEBUG_HOOKS: '1' },
      })
      page = await waitForPageWithUrl(l.app, `://${kind}/`)
      if (kind === 'docs')
        await page.waitForFunction(() =>
          (window as any).__aidocs?.editor?.state.doc.textContent.includes(
            'GENOFFICE_ROUNDTRIP_EDIT',
          ),
        )
      else if (kind === 'sheets') {
        await page.waitForFunction(
          () =>
            (window as any).__genofficeDebug?.univerAPI?.getActiveWorkbook()?.getSheets().length ===
            2,
        )
        const result = await page.evaluate(() => {
          const ss = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getSheets()
          return {
            quantity: ss[0].getRange('B4').getValue(),
            amount: ss[0].getRange('D4').getValue(),
            total: ss[0].getRange('D8').getValue(),
            cross: ss[1].getRange('B2:B4').getValues(),
          }
        })
        await writeFile(
          join(root, 'checks/office-sheets-reopen-values.json'),
          JSON.stringify(result, null, 2),
        )
        expect(result.quantity).toBe(4)
        const z = await zip(path)
        const sheet1 = await z.file('xl/worksheets/sheet1.xml')!.async('string')
        const sheet2 = await z.file('xl/worksheets/sheet2.xml')!.async('string')
        expect(sheet1).toMatch(/r="D4"[^>]*>[\s\S]*?<f>B4\*C4<\/f><v>5002<\/v>/)
        expect(sheet2).toContain('<v>28456.5</v>')
        await page.evaluate(() => {
          ;(window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getSheets()[1].activate()
        })
        await page.waitForTimeout(1500)
        await page.screenshot({
          path: join(root, 'screenshots/office-sheets-genoffice-reopened-sheet2.png'),
        })
      } else {
        await expect(page.locator('.slide-list .thumb')).toHaveCount(3)
        expect(
          await page.evaluate(async () =>
            JSON.stringify(await (window as any).slidesApi.getRenderSlides()),
          ),
        ).toContain('GENOFFICE_ROUNDTRIP_EDIT')
      }
      await page.screenshot({
        path: join(root, `screenshots/office-${kind}-genoffice-reopened.png`),
      })
      await writeFile(
        join(root, `checks/office-${kind}-history-reopen.json`),
        JSON.stringify({ edit: true, undo: true, redo: true, save: true, reopen: true }, null, 2),
      )
    } finally {
      await closeAndSaveVideo(l, `office-import-${kind}`)
    }
  })

for (const kind of ['docs', 'sheets', 'slides'] as const)
  test(`preview GenOffice ${kind} source after save`, async () => {
    test.setTimeout(120_000)
    const ext = { docs: 'docx', sheets: 'xlsx', slides: 'pptx' }[kind]
    const l = await launchShell({
      onboardingSeen: true,
      openFile: join(root, `genoffice/genoffice-${kind}.${ext}`),
      videoDir: `gen-source-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      const page = await waitForPageWithUrl(l.app, `://${kind}/`)
      if (kind === 'docs')
        await page.waitForFunction(() =>
          (window as any).__aidocs?.editor?.state.doc.textContent.includes('GENOFFICE_EDIT_MARKER'),
        )
      else if (kind === 'sheets')
        await page.waitForFunction(
          () =>
            (window as any).__genofficeDebug?.univerAPI
              ?.getActiveWorkbook()
              ?.getActiveSheet()
              ?.getRange('A1')
              .getValue() === 'GenOffice Native Workbook',
        )
      else await expect(page.locator('.slide-list .thumb')).toHaveCount(2)
      await page.screenshot({
        path: join(root, `screenshots/genoffice-${kind}-source-reopened.png`),
      })
      await exportGenPdf(l, page, kind, `genoffice-${kind}`)
    } finally {
      await closeAndSaveVideo(l, `gen-source-${kind}`)
    }
  })

for (const kind of ['docs', 'sheets', 'slides'] as const)
  test(`reopen ${kind} after editing and saving in Microsoft Office`, async () => {
    test.setTimeout(120_000)
    const ext = { docs: 'docx', sheets: 'xlsx', slides: 'pptx' }[kind]
    const native = { docs: 'word', sheets: 'excel', slides: 'powerpoint' }[kind]
    const path = join(root, `roundtrip/genoffice-${kind}-${native}.${ext}`)
    const l = await launchShell({
      onboardingSeen: true,
      openFile: path,
      videoDir: `native-roundtrip-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      const page = await waitForPageWithUrl(l.app, `://${kind}/`)
      if (kind === 'docs') {
        await page.waitForFunction(() =>
          (window as any).__aidocs?.editor?.state.doc.textContent.includes('WORD_ROUNDTRIP_EDIT'),
        )
        const x = await (await zip(path)).file('word/document.xml')!.async('string')
        expect(x).toContain(' TOC ')
        expect(x).toContain('<m:oMath')
        expect(x).toContain('w:gridSpan')
        expect(x).toContain('w:vMerge')
        await page.locator('table.doc-table').first().waitFor()
      } else if (kind === 'sheets') {
        await page.waitForFunction(
          () =>
            (window as any).__genofficeDebug?.univerAPI
              ?.getActiveWorkbook()
              ?.getActiveSheet()
              ?.getRange('F8')
              .getValue() === 'EXCEL_ROUNDTRIP_EDIT',
        )
        const x = await (await zip(path)).file('xl/worksheets/sheet1.xml')!.async('string')
        expect(x).toContain('<f>B4*C4</f>')
        expect(x).toContain('ref="F4:F5"')
      } else {
        await expect(page.locator('.slide-list .thumb')).toHaveCount(2)
        const model = await page.evaluate(() => (window as any).slidesApi.getRenderSlides())
        expect(JSON.stringify(model)).toContain('POWERPOINT_ROUNDTRIP_EDIT')
        const z = await zip(path)
        const x = await z.file('ppt/slides/slide1.xml')!.async('string')
        expect(x).toContain('<a:tbl>')
        expect(x).toContain('<p:pic>')
        expect(Object.keys(z.files).some((x) => /^ppt\/charts\/chart\d+\.xml$/.test(x))).toBe(true)
        await writeFile(
          join(root, 'checks/genoffice-slides-native-roundtrip-model.json'),
          JSON.stringify(model, null, 2),
        )
      }
      await page.screenshot({
        path: join(root, `screenshots/genoffice-${kind}-native-roundtrip-reopened.png`),
      })
      await writeFile(
        join(root, `checks/genoffice-${kind}-native-roundtrip.json`),
        JSON.stringify(
          {
            nativeEditMarkerPresent: true,
            nativeStructuresPresent: true,
            reopenedInGenOffice: true,
          },
          null,
          2,
        ),
      )
    } finally {
      await closeAndSaveVideo(l, `native-roundtrip-${kind}`)
    }
  })

test('Sheets: settled formula values and PDF after complete loading', async () => {
  test.setTimeout(180_000)
  for (const origin of ['office', 'genoffice']) {
    const old = join(root, `screenshots/rendered/compare-${origin}-sheets-1.png`)
    await copyFile(old, join(root, `checks/early-${origin}-sheets-comparison.png`)).catch(
      () => undefined,
    )
    const l = await launchShell({
      onboardingSeen: true,
      openFile: join(root, `${origin}/${origin}-sheets.xlsx`),
      videoDir: `settled-${origin}-sheets`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      const page = await waitForPageWithUrl(l.app, '://sheets/')
      const numeric = () =>
        page.evaluate(() => {
          const s = (window as any).__genofficeDebug?.univerAPI?.getActiveWorkbook()?.getSheets()[0]
          const v = s?.getRange('D4').getValue()
          return v === null || v === undefined ? null : Number(String(v).replace(/,/g, ''))
        })
      await expect.poll(numeric, { timeout: 60_000 }).toBe(2501)
      if (origin === 'office') {
        await page.getByText('Sheet2', { exact: true }).last().click()
        const cross = () =>
          page.evaluate(() => {
            const s = (window as any).__genofficeDebug.univerAPI
              .getActiveWorkbook()
              .getActiveSheet()
            return s
              .getRange('B2:B4')
              .getValues()
              .map((a: any[]) =>
                a[0] === null
                  ? null
                  : a[0] === true || a[0] === 1
                    ? 'TRUE'
                    : String(a[0]).replace(/,/g, '').toUpperCase(),
              )
          })
        await expect.poll(cross, { timeout: 30_000 }).toEqual(['11727.25', '23454.5', 'TRUE'])
        await page.screenshot({ path: join(root, 'screenshots/office-sheets-settled-sheet2.png') })
        await page.getByText('Sheet1', { exact: true }).last().click()
      }
      await page.screenshot({ path: join(root, `screenshots/${origin}-sheets-settled.png`) })
      await exportGenPdf(l, page, 'sheets', `${origin}-sheets`)
      await writeFile(
        join(root, `checks/${origin}-sheets-settled.json`),
        JSON.stringify(
          {
            D4: 2501,
            formulaValuesSettledBeforeExport: true,
            crossSheet: origin === 'office' ? ['11727.25', '23454.5', 'TRUE'] : null,
          },
          null,
          2,
        ),
      )
    } finally {
      await closeAndSaveVideo(l, `settled-${origin}-sheets`)
    }
  }
})
