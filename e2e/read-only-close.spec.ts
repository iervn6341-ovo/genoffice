import { expect, test } from '@playwright/test'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'
import { fixture, kinds, openFreshEditor } from './editor-lifecycle-helpers'

for (const kind of kinds) {
  test(`${kind}: reading, selecting and scrolling does not ask to save`, async () => {
    test.setTimeout(90_000)
    const root = await mkdtemp(join(tmpdir(), 'genoffice-read-only-'))
    const path = await fixture(root, kind)
    const original = await readFile(path)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: `read-only-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      await launched.app.evaluate(({ dialog, ipcMain }) => {
        ;(globalThis as any).__readingPrompts = []
        ;(globalThis as any).__readingCloseStates = []
        ipcMain.on('docs:close-check-result', (_event, state) => {
          ;(globalThis as any).__readingCloseStates.push(state)
        })
        dialog.showMessageBox = (async (...args: unknown[]) => {
          ;(globalThis as any).__readingPrompts.push(args.at(-1))
          return { response: 2, checkboxChecked: false }
        }) as never
      })
      await launched.page.evaluate((path) => (window as any).aiOffice.openPath(path), path)
      const page = await waitForPageWithUrl(launched.app, `://${kind}/`, 30_000)
      if (kind === 'docs' || kind === 'markdown') {
        const editor = page
          .locator(kind === 'docs' ? '.editor-scroll [contenteditable="true"]' : '.doc-editor')
          .first()
        await expect(editor).toContainText('Preview without editing')
        if (kind === 'docs') {
          await expect(page.locator('.editor-scroll')).toHaveAttribute('aria-busy', 'false')
          await page.evaluate(() => {
            const editor = (window as any).__aidocs.editor
            ;(window as any).__readingTransactions = []
            editor.on('update', ({ transaction, appendedTransactions }: any) => {
              ;(window as any).__readingTransactions.push(
                [transaction, ...appendedTransactions].map((tr: any) => ({
                  steps: tr.steps.map((step: any) => step.toJSON()),
                  meta: tr.meta,
                })),
              )
            })
          })
        }
        await editor.locator('p').first().click()
        await page.keyboard.press('ArrowRight')
        await page.keyboard.press('Shift+ArrowRight')
        await page.locator('.editor-scroll').evaluate((el) => {
          el.scrollTop = 1200
        })
        await page.locator('.editor-scroll').evaluate((el) => {
          el.scrollTop = 0
        })
      } else if (kind === 'pdf') {
        await expect(page.locator('.pdf-page canvas').first()).toBeVisible()
        await page.locator('.textLayer span').first().click()
        await page.locator('.pdf-scroll').evaluate((el) => {
          el.scrollTop = 1000
        })
        await page.locator('.pdf-scroll').evaluate((el) => {
          el.scrollTop = 0
        })
      } else if (kind === 'html') {
        const heading = page.frameLocator('.preview-frame').locator('h1')
        await expect(heading).toHaveText('Preview without editing')
        await heading.click()
        await heading.dblclick()
        await page.locator('.ribbon').click({ position: { x: 10, y: 10 } })
        await page.locator('.rb-view', { hasText: /Source/ }).click()
        await page.locator('.source-editor .cm-content').click()
        await page.keyboard.press('ArrowDown')
        await page.keyboard.press('Shift+ArrowRight')
      } else if (kind === 'sheets') {
        await page.waitForFunction(
          () =>
            Boolean(
              (window as any).__genofficeDebug?.univerAPI
                ?.getActiveWorkbook()
                ?.getId()
                .startsWith('file-'),
            ),
          null,
          { timeout: 30_000 },
        )
        await page.evaluate(() => {
          ;(window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange('B2')
            .activate()
        })
        await page.keyboard.press('ArrowDown')
        await page.keyboard.press('ArrowRight')
        await page.locator('canvas[id^="univer-sheet-main-canvas_"]:visible').first().hover()
        await page.mouse.wheel(0, 600)
        await page.mouse.wheel(0, -600)
        await page.keyboard.press('Escape')
        await page
          .locator('canvas[id^="univer-sheet-main-canvas_"]:visible')
          .first()
          .click({ position: { x: 160, y: 100 } })
        await page.keyboard.press('Escape')
        await page.evaluate(async () => {
          const api = (window as any).__genofficeDebug.univerAPI
          const workbook = api.getActiveWorkbook()
          await api.executeCommand('sheet.command.set-zoom-ratio', {
            unitId: workbook.getId(),
            subUnitId: workbook.getActiveSheet().getSheetId(),
            zoomRatio: 0.85,
          })
        })
        await expect
          .poll(() =>
            page.evaluate(() =>
              (window as any).__genofficeDebug.univerAPI
                .getActiveWorkbook()
                .getActiveSheet()
                .getZoom(),
            ),
          )
          .toBe(0.85)
      } else {
        await expect(page.locator('.stage-rel')).toBeVisible()
        const position = await page.evaluate(async () => {
          const slides = await (window as any).slidesApi.getRenderSlides()
          const slide = slides[0]
          const node = slide.nodes.find((node: any) => node.text && node.box)
          const rect = document.querySelector('.stage-rel')!.getBoundingClientRect()
          return {
            x: rect.x + ((node.box.x + node.box.w / 2) * rect.width) / slide.widthPx,
            y: rect.y + ((node.box.y + node.box.h / 2) * rect.height) / slide.heightPx,
          }
        })
        await page.mouse.click(position.x, position.y)
        await page.mouse.dblclick(position.x, position.y)
        if (await page.locator('.slide-text-editor').isVisible())
          await page.keyboard.press('Escape')
        await page.locator('.thumb').first().click()
      }
      // Let delayed rendering, layout and blur commits settle before closing.
      await page.waitForTimeout(1000)
      const diagnostics =
        kind === 'docs'
          ? await page.evaluate(() => (window as any).__readingTransactions)
          : undefined
      await launched.page.locator('.tab-item.active .tab-close').click()
      const prompts = await launched.app.evaluate(() => (globalThis as any).__readingPrompts)
      expect(prompts, JSON.stringify({ kind, prompts, diagnostics })).toEqual([])
      await expect(launched.page.locator('.tab-close')).toHaveCount(0)
      expect(await readFile(path)).toEqual(original)

      // Positive control: authoring must still be protected by the close prompt.
      const edited = await openFreshEditor(launched, path, kind)
      if (kind === 'docs' || kind === 'markdown') {
        const editor = edited
          .locator(kind === 'docs' ? '.editor-scroll [contenteditable="true"]' : '.doc-editor')
          .first()
        await expect(editor).toContainText('Preview without editing')
        if (kind === 'docs') {
          await edited.waitForFunction(() => (window as any).__aidocs?.editor.isEditable)
          await edited.locator('.autosave-toggle input').uncheck({ force: true })
        }
        await editor.locator('p').first().click()
        await edited.keyboard.type(' Actual edit')
        await expect(editor).toContainText('Actual edit')
        if (kind === 'docs')
          await edited.waitForFunction(() =>
            (window as any).__aidocs.editor.state.doc.textContent.includes('Actual edit'),
          )
      } else if (kind === 'pdf') {
        await edited.getByRole('button', { name: 'Pages', exact: true }).click()
        await edited.getByRole('button', { name: 'Rotate right', exact: true }).click()
      } else if (kind === 'html') {
        await edited.locator('.rb-view', { hasText: /Source/ }).click()
        await edited.locator('.source-editor .cm-content').click()
        await edited.keyboard.type('<!--Actual edit-->')
      } else if (kind === 'sheets') {
        await edited.waitForFunction(() =>
          Boolean(
            (window as any).__genofficeDebug?.univerAPI
              ?.getActiveWorkbook()
              ?.getId()
              .startsWith('file-'),
          ),
        )
        // The Facade exists before the final import pass releases journal suppression.
        await expect(edited.locator('.status-msg')).toContainText('Workbook fully loaded')
        await edited.evaluate(() => {
          ;(window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange('B2')
            .setValue('Actual edit')
        })
      } else {
        await edited.locator('.ribbon-tab', { hasText: 'Insert' }).click()
        await edited.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
        await expect(edited.locator('.slide-text-editor')).toBeVisible()
        await edited.keyboard.type('Actual edit')
        await edited.keyboard.press('Escape')
      }
      await launched.page.locator('.tab-item.active .tab-close').click()
      await expect
        .poll(async () => {
          const state = await launched.app.evaluate(() => ({
            prompts: (globalThis as any).__readingPrompts.length,
            states: (globalThis as any).__readingCloseStates,
          }))
          return state.prompts
        })
        .toBe(1)
      await expect(launched.page.locator('.tab-close')).toHaveCount(1)
      expect(await readFile(path)).toEqual(original)
    } finally {
      await closeAndSaveVideo(launched, `read-only-${kind}`)
      await rm(root, { recursive: true, force: true })
    }
  })
}

test('Docs: navigating below a final table stays clean; typing still prompts before close', async () => {
  test.setTimeout(90_000)
  const root = await mkdtemp(join(tmpdir(), 'genoffice-table-reading-'))
  const path = join(root, 'table.docx')
  await writeFile(
    path,
    await buildDocx({
      bodyXml:
        '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Read table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    }),
  )
  const original = await readFile(path)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'read-only-table' })
  try {
    await launched.app.evaluate(({ dialog, ipcMain }) => {
      ;(globalThis as any).__readingPrompts = []
      ;(globalThis as any).__readingCloseStates = []
      ipcMain.on('docs:close-check-result', (_event, state) => {
        ;(globalThis as any).__readingCloseStates.push(state)
      })
      dialog.showMessageBox = (async (...args: unknown[]) => {
        ;(globalThis as any).__readingPrompts.push(args.at(-1))
        return { response: 2, checkboxChecked: false }
      }) as never
    })
    const open = async () => {
      const page = await openFreshEditor(launched, path, 'docs')
      await expect(page.locator('.editor-scroll')).toHaveAttribute('aria-busy', 'false')
      await page.locator('.autosave-toggle input').uncheck({ force: true })
      const table = page.locator('.editor-scroll table').first()
      await expect(table).toContainText('Read table cell')
      const box = (await table.boundingBox())!
      await page.mouse.click(box.x + 20, box.y + box.height + 12)
      await page.waitForFunction(
        () => (window as any).__aidocs.editor.state.doc.lastChild?.type.name === 'docParagraph',
      )
      return page
    }
    await open()
    await launched.page.locator('.tab-item.active .tab-close').click()
    await expect(launched.page.locator('.tab-close')).toHaveCount(0)
    expect(await launched.app.evaluate(() => (globalThis as any).__readingPrompts)).toEqual([])
    const page = await open()
    await page.keyboard.type('Actual edit below table')
    await expect(page.locator('.editor-scroll')).toContainText('Actual edit below table')
    await page.waitForFunction(() =>
      (window as any).__aidocs.editor.state.doc.textContent.includes('Actual edit below table'),
    )
    await launched.page.locator('.tab-item.active .tab-close').click()
    await expect
      .poll(() => launched.app.evaluate(() => (globalThis as any).__readingPrompts.length))
      .toBe(1)
    await expect(launched.page.locator('.tab-close')).toHaveCount(1)
    expect(await readFile(path)).toEqual(original)
  } finally {
    await closeAndSaveVideo(launched, 'read-only-table')
    await rm(root, { recursive: true, force: true })
  }
})
