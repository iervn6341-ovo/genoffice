import { expect, test } from '@playwright/test'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { fixture, kinds, openFreshEditor } from './editor-lifecycle-helpers'
import { closeAndSaveVideo, launchShell } from './helpers'

for (const kind of kinds) {
  test(`${kind}: Undo restores clean state, Redo restores save protection`, async () => {
    test.setTimeout(100_000)
    const root = await mkdtemp(join(tmpdir(), 'genoffice-undo-clean-'))
    const path = await fixture(root, kind)
    const original = await readFile(path)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: `undo-clean-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      await launched.app.evaluate(({ dialog }) => {
        ;(globalThis as any).__undoPrompts = 0
        dialog.showMessageBox = (async () => {
          ;(globalThis as any).__undoPrompts++
          return { response: 2, checkboxChecked: false }
        }) as never
      })
      const page = await openFreshEditor(launched, path, kind)
      let read: () => Promise<unknown>
      let undo: () => Promise<unknown> = () => page.keyboard.press('ControlOrMeta+z')
      let redo: () => Promise<unknown> = () => page.keyboard.press('ControlOrMeta+Shift+z')
      let before: unknown
      if (kind === 'docs' || kind === 'markdown') {
        const editor = page
          .locator(kind === 'docs' ? '.editor-scroll [contenteditable="true"]' : '.doc-editor')
          .first()
        await expect(editor).toContainText('Preview without editing')
        if (kind === 'docs') {
          await expect(page.locator('.editor-scroll')).toHaveAttribute('aria-busy', 'false')
          await page.locator('.autosave-toggle input').uncheck({ force: true })
          read = () => page.evaluate(() => (window as any).__aidocs.editor.state.doc.textContent)
        } else read = () => editor.textContent()
        before = await read()
        await editor.locator('p').first().click()
        await page.keyboard.type(' Undo verification')
      } else if (kind === 'html') {
        await page.locator('.rb-view', { hasText: /Source/ }).click()
        const content = page.locator('.source-editor .cm-content')
        await expect(content).toContainText('Preview without editing')
        read = () => content.textContent()
        before = await read()
        await content.click()
        await page.keyboard.type('<!--Undo verification-->')
      } else if (kind === 'sheets') {
        await page.waitForFunction(() =>
          Boolean(
            (window as any).__genofficeDebug?.univerAPI
              ?.getActiveWorkbook()
              ?.getId()
              .startsWith('file-'),
          ),
        )
        read = () =>
          page.evaluate(() =>
            (window as any).__genofficeDebug.univerAPI
              .getActiveWorkbook()
              .getActiveSheet()
              .getRange('B2')
              .getValue(),
          )
        undo = () => page.evaluate(() => (window as any).__genofficeDebug.univerAPI.undo())
        redo = () => page.evaluate(() => (window as any).__genofficeDebug.univerAPI.redo())
        before = await read()
        await page.evaluate(() =>
          (window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange('B2')
            .setValue('Undo verification'),
        )
      } else if (kind === 'pdf') {
        const canvas = page.locator('.pdf-page canvas').first()
        await expect(canvas).toBeVisible()
        read = () =>
          canvas.evaluate((el) => ({
            width: el.getBoundingClientRect().width,
            height: el.getBoundingClientRect().height,
          }))
        before = await read()
        await page.getByRole('button', { name: 'Pages', exact: true }).click()
        await page.getByRole('button', { name: 'Rotate right', exact: true }).click()
      } else {
        await expect(page.locator('.stage-rel')).toBeVisible()
        read = () =>
          page.evaluate(async () =>
            JSON.stringify(await (window as any).slidesApi.getRenderSlides()),
          )
        undo = () => page.evaluate(() => (window as any).slidesApi.undo())
        redo = () => page.evaluate(() => (window as any).slidesApi.redo())
        before = await read()
        await page.evaluate(async () => {
          const slide = (await (window as any).slidesApi.getRenderSlides())[0]
          const node = slide.nodes.find((node: any) => node.text && node.box)
          await (window as any).slidesApi.setElementFont({
            slideIndex: 0,
            sourceIds: [node.sourceId],
            bold: true,
          })
        })
      }
      await expect.poll(read).not.toEqual(before)
      await undo()
      await expect.poll(read).toEqual(before)
      await redo()
      await expect.poll(read).not.toEqual(before)
      await launched.page.locator('.tab-item.active .tab-close').click()
      await expect
        .poll(() => launched.app.evaluate(() => (globalThis as any).__undoPrompts))
        .toBe(1)
      await undo()
      await expect.poll(read).toEqual(before)
      await launched.page.locator('.tab-item.active .tab-close').click()
      await expect(launched.page.locator('.tab-close')).toHaveCount(0)
      expect(await launched.app.evaluate(() => (globalThis as any).__undoPrompts)).toBe(1)
      expect(await readFile(path)).toEqual(original)
    } finally {
      await closeAndSaveVideo(launched, `undo-clean-${kind}`)
      await rm(root, { recursive: true, force: true })
    }
  })
}
