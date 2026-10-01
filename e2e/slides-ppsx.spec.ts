import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('PPSX opens in Slides, saves edits, and converts to PPTX with the correct package type', async () => {
  const dir =
    process.env.E2E_PPSX_OUTPUT_DIR ?? (await mkdtemp(join(tmpdir(), 'genoffice-ppsx-e2e-')))
  const source = process.env.E2E_PPSX_SOURCE
  const input = join(dir, 'input.PPSX')
  if (source) await writeFile(input, await readFile(source))
  else {
    const zip = await JSZip.loadAsync(
      await readFile('packages/pptx-engine/tests/fixtures/01_standard_business.pptx'),
    )
    const types = await zip.file('[Content_Types].xml')!.async('string')
    zip.file(
      '[Content_Types].xml',
      types.replace('presentationml.presentation.main+xml', 'presentationml.slideshow.main+xml'),
    )
    await writeFile(input, await zip.generateAsync({ type: 'nodebuffer' }))
  }
  const launched = await launchShell({ onboardingSeen: true, openFile: input, videoDir: 'ppsx' })
  try {
    const s = await waitForPageWithUrl(launched.app, '://slides/', 30000)
    const editor = s.locator('.notes-editor')
    await expect(editor).toBeVisible()
    await editor.click()
    await s.keyboard.press('ControlOrMeta+End')
    await s.keyboard.type(' PPSX roundtrip check')
    await s.locator('.stage-wrap').click({ position: { x: 6, y: 6 } })
    await expect
      .poll(() =>
        s.evaluate(async () => JSON.stringify(await (window as any).slidesApi.getNotesRich(0))),
      )
      .toContain('PPSX roundtrip check')
    expect(await s.evaluate(() => (window as any).slidesApi.save())).toMatchObject({ ok: true })
    const show = await JSZip.loadAsync(await readFile(input))
    expect(await show.file('[Content_Types].xml')!.async('string')).toContain(
      'presentationml.slideshow.main+xml',
    )
    const output = join(dir, 'converted.pptx')
    await launched.app.evaluate(({ dialog }, file) => {
      dialog.showSaveDialog = (async (_parent: unknown, options: any) => {
        if (!options.filters.some((f: any) => f.extensions.includes('ppsx')))
          throw Error('Missing PPSX save filter')
        return { canceled: false, filePath: file }
      }) as never
    }, output)
    expect(await s.evaluate(() => (window as any).slidesApi.saveAs('input.PPSX'))).toMatchObject({
      ok: true,
    })
    const deck = await JSZip.loadAsync(await readFile(output))
    expect(await deck.file('[Content_Types].xml')!.async('string')).toContain(
      'presentationml.presentation.main+xml',
    )
    for (const name of Object.keys(show.files)) {
      if (name !== '[Content_Types].xml' && !show.files[name].dir) {
        expect(await deck.file(name)!.async('nodebuffer')).toEqual(
          await show.file(name)!.async('nodebuffer'),
        )
      }
    }
    const reopened = await launchShell({
      onboardingSeen: true,
      openFile: input,
      videoDir: 'ppsx-reopen',
    })
    try {
      const next = await waitForPageWithUrl(reopened.app, '://slides/', 30000)
      await expect(next.locator('.notes-editor')).toContainText('PPSX roundtrip check')
    } finally {
      await closeAndSaveVideo(reopened, 'ppsx-reopen')
    }
  } finally {
    await closeAndSaveVideo(launched, 'ppsx')
  }
})
