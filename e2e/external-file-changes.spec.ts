import { expect, test, type Page } from '@playwright/test'
import { mkdtemp, readFile, writeFile, rename, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { PDFDocument } from 'pdf-lib'
import { fixture, kinds, openFreshEditor } from './editor-lifecycle-helpers'
import { launchShell, closeAndSaveVideo, type LaunchedApp } from './helpers'

type Kind = (typeof kinds)[number]
const LOCAL = 'Local unsaved edit'

async function externalVersion(original: Buffer, kind: Kind, version = 1): Promise<Buffer> {
  if (kind === 'docs' || kind === 'sheets' || kind === 'slides') {
    const zip = await JSZip.loadAsync(original)
    const part =
      kind === 'docs'
        ? 'word/document.xml'
        : kind === 'sheets'
          ? 'xl/worksheets/sheet1.xml'
          : 'ppt/slides/slide1.xml'
    const xml = await zip.file(part)!.async('string')
    const changed =
      kind === 'docs'
        ? xml.replace('Preview without editing', `External version ${version}`)
        : kind === 'slides'
          ? xml.replace(/(<a:t>)[^<]*/, `$1External version ${version}`)
          : xml.replace(
              /<c\b[^>]*\br="B1"[^>]*>[\s\S]*?<\/c>/,
              `<c r="B1" t="inlineStr"><is><t>External version ${version}</t></is></c>`,
            )
    expect(changed).not.toBe(xml)
    zip.file(part, changed)
    if (kind === 'docs' && zip.file('word/media/lazy.png'))
      zip.file('word/media/lazy.png', Buffer.alloc(16 * 1024 * 1024, 0x42 + version))
    return zip.generateAsync({ type: 'nodebuffer' })
  }
  if (kind === 'pdf') {
    const pdf = await PDFDocument.load(original)
    for (let i = 0; i < version; i++) pdf.addPage()
    return Buffer.from(await pdf.save())
  }
  return Buffer.from(
    original.toString().replace('Preview without editing', `External version ${version}`),
  )
}

async function replace(path: string, bytes: Buffer) {
  await writeFile(`${path}.external`, bytes)
  await rename(`${path}.external`, path)
}

async function dialogs(
  launched: LaunchedApp,
  action: 'reload' | 'keep' | 'copy',
  copyPath?: string,
) {
  await launched.app.evaluate(
    ({ dialog }, { action, copyPath }) => {
      ;(globalThis as any).__externalAction = action
      ;(globalThis as any).__externalNotices = []
      dialog.showMessageBox = (async (...args: any[]) => {
        const options = args[args.length - 1]
        if (options.title === 'File changed outside GenOffice') {
          ;(globalThis as any).__externalNotices.push({
            buttons: options.buttons,
            detail: options.detail,
          })
          const action = (globalThis as any).__externalAction
          const index =
            action === 'reload'
              ? options.buttons.indexOf('Reload')
              : action === 'copy'
                ? options.buttons.indexOf('Save a Copy…')
                : options.cancelId
          return { response: index, checkboxChecked: false }
        }
        return { response: 1, checkboxChecked: false }
      }) as never
      dialog.showSaveDialog = (async () => ({ canceled: !copyPath, filePath: copyPath })) as never
    },
    { action, copyPath },
  )
}

const notices = (launched: LaunchedApp) =>
  launched.app.evaluate(
    () => (globalThis as any).__externalNotices as { buttons: string[]; detail: string }[],
  )

async function ready(page: Page, kind: Kind) {
  if (kind === 'docs') {
    await page.waitForFunction(() =>
      Boolean((window as any).__aidocs?.editor.state.doc.textContent.trim()),
    )
    await expect(page.locator('.editor-scroll')).toHaveAttribute('aria-busy', 'false')
    await page.locator('.autosave-toggle input').uncheck({ force: true })
  } else if (kind === 'sheets') {
    await page.waitForFunction(
      () => (window as any).__genofficeDebug?.undoHistoryDiagnostic().fileSessionId,
    )
    await expect(page.locator('.status-msg')).toContainText('Workbook fully loaded')
  } else if (kind === 'slides') await expect(page.locator('.stage-rel')).toBeVisible()
  else if (kind === 'pdf') await expect(page.locator('.pdf-page canvas').first()).toBeVisible()
  else if (kind === 'markdown')
    await expect(page.locator('.doc-editor')).toContainText(
      /Preview without editing|External version|Local unsaved edit/,
    )
  else await expect(page.locator('.status-bar')).toBeVisible()
}

async function read(page: Page, kind: Kind): Promise<unknown> {
  if (kind === 'docs')
    return page.evaluate(() => (window as any).__aidocs.editor.state.doc.textContent)
  if (kind === 'markdown') return page.locator('.doc-editor').textContent()
  if (kind === 'html') {
    if (!(await page.locator('.source-editor .cm-content').isVisible()))
      await page.locator('.rb-view', { hasText: /Source/ }).click()
    return page.locator('.source-editor .cm-content').textContent()
  }
  if (kind === 'sheets')
    return page.evaluate(() =>
      (window as any).__genofficeDebug.univerAPI
        .getActiveWorkbook()
        .getActiveSheet()
        .getRange('B1')
        .getValue(),
    )
  if (kind === 'slides')
    return page.evaluate(async () => {
      const slide = (await (window as any).slidesApi.getRenderSlides())[0]
      const node = slide.nodes.find((n: any) => n.text && n.box)
      const runs = node.text.lines.flatMap((line: any) => line.runs)
      return { text: runs.map((r: any) => r.text).join(''), bold: runs[0].bold }
    })
  const canvas = page.locator('.pdf-page canvas').first()
  return {
    pages: await page.locator('.pdf-page').count(),
    landscape: await canvas.evaluate(
      (el) => el.getBoundingClientRect().width > el.getBoundingClientRect().height,
    ),
  }
}

async function edit(page: Page, kind: Kind) {
  if (kind === 'docs' || kind === 'markdown') {
    await page
      .locator(kind === 'docs' ? '.editor-scroll [contenteditable="true"] p' : '.doc-editor p')
      .first()
      .click()
    await page.keyboard.type(LOCAL)
  } else if (kind === 'html') {
    await read(page, kind)
    await page.locator('.source-editor .cm-content').click()
    await page.keyboard.type(`<!--${LOCAL}-->`)
  } else if (kind === 'sheets')
    await page.evaluate(
      (local) =>
        (window as any).__genofficeDebug.univerAPI
          .getActiveWorkbook()
          .getActiveSheet()
          .getRange('B1')
          .setValue(local),
      LOCAL,
    )
  else if (kind === 'pdf') {
    await page.getByRole('button', { name: 'Pages', exact: true }).click()
    await page.getByRole('button', { name: 'Rotate right', exact: true }).click()
  } else
    await page.evaluate(async () => {
      const slide = (await (window as any).slidesApi.getRenderSlides())[0]
      const node = slide.nodes.find((n: any) => n.text && n.box)
      await (window as any).slidesApi.setElementFont({
        slideIndex: 0,
        sourceIds: [node.sourceId],
        bold: true,
      })
    })
}

async function attemptSave(
  launched: LaunchedApp,
  page: Page,
  kind: Kind,
  path: string,
  original: Buffer,
) {
  if (kind === 'docs')
    return page.evaluate(
      ({ path, bytes }) =>
        (window as any).desktop.saveDocx(path, new Uint8Array(bytes).buffer, true),
      { path, bytes: [...original] },
    )
  if (kind === 'slides') return page.evaluate(() => (window as any).slidesApi.save())
  if (kind === 'pdf')
    return page.evaluate(
      (path) =>
        (window as any).pdfApi.save({
          path,
          markups: [],
          drawings: [],
          formValues: [],
          stamps: [],
        }),
      path,
    )
  if (kind === 'markdown' || kind === 'html')
    return page.evaluate(
      ({ kind }) =>
        (window as any)[`${kind}Api`].save({
          mode: 'save',
          text: 'Must not overwrite the external version',
        }),
      { kind },
    )
  await launched.app.evaluate(({ webContents }) =>
    webContents
      .getAllWebContents()
      .find((wc) => wc.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save'),
  )
  await expect(page.locator('.status-msg')).toContainText(/changed|modified/i)
  return { ok: false }
}

for (const kind of kinds) {
  test(`${kind}: its own save advances the baseline without an external-change notice`, async () => {
    const root = await mkdtemp(join(tmpdir(), 'external-own-save-'))
    const path = await fixture(root, kind)
    const original = await readFile(path)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: `external-own-save-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      await dialogs(launched, 'keep')
      const page = await openFreshEditor(launched, path, kind)
      await ready(page, kind)
      await edit(page, kind)
      await launched.app.evaluate(({ webContents }, kind) => {
        const wc = webContents
          .getAllWebContents()
          .find((wc) => wc.getURL().includes(`://${kind}/`))!
        const channel = {
          docs: 'menu:command',
          sheets: 'menu:action',
          slides: 'slides:menu',
          markdown: 'markdown:save-request',
          html: 'html:save-request',
          pdf: 'pdf:close-save-request',
        }[kind]
        wc.send(channel, 'save')
      }, kind)
      await expect.poll(async () => (await readFile(path)).equals(original)).toBe(false)
      // Cover two monitor ticks after the completed write, not just the immediate save result.
      await page.waitForTimeout(3200)
      expect(await notices(launched)).toEqual([])
      await launched.page.locator('.tab-item.active .tab-close').click()
      await expect(launched.page.locator('.tab-close')).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, `external-own-save-${kind}`)
      await rm(root, { recursive: true, force: true })
    }
  })

  test(`${kind}: detects atomic external replacement and reloads the latest version`, async () => {
    const root = await mkdtemp(join(tmpdir(), 'external-reload-'))
    const path = await fixture(root, kind)
    const original = await readFile(path)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: `external-reload-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      await dialogs(launched, 'reload')
      const page = await openFreshEditor(launched, path, kind)
      await ready(page, kind)
      const before = await read(page, kind)
      const external = await externalVersion(original, kind)
      await replace(path, external)
      await expect.poll(() => notices(launched)).toHaveLength(1)
      await expect.poll(() => read(page, kind)).not.toEqual(before)
      await ready(page, kind)
      if (kind === 'pdf') expect(await read(page, kind)).toEqual({ pages: 4, landscape: false })
      else expect(JSON.stringify(await read(page, kind))).toContain('External version 1')
      expect((await notices(launched))[0].buttons).toEqual(['Reload', 'Later'])
      expect(await readFile(path)).toEqual(external)
      await launched.page.locator('.tab-item.active .tab-close').click()
      await expect(launched.page.locator('.tab-close')).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, `external-reload-${kind}`)
      await rm(root, { recursive: true, force: true })
    }
  })

  test(`${kind}: retains unsaved edits, blocks overwriting, and saves a reopenable copy`, async () => {
    test.setTimeout(120_000)
    const root = await mkdtemp(join(tmpdir(), 'external-copy-'))
    const path = await fixture(root, kind)
    const copy = path.replace(/(\.[^.]+)$/, '-copy$1')
    if (kind === 'docs') {
      const zip = await JSZip.loadAsync(await readFile(path))
      zip.file('word/media/lazy.png', Buffer.alloc(16 * 1024 * 1024, 0x41))
      await writeFile(path, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
    }
    const original = await readFile(path)
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: `external-copy-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      await dialogs(launched, 'keep', copy)
      const page = await openFreshEditor(launched, path, kind)
      await ready(page, kind)
      const before = await read(page, kind)
      await edit(page, kind)
      await expect.poll(() => read(page, kind)).not.toEqual(before)
      const edited = await read(page, kind)
      await replace(path, await externalVersion(original, kind))
      await expect.poll(() => notices(launched)).toHaveLength(1)
      expect((await notices(launched))[0].buttons).toEqual([
        'Keep Editing',
        'Save a Copy…',
        'Reload',
      ])
      expect(await read(page, kind)).toEqual(edited)
      const external = await readFile(path)
      const result = await attemptSave(launched, page, kind, path, original)
      expect(result.ok).toBe(false)
      expect(await readFile(path)).toEqual(external)
      expect(await read(page, kind)).toEqual(edited)
      await launched.app.evaluate(() => {
        ;(globalThis as any).__externalAction = 'copy'
      })
      const external2 = await externalVersion(original, kind, 2)
      await replace(path, external2)
      await expect.poll(() => notices(launched)).toHaveLength(2)
      await expect
        .poll(async () => {
          try {
            return (await readFile(copy)).length
          } catch {
            return 0
          }
        })
        .toBeGreaterThan(0)
      expect(await readFile(path)).toEqual(external2)
      if (kind === 'docs') {
        const zip = await JSZip.loadAsync(await readFile(copy))
        expect(await zip.file('word/document.xml')!.async('string')).toContain(
          'w:footnoteReference',
        )
        expect(
          (await zip.file('word/media/lazy.png')!.async('nodebuffer')).equals(
            Buffer.alloc(16 * 1024 * 1024, 0x41),
          ),
        ).toBe(true)
      }
      await closeAndSaveVideo(launched, `external-copy-${kind}`)
      launched = await launchShell({
        onboardingSeen: true,
        videoDir: `external-copy-reopen-${kind}`,
        env: { GENOFFICE_DEBUG_HOOKS: '1' },
      })
      await dialogs(launched, 'keep')
      const reopened = await openFreshEditor(launched, copy, kind)
      await ready(reopened, kind)
      const content = await read(reopened, kind)
      if (kind === 'pdf' || kind === 'slides') expect(content).toEqual(edited)
      else await expect.poll(async () => String(await read(reopened, kind))).toContain(LOCAL)
      expect(await readFile(path)).toEqual(external2)
    } finally {
      await closeAndSaveVideo(launched, `external-copy-${kind}`)
      await rm(root, { recursive: true, force: true })
    }
  })

  test(`${kind}: explicitly reloads over unsaved edits only when the user chooses Reload`, async () => {
    const root = await mkdtemp(join(tmpdir(), 'external-discard-'))
    const path = await fixture(root, kind)
    const original = await readFile(path)
    const launched = await launchShell({
      onboardingSeen: true,
      videoDir: `external-discard-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      await dialogs(launched, 'reload')
      const page = await openFreshEditor(launched, path, kind)
      await ready(page, kind)
      await edit(page, kind)
      const external = await externalVersion(original, kind)
      await replace(path, external)
      await expect.poll(() => notices(launched)).toHaveLength(1)
      await ready(page, kind)
      if (kind === 'pdf')
        await expect.poll(() => read(page, kind)).toEqual({ pages: 4, landscape: false })
      else
        await expect
          .poll(async () => JSON.stringify(await read(page, kind)))
          .toContain('External version 1')
      expect((await notices(launched))[0].detail).toContain('unsaved edits')
      expect(await readFile(path)).toEqual(external)
    } finally {
      await closeAndSaveVideo(launched, `external-discard-${kind}`)
      await rm(root, { recursive: true, force: true })
    }
  })
}
