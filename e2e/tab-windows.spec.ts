import { test, expect } from '@playwright/test'
import { mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fixture, openFreshEditor } from './editor-lifecycle-helpers'
import { launchShell, closeAndSaveVideo } from './helpers'

for (const kind of ['docs', 'sheets', 'slides'] as const) {
  test(`${kind}: detach and merge the live unsaved editor, retaining undo and redo`, async () => {
    test.setTimeout(150_000)
    const root = await mkdtemp(join(tmpdir(), 'genoffice-tab-windows-'))
    const file = await fixture(root, kind)
    const l = await launchShell({
      onboardingSeen: true,
      videoDir: `tab-windows-${kind}`,
      settings: { autoSave: false },
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      let page = await openFreshEditor(l, file, kind)
      if (kind === 'docs') {
        await page.waitForFunction(() =>
          (window as any).__aidocs?.editor?.state.doc.textContent.includes(
            'Preview without editing',
          ),
        )
        await expect(page.locator('.editor-scroll')).toHaveAttribute('aria-busy', 'false')
      }
      if (kind === 'sheets')
        await page.waitForFunction(
          () => (window as any).__genofficeDebug?.undoHistoryDiagnostic().fileSessionId,
        )
      if (kind === 'slides') await expect(page.locator('.stage-rel')).toBeVisible()
      const marker = 'Retained across windows'
      await page.evaluate(
        async ({ kind, marker }) => {
          const w = window as any
          if (kind === 'docs') w.__aidocs.editor.commands.insertContent(`<p>${marker}</p>`)
          else if (kind === 'sheets')
            w.__genofficeDebug.univerAPI
              .getActiveWorkbook()
              .getActiveSheet()
              .getRange('B1')
              .setValue(marker)
          else
            await w.slidesApi.addElement({
              slideIndex: 0,
              kind: 'textbox',
              xPx: 100,
              yPx: 100,
              wPx: 400,
              hPx: 60,
              fitWidthPx: 1280,
              text: marker,
            })
        },
        { kind, marker },
      )
      const read = () =>
        page.evaluate(async (kind) => {
          const w = window as any
          if (kind === 'docs') return w.__aidocs?.editor?.state.doc.textContent ?? ''
          if (kind === 'sheets')
            return String(
              w.__genofficeDebug.univerAPI
                .getActiveWorkbook()
                .getActiveSheet()
                .getRange('B1')
                .getValue(),
            )
          return ((await w.slidesApi.getRenderSlides()) ?? [])
            .flatMap((s: any) => s.nodes)
            .flatMap((n: any) => n.text?.lines ?? [])
            .map((line: any) => line.runs.map((r: any) => r.text).join(''))
            .join('\n')
        }, kind)
      await expect.poll(read).toContain(marker)
      const tabs = await l.page.evaluate(() => window.aiOfficeTabs.list())
      const id = tabs.find((t) => t.kind === kind)!.id
      const bounds = await l.app.evaluate(({ BrowserWindow }) =>
        BrowserWindow.getAllWindows()[0].getContentBounds(),
      )
      const born = l.app.waitForEvent('window')
      expect(
        await l.page.evaluate(({ id, x, y }) => window.aiOfficeTabs.drop(id, x, y), {
          id,
          x: bounds.x + 220,
          y: bounds.y + 200,
        }),
      ).toBe(true)
      const second = await born
      await second.waitForFunction(() => !!window.aiOfficeTabs)
      await expect
        .poll(() => second.evaluate(() => window.aiOfficeTabs.list()))
        .toMatchObject([{ kind: 'home' }, { id, kind, active: true }])
      expect(
        (await l.page.evaluate(() => window.aiOfficeTabs.list())).some((t) => t.id === id),
      ).toBe(false)
      await expect.poll(read).toContain(marker)
      await page.evaluate(async (kind) => {
        const w = window as any
        if (kind === 'docs') w.__aidocs.editor.commands.undo()
        else if (kind === 'sheets') await w.__genofficeDebug.univerAPI.undo()
        else await w.slidesApi.undo()
      }, kind)
      await expect.poll(read).not.toContain(marker)
      await page.evaluate(async (kind) => {
        const w = window as any
        if (kind === 'docs') w.__aidocs.editor.commands.redo()
        else if (kind === 'sheets') await w.__genofficeDebug.univerAPI.redo()
        else await w.slidesApi.redo()
      }, kind)
      await expect.poll(read).toContain(marker)
      const saved = join(
        root,
        `moved.${kind === 'docs' ? 'docx' : kind === 'sheets' ? 'xlsx' : 'pptx'}`,
      )
      await l.app.evaluate(
        ({ dialog, webContents }, { kind, saved }) => {
          dialog.showSaveDialog = (async () => ({ canceled: false, filePath: saved })) as any
          const editor = webContents
            .getAllWebContents()
            .find((w) => w.getURL().includes(`://${kind}/`))!
          editor.send(
            kind === 'docs' ? 'menu:command' : kind === 'slides' ? 'slides:menu' : 'menu:action',
            'save-as',
          )
        },
        { kind, saved },
      )
      await expect
        .poll(() =>
          second.evaluate(() =>
            window.aiOfficeTabs.list().then((ts) => ts.find((t) => t.kind !== 'home')?.filePath),
          ),
        )
        .toBe(saved)
      expect((await readFile(saved)).length).toBeGreaterThan(0)
      expect(
        await second.evaluate(({ id, x, y }) => window.aiOfficeTabs.drop(id, x, y), {
          id,
          x: bounds.x + 250,
          y: bounds.y + 20,
        }),
      ).toBe(true)
      await expect.poll(() => second.isClosed()).toBe(true)
      expect(
        (await l.page.evaluate(() => window.aiOfficeTabs.list())).find((t) => t.id === id)?.active,
      ).toBe(true)
      await expect.poll(read).toContain(marker)
      expect(page.isClosed()).toBe(false)
      await l.page.evaluate((id) => window.aiOfficeTabs.close(id), id)
      page = await openFreshEditor(l, saved, kind)
      if (kind === 'sheets')
        await page.waitForFunction(
          () => (window as any).__genofficeDebug?.undoHistoryDiagnostic().fileSessionId,
        )
      await expect.poll(read).toContain(marker)
      // Discard only the disposable fixture when shutting down the test instance.
      await l.app.evaluate(({ dialog }) => {
        dialog.showMessageBox = (async () => ({ response: 1, checkboxChecked: false })) as any
      })
    } finally {
      await closeAndSaveVideo(l, `tab-windows-${kind}`)
    }
  })
}

test('tab pointer gestures reorder, cancel, and detach below the strip', async () => {
  test.setTimeout(120_000)
  const root = await mkdtemp(join(tmpdir(), 'genoffice-tab-gestures-'))
  const l = await launchShell({
    onboardingSeen: true,
    videoDir: 'tab-gestures',
    settings: { autoSave: false },
  })
  try {
    await openFreshEditor(l, await fixture(root, 'markdown'), 'markdown')
    await openFreshEditor(l, await fixture(root, 'html'), 'html')
    const list = await l.page.evaluate(() => window.aiOfficeTabs.list())
    const md = list.find((t) => t.kind === 'markdown')!.id
    const html = list.find((t) => t.kind === 'html')!.id
    const row = l.page.locator('.tab-strip .tab-item')
    let a = (await row.nth(1).boundingBox())!
    const b = (await row.nth(2).boundingBox())!
    await l.page.mouse.move(a.x + a.width / 2, a.y + a.height / 2)
    await l.page.mouse.down()
    await l.page.mouse.move(b.x + b.width - 5, b.y + b.height / 2, { steps: 10 })
    await l.page.mouse.up()
    await expect
      .poll(() =>
        l.page.evaluate(() => window.aiOfficeTabs.list().then((ts) => ts.map((t) => t.id))),
      )
      .toEqual(['home', html, md])
    expect(await l.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(
      1,
    )
    a = (await row.nth(2).boundingBox())!
    await l.page.mouse.move(a.x + a.width / 2, a.y + a.height / 2)
    await l.page.mouse.down()
    await l.page.mouse.move(a.x + a.width / 2, a.y + 100, { steps: 8 })
    await l.page.keyboard.press('Escape')
    await l.page.mouse.up()
    expect(await l.app.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows().length)).toBe(
      1,
    )
    const born = l.app.waitForEvent('window')
    await l.page.mouse.move(a.x + a.width / 2, a.y + a.height / 2)
    await l.page.mouse.down()
    await l.page.mouse.move(a.x + a.width / 2, a.y + 150, { steps: 8 })
    await l.page.mouse.up()
    const second = await born
    await second.waitForFunction(() => !!window.aiOfficeTabs)
    await expect
      .poll(() =>
        second.evaluate(() => window.aiOfficeTabs.list().then((ts) => ts.map((t) => t.id))),
      )
      .toEqual(['home', md])
    expect((await l.page.evaluate(() => window.aiOfficeTabs.list())).map((t) => t.id)).toEqual([
      'home',
      html,
    ])
  } finally {
    await closeAndSaveVideo(l, 'tab-gestures')
  }
})
