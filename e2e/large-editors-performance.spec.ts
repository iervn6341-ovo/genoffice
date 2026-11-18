import { expect, test, type Page } from '@playwright/test'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import JSZip from 'jszip'
import pptxgen from 'pptxgenjs'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(!process.env.GENOFFICE_LARGE_PERF, 'Opt-in large editor performance and roundtrip checks')
const root = '/private/tmp/genoffice-large-editors'
const rows = 50_000
const slideCount = 120

async function fixtures() {
  await mkdir(root, { recursive: true })
  const xlsx = join(root, 'large.xlsx')
  const pptx = join(root, 'large.pptx')
  const zip = new JSZip()
  const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  zip.file(
    '[Content_Types].xml',
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<workbook xmlns="${ns}" xmlns:r="${rel}"><sheets><sheet name="Large" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/></Relationships>`,
  )
  // A complete Normal style is required by the sidecar's strict formula importer.
  zip.file(
    'xl/styles.xml',
    `<styleSheet xmlns="${ns}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  )
  zip.file(
    '[Content_Types].xml',
    (await zip.file('[Content_Types].xml')!.async('string')).replace(
      '</Types>',
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    ),
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<worksheet xmlns="${ns}"><dimension ref="A1:P${rows}"/><sheetData>${Array.from({ length: rows }, (_, r) => `<row r="${r + 1}">${Array.from({ length: 16 }, (_, c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}">${c === 15 ? `<f>A${r + 1}*2</f>` : ''}<v>${c === 15 ? (r + 1) * 2 : r + c + 1}</v></c>`).join('')}</row>`).join('')}</sheetData></worksheet>`,
  )
  await writeFile(xlsx, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
  const deck = new pptxgen()
  deck.layout = 'LAYOUT_WIDE'
  for (let i = 0; i < slideCount; i++) {
    const s = deck.addSlide()
    s.addText(`Slide ${i + 1}`, { x: 0.5, y: 0.4, w: 10, h: 0.5, fontSize: 28 })
    for (let n = 0; n < 16; n++)
      s.addText(`Item ${n + 1} on page ${i + 1}`, {
        x: 0.5 + (n % 4) * 3,
        y: 1.3 + Math.floor(n / 4) * 1.2,
        w: 2.8,
        h: 0.8,
        fontSize: 18,
      })
  }
  await deck.writeFile({ fileName: pptx })
  return { sheets: xlsx, slides: pptx }
}

async function ready(page: Page, kind: 'sheets' | 'slides') {
  if (kind === 'sheets')
    await page.waitForFunction(
      () =>
        (window as any).__genofficeDebug?.univerAPI
          ?.getActiveWorkbook()
          ?.getActiveSheet()
          ?.getRange('A1')
          .getValue() === 1,
      null,
      { timeout: 90_000 },
    )
  else await expect(page.locator('.slide-list .thumb')).toHaveCount(slideCount, { timeout: 90_000 })
}

for (const kind of ['sheets', 'slides'] as const)
  test(`${kind}: large open, native resize, edit/history/save/reopen`, async () => {
    test.setTimeout(240_000)
    const paths = await fixtures()
    const originalLastSlide =
      kind === 'slides'
        ? await (
            await JSZip.loadAsync(await readFile(paths.slides))
          )
            .file(`ppt/slides/slide${slideCount}.xml`)!
            .async('string')
        : null
    let launched = await launchShell({
      onboardingSeen: true,
      videoDir: `large-${kind}`,
      env: { GENOFFICE_DEBUG_HOOKS: '1' },
    })
    try {
      await launched.page
        .locator('.quick-card', { hasText: kind === 'sheets' ? 'AI Sheets' : 'AI Slides' })
        .click()
      let page = await waitForPageWithUrl(launched.app, `://${kind}/`)
      await page.waitForSelector('[data-ribbon-body]')
      await page.waitForTimeout(1800)
      const cdp = await page.context().newCDPSession(page)
      await cdp.send('Performance.enable')
      await cdp.send('Profiler.enable')
      const metrics = async () =>
        Object.fromEntries(
          (await cdp.send('Performance.getMetrics')).metrics
            .filter((m: any) =>
              ['TaskDuration', 'ScriptDuration', 'LayoutDuration', 'LayoutCount'].includes(m.name),
            )
            .map((m: any) => [m.name, m.value]),
        )
      const cpu = () =>
        launched.app.evaluate(({ app }) =>
          app
            .getAppMetrics()
            .map((p) => ({ pid: p.pid, type: p.type, seconds: p.cpu.cumulativeCPUUsage ?? 0 })),
        )
      await page.evaluate(() => {
        ;(window as any).__longTasks = []
        new PerformanceObserver((list) => {
          for (const e of list.getEntries()) (window as any).__longTasks.push(e.duration)
        }).observe({ type: 'longtask' })
      })
      await launched.app.evaluate(({ dialog }, path) => {
        dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [path] })
      }, paths[kind])
      const beforeOpen = await metrics()
      const start = Date.now()
      await cdp.send('Profiler.start')
      await launched.app.evaluate(({ webContents }, k) => {
        webContents
          .getAllWebContents()
          .find((w) => w.getURL().includes(`://${k}/`))
          ?.send(k === 'sheets' ? 'menu:action' : 'slides:menu', 'open')
      }, kind)
      await ready(page, kind)
      const visibleMs = Date.now() - start
      await page.waitForTimeout(3000)
      const afterOpen = await metrics()
      const openProfile = (await cdp.send('Profiler.stop')).profile
      const openTasks = await page.evaluate(() => (window as any).__longTasks)
      const canvases = await page.locator('canvas').count()
      await cdp.send('Profiler.start')
      const beforeResize = await metrics()
      const cpuBefore = await cpu()
      await launched.app.evaluate(async ({ BrowserWindow }) => {
        const win = BrowserWindow.getAllWindows()[0]!
        for (let i = 0; i < 60; i++) {
          win.setSize(850 + (i % 15) * 28, 700 + (i % 5) * 12)
          await new Promise((r) => setTimeout(r, 16))
        }
        win.setSize(1280, 800)
      })
      await page.waitForTimeout(2000)
      const afterResize = await metrics()
      const cpuAfter = await cpu()
      const resizeProfile = (await cdp.send('Profiler.stop')).profile
      const report = {
        kind,
        fixture: paths[kind],
        visibleMs,
        canvases,
        beforeOpen,
        afterOpen,
        beforeResize,
        afterResize,
        cpuBefore,
        cpuAfter,
        openTasks,
      }
      const prefix = `${process.env.GENOFFICE_LARGE_PERF_OUTPUT ?? join(root, 'measure')}-${kind}`
      await writeFile(`${prefix}.json`, JSON.stringify(report, null, 2))
      await writeFile(`${prefix}-open.cpuprofile`, JSON.stringify(openProfile))
      await writeFile(`${prefix}-resize.cpuprofile`, JSON.stringify(resizeProfile))
      await page.screenshot({ path: `${prefix}.png` })
      console.log(JSON.stringify(report))
      if (!process.env.GENOFFICE_LARGE_PERF_ROUNDTRIP) return
      // The UI history/save paths below are shared with normal interactive edits.
      if (kind === 'sheets') {
        await page.evaluate(() => {
          ;(window as any).__genofficeDebug.univerAPI
            .getActiveWorkbook()
            .getActiveSheet()
            .getRange('A1')
            .setValue(12345)
        })
        const value = () =>
          page.evaluate(() =>
            (window as any).__genofficeDebug.univerAPI
              .getActiveWorkbook()
              .getActiveSheet()
              .getRange('A1')
              .getValue(),
          )
        await expect.poll(value).toBe(12345)
        await page.locator('button[aria-label^="Undo"]').first().click()
        await expect.poll(value).toBe(1)
        await page.locator('button[aria-label^="Redo"]').first().click()
        await expect.poll(value).toBe(12345)
        await page.locator('button[aria-label^="Save ("]').first().click()
      } else {
        // Edit the first existing title through the actual canvas text overlay.
        const box = await page.locator('.stage-rel').boundingBox()
        expect(box).not.toBeNull()
        await page.mouse.dblclick(box!.x + box!.width * 0.2, box!.y + box!.height * 0.09)
        await expect(page.locator('.slide-text-editor')).toBeVisible()
        await page.locator('.slide-text-editor').fill('Performance edited')
        await page.keyboard.press('Escape')
        const text = () =>
          page.evaluate(async () =>
            (await (window as any).slidesApi.getRenderSlides())[0].nodes[0].text.lines
              .map((line: any) => line.runs.map((run: any) => run.text).join(''))
              .join('\n'),
          )
        await expect.poll(text).toContain('Performance edited')
        await page.locator('.qa-btn[data-tip^="Undo"]').click()
        await expect.poll(text).not.toContain('Performance edited')
        await page.locator('.qa-btn[data-tip^="Redo"]').click()
        await expect.poll(text).toContain('Performance edited')
        await page.locator('.qa-btn[aria-label^="Save"]').first().click()
      }
      await expect
        .poll(
          async () => {
            // A write can briefly expose an incomplete zip; retry until the save settles.
            try {
              const z = await JSZip.loadAsync(await readFile(paths[kind]))
              const xml = await z
                .file(kind === 'sheets' ? 'xl/worksheets/sheet1.xml' : 'ppt/slides/slide1.xml')!
                .async('string')
              return kind === 'sheets'
                ? /<c\b[^>]*\br="A1"[^>]*>[\s\S]*?<v>12345<\/v>[\s\S]*?<\/c>/.test(
                    xml.match(/<c\b[^>]*\br="A1"[^>]*>[\s\S]*?<\/c>/)?.[0] ?? '',
                  )
                : xml.includes('Performance edited')
            } catch {
              return false
            }
          },
          { timeout: 60_000 },
        )
        .toBe(true)
      await closeAndSaveVideo(launched, `large-${kind}-saved`)
      launched = await launchShell({
        onboardingSeen: true,
        openFile: paths[kind],
        videoDir: `large-${kind}-reopened`,
        env: { GENOFFICE_DEBUG_HOOKS: '1' },
      })
      page = await waitForPageWithUrl(launched.app, `://${kind}/`)
      if (kind === 'sheets') {
        await page.waitForFunction(
          () =>
            (window as any).__genofficeDebug?.univerAPI
              ?.getActiveWorkbook()
              ?.getActiveSheet()
              ?.getRange('A1')
              .getValue() === 12345,
          null,
          { timeout: 90_000 },
        )
        // Far-away rows and formula caches remain in the saved package.
        const z = await JSZip.loadAsync(await readFile(paths.sheets))
        const xml = await z.file('xl/worksheets/sheet1.xml')!.async('string')
        const dependent = xml.match(/<c\b[^>]*\br="P1"[^>]*>[\s\S]*?<\/c>/)?.[0]
        expect(dependent).toContain('<f>A1*2</f>')
        expect(dependent).toContain('<v>24690</v>')
        expect(xml).toContain('<f>A50000*2</f>')
        expect(xml).toContain('<v>100000</v>')
      } else {
        await ready(page, kind)
        const savedZip = await JSZip.loadAsync(await readFile(paths.slides))
        expect(await savedZip.file(`ppt/slides/slide${slideCount}.xml`)!.async('string')).toBe(
          originalLastSlide,
        )
        expect(
          await page.evaluate(async () =>
            (await (window as any).slidesApi.getRenderSlides())[0].nodes[0].text.lines
              .map((line: any) => line.runs.map((run: any) => run.text).join(''))
              .join('\n'),
          ),
        ).toContain('Performance edited')
        await page.locator('.slide-list .thumb').last().scrollIntoViewIfNeeded()
        await page.locator('.slide-list .thumb').last().click()
        await expect(page.locator('.slide-list .thumb').last()).toHaveClass(/active/)
        await expect(page.locator('.slide-list .thumb').last().locator('canvas')).toHaveCount(1)
        expect(await page.locator('canvas').count()).toBeLessThan(20)
      }
    } finally {
      await closeAndSaveVideo(launched, `large-${kind}`)
    }
  })
