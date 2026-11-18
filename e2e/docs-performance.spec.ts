import { expect, test } from '@playwright/test'
import { writeFile } from 'node:fs/promises'
import { resolve } from 'node:path'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

test.skip(!process.env.GENOFFICE_PERF_VERIFY, 'Opt-in performance measurement')

test('measure TOC document opening and rapid native window resizing', async () => {
  test.setTimeout(180_000)
  const fixture = resolve(__dirname, '../我的碩士論文計畫書_修訂版.docx')
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-performance' })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI Docs' }).click()
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
    const cdp = await page.context().newCDPSession(page)
    await cdp.send('Performance.enable')
    await cdp.send('Profiler.enable')
    const cpu = () =>
      launched.app.evaluate(({ app }) =>
        app.getAppMetrics().map((p) => ({
          pid: p.pid,
          type: p.type,
          seconds: p.cpu.cumulativeCPUUsage ?? 0,
        })),
      )
    const metrics = async () =>
      Object.fromEntries(
        (await cdp.send('Performance.getMetrics')).metrics
          .filter((m: any) =>
            [
              'TaskDuration',
              'ScriptDuration',
              'LayoutDuration',
              'RecalcStyleDuration',
              'LayoutCount',
            ].includes(m.name),
          )
          .map((m: any) => [m.name, m.value]),
      )
    await page.evaluate(() => {
      ;(window as any).__perfLongTasks = []
      new PerformanceObserver((list) => {
        for (const entry of list.getEntries())
          (window as any).__perfLongTasks.push({ start: entry.startTime, duration: entry.duration })
      }).observe({ type: 'longtask', buffered: false })
    })
    await cdp.send('Profiler.start')
    const beforeOpen = await metrics()
    const openStarted = Date.now()
    await page.evaluate((path) => (window as any).__aidocs.openPath(path), fixture)
    await expect(page.locator('.doc-toc-line').first()).toBeVisible({ timeout: 60_000 })
    const visibleMs = Date.now() - openStarted
    await page.waitForFunction(() => (window as any).__aidocs.editor.isEditable, null, {
      timeout: 60_000,
    })
    const editableMs = Date.now() - openStarted
    await page.evaluate(() => document.fonts.ready)
    await page.waitForTimeout(2000)
    const afterOpen = await metrics()
    const openProfile = (await cdp.send('Profiler.stop')).profile
    const openState = await page.evaluate(() => ({
      nodes: (window as any).__aidocs.editor.state.doc.childCount,
      text: (window as any).__aidocs.editor.state.doc.textContent.length,
      toc: document.querySelectorAll('.doc-toc-line').length,
      longTasks: (window as any).__perfLongTasks,
      pagination: (window as any).__pageDebug,
    }))
    await cdp.send('Profiler.start')
    const beforeResize = await metrics()
    const cpuBefore = await cpu()
    const resizeStarted = Date.now()
    await launched.app.evaluate(async ({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0]!
      for (let i = 0; i < 60; i++) {
        win.setSize(850 + (i % 15) * 28, 700 + (i % 5) * 12)
        await new Promise((r) => setTimeout(r, 16))
      }
      win.setSize(1280, 800)
    })
    await page.waitForTimeout(2000)
    const resizeMs = Date.now() - resizeStarted
    const cpuAfter = await cpu()
    const afterResize = await metrics()
    const resizeProfile = (await cdp.send('Profiler.stop')).profile
    const resizeState = await page.evaluate(() => ({
      nodes: (window as any).__aidocs.editor.state.doc.childCount,
      text: (window as any).__aidocs.editor.state.doc.textContent.length,
      zoom: document.querySelector('.zoom-value')?.textContent,
      pagination: (window as any).__pageDebug,
    }))
    expect(resizeState.nodes).toBe(openState.nodes)
    expect(resizeState.text).toBe(openState.text)
    expect(resizeState.zoom).toBe('100%')
    const report = {
      fixture,
      visibleMs,
      editableMs,
      resizeMs,
      beforeOpen,
      afterOpen,
      beforeResize,
      afterResize,
      cpuBefore,
      cpuAfter,
      openState,
      resizeState,
    }
    const prefix = process.env.GENOFFICE_PERF_OUTPUT ?? '/private/tmp/genoffice-perf'
    await writeFile(`${prefix}.json`, JSON.stringify(report, null, 2))
    await writeFile(`${prefix}-open.cpuprofile`, JSON.stringify(openProfile))
    await writeFile(`${prefix}-resize.cpuprofile`, JSON.stringify(resizeProfile))
    console.log(
      JSON.stringify({
        visibleMs,
        editableMs,
        resizeMs,
        beforeOpen,
        afterOpen,
        beforeResize,
        afterResize,
        cpuDelta: cpuAfter.map((p) => ({
          type: p.type,
          seconds: p.seconds - (cpuBefore.find((b) => b.pid === p.pid)?.seconds ?? p.seconds),
        })),
        longTasks: openState.longTasks.length,
        longestTask: Math.max(0, ...openState.longTasks.map((t: any) => t.duration)),
        nodes: openState.nodes,
        toc: openState.toc,
      }),
    )
  } finally {
    await closeAndSaveVideo(launched, 'docs-performance')
  }
})
