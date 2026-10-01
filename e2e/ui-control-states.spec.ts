import { test, expect, type Page } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

async function stableControlShapes(page: Page, scope: string) {
  return page.evaluate((selector) => {
    const failures: string[] = []
    const controls = Array.from(document.querySelectorAll<HTMLButtonElement>(selector)).filter(
      (button) => button.getBoundingClientRect().width > 0,
    )
    const shape = (button: HTMLButtonElement) => {
      const css = getComputedStyle(button)
      return [
        css.borderTopLeftRadius,
        css.borderTopRightRadius,
        css.borderBottomRightRadius,
        css.borderBottomLeftRadius,
        css.padding,
        css.height,
      ].join('|')
    }
    for (const button of controls) {
      const original = button.className
      const disabled = button.disabled
      const pressed = button.getAttribute('aria-pressed')
      const baseline = shape(button)
      const states = ['active', 'selected', 'on', 'is-active']
      for (const state of [...states, 'disabled', 'pressed']) {
        button.className = original
        button.disabled = disabled
        if (pressed !== null) button.setAttribute('aria-pressed', pressed)
        if (state === 'disabled') button.disabled = !disabled
        else if (state === 'pressed') {
          if (pressed === null) continue
          button.setAttribute('aria-pressed', pressed === 'true' ? 'false' : 'true')
        } else button.classList.toggle(state)
        const after = shape(button)
        if (after !== baseline)
          failures.push(
            `${button.textContent?.trim() || button.getAttribute('aria-label') || original}: ${state}: ${baseline} -> ${after}`,
          )
      }
      button.className = original
      button.disabled = disabled
      if (pressed !== null) button.setAttribute('aria-pressed', pressed)
    }
    return { count: controls.length, failures }
  }, scope)
}

async function prepare(page: Page, theme: string) {
  await page.evaluate((value) => document.documentElement.setAttribute('data-theme', value), theme)
  await page.addStyleTag({
    content: '*, *::before, *::after { transition: none !important; animation: none !important; }',
  })
}

async function chartDialog(page: Page) {
  await page.locator('.ribbon-tab', { hasText: /^插入$/ }).click()
  await page.locator('.ribbon-body button[data-tip^="插入圖表"]').click()
  await expect(page.locator('.modal-chart')).toBeVisible()
}

for (const theme of ['light', 'dark']) {
  test(`Docs chart buttons keep geometry in every state (${theme})`, async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'zh-TW',
      videoDir: 'ui-controls',
    })
    try {
      await launched.page.locator('.quick-card').first().click()
      const page = await waitForPageWithUrl(launched.app, '://docs/')
      await page.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
      await prepare(page, theme)
      await chartDialog(page)
      const choices = page.locator('.chart-kind-button')
      const geometry = () =>
        choices.evaluateAll((buttons) =>
          buttons.map((button) => {
            const rect = button.getBoundingClientRect()
            return [rect.width, rect.height, getComputedStyle(button).borderRadius]
          }),
        )
      const before = await geometry()
      for (const size of before) {
        // Equal flex tracks can differ by a fraction of a CSS pixel at Windows display scaling.
        expect(Number(size[0])).toBeCloseTo(Number(before[0][0]), 0)
        expect(size[1]).toBe(before[0][1])
        expect(size[2]).toBe(before[0][2])
      }
      expect(parseFloat(String(before[0][2]))).toBeGreaterThan(0)
      for (let index = 0; index < 3; index++) {
        await choices.nth(index).click()
        await expect(choices.nth(index)).toHaveAttribute('aria-pressed', 'true')
        await expect(page.locator('.chart-kind-button[aria-pressed="true"]')).toHaveCount(1)
        expect(await geometry()).toEqual(before)
        await choices.nth((index + 1) % 3).hover()
        expect(await geometry()).toEqual(before)
      }
      await expect(page.locator('.chart-data-action').nth(1)).toBeDisabled()
      const audit = await stableControlShapes(page, '.modal-chart button')
      expect(audit.failures).toEqual([])
      await choices.nth(1).click()
      await page
        .locator('.modal-chart')
        .screenshot({ path: test.info().outputPath(`chart-${theme}.png`) })
      await page.locator('.modal-chart > .modal-row input').fill('Review chart')
      await page.locator('.chart-data-grid tbody td input').first().fill('42')
      await page.locator('.modal-chart .modal-actions .btn-primary').click()
      await expect(page.locator('.modal-chart')).toHaveCount(0)
      const data = await page.evaluate(() => (window as any).__aidocs.editor.getJSON())
      const chart = data.content.find((node: any) => node.attrs?.genChart)?.attrs.genChart
      expect(chart).toMatchObject({ kind: 'line', title: 'Review chart' })
      expect(chart.series[0].values[0]).toBe(42)
    } finally {
      await closeAndSaveVideo(launched, `ui-chart-${theme}`)
    }
  })
}

test('Docs chart data scrolls inside a small dialog and Escape cancels', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'zh-TW',
    videoDir: 'ui-controls',
  })
  try {
    await launched.page.locator('.quick-card').first().click()
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
    await launched.app.evaluate(({ BrowserWindow }) =>
      BrowserWindow.getAllWindows()[0].setContentSize(880, 620),
    )
    await prepare(page, 'dark')
    await chartDialog(page)
    for (let i = 0; i < 10; i++) {
      await page.locator('.chart-data-action').first().click()
      await page.locator('.chart-data-action').nth(1).click()
    }
    const overflow = await page.locator('.chart-data-scroll').evaluate((el) => ({
      x: el.scrollWidth > el.clientWidth,
      y: el.scrollHeight > el.clientHeight,
    }))
    expect(overflow).toEqual({ x: true, y: true })
    await expect(page.locator('.modal-chart h2')).toBeInViewport()
    await expect(page.locator('.modal-chart .modal-actions button').last()).toBeInViewport()
    const last = page.locator('.chart-data-grid input').last()
    await last.fill('99')
    await expect(last).toBeInViewport()
    await page
      .locator('.modal-chart')
      .screenshot({ path: test.info().outputPath('chart-overflow.png') })
    await page.keyboard.press('Escape')
    await expect(page.locator('.modal-chart')).toHaveCount(0)
    expect(
      await page.evaluate(() =>
        (window as any).__aidocs.editor.getJSON().content.some((node: any) => node.attrs?.genChart),
      ),
    ).toBe(false)
  } finally {
    await closeAndSaveVideo(launched, 'ui-chart-scroll')
  }
})

function blankPdf() {
  let body = '%PDF-1.4\n'
  const offsets: number[] = []
  const objects = [
    '<</Type/Catalog/Pages 2 0 R>>',
    '<</Type/Pages/Kids[3 0 R]/Count 1>>',
    '<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]>>',
  ]
  objects.forEach((object, i) => {
    offsets.push(body.length)
    body += `${i + 1} 0 obj\n${object}\nendobj\n`
  })
  const start = body.length
  body += `xref\n0 4\n0000000000 65535 f \n${offsets.map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`).join('')}trailer\n<</Size 4/Root 1 0 R>>\nstartxref\n${start}\n%%EOF\n`
  return Buffer.from(body, 'latin1')
}

for (const app of ['shell', 'docs', 'sheets', 'slides', 'pdf', 'markdown', 'html']) {
  test(`${app}: button state and focus shape audit in both themes`, async () => {
    let openFile: string | undefined
    if (['pdf', 'markdown', 'html'].includes(app)) {
      const directory = await mkdtemp(join(tmpdir(), 'genoffice-control-states-'))
      openFile = join(directory, app === 'markdown' ? 'audit.md' : `audit.${app}`)
      await writeFile(
        openFile,
        app === 'pdf'
          ? blankPdf()
          : app === 'markdown'
            ? '# Audit\n\nUI review.'
            : '<html><body><p>UI review.</p></body></html>',
      )
    }
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'ui-controls', openFile })
    try {
      const card = ['docs', 'sheets', 'slides'].indexOf(app)
      if (card >= 0) await launched.page.locator('.quick-card').nth(card).click()
      const page =
        app === 'shell' ? launched.page : await waitForPageWithUrl(launched.app, `://${app}/`)
      await page.waitForFunction(() => document.querySelectorAll('button').length > 5)
      const reports: unknown[] = []
      for (const theme of ['light', 'dark']) {
        await prepare(page, theme)
        const labels = [
          'Home',
          'Insert',
          'Draw',
          'Design',
          'Layout',
          'Page Layout',
          'Formulas',
          'Data',
          'Transitions',
          'Animations',
          'Review',
          'View',
        ]
        const tabs = page.locator('.ribbon-tab, .ribbon-tabs > button:not(.qa-btn)')
        for (const label of ['initial', ...labels]) {
          if (label !== 'initial') {
            const tab = tabs.filter({ hasText: new RegExp(`^${label}$`) }).first()
            if (!(await tab.isVisible())) continue
            await tab.click()
          }
          const audit = await stableControlShapes(page, 'button')
          reports.push({ theme, tab: label, ...audit })
          expect.soft(audit.failures, `${app}/${theme}/${label}`).toEqual([])
        }
        if (app !== 'shell') {
          // These controls require an AI run to appear. Probe their real, built CSS
          // in the editor without making paid network requests or fabricating AI output.
          const focusAudit = await page.evaluate((isSlides) => {
            const host = document.createElement('div')
            host.style.cssText = 'position:fixed;left:-10000px;top:0'
            document.body.append(host)
            const results = []
            for (const cls of [
              'ai-work-group-summary',
              'ai-step-title clickable',
              ...(isSlides ? ['deck-progress-head'] : []),
            ]) {
              const button = document.createElement('button')
              button.className = cls
              button.textContent = 'Audit'
              host.append(button)
              const before = getComputedStyle(button).borderRadius
              button.focus()
              const after = getComputedStyle(button).borderRadius
              results.push({ cls, before, after })
            }
            host.remove()
            return results
          }, app === 'slides')
          for (const result of focusAudit) {
            expect(result.after).toBe(result.before)
            expect(parseFloat(result.before)).toBeGreaterThan(0)
          }
        }
      }
      await test.info().attach(`${app}-state-audit`, {
        body: JSON.stringify(reports, null, 2),
        contentType: 'application/json',
      })
    } finally {
      await closeAndSaveVideo(launched, `ui-states-${app}`)
    }
  })
}
