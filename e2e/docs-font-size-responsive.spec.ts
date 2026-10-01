import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

async function openDocument() {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-font-size' })
  await launched.page.locator('.quick-card').first().click()
  const page = await waitForPageWithUrl(launched.app, '://docs/')
  await page.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
  await page.locator('.doc-page').first().click()
  await page.keyboard.type('asdasdasd')
  return { launched, page }
}

async function paint(page: Page) {
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  )
}

async function fontSize(page: Page) {
  return page.evaluate(
    () => (window as any).__aidocs.editor.getAttributes('docTextStyle').sizeHalfPoints / 2,
  )
}

for (const paragraphCount of [1, 200]) {
  test(`font steps repaint throughout a burst in a ${paragraphCount}-paragraph document`, async () => {
    const testInfo = test.info()
    const { launched, page } = await openDocument()
    try {
      await page.evaluate((count) => {
        const editor = (window as any).__aidocs.editor
        editor.commands.setContent({
          type: 'doc',
          content: Array.from({ length: count }, () => ({
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'Selected text for font sizing.',
                marks: [{ type: 'docTextStyle', attrs: { sizeHalfPoints: 22 } }],
              },
            ],
          })),
        })
        editor.commands.selectAll()
      }, paragraphCount)
      const samples = await page.evaluate(async () => {
        const editor = (window as any).__aidocs.editor
        const button = document.querySelector<HTMLButtonElement>(
          'button[data-tip="Increase Font Size"]',
        )!
        const samples = []
        for (const expected of [12, 14, 16, 18, 20, 22, 24]) {
          const start = performance.now()
          button.click()
          await new Promise<void>((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
          )
          samples.push({
            expected,
            actual: editor.getAttributes('docTextStyle').sizeHalfPoints / 2,
            elapsedMs: performance.now() - start,
          })
        }
        return samples
      })
      expect(samples.map((sample) => sample.actual)).toEqual([12, 14, 16, 18, 20, 22, 24])
      await testInfo.attach('font-step-timing', {
        body: JSON.stringify(samples, null, 2),
        contentType: 'application/json',
      })
      for (const expected of [22, 20, 18, 16, 14, 12, 11]) {
        await page.getByRole('button', { name: 'Decrease Font Size', exact: true }).click()
        await paint(page)
        expect(await fontSize(page)).toBe(expected)
      }
      await page.keyboard.press('ControlOrMeta+Shift+.')
      await paint(page)
      expect(await fontSize(page)).toBe(12)
      await page.keyboard.press('ControlOrMeta+Shift+,')
      await paint(page)
      expect(await fontSize(page)).toBe(11)
    } finally {
      await closeAndSaveVideo(launched, 'docs-font-size')
    }
  })
}

test('decimal markers keep a visible gap without jumping to the next tab stop', async () => {
  const testInfo = test.info()
  const { launched, page } = await openDocument()
  try {
    await page.locator('button[data-tip="Numbering"]').first().click()
    const item = page.locator('.doc-li')
    await expect(item).toHaveCount(1)
    const observations = []
    for (const start of [1, 10, 100]) {
      await page.evaluate((value) => {
        const editor = (window as any).__aidocs.editor
        const numId = editor.state.doc.firstChild.attrs.numId
        const definitions = new Map(editor.storage.listNumbering.defs)
        const definition = definitions.get(numId) as any
        definitions.set(numId, {
          ...definition,
          startOverrides: { ...definition.startOverrides, 0: value },
        })
        editor.storage.listNumbering.defs = definitions
        editor.view.dispatch(editor.state.tr)
      }, start)
      const offsets: number[] = []
      for (const size of [22, 24, 36, 24, 22]) {
        await page.evaluate(
          (pt) =>
            (window as any).__aidocs.editor
              .chain()
              .selectAll()
              .setMark('docTextStyle', { sizeHalfPoints: pt * 2 })
              .run(),
          size,
        )
        await paint(page)
        await expect(item).toHaveAttribute('data-marker-flow', '')
        await expect(item).toHaveAttribute('data-marker', `${start}.`)
        const geometry = await item.evaluate((el) => {
          const range = document.createRange()
          const text = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode()!
          range.setStart(text, 0)
          range.setEnd(text, 1)
          const marker = getComputedStyle(el, '::before')
          const markerX =
            parseFloat(getComputedStyle(el).paddingInlineStart) +
            parseFloat(marker.marginInlineStart)
          const textX = range.getBoundingClientRect().left - el.getBoundingClientRect().left
          return { textX, gap: textX - markerX - parseFloat(marker.width) }
        })
        expect(geometry.gap).toBeCloseTo(8, 0)
        offsets.push(geometry.textX)
        observations.push({ start, size, ...geometry })
        if (start === 1 && offsets.length <= 2) {
          await item.screenshot({ path: testInfo.outputPath(`numbering-${size}.png`) })
        }
      }
      expect(offsets[1] - offsets[0]).toBeGreaterThan(0)
      expect(offsets[1] - offsets[0]).toBeLessThan(8)
      expect(offsets[3]).toBeCloseTo(offsets[1], 1)
      expect(offsets[4]).toBeCloseTo(offsets[0], 1)
    }
    await testInfo.attach('numbering-geometry', {
      body: JSON.stringify(observations, null, 2),
      contentType: 'application/json',
    })
  } finally {
    await closeAndSaveVideo(launched, 'docs-font-size')
  }
})

test('explicit paragraph tab stops keep their positioning when font size changes', async () => {
  const { launched, page } = await openDocument()
  try {
    await page.locator('button[data-tip="Numbering"]').first().click()
    await page.evaluate(() =>
      (window as any).__aidocs.editor
        .chain()
        .selectAll()
        .setMark('docTextStyle', { sizeHalfPoints: 48 })
        .updateAttributes('docListItem', { tabStops: JSON.stringify([{ pos: 1440, val: 'left' }]) })
        .run(),
    )
    await paint(page)
    const item = page.locator('.doc-li')
    await expect(item).not.toHaveAttribute('data-marker-flow')
    expect(
      await item.evaluate((el) => (el as HTMLElement).style.getPropertyValue('--li-tab').trim()),
    ).toBe('54pt')
    await page.getByRole('button', { name: 'Increase Font Size', exact: true }).click()
    await paint(page)
    expect(await fontSize(page)).toBe(26)
    expect(
      await item.evaluate((el) => (el as HTMLElement).style.getPropertyValue('--li-tab').trim()),
    ).toBe('54pt')
  } finally {
    await closeAndSaveVideo(launched, 'docs-font-size')
  }
})
