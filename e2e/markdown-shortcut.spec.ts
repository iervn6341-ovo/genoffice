import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Typing "#" … "######" at the start of a plain paragraph shows a hint; Enter turns the
 * paragraph into that heading instead of splitting it, and ⌘Z brings the marker back.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

const blocks = (p: Page) =>
  p.evaluate(() => {
    const out: Array<{ type: string; text: string; level: number | null }> = []
    ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((n: any) =>
      out.push({ type: n.type.name, text: n.textContent, level: n.attrs.level ?? null }),
    )
    return out
  })

test('docs: Markdown hashes + Enter make a heading', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'md-shortcut' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()

    await p.keyboard.type('##')
    await expect(p.locator('.md-shortcut-hint')).toHaveCount(1)
    await expect(p.locator('.md-shortcut-hint')).toContainText('Heading 2')

    await p.keyboard.press('Enter')
    await expect
      .poll(() => blocks(p).then((b) => b[0]))
      .toMatchObject({
        type: 'docHeading',
        text: '',
        level: 2,
      })
    await expect(p.locator('.md-shortcut-hint')).toHaveCount(0)

    // the heading is live: typing goes into it, Enter then splits as usual
    await p.keyboard.type('Intro')
    await p.keyboard.press('Enter')
    await expect.poll(() => blocks(p).then((b) => b.length)).toBe(2)

    // marker followed by a title converts in place
    await p.keyboard.type('### Details')
    await p.keyboard.press('Enter')
    await expect
      .poll(() => blocks(p).then((b) => b[1]))
      .toMatchObject({
        type: 'docHeading',
        text: 'Details',
        level: 3,
      })
    expect((await blocks(p)).length).toBe(2)

    // its own undo step: the hashes come back as text
    await p.keyboard.press('ControlOrMeta+z')
    await expect
      .poll(() => blocks(p).then((b) => b[1]))
      .toMatchObject({
        type: 'docParagraph',
        text: '### Details',
      })

    // deleting back to "###" shows the hint again
    await p.keyboard.press('End')
    await p.keyboard.press('Backspace')
    await p.keyboard.press('Backspace')
    await expect(p.locator('.md-shortcut-hint')).toHaveCount(1)
  } finally {
    await closeAndSaveVideo(launched, 'md-shortcut')
  }
})

test('docs: typed Markdown table rows + Enter make a table', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'md-table' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()

    // one row alone is just text: no hint, Enter splits as usual
    await p.keyboard.type('| Hi | Hi2 |')
    await expect(p.locator('.md-shortcut-hint')).toHaveCount(0)
    await p.keyboard.press('Enter')
    await p.keyboard.type('| -- | -- |')
    await expect(p.locator('.md-shortcut-hint')).toContainText('table')

    await p.keyboard.press('Enter')
    await expect
      .poll(() => blocks(p).then((b) => b.map((x) => x.type)))
      .toEqual(['docTable', 'docParagraph'])
    const shape = await p.evaluate(() => {
      const table = (window as unknown as DocsWindow).__aidocs!.editor.state.doc.child(0)
      const rows: string[][] = []
      table.forEach((row: any) => {
        const cells: string[] = []
        row.forEach((cell: any) => cells.push(cell.textContent))
        rows.push(cells)
      })
      return { rows, borders: Boolean(table.attrs.borders) }
    })
    // header + the empty body row the caret lands in
    expect(shape).toEqual({
      rows: [
        ['Hi', 'Hi2'],
        ['', ''],
      ],
      borders: true,
    })

    // the caret is in the first body cell: typing fills it
    await p.keyboard.type('21')
    await p.keyboard.press('Tab')
    await p.keyboard.type('22')
    const cells = await p.evaluate(() => {
      const table = (window as unknown as DocsWindow).__aidocs!.editor.state.doc.child(0)
      return table.child(1).textContent
    })
    expect(cells).toBe('2122')
  } finally {
    await closeAndSaveVideo(launched, 'md-table')
  }
})

test('docs: undo brings typed table rows back; header-less rows convert too', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'md-table2' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    await p.keyboard.type('| a | b |')
    await p.keyboard.press('Enter')
    await p.keyboard.type('| 1 | 2 |')
    await p.keyboard.press('Enter')
    await expect.poll(() => blocks(p).then((b) => b[0]!.type)).toBe('docTable')
    await p.keyboard.press('ControlOrMeta+z')
    await expect
      .poll(() => blocks(p).then((b) => b.map((x) => x.text)))
      .toEqual(['| a | b |', '| 1 | 2 |'])
  } finally {
    await closeAndSaveVideo(launched, 'md-table2')
  }
})

test('docs: table rows typed with full-width bars (CJK input) are recognised', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'zh-TW', videoDir: 'md-table3' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    // an input method commits text through insertText, not key events
    await p.keyboard.insertText('｜ 名稱 ｜ 數量 ｜')
    await p.keyboard.press('Enter')
    await p.keyboard.insertText('｜ －－ ｜ －－ ｜')
    await expect(p.locator('.md-shortcut-hint')).toContainText('表格')
    await p.keyboard.press('Enter')
    await expect.poll(() => blocks(p).then((b) => b[0]!.type)).toBe('docTable')
  } finally {
    await closeAndSaveVideo(launched, 'md-table3')
  }
})

test('docs: table rows typed with Shift+Enter (one paragraph, line breaks) become a table', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'md-table4' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    await p.keyboard.type('| Hi| Hi |')
    await p.keyboard.press('Shift+Enter')
    await p.keyboard.type('| -- | -- |')
    await p.keyboard.press('Shift+Enter')
    await p.keyboard.type('|hi2 | 2 |')
    await expect(p.locator('.md-shortcut-hint')).toContainText('table')
    await p.keyboard.press('Enter')
    await expect.poll(() => blocks(p).then((b) => b[0]!.type)).toBe('docTable')
    const rows = await p.evaluate(() => {
      const table = (window as unknown as DocsWindow).__aidocs!.editor.state.doc.child(0)
      const out: string[][] = []
      table.forEach((row: any) => {
        const cells: string[] = []
        row.forEach((cell: any) => cells.push(cell.textContent))
        out.push(cells)
      })
      return out
    })
    expect(rows).toEqual([
      ['Hi', 'Hi'],
      ['hi2', '2'],
    ])
  } finally {
    await closeAndSaveVideo(launched, 'md-table4')
  }
})

test('docs: Esc pauses the Markdown hint until the caret leaves the block', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'md-esc' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    const hint = p.locator('.md-shortcut-hint')

    // heading marker: Esc hides the hint and Enter then splits like normal text
    await p.keyboard.type('##')
    await expect(hint).toHaveCount(1)
    await p.keyboard.press('Escape')
    await expect(hint).toHaveCount(0)
    await p.keyboard.press('Enter')
    await expect
      .poll(() => blocks(p).then((b) => b.map((x) => `${x.type}:${x.text}`)))
      .toEqual(['docParagraph:##', 'docParagraph:'])
    // back into the paused block: the hint is armed again
    await p.keyboard.press('ArrowUp')
    await expect(hint).toHaveCount(1)

    // table rows: same
    await p.keyboard.press('ArrowDown')
    await p.keyboard.type('| a | b |')
    await p.keyboard.press('Enter')
    await p.keyboard.type('| -- | -- |')
    await expect(hint).toHaveCount(1)
    await p.keyboard.press('Escape')
    await expect(hint).toHaveCount(0)
    await p.keyboard.type(' ')
    await expect(hint).toHaveCount(0)
    await p.keyboard.press('Enter')
    expect((await blocks(p)).some((b) => b.type === 'docTable')).toBe(false)
  } finally {
    await closeAndSaveVideo(launched, 'md-esc')
  }
})

test('docs: a half-typed separator row ("| -- | |") still converts to a table', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'zh-TW', videoDir: 'md-table5' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    await p.keyboard.type('| Hi | Hi2|')
    await p.keyboard.press('Enter')
    await p.keyboard.type('| -- | |')
    await expect(p.locator('.md-shortcut-hint')).toHaveCount(1)
    await p.keyboard.press('Enter')
    await expect.poll(() => blocks(p).then((b) => b[0]!.type)).toBe('docTable')
  } finally {
    await closeAndSaveVideo(launched, 'md-table5')
  }
})
