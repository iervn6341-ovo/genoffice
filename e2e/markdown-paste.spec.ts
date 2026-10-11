import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/**
 * Pasted text that looks like Markdown is inserted as typed and a "Markdown detected" chip offers
 * the conversion — for plain text and for the colored <div>/<span> HTML editors put on the clipboard.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

const SAMPLE = '| Hi | Hi2|\n| -- | -- |\n| 21 | 21 |\n\n\n\n| Hi | Hi2|\n| 21 | 21 |'
const EDITOR_HTML =
  '<meta charset="utf-8"><div style="color:#d4d4d4;background-color:#1e1e1e;white-space:pre"><div><span>| Hi | Hi2|</span></div><div><span>| -- | -- |</span></div><div><span>| 21 | 21 |</span></div></div>'

const paste = (p: Page, text: string, html = '') =>
  p.evaluate(
    ({ text, html }) => {
      const ed = (window as unknown as DocsWindow).__aidocs!.editor
      ed.commands.focus()
      const dt = new DataTransfer()
      dt.setData('text/plain', text)
      if (html) dt.setData('text/html', html)
      ed.view.dom.dispatchEvent(
        new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }),
      )
    },
    { text, html },
  )

const clear = (p: Page) =>
  p.evaluate(() => {
    const ed = (window as unknown as DocsWindow).__aidocs!.editor
    ed.commands.clearContent()
    ed.commands.focus()
  })

const types = (p: Page) =>
  p.evaluate(() => {
    const out: string[] = []
    ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((n: any) =>
      out.push(n.type.name),
    )
    return out
  })

test('docs: pasted Markdown offers a conversion chip', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'md-paste' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    const chip = p.locator('[data-testid="markdown-paste-chip"]')

    // plain text: literal first, chip, Convert → two tables (the header-less one is healed)
    await paste(p, SAMPLE)
    await expect(chip).toHaveCount(1)
    expect((await types(p)).includes('docTable')).toBe(false)
    await chip.getByRole('button', { name: 'Convert to formatted text' }).click()
    await expect.poll(() => types(p).then((t) => t.filter((x) => x === 'docTable').length)).toBe(2)
    await expect(chip).toHaveCount(0)
    // converted tables are visible (default borders) and kept apart by a paragraph
    expect(await types(p)).toEqual(['docTable', 'docParagraph', 'docTable', 'docParagraph'])
    expect(
      await p.evaluate(() => {
        let missing = 0
        ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((n: any) => {
          if (n.type.name === 'docTable' && !n.attrs.borders) missing++
        })
        return missing
      }),
    ).toBe(0)

    // editor HTML (colored div/span soup + the same text/plain): same offer
    await clear(p)
    await paste(p, SAMPLE.split('\n\n')[0]!, EDITOR_HTML)
    await expect(chip).toHaveCount(1)
    await chip.getByRole('button', { name: 'Convert to formatted text' }).click()
    await expect.poll(() => types(p).then((t) => t.filter((x) => x === 'docTable').length)).toBe(1)

    // typing something else dismisses it
    await clear(p)
    await paste(p, '## Title\n\nSome **bold** text.')
    await expect(chip).toHaveCount(1)
    await p.keyboard.type('x')
    await expect(chip).toHaveCount(0)

    // the close button dismisses it and keeps the literal text
    await clear(p)
    await paste(p, '- one\n- two')
    await expect(chip).toHaveCount(1)
    await chip.getByRole('button', { name: 'Dismiss' }).click()
    await expect(chip).toHaveCount(0)
    expect((await types(p)).includes('docListItem')).toBe(false)

    // rich HTML keeps its own structure and never shows the chip
    await clear(p)
    await paste(p, '## Title', '<h2>Title</h2><p>body</p>')
    await expect.poll(() => types(p).then((t) => t.includes('docHeading'))).toBe(true)
    await expect(chip).toHaveCount(0)

    // ordinary prose is left alone
    await clear(p)
    await paste(p, 'Meeting notes from Monday.\nBudget is $100.')
    await p.waitForTimeout(300)
    await expect(chip).toHaveCount(0)

    // a big paste stays responsive: literal paste and conversion each well under a few seconds
    await clear(p)
    const big = Array.from(
      { length: 1500 },
      (_, i) =>
        `## Section ${i}\n\n- item **${i}**\n- another\n\n| a | b |\n| - | - |\n| ${i} | x |\n`,
    ).join('\n')
    const t0 = Date.now()
    await paste(p, big)
    await expect(chip).toHaveCount(1, { timeout: 15_000 })
    const pasteMs = Date.now() - t0
    const t1 = Date.now()
    await chip.getByRole('button', { name: 'Convert to formatted text' }).click()
    await expect
      .poll(() => types(p).then((t) => t.includes('docHeading')), { timeout: 30_000 })
      .toBe(true)
    const convertMs = Date.now() - t1
    console.log(`BIG paste ${pasteMs}ms convert ${convertMs}ms`)
    expect(pasteMs).toBeLessThan(8000)
    expect(convertMs).toBeLessThan(15000)
  } finally {
    await closeAndSaveVideo(launched, 'md-paste')
  }
})

test('docs: converting never loses the text around the caret, and foreign tables get borders', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'en', videoDir: 'md-paste-ctx' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    const chip = p.locator('[data-testid="markdown-paste-chip"]')
    const table = '| Hi | Hi2|\n| -- | -- |\n| 21 | 21 |'

    const outline = (p2: Page) =>
      p2.evaluate(() => {
        const out: string[] = []
        ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((n: any) =>
          out.push(n.type.name === 'docTable' ? 'TABLE' : n.textContent),
        )
        return out
      })
    const place = (setup: string, where: 'start' | 'mid' | 'end') =>
      p.evaluate(
        ({ setup, where }) => {
          const ed = (window as unknown as DocsWindow).__aidocs!.editor
          ed.commands.setContent(setup)
          ed.commands.focus()
          const size = ed.state.doc.content.size
          ed.commands.setTextSelection(where === 'start' ? 1 : where === 'mid' ? 4 : size - 1)
        },
        { setup, where },
      )

    for (const [where, expected] of [
      ['end', ['abcdef', 'TABLE', '']],
      ['mid', ['abc', 'TABLE', 'def']],
      ['start', ['TABLE', 'abcdef']],
    ] as const) {
      await place('<p>abcdef</p>', where)
      await paste(p, table)
      await expect(chip).toHaveCount(1)
      await chip.getByRole('button', { name: 'Convert to formatted text' }).click()
      await expect.poll(() => outline(p)).toEqual(expected)
    }

    // right under an existing table: a paragraph keeps the two tables apart
    await place('<table><tbody><tr><td>x</td></tr></tbody></table><p></p>', 'end')
    await paste(p, table)
    await expect(chip).toHaveCount(1)
    await chip.getByRole('button', { name: 'Convert to formatted text' }).click()
    await expect.poll(() => outline(p)).toEqual(['TABLE', '', 'TABLE', ''])

    // a table copied from a web page / Excel carries no borders: it gets the default ones
    await clear(p)
    await paste(
      p,
      'a\tb\n1\t2',
      '<table><tbody><tr><td>a</td><td>b</td></tr><tr><td>1</td><td>2</td></tr></tbody></table>',
    )
    await expect.poll(() => types(p).then((t) => t.includes('docTable'))).toBe(true)
    expect(
      await p.evaluate(() => {
        let bordered = 0
        ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((n: any) => {
          if (n.type.name === 'docTable' && n.attrs.borders?.insideV) bordered++
        })
        return bordered
      }),
    ).toBe(1)
  } finally {
    await closeAndSaveVideo(launched, 'md-paste-ctx')
  }
})
