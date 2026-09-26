import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl } from './helpers'

/**
 * Paragraph dialog indents checked against Word for Mac (zh-TW):
 *  - Special = (none) / First line / Hanging with a "By" value; East Asian
 *    Word counts indents in characters (字元), 2 字元 being the default;
 *  - Left shows where the first line starts (a hanging paragraph's w:left
 *    minus the hang), and OK on untouched fields never rewrites them;
 *  - ⌘T steps a hanging indent by the document's default tab stop
 *    (zh-TW blank document: 480 twips = 24pt, the same as Word).
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

const paraIndents = (p: Page) =>
  p.evaluate(() => {
    const out: { first: number | null; left: number | null }[] = []
    ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((n: any) =>
      out.push({ first: n.attrs.indentFirstLine, left: n.attrs.indentLeft }),
    )
    return out
  })

async function openParagraphDialog(p: Page, text: string) {
  await p.getByText(text).click({ button: 'right' })
  await p.locator('.ctx-item', { hasText: '段落…' }).click()
  const dialog = p.locator('.modal', { has: p.getByRole('heading', { name: '段落' }) })
  await expect(dialog).toBeVisible()
  return dialog
}

test('paragraph dialog: hanging / first-line indent in characters, ⌘T uses the tab grid', async () => {
  const { mkdtempSync, existsSync } = await import('node:fs')
  const { execFileSync } = await import('node:child_process')
  const { tmpdir } = await import('node:os')
  const { join } = await import('node:path')
  const out = join(mkdtempSync(join(tmpdir(), 'docs-indent-')), 'indent.docx')
  const launched = await launchShell({ onboardingSeen: true, lang: 'zh-TW', videoDir: 'docs-indent' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(() => Boolean((window as unknown as DocsWindow).__aidocs?.editor), undefined, {
      timeout: 30_000,
    })
    await launched.app.evaluate(({ dialog }, file) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
    }, out)
    await p.locator('.doc-page').first().click()
    await p.keyboard.type('懸掛段落')
    await p.keyboard.press('Enter')
    await p.keyboard.type('首行段落')
    await p.keyboard.press('Enter')
    await p.keyboard.type('快速鍵段落')

    // one character = the paragraph's rendered font size, in twips
    const charTwips = await p.evaluate(() => {
      const view = (window as unknown as DocsWindow).__aidocs!.editor.view
      let node: Node | null = view.domAtPos(2).node
      if (node && node.nodeType !== Node.ELEMENT_NODE) node = node.parentElement
      return Math.round(parseFloat(getComputedStyle(node as Element).fontSize) * 0.75 * 20)
    })

    // Hanging: By defaults to 2 字元; Left stays 0 (first line at the margin)
    let dialog = await openParagraphDialog(p, '懸掛段落')
    await dialog.getByRole('button', { name: '特殊' }).click()
    await p.getByRole('option', { name: '凸排' }).click()
    await expect(dialog.getByText('位移點數')).toBeVisible()
    await expect(dialog.locator('label', { hasText: '位移點數' }).locator('input')).toHaveValue('2')
    await expect(dialog.locator('label', { hasText: '位移點數' })).toContainText('字元')
    await dialog.getByRole('button', { name: '確定' }).click()
    await expect.poll(() => paraIndents(p).then((r) => r[0])).toEqual({
      first: -2 * charTwips,
      left: 2 * charTwips,
    })

    // reopened: Left 0, Hanging 2 字元; OK without edits keeps the exact twips
    dialog = await openParagraphDialog(p, '懸掛段落')
    await expect(dialog.locator('label', { hasText: '左縮排' }).locator('input')).toHaveValue('0')
    await expect(dialog.getByRole('button', { name: '特殊' })).toContainText('凸排')
    await dialog.getByRole('button', { name: '確定' }).click()
    await expect.poll(() => paraIndents(p).then((r) => r[0])).toEqual({
      first: -2 * charTwips,
      left: 2 * charTwips,
    })

    // First line: 2 字元, body lines at the margin
    dialog = await openParagraphDialog(p, '首行段落')
    await dialog.getByRole('button', { name: '特殊' }).click()
    await p.getByRole('option', { name: '首行' }).click()
    await dialog.getByRole('button', { name: '確定' }).click()
    await expect.poll(() => paraIndents(p).then((r) => r[1])).toEqual({ first: 2 * charTwips, left: null })

    // ⌘T: Word (zh-TW) hangs by one default tab stop — 480 twips
    await p.getByText('快速鍵段落').click()
    // let the editor read the click's caret before the shortcut
    await expect
      .poll(() =>
        p.evaluate(
          () =>
            (window as unknown as DocsWindow).__aidocs!.editor.state.selection.$from.parent
              .textContent,
        ),
      )
      .toBe('快速鍵段落')
    await p.keyboard.press('ControlOrMeta+t')
    await expect.poll(() => paraIndents(p).then((r) => r[2])).toEqual({ first: -480, left: 480 })

    await p.keyboard.press('ControlOrMeta+Shift+s')
    await expect.poll(() => existsSync(out), { timeout: 20_000 }).toBe(true)
    await expect
      .poll(() =>
        execFileSync('unzip', ['-p', out, 'word/document.xml'])
          .toString()
          .match(/<w:ind [^>]*>/g),
      )
      .toEqual([
        `<w:ind w:left="${2 * charTwips}" w:hanging="${2 * charTwips}"/>`,
        `<w:ind w:firstLine="${2 * charTwips}"/>`,
        '<w:ind w:left="480" w:hanging="480"/>',
      ])
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
