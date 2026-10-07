import { test, expect } from '@playwright/test'
import type { Page } from '@playwright/test'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

/**
 * Word (probed 2026-10-06): Backspace at the very start of a list item —
 * empty or not — takes its bullet off and leaves the text where it was; only
 * later Backspaces walk the indent back and finally join the previous line.
 * The same holds for a PowerPoint text box bullet and a Markdown list item.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

const docsBlocks = (p: Page) =>
  p.evaluate(() => {
    const out: Array<{ type: string; text: string; left: number | null }> = []
    ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((n: any) =>
      out.push({ type: n.type.name, text: n.textContent, left: n.attrs.indentLeft ?? null }),
    )
    return out
  })

test('docs: Backspace at a list item start removes the bullet before joining', async () => {
  const launched = await launchShell({ onboardingSeen: true, lang: 'zh-TW', videoDir: 'list-bs' })
  try {
    await launched.page.locator('.quick-card').first().click()
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().click()
    await p.locator('[aria-label="項目符號"]').first().click()
    await p.keyboard.type('asdasd')
    await p.keyboard.press('Enter')
    await p.keyboard.type('asdasd')
    await p.keyboard.press('Enter')
    await expect
      .poll(() => docsBlocks(p).then((b) => b.map((x) => x.type)))
      .toEqual(['docListItem', 'docListItem', 'docListItem'])

    // 1st: the empty item becomes a plain paragraph at the list's text indent
    await p.keyboard.press('Backspace')
    await expect
      .poll(() => docsBlocks(p).then((b) => b[2]))
      .toMatchObject({ type: 'docParagraph', text: '', left: 720 })
    // its own undo step: the bullet comes back
    await p.keyboard.press('ControlOrMeta+z')
    await expect.poll(() => docsBlocks(p).then((b) => b[2]?.type)).toBe('docListItem')
    await p.keyboard.press('Backspace')
    await expect.poll(() => docsBlocks(p).then((b) => b[2]?.type)).toBe('docParagraph')
    // then back to where the bullet stood (720 − 360 hanging), then the margin
    await p.keyboard.press('Backspace')
    await expect.poll(() => docsBlocks(p).then((b) => b[2]?.left)).toBe(360)
    await p.keyboard.press('Backspace')
    await expect.poll(() => docsBlocks(p).then((b) => b[2]?.left)).toBeNull()
    // only the next press joins the line into the previous one
    await p.keyboard.press('Backspace')
    await expect.poll(() => docsBlocks(p).then((b) => b.length)).toBe(2)

    // a non-empty item: the caret at its start, Backspace keeps the text on its own line
    await p.evaluate(() => {
      const ed = (window as unknown as DocsWindow).__aidocs!.editor
      let pos = -1
      ed.state.doc.forEach((n: any, off: number) => {
        if (n.type.name === 'docListItem') pos = off + 1
      })
      ed.commands.setTextSelection(pos)
      ed.commands.focus()
    })
    await p.keyboard.press('Backspace')
    await expect
      .poll(() => docsBlocks(p))
      .toEqual([
        { type: 'docListItem', text: 'asdasd', left: null },
        { type: 'docParagraph', text: 'asdasd', left: 720 },
      ])
  } finally {
    await closeAndSaveVideo(launched, 'cleanup')
  }
})

test('docs: a ListParagraph-styled item does not snap back to the style indent', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'zh-TW',
    openFile: join(__dirname, 'assets/list-paragraph-style-bullets.docx'),
    videoDir: 'list-bs-styled',
  })
  try {
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().waitFor()
    await expect.poll(() => docsBlocks(p).then((b) => b.length)).toBe(5)
    // caret into the empty 4th paragraph (a bullet whose style indents by 720)
    await p.evaluate(() => {
      const ed = (window as unknown as DocsWindow).__aidocs!.editor
      let pos = -1
      let i = 0
      ed.state.doc.forEach((_n: any, off: number) => {
        if (i++ === 3) pos = off + 1
      })
      ed.commands.setTextSelection(pos)
      ed.commands.focus()
    })
    await expect
      .poll(() =>
        p.evaluate(() => {
          const sel = (window as unknown as DocsWindow).__aidocs!.editor.state.selection
          return `${sel.$from.parent.type.name}:${sel.$from.parent.content.size}`
        }),
      )
      .toBe('docListItem:0')
    // caret distance from the left margin in twips, from the rendered position
    // (what the user sees: the style's indent must not pull the line back)
    const caretTwips = () =>
      p.evaluate(() => {
        const ed = (window as unknown as DocsWindow).__aidocs!.editor
        const page = document.querySelector('.doc-page')!.getBoundingClientRect()
        const marginPx = parseFloat(
          getComputedStyle(document.querySelector('.doc-page')!).paddingLeft,
        )
        return Math.round(
          ((ed.view.coordsAtPos(ed.state.selection.from).left - page.left - marginPx) / 96) * 1440,
        )
      })
    // Word, same file: text indent 720 → bullet position 360 → margin → join
    await p.keyboard.press('Backspace')
    await expect.poll(() => docsBlocks(p).then((b) => b[3]?.type)).toBe('docParagraph')
    await expect.poll(caretTwips).toBeCloseTo(720, -1)
    await p.keyboard.press('Backspace')
    await expect.poll(caretTwips).toBeCloseTo(360, -1)
    await p.keyboard.press('Backspace')
    await expect.poll(caretTwips).toBeCloseTo(0, -1)
    await p.keyboard.press('Backspace')
    await expect.poll(() => docsBlocks(p).then((b) => b.length)).toBe(4)
    expect((await docsBlocks(p))[3]!.text).toBe('tail')
  } finally {
    await closeAndSaveVideo(launched, 'cleanup')
  }
})

test('docs: Backspace walks to the bullet position, then the margin (Word-probed lists)', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'zh-TW',
    openFile: join(__dirname, 'assets/list-backspace-matrix.docx'),
    videoDir: 'list-bs-matrix',
  })
  try {
    const p = await waitForPageWithUrl(launched.app, '://docs/')
    await p.waitForFunction(
      () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
      undefined,
      { timeout: 30_000 },
    )
    await p.locator('.doc-page').first().waitFor()
    await expect.poll(() => docsBlocks(p).then((b) => b.length)).toBe(15)
    // caret into the empty item that follows the "<label>x" item
    const enter = async (label: string) => {
      await p.evaluate((lab) => {
        const ed = (window as unknown as DocsWindow).__aidocs!.editor
        let pos = -1
        let prev: any = null
        ed.state.doc.forEach((n: any, off: number) => {
          if (prev?.textContent === `${lab}x` && n.content.size === 0) pos = off + 1
          prev = n
        })
        ed.commands.setTextSelection(pos)
        ed.commands.focus()
      }, label)
      // let the editor read the selection before the key (synthetic keys outrun selectionchange)
      await expect
        .poll(() =>
          p.evaluate(() => {
            const sel = (window as unknown as DocsWindow).__aidocs!.editor.state.selection
            return `${sel.$from.parent.type.name}:${sel.$from.parent.content.size}`
          }),
        )
        .toBe('docListItem:0')
      await p.waitForTimeout(300)
    }
    const empties = (label: string) =>
      docsBlocks(p).then((b) => {
        const i = b.findIndex((x) => x.text === `${label}x`)
        return b.slice(i + 1).find((x) => x.text === '' || x.type !== 'docListItem') ?? null
      })

    // B: bullet level left 1440 / hanging 360 under a style indent of 720 —
    // Word stops at the bullet position 1080, then goes to the margin
    await enter('B')
    await p.keyboard.press('Backspace')
    await expect.poll(() => empties('B').then((x) => x?.left)).toBe(1440)
    await p.keyboard.press('Backspace')
    await expect.poll(() => empties('B').then((x) => x?.left)).toBe(1080)
    await p.keyboard.press('Backspace')
    await expect.poll(() => empties('B').then((x) => x?.left)).toBe(0)

    // D: the bullet already stood on the margin (left 720 / hanging 720) —
    // the second press is already at the margin, the third joins
    await enter('D')
    await p.keyboard.press('Backspace')
    await expect.poll(() => empties('D').then((x) => x?.left)).toBe(720)
    await p.keyboard.press('Backspace')
    await expect.poll(() => empties('D').then((x) => x?.left ?? 0)).toBe(0)
    const before = (await docsBlocks(p)).length
    await p.keyboard.press('Backspace')
    await expect.poll(() => docsBlocks(p).then((b) => b.length)).toBe(before - 1)
  } finally {
    await closeAndSaveVideo(launched, 'cleanup')
  }
})

test('markdown: Backspace at a list item start lifts it out of the list', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'list-bs-md' })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI Markdown' }).click()
    const p = await waitForPageWithUrl(launched.app, '://markdown/')
    const ed = p.locator('.doc-editor').first()
    await ed.click()
    await p.keyboard.type('- asdasd')
    await p.keyboard.press('Enter')
    await p.keyboard.type('asdasd')
    await p.keyboard.press('Enter')
    await expect(ed.locator('li')).toHaveCount(3)
    await p.keyboard.press('Backspace')
    await expect(ed.locator('li')).toHaveCount(2)
    await expect(ed.locator(':scope > p')).not.toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'cleanup')
  }
})

const editorParas = (s: Page) =>
  s.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLElement>('.slide-text-editor > div')).map((d) => ({
      text: d.textContent,
      bullet: d.dataset.bullet ?? d.dataset.hadBullet ?? null,
    })),
  )

test('slides: Backspace at a bulleted paragraph start removes the bullet first', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'list-bs-slides' })
  try {
    await launched.page.locator('.quick-card').nth(2).click()
    const s = await waitForPageWithUrl(launched.app, '://slides/', 20_000)
    await s.waitForSelector('.ribbon', { timeout: 15_000 })
    await setEditorLayoutWidth(launched.app, '://slides/', 2200)
    await s.locator('.ribbon-tab', { hasText: 'Insert' }).click()
    await s.locator('.ribbon-body button[data-tip^="Insert a text box"]').click()
    await s.waitForSelector('.slide-text-editor')
    await s.locator('.ribbon-tab', { hasText: 'Home' }).click()
    await s.keyboard.type('asdasd')
    await s.locator('button[aria-label="Bullets"]').first().click()
    await s.keyboard.press('Enter')
    await s.keyboard.type('asdasd')
    await s.keyboard.press('Enter')
    await expect
      .poll(() => editorParas(s))
      .toEqual([
        { text: 'asdasd', bullet: 'char' },
        { text: 'asdasd', bullet: 'char' },
        { text: '', bullet: 'char' },
      ])
    await s.keyboard.press('Backspace')
    await expect.poll(() => editorParas(s).then((x) => x[2])).toEqual({ text: '', bullet: 'none' })
    await s.keyboard.press('Backspace')
    await expect.poll(() => editorParas(s).then((x) => x.length)).toBe(2)
    for (let i = 0; i < 6; i++) await s.keyboard.press('ArrowLeft')
    await s.keyboard.press('Backspace')
    await expect
      .poll(() => editorParas(s))
      .toEqual([
        { text: 'asdasd', bullet: 'char' },
        { text: 'asdasd', bullet: 'none' },
      ])
    // committed: the second paragraph carries no bullet
    await s.keyboard.press('Escape')
    await expect
      .poll(() =>
        s.evaluate(async () => {
          const r = await (window as any).slidesApi.getRenderSlides()
          const n = r[0].nodes.find((x: any) => x.text)
          return n.text.lines.map((l: any) => l.runs.some((x: any) => x.isBullet))
        }),
      )
      .toEqual([true, false])
  } finally {
    await closeAndSaveVideo(launched, 'cleanup')
  }
})
