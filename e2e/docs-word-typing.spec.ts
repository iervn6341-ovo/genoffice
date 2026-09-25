import { test, expect, type Page } from '@playwright/test'
import { launchShell, waitForPageWithUrl } from './helpers'

/**
 * Typing conventions checked against Word for Mac:
 *  - Home / End move to the start / end of the line (Chromium on macOS only scrolls),
 *    ⇧ extends the selection, ⌘Home / ⌘End go to the start / end of the document;
 *  - AutoFormat As You Type: "* ", "- ", "1. ", "a) " at a paragraph start begin a list,
 *    Enter on an empty item leaves it, and ⌘Z right after brings the typed marker back.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

const blocks = (p: Page) =>
  p.evaluate(() => {
    const ed = (window as unknown as DocsWindow).__aidocs!.editor
    const out: { type: string; kind?: string; numId?: string; text: string }[] = []
    ed.state.doc.forEach((n: any) =>
      out.push({
        type: n.type.name,
        ...(n.type.name === 'docListItem'
          ? { kind: n.attrs.kind, numId: String(n.attrs.numId) }
          : {}),
        text: n.textContent,
      }),
    )
    return out
  })

const selectedText = (p: Page) =>
  p.evaluate(() => {
    const ed = (window as unknown as DocsWindow).__aidocs!.editor
    const { from, to } = ed.state.selection
    return ed.state.doc.textBetween(from, to) as string
  })

async function blankDoc() {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-word-typing' })
  await launched.page.locator('.quick-card').first().click()
  const p = await waitForPageWithUrl(launched.app, '://docs/')
  await p.waitForFunction(
    () => Boolean((window as unknown as DocsWindow).__aidocs?.editor),
    undefined,
    {
      timeout: 30_000,
    },
  )
  await p.locator('.doc-page').first().click()
  return { p, launched }
}

test('Home / End move by line, ⇧ extends, ⌘ goes to the document ends', async () => {
  const { p, launched } = await blankDoc()
  try {
    await p.keyboard.type('Hello world')
    await p.keyboard.press('Home')
    await p.keyboard.type('X')
    await p.keyboard.press('End')
    await p.keyboard.type('Y')
    await expect.poll(() => blocks(p)).toEqual([{ type: 'docParagraph', text: 'XHello worldY' }])
    await p.keyboard.press('Home')
    await p.keyboard.press('Shift+End')
    await expect.poll(() => selectedText(p)).toBe('XHello worldY')
    await p.keyboard.press('End')
    // let the editor read the collapsed selection before Enter (a synthetic
    // key burst can outrun the selectionchange; a person cannot)
    await expect.poll(() => selectedText(p)).toBe('')
    await p.keyboard.press('Enter')
    await p.keyboard.type('second line')
    await p.keyboard.press('ControlOrMeta+Home')
    await p.keyboard.type('S')
    await p.keyboard.press('ControlOrMeta+End')
    await p.keyboard.type('E')
    await expect
      .poll(() => blocks(p).then((b) => b.map((x) => x.text)))
      .toEqual(['SXHello worldY', 'second lineE'])
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('list AutoFormat: markers start lists, Enter on an empty item ends one, ⌘Z restores the marker', async () => {
  const { p, launched } = await blankDoc()
  try {
    const lines: [string, string][] = [
      ['* item one', 'item two'],
      ['1. first', 'second'],
      ['- dash', ''],
      ['a) alpha', ''],
    ]
    for (const [first, next] of lines) {
      await p.keyboard.type(first)
      await p.keyboard.press('Enter')
      if (next) {
        await p.keyboard.type(next)
        await p.keyboard.press('Enter')
      }
      await p.keyboard.press('Enter') // empty item: back to a plain paragraph
    }
    await p.keyboard.type('plain')
    const b = await blocks(p)
    expect(b.map((x) => [x.type, x.kind ?? '', x.text])).toEqual([
      ['docListItem', 'bullet', 'item one'],
      ['docListItem', 'bullet', 'item two'],
      ['docListItem', 'ordered', 'first'],
      ['docListItem', 'ordered', 'second'],
      ['docListItem', 'bullet', 'dash'],
      ['docListItem', 'ordered', 'alpha'],
      ['docParagraph', '', 'plain'],
    ])
    // each marker family is its own list: "- " is not the "* " list, "a) " restarts at a)
    expect(new Set([b[0]!.numId, b[2]!.numId, b[4]!.numId, b[5]!.numId]).size).toBe(4)
    const level0 = (numId: string) =>
      p.evaluate((id) => {
        const def = (
          window as unknown as DocsWindow
        ).__aidocs!.editor.storage.listNumbering.defs.get(id)
        return def ? { numFmt: def.levels[0].numFmt, lvlText: def.levels[0].lvlText } : null
      }, numId)
    expect(await level0(b[2]!.numId!)).toEqual({ numFmt: 'decimal', lvlText: '%1.' })
    expect(await level0(b[4]!.numId!)).toEqual({ numFmt: 'bullet', lvlText: '–' })
    expect(await level0(b[5]!.numId!)).toEqual({ numFmt: 'lowerLetter', lvlText: '%1)' })

    await p.keyboard.press('Enter')
    await p.keyboard.type('* ')
    await expect.poll(() => blocks(p).then((x) => x.at(-1)?.type)).toBe('docListItem')
    await p.keyboard.press('ControlOrMeta+z')
    await expect
      .poll(() => blocks(p).then((x) => x.at(-1)))
      .toEqual({ type: 'docParagraph', text: '* ' })
    // mid-paragraph and in headings the marker stays text
    await p.keyboard.type('stays')
    await p.keyboard.press('Enter')
    await p.keyboard.type('x * y')
    expect((await blocks(p)).at(-1)).toEqual({ type: 'docParagraph', text: 'x * y' })
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('Clear Formatting at the caret types plain; a reopened heading restyled saves its new style', async () => {
  const { mkdtempSync, existsSync } = await import('node:fs')
  const { execFileSync } = await import('node:child_process')
  const { tmpdir } = await import('node:os')
  const { join } = await import('node:path')
  const out = join(mkdtempSync(join(tmpdir(), 'docs-typing-')), 'restyle.docx')
  const pStyles = () =>
    execFileSync('unzip', ['-p', out, 'word/document.xml'])
      .toString()
      .match(/<w:pStyle w:val="[^"]*"/g)
  const { p, launched } = await blankDoc()
  try {
    await launched.app.evaluate(({ dialog }, file) => {
      dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
    }, out)
    await p.keyboard.press('ControlOrMeta+b')
    await p.keyboard.type('bold ')
    await p.locator('.ribbon-body button[aria-label^="Clear All Formatting"]').first().click()
    await p.keyboard.type('plain')
    const runs = await p.evaluate(() =>
      (window as unknown as DocsWindow)
        .__aidocs!.editor.getJSON()
        .content[0].content.map((r: any) => [r.text, (r.marks ?? []).map((m: any) => m.type)]),
    )
    expect(runs).toEqual([
      ['bold ', ['bold']],
      ['plain', []],
    ])
    await p.keyboard.press('Enter')
    await p.keyboard.type('Heading text')
    await p.locator('.style-card', { hasText: 'Heading 2' }).click()
    await p.keyboard.press('ControlOrMeta+Shift+s')
    await expect.poll(() => existsSync(out), { timeout: 20_000 }).toBe(true)
    await expect.poll(pStyles).toEqual(['<w:pStyle w:val="Heading2"'])
  } finally {
    launched.app.process().kill('SIGKILL')
  }
  // reopened, the heading carries styleId Heading2 — restyling must replace it
  const re = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-word-typing',
    openFile: out,
  })
  try {
    const p2 = await waitForPageWithUrl(re.app, '://docs/')
    await p2.waitForFunction(() => Boolean((window as unknown as DocsWindow).__aidocs?.editor))
    await p2.getByText('Heading text').click()
    await p2.locator('.style-card', { hasText: 'Heading 3' }).click()
    let before = execFileSync('unzip', ['-p', out, 'word/document.xml']).toString()
    await p2.keyboard.press('ControlOrMeta+s')
    await expect.poll(pStyles, { timeout: 15_000 }).toEqual(['<w:pStyle w:val="Heading3"'])
    await p2.getByText('Heading text').click()
    await p2.locator('.style-card', { hasText: 'Normal' }).click()
    before = execFileSync('unzip', ['-p', out, 'word/document.xml']).toString()
    await p2.keyboard.press('ControlOrMeta+s')
    await expect
      .poll(() => execFileSync('unzip', ['-p', out, 'word/document.xml']).toString() !== before, {
        timeout: 15_000,
      })
      .toBe(true)
    expect(pStyles()).toBeNull()
  } finally {
    re.app.process().kill('SIGKILL')
  }
})

test('a double-click right of a line never swallows the paragraph break', async () => {
  const { p, launched } = await blankDoc()
  try {
    await p.keyboard.type('Second')
    await p.keyboard.press('Enter')
    await p.keyboard.type('Third line')
    const line = (await p.locator('.doc-page p').first().boundingBox())!
    // the empty space well right of "Second" on its own line
    await p.mouse.dblclick(line.x + line.width - 40, line.y + line.height / 2)
    await p.keyboard.type('X')
    await expect
      .poll(() => blocks(p).then((b) => b.map((x) => x.text)))
      .toEqual(['SecondX', 'Third line'])
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
