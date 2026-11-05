import { expect, test, type Page } from '@playwright/test'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument, PDFName, PDFString, StandardFonts } from 'pdf-lib'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'

async function resize(launched: LaunchedApp, width: number) {
  await launched.app.evaluate(({ BrowserWindow }, width) => {
    BrowserWindow.getAllWindows()[0]!.setSize(width, 900)
  }, width)
}

async function verifyFloatingControls(launched: LaunchedApp, page: Page, count = 8) {
  await resize(launched, 1500)
  const trigger = page.locator('.floating-document-outline-trigger')
  const card = page.locator('.floating-document-outline-card')
  await expect(trigger).toBeVisible()
  await trigger.hover()
  await expect(card).toBeVisible()
  await expect(card.locator('.floating-document-outline-item')).toHaveCount(count)
  const search = card.getByRole('searchbox')
  await search.fill('Section 7')
  await expect(card.locator('.floating-document-outline-item')).toHaveCount(1)
  await search.fill('no such heading')
  await expect(card.getByRole('status')).toHaveText(/No results/i)
  await search.fill('')
  await page.mouse.move(50, 200)
  await expect(card).toHaveCount(0)
  await trigger.click()
  await page.mouse.move(50, 200)
  await expect(card).toBeVisible()
  await resize(launched, 800)
  await expect
    .poll(async () => {
      const box = await card.boundingBox()
      return !!box && box.x >= 0 && box.x + box.width <= 800 && box.y >= 0
    })
    .toBe(true)
  await page.screenshot({
    path: `/private/tmp/genoffice-floating-${page.url().match(/:\/\/(\w+)\//)?.[1]}-narrow.png`,
    animations: 'disabled',
  })
  await card.locator('header button').click()
  await expect(card).toHaveCount(0)
  await expect(trigger).toBeFocused()
  await page.keyboard.press('Enter')
  await expect(card).toBeVisible()
  await page.keyboard.press('Escape')
  await expect(card).toHaveCount(0)
  await resize(launched, 1500)
  await trigger.click()
  await page.screenshot({
    path: `/private/tmp/genoffice-floating-${page.url().match(/:\/\/(\w+)\//)?.[1]}-wide.png`,
    animations: 'disabled',
  })
  return card
}

test('Markdown floating outline navigates, follows edits and history, and survives save/reopen', async () => {
  test.setTimeout(120_000)
  const root = await mkdtemp(join(tmpdir(), 'genoffice-floating-md-'))
  const path = join(root, 'outline.md')
  await writeFile(
    path,
    Array.from(
      { length: 8 },
      (_, i) =>
        `${i === 0 ? '#' : '##'} Section ${i + 1}\n\n${'A paragraph with body text.\n\n'.repeat(12)}`,
    ).join('\n'),
  )
  let launched: LaunchedApp | undefined
  try {
    launched = await launchShell({
      onboardingSeen: true,
      openFile: path,
      videoDir: 'floating-markdown',
    })
    const page = await waitForPageWithUrl(launched.app, '://markdown/')
    const card = await verifyFloatingControls(launched, page)
    const caret = await page.evaluate(() => ({
      text: getSelection()?.anchorNode?.textContent,
      offset: getSelection()?.anchorOffset,
    }))
    await page.locator('.editor-scroll').evaluate((scroller) => {
      const target = [...scroller.querySelectorAll('h2')].find(
        (node) => node.textContent === 'Section 4',
      )!
      scroller.scrollTop +=
        target.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 24
    })
    await expect(
      card.locator('[aria-current="location"] .floating-document-outline-text'),
    ).toHaveText('Section 4')
    expect(
      await page.evaluate(() => ({
        text: getSelection()?.anchorNode?.textContent,
        offset: getSelection()?.anchorOffset,
      })),
    ).toEqual(caret)
    await page.locator('.editor-scroll').evaluate((scroller) => {
      scroller.scrollTop = 0
    })
    await expect(
      card.locator('[aria-current="location"] .floating-document-outline-text'),
    ).toHaveText('Section 1')
    await card.getByRole('button', { name: 'Section 7', exact: true }).click()
    await expect
      .poll(() => page.locator('.editor-scroll').evaluate((el) => el.scrollTop))
      .toBeGreaterThan(500)
    const heading = page.locator('.doc-editor h2', { hasText: 'Section 7' })
    await heading.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' edited')
    await expect(
      card.locator('.floating-document-outline-item', { hasText: 'Section 7 edited' }),
    ).toBeVisible()
    await page.keyboard.press('ControlOrMeta+z')
    await expect(card.getByRole('button', { name: 'Section 7', exact: true })).toBeVisible()
    await page.keyboard.press('ControlOrMeta+Shift+z')
    await expect(card.getByRole('button', { name: 'Section 7 edited', exact: true })).toBeVisible()
    await page.locator('.editor-scroll').evaluate((scroller) => {
      const target = [...scroller.querySelectorAll('h2')].find(
        (node) => node.textContent === 'Section 7 edited',
      )!
      scroller.scrollTop +=
        target.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 24
    })
    await expect(
      card.locator('[aria-current="location"] .floating-document-outline-text'),
    ).toHaveText('Section 7 edited')
    await page.keyboard.press('ControlOrMeta+s')
    await expect.poll(() => readFile(path, 'utf8')).toContain('Section 7 edited')
    await closeAndSaveVideo(launched, 'floating-markdown')
    launched = undefined
    launched = await launchShell({
      onboardingSeen: true,
      openFile: path,
      videoDir: 'floating-markdown-reopen',
    })
    const reopened = await waitForPageWithUrl(launched.app, '://markdown/')
    await reopened.locator('.floating-document-outline-trigger').click()
    await expect(
      reopened.locator('.floating-document-outline-item', { hasText: 'Section 7 edited' }),
    ).toBeVisible()
  } finally {
    if (launched) await closeAndSaveVideo(launched, 'floating-markdown')
    await rm(root, { recursive: true, force: true })
  }
})

async function makePdf(path: string, bookmarks: boolean, samePage = false) {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (let i = 0; i < 8; i++) {
    const page = samePage
      ? (doc.getPages()[0] ?? doc.addPage([595, 2400]))
      : doc.addPage([595, 842])
    const y = samePage ? 2300 - i * 280 : 760
    page.drawText(`Section ${i + 1}`, { x: 50, y, size: 22, font })
    for (let j = 0; j < 12; j++)
      page.drawText('Body paragraph with enough text to identify the normal font size.', {
        x: 50,
        y: y - 60 - j * 16,
        size: 11,
        font,
      })
  }
  if (bookmarks) {
    const root = doc.context.obj({ Type: 'Outlines', Count: 8 })
    const rootRef = doc.context.register(root)
    const entries = Array.from({ length: 8 }, (_, i) =>
      doc.context.obj({
        Title: PDFString.of(`Section ${i + 1}`),
        Parent: rootRef,
        Dest: samePage
          ? [doc.getPage(0).ref, PDFName.of('XYZ'), 50, 2300 - i * 280, null]
          : [doc.getPage(i).ref, PDFName.of('Fit')],
      }),
    )
    const refs = entries.map((entry) => doc.context.register(entry))
    entries.forEach((entry, i) => {
      if (i > 0) entry.set(PDFName.of('Prev'), refs[i - 1]!)
      if (i < refs.length - 1) entry.set(PDFName.of('Next'), refs[i + 1]!)
    })
    root.set(PDFName.of('First'), refs[0]!)
    root.set(PDFName.of('Last'), refs.at(-1)!)
    doc.catalog.set(PDFName.of('Outlines'), rootRef)
  }
  await writeFile(path, await doc.save())
}

for (const [bookmarks, samePage] of [
  [true, false],
  [false, false],
  [true, true],
] as const) {
  test(`PDF floating outline uses ${samePage ? 'same-page destinations' : bookmarks ? 'embedded bookmarks' : 'detected headings'} without changing the document`, async () => {
    test.setTimeout(90_000)
    const root = await mkdtemp(join(tmpdir(), 'genoffice-floating-pdf-'))
    const path = join(root, 'outline.pdf')
    await makePdf(path, bookmarks, samePage)
    const original = await readFile(path)
    const launched = await launchShell({
      onboardingSeen: true,
      openFile: path,
      videoDir: `floating-pdf-${bookmarks}`,
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '://pdf/')
      const card = await verifyFloatingControls(launched, page)
      if (!bookmarks) await expect(card).toContainText('Generated from headings')
      await page.locator('.pdf-scroll').evaluate(
        (scroller, { samePage, bookmarks }) => {
          const target = scroller.querySelector(samePage ? '.pdf-page' : '.pdf-row[data-idx="3"]')!
          scroller.scrollTop +=
            target.getBoundingClientRect().top -
            scroller.getBoundingClientRect().top +
            (samePage
              ? ((100 + 3 * 280) * target.getBoundingClientRect().height) / 2400 - 24
              : !bookmarks
                ? (82 * scroller.querySelector('.pdf-page')!.getBoundingClientRect().height) / 842 -
                  24
                : 0)
        },
        { samePage, bookmarks },
      )
      await expect(
        card.locator('[aria-current="location"] .floating-document-outline-text'),
      ).toHaveText('Section 4')
      await page.locator('.pdf-scroll').evaluate((scroller) => {
        scroller.scrollTop = 0
      })
      await expect(
        card.locator('[aria-current="location"] .floating-document-outline-text'),
      ).toHaveText('Section 1')
      await card.getByRole('button', { name: 'Section 7', exact: true }).click()
      await expect(page.locator('.status-left')).toContainText(
        samePage ? 'Page 1 of 1' : 'Page 7 of 8',
      )
      await expect(card.getByRole('button', { name: 'Section 7', exact: true })).toHaveAttribute(
        'aria-current',
        'location',
      )
      await expect(page.locator('.tb-save-pending')).toHaveCount(0)
      expect(await readFile(path)).toEqual(original)
    } finally {
      await closeAndSaveVideo(launched, `floating-pdf-${bookmarks}`)
      await rm(root, { recursive: true, force: true })
    }
  })
}

test('Docs uses the shared floating outline and follows scrolling through native TOC targets', async () => {
  test.setTimeout(120_000)
  const root = await mkdtemp(join(tmpdir(), 'genoffice-unified-docs-'))
  const path = join(root, 'outline.docx')
  const toc = Array.from(
    { length: 40 },
    (_, i) =>
      `<w:p><w:pPr><w:pStyle w:val="TOC${i === 0 ? 1 : 2}"/></w:pPr><w:hyperlink w:anchor="_Toc${i}"><w:r><w:t>Section ${i + 1}</w:t></w:r><w:r><w:tab/><w:t>${i + 1}</w:t></w:r></w:hyperlink></w:p>`,
  ).join('')
  const body = Array.from(
    { length: 40 },
    (_, i) =>
      `<w:p><w:pPr><w:pStyle w:val="Heading${i === 0 ? 1 : 2}"/></w:pPr><w:bookmarkStart w:id="${i}" w:name="_Toc${i}"/><w:r><w:t>Section ${i + 1}</w:t></w:r><w:bookmarkEnd w:id="${i}"/></w:p>${'<w:p><w:r><w:t>Body paragraph in this section.</w:t></w:r></w:p>'.repeat(2)}`,
  ).join('')
  const bytes = await buildDocx({ bodyXml: toc + body })
  await writeFile(path, bytes)
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: path,
    videoDir: 'unified-docs-outline',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '://docs/')
    await page.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
    const card = await verifyFloatingControls(launched, page, 40)
    await expect(card.locator('.floating-document-outline-page')).toHaveCount(40)
    const before = await page.evaluate(() => ({
      selection: (window as any).__aidocs.editor.state.selection.from,
      json: JSON.stringify((window as any).__aidocs.editor.state.doc.toJSON()),
    }))
    for (const section of [4, 7, 2, 40]) {
      await page.evaluate((section) => {
        const editor = (window as any).__aidocs.editor
        let target: HTMLElement | null = null
        editor.state.doc.forEach((node: any, pos: number) => {
          if (node.type.name === 'docHeading' && node.textContent === `Section ${section}`)
            target = editor.view.nodeDOM(pos)
        })
        const scroller = document.querySelector('.editor-scroll')!
        const element = target as HTMLElement | null
        if (!element) throw new Error('missing heading')
        scroller.scrollTop +=
          element.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 24
      }, section)
      await expect(
        card.locator('[aria-current="location"] .floating-document-outline-text'),
      ).toHaveText(`Section ${section}`)
    }
    await expect
      .poll(() =>
        card.locator('.floating-document-outline-list').evaluate((list) => {
          const current = list.querySelector('[aria-current="location"]')!
          const box = current.getBoundingClientRect()
          const viewport = list.getBoundingClientRect()
          return (
            list.scrollTop > 0 && box.top >= viewport.top - 1 && box.bottom <= viewport.bottom + 1
          )
        }),
      )
      .toBe(true)
    expect(
      await page.evaluate(() => ({
        selection: (window as any).__aidocs.editor.state.selection.from,
        json: JSON.stringify((window as any).__aidocs.editor.state.doc.toJSON()),
      })),
    ).toEqual(before)
    await card.getByTitle('Section 7', { exact: true }).click()
    await expect(
      card.locator('[aria-current="location"] .floating-document-outline-text'),
    ).toHaveText('Section 7')
    expect(await readFile(path)).toEqual(Buffer.from(bytes))
  } finally {
    await closeAndSaveVideo(launched, 'unified-docs-outline')
    await rm(root, { recursive: true, force: true })
  }
})

test('blank Markdown and PDF do not display an empty floating outline', async () => {
  test.setTimeout(90_000)
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'floating-empty' })
  try {
    await launched.page.locator('.quick-card', { hasText: 'AI Markdown' }).click()
    const md = await waitForPageWithUrl(launched.app, '://markdown/')
    await expect(md.locator('.doc-editor')).toBeVisible()
    await expect(md.locator('.floating-document-outline-host')).toHaveCount(0)
    await launched.page.locator('.tab-home').click()
    await launched.page.locator('.quick-card', { hasText: 'AI PDF' }).click()
    const pdf = await waitForPageWithUrl(launched.app, '://pdf/')
    await expect(pdf.locator('.pdf-page canvas').first()).toBeVisible()
    await expect(pdf.locator('.floating-document-outline-host')).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'floating-empty')
  }
})
