import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test, type Page } from '@playwright/test'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Word features the Docs ribbon was missing (Layout → Columns / Breaks / Line Numbers /
 * Hyphenation / Margins / Paper Size, Home → borders, spacing, styles, paste options, Design →
 * page borders, paragraph spacing). Each test drives the real menus, then reads the saved
 * document.xml / settings.xml so the check covers the file Word would open.
 */
async function openDocs(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-ribbon-gaps' })
  await launched.page.locator('.quick-card').first().click()
  const p = await waitForPageWithUrl(launched.app, '://docs/', 30_000)
  await p.waitForSelector('.doc-page', { timeout: 20_000 })
  await p.waitForFunction(() => Boolean((window as any).__aidocs?.editor))
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')

async function saveTo(l: LaunchedApp, p: Page, out: string): Promise<void> {
  await l.app.evaluate(({ dialog }, file) => {
    dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
  }, out)
  await p.keyboard.press('Meta+Shift+s')
  await expect.poll(() => existsSync(out), { timeout: 20_000 }).toBe(true)
}
const part = (file: string, name: string): string =>
  execFileSync('unzip', ['-p', file, name]).toString()

const tab = (p: Page, name: string) =>
  p.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).click()
const ribbonBtn = (p: Page, name: RegExp | string) =>
  p.locator('.ribbon-body button:not(.rb-fold-btn)', { hasText: name }).first()
const styleCard = (p: Page, label: string) =>
  p
    .locator('.style-gallery-menu .style-card')
    .filter({ has: p.locator('.style-card-label', { hasText: new RegExp(`^${label}$`) }) })
const menuItem = (p: Page, name: RegExp | string) =>
  p.locator('[data-rb-panel] button', { hasText: name }).first()

test.describe('Layout tab: the Page Setup dropdowns Word has', () => {
  test('columns, breaks, line numbers, hyphenation, mirrored margins and paper size reach the file', async () => {
    test.setTimeout(180_000)
    const { launched, p } = await openDocs()
    const out = join(mkdtempSync(join(tmpdir(), 'docs-gaps-')), 'layout.docx')
    try {
      await p.locator('.doc-page').click()
      await p.keyboard.type('First line of text', { delay: 5 })
      await p.keyboard.press('Enter')
      await p.keyboard.type('Second line of text', { delay: 5 })
      await tab(p, 'Layout')

      // Columns ▾ → Left (a narrow first column), then More Columns… (3 equal columns + rule)
      await ribbonBtn(p, /^Columns$/).click()
      for (const name of ['One', 'Two', 'Three', 'Left', 'Right', 'More Columns']) {
        await expect(menuItem(p, new RegExp(`^${name}`))).toBeVisible()
      }
      await menuItem(p, /^Left$/).click()
      await ribbonBtn(p, /^Columns$/).click()
      await expect(menuItem(p, /^Left$/)).toHaveClass(/active/)
      await menuItem(p, /^More Columns/).click()
      const dialog = p.locator('.cols-dialog')
      await expect(dialog).toBeVisible()
      await dialog.getByRole('button', { name: 'Three' }).click()
      await dialog.getByLabel('Line between').check()
      await dialog.getByRole('button', { name: /^OK$/ }).click()
      await expect(dialog).toBeHidden()

      // Breaks ▾ → Page, Column, Text Wrapping
      await p.keyboard.press('End')
      for (const name of ['Page', 'Column', 'Text Wrapping']) {
        await ribbonBtn(p, /^Breaks$/).click()
        await menuItem(p, new RegExp(`^${name}`)).click()
      }

      // Line Numbers ▾ → Restart Each Section; Hyphenation ▾ → Automatic
      await ribbonBtn(p, /Line Numbers$/).click()
      await menuItem(p, /^Restart Each Section$/).click()
      await ribbonBtn(p, /Hyphenation$/).click()
      await menuItem(p, /^Automatic$/).click()
      await expect(p.locator('.doc-page')).toHaveAttribute('data-hyphenate', 'on')

      // Margins ▾ → Mirrored, Size ▾ → More Paper Sizes… (20 × 30 cm)
      await ribbonBtn(p, /^Margins$/).click()
      await menuItem(p, /^Mirrored/).click()
      await ribbonBtn(p, /^Size$/).click()
      await menuItem(p, /^More Paper Sizes/).click()
      const paper = p.locator('.modal', { hasText: 'Paper Size' })
      await paper.getByLabel(/^Width/).fill('20')
      await paper.getByLabel(/^Height/).fill('30')
      await paper.getByRole('button', { name: /^OK$/ }).click()

      await saveTo(launched, p, out)
      const doc = part(out, 'word/document.xml')
      expect(doc).toMatch(/<w:cols [^>]*w:num="3"[^>]*w:sep="1"|<w:cols [^>]*w:sep="1"[^>]*w:num="3"/)
      expect(doc).toContain('<w:lnNumType')
      expect(doc).toContain('w:restart="newSection"')
      expect(doc).toMatch(/<w:pgSz [^>]*w:w="11339"[^>]*w:h="17008"/)
      expect(doc.match(/<w:pgMar[^>]*>/)?.[0]).toContain('w:left="1800"')
      expect(doc).toContain('<w:br w:type="column"/>')
      expect(doc).toMatch(/<w:br w:type="page"\/>|<w:pageBreakBefore\/>/)
      const settings = part(out, 'word/settings.xml')
      expect(settings).toContain('<w:autoHyphenation/>')
      expect(settings).toContain('<w:mirrorMargins/>')
    } finally {
      await kill(launched)
    }
  })

  test('layout changes join the ⌘Z stack, including the document-wide switches', async () => {
    test.setTimeout(120_000)
    const { launched, p } = await openDocs()
    try {
      await p.locator('.doc-page').click()
      await p.keyboard.type('Undo me', { delay: 5 })
      await tab(p, 'Layout')
      await ribbonBtn(p, /Hyphenation$/).click()
      await menuItem(p, /^Automatic$/).click()
      await expect(p.locator('.doc-page')).toHaveAttribute('data-hyphenate', 'on')
      await p.keyboard.press('Meta+z')
      await expect(p.locator('.doc-page')).toHaveAttribute('data-hyphenate', 'off')
      await p.keyboard.press('Meta+Shift+z')
      await expect(p.locator('.doc-page')).toHaveAttribute('data-hyphenate', 'on')
    } finally {
      await kill(launched)
    }
  })
})

test.describe('Home tab: borders, spacing, styles and paste options', () => {
  test('the Borders menu toggles sides, writes between, and Borders and Shading sets style and color', async () => {
    test.setTimeout(180_000)
    const { launched, p } = await openDocs()
    const out = join(mkdtempSync(join(tmpdir(), 'docs-gaps-')), 'borders.docx')
    try {
      await p.locator('.doc-page').click()
      await p.keyboard.type('First paragraph', { delay: 5 })
      await p.keyboard.press('Enter')
      await p.keyboard.type('Second paragraph', { delay: 5 })

      const bordersMenu = () => p.locator('.ribbon-body button[aria-label="Paragraph Borders"]').click()
      // Bottom then Top: the entries combine (Word toggles each side), a second click removes it
      await bordersMenu()
      await menuItem(p, /^Bottom Border$/).click()
      await bordersMenu()
      await expect(menuItem(p, /^Bottom Border$/)).toHaveClass(/active/)
      await menuItem(p, /^Top Border$/).click()
      await bordersMenu()
      await expect(menuItem(p, /^Top Border$/)).toHaveClass(/active/)
      await expect(menuItem(p, /^Bottom Border$/)).toHaveClass(/active/)
      await menuItem(p, /^Bottom Border$/).click()
      await bordersMenu()
      await expect(menuItem(p, /^Bottom Border$/)).not.toHaveClass(/active/)
      // the vertical entry is shown but unavailable for paragraphs, as in Word
      await expect(menuItem(p, /^Inside Vertical Border$/)).toBeDisabled()
      await menuItem(p, /^Bottom Border$/).click() // closes the menu, top + bottom again

      // both paragraphs: All Borders adds the rule between them
      await p.keyboard.press('Meta+a')
      await bordersMenu()
      await menuItem(p, /^All Borders$/).click()
      // Borders and Shading…: double red 3 pt
      await bordersMenu()
      await menuItem(p, /^Borders and Shading/).click()
      const dialog = p.locator('.borders-dialog')
      await expect(dialog).toBeVisible()
      await dialog.getByLabel('Style').selectOption('double')
      await dialog.getByLabel(/^Border weight/).selectOption('3')
      await dialog.getByLabel('Automatic').uncheck()
      await dialog.locator('input[type="color"]').first().fill('#ff0000')
      await dialog.getByRole('button', { name: /^OK$/ }).click()
      await expect(dialog).toBeHidden()

      // Horizontal Line adds a bottom-bordered paragraph after the caret's paragraph;
      // View Gridlines is a display toggle
      await p.getByText('First paragraph').click()
      await bordersMenu()
      await menuItem(p, /^Horizontal Line$/).click()
      await bordersMenu()
      await menuItem(p, /^View Gridlines$/).click()
      await expect(p.locator('.doc-page')).toHaveClass(/show-table-gridlines/)

      await saveTo(launched, p, out)
      const doc = part(out, 'word/document.xml')
      const pBdr = doc.match(/<w:pBdr>.*?<\/w:pBdr>/g) ?? []
      expect(pBdr.length).toBeGreaterThanOrEqual(3) // two paragraphs + the horizontal line
      expect(pBdr.some((b) => b.includes('<w:between') && b.includes('w:val="double"') && b.includes('w:color="FF0000"') && b.includes('w:sz="24"'))).toBe(true)
      // the line itself: an empty paragraph carrying only a bottom border, between the two texts
      expect(doc).toMatch(
        /First paragraph<\/w:t><\/w:r><\/w:p><w:p><w:pPr><w:pBdr><w:bottom [^>]*\/><\/w:pBdr><\/w:pPr><\/w:p>/,
      )
    } finally {
      await kill(launched)
    }
  })

  test('Line Spacing → Add / Remove Space Before and After Paragraph', async () => {
    test.setTimeout(120_000)
    const { launched, p } = await openDocs()
    const out = join(mkdtempSync(join(tmpdir(), 'docs-gaps-')), 'spacing.docx')
    try {
      await p.locator('.doc-page').click()
      await p.keyboard.type('Spaced paragraph', { delay: 5 })
      const spacing = () => p.locator('.ribbon-body button[aria-label="Line Spacing"]').click()
      await spacing()
      await menuItem(p, /^Add Space Before Paragraph$/).click()
      await spacing()
      await menuItem(p, /^Add Space After Paragraph$/).click()
      await spacing()
      await expect(menuItem(p, /^Remove Space Before Paragraph$/)).toBeVisible()
      await expect(menuItem(p, /^Remove Space After Paragraph$/)).toBeVisible()
      await menuItem(p, /^Remove Space After Paragraph$/).click()
      await saveTo(launched, p, out)
      const spacingTag = part(out, 'word/document.xml').match(/<w:spacing [^>]*>/g)?.join(' ') ?? ''
      expect(spacingTag).toContain('w:before="240"')
      expect(spacingTag).toMatch(/w:after="0"/)
    } finally {
      await kill(launched)
    }
  })

  test('the Styles gallery offers Title / Quote and the Styles pane applies them, saved as real styles', async () => {
    test.setTimeout(180_000)
    const { launched, p } = await openDocs()
    const out = join(mkdtempSync(join(tmpdir(), 'docs-gaps-')), 'styles.docx')
    try {
      await p.locator('.doc-page').click()
      await p.keyboard.type('A title line', { delay: 5 })
      await p.keyboard.press('Enter')
      await p.keyboard.type('A quoted line', { delay: 5 })
      // open the gallery expander → Styles Pane
      await p.locator('.style-gallery-more').click()
      await expect(styleCard(p, 'Title')).toBeVisible()
      await expect(styleCard(p, 'Intense Quote')).toBeVisible()
      await p.locator('.style-gallery-footer').click()
      const pane = p.locator('.styles-pane')
      await expect(pane).toBeVisible()
      await pane.locator('.styles-pane-item', { hasText: /^Quote/ }).click()
      await expect(p.locator('.doc-page [data-style="Quote"]')).toHaveCount(1)
      await expect(pane.locator('.styles-pane-item.active')).toContainText('Quote')
      // the quote is centered and italic at once (display info from the built-in definition)
      const css = await p.locator('.doc-page [data-style="Quote"]').evaluate((el) => {
        const cs = getComputedStyle(el)
        return { align: cs.textAlign, italic: cs.fontStyle }
      })
      expect(css).toEqual({ align: 'center', italic: 'italic' })
      // Title from the gallery card on the first line
      await p.getByText('A title line').click()
      await p.locator('.style-gallery-more').click()
      await styleCard(p, 'Title').click()
      await expect(p.locator('.doc-page [data-style="Title"]')).toHaveCount(1)

      await saveTo(launched, p, out)
      const doc = part(out, 'word/document.xml')
      expect(doc).toContain('<w:pStyle w:val="Title"/>')
      expect(doc).toContain('<w:pStyle w:val="Quote"/>')
      const styles = part(out, 'word/styles.xml')
      expect(styles).toContain('w:styleId="Title"')
      expect(styles).toContain('w:styleId="Quote"')
    } finally {
      await kill(launched)
    }
  })

  test('Paste ▾ offers Keep Source / Merge / Keep Text Only for one paste', async () => {
    test.setTimeout(120_000)
    const { launched, p } = await openDocs()
    try {
      await p.locator('.doc-page').click()
      await p.keyboard.type('x', { delay: 5 })
      await launched.app.evaluate(({ clipboard }) => {
        clipboard.write({
          html: `<meta charset='utf-8'><p style="font-family: Arial; font-size: 14px;">pasted <b>bold</b></p>`,
          text: 'pasted bold',
        })
      })
      await p.locator('.ribbon-body button[aria-label="Paste Options"]').first().click()
      for (const name of ['Keep Source Formatting', 'Merge Formatting', 'Keep Text Only']) {
        await expect(menuItem(p, new RegExp(`^${name}$`))).toBeVisible()
      }
      await menuItem(p, /^Keep Text Only$/).click()
      await expect(p.locator('.doc-page')).toContainText('pasted bold')
      // text only: the bold run is gone
      expect(await p.locator('.doc-page strong, .doc-page b').count()).toBe(0)
    } finally {
      await kill(launched)
    }
  })
})

test.describe('Design tab: paragraph spacing presets and the Page Border dialog', () => {
  test('Paragraph Spacing rewrites the document spacing, undoes with ⌘Z and saves into Normal', async () => {
    test.setTimeout(180_000)
    const { launched, p } = await openDocs()
    const out = join(mkdtempSync(join(tmpdir(), 'docs-gaps-')), 'design.docx')
    try {
      await p.locator('.doc-page').click()
      await p.keyboard.type('First paragraph', { delay: 5 })
      await p.keyboard.press('Enter')
      await p.keyboard.type('Second paragraph', { delay: 5 })
      const gap = () =>
        p.locator('.doc-page p').first().evaluate((el) => {
          const cs = getComputedStyle(el)
          return { after: parseFloat(cs.marginBottom), line: parseFloat(cs.lineHeight) / parseFloat(cs.fontSize) }
        })
      const before = await gap()
      await tab(p, 'Design')
      await ribbonBtn(p, /^Paragraph Spacing$/).click()
      for (const name of ['Default', 'No Paragraph Space', 'Compact', 'Tight', 'Open', 'Relaxed', 'Double']) {
        await expect(menuItem(p, new RegExp(`^${name}`))).toBeVisible()
      }
      await menuItem(p, /^No Paragraph Space/).click()
      await expect.poll(async () => (await gap()).after).toBe(0)
      await ribbonBtn(p, /^Paragraph Spacing$/).click()
      await expect(menuItem(p, /^No Paragraph Space/)).toHaveClass(/active/)
      await menuItem(p, /^Double/).click()
      await expect.poll(async () => (await gap()).line).toBeGreaterThan(1.9)
      // ⌘Z puts the previous preset back, then the file's own spacing
      await p.locator('.doc-page').click()
      await p.keyboard.press('Meta+z')
      await expect.poll(async () => (await gap()).after).toBe(0)
      await expect.poll(async () => (await gap()).line).toBeLessThan(1.5)
      await p.keyboard.press('Meta+z')
      await expect.poll(async () => (await gap()).after).toBeCloseTo(before.after, 0)
      await p.keyboard.press('Meta+Shift+z')
      await p.keyboard.press('Meta+Shift+z')

      await saveTo(launched, p, out)
      const styles = part(out, 'word/styles.xml')
      const normal = styles.match(/<w:style [^>]*w:styleId="Normal".*?<\/w:style>/s)?.[0] ?? ''
      expect(normal).toMatch(/<w:spacing [^>]*w:after="160"[^>]*w:line="480"|<w:spacing [^>]*w:line="480"[^>]*w:after="160"/)
    } finally {
      await kill(launched)
    }
  })

  test('Page Borders… sets style, color and the pages it covers', async () => {
    test.setTimeout(120_000)
    const { launched, p } = await openDocs()
    const out = join(mkdtempSync(join(tmpdir(), 'docs-gaps-')), 'pageborder.docx')
    try {
      await p.locator('.doc-page').click()
      await p.keyboard.type('A bordered page', { delay: 5 })
      await tab(p, 'Design')
      await ribbonBtn(p, /^Page Borders$/).click()
      await menuItem(p, /^Borders and Shading/).click()
      const dialog = p.locator('.borders-dialog')
      await expect(dialog).toBeVisible()
      await dialog.getByLabel('Style').selectOption('double')
      await dialog.getByLabel(/^Border weight/).selectOption('3')
      await dialog.getByLabel('Automatic').uncheck()
      await dialog.locator('input[type="color"]').first().fill('#0000ff')
      await dialog.getByLabel('Apply to').selectOption('firstPage')
      await dialog.getByLabel('Left Border').uncheck()
      await dialog.getByRole('button', { name: /^OK$/ }).click()
      await expect(dialog).toBeHidden()
      await saveTo(launched, p, out)
      const doc = part(out, 'word/document.xml')
      const box = doc.match(/<w:pgBorders[^>]*>.*?<\/w:pgBorders>/s)?.[0] ?? ''
      expect(box).toContain('w:display="firstPage"')
      expect(box).toContain('<w:top w:val="double" w:sz="24" w:space="24" w:color="0000FF"/>')
      expect(box).toContain('<w:bottom')
      expect(box).toContain('<w:right')
      expect(box).not.toContain('<w:left')
    } finally {
      await kill(launched)
    }
  })
})
