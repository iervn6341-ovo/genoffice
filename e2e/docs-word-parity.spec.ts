import { existsSync } from 'node:fs'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * Docs against Word 16 itself. The expected XML is what Word wrote (Word.Application through COM,
 * SaveAs2 to .docx) after the same ribbon choice: Home → Borders (the gap Word leaves between a
 * paragraph border and its text), Layout → Line Numbers, Design → Page Borders ▸ Box.
 */
async function openDocs(): Promise<{ launched: LaunchedApp; p: Page }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-word-parity' })
  await launched.page.locator('.quick-card').first().click()
  const p = await waitForPageWithUrl(launched.app, '://docs/', 30_000)
  await p.waitForSelector('.doc-page', { timeout: 20_000 })
  await p.waitForFunction(() =>
    Boolean((window as unknown as { __aidocs?: { editor?: unknown } }).__aidocs?.editor),
  )
  return { launched, p }
}
const kill = (l: LaunchedApp) => closeAndSaveVideo(l, 'cleanup')

async function saveTo(l: LaunchedApp, p: Page, out: string): Promise<string> {
  await l.app.evaluate(({ dialog }, file) => {
    dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
  }, out)
  await p.keyboard.press('Control+Shift+s')
  await expect.poll(() => existsSync(out), { timeout: 20_000 }).toBe(true)
  await new Promise((resolve) => setTimeout(resolve, 1200))
  const zip = await JSZip.loadAsync(await readFile(out))
  return zip.file('word/document.xml')!.async('string')
}
const tab = (p: Page, name: string) =>
  p.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) }).click()
const ribbonBtn = (p: Page, name: RegExp | string) =>
  p.locator('.ribbon-body button:not(.rb-fold-btn)', { hasText: name }).first()
const menuItem = (p: Page, name: RegExp | string) =>
  p.locator('[data-rb-panel] button', { hasText: name }).first()

// what Word wrote for the paragraph borders of Home → Borders (Word 16, saved .docx)
const WORD_BOTTOM =
  '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>'
const WORD_OUTSIDE =
  '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>' +
  '<w:left w:val="single" w:sz="4" w:space="4" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="1" w:color="auto"/>' +
  '<w:right w:val="single" w:sz="4" w:space="4" w:color="auto"/></w:pBdr>'

test('Home → Borders leaves the gap Word leaves between the border and the text', async () => {
  test.setTimeout(180_000)
  const { launched, p } = await openDocs()
  const dir = await mkdtemp(join(tmpdir(), 'docs-word-parity-'))
  try {
    await p.locator('.doc-page').click()
    await p.keyboard.type('A bordered paragraph', { delay: 5 })
    await p.keyboard.press('Enter')
    await p.keyboard.type('Another one', { delay: 5 })
    await tab(p, 'Home')

    // Bottom Border on the second paragraph
    await p.locator('button[aria-label="Paragraph Borders"]').first().click()
    await menuItem(p, /^Bottom Border/).click()
    // Outside Borders on the first
    await p.keyboard.press('Control+Home')
    await p.locator('button[aria-label="Paragraph Borders"]').first().click()
    await menuItem(p, /^Outside Borders/).click()

    const doc = (await saveTo(launched, p, join(dir, 'borders.docx'))).replace(/\s+/g, ' ')
    expect(doc).toContain(WORD_OUTSIDE)
    expect(doc).toContain(WORD_BOTTOM)
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})

test('Layout → Line Numbers ▸ Continuous and Design → Page Borders ▸ Box write Word’s elements', async () => {
  test.setTimeout(180_000)
  const { launched, p } = await openDocs()
  const dir = await mkdtemp(join(tmpdir(), 'docs-word-parity2-'))
  try {
    await p.locator('.doc-page').click()
    await p.keyboard.type('Numbered line', { delay: 5 })
    await tab(p, 'Layout')
    await ribbonBtn(p, /Line Numbers$/).click()
    await menuItem(p, /^Continuous$/).click()

    const doc = (await saveTo(launched, p, join(dir, 'lines.docx'))).replace(/\s+/g, ' ')
    // Word: <w:lnNumType w:countBy="1" w:restart="continuous"/>
    expect(doc).toMatch(/<w:lnNumType w:countBy="1"[^>]*w:restart="continuous"\/>/)
  } finally {
    await kill(launched)
    await rm(dir, { recursive: true, force: true })
  }
})
