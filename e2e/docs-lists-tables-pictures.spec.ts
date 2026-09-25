import { test, expect, type Page } from '@playwright/test'
import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { deflateSync } from 'node:zlib'
import { launchShell, waitForPageWithUrl } from './helpers'

/**
 * Word parity around lists, tables and floating pictures:
 *  - Continue Numbering joins the nearest earlier list of the same format, never a bullet
 *    list sitting in between;
 *  - a whole table picked with its move handle keeps Table Design / Table Layout, and their
 *    commands act on the whole table;
 *  - a picture dragged off its inline spot saves its position and comes back there.
 */

interface DocsWindow {
  __aidocs?: { editor?: any }
}

function png(w: number, h: number, rgb: [number, number, number]): Buffer {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    return c >>> 0
  })
  const crc = (buf: Buffer) => {
    let c = 0xffffffff
    for (const b of buf) c = crcTable[(c ^ b) & 0xff]! ^ (c >>> 8)
    return (c ^ 0xffffffff) >>> 0
  }
  const chunk = (type: string, data: Buffer) => {
    const len = Buffer.alloc(4)
    len.writeUInt32BE(data.length)
    const td = Buffer.concat([Buffer.from(type), data])
    const c = Buffer.alloc(4)
    c.writeUInt32BE(crc(td))
    return Buffer.concat([len, td, c])
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(w, 0)
  ihdr.writeUInt32BE(h, 4)
  ihdr[8] = 8
  ihdr[9] = 2
  const row = Buffer.concat([Buffer.from([0]), Buffer.from(Array(w).fill(rgb).flat())])
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(Buffer.concat(Array(h).fill(row)))),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

async function blankDoc() {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'docs-lists-tables' })
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

const lists = (p: Page) =>
  p.evaluate(() => {
    const out: [string, string][] = []
    ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.forEach((n: any) => {
      if (n.type.name === 'docListItem') out.push([n.textContent, String(n.attrs.numId)])
    })
    return out
  })

const ribbonTabs = (p: Page) => p.locator('.ribbon-tab').allInnerTexts()

test('Continue Numbering skips a bullet list in between', async () => {
  const { p, launched } = await blankDoc()
  try {
    for (const [first, next] of [
      ['1. one', 'two'],
      ['* dot', ''],
      ['1. three', ''],
    ]) {
      await p.keyboard.type(first)
      await p.keyboard.press('Enter')
      if (next) {
        await p.keyboard.type(next)
        await p.keyboard.press('Enter')
      }
      await p.keyboard.press('Enter')
    }
    const before = Object.fromEntries(await lists(p))
    await p.getByText('three').click({ button: 'right' })
    await p.getByText('Continue Numbering', { exact: true }).click()
    await expect.poll(() => lists(p).then((l) => Object.fromEntries(l).three)).toBe(before.one)
    expect(Object.fromEntries(await lists(p)).dot).toBe(before.dot)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('a table selected with its move handle keeps the Table tabs, and Insert Below works on it', async () => {
  const { p, launched } = await blankDoc()
  try {
    await p.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    await p.locator('.ribbon-body button', { hasText: 'Table' }).first().click()
    const cell = p.locator('.table-picker-grid .table-cell').nth(11) // 2 × 2
    await cell.hover()
    await cell.click()
    for (const t of ['a', 'b', 'c', 'd']) {
      await p.keyboard.type(t)
      await p.keyboard.press('Tab')
    }
    // Tab out of the last cell adds a row, as in Word
    const rows = () =>
      p.evaluate(() => {
        let n = 0
        ;(window as unknown as DocsWindow).__aidocs!.editor.state.doc.descendants((x: any) => {
          if (x.type.name === 'docTableRow') n++
        })
        return n
      })
    const grown = await rows()
    const tbl = (await p.locator('.doc-page table').first().boundingBox())!
    await p.mouse.move(tbl.x + 20, tbl.y + 10)
    await p.locator('.doc-table-handle').first().click({ force: true })
    await expect
      .poll(() =>
        p.evaluate(
          () => (window as unknown as DocsWindow).__aidocs!.editor.state.selection.node?.type.name,
        ),
      )
      .toBe('docTable')
    await expect.poll(() => ribbonTabs(p)).toContain('Table Layout')
    await p.locator('.ribbon-tab', { hasText: 'Table Layout' }).click()
    await p
      .locator('.ribbon-body button', { hasText: /Insert Below/ })
      .first()
      .click()
    await expect.poll(rows).toBeGreaterThan(grown)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('a dragged picture keeps its place through save and reopen', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'docs-picture-'))
  const out = join(dir, 'picture.docx')
  const pic = join(dir, 'picture.png')
  writeFileSync(pic, png(120, 80, [30, 120, 200]))
  const { p, launched } = await blankDoc()
  let offset: { dx: number; dy: number }
  try {
    await launched.app.evaluate(
      ({ dialog }, [file, picture]) => {
        dialog.showSaveDialog = (async () => ({ canceled: false, filePath: file })) as never
        dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [picture] })) as never
      },
      [out, pic],
    )
    await p.keyboard.type('Text before the picture.')
    await p.keyboard.press('Enter')
    await p.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    await p.locator('.ribbon-body button', { hasText: 'Picture' }).first().click()
    const img = p.locator('.doc-page img').first()
    await expect(img).toBeVisible()
    const box = (await img.boundingBox())!
    await p.mouse.move(box.x + box.width / 2, box.y + box.height / 2)
    await p.mouse.down()
    await p.mouse.move(box.x + box.width / 2 + 150, box.y + box.height / 2 + 80, { steps: 12 })
    await p.mouse.up()
    const anchor = (await p.getByText('Text before the picture.').boundingBox())!
    const moved = (await img.boundingBox())!
    offset = { dx: moved.x - anchor.x, dy: moved.y - anchor.y }
    expect(offset.dx).toBeGreaterThan(100)
    await p.keyboard.press('Meta+Shift+s')
    await expect.poll(() => existsSync(out), { timeout: 20_000 }).toBe(true)
    await expect
      .poll(() => execFileSync('unzip', ['-p', out, 'word/document.xml']).toString(), {
        timeout: 15_000,
      })
      .toMatch(/<wp:positionH relativeFrom="column"><wp:posOffset>\d+<\/wp:posOffset>/)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
  const re = await launchShell({
    onboardingSeen: true,
    videoDir: 'docs-lists-tables',
    openFile: out,
  })
  try {
    const p2 = await waitForPageWithUrl(re.app, '://docs/')
    const img = p2.locator('.doc-page img').first()
    await expect(img).toBeVisible({ timeout: 30_000 })
    const anchor = (await p2.getByText('Text before the picture.').boundingBox())!
    const at = (await img.boundingBox())!
    expect(at.x - anchor.x).toBeCloseTo(offset.dx, 0)
    expect(at.y - anchor.y).toBeCloseTo(offset.dy, 0)
  } finally {
    re.app.process().kill('SIGKILL')
  }
})
