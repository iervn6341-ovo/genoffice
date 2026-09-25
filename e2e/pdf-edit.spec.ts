import { test, expect, type Page } from '@playwright/test'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { deflateSync } from 'node:zlib'
import { launchShell, waitForPageWithUrl, type LaunchedApp } from './helpers'

/**
 * PDF editing as a user does it: restyle a text run (size / bold / font / colour), move and
 * resize a picture, insert a picture and new text — then save and read the file back with
 * pdf.js. The fixture draws its heading through a 0.75 matrix like Word / Chrome exports do,
 * which is where a picked size used to come out 25% too small.
 */

const req = createRequire(join(process.cwd(), 'package.json'))

/** Solid-colour RGB PNG (minimal encoder: IHDR + one IDAT + IEND) */
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

async function fixture(): Promise<{ dir: string; pdf: string; picture: string }> {
  const dir = mkdtempSync(join(tmpdir(), 'genoffice-pdf-edit-'))
  const {
    PDFDocument,
    StandardFonts,
    concatTransformationMatrix,
    pushGraphicsState,
    popGraphicsState,
  } = await import('pdf-lib')
  const doc = await PDFDocument.create()
  const page = doc.addPage([612, 792])
  const font = await doc.embedFont(StandardFonts.Helvetica)
  // Tf 16 through a 0.75 matrix: a 12pt heading at (60, 675)
  page.pushOperators(pushGraphicsState(), concatTransformationMatrix(0.75, 0, 0, 0.75, 0, 0))
  page.drawText('Scaled heading text', { x: 80, y: 900, size: 16, font })
  page.pushOperators(popGraphicsState())
  page.drawText('Body line stays put', { x: 60, y: 640, size: 12, font })
  const img = await doc.embedPng(png(80, 60, [30, 120, 200]))
  page.drawImage(img, { x: 100, y: 330, width: 240, height: 180 })
  // more pages than the sidebar shows at once: its scrollbar narrows the thumbnails below
  // the nominal raster width, which is where pending-edit marks used to drift
  for (let i = 0; i < 5; i++)
    doc.addPage([612, 792]).drawText(`Page ${i + 2}`, { x: 60, y: 700, size: 12, font })
  const pdf = join(dir, 'fixture.pdf')
  writeFileSync(pdf, await doc.save({ useObjectStreams: false }))
  const picture = join(dir, 'picture.png')
  writeFileSync(picture, png(40, 20, [220, 40, 40])) // 2:1, unlike the 4:3 picture
  return { dir, pdf, picture }
}

interface PdfContent {
  items: { str: string; font: string; size: number; x: number; y: number }[]
  images: { w: number; h: number; x: number; y: number }[]
  fills: string[]
}

/** Page 1 as pdf.js reads the saved file: text runs (real font names), placed images, fills */
async function readPage(file: string): Promise<PdfContent> {
  const pdfjs = await import(req.resolve('pdfjs-dist/legacy/build/pdf.mjs'))
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(readFileSync(file)),
    disableFontFace: true,
    verbosity: 0,
  }).promise
  try {
    const page = await doc.getPage(1)
    const tc = await page.getTextContent()
    const ops = await page.getOperatorList()
    const fontOf = (id: string) => {
      try {
        return String(page.commonObjs.get(id).name ?? '')
      } catch {
        return id
      }
    }
    const items = tc.items
      .filter((it: { str?: string }) => it.str)
      .map((it: { str: string; fontName: string; transform: number[] }) => ({
        str: it.str,
        font: fontOf(it.fontName),
        size: Math.hypot(it.transform[2]!, it.transform[3]!),
        x: it.transform[4]!,
        y: it.transform[5]!,
      }))
    const mul = (m: number[], n: number[]) => [
      m[0]! * n[0]! + m[1]! * n[2]!,
      m[0]! * n[1]! + m[1]! * n[3]!,
      m[2]! * n[0]! + m[3]! * n[2]!,
      m[2]! * n[1]! + m[3]! * n[3]!,
      m[4]! * n[0]! + m[5]! * n[2]! + n[4]!,
      m[4]! * n[1]! + m[5]! * n[3]! + n[5]!,
    ]
    let ctm = [1, 0, 0, 1, 0, 0]
    const stack: number[][] = []
    const images: PdfContent['images'] = []
    const fills: string[] = []
    ops.fnArray.forEach((fn: number, i: number) => {
      if (fn === pdfjs.OPS.save) stack.push(ctm)
      if (fn === pdfjs.OPS.restore) ctm = stack.pop() ?? [1, 0, 0, 1, 0, 0]
      if (fn === pdfjs.OPS.transform) ctm = mul(ops.argsArray[i], ctm)
      if (fn === pdfjs.OPS.setFillRGBColor)
        fills.push(String(ops.argsArray[i][0] ?? ops.argsArray[i]))
      if (fn === pdfjs.OPS.paintImageXObject)
        images.push({ w: ctm[0]!, h: ctm[3]!, x: ctm[4]!, y: ctm[5]! })
    })
    return { items, images, fills }
  } finally {
    await doc.loadingTask.destroy()
  }
}

async function openPdf(file: string): Promise<{ p: Page; launched: LaunchedApp }> {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'pdf-edit', openFile: file })
  const p = await waitForPageWithUrl(launched.app, '://pdf/', 30_000)
  await p.waitForSelector('.pdf-page canvas', { timeout: 20_000 })
  await p.waitForTimeout(1500)
  await p.getByText('Edit', { exact: true }).first().click()
  return { p, launched }
}

const pageBox = async (p: Page) => (await p.locator('.pdf-page').first().boundingBox())!
/** ⌘S, then wait for the file itself: the "Unsaved" badge clears before the write lands */
const save = async (p: Page, file: string) => {
  const before = readFileSync(file)
  await p.keyboard.press('ControlOrMeta+s')
  await expect.poll(() => !readFileSync(file).equals(before), { timeout: 15_000 }).toBe(true)
  await expect(p.getByText('Unsaved')).toHaveCount(0, { timeout: 15_000 })
}

test('edit text: picked size, bold, font and colour save as shown; the thumbnail mirrors it in place', async () => {
  const f = await fixture()
  const { p, launched } = await openPdf(f.pdf)
  try {
    await p.locator('button', { hasText: 'Edit text' }).first().click()
    const run = (await p.getByText('Scaled heading text').first().boundingBox())!
    await p.mouse.click(run.x + run.width / 2, run.y + run.height / 2)
    await p.keyboard.press('ControlOrMeta+a')
    await p.keyboard.type('Renamed heading')
    const size = p.locator('.pdf-textedit-sizenum')
    await expect(size).toHaveValue('12') // the page-space size, not the Tf 16
    await size.fill('14')
    await p.locator('.pdf-textedit-toggle', { hasText: 'B' }).click()
    await p.locator('.pdf-textedit-fontsel .gs-dd-btn').click()
    await p.locator('.gs-dd-item', { hasText: 'Times New Roman' }).click()
    await p.locator('.pdf-textedit-bar button[data-tip="Color"]').click()
    await p.locator('.pdf-color-popover .gcp-standard-row .gcp-swatch').nth(1).click()
    const pb = await pageBox(p)
    await p.mouse.click(pb.x + pb.width - 10, pb.y + 10) // commit
    await expect(p.locator('.pdf-page .pdf-textedit-preview')).toHaveCount(1)

    // the sidebar thumbnail draws the pending edit where the page does
    const where = await p.evaluate(() => {
      const rel = (el: Element, box: Element) => {
        const r = el.getBoundingClientRect()
        const b = box.getBoundingClientRect()
        return { x: (r.left - b.left) / b.width, y: (r.top - b.top) / b.height }
      }
      const thumb = document.querySelector('.pdf-thumb-box')!
      const main = document.querySelector('.pdf-page .pdf-textedit-preview')!
      return {
        thumb: rel(thumb.querySelector('.pdf-textedit-preview')!, thumb.querySelector('canvas')!),
        page: rel(main, main.closest('.pdf-page')!),
      }
    })
    expect(where.thumb.x).toBeCloseTo(where.page.x, 2)
    expect(where.thumb.y).toBeCloseTo(where.page.y, 2)

    await save(p, f.pdf)
    const { items, fills } = await readPage(f.pdf)
    const edited = items.find((i) => i.str === 'Renamed heading')!
    expect(edited.size).toBeCloseTo(14, 1)
    // Linux has no Times New Roman: fontconfig substitutes the metric-compatible Liberation Serif
    expect(edited.font).toMatch(/(TimesNewRoman|LiberationSerif).*Bold/)
    expect(fills).toContain('#ff0000')
    expect(items.find((i) => i.str === 'Body line stays put')?.size).toBeCloseTo(12, 1)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('images: a moved picture stays selected and resizes; Insert image places a new one', async () => {
  const f = await fixture()
  const { p, launched } = await openPdf(f.pdf)
  try {
    await p.locator('button', { hasText: 'Edit images' }).first().click()
    const pb = await pageBox(p)
    // the picture: x 100..340, y 330..510 of a 612×792 page
    const k = pb.width / 612
    const cx = pb.x + 220 * k
    const cy = pb.y + (792 - 420) * k
    await p.mouse.move(cx, cy)
    await p.mouse.down()
    await p.mouse.move(cx + 46 * k, cy + 23 * k, { steps: 8 })
    await p.mouse.up()
    // still selected: the corner handles stay for an immediate resize (PowerPoint / Acrobat)
    await expect(p.locator('.pdf-imgedit-handle')).toHaveCount(4)
    const h = (await p.locator('.pdf-imgedit-handle').last().boundingBox())!
    await p.mouse.move(h.x + h.width / 2, h.y + h.height / 2)
    await p.mouse.down()
    await p.mouse.move(h.x - 60 * k, h.y - 45 * k, { steps: 8 })
    await p.mouse.up()
    await expect(p.locator('.pdf-imgedit-handle')).toHaveCount(4)

    const [chooser] = await Promise.all([
      p.waitForEvent('filechooser'),
      p.locator('button', { hasText: 'Insert image' }).first().click(),
    ])
    await chooser.setFiles(f.picture)
    // placement mode once the picked file is decoded
    await expect(p.locator('button', { hasText: 'Insert image' }).first()).toHaveClass(/active/)
    await p.mouse.click(pb.x + 450 * k, pb.y + 150 * k)
    await expect(p.locator('button', { hasText: 'Insert image' }).first()).not.toHaveClass(/active/)
    await save(p, f.pdf)

    const { images } = await readPage(f.pdf)
    expect(images).toHaveLength(2)
    const moved = images.find((i) => Math.abs(i.w / i.h - 240 / 180) < 0.02)!
    expect(moved.x).toBeCloseTo(146, -1) // moved 46pt right
    expect(moved.w).toBeLessThan(230) // and made smaller, aspect kept
    const inserted = images.find((i) => i !== moved)!
    expect(inserted.w / inserted.h).toBeCloseTo(2, 1)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('Insert text: the dialog’s font, bold and italic land in the file', async () => {
  const f = await fixture()
  const { p, launched } = await openPdf(f.pdf)
  try {
    await p.locator('button', { hasText: 'Insert text' }).first().click()
    await p.keyboard.type('Added remark')
    await p.locator('.pdf-insert-face .gs-dd-btn').click()
    await p.locator('.gs-dd-item', { hasText: 'Courier New' }).click()
    await p.locator('.pdf-insert-face button[aria-label="Bold"]').click()
    await p.locator('.pdf-insert-face button[aria-label="Italic"]').click()
    await expect(p.locator('.pdf-modal-textarea')).toHaveCSS('font-style', 'italic')
    await p.locator('.pdf-modal button', { hasText: 'OK' }).click()
    await expect(p.locator('.pdf-modal')).toHaveCount(0)
    const pb = await pageBox(p)
    await p.mouse.move(pb.x + pb.width * 0.1, pb.y + pb.height * 0.05)
    await p.mouse.click(pb.x + pb.width * 0.1, pb.y + pb.height * 0.05)
    await expect(p.locator('.pdf-page .pdf-textinsert-preview')).toHaveCount(1)
    await save(p, f.pdf)
    const { items } = await readPage(f.pdf)
    // pdf.js splits the run at word boundaries: read the whole baseline
    const added = items.filter(
      (i) => Math.abs(i.y - (items.find((j) => j.str === 'Added')?.y ?? -1)) < 1,
    )
    expect(added.map((i) => i.str).join('')).toBe('Added remark')
    for (const run of added) {
      // Linux has no Courier New: fontconfig substitutes the metric-compatible Liberation Mono
      expect(run.font).toMatch(/(CourierNew|LiberationMono).*BoldItalic/)
      expect(run.size).toBeCloseTo(14, 1)
    }
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

/** Every page's text, in file order */
async function pageTexts(file: string): Promise<string[]> {
  const pdfjs = await import(req.resolve('pdfjs-dist/legacy/build/pdf.mjs'))
  const doc = await pdfjs.getDocument({
    data: new Uint8Array(readFileSync(file)),
    disableFontFace: true,
    verbosity: 0,
  }).promise
  try {
    const out: string[] = []
    for (let n = 1; n <= doc.numPages; n++) {
      const tc = await (await doc.getPage(n)).getTextContent()
      out.push(tc.items.map((it: { str?: string }) => it.str ?? '').join(''))
    }
    return out
  } finally {
    await doc.loadingTask.destroy()
  }
}

test('Import pages after a deleted page lands behind the page that was right-clicked', async () => {
  const f = await fixture()
  const { PDFDocument, StandardFonts } = await import('pdf-lib')
  const extra = await PDFDocument.create()
  extra.addPage([612, 792]).drawText('Imported page', {
    x: 60,
    y: 700,
    size: 12,
    font: await extra.embedFont(StandardFonts.Helvetica),
  })
  const extraPath = join(f.dir, 'extra.pdf')
  writeFileSync(extraPath, await extra.save())
  const { p, launched } = await openPdf(f.pdf)
  try {
    await launched.app.evaluate(({ dialog }, picked) => {
      dialog.showOpenDialog = (async () => ({ canceled: false, filePaths: [picked] })) as never
    }, extraPath)
    const thumbs = p.locator('.pdf-thumb')
    await expect(thumbs).toHaveCount(6)
    await thumbs.first().click({ button: 'right' })
    await p.locator('.thumb-menu button', { hasText: 'Delete page' }).click()
    await expect(thumbs).toHaveCount(5)
    // "Page 2" is now first on screen; import after it — the pending delete shifts file
    // indices, which is where the import used to land one page too late
    await thumbs.first().click({ button: 'right' })
    await p.locator('.thumb-menu button', { hasText: 'Import pages' }).click()
    await expect(thumbs).toHaveCount(6, { timeout: 20_000 })
    await expect
      .poll(() => pageTexts(f.pdf), { timeout: 15_000 })
      .toEqual(['Page 2', 'Imported page', 'Page 3', 'Page 4', 'Page 5', 'Page 6'])
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('redaction marks and Clear marks go through undo / redo', async () => {
  const f = await fixture()
  const { p, launched } = await openPdf(f.pdf)
  try {
    await p.locator('.ribbon-tab', { hasText: 'Annotate' }).click()
    await p.locator('button', { hasText: 'Redact area' }).first().click()
    const marks = p.locator('.pdf-redaction-mark')
    const pb = await pageBox(p)
    const drag = async (x: number, y: number) => {
      await p.mouse.move(pb.x + pb.width * x, pb.y + pb.height * y)
      await p.mouse.down()
      await p.mouse.move(pb.x + pb.width * (x + 0.2), pb.y + pb.height * (y + 0.05), { steps: 6 })
      await p.mouse.up()
    }
    await drag(0.1, 0.1)
    await drag(0.1, 0.3)
    await expect(marks).toHaveCount(2)
    await p.keyboard.press('ControlOrMeta+z')
    await expect(marks).toHaveCount(1)
    await p.keyboard.press('ControlOrMeta+Shift+z')
    await expect(marks).toHaveCount(2)
    await p.locator('button', { hasText: 'Clear marks' }).click()
    await expect(marks).toHaveCount(0)
    await p.keyboard.press('ControlOrMeta+z')
    await expect(marks).toHaveCount(2)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})

test('Crop pages: the dialog fits the window under the ribbon and crops the page', async () => {
  const f = await fixture()
  const { p, launched } = await openPdf(f.pdf)
  try {
    await p.locator('.ribbon-tab', { hasText: 'Pages' }).click()
    await p.locator('.ribbon-body button', { hasText: 'Crop pages' }).first().click()
    const frame = p.locator('.pdf-modal div[style*="cursor: move"]').first()
    await expect(frame).toBeVisible()
    // the whole dialog — bottom handles and Apply — stays inside the window (it used to be
    // sized against the full window height and hang below the ribbon-shortened pane)
    const fits = await p.evaluate(() => {
      const modal = document.querySelector('.pdf-modal')!.getBoundingClientRect()
      return modal.bottom <= innerHeight && modal.top >= 0
    })
    expect(fits).toBe(true)
    const b = (await frame.boundingBox())!
    await p.mouse.move(b.x + b.width - 2, b.y + b.height - 2)
    await p.mouse.down()
    await p.mouse.move(b.x + b.width * 0.5, b.y + b.height * 0.5, { steps: 8 })
    await p.mouse.up()
    const before = readFileSync(f.pdf)
    await p.locator('.pdf-modal button', { hasText: 'Apply' }).click()
    await expect.poll(() => !readFileSync(f.pdf).equals(before), { timeout: 15_000 }).toBe(true)
    const { PDFDocument } = await import('pdf-lib')
    const crop = (await PDFDocument.load(readFileSync(f.pdf))).getPage(0).getCropBox()
    expect(crop.width).toBeCloseTo(306, -1)
    expect(crop.height).toBeCloseTo(396, -1)
  } finally {
    launched.app.process().kill('SIGKILL')
  }
})
