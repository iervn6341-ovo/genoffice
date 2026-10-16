import { expect, test, type Page } from '@playwright/test'
import JSZip from 'jszip'
import { createRequire } from 'node:module'
import { copyFile, mkdir, mkdtemp, readFile, readdir } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { tmpdir } from 'node:os'

import { closeAndSaveVideo, launchShell, setEditorLayoutWidth, waitForPageWithUrl } from './helpers'

const evidence =
  process.env.GENOFFICE_HF_PICTURE_VERIFICATION_DIR ??
  join(process.cwd(), 'docs/qa/office-mac-hf-picture-evidence-2026-10-13/sheets')
const requireModule = createRequire(join(process.cwd(), 'package.json'))

test.skip(
  !process.env.GENOFFICE_HF_PICTURE_VERIFY,
  'Opt-in native header/footer picture verification',
)
test.use({ actionTimeout: 10_000 })

async function sheetsPage(launched: Awaited<ReturnType<typeof launchShell>>): Promise<Page> {
  const page = await waitForPageWithUrl(launched.app, '://sheets/', 30_000)
  await page.waitForFunction(() => Boolean((window as any).__genofficeDebug?.univerAPI))
  await setEditorLayoutWidth(launched.app, '://sheets/', 2200)
  return page
}

async function openHeaderFooter(launched: Awaited<ReturnType<typeof launchShell>>, page: Page) {
  await setEditorLayoutWidth(launched.app, '://sheets/', 2200)
  await page
    .locator('nav.ribbon-tabs > button', { hasText: /^Insert$/ })
    .first()
    .click()
  await page.locator('[data-ribbon-body] button[data-tip="Header & Footer"]').click()
  const dialog = page.getByRole('dialog', { name: 'Header & Footer' })
  await expect(dialog).toBeVisible()
  await expect(dialog.getByRole('button', { name: 'OK' })).toBeEnabled()
  return dialog
}

async function save(
  launched: Awaited<ReturnType<typeof launchShell>>,
  path: string,
  verify: (zip: JSZip) => Promise<boolean>,
): Promise<JSZip> {
  await launched.app.evaluate(({ webContents }) => {
    webContents
      .getAllWebContents()
      .find((wc) => wc.getURL().includes('://sheets/'))
      ?.send('menu:action', 'save')
  })
  await expect
    .poll(async () => verify(await JSZip.loadAsync(await readFile(path))), { timeout: 30_000 })
    .toBe(true)
  return JSZip.loadAsync(await readFile(path))
}

const xml = async (zip: JSZip, path: string): Promise<string> =>
  (await zip.file(path)?.async('string')) ?? ''

async function imageFile(page: Page, kind: 'png' | 'jpeg', color: string) {
  const data = await page.evaluate(
    ({ kind, color }) => {
      const canvas = document.createElement('canvas')
      canvas.width = 64
      canvas.height = 32
      const context = canvas.getContext('2d')!
      context.fillStyle = color
      context.fillRect(0, 0, 64, 32)
      return canvas.toDataURL(kind === 'png' ? 'image/png' : 'image/jpeg')
    },
    { kind, color },
  )
  return {
    name: `hf-${color.slice(1)}.${kind === 'jpeg' ? 'jpg' : 'png'}`,
    mimeType: kind === 'png' ? 'image/png' : 'image/jpeg',
    buffer: Buffer.from(data.slice(data.indexOf(',') + 1), 'base64'),
  }
}

function slot(dialog: ReturnType<Page['getByRole']>, index: number) {
  return dialog.locator('.dialog-grid > div').nth(index)
}

async function pdfImageBoxes(path: string): Promise<{ top: number; bottomDistance: number }[][]> {
  const pdfjs = await import(requireModule.resolve('pdfjs-dist/legacy/build/pdf.mjs'))
  const document = await pdfjs.getDocument({
    data: new Uint8Array(await readFile(path)),
    disableFontFace: true,
    verbosity: 0,
  }).promise
  try {
    const pages: { top: number; bottomDistance: number }[][] = []
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const page = await document.getPage(pageNumber)
      const ops = await page.getOperatorList()
      let matrix = [1, 0, 0, 1, 0, 0]
      const stack: number[][] = []
      const boxes: { top: number; bottomDistance: number }[] = []
      for (let index = 0; index < ops.fnArray.length; index += 1) {
        const operator = ops.fnArray[index]
        if (operator === pdfjs.OPS.save) stack.push([...matrix])
        else if (operator === pdfjs.OPS.restore) matrix = stack.pop() ?? [1, 0, 0, 1, 0, 0]
        else if (operator === pdfjs.OPS.transform)
          matrix = pdfjs.Util.transform(matrix, ops.argsArray[index])
        else if (
          operator === pdfjs.OPS.paintImageXObject ||
          operator === pdfjs.OPS.paintInlineImageXObject
        ) {
          const y = [
            matrix[5]!,
            matrix[1]! + matrix[5]!,
            matrix[3]! + matrix[5]!,
            matrix[1]! + matrix[3]! + matrix[5]!,
          ]
          boxes.push({
            top: page.getViewport({ scale: 1 }).height - Math.max(...y),
            bottomDistance: Math.min(...y),
          })
        }
      }
      pages.push(boxes)
    }
    return pages
  } finally {
    await document.loadingTask.destroy()
  }
}

test('native &G images survive UI edit, cross-save Undo/Redo, reopen, and PDF', async () => {
  test.setTimeout(300_000)
  await mkdir(evidence, { recursive: true })
  const scratch = await mkdtemp(join(tmpdir(), 'genoffice-hf-picture-'))
  const launched = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-hf-picture',
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  let workbook: string
  try {
    await launched.app.evaluate(({ app }, path) => app.setPath('documents', path), scratch)
    await launched.page.locator('.quick-card').nth(1).click()
    const page = await sheetsPage(launched)
    await page.evaluate(() => {
      const sheet = (window as any).__genofficeDebug.univerAPI.getActiveWorkbook().getActiveSheet()
      sheet
        .getRange('A1:A180')
        .setValues(Array.from({ length: 180 }, (_, i) => [`IMAGE_ROW_${i + 1}`]))
    })
    workbook = join(
      scratch,
      'GenOffice',
      (await readdir(join(scratch, 'GenOffice'))).find((name) => name.endsWith('.xlsx'))!,
    )
    const png = await imageFile(page, 'png', '#1969c9')
    const jpeg = await imageFile(page, 'jpeg', '#c95b19')
    const dialog = await openHeaderFooter(launched, page)
    await slot(dialog, 0).locator('input[type="file"]').setInputFiles(png)
    await expect(dialog.getByLabel('Header · Left')).toHaveValue('&G')
    await dialog.getByRole('tab', { name: 'First page' }).click()
    await dialog.getByLabel('Different first page').check()
    await slot(dialog, 5).locator('input[type="file"]').setInputFiles(jpeg)
    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await dialog.getByLabel('Different odd and even pages').check()
    await slot(dialog, 1).locator('input[type="file"]').setInputFiles(png)
    await dialog.getByRole('button', { name: 'OK' }).click()
    await expect(dialog).toBeHidden()

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    const reopenedDialog = await openHeaderFooter(launched, page)
    await expect(reopenedDialog.getByLabel('Header · Left')).toHaveValue('')
    await reopenedDialog.getByRole('button', { name: 'Cancel' }).click()
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    const first = await save(launched, workbook, async (zip) => {
      const sheet = await xml(zip, 'xl/worksheets/sheet1.xml')
      const vml = await xml(zip, 'xl/drawings/vmlDrawing1.vml')
      return (
        sheet.includes('&amp;G') &&
        ['LH', 'RFFIRST', 'CHEVEN'].every((position) => vml.includes(`id="${position}"`))
      )
    })
    expect(
      (await xml(first, 'xl/drawings/_rels/vmlDrawing1.vml.rels')).match(/relationships\/image/g),
    ).toHaveLength(3)
    await copyFile(workbook, join(evidence, 'hf-pictures-saved.xlsx'))

    await page.getByRole('button', { name: 'Undo', exact: true }).click()
    await save(
      launched,
      workbook,
      async (zip) => !(await xml(zip, 'xl/drawings/vmlDrawing1.vml')).includes('id="LH"'),
    )
    await page.getByRole('button', { name: 'Redo', exact: true }).click()
    await save(launched, workbook, async (zip) =>
      (await xml(zip, 'xl/drawings/vmlDrawing1.vml')).includes('id="LH"'),
    )
    await copyFile(workbook, join(evidence, 'hf-pictures-redone.xlsx'))
  } finally {
    await closeAndSaveVideo(launched, 'sheets-hf-picture')
  }

  const reopened = await launchShell({
    onboardingSeen: true,
    lang: 'en',
    videoDir: 'sheets-hf-picture-reopen',
    openFile: workbook,
    env: { GENOFFICE_DEBUG_HOOKS: '1' },
  })
  try {
    const page = await sheetsPage(reopened)
    const dialog = await openHeaderFooter(reopened, page)
    await expect(slot(dialog, 0).locator('img')).toBeVisible()
    const jpeg = await imageFile(page, 'jpeg', '#49a672')
    await slot(dialog, 0).locator('input[type="file"]').setInputFiles(jpeg)
    await dialog.getByRole('tab', { name: 'First page' }).click()
    await slot(dialog, 5).getByLabel('Width (pt)').fill('70')
    await dialog.getByRole('tab', { name: 'Even pages' }).click()
    await slot(dialog, 1).getByRole('button', { name: 'Remove picture' }).click()
    await dialog.getByRole('button', { name: 'OK' }).click()
    const edited = await save(reopened, workbook, async (zip) => {
      const vml = await xml(zip, 'xl/drawings/vmlDrawing1.vml')
      return vml.includes('id="LH"') && !vml.includes('id="CHEVEN"') && vml.includes('width:70pt')
    })
    expect((await xml(edited, 'xl/worksheets/sheet1.xml')).match(/&amp;G/g)).toHaveLength(2)
    await copyFile(workbook, join(evidence, 'hf-pictures-edited.xlsx'))
    await page.screenshot({ path: join(evidence, 'hf-pictures-reopened.png') })
    const pdf = join(dirname(workbook), 'hf-pictures.pdf')
    await reopened.app.evaluate(({ dialog: native }, path) => {
      native.showSaveDialog = (async () => ({ canceled: false, filePath: path })) as never
    }, pdf)
    await reopened.app.evaluate(({ webContents }) => {
      webContents
        .getAllWebContents()
        .find((wc) => wc.getURL().includes('://sheets/'))
        ?.send('menu:action', 'export-pdf')
    })
    await expect
      .poll(async () => (await readFile(pdf).catch(() => Buffer.alloc(0))).length, {
        timeout: 30_000,
      })
      .toBeGreaterThan(1_000)
    const imageBoxes = await pdfImageBoxes(pdf)
    const imageCounts = imageBoxes.map((boxes) => boxes.length)
    expect(imageCounts.length).toBeGreaterThan(2)
    expect(imageCounts).toEqual(
      imageCounts.map((_, index) => (index === 0 || index % 2 === 0 ? 1 : 0)),
    )
    expect(Math.abs(imageBoxes[0]![0]!.bottomDistance - 21.6)).toBeLessThan(1)
    for (let index = 2; index < imageBoxes.length; index += 2)
      expect(Math.abs(imageBoxes[index]![0]!.top - 21.6)).toBeLessThan(1)
    await copyFile(pdf, join(evidence, 'hf-pictures.pdf'))
  } finally {
    await closeAndSaveVideo(reopened, 'sheets-hf-picture-reopen')
  }
})
